"""
Print Job Queue
================
Thread-safe per-printer job queue with auto-retry.
Replaces the CUPS spooler for Windows -- pure Python, no system dependencies.

Each printer gets its own worker thread that processes jobs sequentially,
preventing TCP socket collisions from concurrent requests.
"""

import uuid
import time
import logging
import threading
from queue import Queue, Empty
from datetime import datetime
from collections import deque

logger = logging.getLogger(__name__)


class PrintQueue:
    """Per-printer job queue with background workers and auto-retry."""

    STATUS_PENDING = "pending"
    STATUS_PRINTING = "printing"
    STATUS_DONE = "done"
    STATUS_FAILED = "failed"
    STATUS_CANCELLED = "cancelled"
    STATUS_RETRYING = "retrying"

    def __init__(self, max_retries=3, retry_base_delay=1.0, history_size=100):
        self.max_retries = max_retries
        self.retry_base_delay = retry_base_delay
        self.history_size = history_size

        self._queues: dict[str, Queue] = {}
        self._workers: dict[str, threading.Thread] = {}
        self._jobs: dict[str, dict] = {}
        self._history: deque[dict] = deque(maxlen=history_size)
        self._lock = threading.Lock()
        self._shutdown = threading.Event()

    def submit(self, printer_name: str, job_type: str, execute_fn, params: dict) -> str:
        """Submit a print job. Returns job_id immediately."""
        job_id = str(uuid.uuid4())
        job = {
            "id": job_id,
            "type": job_type,
            "printer": printer_name,
            "params": params,
            "execute_fn": execute_fn,
            "status": self.STATUS_PENDING,
            "retries": 0,
            "created_at": datetime.now().isoformat(),
            "started_at": None,
            "completed_at": None,
            "error": None,
        }

        with self._lock:
            self._jobs[job_id] = job
            if printer_name not in self._queues:
                self._queues[printer_name] = Queue()
            self._queues[printer_name].put(job_id)
            self._ensure_worker(printer_name)

        logger.info(f"[QUEUE] Job {job_id[:8]} queued for {printer_name} (type={job_type})")
        return job_id

    def get_job(self, job_id: str) -> dict | None:
        """Get a job's status (without the execute_fn)."""
        with self._lock:
            job = self._jobs.get(job_id)
            if job is None:
                for h in self._history:
                    if h["id"] == job_id:
                        return h
                return None
            return self._sanitize_job(job)

    def get_queue(self, printer_name: str = None) -> list[dict]:
        """Get all active + recent jobs, optionally filtered by printer."""
        with self._lock:
            active = [
                self._sanitize_job(j) for j in self._jobs.values()
                if printer_name is None or j["printer"] == printer_name
            ]
            history = [
                h for h in self._history
                if printer_name is None or h["printer"] == printer_name
            ]
        return active + list(history)

    def cancel_job(self, job_id: str) -> bool:
        """Cancel a pending job. Returns True if cancelled, False otherwise."""
        with self._lock:
            job = self._jobs.get(job_id)
            if job and job["status"] == self.STATUS_PENDING:
                job["status"] = self.STATUS_CANCELLED
                job["completed_at"] = datetime.now().isoformat()
                self._move_to_history(job_id)
                logger.info(f"[QUEUE] Job {job_id[:8]} cancelled")
                return True
        return False

    def get_status(self) -> dict:
        """Get overall queue health."""
        with self._lock:
            printers = {}
            for pname, q in self._queues.items():
                pending = sum(
                    1 for j in self._jobs.values()
                    if j["printer"] == pname and j["status"] == self.STATUS_PENDING
                )
                printing = sum(
                    1 for j in self._jobs.values()
                    if j["printer"] == pname and j["status"] == self.STATUS_PRINTING
                )
                worker_alive = (
                    pname in self._workers and self._workers[pname].is_alive()
                )
                printers[pname] = {
                    "pending": pending,
                    "printing": printing,
                    "worker_alive": worker_alive,
                }
            total_active = len(self._jobs)
            total_history = len(self._history)

        return {
            "printers": printers,
            "active_jobs": total_active,
            "history_size": total_history,
            "max_retries": self.max_retries,
        }

    def shutdown(self):
        """Signal all workers to stop."""
        self._shutdown.set()

    def _ensure_worker(self, printer_name: str):
        """Start a worker thread for this printer if not already running."""
        if printer_name in self._workers and self._workers[printer_name].is_alive():
            return
        t = threading.Thread(
            target=self._worker_loop,
            args=(printer_name,),
            daemon=True,
            name=f"print-worker-{printer_name}",
        )
        self._workers[printer_name] = t
        t.start()
        logger.info(f"[QUEUE] Worker started for {printer_name}")

    def _worker_loop(self, printer_name: str):
        """Pull jobs from this printer's queue and execute them sequentially."""
        q = self._queues[printer_name]
        while not self._shutdown.is_set():
            try:
                job_id = q.get(timeout=5)
            except Empty:
                with self._lock:
                    if q.empty() and not any(
                        j["status"] in (self.STATUS_PENDING, self.STATUS_PRINTING, self.STATUS_RETRYING)
                        for j in self._jobs.values()
                        if j["printer"] == printer_name
                    ):
                        break
                continue

            with self._lock:
                job = self._jobs.get(job_id)
                if not job or job["status"] == self.STATUS_CANCELLED:
                    q.task_done()
                    continue
                job["status"] = self.STATUS_PRINTING
                job["started_at"] = datetime.now().isoformat()

            execute_fn = job["execute_fn"]
            try:
                execute_fn()
                with self._lock:
                    job["status"] = self.STATUS_DONE
                    job["completed_at"] = datetime.now().isoformat()
                    job["error"] = None
                    self._move_to_history(job_id)
                logger.info(f"[QUEUE] Job {job_id[:8]} completed on {printer_name}")
            except Exception as e:
                retry_count = job["retries"] + 1
                if retry_count <= self.max_retries:
                    delay = self.retry_base_delay * (2 ** (retry_count - 1))
                    with self._lock:
                        job["status"] = self.STATUS_RETRYING
                        job["retries"] = retry_count
                        job["error"] = str(e)
                    logger.warning(
                        f"[QUEUE] Job {job_id[:8]} failed (attempt {retry_count}/{self.max_retries}), "
                        f"retrying in {delay:.1f}s: {e}"
                    )
                    time.sleep(delay)
                    with self._lock:
                        job["status"] = self.STATUS_PENDING
                    q.put(job_id)
                else:
                    with self._lock:
                        job["status"] = self.STATUS_FAILED
                        job["completed_at"] = datetime.now().isoformat()
                        job["error"] = str(e)
                        self._move_to_history(job_id)
                    logger.error(
                        f"[QUEUE] Job {job_id[:8]} failed permanently after "
                        f"{self.max_retries} retries: {e}"
                    )
            finally:
                q.task_done()

        logger.info(f"[QUEUE] Worker for {printer_name} exiting (queue empty)")

    def _move_to_history(self, job_id: str):
        """Move a finished job from active to history. Must hold _lock."""
        job = self._jobs.pop(job_id, None)
        if job:
            sanitized = self._sanitize_job(job)
            self._history.appendleft(sanitized)

    @staticmethod
    def _sanitize_job(job: dict) -> dict:
        """Return a copy without the execute_fn (not serializable)."""
        return {k: v for k, v in job.items() if k != "execute_fn"}

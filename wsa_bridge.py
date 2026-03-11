"""
WSA Bridge (adb reverse)
=========================
Sets up adb reverse port forwarding so that apps running inside
Windows Subsystem for Android (WSA) can reach the printer server
on the host via localhost.

How it works:
  1. Connects to WSA's ADB interface (127.0.0.1:58526 by default)
  2. Sets up `adb reverse tcp:3006 tcp:3006`
  3. Now WSA apps can reach the printer server at localhost:3006

This solves the problem where WSA can access public internet
but cannot see the local network (Hyper-V NAT isolation).
"""

import subprocess
import shutil
import logging
import time
import threading
import os
import platform

logger = logging.getLogger(__name__)

_SUBPROCESS_FLAGS = (
    subprocess.CREATE_NO_WINDOW
    if platform.system() == "Windows"
    else 0
)


class WSABridge:
    """Manage adb reverse port forwarding for WSA."""

    def __init__(self, adb_port=58526, server_port=3006, auto_reconnect=True):
        self.adb_port = adb_port
        self.server_port = server_port
        self.auto_reconnect = auto_reconnect
        self.connected = False
        self._monitor_thread = None
        self._stop_event = threading.Event()

    def setup(self):
        """Set up the adb reverse tunnel."""
        adb_path = self._find_adb()
        if not adb_path:
            logger.warning(
                "ADB not found. WSA bridge disabled.\n"
                "  To enable: Install Android SDK Platform Tools and add to PATH.\n"
                "  Download: https://developer.android.com/tools/releases/platform-tools"
            )
            return False

        logger.info("Setting up WSA bridge (adb reverse)...")

        # Step 1: Connect to WSA's ADB
        if not self._adb_connect(adb_path):
            logger.warning(
                f"Could not connect to WSA ADB on 127.0.0.1:{self.adb_port}.\n"
                "  Make sure WSA is running and Developer Mode is enabled:\n"
                "  WSA Settings > Developer > Enable Developer Mode"
            )
            return False

        # Step 2: Set up reverse port forwarding
        if not self._adb_reverse(adb_path):
            logger.warning("Failed to set up adb reverse. WSA bridge disabled.")
            return False

        self.connected = True
        logger.info(
            f"WSA bridge active! WSA apps can reach printer server at "
            f"localhost:{self.server_port}"
        )

        # Step 3: Start auto-reconnect monitor (WSA can restart)
        if self.auto_reconnect:
            self._start_monitor(adb_path)

        return True

    def teardown(self):
        """Remove the adb reverse tunnel."""
        self._stop_event.set()
        adb_path = self._find_adb()
        if adb_path:
            try:
                subprocess.run(
                    [adb_path, "reverse", "--remove", f"tcp:{self.server_port}"],
                    capture_output=True,
                    timeout=5,
                    creationflags=_SUBPROCESS_FLAGS,
                )
                logger.info("WSA bridge removed.")
            except Exception:
                pass
        self.connected = False

    def _find_adb(self):
        """Find the adb executable."""
        # Check PATH
        adb = shutil.which("adb")
        if adb:
            return adb

        # Common Windows locations
        common_paths = [
            os.path.expandvars(r"%LOCALAPPDATA%\Android\Sdk\platform-tools\adb.exe"),
            os.path.expandvars(r"%USERPROFILE%\AppData\Local\Android\Sdk\platform-tools\adb.exe"),
            r"C:\platform-tools\adb.exe",
            r"C:\Android\platform-tools\adb.exe",
        ]

        for path in common_paths:
            if os.path.isfile(path):
                return path

        return None

    def _adb_connect(self, adb_path):
        """Connect to WSA's ADB server."""
        target = f"127.0.0.1:{self.adb_port}"

        try:
            result = subprocess.run(
                [adb_path, "connect", target],
                capture_output=True,
                text=True,
                timeout=10,
                creationflags=_SUBPROCESS_FLAGS,
            )

            output = result.stdout.strip() + result.stderr.strip()
            logger.debug(f"adb connect output: {output}")

            # "connected to" or "already connected" means success
            if "connected" in output.lower():
                logger.info(f"  Connected to WSA ADB at {target}")
                return True
            elif "refused" in output.lower():
                logger.debug(f"  ADB connection refused at {target}")
                return False
            else:
                logger.debug(f"  ADB connect unexpected: {output}")
                return False

        except subprocess.TimeoutExpired:
            logger.debug(f"  ADB connect timed out for {target}")
            return False
        except FileNotFoundError:
            return False

    def _adb_reverse(self, adb_path):
        """Set up adb reverse tcp forwarding."""
        try:
            # Remove existing reverse if any
            subprocess.run(
                [adb_path, "reverse", "--remove-all"],
                capture_output=True,
                timeout=5,
                creationflags=_SUBPROCESS_FLAGS,
            )

            # Set up new reverse
            result = subprocess.run(
                [adb_path, "reverse",
                 f"tcp:{self.server_port}",
                 f"tcp:{self.server_port}"],
                capture_output=True,
                text=True,
                timeout=10,
                creationflags=_SUBPROCESS_FLAGS,
            )

            if result.returncode == 0:
                logger.info(
                    f"  adb reverse tcp:{self.server_port} -> "
                    f"tcp:{self.server_port} (OK)"
                )
                return True
            else:
                logger.error(f"  adb reverse failed: {result.stderr.strip()}")
                return False

        except subprocess.TimeoutExpired:
            logger.error("  adb reverse timed out")
            return False
        except FileNotFoundError:
            return False

    def _start_monitor(self, adb_path):
        """Start a background thread to monitor and re-establish the tunnel."""
        self._stop_event.clear()
        self._monitor_thread = threading.Thread(
            target=self._monitor_loop,
            args=(adb_path,),
            daemon=True,
            name="wsa-bridge-monitor",
        )
        self._monitor_thread.start()

    def _monitor_loop(self, adb_path):
        """Periodically check if the adb reverse tunnel is still alive."""
        while not self._stop_event.is_set():
            self._stop_event.wait(30)  # Check every 30 seconds
            if self._stop_event.is_set():
                break

            try:
                result = subprocess.run(
                    [adb_path, "reverse", "--list"],
                    capture_output=True,
                    text=True,
                    timeout=5,
                    creationflags=_SUBPROCESS_FLAGS,
                )
                if f"tcp:{self.server_port}" not in result.stdout:
                    logger.warning("WSA bridge tunnel lost. Reconnecting...")
                    if self._adb_connect(adb_path) and self._adb_reverse(adb_path):
                        logger.info("WSA bridge reconnected.")
                        self.connected = True
                    else:
                        self.connected = False
            except Exception:
                pass


# ─── Standalone test ─────────────────────────────────────────
if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG, format="%(message)s")
    bridge = WSABridge(server_port=3006)
    if bridge.setup():
        print("\nWSA bridge is active. Press Ctrl+C to stop.")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            bridge.teardown()
    else:
        print("\nWSA bridge setup failed. Check the logs above.")

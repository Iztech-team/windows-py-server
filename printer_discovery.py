"""
Network Printer Discovery
==========================
Scans the local network for thermal printers on port 9100.
Uses parallel TCP connection attempts + multiple strategies
to resolve the real MAC address of each printer:

  1. SNMP query (raw UDP, no pysnmp dependency)
  2. Gateway-aware ARP filtering
  3. IP-based unique fallback ID

Works on Windows, Linux, and macOS.
"""

import socket
import subprocess
import re
import logging
import platform
from concurrent.futures import ThreadPoolExecutor, as_completed

logger = logging.getLogger(__name__)

# Prevent console windows from flashing when running under pythonw.exe
_SUBPROCESS_FLAGS = (
    subprocess.CREATE_NO_WINDOW
    if platform.system() == "Windows"
    else 0
)


class PrinterDiscovery:
    """Discover network printers by scanning for open port 9100."""

    PRINTER_PORT = 9100
    SCAN_TIMEOUT = 3  # seconds per host
    MAX_WORKERS = 50  # parallel scan threads

    def __init__(self, port=9100, timeout=3):
        self.PRINTER_PORT = port
        self.SCAN_TIMEOUT = timeout
        self._gateway_mac = None

    def scan(self, subnet=None):
        """
        Scan the local network for printers.
        Returns dict: { ip: { port, mac, responsive } }
        """
        if subnet is None:
            subnet = self._detect_subnet()

        if not subnet:
            logger.error("Could not detect network subnet.")
            return {}

        logger.info(f"Scanning {subnet}.0/24 on port {self.PRINTER_PORT}...")

        self._gateway_mac = self._detect_gateway_mac(subnet)
        if self._gateway_mac:
            logger.info(f"Gateway MAC detected: {self._gateway_mac} (will be ignored for printers)")

        targets = [f"{subnet}.{i}" for i in range(1, 255)]

        found = {}
        with ThreadPoolExecutor(max_workers=self.MAX_WORKERS) as executor:
            futures = {
                executor.submit(self._check_port, ip, self.PRINTER_PORT): ip
                for ip in targets
            }
            for future in as_completed(futures):
                ip = futures[future]
                try:
                    is_open = future.result()
                    if is_open:
                        mac = self._get_real_mac(ip)
                        found[ip] = {
                            "port": self.PRINTER_PORT,
                            "mac": mac,
                            "responsive": True,
                        }
                        logger.info(f"  Found printer: {ip} (MAC: {mac or 'unknown'})")
                except Exception as e:
                    logger.debug(f"  Error scanning {ip}: {e}")

        logger.info(f"Scan complete. Found {len(found)} printer(s).")
        return found

    def check_printer(self, ip, port=None):
        """Check if a specific printer is reachable."""
        port = port or self.PRINTER_PORT
        return self._check_port(ip, port)

    def _check_port(self, ip, port):
        """Try to connect to ip:port. Returns True if open."""
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(self.SCAN_TIMEOUT)
            result = sock.connect_ex((ip, port))
            sock.close()
            return result == 0
        except (socket.timeout, OSError):
            return False

    def _detect_subnet(self):
        """Detect the local subnet (e.g., '192.168.1')."""
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            local_ip = s.getsockname()[0]
            s.close()

            parts = local_ip.split(".")
            if len(parts) == 4:
                subnet = ".".join(parts[:3])
                logger.info(f"Local IP: {local_ip}, Subnet: {subnet}.0/24")
                return subnet
        except Exception as e:
            logger.error(f"Failed to detect subnet: {e}")

        return None

    # ─── Gateway detection ────────────────────────────────────

    def _detect_gateway_mac(self, subnet):
        """Find the default gateway's MAC so we can ignore it on printer entries."""
        gateway_ip = self._get_gateway_ip(subnet)
        if not gateway_ip:
            return None
        logger.info(f"Default gateway IP: {gateway_ip}")
        return self._get_arp_mac(gateway_ip)

    def _get_gateway_ip(self, subnet):
        """Get the default gateway IP address."""
        system = platform.system().lower()
        try:
            if system == "windows":
                result = subprocess.run(
                    ["ipconfig"],
                    capture_output=True, text=True, timeout=10,
                    creationflags=_SUBPROCESS_FLAGS,
                )
                for line in result.stdout.splitlines():
                    if "default gateway" in line.lower() or "passerelle par" in line.lower():
                        match = re.search(r"(\d+\.\d+\.\d+\.\d+)", line)
                        if match and match.group(1).startswith(subnet.rsplit(".", 2)[0]):
                            return match.group(1)
            else:
                result = subprocess.run(
                    ["ip", "route", "show", "default"],
                    capture_output=True, text=True, timeout=5,
                    creationflags=_SUBPROCESS_FLAGS,
                )
                match = re.search(r"via\s+(\d+\.\d+\.\d+\.\d+)", result.stdout)
                if match:
                    return match.group(1)
        except Exception as e:
            logger.debug(f"Could not detect gateway: {e}")

        # Fallback: assume .1 or .254
        for last_octet in ("1", "254"):
            candidate = f"{subnet}.{last_octet}"
            arp_mac = self._get_arp_mac(candidate)
            if arp_mac:
                return candidate
        return None

    # ─── Real MAC resolution (multi-strategy) ─────────────────

    def _get_real_mac(self, ip):
        """
        Try multiple methods to get the printer's *real* MAC address:
          1. HTTP web interface (JK-E02 and similar ethernet modules)
          2. SNMP query for ifPhysAddress
          3. ARP cache, but only if it differs from the gateway MAC
          4. None if all methods fail
        """
        # Strategy 1: HTTP web interface (most reliable for JK-E02 modules)
        http_mac = self._get_mac_http(ip)
        if http_mac:
            logger.info(f"  {ip}: got real MAC via HTTP web interface: {http_mac}")
            return http_mac

        # Strategy 2: SNMP
        snmp_mac = self._get_mac_snmp(ip)
        if snmp_mac:
            logger.info(f"  {ip}: got real MAC via SNMP: {snmp_mac}")
            return snmp_mac

        # Strategy 3: ARP with gateway filtering
        arp_mac = self._get_arp_mac(ip)
        if arp_mac and arp_mac != self._gateway_mac:
            logger.debug(f"  {ip}: got MAC via ARP (not gateway): {arp_mac}")
            return arp_mac

        if arp_mac and arp_mac == self._gateway_mac:
            logger.warning(
                f"  {ip}: ARP returned gateway MAC {arp_mac} — "
                f"printer is likely behind a secondary router. MAC will be unknown."
            )

        return None

    # ─── HTTP web interface MAC resolution ────────────────────

    def _get_mac_http(self, ip):
        """
        Many thermal printer ethernet modules (JK-E02, etc.) expose a web
        interface on port 80. The status page contains the real MAC address
        in decimal-separated format (e.g. "168-1-87-59-209-132").
        """
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(3)
            result = sock.connect_ex((ip, 80))
            if result != 0:
                sock.close()
                return None

            sock.sendall(
                b"GET /port_stat.htm HTTP/1.0\r\nHost: "
                + ip.encode()
                + b"\r\n\r\n"
            )

            data = b""
            while True:
                try:
                    chunk = sock.recv(4096)
                    if not chunk:
                        break
                    data += chunk
                except socket.timeout:
                    break
            sock.close()

            text = data.decode("ascii", errors="replace")
            return self._parse_http_mac(text)
        except (socket.timeout, OSError) as e:
            logger.debug(f"  HTTP MAC query to {ip} failed: {e}")
        return None

    @staticmethod
    def _parse_http_mac(html):
        """
        Parse MAC address from printer web interface HTML.
        Supports two formats:
          - Decimal: "168-1-87-59-209-132" (JK-E02 style)
          - Hex:     "A8:01:57:3B:D1:84" or "A8-01-57-3B-D1-84"
        """
        # Format 1: Standard hex MAC (XX:XX:XX:XX:XX:XX or XX-XX-XX-XX-XX-XX)
        hex_match = re.search(
            r"[Mm]ac.*?([0-9a-fA-F]{2}[:\-][0-9a-fA-F]{2}[:\-]"
            r"[0-9a-fA-F]{2}[:\-][0-9a-fA-F]{2}[:\-]"
            r"[0-9a-fA-F]{2}[:\-][0-9a-fA-F]{2})",
            html,
        )
        if hex_match:
            return hex_match.group(1).upper().replace("-", ":")

        # Format 2: Decimal-separated (JK-E02 modules: "168-1-87-59-209-132")
        dec_match = re.search(
            r"[Mm]ac.*?Address.*?(\d{1,3})-(\d{1,3})-(\d{1,3})"
            r"-(\d{1,3})-(\d{1,3})-(\d{1,3})",
            html,
            re.DOTALL,
        )
        if dec_match:
            octets = [int(dec_match.group(i)) for i in range(1, 7)]
            if all(0 <= o <= 255 for o in octets):
                return ":".join(f"{o:02X}" for o in octets)

        return None

    # ─── SNMP MAC resolution (raw UDP, no dependencies) ───────

    def _get_mac_snmp(self, ip, community=b"public"):
        """
        Send a raw SNMPv1 GET-NEXT request for the interface table
        (OID 1.3.6.1.2.1.2.2.1.6 = ifPhysAddress) and parse the
        first MAC-address response. No pysnmp needed.
        """
        try:
            oid = [1, 3, 6, 1, 2, 1, 2, 2, 1, 6]
            packet = self._build_snmp_getnext(community, oid)

            sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            sock.settimeout(2)
            sock.sendto(packet, (ip, 161))

            data, _ = sock.recvfrom(4096)
            sock.close()

            mac_bytes = self._parse_snmp_mac_response(data)
            if mac_bytes and len(mac_bytes) == 6:
                mac_str = ":".join(f"{b:02X}" for b in mac_bytes)
                if mac_str != "00:00:00:00:00:00":
                    return mac_str
        except (socket.timeout, OSError):
            pass
        except Exception as e:
            logger.debug(f"  SNMP query to {ip} failed: {e}")
        return None

    @staticmethod
    def _encode_oid(oid):
        """BER-encode an OID value (without tag+length wrapper)."""
        if len(oid) < 2:
            return bytes([0])
        result = bytes([40 * oid[0] + oid[1]])
        for sub_id in oid[2:]:
            if sub_id < 0x80:
                result += bytes([sub_id])
            else:
                # multi-byte encoding for sub-identifiers >= 128
                pieces = []
                val = sub_id
                pieces.append(val & 0x7F)
                val >>= 7
                while val > 0:
                    pieces.append(0x80 | (val & 0x7F))
                    val >>= 7
                result += bytes(reversed(pieces))
        return result

    @staticmethod
    def _encode_length(length):
        """BER-encode a length field."""
        if length < 0x80:
            return bytes([length])
        elif length < 0x100:
            return bytes([0x81, length])
        else:
            return bytes([0x82, (length >> 8) & 0xFF, length & 0xFF])

    @staticmethod
    def _encode_tlv(tag, value):
        """Wrap value bytes in a TLV (tag-length-value) structure."""
        length = PrinterDiscovery._encode_length(len(value))
        return bytes([tag]) + length + value

    def _build_snmp_getnext(self, community, oid):
        """Build a raw SNMPv1 GET-NEXT request packet."""
        # Integer: version = 0 (SNMPv1)
        version = self._encode_tlv(0x02, b'\x00')
        # Octet string: community
        comm = self._encode_tlv(0x04, community)
        # OID
        oid_val = self._encode_oid(oid)
        oid_tlv = self._encode_tlv(0x06, oid_val)
        # NULL value
        null_val = self._encode_tlv(0x05, b'')
        # Varbind: SEQUENCE { oid, null }
        varbind = self._encode_tlv(0x30, oid_tlv + null_val)
        # VarbindList: SEQUENCE { varbind }
        varbind_list = self._encode_tlv(0x30, varbind)
        # Request-ID
        req_id = self._encode_tlv(0x02, b'\x01')
        # Error status
        error_status = self._encode_tlv(0x02, b'\x00')
        # Error index
        error_index = self._encode_tlv(0x02, b'\x00')
        # GET-NEXT PDU (tag 0xA1)
        pdu = self._encode_tlv(0xA1, req_id + error_status + error_index + varbind_list)
        # Full SNMP message
        return self._encode_tlv(0x30, version + comm + pdu)

    @staticmethod
    def _parse_snmp_mac_response(data):
        """
        Walk the BER-encoded SNMP response and extract the first
        OCTET STRING value that is 6 bytes long (a MAC address).
        """
        try:
            i = 0
            while i < len(data) - 7:
                # Look for OCTET STRING (tag 0x04) with length 6
                if data[i] == 0x04 and data[i + 1] == 0x06:
                    return data[i + 2 : i + 8]
                i += 1
        except Exception:
            pass
        return None

    # ─── ARP MAC resolution ───────────────────────────────────

    def _get_arp_mac(self, ip):
        """Get MAC address from ARP cache (ping first to populate)."""
        system = platform.system().lower()

        try:
            if system == "windows":
                subprocess.run(
                    ["ping", "-n", "1", "-w", "1000", ip],
                    capture_output=True, timeout=5,
                    creationflags=_SUBPROCESS_FLAGS,
                )
            else:
                subprocess.run(
                    ["ping", "-c", "1", "-W", "2", ip],
                    capture_output=True, timeout=5,
                    creationflags=_SUBPROCESS_FLAGS,
                )
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pass

        try:
            if system == "windows":
                return self._get_mac_windows(ip)
            else:
                return self._get_mac_linux(ip)
        except Exception as e:
            logger.debug(f"Could not get ARP MAC for {ip}: {e}")
            return None

    def _get_mac_windows(self, ip):
        """Get MAC from Windows ARP cache."""
        try:
            result = subprocess.run(
                ["arp", "-a", ip],
                capture_output=True, text=True, timeout=5,
                creationflags=_SUBPROCESS_FLAGS,
            )
            for line in result.stdout.splitlines():
                if ip in line:
                    match = re.search(
                        r"([0-9a-fA-F]{2}[:-]){5}[0-9a-fA-F]{2}", line
                    )
                    if match:
                        return match.group(0).upper().replace("-", ":")
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pass
        return None

    def _get_mac_linux(self, ip):
        """Get MAC from Linux ARP cache."""
        try:
            result = subprocess.run(
                ["ip", "neighbor", "show", ip],
                capture_output=True, text=True, timeout=5,
                creationflags=_SUBPROCESS_FLAGS,
            )
            match = re.search(
                r"([0-9a-fA-F]{2}:){5}[0-9a-fA-F]{2}", result.stdout
            )
            if match:
                return match.group(0).upper()
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pass

        try:
            result = subprocess.run(
                ["arp", "-n", ip],
                capture_output=True, text=True, timeout=5,
                creationflags=_SUBPROCESS_FLAGS,
            )
            for line in result.stdout.splitlines():
                if ip in line:
                    match = re.search(
                        r"([0-9a-fA-F]{2}:){5}[0-9a-fA-F]{2}", line
                    )
                    if match:
                        return match.group(0).upper()
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pass

        return None


# ─── Standalone test ─────────────────────────────────────────
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    d = PrinterDiscovery()
    printers = d.scan()
    if printers:
        print(f"\nFound {len(printers)} printer(s):")
        for ip, info in printers.items():
            print(f"  {ip}:{info['port']} (MAC: {info.get('mac', 'unknown')})")
    else:
        print("\nNo printers found on this network.")

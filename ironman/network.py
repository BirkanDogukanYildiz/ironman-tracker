# -*- coding: utf-8 -*-
"""Yerel ağ adresi tespiti — telefondan bağlanmak için.

Bir bilgisayarda birden çok ağ arayüzü olabilir (Wi-Fi, Ethernet, WSL, Docker,
VirtualBox, VPN). Yanlış adresi göstermek telefondan bağlanamamanın en yaygın
sebebidir; bu yüzden hepsini bulup en olası olanı öne alıyoruz.
"""
from __future__ import annotations

import ipaddress
import socket
from dataclasses import dataclass


# Sanal makine / konteyner arayüzlerinin tipik aralıkları — bunlar telefondan
# erişilemez, listenin sonuna atılırlar.
VIRTUAL_HINTS = (
    ipaddress.ip_network("172.17.0.0/16"),   # Docker
    ipaddress.ip_network("172.18.0.0/16"),
    ipaddress.ip_network("192.168.56.0/24"), # VirtualBox host-only
    ipaddress.ip_network("192.168.99.0/24"), # Docker Machine
    ipaddress.ip_network("10.0.75.0/24"),    # Hyper-V / Docker Desktop
    ipaddress.ip_network("172.16.0.0/16"),   # sık kullanılan VM aralığı
)


@dataclass
class Address:
    ip: str
    likely: bool = True
    note: str = ""

    @property
    def is_private(self) -> bool:
        try:
            return ipaddress.ip_address(self.ip).is_private
        except ValueError:
            return False


def _primary_ip() -> str | None:
    """İnternete çıkarken kullanılan arayüzün adresi (paket göndermez)."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("10.255.255.255", 1))
        return s.getsockname()[0]
    except OSError:
        return None
    finally:
        s.close()


def _all_ips() -> list[str]:
    out: list[str] = []
    try:
        host = socket.gethostname()
        for info in socket.getaddrinfo(host, None, socket.AF_INET):
            ip = info[4][0]
            if ip not in out:
                out.append(ip)
    except (socket.gaierror, OSError):
        pass
    return out


def _is_virtual(ip: str) -> bool:
    try:
        addr = ipaddress.ip_address(ip)
    except ValueError:
        return True
    return any(addr in net for net in VIRTUAL_HINTS)


def lan_addresses() -> list[Address]:
    """Telefondan denenebilecek adresler — en olası olan başta."""
    found: list[str] = []
    primary = _primary_ip()
    if primary:
        found.append(primary)
    for ip in _all_ips():
        if ip not in found:
            found.append(ip)

    out: list[Address] = []
    for ip in found:
        if ip.startswith("127."):
            continue
        try:
            if not ipaddress.ip_address(ip).is_private:
                continue
        except ValueError:
            continue
        virtual = _is_virtual(ip)
        out.append(Address(
            ip=ip,
            likely=not virtual,
            note="sanal ağ olabilir (Docker / VM / WSL)" if virtual else "",
        ))
    out.sort(key=lambda a: (not a.likely, a.ip != primary))
    return out


def best_address() -> str:
    addrs = lan_addresses()
    return addrs[0].ip if addrs else "127.0.0.1"


def phone_url(port: int = 5000, ip: str | None = None) -> str:
    return f"http://{ip or best_address()}:{port}"


def qr_svg(data: str, scale: int = 7) -> str | None:
    """QR kodu SVG olarak döner. `segno` kurulu değilse None."""
    try:
        import segno
    except ImportError:
        return None
    import io
    qr = segno.make(data, error="m")
    buf = io.BytesIO()
    qr.save(buf, kind="svg", scale=scale, border=2,
            dark="#0b1220", light=None, svgclass="qr", lineclass=None,
            omitsize=True, xmldecl=False, svgns=True)
    return buf.getvalue().decode("utf-8")

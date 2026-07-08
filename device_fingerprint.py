"""
设备指纹生成器
跨平台（Windows/Linux）设备唯一标识生成工具

该模块提供设备指纹生成功能，通过收集 CPU ID、硬盘序列号、MAC 地址等
硬件信息，生成唯一的设备标识符（SHA256 哈希字符串）。
"""

import hashlib
import platform
import subprocess
import uuid
from typing import Optional


def get_cpu_id() -> str:
    """
    获取 CPU ID。

    Returns:
        str: CPU ID 字符串，如果获取失败则返回空字符串。
    """
    system = platform.system()

    if system == "Windows":
        try:
            # Windows: 使用 wmic 获取 ProcessorId
            result = subprocess.run(
                ["wmic", "cpu", "get", "ProcessorId"],
                capture_output=True,
                text=True,
                timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW  # type: ignore
            )
            lines = result.stdout.strip().split("\n")
            if len(lines) > 1:
                return lines[1].strip()
        except Exception:
            pass

    elif system == "Linux":
        try:
            # Linux: 读取 /proc/cpuinfo 获取序列号
            with open("/proc/cpuinfo", "r") as f:
                for line in f:
                    if line.startswith("serial"):
                        cpu_serial = line.split(":")[1].strip()
                        if cpu_serial and cpu_serial != "0":
                            return cpu_serial
            # 备用方案：使用 dmidecode
            result = subprocess.run(
                ["dmidecode", "-t", "processor"],
                capture_output=True,
                text=True,
                timeout=10
            )
            for line in result.stdout.split("\n"):
                if "ID" in line and "Socket" not in line:
                    parts = line.split(":")
                    if len(parts) > 1:
                        return parts[1].strip()
        except Exception:
            pass

    return ""


def get_disk_serial() -> str:
    """
    获取硬盘序列号。

    Returns:
        str: 硬盘序列号字符串，如果获取失败则返回空字符串。
    """
    system = platform.system()

    if system == "Windows":
        try:
            # Windows: 使用 wmic 获取硬盘序列号
            result = subprocess.run(
                ["wmic", "diskdrive", "get", "serialnumber"],
                capture_output=True,
                text=True,
                timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW  # type: ignore
            )
            lines = result.stdout.strip().split("\n")
            for line in lines[1:]:
                serial = line.strip()
                if serial:
                    return serial
        except Exception:
            pass

    elif system == "Linux":
        try:
            # Linux: 尝试读取 /dev/sda 或 /dev/nvme0n1 的序列号
            for disk in ["/dev/sda", "/dev/nvme0n1", "/dev/vda"]:
                try:
                    result = subprocess.run(
                        ["hdparm", "-I", disk],
                        capture_output=True,
                        text=True,
                        timeout=10
                    )
                    for line in result.stdout.split("\n"):
                        if "Serial Number" in line:
                            return line.split(":")[1].strip()
                except Exception:
                    continue

            # 备用方案：读取 sysfs
            for disk_path in ["/sys/block/sda", "/sys/block/nvme0n1"]:
                try:
                    serial_file = f"{disk_path}/serial"
                    with open(serial_file, "r") as f:
                        serial = f.read().strip()
                        if serial:
                            return serial
                except Exception:
                    continue
        except Exception:
            pass

    return ""


def get_mac_address() -> str:
    """
    获取 MAC 地址。

    Returns:
        str: MAC 地址字符串（格式：XX:XX:XX:XX:XX:XX），如果获取失败则返回空字符串。
    """
    try:
        # 获取所有网络接口的 MAC 地址
        mac = uuid.getnode()
        if mac:
            return ":".join(f"{(mac >> elements) & 0xFF:02x}"
                           for elements in range(0, 48, 8))[::-1]
    except Exception:
        pass

    system = platform.system()

    if system == "Windows":
        try:
            result = subprocess.run(
                ["getmac", "/fo", "csv", "/nh"],
                capture_output=True,
                text=True,
                timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW  # type: ignore
            )
            for line in result.stdout.strip().split("\n"):
                parts = line.split(",")
                for part in parts:
                    mac_addr = part.strip().replace("-", ":")
                    if len(mac_addr) == 17 and ":" in mac_addr:
                        return mac_addr.upper()
        except Exception:
            pass

    elif system == "Linux":
        try:
            # 尝试读取 /sys/class/net 下的 MAC 地址
            import os
            net_path = "/sys/class/net"
            if os.path.exists(net_path):
                for iface in os.listdir(net_path):
                    if iface != "lo":  # 跳过回环接口
                        try:
                            with open(f"{net_path}/{iface}/address", "r") as f:
                                mac = f.read().strip()
                                if mac and mac != "00:00:00:00:00:00":
                                    return mac.upper()
                        except Exception:
                            continue
        except Exception:
            pass

        try:
            # 备用方案：使用 ip 命令
            result = subprocess.run(
                ["ip", "link"],
                capture_output=True,
                text=True,
                timeout=10
            )
            for line in result.stdout.split("\n"):
                if "link/ether" in line:
                    parts = line.strip().split()
                    for i, part in enumerate(parts):
                        if part == "link/ether" and i + 1 < len(parts):
                            return parts[i + 1].upper()
        except Exception:
            pass

    return ""


def generate_device_fingerprint() -> str:
    """
    生成设备指纹。

    通过收集 CPU ID、硬盘序列号、MAC 地址等硬件信息，
    生成唯一的设备标识符（SHA256 哈希字符串）。

    该函数跨平台兼容 Windows 和 Linux 系统。如果某些硬件信息无法获取，
    将使用可用信息组合生成指纹。如果所有硬件信息都获取失败，将使用
    UUID 作为后备方案。

    Returns:
        str: SHA256 哈希字符串（64 个十六进制字符），作为设备的唯一指纹。

    Example:
        >>> fingerprint = generate_device_fingerprint()
        >>> print(f"Device fingerprint: {fingerprint}")
        >>> print(f"Length: {len(fingerprint)}")  # 64
    """
    # 收集硬件信息
    cpu_id = get_cpu_id()
    disk_serial = get_disk_serial()
    mac_address = get_mac_address()

    # 组合所有硬件信息
    hardware_info = f"{cpu_id}|{disk_serial}|{mac_address}"

    # 如果所有硬件信息都为空，使用 UUID 作为后备
    if not hardware_info.strip("|"):
        hardware_info = str(uuid.uuid4())

    # 生成 SHA256 哈希
    fingerprint = hashlib.sha256(hardware_info.encode("utf-8")).hexdigest()

    return fingerprint


if __name__ == "__main__":
    # 测试代码
    print("设备指纹生成器测试")
    print("=" * 50)

    print(f"\nCPU ID: {get_cpu_id()}")
    print(f"硬盘序列号：{get_disk_serial()}")
    print(f"MAC 地址：{get_mac_address()}")

    fingerprint = generate_device_fingerprint()
    print(f"\n设备指纹：{fingerprint}")
    print(f"指纹长度：{len(fingerprint)} 字符")

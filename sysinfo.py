#!/usr/bin/env python3

# Imports {{{
# builtins
import subprocess
import json
import sys
import os
import pathlib
from platform import system
from dataclasses import dataclass
from typing import Any, Callable, Collection, List, Optional

# 3rd party
from PySide2.QtWidgets import (
    QApplication,
    QHBoxLayout,
    QVBoxLayout,
    QLabel,
    QWidget,
    QFrame,
)
from PySide2.QtGui import (
    QPixmap,
    QColor,
)
from PySide2.QtCore import Qt
import qrcode
import PIL.ImageQt
import pint
import psutil
from cpuinfo import get_cpu_info

# local modules

# }}}


units = pint.UnitRegistry()


def subprocess_args(include_stdout=True) -> dict:
    """
    Create a set of arguments which make a ``subprocess.Popen`` (and
    variants) call work with or without Pyinstaller, ``--noconsole`` or
    not, on Windows and Linux. Typical use::

      subprocess.call(['program_to_run', 'arg_1'], **subprocess_args())

    When calling ``check_output``::

      subprocess.check_output(['program_to_run', 'arg_1'],
                              **subprocess_args(False))
    """
    # The following is true only on Windows.
    if hasattr(subprocess, "STARTUPINFO"):
        # On Windows, subprocess calls will pop up a command window by default
        # when run from Pyinstaller with the ``--noconsole`` option. Avoid this
        # distraction.
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        # Windows doesn't search the path by default. Pass it an environment so
        # it will.
        env = os.environ
    else:
        si = None
        env = None

    # ``subprocess.check_output`` doesn't allow specifying ``stdout``::
    #
    #   Traceback (most recent call last):
    #     File "test_subprocess.py", line 58, in <module>
    #       **subprocess_args(stdout=None))
    #     File "C:\Python27\lib\subprocess.py", line 567, in check_output
    #       raise ValueError('stdout argument not allowed, it will be overridden.')
    #   ValueError: stdout argument not allowed, it will be overridden.
    #
    # So, add it only if it's needed.
    if include_stdout:
        ret = {"stdout": subprocess.PIPE}
    else:
        ret = {}

    # On Windows, running this from the binary produced by Pyinstaller
    # with the ``--noconsole`` option requires redirecting everything
    # (stdin, stdout, stderr) to avoid an OSError exception
    # "[Error 6] the handle is invalid."
    ret.update(
        {
            "stdin": subprocess.PIPE,
            "stderr": subprocess.DEVNULL,
            "startupinfo": si,
            "env": env,
        }
    )

    return ret


def run_powershell(cmd: str) -> str:
    """
    Run a Powershell command.

    The cmd argument should be a string, not a list. It will be passed directly
    to an underlying Powershell call. Returns the process' stdout.
    """
    stdout = subprocess.check_output(
        [
            "powershell.exe",
            "-NoProfile",
            "-NonInteractive",
            "-NoLogo",
            "-WindowStyle",
            "hidden",
            "-Command",
            cmd,
        ],
        text=True,
        **subprocess_args(False),
    )
    return stdout


def run_command(cmd: list) -> str:
    """
    Run a linux command.

    The cmd argument should be a list that can be passed directly to a
    subprocess.run call. Returns the process' stdout.
    """
    stdout = subprocess.check_output(
        cmd,
        text=True,
        **subprocess_args(False),
    )
    return stdout


class MacAddress(str):
    """
    Automatically normalizes MAC addresses to use colons instead of dashes.

    Call .with_dashes() to get a dashed version.
    """

    def __new__(cls, content):
        return str.__new__(cls, content.replace("-", ":"))

    def with_dashes(self):
        return self.replace(":", "-")


@dataclass
class Interface(object):
    """
    Simple representation of a network interface.
    """

    name: str
    description: str
    mac: MacAddress
    logical: str = None


def partition(
    iterable: Collection, sentinel: Callable[[Any], bool], include_boundary=False
) -> list:
    """
    Partition an iterable based on a sentinel function.

    Each item in the iterable is tested via the sentinel function, and when the
    sentinel returns a truthy value, the iterable is split at the current
    position. The final list of lists is then returned when iteration has
    finished.

    If include_boundary is True, then the item that triggered a partition will
    be included as a part of the group that is created. Otherwise, it is
    discarded.
    """
    partitions = []
    current = []
    for item in iterable:
        if sentinel(item):
            if include_boundary:
                current.append(item)
            partitions.append(current)
            current = []
        else:
            current.append(item)
    partitions.append(current)

    return partitions


@dataclass
class SystemInfo(object):
    """
    Get general info about the local machine.

    Works for Windows and Linux. This class provides an easy way to get relevant
    hardware info about the local machine. Instantiate an instance of the class
    and then reference its properties as needed.
    """

    @property
    def os(self) -> Optional[str]:
        """
        The full, human-readable string representing the installed OS.

        On Windows, this is the release and edition (ex. Windows 10 Professional).
        On Linux, this is the distro and version (ex. Ubuntu 20.04.1 LTS).
        """

        if system() == "Windows":
            cmd = """
                $edition = @{label="Edition";expression={$_.OsName.Substring(10)}} ;
                Get-ComputerInfo | Select $edition | ConvertTo-Json
            """
            cmd = " ".join([line.strip() for line in cmd.split("\n") if line.strip()])
            os = json.loads(run_powershell(cmd))["Edition"]
        elif system() == "Linux":
            release_info = {
                line.split("=")[0]: line.split("=")[1].replace('"', "")
                for line in open("/etc/os-release").read().splitlines()
            }
            os = release_info["PRETTY_NAME"]
        else:
            os = None
        return os

    @property
    def model(self) -> Optional[str]:
        """
        Manufacturer model name of the machine.

        This should be the same regardless of OS.
        """

        if system() == "Windows":
            cmd = "Get-ComputerInfo | Select CsModel | ConvertTo-Json"
            model = json.loads(run_powershell(cmd))["CsModel"]
        elif system() == "Linux":
            proc = run_command(["dmidecode", "-t1"]).splitlines()
            lines = partition(proc, lambda line: "System Information" in line)[-1]
            lines = {
                line.split(":")[0].strip(): line.split(":")[1].strip()
                for line in lines
                if line
            }
            if lines["Version"]:
                model = lines["Version"]
            else:
                model = None
        else:
            model = None
        return model

    @property
    def serial(self) -> Optional[str]:
        """Serial number of the computer, as given by the manufacturer."""

        if system() == "Windows":
            cmd = "(Get-WMIObject -Class WIN32_SystemEnclosure -ComputerName $env:ComputerName).SerialNumber | Select -First 1"
            serial = run_powershell(cmd).strip()
        elif system() == "Linux":
            cmd = ["dmidecode", "-t", "system"]
            result = run_command(cmd)
            serial_line = next(
                line for line in result.splitlines() if "serial" in line.lower()
            )
            serial = serial_line.split(":")[-1].strip()
        else:
            serial = None
        return serial

    @property
    def interfaces(self) -> List[Interface]:
        """
        A list of network interfaces found on the machine.

        If on Linux, the interface objects returned should also have a "logical"
        attribute that is the logical name of the interface (ex. eth0)
        """

        if system() == "Windows":
            cmd = "Get-NetAdapter -Physical | Select Name, InterfaceDescription, MacAddress | ConvertTo-Json"
            data = json.loads(run_powershell(cmd))
            interfaces = [
                Interface(
                    iface["Name"],
                    iface["InterfaceDescription"],
                    MacAddress(iface["MacAddress"]),
                )
                for iface in data
            ]
        elif system() == "Linux":

            def get_physical_interfaces():
                """
                Equivalent to `ls -d /sys/class/net/*/device | cut -d/ -f5`.
                """
                return [
                    dir.name
                    for dir in pathlib.Path("/sys/class/net").iterdir()
                    if (dir / "device").exists()
                ]

            cmd = ["lshw", "-class", "network"]
            result = run_command(cmd)
            # break up by interface
            devs = partition(result.splitlines(), lambda line: "*-network" in line)
            # filter out empty partitions
            devs = [dev for dev in devs if dev]
            # convert to dicts
            devs = [
                {
                    line.split(":")[0].strip(): ":".join(line.split(":")[1:]).strip()
                    for line in dev
                }
                for dev in devs
            ]
            interfaces = [
                Interface(
                    iface.get("product"),
                    iface.get("description").replace(" interface", ""),
                    MacAddress(iface.get("serial")),
                    logical=iface.get("logical name"),
                )
                for iface in devs
                if iface.get("logical name") in get_physical_interfaces()
            ]
        else:
            interfaces = []
        return interfaces

    @property
    def processor(self) -> Optional[str]:
        """Brand name of the processor."""
        return get_cpu_info().get("brand_raw")

    @property
    def memory(self) -> pint.Quantity:
        """Total installed physical memory on the machine, in gigabytes."""
        bytes = units.bytes * psutil.virtual_memory().total
        return bytes.to(units.gigabytes)


def make_qr(data, pixmap=True):
    """Make a QR from arbitrary data."""

    error_level = qrcode.constants.ERROR_CORRECT_H
    qr = qrcode.QRCode(error_correction=error_level, box_size=10)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image()
    if pixmap:
        img = PIL.ImageQt.toqpixmap(img)
    return img


class QCaptionedImage(QWidget):
    def __init__(self, caption=None, image=None, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.window = QLabel()
        self.window.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        self.window.setLineWidth(2)

        self.caption = QLabel()
        self.caption.setAlignment(Qt.AlignCenter)
        self.caption.setWordWrap(True)

        # placeholder background when nothing is shown
        self.blank_background = QPixmap(290, 290)
        self.blank_background.fill(QColor(0, 0, 0, 32))

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.caption)
        self.layout.addWidget(self.window)
        self.setLayout(self.layout)

        if caption:
            self.caption.setText(caption)

        if image is None:
            self.clearImage()
        else:
            self.setImage(image)

    def clearImage(self):
        self.window.setPixmap(self.blank_background)

    def setImage(self, image: QPixmap):
        self.window.setPixmap(image)


class MainWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        info = SystemInfo()
        title_strings = [
            string
            for string in (
                info.model,
                info.os,
                info.processor,
                f"{int(info.memory.magnitude)} GB RAM",
            )
            if string is not None
        ]

        self.setWindowTitle("   —   ".join(title_strings))

        self.interfaces = []
        for interface in info.interfaces:
            image = make_qr(interface.mac)
            logical_name = (
                f" ({interface.logical})" if interface.logical is not None else ""
            )
            caption = f"{interface.name} — {interface.description}{logical_name}"

            self.interfaces.append(QCaptionedImage(caption, image))

        image = make_qr(info.serial)
        self.serial = QCaptionedImage(f"Serial Number ({info.serial})", image)

        layout = QHBoxLayout()
        for widget in self.interfaces + [self.serial]:
            layout.addWidget(widget)
        self.setLayout(layout)


def main():
    if system() == "Linux" and os.geteuid() != 0:
        print("You must run this as root!")
        sys.exit(1)
    app = QApplication()
    ui = MainWindow()

    ui.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

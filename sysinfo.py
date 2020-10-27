#!/usr/bin/env python3

# builtins
import subprocess
import json
import sys
import os

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

# my modules


# Create a set of arguments which make a ``subprocess.Popen`` (and
# variants) call work with or without Pyinstaller, ``--noconsole`` or
# not, on Windows and Linux. Typical use::
#
#   subprocess.call(['program_to_run', 'arg_1'], **subprocess_args())
#
# When calling ``check_output``::
#
#   subprocess.check_output(['program_to_run', 'arg_1'],
#                           **subprocess_args(False))
def subprocess_args(include_stdout=True):
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
            "stderr": subprocess.PIPE,
            "startupinfo": si,
            "env": env,
        }
    )

    return ret


def run_powershell(cmd):
    proc = subprocess.run(
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
        **subprocess_args(),
    )
    return proc.stdout


def get_interfaces():
    cmd = "Get-NetAdapter -Physical | Select Name, InterfaceDescription, MacAddress | ConvertTo-Json"
    result = run_powershell(cmd)
    data = json.loads(result)
    for interface in data:
        interface["MacAddress"] = interface["MacAddress"].replace("-", ":")
    return data


def get_serial():
    cmd = "(Get-WMIObject -Class WIN32_SystemEnclosure -ComputerName $env:ComputerName).SerialNumber | Select -First 1"
    result = run_powershell(cmd)
    return result.strip()


def get_sysinfo():
    cmd = """
        $processor = @{label="Processor";expression={$_.CsProcessors[0].name}} ;
        $mem_size = @{label="Memory";expression={"$([math]::round($_.CsPhyicallyInstalledMemory*1Kb/1Gb))GB RAM"}} ;
        $edition = @{label="Edition";expression={$_.OsName.Substring(10)}} ;
        Get-ComputerInfo | Select CsModel,$edition,$processor,$mem_size
        | ConvertTo-Json
    """
    cmd = " ".join([line.strip() for line in cmd.split("\n") if line.strip()])
    data = json.loads(run_powershell(cmd))
    return "   —   ".join(data.values())


def make_qr(data, pixmap=True):
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

        self.setWindowTitle(get_sysinfo())

        self.interfaces = []
        for interface in get_interfaces():
            name = interface["Name"]
            desc = interface["InterfaceDescription"]

            image = make_qr(interface["MacAddress"])
            caption = f"{name} — ({desc})"

            self.interfaces.append(QCaptionedImage(caption, image))

        serial = get_serial()
        image = make_qr(serial)
        self.serial = QCaptionedImage("Serial Number", image)

        layout = QHBoxLayout()
        for widget in self.interfaces + [self.serial]:
            layout.addWidget(widget)
        self.setLayout(layout)


def main():
    app = QApplication()
    ui = MainWindow()

    ui.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

import concurrent.futures
import json
import logging
import os
import pathlib
import subprocess
import sys
import threading
import time
import tkinter as tk
import tkinter.font as font
import tkinter.messagebox
from enum import Enum, auto
from tkinter import scrolledtext
from tkinter.constants import RAISED, SUNKEN
from typing import Dict, List
import ctypes
import csv
import numpy as np
import serial
import serial.tools.list_ports
import zope.event
import configClass
import pyoto.otoProtocol.otoCommands as pyoto
import pyoto.otoProtocol.otoMessageDefs as otoMessageDefs

# otatool.py lives in $IDF_PATH/components/app_update
IDF_PATH = os.path.join(os.path.dirname(__file__), "esp-idf")
# otatool.py lives in $IDF_PATH/components/esptool_py/esptool
ESPTOOL_DIR = (pathlib.Path(__file__).parent / "esp-idf" / "components" / "esptool_py" / "esptool").resolve()
ESPTOOL_PY_DIR = (ESPTOOL_DIR / "esptool.py").resolve()
# -------- Basic Settings --------
FIRMWARE_VERSION = "v2.4.0.0-v5"  # firmware for v5.3 boards
BOM_NUMBER = "SOMETHING IS WRONG, FILL IN THE BOM NUMBER FOR CONFIG.YML"  # this will be taken from the config.yml, if
# this message shows, then user will know to fix.
CONFIG_YAML_PATH = "config.yml"
# updated values from v4.3 boards Mar 22, 2023 per 537 pieces with new power supply, flyers removed, ±4σ
ADC_BATT_LOW_LIMIT = 2113
ADC_BATT_HIGH_LIMIT = 2631
# -------- Other Settings --------
# USB VID and PID of OtO flasher board
VALID_VID = 0x10C4
VALID_PID = 0xEA60
# Max number of workers
WORKERS = 20
lock = threading.Lock()
globalLoggingLevel = logging.INFO
# Set up logging
logging.basicConfig(level=globalLoggingLevel, format="%(message)s")
mainLogger = logging.getLogger(__name__)
DEFAULT_FLASHER_ARGS_JSON = {
    "write_flash_args": ["--flash_mode", "dio",
        "--flash_size", "detect",
        "--flash_freq", "40m"],
    "flash_settings": {
        "flash_mode": "dio",
        "flash_size": "detect",
        "flash_freq": "40m",
    },
    "flash_files": {
        "0x1000": "bootloader/bootloader.bin",
        "0x50000": "OtO-Firmware.bin",
        "0x8000": "partition_table/partition-table.bin",
        "0x49000": "ota_data_initial.bin"
    },
    "extra_esptool_args": {
        "after": "hard_reset",
        "before": "default_reset",
        "stub": True,
        "chip": "esp32",
    },
    "bootloader": {
        "offset": "0x1000",
        "file": "bootloader/bootloader.bin",
        "encrypted": "false"
    },
    "app": {
        "offset": "0x50000",
        "file": "OtO-Firmware.bin",
        "encrypted": "false"
    },
    "partition-table": {
        "offset": "0x8000",
        "file": "partition_table/partition-table.bin",
        "encrypted": "false"
    },
    "otadata": {
        "offset": "0x49000",
        "file": "ota_data_initial.bin",
        "encrypted": "false"
    }
}

class ButtonState:
    DISABLED = "disabled"
    NORMAL = "normal"

class bcolors:
    """Console colors"""
    HEADER = "\033[95m"
    OKBLUE = "\033[94m"
    OKCYAN = "\033[96m"
    OKGREEN = "\033[92m"
    WARNING = "\033[93m"
    FAIL = "\033[91m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"
    UNDERLINE = "\033[4m"

def threaded(fn):
    def run(*k, **kw):
        t = threading.Thread(target=fn, args=k, kwargs=kw)
        t.start()
        return t

    return run

class EventType(Enum):
    UPDATE_FLASH_ALL = auto()
    DISABLE_PACK = auto()
    ENABLE_PACK = auto()

class SerialBoardCard(tk.Frame):
    """Com Port Card Gui and Functional Class
    One instance of this class is created for every flasher board with a serial"""
    IDLE_COLOR = "#dddddd"
    BUSY_COLOR = "#e9c7ff"
    OK_COLOR = "#40ff40"
    ERROR_COLOR = "#ff145b"

    class PortStatus(Enum):
        IDLE = auto()
        FLASHING = auto()
        FLASHING_1 = auto()
        FLASHING_2 = auto()
        FLASHING_3 = auto()
        READING_ADC = auto()
        CONNECTING = auto()
        CHECK_PRESSURE = auto()
        CONNECTED = auto()
        WRITING = auto()
        SUCCESS = auto()
        ADC_READING_SUCCESS = auto()
        CALIBRATION_SUCCESS = auto()
        FAIL = auto()
        FAIL_ADC = auto()
        FAIL_PRESSURE = auto()
        FAIL_CALIBRATION = auto()
        CONNECT_FLASHER = auto()
        CHECK_POWER = auto()
        FAILBATTERYVOLTAGE = auto()

    class TextHandler(logging.Handler):
        # This class allows you to log to a Tkinter Text or ScrolledText widget
        MAX_LINES = 100

        def __init__(self, scrolledTextWidget: scrolledtext.ScrolledText):
            logging.Handler.__init__(self)
            # Store a reference to the scrolledText it will log to
            self.scrolledTextWidget = scrolledTextWidget

        def limit_lines(self):
            # while the total number of lines is greater than max lines
            while float(self.scrolledTextWidget.index("end-1c")) > self.MAX_LINES:
                # remove the first line
                self.scrolledTextWidget.delete("1.0", "2.0")

        def emit(self, record: logging.LogRecord):
            msg = self.format(record)
            self.scrolledTextWidget.configure(state="normal")
            self.scrolledTextWidget.insert(tk.END, msg + "\n", record.levelno)
            self.limit_lines()
            self.scrolledTextWidget.configure(state="disabled")
            # Autoscroll to the bottom
            self.scrolledTextWidget.yview(tk.END)

    def __init__(
        self,
        master,
        flasherSerial: str,
        text: str,
        config_object: configClass.OtoFlasherConfigObject,
    ):

        # ---- Init Self Widget ----
        super().__init__(master=master, width=100, height=300, borderwidth=1, relief=RAISED)
        self.logger = logging.getLogger(flasherSerial)
        fontComPortTitle = font.Font(size=10)
        fontComPortStatus = font.Font(size=12)
        self.pack(side="left", expand=True, fill="both", pady=2, padx=2)

        # ---- Init Port Name Label Widget ----
        self.labelPortName = tk.Label(self, font=fontComPortTitle, text=str(text))
        self.labelPortName.pack(pady=2)

        # ---- Init Status Label Widget ----
        self.labelStatus = tk.Label(self, font=fontComPortStatus)
        self.labelStatus.pack(side=tk.TOP, pady=4, anchor=tk.CENTER, expand=False, fill="both")

        # Disable pack propagate here because ScrolledText seems to mess with card size
        self.pack_propagate(False)
        infoBoxFont = font.Font(family = "Microsoft YaHei UI", size = 8)
        self.infoBox = scrolledtext.ScrolledText(self, wrap = tk.CHAR, width = 10, height = 8, font = infoBoxFont)

        # Remove existing handlers
        while len(self.logger.handlers):
            self.logger.removeHandler(self.logger.handlers[0])

        # Add handler to direct logs to infoBox
        self.logger.addHandler(SerialBoardCard.TextHandler(self.infoBox))
        self.infoBox.pack(side = tk.BOTTOM, padx = 2, pady = 2, expand = True, fill = tk.BOTH)

        # Define styling for logging levels
        self.infoBox.tag_config(logging.NOTSET, foreground="black")
        self.infoBox.tag_config(logging.DEBUG, foreground="gray")
        self.infoBox.tag_config(logging.INFO, foreground="black")
        self.infoBox.tag_config(logging.WARNING, foreground="orange")
        self.infoBox.tag_config(logging.ERROR, foreground="red")
        self.infoBox.tag_config(logging.CRITICAL, foreground="red", underline = 1)

        # ---- Set other self properties ----
        self.port = None
        self.unitSerial = None
        self.flasherSerial = flasherSerial
        self.status = SerialBoardCard.PortStatus.IDLE
        self.MACAddress = None
        self.config_object = config_object
        self.ADC_Failed = False
        self.Pressure_Failed = False
        self.adcReading = 0
        self.ZeroPressure = 0
        self.ZeroPressureSTD = 0
        self.bomNumber: str = None

    def __str__(self):
        return self.labelPortName["text"]

    @property
    def status(self):
        return self._status

    @status.setter
    def status(self, new_status: PortStatus):

        # Route to corresponding state change function
        if new_status == SerialBoardCard.PortStatus.IDLE:
            self._setStatusIdle()
        elif new_status == SerialBoardCard.PortStatus.FLASHING:
            self._setStatusFlashing()
        elif new_status == SerialBoardCard.PortStatus.READING_ADC:
            self._setStatusReadingADC()
        elif new_status == SerialBoardCard.PortStatus.WRITING:
            self._setStatusWriting()
        elif new_status == SerialBoardCard.PortStatus.CONNECTING:
            self._setStatusConnecting()
        elif new_status == SerialBoardCard.PortStatus.CHECK_PRESSURE:
            self._setStatusCheckPressure()
        elif new_status == SerialBoardCard.PortStatus.CONNECTED:
            self._setStatusConnected()
        elif new_status == SerialBoardCard.PortStatus.FAIL:
            self._setStatusFail()
        elif new_status == SerialBoardCard.PortStatus.SUCCESS:
            self._setStatusSuccess()
        elif new_status == SerialBoardCard.PortStatus.ADC_READING_SUCCESS:
            self._setStatusADCReadingSuccess()
        elif new_status == SerialBoardCard.PortStatus.FAIL_ADC:
            self._setStatusFailADC()
        elif new_status == SerialBoardCard.PortStatus.FAIL_PRESSURE:
            self._setStatusFailPressure()
        elif new_status == SerialBoardCard.PortStatus.CALIBRATION_SUCCESS:
            self._setStatusCalibrationSuccess()
        elif new_status == SerialBoardCard.PortStatus.FAIL_CALIBRATION:
            self._setStatusFailCalibration()
        elif new_status == SerialBoardCard.PortStatus.CONNECT_FLASHER:
            self._setStatusConnectFlasher()
        elif new_status == SerialBoardCard.PortStatus.CHECK_POWER:
            self._setStatusCheckPowerSupply()
        elif new_status == SerialBoardCard.PortStatus.FAILBATTERYVOLTAGE:
            self._setStatusFailBatteryVoltage()
        else:
            self.logger.warning(f"Invalid status: {new_status}")

    @property
    def isBusy(self):
        """Check if current status of port is busy"""
        busyStates = [
            SerialBoardCard.PortStatus.FLASHING,
            SerialBoardCard.PortStatus.FLASHING_1,
            SerialBoardCard.PortStatus.FLASHING_2,
            SerialBoardCard.PortStatus.FLASHING_3,
            SerialBoardCard.PortStatus.READING_ADC,
            SerialBoardCard.PortStatus.CONNECTING,
            SerialBoardCard.PortStatus.CONNECTED,
            SerialBoardCard.PortStatus.WRITING,
            SerialBoardCard.PortStatus.ADC_READING_SUCCESS,
            SerialBoardCard.PortStatus.SUCCESS,
            SerialBoardCard.PortStatus.CALIBRATION_SUCCESS,
        ]

        # True if current state is one of the busy states
        if self.status in busyStates:
            return True
        return False

    @isBusy.setter
    def isBusy(self, new_busy):
        self.logger.warning(f"Cannot set isBusy to {new_busy}, Read-only Property")

    def flashButtonCallback(self):
        """Flash Unit, calibrate voltage, check pressure ADC"""

        self.ADC_Failed = False  # Resetting flag to False so that the previous flash does not carry onto the next
        self.Pressure_Failed = False
        Battery: list = []

        self.logger.info("Started processing\n启动程序...")

        # Update GUI
        self.status = SerialBoardCard.PortStatus.FLASHING
        zope.event.notify(EventType.UPDATE_FLASH_ALL)

        # STEP 1 Find COM port
        error = self.getSerialPortFromUSBSerial()
        if error is not None:
            self.logger.error(error)
            self.status = SerialBoardCard.PortStatus.CONNECT_FLASHER
            zope.event.notify(EventType.UPDATE_FLASH_ALL)
            return error

        # STEP 2 Flash Unit
        error = self.flashUnitEspTool(self.port, FIRMWARE_VERSION)
        if error is not None:
            self.logger.error(error)
            self.status = SerialBoardCard.PortStatus.FAIL
            zope.event.notify(EventType.UPDATE_FLASH_ALL)
            return error

        # STEP 3 Connect to OtO after making it reboot
        error = self.OtOConnect()
        if error is not None:
            self.logger.error(error)
            self.status = SerialBoardCard.PortStatus.FAIL
            zope.event.notify(EventType.UPDATE_FLASH_ALL)
            return error

        # STEP 4 Get pressure sensor version so we can set appropriate zero pressure limits
        error = self.getPressureSensorVersion()
        if error is not None:
            self.logger.error(error)
            self.status = SerialBoardCard.PortStatus.FAIL
            zope.event.notify(EventType.UPDATE_FLASH_ALL)
            return error

        # Step 5 Zero Pressure Check to confirm sensor is usable
        error = self.EstablishZeroPressure(data_collection_time = 3.0)
        if error is not None:
            self.logger.error(error)
            self.status = SerialBoardCard.PortStatus.FAIL_PRESSURE
            zope.event.notify(EventType.UPDATE_FLASH_ALL)
            return error

        # STEP 6 Reading ADC from Oto
        error = self.getBattCalibrationValue()
        error = self.readBatteryAdcSerialPyoto()
        if error is not None:
            self.logger.error(error)
            self.status = SerialBoardCard.PortStatus.FAIL_ADC
            zope.event.notify(EventType.UPDATE_FLASH_ALL)
            return error

        # Step 7 Writing ADC values to flash
        error = self.writeCalibrationVoltagesSerialPyoto()
        if error is not None:
            self.logger.error(error)
            self.status = SerialBoardCard.PortStatus.FAIL_CALIBRATION
            zope.event.notify(EventType.UPDATE_FLASH_ALL)
            return error

        # Step 8 Check Battery Voltage is 4.1±0.05 after resetting OtO
        self.pyoto_instance.stop_connection()
        error = self.OtOConnect()
        if error is not None:
            self.logger.error(error)
            self.status = SerialBoardCard.PortStatus.FAIL
            zope.event.notify(EventType.UPDATE_FLASH_ALL)
            return error
        for x in range(1, 10):
            Battery.extend([self.pyoto_instance.get_voltages().battery_voltage_v])
            time.sleep(0.1)
        BattVoltage = round(float(np.average(Battery)), 3)
        if BattVoltage > 4.15 or BattVoltage < 4.05:
            error = f"Board voltage is not 4.1±0.05V\n线路板没有读取4.1±0.05V: {BattVoltage} V"
            self.logger.error(error)
            self.status = SerialBoardCard.PortStatus.FAILBATTERYVOLTAGE
            zope.event.notify(EventType.UPDATE_FLASH_ALL)
            return error
        else:
            self.logger.info(f"Actual Battery 实际电池电压: {BattVoltage}V")

        self.logger.info("All steps completed.\n完成所有步骤")
        return None

    def _setStatusIdle(self):
        self.configure(background=self.IDLE_COLOR)
        self._status = SerialBoardCard.PortStatus.IDLE
        self.labelPortName["bg"] = self.IDLE_COLOR
        self.labelStatus["text"] = "Idle 闲置中"
        self.labelStatus["bg"] = self.IDLE_COLOR

    def _setStatusBusy(self):
        self.configure(background=self.BUSY_COLOR)
        self.labelPortName["bg"] = self.BUSY_COLOR
        self.labelStatus["text"] = "Busy 忙碌"
        self.labelStatus["bg"] = self.BUSY_COLOR

    def _setStatusSuccess(self):
        self._setStatusBusy()
        self.configure(background=self.OK_COLOR)
        self._status = SerialBoardCard.PortStatus.SUCCESS
        self.labelPortName["bg"] = self.OK_COLOR
        self.labelStatus["text"] = "DONE 完成"
        self.labelStatus["bg"] = self.OK_COLOR

    def _setStatusFail(self):
        self.configure(background=self.ERROR_COLOR)
        self._status = SerialBoardCard.PortStatus.FAIL
        self.labelPortName["bg"] = self.ERROR_COLOR
        self.labelStatus["text"] = "REFLASH 重新存储"
        self.labelStatus["bg"] = self.ERROR_COLOR

    def _setStatusFailADC(self):
        self.configure(background=self.ERROR_COLOR)
        self._status = SerialBoardCard.PortStatus.FAIL_ADC
        self.labelPortName["bg"] = self.ERROR_COLOR
        self.labelStatus["text"] = "Failed ADC Reading\nADC读数错误"
        self.labelStatus["bg"] = self.ERROR_COLOR

    def _setStatusFailPressure(self):
        self.configure(background=self.ERROR_COLOR)
        self._status = SerialBoardCard.PortStatus.FAIL_PRESSURE
        self.labelPortName["bg"] = self.ERROR_COLOR
        self.labelStatus["text"] = "Failed Pressure Sensor\n压力传感器故障"
        self.labelStatus["bg"] = self.ERROR_COLOR

    def _setStatusFailCalibration(self):
        self.configure(background=self.ERROR_COLOR)
        self._status = SerialBoardCard.PortStatus.FAIL_CALIBRATION
        self.labelPortName["bg"] = self.ERROR_COLOR
        self.labelStatus["text"] = "Failed Calibration\n校准失败"
        self.labelStatus["bg"] = self.ERROR_COLOR

    def _setStatusCalibrationSuccess(self):
        self._setStatusBusy()
        self._status = SerialBoardCard.PortStatus.CALIBRATION_SUCCESS
        self.labelStatus["text"] = "Calibration Successful\n校准成功"

    def _setStatusFailBatteryVoltage(self):
        self.configure(background=self.ERROR_COLOR)
        self._status = SerialBoardCard.PortStatus.FAILBATTERYVOLTAGE
        self.labelPortName["bg"] = self.ERROR_COLOR
        self.labelStatus["text"] = "Failed Board Voltage\n错误的线路板电压"
        self.labelStatus["bg"] = self.ERROR_COLOR

    def _setStatusCheckPowerSupply(self):
        self.configure(background=self.ERROR_COLOR)
        self._status = SerialBoardCard.PortStatus.CHECK_POWER
        self.labelPortName["bg"] = self.ERROR_COLOR
        self.labelStatus["text"] = "CHECK POWER SUPPLY\n检查电源"
        self.labelStatus["bg"] = self.ERROR_COLOR

    def _setStatusADCReadingSuccess(self):
        self._setStatusBusy()
        self._status = SerialBoardCard.PortStatus.ADC_READING_SUCCESS
        self.labelStatus["text"] = "ADC Reading Successful\nADC读数成功"

    def _setStatusFlashing(self, flashing_stage=0, total_stages=3):
        self._setStatusBusy()
        self._status = SerialBoardCard.PortStatus.FLASHING
        self.labelStatus["text"] = "Flashing 刷新"

    def _setStatusReadingADC(self):
        self._setStatusBusy()
        self._status = SerialBoardCard.PortStatus.READING_ADC
        self.labelStatus["text"] = "Reading ADC 读取ADC"

    def _setStatusConnecting(self):
        self._setStatusBusy()
        self._status = SerialBoardCard.PortStatus.CONNECTING
        self.labelStatus["text"] = "Connecting 连接"

    def _setStatusCheckPressure(self):
        self._setStatusBusy()
        self._status = SerialBoardCard.PortStatus.CONNECTING
        self.labelStatus["text"] = "Zero pressure ADC\n零压力ADC"

    def _setStatusConnected(self):
        self._setStatusBusy()
        self._status = SerialBoardCard.PortStatus.CONNECTED
        self.labelStatus["text"] = "Connected 连接成功"

    def _setStatusWriting(self):
        self._setStatusBusy()
        self._status = SerialBoardCard.PortStatus.WRITING
        self.labelStatus["text"] = "Writing 存储"

    def _setStatusConnectFlasher(self):
        self.configure(background=self.ERROR_COLOR)
        self._status = SerialBoardCard.PortStatus.CONNECT_FLASHER
        self.labelPortName["bg"] = self.ERROR_COLOR
        self.labelStatus["text"] = "CONNECT FLASHER\n连接USB通信测试板"
        self.labelStatus["bg"] = self.ERROR_COLOR

    def read_assemble_flash_args(self, build_folder_path, port: str):
        """Read flasher_args.json file and assemble args list for esptool.py"""

        # Get flash args from flash_project_args file
        flash_args_json_file_path = (pathlib.Path(build_folder_path) / "flasher_args.json").resolve()
        print("flash_args_file_path:", flash_args_json_file_path.absolute())

        # Read json file
        try:
            with open(flash_args_json_file_path, "r") as f:
                flasher_args_json = json.load(f)
        except FileNotFoundError:
            # If flasher_args.json not found, fall back to default flasher_args.json
            # print(f"File {flash_args_json_file_path} not found, Falling back to default flasher_args")
            flasher_args_json = DEFAULT_FLASHER_ARGS_JSON

        # Get bin paths
        offset_path_list = list()
        flash_files_dict: Dict[str, str] = flasher_args_json["flash_files"]
        for key, value in flash_files_dict.items():
            offset = key
            bin_file_path: pathlib.Path = (pathlib.Path(build_folder_path) / value).resolve()
            assert bin_file_path.exists(), f"Binary file {bin_file_path} not found\n"
            offset_path_list.append(offset)
            offset_path_list.append(str(bin_file_path))

        flash_args_list: List[str] = [
            "--chip",
            "esp32",
            "--baud",
            "921600",
            "--port",
            port,
            "--connect-attempts",
            "20",
            "write_flash",
            *flasher_args_json["write_flash_args"],
            *offset_path_list,
        ]
        return flash_args_list

    def flashUnitEspTool(self, port: str, firmware_version: str):
        """Flash the esp on the given port with the given firmware folder name
        Args:       port = name of com port
                    firmware_version = name of folder that contains the firmware to use
        Returns:    None if success
                    Error string if fail"""

        self.status = SerialBoardCard.PortStatus.FLASHING
        build_folder_path = pathlib.Path(__file__).parent / "binaries" / firmware_version / "build"
        flash_args_list = self.read_assemble_flash_args(build_folder_path=build_folder_path, port=port)
        # included explicit path to improve path search speed during execution
        try:
            # Run esptool with args
            # esptool.main(flash_args_list)
            # included explicit path to improve path search speed during execution
            popen = subprocess.Popen(args=[str(sys.executable), str(ESPTOOL_PY_DIR), *flash_args_list], stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
            while popen.poll() is None:
                text = popen.stdout.readline()
                text = text[:len(text) - 1]
                self.logger.info(f"\r{text}")

            return_code = popen.wait()
            if return_code != 0:
                raise Exception(popen.returncode, popen.stderr.read())
        except Exception:
            self.logger.exception("Exception 例外: ")
            return (f"Failed to flash PCB on {port}.\n存储{port}上的PCB失败")

        # DONE
        return None

    def getSerialPortFromUSBSerial(self):
        """Match port with serial card by matching the serial numbers
        Uses self.flasherSerial

        Saves as string to self.port"""

        self.port = None
        allPorts = serial.tools.list_ports.comports()
        for port in allPorts:
            if port.serial_number == self.flasherSerial and port.vid == VALID_VID and port.pid == VALID_PID:
                self.port = port.name

        if self.port is None:
            return f"COM port with serial {self.flasherSerial}, vid {VALID_VID}, pid {VALID_PID} not found. Ensure flashing boards are connected to the computer\n确认USB通信测试板与计算机连接"

        return None

    def readBatteryAdcSerialPyoto(self):
        ADCValues: list = []
        try:
            self.status = SerialBoardCard.PortStatus.READING_ADC
            zope.event.notify(EventType.UPDATE_FLASH_ALL)
            for x in range(1, 10):
                return_message = self.pyoto_instance.get_voltages_adc()
                ADCValues.extend([return_message.battery_voltage_adc])
                time.sleep(0.1)

            self.batt_voltage_adc = round(float(np.average(ADCValues)), 0)
            self.logger.info(f"New voltage value from board: {self.batt_voltage_adc:,} ADC\n线路板的新电压值")

            self.adcReading = self.batt_voltage_adc

            if self.batt_voltage_adc < ADC_BATT_LOW_LIMIT:
                self.logger.info("Voltage ADC value too low!\n电压太低")
                self.ADC_Failed = True
                return f"Measured voltage ADC value too low on port {self.port}\n测试电压太低"

            if self.batt_voltage_adc > ADC_BATT_HIGH_LIMIT:
                self.logger.info("Voltage ADC value too high!\n电压太高")
                self.ADC_Failed = True
                return f"Measured voltage ADC value too high on port {self.port}\n测试电压太高"

        except Exception as error:
            self.logger.exception("Failed to read voltage ADC from board\n线路板电压读取失败")
            return f"PyOtO failed to read voltage ADC from board on port {self.port}:\n线路板电压读取失败\n{repr(error)}"

    def writeCalibrationVoltagesSerialPyoto(self):
        """ Writes the adc values to flash
        Returns:    None if success
                    str of error if fail
        """
        self.status = SerialBoardCard.PortStatus.WRITING
        v41 = self.adcReading
        try:
            self.pyoto_instance.set_calibration_voltages(v41)
            self.status = SerialBoardCard.PortStatus.CALIBRATION_SUCCESS
            zope.event.notify(EventType.UPDATE_FLASH_ALL)
        except Exception as error:
            self.status = SerialBoardCard.PortStatus.FAIL_CALIBRATION
            zope.event.notify(EventType.UPDATE_FLASH_ALL)
            self.logger.exception("Failed writing calibration value to board\n存储校正值到线路板失败")
            return f"PyOtO failed to connect to board on port {self.port}:\n存储校正值到线路板失败\n{repr(error)}"

    def EstablishZeroPressure(self, data_collection_time: float):
        self.status = SerialBoardCard.PortStatus.CHECK_PRESSURE
        number_of_trials: int = 2
        trial_count: int = 1
        loop_check: bool = True
        STD_check = True
        ADC_check = True
        self.logger.info("Checking zero pressure ADC value...\n检查零压力ADC值")
        self.pyoto_instance.use_moving_average_filter(True)
        while loop_check and trial_count <= number_of_trials:
            main_loop_start_time = time.perf_counter()
            Sensor_Read_List: List[pyoto.otoMessageDefs.SensorReadMessage] = []
            pressureReading: list = []
            dataCount: int = 0
            self.ZeroPressure = 0
            self.ZeroPressureSTD = 0
            self.pyoto_instance.set_sensor_subscribe(subscribe_frequency=pyoto.SensorSubscribeFrequencyEnum.SENSOR_SUBSCRIBE_FREQUENCY_100Hz)
            time.sleep(0.1)
            self.pyoto_instance.clear_incoming_packet_log()
            while time.perf_counter() - main_loop_start_time <= data_collection_time:
                Sensor_Read_List.extend(self.pyoto_instance.read_all_sensor_packets(limit=None, consume=True))
            self.pyoto_instance.set_sensor_subscribe(subscribe_frequency=pyoto.SensorSubscribeFrequencyEnum.SENSOR_SUBSCRIBE_FREQUENCY_OFF)
            ElapsedTime = time.perf_counter() - main_loop_start_time

            if not Sensor_Read_List:
                return "No pressure data was collected.\n未收集压力数值"

            for message in Sensor_Read_List:
                pressureReading.append(int(message.pressure_adc))
                dataCount += 1

            self.ZeroPressure = round(np.mean(pressureReading), 0)
            self.ZeroPressureSTD = round(np.std(pressureReading), 1)

            trial_count += 1
            if (self.ZeroPressureSTD <= self.max_acceptable_STD and self.ZeroPressureSTD >= self.min_acceptable_STD) and (self.ZeroPressure <= self.max_acceptable_ADC and self.ZeroPressure >= self.min_acceptable_ADC):
                loop_check = False

        if self.ZeroPressureSTD > self.max_acceptable_STD or self.ZeroPressureSTD < self.min_acceptable_STD:
            STD_check = False
        if self.ZeroPressure > self.max_acceptable_ADC or self.ZeroPressure < self.min_acceptable_ADC:
            ADC_check = False

        if (not STD_check) or (not ADC_check):
            if (not STD_check) and (not ADC_check):
                return f"Failed Zero Pressure Check on Mean AND STD: Pressure: {self.ZeroPressure:,} ADC, sigma: {self.ZeroPressureSTD:,} ADC\n平均和标准差值零压力检查失败"
            if not STD_check:
                return f"Failed Zero Pressure Check on STD: Pressure: {self.ZeroPressure:,} ADC, sigma: {self.ZeroPressureSTD:,} ADC\n标准差值零压力检查失败"
            if not ADC_check:
                return f"Failed Zero Pressure Check on Mean: Pressure: {self.ZeroPressure:,} ADC, sigma: {self.ZeroPressureSTD:,} ADC\n平均值零压力检查失败"

        self.logger.info(f"Zero pressure 零压力: {self.ZeroPressure:,} ADC\nSTD 标准差值: {self.ZeroPressureSTD:,} ADC\npoints 数据数量: {dataCount}, Elapsed 经过的时间: {ElapsedTime:0.4f} sec")
        return None

    def OtOConnect(self):
        try:
            self.pyoto_instance = pyoto.OtoInterface(connection_type=pyoto.ConnectionType.UART, logger=None)
            self.pyoto_instance.logger.setLevel(logging.INFO)
            self.logger.info("Waiting for board to reboot...\n等待线路板重启")
            self.pyoto_instance.start_connection(port=self.port, reset_on_connect=True)
            self.status = SerialBoardCard.PortStatus.CONNECTED
            self.MACAddress = self.pyoto_instance.get_mac_address().string
        except Exception as error:
            self.logger.exception("Failed to connect to board\n连接线路板失败")
            return f"Failed to connect to board on port {self.port}:\n连接线路板失败\n{repr(error)}"

    def getPressureSensorVersion(self):
        # assumes pyoto connection is open
        try:
            return_message = self.pyoto_instance.get_pressure_sensor_version()
        except Exception as error:
            self.logger.exception("Failed to get pressure sensor address\n无法获得压力传感器地址")
            return f"Failed to get pressure sensor address on port {self.port}:\n无法获得压力传感器地址\n{repr(error)}"

        returned_pressure_sensor_version = return_message.pressure_sensor_version  # should be an int

        # see PressureSensorVersionEnum in otoMessageDefs for breakdown
        if returned_pressure_sensor_version == otoMessageDefs.PressureSensorVersionEnum.PRESSURE_SENSOR_UNINITIALIZED.value:
            return "Uninitialized pressure sensor\n压力传感器未能启动"
        elif returned_pressure_sensor_version == otoMessageDefs.PressureSensorVersionEnum.TPBD_15_PSI_GAUGE.value:
            return "This is an old design board and should not be flashed on this station.\n这是一个旧的线路板设计, 不应存储到这个工作台"
        elif returned_pressure_sensor_version == otoMessageDefs.PressureSensorVersionEnum.MPRL_15_PSI_GAUGE.value:
            self.bomNumber += "1"
            self.logger.info("15 psi pressure sensor detected\n检测到0.1MPa压力传感器")
            self.max_acceptable_STD: float = 387.8  # Jan 2023 ±4σ
            self.min_acceptable_STD: float = 96.5  # Jan 2023 ±4σ
            self.max_acceptable_ADC: float = 1786755  # Jan 2023 ±4σ
            self.min_acceptable_ADC: float = 1611555  # Jan 2023 ±4σ
            return None
        elif returned_pressure_sensor_version == otoMessageDefs.PressureSensorVersionEnum.MPRL_30_PSI_GAUGE.value:
            self.logger.info("30 psi pressure sensor detected\n检测到0.21MPa压力传感器")
            self.max_acceptable_STD: float = 206.9  # Jan 2023 ±4σ
            self.min_acceptable_STD: float = 66.5  # Jan 2023 ±4σ
            self.max_acceptable_ADC: float = 1764145  # Jan 2023 ±4σ 1764145
            self.min_acceptable_ADC: float = 1630925  # Jan 2023 ±4σ 1630925
            return None
        else:
            return "Unknown pressure sensor detected\n检查到未知压力传感器"

    def getBattCalibrationValue(self):
        try:
            return_message: otoMessageDefs.ReadCalibrationVoltagesMessage = self.pyoto_instance.get_calibration_voltages()
        except pyoto.NotInitializedException:
            self.logger.info("Board voltage not calibrated...\n线路板电压没被校准")
            return None
        except Exception as error:
            self.logger.exception("Unable to read a voltage calibration value\n无从读取电压校准值")
            return f"Failed to read a voltage calibration value on port {self.port}:\n无从读取电压校准值\n{repr(error)}"
        self.adcReading = return_message.calib_4v1
        self.logger.info(f"Present value 当前数值: {self.adcReading:,} ADC\nCalibrating again 再次校准...")
        return None

class FlashAllButton(tk.Button):
    FONT_FAMILY = "Microsoft YaHei UI"
    FONT_SIZE = 22
    FONT_WEIGHT = "normal"
    NORMAL_COLOR_BG = "#0093f5"
    NORMAL_COLOR_FG = "#000000"
    DISABLED_COLOR_BG = "#00dfed"
    DISABLED_COLOR_FG = "#ffffff"

    def __init__(self, master=None, command=None):
        tk.Button.__init__(
            self,
            master,
            font=font.Font(family=self.FONT_FAMILY, size=self.FONT_SIZE, weight=self.FONT_WEIGHT),
            text="Flash All 存储所有线路板",
            command=command,
        )
        self.enable()

    def enable(self):
        self["state"] = ButtonState.NORMAL
        self["fg"] = self.NORMAL_COLOR_FG
        self["bg"] = self.NORMAL_COLOR_BG

    def disable(self):
        self["state"] = ButtonState.DISABLED
        self["fg"] = self.DISABLED_COLOR_FG
        self["bg"] = self.DISABLED_COLOR_BG

class VersionBox(tk.LabelFrame):
    def __init__(
        self,
        firmwareVersion,
    ):
        tk.LabelFrame.__init__(self, text="Version  版本")
        labelText = f"Firmware Version  固件版本: {firmwareVersion}"
        self.labelScriptVersion = tk.Label(self, text=labelText)
        self.labelScriptVersion.pack(side="top")

class Application(tk.Frame):
    portCardList: List[SerialBoardCard] = list()
    listenToPorts = True
    fullScreenState = False

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.grid(row=0, column=0, sticky="nsew", pady=2, padx=2)

        # Load config from yaml
        self.config_object = configClass.OtoFlasherConfigObject()
        self.read_validate_yaml_config()

        # Setup gui widgets
        self.createWidgets()
        self.createPortCards()

        # Add Toggle and escape fullscreen mode hotkeys
        self.winfo_toplevel().bind("<F11>", self.toggleFullScreen)
        self.winfo_toplevel().bind("<Escape>", self.quitFullScreen)

        zope.event.subscribers.append(self.updateFlashAllButton)

    def updateFlashAllButton(self, event):
        """Update the status of the flash all button"""
        if event == EventType.UPDATE_FLASH_ALL:
            if any([True for x in self.portCardList if x.isBusy]):
                self.buttonFlashAll.disable()
            else:
                self.buttonFlashAll.enable()

    def disablePack(self, event):
        """Disable pack propogate for the portGuiWindow"""
        if event == EventType.DISABLE_PACK:
            self.guiWindow.pack_propagate(False)

    def enablePack(self, event):
        """Enable pack propogate for the portGuiWindow"""
        if event == EventType.ENABLE_PACK:
            self.guiWindow.pack_propagate(True)

    def createWidgets(self):

        self.grid_columnconfigure(0, weight=3)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=3)

        # Flash All Button
        self.buttonFlashAll = FlashAllButton(self, self.flashAll)
        self.buttonFlashAll.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)

        # Port Gui Window
        self.guiWindow = tk.Frame(self, borderwidth=2, relief=SUNKEN, bg="#e0e0e0", height=100)
        self.guiWindow.grid(row=1, column=0,  rowspan = 5, sticky="nsew", padx=1, pady=1)

        # Version Info Box
        self.versionBox = VersionBox(FIRMWARE_VERSION)
        self.versionBox.grid(row=6, column=0, sticky="ns", padx=1, pady=1)

    def read_validate_yaml_config(self):
        try:
            self.config_object.from_yaml_file()

        except FileNotFoundError:
            tkinter.messagebox.showerror(
                title = "File Not Found  文件未找到",
                message = f"File {self.config_object.yaml_file_path} not found.",
            )
            on_closing()
        except Exception:
            tkinter.messagebox.showerror(
                title = "Invalid File",
                message=f"{self.config_object.yaml_file_path} is not in a valid format.",
            )
            on_closing()

    def createPortCards(self):
        # Add serial port cards
        serialList = [x.serial for x in self.config_object.flasher_list]

        if serialList:
            for index, serialItem in enumerate(serialList):
                self.portCardList.append(
                    SerialBoardCard(
                        self.guiWindow,
                        flasherSerial=serialItem,
                        text=str(index + 1),
                        config_object=self.config_object,
                    )
                )
                root.minsize(width=len(self.portCardList) * 105, height=450)

    @threaded
    def flashAll(self):
        """Flash all connected COM Ports"""
        self.buttonFlashAll.disable()
        self.listenToPorts = True
        resultList = list()
        futureList: List[concurrent.futures.Future] = list()

        with concurrent.futures.ThreadPoolExecutor(max_workers=WORKERS, thread_name_prefix="flashing") as executor:
            for portGui in self.portCardList:
                futureList.append(executor.submit(portGui.flashButtonCallback))
                concurrent.futures.wait(futureList, timeout=1, return_when="ALL_COMPLETED")
        # Checking if an ADC reading has failed
        ADC_reading_failed = False
        for portGui in self.portCardList:
            if portGui.ADC_Failed is True:
                ADC_reading_failed = True
        resultList = [x.result() for x in futureList]
        index = 0
        portList = map(str, self.portCardList)
        for port, result in zip(portList, resultList):
            if ADC_reading_failed is False:
                if result is None:
                    self.portCardList[index].status = SerialBoardCard.PortStatus.SUCCESS
                    zope.event.notify(EventType.UPDATE_FLASH_ALL)
            else:
                self.portCardList[index].status = SerialBoardCard.PortStatus.CHECK_POWER
                zope.event.notify(EventType.UPDATE_FLASH_ALL)
                self.portCardList[index].logger.error("A BOARD HAS FAILED THE ADC READING. CHECK POWERSUPPLY IS SET TO 4.1V AND REFLASH")
            index += 1
        with open("readings.csv", "a", newline='') as csvfile:
            dataWriter = csv.writer(csvfile)
            for count, x in enumerate(self.portCardList):
                if self.portCardList[count].MACAddress is not None:
                    row = [self.portCardList[count].MACAddress, self.portCardList[count].adcReading, self.portCardList[count].ZeroPressure, self.portCardList[count].ZeroPressureSTD]
                    dataWriter.writerow(row)
        self.buttonFlashAll.enable()
        self.listenToPorts = True

    def getValidPorts(self, VID=None, PID=None):
        """Return a list of all ports matching given usb vid and pid"""
        validPorts: List[str] = list()
        allPorts = serial.tools.list_ports.comports()
        for port in allPorts:
            if (VID is None or port.vid == VID) and (PID is None or port.pid == PID):
                validPorts.append(port.name)
        return validPorts

    def toggleFullScreen(self, event):
        self.fullScreenState = not self.fullScreenState
        self.winfo_toplevel().attributes("-fullscreen", self.fullScreenState)

    def quitFullScreen(self, event):
        self.fullScreenState = False
        self.winfo_toplevel().attributes("-fullscreen", self.fullScreenState)

def on_closing():
    sys.exit()

if __name__ == "__main__":
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    root = tk.Tk()
    root.title("Flash Utility 0281(v5.3)-2024  存储实用程序0281(v5.3)-2024")
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)
    root.minsize(width=105, height=450)
    root.state("zoomed")
    app = Application(master=root)
    root.mainloop()

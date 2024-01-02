import concurrent.futures
import logging
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
from datetime import datetime
import ctypes
import csv
import numpy as np
import serial
import serial.tools.list_ports
import zope.event
import configClass
import pyoto.otoProtocol.otoCommands as pyoto
import pyoto.otoProtocol.otoMessageDefs as otoMessageDefs

CONFIG_YAML_PATH = "config.yml"
# -------- Other Settings --------
# USB VID and PID of OtO flasher board
VALID_VID = 0x10C4
VALID_PID = 0xEA60
# Max number of workers
WORKERS = 20
# lock = threading.Lock()
globalLoggingLevel = logging.INFO
# Set up logging
logging.basicConfig(level=globalLoggingLevel, format="%(message)s")
mainLogger = logging.getLogger(__name__)

OUTPUTMIN = 0.1 * (2**24)
ADCtokPa = 206.8427 / (0.8 * (2**24)) # 206.8427 kPa = 30 psi
TIMEINTERVAL = 120  # time in seconds to wait between samples
TOTALTIME = 15 * TIMEINTERVAL  # time in seconds to collect data over

def tokPa(ADC):
    return (ADC - OUTPUTMIN) * ADCtokPa

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
    UPDATE_ALL = auto()
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
        CONNECTING = auto()
        CHECK_PRESSURE = auto()
        CONNECTED = auto()
        SUCCESS = auto()
        FAIL = auto()
        FAIL_PRESSURE = auto()
        CONNECT_FLASHER = auto()
        WAITING = auto()

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
        self.Pressure_Failed = False
        self.PressureAve: float = 0
        self.PressureSTD: float = 0
        self.Pressures = []
        self.STDs = []
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
        elif new_status == SerialBoardCard.PortStatus.FAIL_PRESSURE:
            self._setStatusFailPressure()
        elif new_status == SerialBoardCard.PortStatus.CONNECT_FLASHER:
            self._setStatusConnectFlasher()
        elif new_status == SerialBoardCard.PortStatus.WAITING:
            self._setStatusWaiting()
        else:
            self.logger.warning(f"Invalid status: {new_status}")

    @property
    def isBusy(self):
        """Check if current status of port is busy"""
        busyStates = [
            SerialBoardCard.PortStatus.CONNECTING,
            SerialBoardCard.PortStatus.CONNECTED,
            SerialBoardCard.PortStatus.SUCCESS,
            SerialBoardCard.PortStatus.CHECK_PRESSURE,
            SerialBoardCard.PortStatus.WAITING
        ]

        # True if current state is one of the busy states
        if self.status in busyStates:
            return True
        return False

    @isBusy.setter
    def isBusy(self, new_busy):
        self.logger.warning(f"Cannot set isBusy to {new_busy}, Read-only Property")

    def ButtonCallback(self):
        """check pressure decay"""

        self.Pressure_Failed = False

        self.logger.info(f"Started Testing for {TOTALTIME/60} minutes, every {TIMEINTERVAL} seconds...")

        # STEP 1 Find COM port
        error = self.getSerialPortFromUSBSerial()
        if error is not None:
            self.logger.error(error)
            self.status = SerialBoardCard.PortStatus.CONNECT_FLASHER
            zope.event.notify(EventType.UPDATE_ALL)
            return error

        # STEP 2 Connect to OtO after making it reboot
        error = self.OtOConnect()
        if error is not None:
            self.logger.error(error)
            self.status = SerialBoardCard.PortStatus.FAIL
            zope.event.notify(EventType.UPDATE_ALL)
            return error

        # STEP 3 Get pressure sensor version so we can set appropriate limits
        error = self.getPressureSensorVersion()
        if error is not None:
            self.logger.error(error)
            self.status = SerialBoardCard.PortStatus.FAIL
            zope.event.notify(EventType.UPDATE_ALL)
            return error

        # Step 4 Pressure Check for x minutes
        MACName = self.MACAddress.replace(":", "-")
        self.Pressures.clear()
        self.STDs.clear()
        StartTime = time.time()
        t0 = StartTime
        t1 = StartTime + 1
        Duration = 0
        self.logger.info(self.MACAddress, datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f"))
        while Duration < TOTALTIME:
            if t0 > t1:
                error = self.PressureCheck(data_collection_time = 3.0)
                logtime = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")
                self.status = SerialBoardCard.PortStatus.WAITING
                zope.event.notify(EventType.UPDATE_ALL)
                if error is not None:
                    self.logger.error(error)
                    self.status = SerialBoardCard.PortStatus.FAIL_PRESSURE
                    zope.event.notify(EventType.UPDATE_ALL)
                    return error
                self.Pressures.append(self.PressureAve)
                self.STDs.append(self.PressureSTD)
                if Duration > 0:
                    AverageRate = round((self.Pressures[0] - self.PressureAve) / (Duration / 60), 3)
                    RateError = round((2.75 * (self.STDs[0] + self.PressureSTD)) / (Duration / 60), 4)
                with open(MACName + " readings.csv", "a", newline='') as csvfile:
                    dataWriter = csv.writer(csvfile)
                    row = [logtime, round(self.PressureAve, 4), round(self.PressureSTD * 2.75, 5), AverageRate, RateError]
                    dataWriter.writerow(row)
                    csvfile.close()
                self.logger.info(f"{round(Duration/60, 1)} minutes: {round(self.PressureAve, 2)}±{round(self.PressureSTD * 2.75, 3)} kPa, {AverageRate}±{RateError} kPa/hr")
                t1 = t0 + TIMEINTERVAL
            else:
                time.sleep(t1 - t0)
            t0 = time.time()
            Duration = t0 - StartTime
        error = self.PressureCheck(data_collection_time = 3.0)
        logtime = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")
        self.status = SerialBoardCard.PortStatus.WAITING
        zope.event.notify(EventType.UPDATE_ALL)
        if error is not None:
            self.logger.error(error)
            self.status = SerialBoardCard.PortStatus.FAIL_PRESSURE
            zope.event.notify(EventType.UPDATE_ALL)
            return error
        self.Pressures.append(self.PressureAve)
        self.STDs.append(self.PressureSTD)
        AverageRate = round((self.Pressures[0] - self.PressureAve) / (TOTALTIME / 60), 3)
        RateError = round((2.75 * (self.STDs[0] + self.PressureSTD)) / (TOTALTIME / 60), 4)
        with open(MACName + " readings.csv", "a", newline='') as csvfile:
            dataWriter = csv.writer(csvfile)
            row = [logtime, round(self.PressureAve, 4), round(self.PressureSTD * 2.75, 5), AverageRate, RateError]
            dataWriter.writerow(row)
            csvfile.close()
        self.logger.info(f"{round(TOTALTIME/60, 1)} minutes: {round(self.PressureAve, 2)}±{round(self.PressureSTD * 2.75, 3)} kPa, {AverageRate}±{RateError} kPa/hr")
        self.logger.info("Test complete.")
        return None

    def _setStatusIdle(self):
        self.configure(background=self.IDLE_COLOR)
        self._status = SerialBoardCard.PortStatus.IDLE
        self.labelPortName["bg"] = self.IDLE_COLOR
        self.labelStatus["text"] = "Idle 闲置中"
        self.labelStatus["bg"] = self.IDLE_COLOR

    def _setStatusWaiting(self):
        self.configure(background=self.IDLE_COLOR)
        self._status = SerialBoardCard.PortStatus.WAITING
        self.labelPortName["bg"] = self.IDLE_COLOR
        self.labelStatus["text"] = "Waiting"
        self.labelStatus["bg"] = self.IDLE_COLOR

    def _setStatusSuccess(self):
        self.configure(background=self.OK_COLOR)
        self._status = SerialBoardCard.PortStatus.SUCCESS
        self.labelPortName["bg"] = self.OK_COLOR
        self.labelStatus["text"] = "DONE 完成"
        self.labelStatus["bg"] = self.OK_COLOR

    def _setStatusFail(self):
        self.configure(background=self.ERROR_COLOR)
        self._status = SerialBoardCard.PortStatus.FAIL
        self.labelPortName["bg"] = self.ERROR_COLOR
        self.labelStatus["text"] = "Test Error!"
        self.labelStatus["bg"] = self.ERROR_COLOR

    def _setStatusFailPressure(self):
        self.configure(background=self.ERROR_COLOR)
        self._status = SerialBoardCard.PortStatus.FAIL_PRESSURE
        self.labelPortName["bg"] = self.ERROR_COLOR
        self.labelStatus["text"] = "Failed Pressure Sensor\n压力传感器故障"
        self.labelStatus["bg"] = self.ERROR_COLOR

    def _setStatusConnecting(self):
        self._status = SerialBoardCard.PortStatus.CONNECTING
        self.labelStatus["text"] = "Connecting 连接"

    def _setStatusCheckPressure(self):
        self.configure(background=self.BUSY_COLOR)
        self._status = SerialBoardCard.PortStatus.CHECK_PRESSURE
        self.labelPortName["bg"] = self.BUSY_COLOR
        self.labelStatus["text"] = "Pressure Check"
        self.labelStatus["bg"] = self.BUSY_COLOR

    def _setStatusConnected(self):
        self._status = SerialBoardCard.PortStatus.CONNECTED
        self.labelStatus["text"] = "Connected 连接成功"

    def _setStatusConnectFlasher(self):
        self.configure(background=self.ERROR_COLOR)
        self._status = SerialBoardCard.PortStatus.CONNECT_FLASHER
        self.labelPortName["bg"] = self.ERROR_COLOR
        self.labelStatus["text"] = "CONNECT FLASHER\n连接USB通信测试板"
        self.labelStatus["bg"] = self.ERROR_COLOR

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

    def PressureCheck(self, data_collection_time: float):
        self.status = SerialBoardCard.PortStatus.CHECK_PRESSURE
        Sensor_Read_List: List[pyoto.otoMessageDefs.SensorReadMessage] = []
        pressureReading: list = []
        self.PressureAve = 0
        self.PressureSTD = 0
        self.pyoto_instance.set_valve_duty(direction = 1, duty_cycle = 100)
        self.pyoto_instance.set_sensor_subscribe(subscribe_frequency=pyoto.SensorSubscribeFrequencyEnum.SENSOR_SUBSCRIBE_FREQUENCY_100Hz)
        time.sleep(0.1)
        self.pyoto_instance.clear_incoming_packet_log()
        main_loop_start_time = time.time()
        while time.time() - main_loop_start_time <= data_collection_time:
            Sensor_Read_List.extend(self.pyoto_instance.read_all_sensor_packets(limit=None, consume=True))
        self.pyoto_instance.set_sensor_subscribe(subscribe_frequency=pyoto.SensorSubscribeFrequencyEnum.SENSOR_SUBSCRIBE_FREQUENCY_OFF)
        self.pyoto_instance.set_valve_duty(direction = 0, duty_cycle = 0)
        if not Sensor_Read_List:
            return "No pressure data was collected.\n未收集压力数值"
        for message in Sensor_Read_List:
            pressureReading.append(int(message.pressure_adc))
        self.PressureAve = round(tokPa(np.mean(pressureReading)), 4)
        self.PressureSTD = round(ADCtokPa * np.std(pressureReading), 5)
        return None

    def OtOConnect(self):
        try:
            self.pyoto_instance = pyoto.OtoInterface(connection_type=pyoto.ConnectionType.UART, logger=None)
            # self.pyoto_instance.logger.setLevel(logging.INFO)
            self.logger.info("Waiting for board to reboot...\n等待线路板重启")
            self.pyoto_instance.start_connection(port=self.port, reset_on_connect=True)
            self.status = SerialBoardCard.PortStatus.CONNECTED
            self.MACAddress = self.pyoto_instance.get_mac_address().string
            self.pyoto_instance.use_moving_average_filter(True)
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
        elif returned_pressure_sensor_version == otoMessageDefs.PressureSensorVersionEnum.TPBD_15_PSI_GAUGE.value or returned_pressure_sensor_version == otoMessageDefs.PressureSensorVersionEnum.MPRL_15_PSI_GAUGE.value:
            return "This is an old design board and cannot be tested on this station."
        elif returned_pressure_sensor_version == otoMessageDefs.PressureSensorVersionEnum.MPRL_30_PSI_GAUGE.value:
            # self.logger.info("30 psi pressure sensor detected\n检测到0.21MPa压力传感器")
            self.GaugeRange = 206.8427  # 206.8427 kPa = 30 psi
            self.max_acceptable_STD: float = 206.9  # Jan 2023 ±4σ
            self.min_acceptable_STD: float = 66.5  # Jan 2023 ±4σ
            self.max_acceptable_ADC: float = 1764145  # Jan 2023 ±4σ 1764145
            self.min_acceptable_ADC: float = 1630925  # Jan 2023 ±4σ 1630925
            return None
        else:
            return "Unknown pressure sensor detected\n检查到未知压力传感器"

class AllButton(tk.Button):
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
            text = "Start Decay Test",
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
        Version,
    ):
        tk.LabelFrame.__init__(self, text="Version")
        labelText = f"Version: {Version}"
        self.labelScriptVersion = tk.Label(self, text=labelText)
        self.labelScriptVersion.pack(side="top")

class Application(tk.Frame):
    portCardList: List[SerialBoardCard] = list()
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

        zope.event.subscribers.append(self.updateAllButton)

    def updateAllButton(self, event):
        """Update the status of the all button"""
        if event == EventType.UPDATE_ALL:
            if any([True for x in self.portCardList if x.isBusy]):
                self.ButtonAll.disable()
            else:
                self.ButtonAll.enable()

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

        # Test All Button
        self.ButtonAll = AllButton(self, self.TestAll)
        self.ButtonAll.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)

        # Port Gui Window
        self.guiWindow = tk.Frame(self, borderwidth=2, relief=SUNKEN, bg="#e0e0e0", height=100)
        self.guiWindow.grid(row=1, column=0,  rowspan = 5, sticky="nsew", padx=1, pady=1)

        # Version Info Box
        self.versionBox = VersionBox("v0.5")
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
                        config_object=self.config_object
                    )
                )
                root.minsize(width=len(self.portCardList)*105, height=450)
        else:
            tkinter.messagebox.showerror(
                title = "Invalid File",
                message=f"{self.config_object.yaml_file_path} is not in a valid format.",
            )
            on_closing()

    @threaded
    def TestAll(self):
        """Test all connected COM Ports"""
        self.ButtonAll.disable()
        resultList = list()
        futureList: List[concurrent.futures.Future] = list()

        with concurrent.futures.ThreadPoolExecutor(max_workers=WORKERS, thread_name_prefix="testing") as executor:
            for portGui in self.portCardList:
                futureList.append(executor.submit(portGui.ButtonCallback))
                concurrent.futures.wait(futureList, timeout = None, return_when = "ALL_COMPLETED")
        resultList = [x.result() for x in futureList]
        index = 0
        portList = map(str, self.portCardList)
        for port, result in zip(portList, resultList):
            if result is None:
                self.portCardList[index].status = SerialBoardCard.PortStatus.SUCCESS
                zope.event.notify(EventType.UPDATE_ALL)
            index += 1
        self.ButtonAll.enable()

    def getValidPorts(self, VID=None, PID=None):
        """Return a list of all ports matching given usb vid and pid"""
        validPorts: List[str] = list()
        allPorts = serial.tools.list_ports.comports()
        for port in allPorts:
            if (VID is None or port.vid == VID) and (PID is None or port.pid == PID):
                validPorts.append(port.name)
        return validPorts

def on_closing():
    sys.exit()

if __name__ == "__main__":
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    root = tk.Tk()
    root.title("OtO Decay Test - 2024")
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)
    root.minsize(width=105, height=450)
    root.state("zoomed")
    app = Application(master=root)
    root.mainloop()

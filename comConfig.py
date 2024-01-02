import sys
import threading
import time
import tkinter as tk
import tkinter.constants
import tkinter.messagebox
from pathlib import Path
from pprint import pprint
from tkinter import ttk
from typing import List

import serial
import serial.tools.list_ports
import yaml
import ctypes

import configClass

# -------- Basic Settings --------

CONFIG_YAML_PATH = "config.yml"


# -------- Other Settings --------

# USB VID and PID of OtO flasher board
VALID_VID = 0x10C4
VALID_PID = 0xEA60

# Max number of workers
WORKERS = 20

# Port Listener Speed
PORT_LISTENER_RATE_HZ = 4

lock = threading.Lock()

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

class ConfigObject:
    """Holds the info stored in the config.yaml file in a structured python class"""

    def __init__(
        self,
        yaml_object: dict = None,
        flasher_list: list = None
    ) -> None:
        if yaml_object is not None:
            self.flasher_list: list = yaml_object.get("flasher_list")
        else:
            self.flasher_list = flasher_list

    def to_dict(self):

        return_dict = dict()
        if self.flasher_list is not None:
            return_dict["flasher_list"] = self.flasher_list

        return return_dict

class configureCOMPort(tk.Frame):
    """GUI to configure COM port order"""

    listenToPorts = True
    stopListenerThread_flag = False

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.loadedSerialList = list()
        self.config_object = configClass.OtoFlasherConfigObject()
        self.createWidgets()
        self.setup_from_config_yaml()
        self.startPortListener(VALID_VID, VALID_PID)
        self.grid(row=0, column=0, sticky="nsew", pady=0, padx=0)

    def okButtonCallback(self):
        """Callback function for OK Button"""

        self.stopListenerThread_flag = True
        self.portListenerThread.join()

        with lock:
            self.set_config_object_to_current_state()
            self.write_current_state_to_yaml()
            self.destroyPopUp()

    def applyButtonCallback(self):
        """Callback function for Apply Button"""

        with lock:
            self.set_config_object_to_current_state()
            self.write_current_state_to_yaml()
            self.clearItems()
            self.setup_from_config_yaml()

    def createWidgets(self):
        self.grid_columnconfigure(0, weight=3, minsize=10)
        self.grid_columnconfigure(1, weight=0, minsize=10)
        self.grid_rowconfigure(0, weight=2, minsize=0)
        self.grid_rowconfigure(1, weight=2, minsize=0)
        self.grid_rowconfigure(2, weight=0, minsize=30)

        self.FlasherSerialTreeView = ttk.Treeview(
            self, selectmode="browse", columns=("index", "serial", "full_serial")
        )
        self.FlasherSerialTreeView.column(
            "index", width=80, anchor=tkinter.constants.CENTER
        )
        self.FlasherSerialTreeView.heading("index", text="Index")
        self.FlasherSerialTreeView.column(
            "serial", width=160, anchor=tkinter.constants.CENTER
        )
        self.FlasherSerialTreeView.heading("serial", text="Serial")
        self.FlasherSerialTreeView.column("full_serial", width=400)
        self.FlasherSerialTreeView.heading("full_serial", text="Full Serial")
        self.FlasherSerialTreeView["show"] = "headings"

        self.FlasherSerialTreeView.tag_configure("green_bg", background="#94ffa8")

        self.FlasherSerialTreeView.grid(
            sticky="nsew",
            row=0,
            column=0,
            padx=(10, 0),
            pady=10,
            rowspan=2,
            columnspan=1,
        )

        self.FlasherSerialTreeView.bind("<Button-1>", self.onTreeviewClick)

        self.moveUpBtn = ttk.Button(self, text="Up", command=self.moveItemUp)
        self.moveUpBtn.grid(row=0, column=1, sticky="nsew", padx=10, pady=(10, 5))
        self.moveDownBtn = ttk.Button(self, text="Down", command=self.moveItemDown)
        self.moveDownBtn.grid(row=1, column=1, sticky="nsew", padx=10, pady=(5, 10))

        # Options frame, contains bom entry, database selector, firmware version selector
        self.optionsFrame = ttk.Frame(self)
        self.optionsFrame.grid(sticky="nsew", row=2, column=0, padx=0, pady=0, rowspan=1, columnspan=2)

        # Bottom Button Frame, contains apply and okay button
        self.bottomButtonFrame = ttk.Frame(self)
        self.bottomButtonFrame.grid(
            sticky="nsew", row=3, column=0, padx=0, pady=0, rowspan=1, columnspan=2
        )
        self.bottomButtonFrame.grid_columnconfigure(0, weight=1)
        self.bottomButtonFrame.grid_columnconfigure(1, weight=1)
        self.bottomButtonFrame.grid_rowconfigure(0, weight=1)

        s.configure("apply.TButton", font=("Microsoft YaHei UI", 10, "bold"))
        self.ApplyBtn = ttk.Button(
            self.bottomButtonFrame,
            text="Apply",
            command=self.applyButtonCallback,
            style="apply.TButton",
        )
        self.ApplyBtn.grid(row=0, column=0, sticky="nsew", padx=(10, 5), pady=(0, 10))
        s.configure("okay.TButton", font=("Helvetica", 10, "bold"))
        self.okayBtn = ttk.Button(
            self.bottomButtonFrame,
            text="Okay",
            command=self.okButtonCallback,
            style="okay.TButton",
        )
        self.okayBtn.grid(row=0, column=1, sticky="nsew", padx=(5, 10), pady=(0, 10))

    def onTreeviewClick(self, event):
        if not self.FlasherSerialTreeView.identify_row(event.y):
            self.FlasherSerialTreeView.selection_remove(
                self.FlasherSerialTreeView.selection()
            )

    def set_config_object_to_current_state(self):
        """Sets self.config_object to current GUI state"""
        allPorts = serial.tools.list_ports.comports()
        flasher_list: List = list()
        for serialNumber in list(self.FlasherSerialTreeView.get_children()):
            for port in allPorts:
                if (
                    (port.vid == VALID_VID)
                    and (port.pid == VALID_PID)
                    and (port.serial_number == serialNumber)
                ):
                    flasher_list.append(
                        configClass.OtoFlasherObject(
                            vid=str(hex(port.vid)),
                            pid=str(hex(port.pid)),
                            serial=str(port.serial_number),
                        )
                    )

        self.config_object.flasher_list = flasher_list

    def getValidSerialNumbers(self, VID=None, PID=None):
        """Return a list of all port serial number matching given usb vid and pid"""
        validPorts: List[str] = list()
        allPorts = serial.tools.list_ports.comports()
        for port in allPorts:
            if (VID is None or port.vid == VID) and (PID is None or port.pid == PID):
                validPorts.append(port.serial_number)
        return validPorts

    def moveItemUp(self, *args):
        """Callback Function for Up Button"""

        selected_item = self.FlasherSerialTreeView.selection()
        if not selected_item:
            return

        selected_index = self.FlasherSerialTreeView.index(selected_item)
        max_index = len(self.FlasherSerialTreeView.get_children())
        new_index = min(max_index, max(0, selected_index - 1))
        self.FlasherSerialTreeView.move(selected_item, "", new_index)
        self.updateIndexColumns()

    def moveItemDown(self, *args):
        """Callback Function for Down Button"""

        selected_item = self.FlasherSerialTreeView.selection()
        if not selected_item:
            return
        selected_index = self.FlasherSerialTreeView.index(selected_item)
        max_index = len(self.FlasherSerialTreeView.get_children())
        new_index = min(max_index, max(0, selected_index + 1))
        self.FlasherSerialTreeView.move(selected_item, "", new_index)
        self.updateIndexColumns()

    def clearItems(self):
        """Clears all items from the treeview"""

        for item in self.FlasherSerialTreeView.get_children():
            self.FlasherSerialTreeView.delete(item)

        self.updateIndexColumns()

    def portListener(self, VID, PID, *args):
        """Listens to COM ports and updates portList as necessary"""

        while not self.stopListenerThread_flag:

            if self.listenToPorts:

                newSerialList = self.getValidSerialNumbers(VID, PID)

                addedSerials = [
                    serial_number
                    for serial_number in newSerialList
                    if serial_number
                    not in list(self.FlasherSerialTreeView.get_children())
                ]

                removedSerials = [
                    serial_number
                    for serial_number in list(self.FlasherSerialTreeView.get_children())
                    if serial_number not in newSerialList
                ]

                # Update list of com ports if ports were changed
                if addedSerials or removedSerials:
                    with lock:
                        for serial_number in addedSerials:

                            self.insertItem(serial_number)

                        for serial_number in removedSerials:

                            # Remove item from Treeview
                            self.FlasherSerialTreeView.delete(serial_number)

                    self.updateIndexColumns()

            time.sleep(1 / PORT_LISTENER_RATE_HZ)

    def insertItem(self, serial_number):
        """Inserts row into the treeview"""

        if serial_number not in self.loadedSerialList:
            tags = ["green_bg"]
        else:
            tags = []

        # Insert new item into Treeview
        self.FlasherSerialTreeView.insert(
            parent="",  # root level item
            index=tkinter.constants.END,  # adds to end of list
            iid=serial_number,  # item id set to serial number
            text=serial_number,  # first column set to shortened serial number
            values=(
                0,
                serial_number[:7],
                serial_number,
            ),  # second column set to full serial number
            tags=tags,
        )
        self.updateIndexColumns()

    def updateIndexColumns(self):
        """Applies the correct index number to each item in treeview"""
        for item in self.FlasherSerialTreeView.get_children():
            self.FlasherSerialTreeView.set(
                item, column="index", value=self.FlasherSerialTreeView.index(item) + 1
            )

    def removeTreeviewItem(self, serial_number):

        if serial_number not in self.loadedSerialList:
            tags = ["green_bg"]
        else:
            tags = []

        # Insert new item into Treeview
        self.FlasherSerialTreeView.insert(
            parent="",  # root level item
            index=tkinter.constants.END,  # adds to end of list
            iid=serial_number,  # item id set to serial number
            text=serial_number[:7],  # first column set to shortened serial number
            values=(serial_number),  # second column set to full serial number
            tags=tags,
        )

    def startPortListener(self, VID: str, PID: str):
        """Runs portListener function as a daemon thread"""
        self.portListenerThread = threading.Thread(
            target=self.portListener, args=(VID, PID)
        )
        self.portListenerThread.daemon = True
        self.portListenerThread.start()

    def setup_from_config_yaml(self):
        """Sets up GUI components with data read from config yaml"""
        # Read from yaml file
        try:
            self.config_object.from_yaml_file()
        except FileNotFoundError:
            pass

        # Set Flasher list
        if self.config_object.flasher_list:
            self.loadedSerialList = [x.serial for x in self.config_object.flasher_list]
        print(self.config_object.flasher_list)
        pprint(self.loadedSerialList)
        for item in self.loadedSerialList:
            self.insertItem(item)

    def write_current_state_to_yaml(self):
        """Writes self.config_object to yaml file"""

        with open(CONFIG_YAML_PATH, "w") as file_handler:
            yaml.dump(self.config_object.to_dict(), file_handler)

    def destroyPopUp(self):
        configurePopUp.withdraw()
        configurePopUp.destroy()

    def hideWindowself(self):
        configurePopUp.withdraw()


def on_closing():
    sys.exit()


if __name__ == "__main__":

    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    configurePopUp = tk.Tk()

    # ====== TKINTER PATCH from https://bugs.python.org/issue36468 ======
    s = ttk.Style()
    # from os import name as OS_Name
    # if configurePopUp.getvar("tk_patchLevel") == "8.6.9":  # and OS_Name=='nt':

    #     def fixed_map(option):
    #         return [
    #             elm
    #             for elm in s.map("Treeview", query_opt=option)
    #             if elm[:2] != ("!disabled", "!selected")
    #         ]

    #     s.map(
    #         "Treeview",
    #         foreground=fixed_map("foreground"),
    #         background=fixed_map("background"),
    #     )

    # ====== TKINTER PATCH END ======

    configurePopUp.attributes("-topmost", "true")
    configurePopUp.title("OtO Flasher Configuration")
    configurePopUp.protocol("WM_DELETE_WINDOW", on_closing)
    configurePopUp.grid_rowconfigure(0, weight=1)
    configurePopUp.grid_columnconfigure(0, weight=1)
    configurePopUp.minsize(width=640, height=480)
    popUP = configureCOMPort(master=configurePopUp)

    configurePopUp.mainloop()

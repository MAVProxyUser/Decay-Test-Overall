import logging
import os
from typing import List
import yaml


class OtoFlasherConfigObject:
    def __init__(self, logger=None) -> None:
        self.base_url = None
        self.flasher_list: List[OtoFlasherObject] = list()
        self.bom_number = None
        self.yaml_file_path = "config.yml"

        if isinstance(logger, logging.Logger):
            self.logger = logger
        else:
            self.logger = logging.getLogger(__name__)

    def from_yaml_file(self):
        """Loads and from a given yaml file
        If yaml file doesn't have a field for an attribute, the attribute will
        be set to None
        Args:       None
        Raises:     FileNotFoundError if file doesn't exist
                    Exception on yaml file read/parse error"""

        with open(self.yaml_file_path, "r") as file_handler:
            self.logger.info(f"Reading yaml from {os.path.realpath(file_handler.name)} ...")
            yaml_object = yaml.load(file_handler, Loader=yaml.SafeLoader)

        self.flasher_list = list()

        # fill out attributes, sets them to none if they don't exist
        if isinstance(yaml_object, dict):
            flasher_dict_list = yaml_object.get("flasher_list")
            if isinstance(flasher_dict_list, list):
                for flasher_dict in flasher_dict_list:
                    current_flasher_object = OtoFlasherObject()
                    current_flasher_object.from_dict(flasher_dict)
                    self.flasher_list.append(current_flasher_object)
            self.bom_number = yaml_object.get("bom_number")
            self.base_url = yaml_object.get("base_url")

    def to_yaml_file(self):
        """Writes to yaml file
        If attribute is none, the field will not be written
        To change file path, edit yaml_file_path attribute
        Args:       None
        Raises:     Exception on write/serialize error"""

        with open(self.yaml_file_path, "w") as file_handler:
            self.logger.info(f"Saving yaml to {os.path.realpath(file_handler.name)} ...")
            self.logger.info(yaml.dump(self.to_dict()))
            yaml.dump(self.to_dict(), file_handler)

    def to_dict(self):
        """Returns a dict representation of self"""
        return_dict = dict()
        if self.flasher_list:
            return_dict["flasher_list"] = [flasher.to_dict() for flasher in self.flasher_list]

        if self.bom_number is not None:
            return_dict["bom_number"] = self.bom_number

        if self.base_url is not None:
            return_dict["base_url"] = self.base_url

        return return_dict


class OtoFlasherObject:
    def __init__(self, vid: str = None, pid: str = None, serial: str = None) -> None:
        self.vid = vid
        self.pid = pid
        self.serial = serial

    def to_dict(self):
        """To dict
        Args:       None
        Returns:    dict representing self
        """
        return self.__dict__

    def from_dict(self, new_dict: dict = None):
        """From dict
        Args:       new_dict
        Returns:    None
        """
        if new_dict is None:
            self.vid = None
            self.pid = None
            self.serial = None
            return

        self.vid = new_dict.get("vid")
        self.pid = new_dict.get("pid")
        self.serial = new_dict.get("serial")
        return

    def __eq__(self, other):
        if isinstance(other, OtoFlasherObject):
            return self.pid == other.pid and self.vid == other.vid and self.serial == other.serial
        return False

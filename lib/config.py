import re
from lib.logging import log

class Config():

    ################################################################# SURCHARGE

    def __init__(self, path:str=None) -> None:
        self._dict = None        # {key1: value1, ...}

        if path:
            self.load(path)

    def __str__(self) -> str:
        substrings = []
        for attribute, value in vars(self).items():
            substrings.append(f"{attribute}: {str(value)}")
        return "\n".join(substrings)

    ################################################################### GETTERS

    @property
    def dict(self) -> str:
        return self._dict
    
    ################################################################### SETTERS

    @dict.setter
    def dict(self, dict:dict) -> None:
        self._dict = dict

    ################################################################### METHODS

    def load(self, path:str="./config.txt") -> bool:
        dictionary = {}
        with open(path, 'r') as file:
            for line in file:
                line = line.strip()
                if not line or line.startswith("#"): 
                    continue
                match = re.match(r'^\s*(.*?)\s*=\s*(.*?)\s*$', line)
                if match:
                    key, value = match.groups()
                    dictionary[key.strip()] = value.strip()
                else:
                    log(f"Unable to parse the line \"{line}\" of the config file \"{path}\". Use \"#\" for comments and \"key1 = value1\" for configurations.", "error")
                    # Quit the method with an error code
                    return False
        # Save the configurations
        self.dict = dictionary
        log(f"Configuration loaded: {self.dict}", "info")
        # Quit the method with a success code
        return True

    def get(self, key:str) -> str:
        if not self.dict:
            log(f"Unable to get a configuration value without loading a configuration file first.", "error")
            return
        if key in self.dict.keys():
            return self.dict[key]
        else:
            log(f"Unable to get the configuration for \"{key}\".", "error")
            return
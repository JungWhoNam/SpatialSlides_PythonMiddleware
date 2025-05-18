from enum import Enum, auto


class ServerAction(Enum):
    NO_ACTION = auto()
    SEND_CURRENT_VIEWS = auto()
    SEND_ALL_VIEWS = auto()
    SEND_CURRENT_MODE = auto()

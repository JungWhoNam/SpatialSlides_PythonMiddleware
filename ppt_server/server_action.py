from enum import Enum, auto


class ServerAction(Enum):
    NO_ACTION = auto()
    SEND_CURRENT_VIEWS = auto()

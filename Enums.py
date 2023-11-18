from enum import Enum


class EnableType(Enum):
    Unable = 0
    Enable = 1


class ColType(Enum):
    Enable = 1
    PicName = 2
    TimeOut = 3
    TimeOutAction = 4
    Interval = 5
    Action = 6


class STYPE(Enum):
    Empty = 0
    Str = 1
    Num = 2


class RunType(Enum):
    Wait = -1
    Stop = 0
    Running = 1

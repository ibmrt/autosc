# Get window size
class StartAcquireWindow:
    def __repr__(self) -> str:
        return "StartAcquireWindow"

class WindowSizeResult:
    def __init__(self, left:int, top:int, width:int, height:int):
        self.left = left
        self.top = top
        self.width = width
        self.height = height

    def __repr__(self) -> str:
        return f"WindowSizeResult({self.left}, {self.top}, {self.width}, {self.height})"

class SaveWindowSizeFailed:
    def __init__(self, reason):
        self.reason = reason

    def __repr__(self) -> str:
        return f"SaveWindowSizeFailed({self.reason})"

# Set path
class SaveFolderPath:
    def __init__(self, path:str):
        self.path = path

    def __repr__(self) -> str:
        return f"SaveFolderPath({self.path})"

class PathSaved:
    def __init__(self, path:str):
        self.path = path

    def __repr__(self) -> str:
        return f"PathSaved({self.path})"

class SavePathFailed:
    def __init__(self, reason:str):
        self.reason = reason

    def __repr__(self) -> str:
        return f"SavePathFailed({self.reason})"

# save setting
class SaveSettings:
    def __repr__(self) -> str:
        return "SaveSettings"

class SettingSaved:
    def __init__(self, path, window_size):
        self.path = path
        self.window_size = window_size

    def __repr__(self) -> str:
        return f"SettingSaved({self.path}, {self.window_size})"

class SaveSettingFailed:
    def __init__(self, reason:str):
        self.reason = reason

    def __repr__(self) -> str:
        return f"SaveSettingFailed({self.reason})"

# process
class StartProcess:
    def __init__(self, process):
        self.process = process

    def __repr__(self) -> str:
        return f"StartProcess({self.process})"

class StartedProcess:
    def __init__(self, path):
        self.path = path

    def path(self):
        return self.path

    def __repr__(self) -> str:
        return f"{self.path}"

class FailedToStartProcess:
    def __init__(self, reason:str):
        self.reason = reason

    def __repr__(self) -> str:
        return f"FailedToStartProcess({self.reason})"

# pptx
class StartCreatePPT:
    def __init__(self, path:str):
        self.path = path

    def __repr__(self) -> str:
        return f"StartCreatePPT({self.path})"

class PPTCreated:
    def __init__(self, path:str):
        self.path = path

    def __repr__(self) -> str:
        return f"PPTCreated({self.path})"

class FailedToCreatePPT:
    def __init__(self, reason:str):
        self.reason = reason

    def __repr__(self) -> str:
        return f"FailedToCreatePPT({self.reason})"
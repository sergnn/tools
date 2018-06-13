import platform
import re
import win32api
from collections import namedtuple

import win32com.client
from winnt import PROCESSOR_ARCHITECTURE_AMD64

VER_NT_WORKSTATION = 0x0000001  # The operating system is Windows 8, Windows 7, Windows Vista, Windows XP Professional, Windows XP Home Edition, or Windows 2000 Professional.
SM_SERVERR2 = 89  # The build number if the system is Windows Server 2003 R2; otherwise, 0.

WinName = namedtuple('WinName', ['full', 'short', 'server'])


class WinVer(object):
    def __init__(self):
        _version = platform.version().split('.')
        if len(_version) == 2:
            self.major = int(_version[0])
            self.minor = int(_version[1])
            self.build = 0
        else:
            self.major = int(_version[0])
            self.minor = int(_version[1])
            self.build = int(_version[2])
        self.wProductType = win32api.GetVersionEx(1)[8]
        self.wProcessorArchitecture = win32api.GetSystemInfo()[0]
        self.bits = 64 if platform.machine().endswith('64') else 32
        _v = self.version
        self.version_full = _v.full
        self.version_short = _v.short
        self.server = _v.server

    @property
    def dict(self):
        return {k: v for k, v in self.__dict__.items() if not k.startswith('_')}

    @property
    def version(self):
        if self.major == 5 and self.minor == 0:
            return WinName('2000', '2000', False)

        if self.major == 5 and self.minor == 1:
            return WinName('XP', 'XP', False)

        if self.major == 5 and self.minor == 2 \
                and self.wProductType == VER_NT_WORKSTATION \
                and self.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64:
            return WinName('XP Professional x64 Edition', 'XP', False)

        if self.major == 5 and self.minor == 2 \
                and win32api.GetSystemMetrics(SM_SERVERR2) == 0:
            return WinName('Server 2003', '2003', True)

        if self.major == 5 and self.minor == 2 \
                and win32api.GetSystemMetrics(SM_SERVERR2) != 0:
            return WinName('Server 2003 R2', '2003R2', True)

        if self.major == 6 and self.minor == 0 \
                and self.wProductType == VER_NT_WORKSTATION:
            return WinName('Vista', 'Vista', False)

        if self.major == 6 and self.minor == 0 \
                and self.wProductType != VER_NT_WORKSTATION:
            return WinName('Server 2008', '2008', True)

        if self.major == 6 and self.minor == 1 \
                and self.wProductType != VER_NT_WORKSTATION:
            return WinName('Server 2008 R2', '2008R2', True)

        if self.major == 6 and self.minor == 1 \
                and self.wProductType == VER_NT_WORKSTATION:
            return WinName('7', '7', False)

        if self.major == 6 and self.minor == 2 \
                and self.wProductType != VER_NT_WORKSTATION:
            return WinName('Server 2012', '2012', True)

        if self.major == 6 and self.minor == 2 \
                and self.wProductType == VER_NT_WORKSTATION:
            return WinName('8', '8', False)

        if self.major == 6 and self.minor == 3 \
                and self.wProductType != VER_NT_WORKSTATION:
            return WinName('Server 2012 R2', '2012R2', True)

        if self.major == 6 and self.minor == 3 \
                and self.wProductType == VER_NT_WORKSTATION:
            return WinName('8.1', '8.1', False)

        if self.major == 10 and self.minor == 0 \
                and self.wProductType != VER_NT_WORKSTATION:
            return WinName('Server 2016', '2016', True)

        if self.major == 10 and self.minor == 0 \
                and self.wProductType == VER_NT_WORKSTATION:

            if self.build <= 10240:
                return WinName('10 Threshold 1', '10', False)
            elif self.build <= 10586:
                return WinName('10 Threshold 2', '10', False)
            elif self.build <= 14393:
                return WinName('10 Redstone 1', '10RS1', False)
            elif self.build <= 15063:
                return WinName('10 Redstone 2', '10RS2', False)
            elif self.build <= 16299:
                return WinName('10 Redstone 3', '10RS3', False)
            elif self.build <= 17134:
                return WinName('10 Redstone 4', '10RS4', False)
            else:
                return WinName('10 Redstone 5', '10RS5', False)


def enum_winupdates():
    wua = win32com.client.Dispatch('Microsoft.Update.Session')
    update_searcher = wua.CreateUpdateSearcher()
    update_searcher.IncludePotentiallySupersededUpdates = True
    history_count = update_searcher.GetTotalHistoryCount()
    search_result = update_searcher.QueryHistory(0, history_count)
    updates = []
    for update in search_result:
        if update.ResultCode == 2 and update.Operation == 1:
            try:
                result = re.search(r'KB\d+', update.Title)
                updates.append(result.group(0))
            except AttributeError:
                pass
    return updates

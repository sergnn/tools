import os
import re
import winreg

import win32com.client


def win_bits() -> int:
    return 64 if "PROGRAMFILES(X86)" in os.environ else 32


def win_release_id() -> int:
    key = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    val = r"ReleaseID"
    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key)
    releaseId = int(winreg.QueryValueEx(key, val)[0])
    winreg.CloseKey(key)
    return releaseId


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


patches = {1709: {'Spectre&Meltdown': 'KB4051892', 'Petya': 'KB0000'}, }

if __name__ == '__main__':
    required = patches[win_release_id()]
    installed = enum_winupdates()
    print([(v, p) for v, p in required.items() if p not in installed])

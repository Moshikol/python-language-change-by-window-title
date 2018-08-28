
# using pywin32 for constants and ctypes for everything else seems a little
import win32con
import win32gui
import win32api
import py_win_keyboard_layout


import sys
import ctypes
import ctypes.wintypes

user32 = ctypes.windll.user32
ole32 = ctypes.windll.ole32
kernel32 = ctypes.windll.kernel32

WinEventProcType = ctypes.WINFUNCTYPE(
    None,
    ctypes.wintypes.HANDLE,
    ctypes.wintypes.DWORD,
    ctypes.wintypes.HWND,
    ctypes.wintypes.LONG,
    ctypes.wintypes.LONG,
    ctypes.wintypes.DWORD,
    ctypes.wintypes.DWORD
)


# The types of events we want to listen for, and the names we'll use for
# them in the log output. Pick from
# http://msdn.microsoft.com/en-us/library/windows/desktop/dd318066(v=vs.85).aspx
eventTypes = {
    # win32con.EVENT_SYSTEM_FOREGROUND: "Foreground",
    win32con.EVENT_OBJECT_FOCUS: "Focus",
}

# limited information would be sufficient, but our platform doesn't have it.
processFlag = getattr(win32con, 'PROCESS_QUERY_LIMITED_INFORMATION',
                      win32con.PROCESS_QUERY_INFORMATION)

threadFlag = getattr(win32con, 'THREAD_QUERY_LIMITED_INFORMATION',
                     win32con.THREAD_QUERY_INFORMATION)


# store last event time for displaying time between events
lastTime = 0


def log(msg):

    print(msg)


def logError(msg):
    sys.stdout.write(msg + '\n')


def getProcessID(dwEventThread, hwnd):
    # It's possible to have a window we can get a PID out of when the thread
    # isn't accessible, but it's also possible to get called with no window,
    # so we have two approaches.

    hThread = kernel32.OpenThread(threadFlag, 0, dwEventThread)

    if hThread:
        try:
            processID = kernel32.GetProcessIdOfThread(hThread)
            if not processID:
                logError("Couldn't get process for thread %s: %s" %
                         (hThread, ctypes.WinError()))
        finally:
            kernel32.CloseHandle(hThread)
    else:
        errors = ["No thread handle for %s: %s" %
                  (dwEventThread, ctypes.WinError(),)]

        if hwnd:
            processID = ctypes.wintypes.DWORD()
            threadID = user32.GetWindowThreadProcessId(
                hwnd, ctypes.byref(processID))
            if threadID != dwEventThread:
                logError("Window thread != event thread? %s != %s" %
                         (threadID, dwEventThread))
            if processID:
                processID = processID.value
            else:
                errors.append(
                    "GetWindowThreadProcessID(%s) didn't work either: %s" % (
                        hwnd, ctypes.WinError()))
                processID = None
        else:
            processID = None

        if not processID:
            for err in errors:
                logError(err)

    return processID


def getProcessFilename(processID):
    hProcess = kernel32.OpenProcess(processFlag, 0, processID)
    if not hProcess:
        logError("OpenProcess(%s) failed: %s" % (processID, ctypes.WinError()))
        return None

    try:
        filenameBufferSize = ctypes.wintypes.DWORD(4096)
        filename = ctypes.create_unicode_buffer(filenameBufferSize.value)
        kernel32.QueryFullProcessImageNameW(hProcess, 0, ctypes.byref(filename),
                                            ctypes.byref(filenameBufferSize))

        return filename.value
    finally:
        kernel32.CloseHandle(hProcess)


def callback(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread,
             dwmsEventTime):
    set_keyboard_layout()
    # lastTime = dwmsEventTime


def set_keyboard_layout():
    w = win32gui
    window_title = w.GetWindowText(w.GetForegroundWindow())
    print(window_title)
    if 'whatsapp' in window_title.lower():
        py_win_keyboard_layout.change_foreground_window_keyboard_layout(-264436723)  # to switch to Heb
        print("Keyboard changed" + str(win32api.GetKeyboardLayout()))
    if any("\u0590" <= c <= "\u05EA" for c in window_title):
        py_win_keyboard_layout.change_foreground_window_keyboard_layout(-264436723)  # to switch to Heb
        print("Keyboard changed" + str(win32api.GetKeyboardLayout()))
    elif 'visual studio code' in window_title.lower():
        py_win_keyboard_layout.change_foreground_window_keyboard_layout(67699721)  # to switch to en
        print("Keyboard changed" + str(win32api.GetKeyboardLayout()))
    elif 'chrome' in window_title.lower():
        py_win_keyboard_layout.change_foreground_window_keyboard_layout(67699721)  # to switch to en
        print("Keyboard changed" + str(win32api.GetKeyboardLayout()))
    elif 'gitkraken' in window_title.lower():
        py_win_keyboard_layout.change_foreground_window_keyboard_layout(67699721)  # to switch to en
        print("Keyboard changed" + str(win32api.GetKeyboardLayout()))
    elif 'cmd' in window_title.lower():
        py_win_keyboard_layout.change_foreground_window_keyboard_layout(67699721)  # to switch to en
        print("Keyboard changed" + str(win32api.GetKeyboardLayout()))
    elif 'ssh' in window_title.lower():
        py_win_keyboard_layout.change_foreground_window_keyboard_layout(67699721)  # to switch to en
        print("Keyboard changed" + str(win32api.GetKeyboardLayout()))


def setHook(WinEventProc, eventType):
    return user32.SetWinEventHook(
        eventType,
        eventType,
        0,
        WinEventProc,
        0,
        0,
        win32con.WINEVENT_OUTOFCONTEXT
    )


def track_window():
    ole32.CoInitialize(0)

    WinEventProc = WinEventProcType(callback)
    user32.SetWinEventHook.restype = ctypes.wintypes.HANDLE

    hookIDs = [setHook(WinEventProc, et) for et in eventTypes.keys()]
    if not any(hookIDs):
        print('SetWinEventHook failed')
        sys.exit(1)

    msg = ctypes.wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), 0, 0, 0) != 0:
        user32.TranslateMessageW(msg)
        user32.DispatchMessageW(msg)

    for hookID in hookIDs:
        user32.UnhookWinEvent(hookID)
    ole32.CoUninitialize()


# if __name__ == '__main__':

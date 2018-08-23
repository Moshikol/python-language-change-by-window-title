import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import trackwindow
import sys


class AppServerSvc (win32serviceutil.ServiceFramework):
    _svc_name_ = "Moshiko_Language_Changer"
    _svc_display_name_ = "Moshiko Language Changer"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        self.ReportServiceStatus(win32service.SERVICE_START_PENDING)
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def main(self):

        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        trackwindow.track_window()


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(AppServerSvc)

import subprocess
import schedule
import geocoder
import pytz
import threading
from suntime import Sun
from win32com.client import Dispatch
from datetime import datetime
import wx
from wx.adv import TaskBarIcon as TaskBarIcon
from time import sleep


class MyTaskBarIcon(TaskBarIcon):
    def __init__(self, frame):
        TaskBarIcon.__init__(self)

        self.frame = frame

        self.SetIcon(wx.Icon('./icon.jpg', wx.BITMAP_TYPE_JPEG), 'Windows Light Changer')

        self.Bind(wx.EVT_MENU, self.OnTaskBarActivate, id=1)
        self.Bind(wx.EVT_MENU, self.OnTaskBarDeactivate, id=2)
        self.Bind(wx.EVT_MENU, self.OnTaskBarClose, id=3)

    def CreatePopupMenu(self):
        menu = wx.Menu()
        menu.Append(1, 'Show')
        menu.Append(2, 'Hide')
        menu.Append(3, 'Close')

        return menu

    def OnTaskBarClose(self, event):
        self.frame.Close()

    def OnTaskBarActivate(self, event):
        if not self.frame.IsShown():
            self.frame.Show()

    def OnTaskBarDeactivate(self, event):
        if self.frame.IsShown():
            self.frame.Hide()
    
    def getSunriseSunset(self):
        g = geocoder.ip('me')
        lat, lng = g.latlng
        sun = Sun(lat, lng)
        return sun.get_local_sunrise_time(), sun.get_local_sunset_time(), sun.get_local_sunrise_time().strftime('%H:%M'), sun.get_local_sunset_time().strftime("%H:%S")

    def run(self, cmd):
        subprocess.run(["powershell", "-Command", cmd])

    def light_mode(self):
        light_mode = 'New-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name SystemUsesLightTheme -Value 1 -Type Dword -Force; New-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name AppsUseLightTheme -Value 1 -Type Dword -Force'
        self.run(light_mode)

    def dark_mode(self):
        dark_mode = 'New-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name SystemUsesLightTheme -Value 0 -Type Dword -Force; New-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name AppsUseLightTheme -Value 0 -Type Dword -Force'
        self.run(dark_mode)

class MyFrame(wx.Frame):
    def __init__(self, parent, id, title):
        wx.Frame.__init__(self, parent, id, title, (-1, -1), (290, 280))
        self.SetIcon(wx.Icon('./icon.jpg', wx.BITMAP_TYPE_JPEG))
        self.SetSize((350, 250))
        self.tskic = MyTaskBarIcon(self)
        self.Bind(wx.EVT_CLOSE, self.OnClose)
        self.Centre()
        
    def OnClose(self, event):
        self.tskic.Destroy()
        self.Destroy()

class MyApp(wx.App):
    def OnInit(self):
        frame = MyFrame(None, -1, 'Windows Light Changer')
        self.frame = frame
        frame.Show(False)
        self.SetTopWindow(frame)
        
        # light = schedule.every().day.at(sunrise).do(frame.tskic.light_mode)
        # dark = schedule.every().day.at(sunset).do(frame.tskic.dark_mode) 
        return True
    
    def timeCheck(self):
        frame = self.frame
        while True:
            sunrise_obj, sunset_obj, sunrise, sunset = frame.tskic.getSunriseSunset()
            utc = pytz.UTC
            curtime = utc.localize(datetime.now())
            print(sunrise, sunset)

            if sunrise_obj <= curtime < sunset_obj:
                frame.tskic.light_mode()
            else:
                frame.tskic.dark_mode()
            sleep(60)
            
    
if __name__ == '__main__':
    app = MyApp(0)
    t = threading.Thread(target=app.timeCheck)
    t.setDaemon(1)
    t.start()
    app.MainLoop()
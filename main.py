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
from time import time, sleep


class MyTaskBarIcon(TaskBarIcon):
    def __init__(self, frame, time=10):
        TaskBarIcon.__init__(self)

        self.frame = frame
        self.updateTime = time

        self.SetIcon(wx.Icon('./icon.jpg', wx.BITMAP_TYPE_JPEG), 'Windows Light Changer')

        self.Bind(wx.EVT_MENU, self.OnTaskBarActivate, id=1)
        self.Bind(wx.EVT_MENU, self.OnTaskBarDeactivate, id=2)
        self.Bind(wx.EVT_MENU, self.dark_mode, id=3)
        self.Bind(wx.EVT_MENU, self.light_mode, id=4)
        self.Bind(wx.EVT_MENU, self.OnTaskBarClose, id=5)

    def CreatePopupMenu(self):
        menu = wx.Menu()
        menu.Append(1, 'Show')
        menu.Append(2, 'Hide')
        menu.Append(3, 'Set Dark')
        menu.Append(4, 'Set Light')
        menu.Append(5, 'Close')

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
        return sun.get_local_sunrise_time(), sun.get_local_sunset_time()

    def run(self, cmd):
        subprocess.run(["powershell", "-Command", cmd])

    def light_mode(self, event=None):
        light_mode = 'New-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name SystemUsesLightTheme -Value 1 -Type Dword -Force; New-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name AppsUseLightTheme -Value 1 -Type Dword -Force'
        self.run(light_mode)

    def dark_mode(self, event=None):
        dark_mode = 'New-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name SystemUsesLightTheme -Value 0 -Type Dword -Force; New-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name AppsUseLightTheme -Value 0 -Type Dword -Force'
        self.run(dark_mode)

class MyFrame(wx.Frame):
    def __init__(self, parent, id, title):
        wx.Frame.__init__(self, parent, id, title)
        self.SetIcon(wx.Icon('./icon.jpg', wx.BITMAP_TYPE_JPEG))
        self.SetSize((350, 250))
        self.tskic = MyTaskBarIcon(self)

        ##Init Main Panel
        main_panel = wx.Panel(self)

        ## Init boxes
        vbox = wx.BoxSizer(wx.VERTICAL)
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        
        ## HBox1 Setup
        self.update_text = wx.TextCtrl(main_panel)
        self.update_text.SetHint("Seconds")
        hbox1.Add(self.update_text, 0, wx.LEFT, 8) 
        updateTimerBtn = wx.Button(main_panel, label='Update Frequency')
        updateTimerBtn.Bind(wx.EVT_BUTTON, self.update_press)
        hbox1.Add(updateTimerBtn, 0, wx.RIGHT, 3)

        ## HBox2 Setup
        darkModeBtn = wx.Button(main_panel, label='Dark Mode')
        darkModeBtn.Bind(wx.EVT_BUTTON, self.tskic.dark_mode)
        lightModeBtn = wx.Button(main_panel, label='Light Mode')
        lightModeBtn.Bind(wx.EVT_BUTTON, self.tskic.light_mode)
        hbox2.Add(darkModeBtn, 0, wx.CENTER, 5)
        hbox2.Add(lightModeBtn, 0, wx.CENTER, 5)

        vbox.Add(hbox1)
        vbox.Add(hbox2)
        main_panel.SetSizer(vbox)
        self.Bind(wx.EVT_CLOSE, self.OnClose)
        self.Centre()
        
    def OnClose(self, event):
        self.tskic.Destroy()
        self.Destroy()
    
    def update_press(self, event):
        value = self.update_text.GetValue()
        if value and value.isdigit():
            self.tskic.updateTime = int(value)
            e.set()

class MyApp(wx.App):
    def OnInit(self):
        frame = MyFrame(None, -1, 'Windows Light Changer')
        self.frame = frame
        frame.Show(True)
        self.SetTopWindow(frame)
    
        return True
    
    def timeCheck(self, event):
        frame = self.frame
        while True:
            sunrise_obj, sunset_obj = frame.tskic.getSunriseSunset()
            utc = pytz.UTC
            curtime = utc.localize(datetime.now())
            if sunrise_obj.time() <= curtime.time() < sunset_obj.time():
                frame.tskic.light_mode()
            else:
                frame.tskic.dark_mode()
            event.wait()
            event.clear()
    
    def timer(self, event):
        while True:
            start = time()
            elapsed = 0
            while elapsed < self.frame.tskic.updateTime:
                elapsed = time() - start
                sleep(0.5)
            event.set()
            
    
if __name__ == '__main__':
    e = threading.Event()
    app = MyApp(0)
    t = threading.Thread(target=app.timeCheck, args=(e,))
    timer = threading.Thread(target=app.timer, args=(e,))
    t.setDaemon(1)
    t.start()
    timer.setDaemon(1)
    timer.start()
    app.MainLoop()
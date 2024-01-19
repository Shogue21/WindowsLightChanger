import subprocess
import schedule
import time
import os, winshell
import geocoder
from suntime import Sun
from win32com.client import Dispatch

def getSunriseSunset():
    g = geocoder.ip('me')
    lat, lng = g.latlng
    sun = Sun(lat, lng)
    return sun.get_local_sunrise_time().strftime('%H:%M'), sun.get_local_sunset_time().strftime("%H:%S")

def run(cmd):
    subprocess.run(["powershell", "-Command", cmd])

def light_mode():
    light_mode = 'New-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name SystemUsesLightTheme -Value 1 -Type Dword -Force; New-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name AppsUseLightTheme -Value 1 -Type Dword -Force'
    run(light_mode)

def dark_mode():
    dark_mode = 'New-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name SystemUsesLightTheme -Value 0 -Type Dword -Force; New-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name AppsUseLightTheme -Value 0 -Type Dword -Force'
    run(dark_mode)

def run_on_startup():
    filepath, ext = os.path.splitext(__file__)
    filename = os.path.basename(filepath)
    cur_dir = os.path.dirname(filepath)
    path = '%appdata%\\Microsoft\\Windows\\Start Menu\\Programs\\Startup'
    target = icon = f"{cur_dir}\\{filename}{ext}"
    print(target)
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(f"{path}\\{filename}.lnk")
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = path
    shortcut.IconLocation = icon
    shortcut.save()
    
def create_shortcut_exe_or_other(arguments='', target='', icon='', new_name=''):
    """
    Create a shortcut for a given target filename
    :param arguments str: full arguments and filename to make link to
    :param target str: what to launch (e.g. python)
    :param icon str: .ICO file
    :return: filename of the created shortcut file
    :rtype: str
    """
    filepath, ext = os.path.splitext(target)
    filename = os.path.basename(filepath)
    working_dir = os.path.dirname(target)
    shell = Dispatch('WScript.Shell')
    if new_name:
        shortcut_filename = new_name + ".lnk"
    else:
        shortcut_filename = filename + ".lnk"
    appdata = os.getenv('APPDATA')
    shortcut_filename = os.path.join(f'{appdata}\\Microsoft\\Windows\\Start Menu\\Programs\\Startup', shortcut_filename)
    shortcut = shell.CreateShortCut(shortcut_filename)
    shortcut.Targetpath = str(target)
    shortcut.Arguments = f'"{arguments}"'
    shortcut.WorkingDirectory = working_dir
    if icon == '':
        pass
    else:
        shortcut.IconLocation = icon
    shortcut.save()
    return shortcut_filename

def main():
    sunrise, sunset = getSunriseSunset()
    print(sunrise, sunset)
    light = schedule.every().day.at(sunrise).do(light_mode)
    dark = schedule.every().day.at(sunset).do(dark_mode)
    # run_on_startup() 
    # create_shortcut_exe_or_other(target=os.path.abspath(__file__))
    while True:
        schedule.run_pending()

main()
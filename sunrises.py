#!/usr/bin/env python
#
# pip install https://github.com/pyinstaller/pyinstaller/tarball/develop
# pip install pywin32 tzlocal==2.1 requests bs4 winevt astral geoip2
# 
# pyinstaller -F -w -i ico.ico --paths C:\Windows\System32\downlevel --clean sunrises.py
#
# NOTE: Do not use it with tzlocal 3.0 or later because it will not work. It was tested with 2.1.
#
# tested on python 3.7.4 - 32 bit with pyinstaller 3.4
# tested on python 3.9.2 - 64 bit with pyinstaller 5.0dev packed with upx


#import pkg_resources.py2_warn
#from tzdata import *
#from tzdata.zoneinfo import *
#import tzdata

import pathlib


import urllib.request

import re
import geoip2.database
import socket

import astral
from astral.sun import sun


import traceback

import os
import sys
import win32api         # package pywin32
import win32con

import ctypes
from typing import List

import pythoncom
import pywintypes

import win32gui
from win32gui import *
from win32com.shell import shell, shellcon
#from win32api import *
#from win32gui import *

import win32gui_struct

import threading
import queue


from tzlocal import get_localzone

import requests

import json

from bs4 import BeautifulSoup

import pytz

import datetime, time


import configparser
from winevt import EventLog
from collections import deque
import logging





cwdir = os.getcwd()

LOCAL_TIMEZONE = str(get_localzone())
print (LOCAL_TIMEZONE)

class SysTrayIcon(object):
    '''TODO'''
    QUIT = 'QUIT'
    SPECIAL_ACTIONS = [QUIT]

    FIRST_ID = 1023

    def __init__(self,
                 icon,
                 hover_text,
                 menu_options,
                 on_quit=None,
                 default_menu_index=None,
                 window_class_name=None):

        self.icon = icon
        self.hover_text = hover_text
        self.on_quit = on_quit

        menu_options = menu_options + (('Quit', None, self.QUIT),)
        self._next_action_id = self.FIRST_ID
        self.menu_actions_by_id = set()
        self.menu_options = self._add_ids_to_menu_options(list(menu_options))
        self.menu_actions_by_id = dict(self.menu_actions_by_id)
        del self._next_action_id


        self.default_menu_index = (default_menu_index or 0)
        self.window_class_name = window_class_name or "SysTrayIconPy"

        message_map = {win32gui.RegisterWindowMessage("TaskbarCreated"): self.restart,
                       win32con.WM_DESTROY: self.destroy,
                       win32con.WM_COMMAND: self.command,
                       win32con.WM_USER+20 : self.notify,}
        # Register the Window class.
        window_class = win32gui.WNDCLASS()
        hinst = window_class.hInstance = win32gui.GetModuleHandle(None)
        window_class.lpszClassName = self.window_class_name
        window_class.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW;
        window_class.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        window_class.hbrBackground = win32con.COLOR_WINDOW
        window_class.lpfnWndProc = message_map # could also specify a wndproc.
        classAtom = win32gui.RegisterClass(window_class)
        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow(classAtom,
                                          self.window_class_name,
                                          style,
                                          0,
                                          0,
                                          win32con.CW_USEDEFAULT,
                                          win32con.CW_USEDEFAULT,
                                          0,
                                          0,
                                          hinst,
                                          None)
        win32gui.UpdateWindow(self.hwnd)
        self.notify_id = None
        self.refresh_icon()

        win32gui.PumpMessages()

    def _add_ids_to_menu_options(self, menu_options):
        result = []
        for menu_option in menu_options:
            option_text, option_icon, option_action = menu_option
            if callable(option_action) or option_action in self.SPECIAL_ACTIONS:
                self.menu_actions_by_id.add((self._next_action_id, option_action))
                result.append(menu_option + (self._next_action_id,))
            elif non_string_iterable(option_action):
                result.append((option_text,
                               option_icon,
                               self._add_ids_to_menu_options(option_action),
                               self._next_action_id))
            else:
                log.debug('Unknown item' + option_text + option_icon + option_action)
            self._next_action_id += 1
        return result

    def refresh_icon(self):
        # Try and find a custom icon
        hinst = win32gui.GetModuleHandle(None)
        if os.path.isfile(self.icon):
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            hicon = win32gui.LoadImage(hinst,
                                       self.icon,
                                       win32con.IMAGE_ICON,
                                       0,
                                       0,
                                       icon_flags)
        else:
            log.debug("Can't find icon file - using default.")
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

        if self.notify_id: message = win32gui.NIM_MODIFY
        else: message = win32gui.NIM_ADD
        self.notify_id = (self.hwnd,
                          0,
                          win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                          win32con.WM_USER+20,
                          hicon,
                          self.hover_text)
        win32gui.Shell_NotifyIcon(message, self.notify_id)

    def restart(self, hwnd, msg, wparam, lparam):
        self.refresh_icon()

    def destroy(self, hwnd, msg, wparam, lparam):
        if self.on_quit: self.on_quit(self)
        nid = (self.hwnd, 0)
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        win32gui.PostQuitMessage(0) # Terminate the app.

    def notify(self, hwnd, msg, wparam, lparam):
        if lparam==win32con.WM_LBUTTONDBLCLK:
            self.execute_menu_option(self.default_menu_index + self.FIRST_ID)
        elif lparam==win32con.WM_RBUTTONUP:
            self.show_menu()
        elif lparam==win32con.WM_LBUTTONUP:
            pass
        return True

    def show_menu(self):
        menu = win32gui.CreatePopupMenu()
        self.create_menu(menu, self.menu_options)
        #win32gui.SetMenuDefaultItem(menu, 1000, 0)

        pos = win32gui.GetCursorPos()
        # See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/menus_0hdi.asp
        win32gui.SetForegroundWindow(self.hwnd)
        win32gui.TrackPopupMenu(menu,
                                win32con.TPM_LEFTALIGN,
                                pos[0],
                                pos[1],
                                0,
                                self.hwnd,
                                None)
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)

    def create_menu(self, menu, menu_options):
        for option_text, option_icon, option_action, option_id in menu_options[::-1]:
            if option_icon:
                option_icon = self.prep_menu_icon(option_icon)

            if option_id in self.menu_actions_by_id:
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                hbmpItem=option_icon,
                                                                wID=option_id)
                win32gui.InsertMenuItem(menu, 0, 1, item)
            else:
                submenu = win32gui.CreatePopupMenu()
                self.create_menu(submenu, option_action)
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                hbmpItem=option_icon,
                                                                hSubMenu=submenu)
                win32gui.InsertMenuItem(menu, 0, 1, item)

    def prep_menu_icon(self, icon):
        # First load the icon.
        ico_x = win32api.GetSystemMetrics(win32con.SM_CXSMICON)
        ico_y = win32api.GetSystemMetrics(win32con.SM_CYSMICON)
        hicon = win32gui.LoadImage(0, icon, win32con.IMAGE_ICON, ico_x, ico_y, win32con.LR_LOADFROMFILE)

        hdcBitmap = win32gui.CreateCompatibleDC(0)
        hdcScreen = win32gui.GetDC(0)
        hbm = win32gui.CreateCompatibleBitmap(hdcScreen, ico_x, ico_y)
        hbmOld = win32gui.SelectObject(hdcBitmap, hbm)
        # Fill the background.
        brush = win32gui.GetSysColorBrush(win32con.COLOR_MENU)
        win32gui.FillRect(hdcBitmap, (0, 0, 16, 16), brush)
        # unclear if brush needs to be feed.  Best clue I can find is:
        # "GetSysColorBrush returns a cached brush instead of allocating a new
        # one." - implies no DeleteObject
        # draw the icon
        win32gui.DrawIconEx(hdcBitmap, 0, 0, hicon, ico_x, ico_y, 0, 0, win32con.DI_NORMAL)
        win32gui.SelectObject(hdcBitmap, hbmOld)
        win32gui.DeleteDC(hdcBitmap)

        return hbm

    def command(self, hwnd, msg, wparam, lparam):
        id = win32gui.LOWORD(wparam)
        self.execute_menu_option(id)

    def execute_menu_option(self, id):
        menu_action = self.menu_actions_by_id[id]
        if menu_action == self.QUIT:
            win32gui.DestroyWindow(self.hwnd)
        else:
            menu_action(self)

def non_string_iterable(obj):
    try:
        iter(obj)
    except TypeError:
        return False
    else:
        return not isinstance(obj, str)

def check_ip_address(ip):
    try:
        socket.inet_aton(ip)
        return True

    except socket.error:
        return False

def get_lat_lon_by_ip():
    ip = None
    try:
        endpoint = 'http://checkip.dyndns.org/'
        response = requests.get(endpoint)
        data = response.text
        ip = re.findall('([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)', data)[0]
        log.debug ('My public IP address is: '+ str(ip))
    except Exception as err:
        traceback.format_exc()
    if not check_ip_address(ip):
        try:
            ip = requests.get('https://api.ipify.org').text
            log.debug ('My public IP address is: '+ str(ip))
        except Exception as err:
            traceback.format_exc()
        if not check_ip_address(ip):
            try:
                ip = urllib.request.urlopen('https://ident.me').read().decode('utf8')
                log.debug ('My public IP address is: '+ str(ip))
            except Exception as err:
                traceback.format_exc()
            if not check_ip_address(ip):
                try:
                    ip = requests.get('https://checkip.amazonaws.com').text.strip()
                    log.debug ('My public IP address is:' + str(ip))
                except Exception as err:
                    traceback.format_exc()
                if not check_ip_address(ip):
                    try:
                        res = requests.get("http://whatismyip.org")
                        ip = re.compile('(\d{1,3}\.){3}\d{1,3}').search(res.text).group()
                        log.debug ('My public IP address is: ' + str(ip))
                    except Exception as err:
                        traceback.format_exc()
                    if not check_ip_address(ip):
                        try:
                            ip = requests.get('https://www.wikipedia.org').headers['X-Client-IP']
                            log.debug ('My public IP address is: ' + str(ip))
                        except Exception as err:
                            traceback.format_exc()

    log.debug ('My public IP address is: ' + str(ip))

    reader = geoip2.database.Reader(cwdir+os.sep+'GeoLite2-City.mmdb')
    response = reader.city(ip)

    lat1 = response.location.latitude
    lon1 = response.location.longitude
    log.debug (str(response.city.name) + ' '+str(lat1) + ' '+str(lon1))
    reader.close()
    """
    driver = requests.get('https://whatismyipaddress.com/ip/'+ip)
    soup = bs4.BeautifulSoup(driver.text, 'html.parser')

    ths = soup.find_all('th')
    try:
        log.debug(ths)
        for th in ths:
            if th.get_text() == 'Latitude:':
                lat = th.next_sibling.get_text().split('(')[0].strip()
                #log.debug (lat)
            if th.get_text() == 'Longitude:':
                lon = th.next_sibling.get_text().split('(')[0].strip()
                #log.debug (lon)
    except UnboundLocalError:
        return None, None
    log.debug(lat +' , '+lon)
    """
    """
    driver = requests.get('https://freegeoip.app/json/'+ip)
    json_data = json.loads(driver.text)
    #log.debug (json_data)
    lat = json_data['latitude']
    lon = json_data['longitude']
    log.debug(lat + ' , ' +lon)
    """
    return lat1, lon1


def calculate_sunrise_sunset(lat,lon):



    city = astral.LocationInfo("", "", "", float(lat), float(lon))
    """
    log.debug((
        f"Information for {city.name}/{city.region}\n"
        f"Timezone: {city.timezone}\n"
        f"Latitude: {city.latitude:.02f}; Longitude: {city.longitude:.02f}\n"
    ))
    """
    s = astral.sun.sun(city.observer, date=(datetime.date.today()), tzinfo=pytz.timezone(str(get_localzone())))
    #log.debug('Dawn:' +s["dawn"].strftime("%H:%M:%S"))
    #log.debug('Dusk:' +s["dusk"].strftime("%H:%M:%S"))

    """
    driver = requests.get('https://api.sunrise-sunset.org/json?lat='+lat+'&lng='+lon)
    # example: https://api.sunrise-sunset.org/json?lat=46.0503&lng=14.5046
    json_data = json.loads(driver.text)

    sunrise_time = datetime.datetime.strptime(str(datetime.date.today())+' '+json_data['results']['sunrise'], '%Y-%m-%d %I:%M:%S %p')
    #log.debug(str(datetime.date.today())+' '+json_data['results']['sunrise'])
    civil_twilight_begin = datetime.datetime.strptime(str(datetime.date.today())+' '+json_data['results']['civil_twilight_begin'], '%Y-%m-%d %I:%M:%S %p')
    sunrise_time = civil_twilight_begin + datetime.timedelta(0,((sunrise_time - civil_twilight_begin).seconds/2))

    sunrise_time = sunrise_time.replace(tzinfo=pytz.timezone('UTC'))
    sunrise_time = sunrise_time.astimezone(pytz.timezone(LOCAL_TIMEZONE))

    sunset_time = datetime.datetime.strptime(str(datetime.date.today())+' '+json_data['results']['sunset'], '%Y-%m-%d %I:%M:%S %p')
    civil_twilight_end = datetime.datetime.strptime(str(datetime.date.today())+' '+json_data['results']['civil_twilight_end'], '%Y-%m-%d %I:%M:%S %p')
    sunset_time = sunset_time + datetime.timedelta(0,((civil_twilight_end-sunset_time).seconds/2))
    sunset_time = sunset_time.replace(tzinfo=pytz.timezone('UTC'))
    sunset_time = sunset_time.astimezone(pytz.timezone(LOCAL_TIMEZONE))

    log.debug (sunrise_time.strftime("%H:%M:%S"))
    log.debug (sunset_time.strftime("%H:%M:%S"))

    """
    #sys.exit()
    sunrise_time = s['dawn'] + ((s["sunrise"] - s["dawn"])/2)
    sunset_time = s['sunset'] + ((s["dusk"] - s["sunset"])/2)

    #log.debug('Dawn: ' +s["dawn"].strftime("%H:%M:%S"))
    #log.debug('Sunrise: '+s["sunrise"].strftime("%H:%M:%S"))
    #log.debug('SUNRISE: '+sunrise_time.strftime("%H:%M:%S"))
    #log.debug('Dusk: ' +s["dusk"].strftime("%H:%M:%S"))
    #log.debug('Sunset: ' +s["sunset"].strftime("%H:%M:%S"))
    #log.debug('SUNSET: ' +sunset_time.strftime("%H:%M:%S"))

    return sunrise_time, sunset_time
    #return s["dawn"], s["dusk"]

class WallpaperThread(threading.Thread):
    def __init__(self, q):
        threading.Thread.__init__(self)
        self.cwd = os.getcwd()

        try:
            self.config = configparser.ConfigParser()
            self.config.read('sunrises.ini')
            self.day_wallpaper_path = self.config['DEFAULT']['day_wallpaper_path']
            self.night_wallpaper_path = self.config['DEFAULT']['night_wallpaper_path']

        #except configparser.MissingSectionHeaderError:
        #    log.debug ('except configparser.MissingSectionHeaderError:')
        except:
            #log.debug (traceback.format_exc())
            self.config['DEFAULT'] = {'day_wallpaper_path': (self.cwd+os.sep+'day.jpg').replace('\\', '/'),
                                'night_wallpaper_path': (self.cwd+os.sep+'night.jpg').replace('\\', '/')}
            with open(self.cwd+os.sep+'sunrises.ini', 'w') as configfile:
                self.config.write(configfile)

            self.day_wallpaper_path = (self.cwd+os.sep+'day.jpg').replace('\\', '/')
            self.night_wallpaper_path = (self.cwd+os.sep+'night.jpg').replace('\\', '/')

        #self.day_wallpaper_path = self.cwd+os.sep+'day.jpg'
        #self.night_wallpaper_path = self.cwd+os.sep+'night.jpg'
    def run(self):


        def sunrise_sunset():
            sunrise_sunset_calculated_day = None
            LOCAL_TIMEZONE = str(get_localzone())
            first_time = True
            #lat, lon = get_lat_lon_by_ip()
            while True:
                try:
                    now1 = datetime.datetime.now()
                    #log.debug (pytz.all_timezones)
                    #query = EventLog.Query("System","Event/System[EventID=107]")
                    query2 = EventLog.Query("System","Event/System[EventID=131]")
                    #event = next(query)
                    #dd = deque(query, maxlen=1)
                    #event = dd.pop()
                    dd = deque(query2, maxlen=1)
                    event2 = dd.pop()
                    #last_sleep_time = datetime.datetime.strptime(event.System.TimeCreated['SystemTime'][:19], '%Y-%m-%dT%H:%M:%S')
                    last_sleep_time2 = datetime.datetime.strptime(event2.System.TimeCreated['SystemTime'][:19], '%Y-%m-%dT%H:%M:%S')
                    #log.debug('last sleep time: ' +last_sleep_time)
                    #log.debug('last sleep time2: ' +str(last_sleep_time2))
                    #sleep_diff = (datetime.datetime.now(pytz.timezone('UTC')) - last_sleep_time.replace(tzinfo=pytz.timezone('UTC'))).seconds
                    #log.debug('now: '+ str(datetime.datetime.now(pytz.timezone('UTC'))))
                    sleep_diff2 = (datetime.datetime.now(pytz.timezone('UTC')) - last_sleep_time2.replace(tzinfo=pytz.timezone('UTC'))).seconds
                    #log.debug('last_sleep_time_UTC: '+ last_sleep_time.replace(tzinfo=pytz.timezone('UTC')))
                    #log.debug ('sleep diff: '+ sleep_diff)
                    #log.debug ('sleep diff2: '+ str(sleep_diff2))
                    if sleep_diff2 < 30:
                        first_time = True
                        log.debug ('Returned from sleep: first_time = True')
                    sleep_calc_time = (datetime.datetime.now() - now1).seconds
                    #log.debug ('sleep calc time: ' + str(sleep_calc_time))
                except IndexError:
                    pass
                    #log.debug ('Sleep log empty... continuing..')
                except Exception as err:
                    log.debug (traceback.format_exc())

                try:
                    if not q.empty():
                        while not q.empty():
                            a,b = q.get(timeout=0)
                            if a == 'day_wallpaper_path':
                                b = b.replace('\\', '/')
                                self.day_wallpaper_path = b
                                self.config['DEFAULT']['day_wallpaper_path'] = b
                                print(b)
                                first_time = True
                            if a == 'night_wallpaper_path':
                                b = b.replace('\\', '/')
                                self.night_wallpaper_path = b
                                self.config['DEFAULT']['night_wallpaper_path'] = b
                                print(b)
                                first_time = True
                            with open(self.cwd+os.sep+'sunrises.ini', 'w') as configfile:
                                self.config.write(configfile)
                except Exception as err:
                    log.debug (traceback.format_exc())

                try:
                    today = (datetime.datetime.now() - datetime.timedelta(hours=2, minutes=10)).date()
                    #today = datetime.datetime.now(pytz.utc).astimezone(pytz.timezone(LOCAL_TIMEZONE)).date()
                    now = datetime.datetime.now(pytz.timezone('utc')).astimezone(pytz.timezone(str(get_localzone())))

                    if sunrise_sunset_calculated_day != today:
                        lat, lon = get_lat_lon_by_ip()
                        #log.debug (lat + ' , '+ lon)
                        sunrise_time, sunset_time = calculate_sunrise_sunset(lat,lon)
                        if sunrise_time == None:
                            sunrise_time, sunset_time = calculate_sunrise_sunset(lat,lon)
                        if sunrise_time == None:
                            sunrise_time, sunset_time = calculate_sunrise_sunset(lat,lon)
                        sunrise_sunset_calculated_day = today   #datetime.datetime.now().date()

                    if ((now < sunset_time) and (now > sunrise_time)) and first_time:
                        log.debug('if ((now < sunset_time) and (now > sunrise_time)) and first_time:')

                        #ctypes.windll.user32.SystemParametersInfoW(20, 0, self.day_wallpaper_path , 3)
                        set_wallpaper(self.day_wallpaper_path)
                        first_time = False
                    elif ((now < sunrise_time) or (now > sunset_time)) and first_time:
                        log.debug('elif ((now < sunrise_time) or (now > sunset_time)) and first_time:')

                        #ctypes.windll.user32.SystemParametersInfoW(20, 0, self.night_wallpaper_path , 3)
                        set_wallpaper(self.night_wallpaper_path)
                        first_time = False

                    #log.debug ((now -sunset_time).seconds)
                    #log.debug ((sunset_time - now).seconds)
                    #log.debug ((sunrise_time - now).seconds)
                    #log.debug ((now - sunrise_time).seconds)

                    #if now < (sunrise_time + datetime.timedelta(0,5)):
                    #if ((sunrise_time + datetime.timedelta(0,5))-now).seconds < 10:
                    if (sunrise_time - now).seconds < 10:
                        log.debug ('if (sunrise_time - now).seconds < 10:')
                        log.debug (str(datetime.datetime.now()))
                        #ctypes.windll.user32.SystemParametersInfoW(20, 0, self.day_wallpaper_path , 3)
                        set_wallpaper(self.day_wallpaper_path)
                    #elif sunrise_time < now < (sunset_time + datetime.timedelta(0,5)):
                    #elif ((sunset_time + datetime.timedelta(0,5))-now).seconds < 10:
                    elif (sunset_time - now).seconds < 10:
                        log.debug ('elif (sunset_time - now).seconds < 10:')
                        log.debug (str(datetime.datetime.now()))
                        #ctypes.windll.user32.SystemParametersInfoW(20, 0, self.night_wallpaper_path , 3)
                        set_wallpaper(self.night_wallpaper_path)

                    time.sleep(5)
                except Exception as err:
                    log.debug (traceback.format_exc())

        sunrise_sunset()


class WindowsBalloonTip:
    def __init__(self, title, msg):

        message_map = {
                win32con.WM_DESTROY: self.OnDestroy,
        }
        # Register the Window class.
        wc = win32gui.WNDCLASS()
        hinst = wc.hInstance = GetModuleHandle(None)
        wc.lpszClassName = "PythonTaskbar"
        wc.lpfnWndProc = message_map # could also specify a wndproc.
        classAtom = RegisterClass(wc)
        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = CreateWindow( classAtom, "Taskbar", style, \
                0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, \
                0, 0, hinst, None)
        UpdateWindow(self.hwnd)
        iconPathName = os.path.abspath(os.path.join( sys.path[0], cwdir+os.sep+"balloon.ico" ))
        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        try:
           hicon = LoadImage(hinst, iconPathName, \
                    win32con.IMAGE_ICON, 0, 0, icon_flags)
        except:
          hicon = LoadIcon(0, win32con.IDI_APPLICATION)
        flags = NIF_ICON | NIF_MESSAGE | NIF_TIP
        nid = (self.hwnd, 0, flags, win32con.WM_USER+20, hicon, "message")
        Shell_NotifyIcon(NIM_ADD, nid)
        Shell_NotifyIcon(NIM_MODIFY, \
                         (self.hwnd, 0, NIF_INFO, win32con.WM_USER+20,\
                          hicon, "Balloon  tooltip",msg,200,title))
        # self.show_balloon(title, msg)
        time.sleep(10)
        DestroyWindow(self.hwnd)
        UnregisterClass(classAtom, hinst)
    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        Shell_NotifyIcon(NIM_DELETE, nid)
        #PostQuitMessage(0) # Terminate the app.

def balloon_tip(title, msg):
    w=WindowsBalloonTip(title, msg)

def _make_filter(class_name: str, title: str):
    """https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enumwindows"""

    def enum_windows(handle: int, h_list: list):
        if not (class_name or title):
            h_list.append(handle)
        if class_name and class_name not in win32gui.GetClassName(handle):
            return True  # continue enumeration
        if title and title not in win32gui.GetWindowText(handle):
            return True  # continue enumeration
        h_list.append(handle)

    return enum_windows


def find_window_handles(parent: int = None, window_class: str = None, title: str = None) -> List[int]:
    cb = _make_filter(window_class, title)
    try:
        handle_list = []
        if parent:
            win32gui.EnumChildWindows(parent, cb, handle_list)
        else:
            win32gui.EnumWindows(cb, handle_list)
        return handle_list
    except pywintypes.error:
        return []


def force_refresh():
    ctypes.windll.user32.UpdatePerUserSystemParameters(1)


def enable_activedesktop():
    """https://stackoverflow.com/a/16351170"""
    try:
        progman = find_window_handles(window_class='Progman')[0]
        cryptic_params = (0x52c, 0, 0, 0, 500, None)
        ctypes.windll.user32.SendMessageTimeoutW(progman, *cryptic_params)
    except IndexError as e:
        raise WindowsError('Cannot enable Active Desktop') from e


def set_wallpaper(image_path: str, use_activedesktop: bool = True):
    if use_activedesktop:
        enable_activedesktop()
    pythoncom.CoInitialize()
    iad = pythoncom.CoCreateInstance(shell.CLSID_ActiveDesktop,
                                     None,
                                     pythoncom.CLSCTX_INPROC_SERVER,
                                     shell.IID_IActiveDesktop)
    iad.SetWallpaper(image_path, 0)
    iad.ApplyChanges(shellcon.AD_APPLY_ALL)
    force_refresh()
#if __name__ == '__main__':
#    balloon_tip("Title for popup", "This is the popup's message")



# Minimal self test. You'll need a bunch of ICO files in the current working
# directory in order for this to work...
if __name__ == '__main__':

    icon = 'ico.ico'
    hover_text = "Sunrise/sunset wallpaper changer"

    #get_lat_lon_by_ip()

    global q
    q = queue.Queue()


    log = logging.getLogger(__name__)
    log.setLevel(logging.DEBUG)
    ch = logging.FileHandler(filename='sunrises.log')
    ch.setLevel(logging.DEBUG)
    # create formatter
    formatter = logging.Formatter('%(asctime)s - Sunrise-sunset - %(message)s')
    # add formatter to ch
    ch.setFormatter(formatter)
    # add ch to logger
    log.addHandler(ch)
    log.debug('Sunrises program started...')



    t = WallpaperThread(q)
    t.daemon = True
    t.start()


    def day(sysTrayIcon):
        try:
            """Ask the user to select a single file.  Return full path"""
            f = 'Picture files\0*.jpg;*.jpeg;*.png;*.bmp\0'
            try:
                ret = win32gui.GetOpenFileNameW(None,
                                                Flags=win32con.OFN_EXPLORER
                                                | win32con.OFN_FILEMUSTEXIST,
                                                Title='', Filter=f)
                #log.debug (ret[0])
                log.debug(ret[0])
            except pywintypes.error:
                log.debug('No file selected..')
                #log.debug('No file selected..')
            try:
                path_is_file = os.path.isfile(ret[0])
            except:
                path_is_file = False
            if path_is_file:
                q.put(('day_wallpaper_path', ret[0]))
                #log.debug("Day wallpaper set.")
                log.debug("Day wallpaper set.")
            else:
                #log.debug("Day wallpaper not set.")
                log.debug("Day wallpaper not set.")
        except Exception as err:
            #traceback.format_exc()
            log.debug(traceback.format_exc())

    def night(sysTrayIcon):
        try:
            """Ask the user to select a single file.  Return full path"""
            f = 'Picture files\0*.jpg;*.jpeg;*.png;*.bmp\0'
            try:
                ret = win32gui.GetOpenFileNameW(None,
                                                Flags=win32con.OFN_EXPLORER
                                                | win32con.OFN_FILEMUSTEXIST,
                                                Title='', Filter=f)
                log.debug(ret[0])
            except pywintypes.error:
                log.debug('No file selected..')
            try:
                path_is_file = os.path.isfile(ret[0])
            except:
                path_is_file = False
            if path_is_file:
                q.put(('night_wallpaper_path', ret[0]))
                log.debug("Night wallpaper set.")
            else:
                log.debug("Night wallpaper not set.")
        except Exception as err:
            log.debug(traceback.format_exc())

    def balloon_tip(sysTrayIcon):
        try:
            lat, lon = get_lat_lon_by_ip()

            sunrise_time, sunset_time = calculate_sunrise_sunset(lat,lon)
            if sunrise_time == None:
                sunrise_time, sunset_time = calculate_sunrise_sunset(lat,lon)
            if sunrise_time == None:
                sunrise_time, sunset_time = calculate_sunrise_sunset(lat,lon)

            w=WindowsBalloonTip('Current sunrise/sunset:', 'Sunrise: '+sunrise_time.strftime("%H:%M:%S")+'\nSunset:  '+sunset_time.strftime("%H:%M:%S")+'\n')
        except Exception as err:
            log.debug (traceback.format_exc())

    menu_options = (('Show current sunrise/sunset', None, balloon_tip),
                    ('Change day wallpaper', None, day),
                    ('Change night wallpaper', None, night),)

    def bye(sysTrayIcon): log.debug('Day-Night quitting!')

    try:
        SysTrayIcon(icon, hover_text, menu_options, on_quit=bye, default_menu_index=1)
    except KeyboardInterrupt:
        log.debug('Keyboard interrupt. Quitting!')
        sys.exit()
    except Exception as err:
        log.debug(traceback.format_exc())


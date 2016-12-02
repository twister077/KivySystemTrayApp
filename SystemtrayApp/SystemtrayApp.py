import kivy
import threading

kivy.require('1.9.1')

from kivy.app import App
from kivy.uix.label import Label
from kivy.lang import Builder
from kivy.config import Config
from kivy.core.window import Window

from win32api import *
from win32gui import *
import win32con
import sys, os
import struct
import time

import win32gui_struct
try:
    import winxpgui as win32gui
except ImportError:
    import win32gui

Window.clearcolor = (1, 1, 1, 1)

kv = '''
BoxLayout:
    size_hint_y: None
    height: '50sp'
    Button:
        text: 'Start service'
        on_press: app.runTask()
    Button:
        text: 'Quit'
        on_press: app.runQuit()

'''

class SystemTrayIcon:
    def __init__(self):
        self.QUIT = 'QUIT'
        message_map = {
                win32con.WM_DESTROY: self.OnDestroy,
                win32con.WM_COMMAND: self.command,
                win32con.WM_USER+20 : self.notify,
        }
        # Register the Window class.
        wc = WNDCLASS()
        hinst = wc.hInstance = GetModuleHandle(None)
        wc.lpszClassName = "PythonTaskbar"
        wc.lpfnWndProc = message_map
        classAtom = RegisterClass(wc)
        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = CreateWindow( classAtom, "Taskbar", style, \
                0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, \
                0, 0, hinst, None)
        UpdateWindow(self.hwnd)
        iconPathName = os.path.abspath(os.path.join( sys.path[0], 'systray.ico'))
        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        try:
           hicon = LoadImage(hinst, iconPathName, \
                    win32con.IMAGE_ICON, 0, 0, icon_flags)
        except:
          hicon = LoadIcon(0, win32con.IDI_APPLICATION)
        flags = NIF_ICON | NIF_MESSAGE | NIF_TIP
        nid = (self.hwnd, 0, flags, win32con.WM_USER+20, hicon, "SysTrayExample")
        Shell_NotifyIcon(NIM_ADD, nid)

    def OnNotify(self, title, msg):
        try:
           hicon = LoadImage(hinst, iconPathName, \
                    win32con.IMAGE_ICON, 0, 0, icon_flags)
        except:
          hicon = LoadIcon(0, win32con.IDI_APPLICATION)
        Shell_NotifyIcon(NIM_MODIFY, \
                         (self.hwnd, 0, NIF_INFO, win32con.WM_USER+20,\
                          hicon, "Balloon  tooltip",msg,200,title))

    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        Shell_NotifyIcon(NIM_DELETE, nid)
        PostQuitMessage(0)

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
                print 'Unknown item', option_text, option_icon, option_action
            self._next_action_id += 1
        return result

    def notify(self, hwnd, msg, wparam, lparam):
        if lparam == win32con.WM_LBUTTONDBLCLK:
            self.execute_menu_option(self._default_menu_index + SysTrayIcon.FIRST_ID)
        elif lparam == win32con.WM_RBUTTONUP:
            self._show_menu()
        elif lparam == win32con.WM_LBUTTONUP:
            pass
        return True

    def command(self, hwnd, msg, wparam, lparam):
        id = win32gui.LOWORD(wparam)
        self.execute_menu_option(id)

    def execute_menu_option(self, id):
        menu_action = self.menu_actions_by_id[id]      
        if menu_action == self.QUIT:
            win32gui.DestroyWindow(self.hwnd)
        else:
            menu_action(self)

    def _show_menu(self):
        self.FIRST_ID = 1023
        self.menu_options = (('Show', None, self.onShow),('Run', None, self.onRun),('Quit', None, self.onQuit))
        self._next_action_id = self.FIRST_ID
        self.menu_actions_by_id = set()
        self.menu_options = self._add_ids_to_menu_options(list(self.menu_options))
        self.menu_actions_by_id = dict(self.menu_actions_by_id)
        del self._next_action_id
        menu = win32gui.CreatePopupMenu()
        self.create_menu(menu, self.menu_options)
        
        pos = win32gui.GetCursorPos()
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

    def onShow(self, id):
        print 'Showing .......'

    def onRun(self, id):
        print 'Run some code .......'

    def onQuit(self, id):
        nid = (self.hwnd, 0)
        Shell_NotifyIcon(NIM_DELETE, nid)
        PostQuitMessage(0)
        exit()


class SysTrayApp(App):

    def __init__(self, **kwargs):
        self.w=SystemTrayIcon()
        return super(SysTrayApp, self).__init__(**kwargs)

    def build(self):
        self.root = Builder.load_string(kv)
        self.icon = 'systray.png'
        return self.root

    def runTask(self):
        self.root_window.minimize
        self.w.OnNotify('SysTray','Task ran successfully')

    def runQuit(self):
        self.w.onQuit(1)

if __name__ == '__main__':
    SysTrayApp().run()

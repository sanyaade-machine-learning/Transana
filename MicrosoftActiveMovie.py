# -*- coding: mbcs -*-
# Created by makepy.py version 0.4.8
# By python version 2.3.3 (#51, Dec 18 2003, 20:22:39) [MSC v.1200 32 bit (Intel)]
# From type library 'amcompat.tlb'
# On Tue Jan 27 16:44:53 2004
"""Microsoft ActiveMovie Control"""
makepy_version = '0.4.8'
python_version = 0x20303f0

import win32com.client.CLSIDToClass, pythoncom
import win32com.client.util
from pywintypes import IID
from win32com.client import Dispatch

# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing and pythoncom.Empty
defaultNamedOptArg=pythoncom.Empty
defaultNamedNotOptArg=pythoncom.Empty
defaultUnnamedArg=pythoncom.Empty

CLSID = IID('{05589FA0-C356-11CE-BF01-00AA0055595A}')
MajorVersion = 2
MinorVersion = 0
LibraryFlags = 8
LCID = 0x0

class constants:
	amv3D                         =0x1        # from enum AppearanceConstants
	amvFlat                       =0x0        # from enum AppearanceConstants
	amvFixedSingle                =0x1        # from enum BorderStyleConstants
	amvNone                       =0x0        # from enum BorderStyleConstants
	amvFrames                     =0x1        # from enum DisplayModeConstants
	amvTime                       =0x0        # from enum DisplayModeConstants
	amvComplete                   =0x4        # from enum ReadyStateConstants
	amvInteractive                =0x3        # from enum ReadyStateConstants
	amvLoading                    =0x1        # from enum ReadyStateConstants
	amvUninitialized              =0x0        # from enum ReadyStateConstants
	amvNotLoaded                  =-1         # from enum StateConstants
	amvPaused                     =0x1        # from enum StateConstants
	amvRunning                    =0x2        # from enum StateConstants
	amvStopped                    =0x0        # from enum StateConstants
	amvDoubleOriginalSize         =0x1        # from enum WindowSizeConstants
	amvOneFourthScreen            =0x3        # from enum WindowSizeConstants
	amvOneHalfScreen              =0x4        # from enum WindowSizeConstants
	amvOneSixteenthScreen         =0x2        # from enum WindowSizeConstants
	amvOriginalSize               =0x0        # from enum WindowSizeConstants

from win32com.client import DispatchBaseClass
class DActiveMovieEvents(DispatchBaseClass):
	"""Event interface for ActiveMovie Control"""
	CLSID = IID('{05589FA3-C356-11CE-BF01-00AA0055595A}')
	coclass_clsid = IID('{05589FA1-C356-11CE-BF01-00AA0055595A}')

	def Click(self):
		return self._oleobj_.InvokeTypes(-600, LCID, 1, (24, 0), (),)

	def DblClick(self):
		return self._oleobj_.InvokeTypes(-601, LCID, 1, (24, 0), (),)

	def Error(self, SCode=defaultNamedNotOptArg, Description=defaultNamedNotOptArg, Source=defaultNamedNotOptArg, CancelDisplay=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(999, LCID, 1, (24, 0), ((2, 0), (8, 0), (8, 0), (16395, 0)),SCode, Description, Source, CancelDisplay)

	def KeyDown(self, KeyCode=defaultNamedNotOptArg, Shift=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(-602, LCID, 1, (24, 0), ((16386, 0), (2, 0)),KeyCode, Shift)

	def KeyPress(self, KeyAscii=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(-603, LCID, 1, (24, 0), ((16386, 0),),KeyAscii)

	def KeyUp(self, KeyCode=defaultNamedNotOptArg, Shift=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(-604, LCID, 1, (24, 0), ((16386, 0), (2, 0)),KeyCode, Shift)

	def MouseDown(self, Button=defaultNamedNotOptArg, Shift=defaultNamedNotOptArg, x=defaultNamedNotOptArg, y=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(-605, LCID, 1, (24, 0), ((2, 0), (2, 0), (3, 0), (3, 0)),Button, Shift, x, y)

	def MouseMove(self, Button=defaultNamedNotOptArg, Shift=defaultNamedNotOptArg, x=defaultNamedNotOptArg, y=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(-606, LCID, 1, (24, 0), ((2, 0), (2, 0), (3, 0), (3, 0)),Button, Shift, x, y)

	def MouseUp(self, Button=defaultNamedNotOptArg, Shift=defaultNamedNotOptArg, x=defaultNamedNotOptArg, y=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(-607, LCID, 1, (24, 0), ((2, 0), (2, 0), (3, 0), (3, 0)),Button, Shift, x, y)

	def OpenComplete(self):
		"""Reports that an asynchronous operation to open a file has completed successfully"""
		return self._oleobj_.InvokeTypes(50, LCID, 1, (24, 0), (),)

	def PositionChange(self, oldPosition=defaultNamedNotOptArg, newPosition=defaultNamedNotOptArg):
		"""Indicates that the current position of the movie has changed"""
		return self._oleobj_.InvokeTypes(2, LCID, 1, (24, 0), ((5, 0), (5, 0)),oldPosition, newPosition)

	def StateChange(self, oldState=defaultNamedNotOptArg, newState=defaultNamedNotOptArg):
		"""Indicates that the current state of the movie has changed"""
		return self._oleobj_.InvokeTypes(1, LCID, 1, (24, 0), ((3, 0), (3, 0)),oldState, newState)

	def Timer(self):
		"""ActiveMovie Control's progress timer"""
		return self._oleobj_.InvokeTypes(3, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
	}
	_prop_map_put_ = {
	}

class DActiveMovieEvents2:
	"""Event interface for ActiveMovie Control"""
	CLSID = CLSID_Sink = IID('{B6CD6553-E9CB-11D0-821F-00A0C91F9CA0}')
	coclass_clsid = IID('{05589FA1-C356-11CE-BF01-00AA0055595A}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		        2 : "OnPositionChange",
		        1 : "OnStateChange",
		     -603 : "OnKeyPress",
		     -604 : "OnKeyUp",
		     -607 : "OnMouseUp",
		        3 : "OnTimer",
		     -606 : "OnMouseMove",
		     -609 : "OnReadyStateChange",
		       51 : "OnDisplayModeChange",
		     -605 : "OnMouseDown",
		     -602 : "OnKeyDown",
		       52 : "OnScriptCommand",
		      999 : "OnError",
		       50 : "OnOpenComplete",
		     -600 : "OnClick",
		     -601 : "OnDblClick",
		}

	def __init__(self, oobj = None):
		if oobj is None:
			self._olecp = None
		else:
			import win32com.server.util
			from win32com.server.policy import EventHandlerPolicy
			cpc=oobj._oleobj_.QueryInterface(pythoncom.IID_IConnectionPointContainer)
			cp=cpc.FindConnectionPoint(self.CLSID_Sink)
			cookie=cp.Advise(win32com.server.util.wrap(self, usePolicy=EventHandlerPolicy))
			self._olecp,self._olecp_cookie = cp,cookie
	def __del__(self):
		try:
			self.close()
		except pythoncom.com_error:
			pass
	def close(self):
		if self._olecp is not None:
			cp,cookie,self._olecp,self._olecp_cookie = self._olecp,self._olecp_cookie,None,None
			cp.Unadvise(cookie)
	def _query_interface_(self, iid):
		import win32com.server.util
		if iid==self.CLSID_Sink: return win32com.server.util.wrap(self)

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnPositionChange(self, oldPosition=defaultNamedNotOptArg, newPosition=defaultNamedNotOptArg):
#		"""Indicates that the current position of the movie has changed"""
#	def OnStateChange(self, oldState=defaultNamedNotOptArg, newState=defaultNamedNotOptArg):
#		"""Indicates that the current state of the movie has changed"""
#	def OnKeyPress(self, KeyAscii=defaultNamedNotOptArg):
#	def OnKeyUp(self, KeyCode=defaultNamedNotOptArg, Shift=defaultNamedNotOptArg):
#	def OnMouseUp(self, Button=defaultNamedNotOptArg, Shift=defaultNamedNotOptArg, x=defaultNamedNotOptArg, y=defaultNamedNotOptArg):
#	def OnTimer(self):
#		"""ActiveMovie Control's progress timer"""
#	def OnMouseMove(self, Button=defaultNamedNotOptArg, Shift=defaultNamedNotOptArg, x=defaultNamedNotOptArg, y=defaultNamedNotOptArg):
#	def OnReadyStateChange(self, ReadyState=defaultNamedNotOptArg):
#		"""Reports that the ReadyState property of the ActiveMovie Control has changed"""
#	def OnDisplayModeChange(self):
#		"""Indicates that the display mode of the movie has changed"""
#	def OnMouseDown(self, Button=defaultNamedNotOptArg, Shift=defaultNamedNotOptArg, x=defaultNamedNotOptArg, y=defaultNamedNotOptArg):
#	def OnKeyDown(self, KeyCode=defaultNamedNotOptArg, Shift=defaultNamedNotOptArg):
#	def OnScriptCommand(self, bstrType=defaultNamedNotOptArg, bstrText=defaultNamedNotOptArg):
#	def OnError(self, SCode=defaultNamedNotOptArg, Description=defaultNamedNotOptArg, Source=defaultNamedNotOptArg, CancelDisplay=defaultNamedNotOptArg):
#	def OnOpenComplete(self):
#		"""Reports that an asynchronous operation to open a file has completed successfully"""
#	def OnClick(self):
#	def OnDblClick(self):


class IActiveMovie(DispatchBaseClass):
	"""ActiveMovie Control"""
	CLSID = IID('{05589FA2-C356-11CE-BF01-00AA0055595A}')
	coclass_clsid = IID('{05589FA1-C356-11CE-BF01-00AA0055595A}')

	def AboutBox(self):
		return self._oleobj_.InvokeTypes(-552, LCID, 1, (24, 0), (),)

	def Pause(self):
		"""Puts the multimedia stream into Paused state"""
		return self._oleobj_.InvokeTypes(1610743810, LCID, 1, (24, 0), (),)

	def Run(self):
		"""Puts the multimedia stream into Running state"""
		return self._oleobj_.InvokeTypes(1610743809, LCID, 1, (24, 0), (),)

	def Stop(self):
		"""Puts the multimedia stream into Stopped state"""
		return self._oleobj_.InvokeTypes(1610743811, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"AllowChangeDisplayMode": (33, 2, (11, 0), (), "AllowChangeDisplayMode", None),
		"AllowHideControls": (31, 2, (11, 0), (), "AllowHideControls", None),
		"AllowHideDisplay": (30, 2, (11, 0), (), "AllowHideDisplay", None),
		"Appearance": (-520, 2, (3, 0), (), "Appearance", None),
		"Author": (6, 2, (8, 0), (), "Author", None),
		"AutoRewind": (41, 2, (11, 0), (), "AutoRewind", None),
		"AutoStart": (40, 2, (11, 0), (), "AutoStart", None),
		"Balance": (20, 2, (3, 0), (), "Balance", None),
		"BorderStyle": (42, 2, (3, 0), (), "BorderStyle", None),
		"Copyright": (8, 2, (8, 0), (), "Copyright", None),
		"CurrentPosition": (13, 2, (5, 0), (), "CurrentPosition", None),
		"CurrentState": (17, 2, (3, 0), (), "CurrentState", None),
		"Description": (9, 2, (8, 0), (), "Description", None),
		"DisplayBackColor": (37, 2, (19, 0), (), "DisplayBackColor", None),
		"DisplayForeColor": (36, 2, (19, 0), (), "DisplayForeColor", None),
		"DisplayMode": (32, 2, (3, 0), (), "DisplayMode", None),
		"Duration": (12, 2, (5, 0), (), "Duration", None),
		"EnableContextMenu": (21, 2, (11, 0), (), "EnableContextMenu", None),
		"EnablePositionControls": (27, 2, (11, 0), (), "EnablePositionControls", None),
		"EnableSelectionControls": (28, 2, (11, 0), (), "EnableSelectionControls", None),
		"EnableTracker": (29, 2, (11, 0), (), "EnableTracker", None),
		"Enabled": (-514, 2, (11, 0), (), "Enabled", None),
		"FileName": (11, 2, (8, 0), (), "FileName", None),
		"FilterGraph": (34, 2, (13, 0), (), "FilterGraph", None),
		"FilterGraphDispatch": (35, 2, (9, 0), (), "FilterGraphDispatch", None),
		"FullScreenMode": (39, 2, (11, 0), (), "FullScreenMode", None),
		"ImageSourceHeight": (5, 2, (3, 0), (), "ImageSourceHeight", None),
		"ImageSourceWidth": (4, 2, (3, 0), (), "ImageSourceWidth", None),
		"Info": (1610743885, 2, (3, 0), (), "Info", None),
		"MovieWindowSize": (38, 2, (3, 0), (), "MovieWindowSize", None),
		"PlayCount": (14, 2, (3, 0), (), "PlayCount", None),
		"Rate": (18, 2, (5, 0), (), "Rate", None),
		"Rating": (10, 2, (8, 0), (), "Rating", None),
		"SelectionEnd": (16, 2, (5, 0), (), "SelectionEnd", None),
		"SelectionStart": (15, 2, (5, 0), (), "SelectionStart", None),
		"ShowControls": (23, 2, (11, 0), (), "ShowControls", None),
		"ShowDisplay": (22, 2, (11, 0), (), "ShowDisplay", None),
		"ShowPositionControls": (24, 2, (11, 0), (), "ShowPositionControls", None),
		"ShowSelectionControls": (25, 2, (11, 0), (), "ShowSelectionControls", None),
		"ShowTracker": (26, 2, (11, 0), (), "ShowTracker", None),
		"Title": (7, 2, (8, 0), (), "Title", None),
		"Volume": (19, 2, (3, 0), (), "Volume", None),
		"hWnd": (-515, 2, (3, 0), (), "hWnd", None),
	}
	_prop_map_put_ = {
		"AllowChangeDisplayMode": ((33, LCID, 4, 0),()),
		"AllowHideControls": ((31, LCID, 4, 0),()),
		"AllowHideDisplay": ((30, LCID, 4, 0),()),
		"Appearance": ((-520, LCID, 4, 0),()),
		"AutoRewind": ((41, LCID, 4, 0),()),
		"AutoStart": ((40, LCID, 4, 0),()),
		"Balance": ((20, LCID, 4, 0),()),
		"BorderStyle": ((42, LCID, 4, 0),()),
		"CurrentPosition": ((13, LCID, 4, 0),()),
		"DisplayBackColor": ((37, LCID, 4, 0),()),
		"DisplayForeColor": ((36, LCID, 4, 0),()),
		"DisplayMode": ((32, LCID, 4, 0),()),
		"EnableContextMenu": ((21, LCID, 4, 0),()),
		"EnablePositionControls": ((27, LCID, 4, 0),()),
		"EnableSelectionControls": ((28, LCID, 4, 0),()),
		"EnableTracker": ((29, LCID, 4, 0),()),
		"Enabled": ((-514, LCID, 4, 0),()),
		"FileName": ((11, LCID, 4, 0),()),
		"FilterGraph": ((34, LCID, 4, 0),()),
		"FullScreenMode": ((39, LCID, 4, 0),()),
		"MovieWindowSize": ((38, LCID, 4, 0),()),
		"PlayCount": ((14, LCID, 4, 0),()),
		"Rate": ((18, LCID, 4, 0),()),
		"SelectionEnd": ((16, LCID, 4, 0),()),
		"SelectionStart": ((15, LCID, 4, 0),()),
		"ShowControls": ((23, LCID, 4, 0),()),
		"ShowDisplay": ((22, LCID, 4, 0),()),
		"ShowPositionControls": ((24, LCID, 4, 0),()),
		"ShowSelectionControls": ((25, LCID, 4, 0),()),
		"ShowTracker": ((26, LCID, 4, 0),()),
		"Volume": ((19, LCID, 4, 0),()),
	}

class IActiveMovie2(DispatchBaseClass):
	"""ActiveMovie Control"""
	CLSID = IID('{B6CD6554-E9CB-11D0-821F-00A0C91F9CA0}')
	coclass_clsid = IID('{05589FA1-C356-11CE-BF01-00AA0055595A}')

	def AboutBox(self):
		return self._oleobj_.InvokeTypes(-552, LCID, 1, (24, 0), (),)

	def IsSoundCardEnabled(self):
		"""Determines whether the sound card is enabled on the machine"""
		return self._oleobj_.InvokeTypes(53, LCID, 1, (11, 0), (),)

	def Pause(self):
		"""Puts the multimedia stream into Paused state"""
		return self._oleobj_.InvokeTypes(1610743810, LCID, 1, (24, 0), (),)

	def Run(self):
		"""Puts the multimedia stream into Running state"""
		return self._oleobj_.InvokeTypes(1610743809, LCID, 1, (24, 0), (),)

	def Stop(self):
		"""Puts the multimedia stream into Stopped state"""
		return self._oleobj_.InvokeTypes(1610743811, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"AllowChangeDisplayMode": (33, 2, (11, 0), (), "AllowChangeDisplayMode", None),
		"AllowHideControls": (31, 2, (11, 0), (), "AllowHideControls", None),
		"AllowHideDisplay": (30, 2, (11, 0), (), "AllowHideDisplay", None),
		"Appearance": (-520, 2, (3, 0), (), "Appearance", None),
		"Author": (6, 2, (8, 0), (), "Author", None),
		"AutoRewind": (41, 2, (11, 0), (), "AutoRewind", None),
		"AutoStart": (40, 2, (11, 0), (), "AutoStart", None),
		"Balance": (20, 2, (3, 0), (), "Balance", None),
		"BorderStyle": (42, 2, (3, 0), (), "BorderStyle", None),
		"Copyright": (8, 2, (8, 0), (), "Copyright", None),
		"CurrentPosition": (13, 2, (5, 0), (), "CurrentPosition", None),
		"CurrentState": (17, 2, (3, 0), (), "CurrentState", None),
		"Description": (9, 2, (8, 0), (), "Description", None),
		"DisplayBackColor": (37, 2, (19, 0), (), "DisplayBackColor", None),
		"DisplayForeColor": (36, 2, (19, 0), (), "DisplayForeColor", None),
		"DisplayMode": (32, 2, (3, 0), (), "DisplayMode", None),
		"Duration": (12, 2, (5, 0), (), "Duration", None),
		"EnableContextMenu": (21, 2, (11, 0), (), "EnableContextMenu", None),
		"EnablePositionControls": (27, 2, (11, 0), (), "EnablePositionControls", None),
		"EnableSelectionControls": (28, 2, (11, 0), (), "EnableSelectionControls", None),
		"EnableTracker": (29, 2, (11, 0), (), "EnableTracker", None),
		"Enabled": (-514, 2, (11, 0), (), "Enabled", None),
		"FileName": (11, 2, (8, 0), (), "FileName", None),
		"FilterGraph": (34, 2, (13, 0), (), "FilterGraph", None),
		"FilterGraphDispatch": (35, 2, (9, 0), (), "FilterGraphDispatch", None),
		"FullScreenMode": (39, 2, (11, 0), (), "FullScreenMode", None),
		"ImageSourceHeight": (5, 2, (3, 0), (), "ImageSourceHeight", None),
		"ImageSourceWidth": (4, 2, (3, 0), (), "ImageSourceWidth", None),
		"Info": (1610743885, 2, (3, 0), (), "Info", None),
		"MovieWindowSize": (38, 2, (3, 0), (), "MovieWindowSize", None),
		"PlayCount": (14, 2, (3, 0), (), "PlayCount", None),
		"Rate": (18, 2, (5, 0), (), "Rate", None),
		"Rating": (10, 2, (8, 0), (), "Rating", None),
		"ReadyState": (-525, 2, (3, 0), (), "ReadyState", None),
		"SelectionEnd": (16, 2, (5, 0), (), "SelectionEnd", None),
		"SelectionStart": (15, 2, (5, 0), (), "SelectionStart", None),
		"ShowControls": (23, 2, (11, 0), (), "ShowControls", None),
		"ShowDisplay": (22, 2, (11, 0), (), "ShowDisplay", None),
		"ShowPositionControls": (24, 2, (11, 0), (), "ShowPositionControls", None),
		"ShowSelectionControls": (25, 2, (11, 0), (), "ShowSelectionControls", None),
		"ShowTracker": (26, 2, (11, 0), (), "ShowTracker", None),
		"Title": (7, 2, (8, 0), (), "Title", None),
		"Volume": (19, 2, (3, 0), (), "Volume", None),
		"hWnd": (-515, 2, (3, 0), (), "hWnd", None),
	}
	_prop_map_put_ = {
		"AllowChangeDisplayMode": ((33, LCID, 4, 0),()),
		"AllowHideControls": ((31, LCID, 4, 0),()),
		"AllowHideDisplay": ((30, LCID, 4, 0),()),
		"Appearance": ((-520, LCID, 4, 0),()),
		"AutoRewind": ((41, LCID, 4, 0),()),
		"AutoStart": ((40, LCID, 4, 0),()),
		"Balance": ((20, LCID, 4, 0),()),
		"BorderStyle": ((42, LCID, 4, 0),()),
		"CurrentPosition": ((13, LCID, 4, 0),()),
		"DisplayBackColor": ((37, LCID, 4, 0),()),
		"DisplayForeColor": ((36, LCID, 4, 0),()),
		"DisplayMode": ((32, LCID, 4, 0),()),
		"EnableContextMenu": ((21, LCID, 4, 0),()),
		"EnablePositionControls": ((27, LCID, 4, 0),()),
		"EnableSelectionControls": ((28, LCID, 4, 0),()),
		"EnableTracker": ((29, LCID, 4, 0),()),
		"Enabled": ((-514, LCID, 4, 0),()),
		"FileName": ((11, LCID, 4, 0),()),
		"FilterGraph": ((34, LCID, 4, 0),()),
		"FullScreenMode": ((39, LCID, 4, 0),()),
		"MovieWindowSize": ((38, LCID, 4, 0),()),
		"PlayCount": ((14, LCID, 4, 0),()),
		"Rate": ((18, LCID, 4, 0),()),
		"SelectionEnd": ((16, LCID, 4, 0),()),
		"SelectionStart": ((15, LCID, 4, 0),()),
		"ShowControls": ((23, LCID, 4, 0),()),
		"ShowDisplay": ((22, LCID, 4, 0),()),
		"ShowPositionControls": ((24, LCID, 4, 0),()),
		"ShowSelectionControls": ((25, LCID, 4, 0),()),
		"ShowTracker": ((26, LCID, 4, 0),()),
		"Volume": ((19, LCID, 4, 0),()),
	}

class IActiveMovie3(DispatchBaseClass):
	"""ActiveMovie Control"""
	CLSID = IID('{265EC140-AE62-11D1-8500-00A0C91F9CA0}')
	coclass_clsid = IID('{05589FA1-C356-11CE-BF01-00AA0055595A}')

	def AboutBox(self):
		return self._oleobj_.InvokeTypes(-552, LCID, 1, (24, 0), (),)

	def IsSoundCardEnabled(self):
		"""Determines whether the sound card is enabled on the machine"""
		return self._oleobj_.InvokeTypes(53, LCID, 1, (11, 0), (),)

	def Pause(self):
		"""Puts the multimedia stream into Paused state"""
		return self._oleobj_.InvokeTypes(1610743810, LCID, 1, (24, 0), (),)

	def Run(self):
		"""Puts the multimedia stream into Running state"""
		return self._oleobj_.InvokeTypes(1610743809, LCID, 1, (24, 0), (),)

	def Stop(self):
		"""Puts the multimedia stream into Stopped state"""
		return self._oleobj_.InvokeTypes(1610743811, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"AllowChangeDisplayMode": (33, 2, (11, 0), (), "AllowChangeDisplayMode", None),
		"AllowHideControls": (31, 2, (11, 0), (), "AllowHideControls", None),
		"AllowHideDisplay": (30, 2, (11, 0), (), "AllowHideDisplay", None),
		"Appearance": (-520, 2, (3, 0), (), "Appearance", None),
		"Author": (6, 2, (8, 0), (), "Author", None),
		"AutoRewind": (41, 2, (11, 0), (), "AutoRewind", None),
		"AutoStart": (40, 2, (11, 0), (), "AutoStart", None),
		"Balance": (20, 2, (3, 0), (), "Balance", None),
		"BorderStyle": (42, 2, (3, 0), (), "BorderStyle", None),
		"Copyright": (8, 2, (8, 0), (), "Copyright", None),
		"CurrentPosition": (13, 2, (5, 0), (), "CurrentPosition", None),
		"CurrentState": (17, 2, (3, 0), (), "CurrentState", None),
		"Description": (9, 2, (8, 0), (), "Description", None),
		"DisplayBackColor": (37, 2, (19, 0), (), "DisplayBackColor", None),
		"DisplayForeColor": (36, 2, (19, 0), (), "DisplayForeColor", None),
		"DisplayMode": (32, 2, (3, 0), (), "DisplayMode", None),
		"Duration": (12, 2, (5, 0), (), "Duration", None),
		"EnableContextMenu": (21, 2, (11, 0), (), "EnableContextMenu", None),
		"EnablePositionControls": (27, 2, (11, 0), (), "EnablePositionControls", None),
		"EnableSelectionControls": (28, 2, (11, 0), (), "EnableSelectionControls", None),
		"EnableTracker": (29, 2, (11, 0), (), "EnableTracker", None),
		"Enabled": (-514, 2, (11, 0), (), "Enabled", None),
		"FileName": (11, 2, (8, 0), (), "FileName", None),
		"FilterGraph": (34, 2, (13, 0), (), "FilterGraph", None),
		"FilterGraphDispatch": (35, 2, (9, 0), (), "FilterGraphDispatch", None),
		"FullScreenMode": (39, 2, (11, 0), (), "FullScreenMode", None),
		"ImageSourceHeight": (5, 2, (3, 0), (), "ImageSourceHeight", None),
		"ImageSourceWidth": (4, 2, (3, 0), (), "ImageSourceWidth", None),
		"Info": (1610743885, 2, (3, 0), (), "Info", None),
		"MediaPlayer": (1111, 2, (9, 0), (), "MediaPlayer", None),
		"MovieWindowSize": (38, 2, (3, 0), (), "MovieWindowSize", None),
		"PlayCount": (14, 2, (3, 0), (), "PlayCount", None),
		"Rate": (18, 2, (5, 0), (), "Rate", None),
		"Rating": (10, 2, (8, 0), (), "Rating", None),
		"ReadyState": (-525, 2, (3, 0), (), "ReadyState", None),
		"SelectionEnd": (16, 2, (5, 0), (), "SelectionEnd", None),
		"SelectionStart": (15, 2, (5, 0), (), "SelectionStart", None),
		"ShowControls": (23, 2, (11, 0), (), "ShowControls", None),
		"ShowDisplay": (22, 2, (11, 0), (), "ShowDisplay", None),
		"ShowPositionControls": (24, 2, (11, 0), (), "ShowPositionControls", None),
		"ShowSelectionControls": (25, 2, (11, 0), (), "ShowSelectionControls", None),
		"ShowTracker": (26, 2, (11, 0), (), "ShowTracker", None),
		"Title": (7, 2, (8, 0), (), "Title", None),
		"Volume": (19, 2, (3, 0), (), "Volume", None),
		"hWnd": (-515, 2, (3, 0), (), "hWnd", None),
	}
	_prop_map_put_ = {
		"AllowChangeDisplayMode": ((33, LCID, 4, 0),()),
		"AllowHideControls": ((31, LCID, 4, 0),()),
		"AllowHideDisplay": ((30, LCID, 4, 0),()),
		"Appearance": ((-520, LCID, 4, 0),()),
		"AutoRewind": ((41, LCID, 4, 0),()),
		"AutoStart": ((40, LCID, 4, 0),()),
		"Balance": ((20, LCID, 4, 0),()),
		"BorderStyle": ((42, LCID, 4, 0),()),
		"CurrentPosition": ((13, LCID, 4, 0),()),
		"DisplayBackColor": ((37, LCID, 4, 0),()),
		"DisplayForeColor": ((36, LCID, 4, 0),()),
		"DisplayMode": ((32, LCID, 4, 0),()),
		"EnableContextMenu": ((21, LCID, 4, 0),()),
		"EnablePositionControls": ((27, LCID, 4, 0),()),
		"EnableSelectionControls": ((28, LCID, 4, 0),()),
		"EnableTracker": ((29, LCID, 4, 0),()),
		"Enabled": ((-514, LCID, 4, 0),()),
		"FileName": ((11, LCID, 4, 0),()),
		"FilterGraph": ((34, LCID, 4, 0),()),
		"FullScreenMode": ((39, LCID, 4, 0),()),
		"MovieWindowSize": ((38, LCID, 4, 0),()),
		"PlayCount": ((14, LCID, 4, 0),()),
		"Rate": ((18, LCID, 4, 0),()),
		"SelectionEnd": ((16, LCID, 4, 0),()),
		"SelectionStart": ((15, LCID, 4, 0),()),
		"ShowControls": ((23, LCID, 4, 0),()),
		"ShowDisplay": ((22, LCID, 4, 0),()),
		"ShowPositionControls": ((24, LCID, 4, 0),()),
		"ShowSelectionControls": ((25, LCID, 4, 0),()),
		"ShowTracker": ((26, LCID, 4, 0),()),
		"Volume": ((19, LCID, 4, 0),()),
	}

from win32com.client import CoClassBaseClass
# This CoClass is known by the name 'AMOVIE.ActiveMovieControl.2'
class ActiveMovie(CoClassBaseClass): # A CoClass
	# Microsoft ActiveMovie Control
	CLSID = IID('{05589FA1-C356-11CE-BF01-00AA0055595A}')
	coclass_sources = [
		DActiveMovieEvents2,
	]
	default_source = DActiveMovieEvents2
	coclass_interfaces = [
		IActiveMovie3,
		DActiveMovieEvents,
		IActiveMovie2,
		IActiveMovie,
	]
	default_interface = IActiveMovie3

IActiveMovie_vtables_dispatch_ = 1
IActiveMovie_vtables_ = [
	(('AboutBox',), -552, (-552, (), [], 1, 1, 4, 0, 28, (3, 0, None, None), 0)),
	(('Run',), 1610743809, (1610743809, (), [], 1, 1, 4, 0, 32, (3, 0, None, None), 0)),
	(('Pause',), 1610743810, (1610743810, (), [], 1, 1, 4, 0, 36, (3, 0, None, None), 0)),
	(('Stop',), 1610743811, (1610743811, (), [], 1, 1, 4, 0, 40, (3, 0, None, None), 0)),
	(('ImageSourceWidth', 'pWidth'), 4, (4, (), [(16387, 10, None, None)], 1, 2, 4, 0, 44, (3, 0, None, None), 0)),
	(('ImageSourceHeight', 'pHeight'), 5, (5, (), [(16387, 10, None, None)], 1, 2, 4, 0, 48, (3, 0, None, None), 0)),
	(('Author', 'pbstrAuthor'), 6, (6, (), [(16392, 10, None, None)], 1, 2, 4, 0, 52, (3, 0, None, None), 0)),
	(('Title', 'pbstrTitle'), 7, (7, (), [(16392, 10, None, None)], 1, 2, 4, 0, 56, (3, 0, None, None), 0)),
	(('Copyright', 'pbstrCopyright'), 8, (8, (), [(16392, 10, None, None)], 1, 2, 4, 0, 60, (3, 0, None, None), 0)),
	(('Description', 'pbstrDescription'), 9, (9, (), [(16392, 10, None, None)], 1, 2, 4, 0, 64, (3, 0, None, None), 0)),
	(('Rating', 'pbstrRating'), 10, (10, (), [(16392, 10, None, None)], 1, 2, 4, 0, 68, (3, 0, None, None), 0)),
	(('FileName', 'pbstrFileName'), 11, (11, (), [(16392, 10, None, None)], 1, 2, 4, 0, 72, (3, 0, None, None), 0)),
	(('FileName', 'pbstrFileName'), 11, (11, (), [(8, 1, None, None)], 1, 4, 4, 0, 76, (3, 0, None, None), 0)),
	(('Duration', 'pValue'), 12, (12, (), [(16389, 10, None, None)], 1, 2, 4, 0, 80, (3, 0, None, None), 0)),
	(('CurrentPosition', 'pValue'), 13, (13, (), [(16389, 10, None, None)], 1, 2, 4, 0, 84, (3, 0, None, None), 0)),
	(('CurrentPosition', 'pValue'), 13, (13, (), [(5, 1, None, None)], 1, 4, 4, 0, 88, (3, 0, None, None), 0)),
	(('PlayCount', 'pPlayCount'), 14, (14, (), [(16387, 10, None, None)], 1, 2, 4, 0, 92, (3, 0, None, None), 0)),
	(('PlayCount', 'pPlayCount'), 14, (14, (), [(3, 1, None, None)], 1, 4, 4, 0, 96, (3, 0, None, None), 0)),
	(('SelectionStart', 'pValue'), 15, (15, (), [(16389, 10, None, None)], 1, 2, 4, 0, 100, (3, 0, None, None), 0)),
	(('SelectionStart', 'pValue'), 15, (15, (), [(5, 1, None, None)], 1, 4, 4, 0, 104, (3, 0, None, None), 0)),
	(('SelectionEnd', 'pValue'), 16, (16, (), [(16389, 10, None, None)], 1, 2, 4, 0, 108, (3, 0, None, None), 0)),
	(('SelectionEnd', 'pValue'), 16, (16, (), [(5, 1, None, None)], 1, 4, 4, 0, 112, (3, 0, None, None), 0)),
	(('CurrentState', 'pState'), 17, (17, (), [(16387, 10, None, None)], 1, 2, 4, 0, 116, (3, 0, None, None), 0)),
	(('Rate', 'pValue'), 18, (18, (), [(16389, 10, None, None)], 1, 2, 4, 0, 120, (3, 0, None, None), 0)),
	(('Rate', 'pValue'), 18, (18, (), [(5, 1, None, None)], 1, 4, 4, 0, 124, (3, 0, None, None), 0)),
	(('Volume', 'pValue'), 19, (19, (), [(16387, 10, None, None)], 1, 2, 4, 0, 128, (3, 0, None, None), 0)),
	(('Volume', 'pValue'), 19, (19, (), [(3, 1, None, None)], 1, 4, 4, 0, 132, (3, 0, None, None), 0)),
	(('Balance', 'pValue'), 20, (20, (), [(16387, 10, None, None)], 1, 2, 4, 0, 136, (3, 0, None, None), 0)),
	(('Balance', 'pValue'), 20, (20, (), [(3, 1, None, None)], 1, 4, 4, 0, 140, (3, 0, None, None), 0)),
	(('EnableContextMenu', 'pEnable'), 21, (21, (), [(16395, 10, None, None)], 1, 2, 4, 0, 144, (3, 0, None, None), 0)),
	(('EnableContextMenu', 'pEnable'), 21, (21, (), [(11, 1, None, None)], 1, 4, 4, 0, 148, (3, 0, None, None), 0)),
	(('ShowDisplay', 'Show'), 22, (22, (), [(16395, 10, None, None)], 1, 2, 4, 0, 152, (3, 0, None, None), 0)),
	(('ShowDisplay', 'Show'), 22, (22, (), [(11, 1, None, None)], 1, 4, 4, 0, 156, (3, 0, None, None), 0)),
	(('ShowControls', 'Show'), 23, (23, (), [(16395, 10, None, None)], 1, 2, 4, 0, 160, (3, 0, None, None), 0)),
	(('ShowControls', 'Show'), 23, (23, (), [(11, 1, None, None)], 1, 4, 4, 0, 164, (3, 0, None, None), 0)),
	(('ShowPositionControls', 'Show'), 24, (24, (), [(16395, 10, None, None)], 1, 2, 4, 0, 168, (3, 0, None, None), 0)),
	(('ShowPositionControls', 'Show'), 24, (24, (), [(11, 1, None, None)], 1, 4, 4, 0, 172, (3, 0, None, None), 0)),
	(('ShowSelectionControls', 'Show'), 25, (25, (), [(16395, 10, None, None)], 1, 2, 4, 0, 176, (3, 0, None, None), 0)),
	(('ShowSelectionControls', 'Show'), 25, (25, (), [(11, 1, None, None)], 1, 4, 4, 0, 180, (3, 0, None, None), 0)),
	(('ShowTracker', 'Show'), 26, (26, (), [(16395, 10, None, None)], 1, 2, 4, 0, 184, (3, 0, None, None), 0)),
	(('ShowTracker', 'Show'), 26, (26, (), [(11, 1, None, None)], 1, 4, 4, 0, 188, (3, 0, None, None), 0)),
	(('EnablePositionControls', 'Enable'), 27, (27, (), [(16395, 10, None, None)], 1, 2, 4, 0, 192, (3, 0, None, None), 0)),
	(('EnablePositionControls', 'Enable'), 27, (27, (), [(11, 1, None, None)], 1, 4, 4, 0, 196, (3, 0, None, None), 0)),
	(('EnableSelectionControls', 'Enable'), 28, (28, (), [(16395, 10, None, None)], 1, 2, 4, 0, 200, (3, 0, None, None), 0)),
	(('EnableSelectionControls', 'Enable'), 28, (28, (), [(11, 1, None, None)], 1, 4, 4, 0, 204, (3, 0, None, None), 0)),
	(('EnableTracker', 'Enable'), 29, (29, (), [(16395, 10, None, None)], 1, 2, 4, 0, 208, (3, 0, None, None), 0)),
	(('EnableTracker', 'Enable'), 29, (29, (), [(11, 1, None, None)], 1, 4, 4, 0, 212, (3, 0, None, None), 0)),
	(('AllowHideDisplay', 'Show'), 30, (30, (), [(16395, 10, None, None)], 1, 2, 4, 0, 216, (3, 0, None, None), 0)),
	(('AllowHideDisplay', 'Show'), 30, (30, (), [(11, 1, None, None)], 1, 4, 4, 0, 220, (3, 0, None, None), 0)),
	(('AllowHideControls', 'Show'), 31, (31, (), [(16395, 10, None, None)], 1, 2, 4, 0, 224, (3, 0, None, None), 0)),
	(('AllowHideControls', 'Show'), 31, (31, (), [(11, 1, None, None)], 1, 4, 4, 0, 228, (3, 0, None, None), 0)),
	(('DisplayMode', 'pValue'), 32, (32, (), [(16387, 10, None, None)], 1, 2, 4, 0, 232, (3, 0, None, None), 0)),
	(('DisplayMode', 'pValue'), 32, (32, (), [(3, 1, None, None)], 1, 4, 4, 0, 236, (3, 0, None, None), 0)),
	(('AllowChangeDisplayMode', 'fAllow'), 33, (33, (), [(16395, 10, None, None)], 1, 2, 4, 0, 240, (3, 0, None, None), 0)),
	(('AllowChangeDisplayMode', 'fAllow'), 33, (33, (), [(11, 1, None, None)], 1, 4, 4, 0, 244, (3, 0, None, None), 0)),
	(('FilterGraph', 'ppFilterGraph'), 34, (34, (), [(16397, 10, None, None)], 1, 2, 4, 0, 248, (3, 0, None, None), 0)),
	(('FilterGraph', 'ppFilterGraph'), 34, (34, (), [(13, 1, None, None)], 1, 4, 4, 0, 252, (3, 0, None, None), 0)),
	(('FilterGraphDispatch', 'pDispatch'), 35, (35, (), [(16393, 10, None, None)], 1, 2, 4, 0, 256, (3, 0, None, None), 0)),
	(('DisplayForeColor', 'ForeColor'), 36, (36, (), [(16403, 10, None, None)], 1, 2, 4, 0, 260, (3, 0, None, None), 0)),
	(('DisplayForeColor', 'ForeColor'), 36, (36, (), [(19, 1, None, None)], 1, 4, 4, 0, 264, (3, 0, None, None), 0)),
	(('DisplayBackColor', 'BackColor'), 37, (37, (), [(16403, 10, None, None)], 1, 2, 4, 0, 268, (3, 0, None, None), 0)),
	(('DisplayBackColor', 'BackColor'), 37, (37, (), [(19, 1, None, None)], 1, 4, 4, 0, 272, (3, 0, None, None), 0)),
	(('MovieWindowSize', 'WindowSize'), 38, (38, (), [(16387, 10, None, None)], 1, 2, 4, 0, 276, (3, 0, None, None), 0)),
	(('MovieWindowSize', 'WindowSize'), 38, (38, (), [(3, 1, None, None)], 1, 4, 4, 0, 280, (3, 0, None, None), 0)),
	(('FullScreenMode', 'pEnable'), 39, (39, (), [(16395, 10, None, None)], 1, 2, 4, 0, 284, (3, 0, None, None), 0)),
	(('FullScreenMode', 'pEnable'), 39, (39, (), [(11, 1, None, None)], 1, 4, 4, 0, 288, (3, 0, None, None), 0)),
	(('AutoStart', 'pEnable'), 40, (40, (), [(16395, 10, None, None)], 1, 2, 4, 0, 292, (3, 0, None, None), 0)),
	(('AutoStart', 'pEnable'), 40, (40, (), [(11, 1, None, None)], 1, 4, 4, 0, 296, (3, 0, None, None), 0)),
	(('AutoRewind', 'pEnable'), 41, (41, (), [(16395, 10, None, None)], 1, 2, 4, 0, 300, (3, 0, None, None), 0)),
	(('AutoRewind', 'pEnable'), 41, (41, (), [(11, 1, None, None)], 1, 4, 4, 0, 304, (3, 0, None, None), 0)),
	(('hWnd', 'hWnd'), -515, (-515, (), [(16387, 10, None, None)], 1, 2, 4, 0, 308, (3, 0, None, None), 0)),
	(('Appearance', 'pAppearance'), -520, (-520, (), [(16387, 10, None, None)], 1, 2, 4, 0, 312, (3, 0, None, None), 0)),
	(('Appearance', 'pAppearance'), -520, (-520, (), [(3, 1, None, None)], 1, 4, 4, 0, 316, (3, 0, None, None), 0)),
	(('BorderStyle', 'pBorderStyle'), 42, (42, (), [(16387, 10, None, None)], 1, 2, 4, 0, 320, (3, 0, None, None), 0)),
	(('BorderStyle', 'pBorderStyle'), 42, (42, (), [(3, 1, None, None)], 1, 4, 4, 0, 324, (3, 0, None, None), 0)),
	(('Enabled', 'pEnabled'), -514, (-514, (), [(16395, 10, None, None)], 1, 2, 4, 0, 328, (3, 0, None, None), 0)),
	(('Enabled', 'pEnabled'), -514, (-514, (), [(11, 1, None, None)], 1, 4, 4, 0, 332, (3, 0, None, None), 0)),
	(('Info', 'ppInfo'), 1610743885, (1610743885, (), [(16387, 10, None, None)], 1, 2, 4, 0, 336, (3, 0, None, None), 64)),
]

IActiveMovie2_vtables_dispatch_ = 1
IActiveMovie2_vtables_ = [
	(('IsSoundCardEnabled', 'pbSoundCard'), 53, (53, (), [(16395, 10, None, None)], 1, 1, 4, 0, 340, (3, 0, None, None), 0)),
	(('ReadyState', 'pValue'), -525, (-525, (), [(16387, 10, None, None)], 1, 2, 4, 0, 344, (3, 0, None, None), 0)),
]

IActiveMovie3_vtables_dispatch_ = 1
IActiveMovie3_vtables_ = [
	(('MediaPlayer', 'ppDispatch'), 1111, (1111, (), [(16393, 10, None, None)], 1, 2, 4, 0, 348, (3, 0, None, None), 0)),
]

RecordMap = {
}

CLSIDToClassMap = {
	'{B6CD6553-E9CB-11D0-821F-00A0C91F9CA0}' : DActiveMovieEvents2,
	'{B6CD6554-E9CB-11D0-821F-00A0C91F9CA0}' : IActiveMovie2,
	'{265EC140-AE62-11D1-8500-00A0C91F9CA0}' : IActiveMovie3,
	'{05589FA1-C356-11CE-BF01-00AA0055595A}' : ActiveMovie,
	'{05589FA2-C356-11CE-BF01-00AA0055595A}' : IActiveMovie,
	'{05589FA3-C356-11CE-BF01-00AA0055595A}' : DActiveMovieEvents,
}
CLSIDToPackageMap = {}
win32com.client.CLSIDToClass.RegisterCLSIDsFromDict( CLSIDToClassMap )
VTablesToPackageMap = {}
VTablesToClassMap = {
	'{265EC140-AE62-11D1-8500-00A0C91F9CA0}' : 'IActiveMovie3',
	'{B6CD6554-E9CB-11D0-821F-00A0C91F9CA0}' : 'IActiveMovie2',
	'{05589FA2-C356-11CE-BF01-00AA0055595A}' : 'IActiveMovie',
}


NamesToIIDMap = {
	'IActiveMovie2' : '{B6CD6554-E9CB-11D0-821F-00A0C91F9CA0}',
	'IActiveMovie3' : '{265EC140-AE62-11D1-8500-00A0C91F9CA0}',
	'DActiveMovieEvents2' : '{B6CD6553-E9CB-11D0-821F-00A0C91F9CA0}',
	'DActiveMovieEvents' : '{05589FA3-C356-11CE-BF01-00AA0055595A}',
	'IActiveMovie' : '{05589FA2-C356-11CE-BF01-00AA0055595A}',
}

win32com.client.constants.__dicts__.append(constants.__dict__)


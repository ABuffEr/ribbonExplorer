# -*- coding: UTF-8 -*-
# Author: Alberto Buffolino
# License: GPLv2
from tones import beep
import globalPluginHandler
import addonHandler
import controlTypes as ct
import ui
import api
import speech
from keyboardHandler import KeyboardInputGesture as InputGesture
import braille
import inspect
from NVDAObjects import NVDAObject
from NVDAObjects.IAccessible import IAccessible
import review
import scriptHandler
from .utils import *
# for compatibility
REASON_FOCUS = ct.OutputReason.FOCUS if hasattr(ct, "OutputReason") else ct.REASON_FOCUS
if hasattr(ct, 'Role'):
	roles = ct.Role
else:
	roles = type('Enum', (), dict([(x.split("ROLE_")[1], getattr(ct, x)) for x in dir(ct) if x.startswith("ROLE_")]))
if hasattr(ct, 'State'):
	states = ct.State
else:
	states = type('Enum', (), dict([(x.split("STATE_")[1], getattr(ct, x)) for x in dir(ct) if x.startswith("STATE_")]))
# to take advantages from NVDA translations
NVDALocale = _
addonHandler.initTranslation()

"""
**Important dev info:**
Ribbon is a nightmare. So, this code is a nightmare upon another one.
It's heavily event-driven, as you see in event_focusEntered and event_gainFocus. But the main, esoteric thing you have to understand is how NVDA decides what object to return, , when you invoke .simple* methods.
That is, according to obj.presentationType property.
This can be:
- content: obj can be retrieved, NVDA returns it;
- layout: obj cannot be retrieved, NVDA returns the first retrievable parent or child (if present), according to specified relation/direction;
- unavailable: obj and its children are not retrievable, NVDA returns the first available obj with specified relation/direction.
So, for example, imagine to have obj1, that has obj2 and obj3 as children, and obj3 has obj4 as child.
Scenario 1: obj3.presentationType is layout and obj1 and obj4 ones are content, then:
- obj1.simpleLastChild returns obj4;
- obj4.simpleParent returns obj1;
- obj2.simpleNext returns obj4.
Scenario 2: obj2.presentationType is unavailable and obj1 and obj3 ones are content, then:
- obj3.simplePrevious is None;
- obj4.simpleParent is obj3;
- obj3.simpleParent is obj1.
Scenario 3: obj3.presentationType is unavailable and obj2 one is content, then:
- obj1.simpleLastChild is obj2;
- obj2.simpleNext is None.
Setting this property following our needs lets the code to be more simple, than managing a lot of exceptions in object navigation.
I said "more simple", not simple, though... even due to three possible views of the Ribbon, that are:
v1: full screen mode;
v2: tabs only (default, maybe);
v3: multifunction bar/tabs and commands (menu associated with current tab is always shown).
It's all, for now...
"""
# todo: fix closing account submenu
# todo: fix expansion of search in v3
# todo: settings dialog
# todo: make a option for superEsc
# todo: return to menubar consistently
# todo: fix alt+downArrow in v1 and v3 expanding menubar items

# speed-up
content = NVDAObject.presType_content
layout = NVDAObject.presType_layout
unavailable = NVDAObject.presType_unavailable

class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	# starting variables
	# to speed-up negative check in event_focusEntered
	supportedApp = False
	# determine if Ribbon exploration is active
	exploring = False
	# save and then restore initial review mode
	startReviewMode = None
	# keep track of object presented to the user by the add-on
	userObj = None
	# keep of userObj focus status (see reportUser)
	userObjHasFocus = False
	# to go back on expanded tab in menubar (see collapseMenu)
	menubar = []
	# list of objects to hide keeping their children
	layoutableObj = []
	# to adjust focus when expanding a menu tab
	isExpandingMenu = False
	# list expanded menu tab (ideally one)
	expandedMenu = []
	# to adjust focus when expanding a submenu
	isExpandingSubmenu = False
	# list expanded submenus (potentially nested)
	expandedSubmenu = []
	# to control events and avoid focus problems/lost
	isCollapsingSubmenu = False
	# list of initial menu item(s) expanded in a submenu
	collapsingMenuItem = []

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		if not self.exploring:
			return
		if not obj:
			return
#		debugLog("Running %s for obj %s,%s"%(inspect.currentframe().f_code.co_name,obj.name, obj.role))
		# speed-up
		global content, layout, unavailable
		objRole = obj.role
		if objRole == roles.EDITABLETEXT and isRibbonRoot(obj.simpleParent):
			clsList.insert(0, EditWithoutSelection)
			return
		if isRibbonRoot(obj):
			# to simplify check of menubar items
#			debugLog("Root, set content")
			obj.presentationType = content
			# speed-up: always return immediately
			return
		elif obj in self.layoutableObj:
#			debugLog("Redundant obj whose we want children of, set layout")
			obj.presentationType = layout
			return
		elif not obj.name and objRole == roles.MENUITEM:
#			debugLog("Anonymous menuitem, set layout")
			obj.presentationType = layout
			return
		elif objRole == roles.TABCONTROL:
			# to simplify menubar exploration (enforcing)
#			debugLog("Role tabcontrol, set layout")
			obj.presentationType = layout
			return
		elif isSubtab(obj):
#			debugLog("Subtab, set content")
			obj.presentationType = content
			return
		elif obj in self.expandedMenu:
			# for expanded menu
#			debugLog("ExpandedMenu, set layout")
			obj.presentationType = layout
			return
		elif objRole == roles.POPUPMENU and not obj.states:
			# to select this as simpleParent when closing submenu
#			debugLog("Role popupmenu without states, set content")
			obj.presentationType = content
			return
		elif hasattr(obj, "UIAElement") and obj.UIAElement.cachedClassName in ("NetUIRepeatButton", "NetUIScrollBar", "NetUIAppFrameHelper"):
			# to hide scrolling and window-action buttons
#			debugLog("CachedClassName %s, set unavailable"%obj.UIAElement.cachedClassName)
			obj.presentationType = unavailable
			return
		elif not obj.name or obj.name.isspace():
			# generic
#			debugLog("Anonymous obj, set layout")
			obj.presentationType = layout
			return
		elif obj.presentationType == unavailable:
			# for menu items not currently available,
			# but which we want to show to users
#			debugLog("PresType unavailable, set content")
			obj.presentationType = content
			return
		elif objRole == roles.DATAGRID and obj.description:
			# to hide in grouping (it should be a grid associated to a visible button)
#			debugLog("Role datagrid has description, set unavailable")
			obj.presentationType = unavailable
			return
		elif objRole == roles.DATAGRID and not obj.description:
			# to show in submenu
			if allObjPassCheck(lambda i: i.role == roles.GROUPING and i.presentationType == i.presType_content, obj.children):
				obj.presentationType = layout
			else:
				# ...manage other cases
#				debugLog("...set content")
				obj.presentationType = content
			return
		elif objRole == roles.LIST:
			# to explore children only
#			debugLog("Role list, set layout")
			obj.presentationType = layout
			return
		elif objRole == roles.GROUPING:
#			debugLog("Role grouping, set content")
			obj.presentationType = content
			return
		elif objRole in (roles.GRAPHIC, roles.STATICTEXT):
			# rare and useless, hide
#			debugLog("Role %s, set unavailable"%objRole)
			obj.presentationType = unavailable

	def event_foreground(self, obj, nextHandler):
		nextHandler()
		if isOfficeApp(obj):
			self.supportedApp = True
		else:
			self.supportedApp = False

	def event_focusEntered(self, obj, nextHandler):
		if not self.supportedApp:
			nextHandler()
			return
		elif not self.exploring and isRibbonRoot(obj):
			debugLog("Exploration starts")
			obj.presentationType = obj.presType_content
			self.explorationStart()
			return
		elif not self.exploring:
			# stop immediately
			nextHandler()
			return
		debugLog("Running %s for obj %s,%s"%(inspect.currentframe().f_code.co_name,obj.name,obj.role))
		if obj.role == roles.MENUITEM and isRibbonRoot(obj.parent):
			# in v1
			obj.presentationType = obj.presType_layout
			return
		elif obj.role == roles.TABCONTROL:
			debugLog("Mute %s"%obj.role)
			obj.presentationType = obj.presType_layout
			return
		elif self.isExpandingMenu:
			# self.userObj should be a menu tab, set by last gainFocus
			# while obj the child of lower multi tab, containing current menu items
			if isSubtab(obj.parent):
				debugLog("Found groupMenu %s"%obj.name)
				self.expandedMenu.append(obj)
			debugLog("Ignore event")
			return
		elif self.isExpandingSubmenu:
			if obj.role in (roles.TOOLBAR, roles.POPUPMENU):
				debugLog("Found submenu")
				self.expandedSubmenu.append(obj)
			debugLog("Ignore event")
			return
		elif self.isCollapsingSubmenu:
			debugLog("Ignore event")
			return
		debugLog("Process event")
		nextHandler()

	def event_gainFocus(self, obj, nextHandler):
		if not self.exploring:
			nextHandler()
			return
		debugLog("Running %s for obj %s,%s"%(inspect.currentframe().f_code.co_name,obj.name,obj.role))
		if self.isExpandingMenu:
			# first gainFocus after expandMenu claims expansion as terminated
			# and performs action for adjusting focus
			debugLog("Event raises expandedMenuAction")
			self.expandedMenuAction()
			return
		elif self.isExpandingSubmenu:
			debugLog("Event raises expandedSubmenuAction")
			self.expandedSubmenuAction()
			return
		elif self.isCollapsingSubmenu:
			# avoid exploration exit when closing submenu
			if obj == self.collapsingMenuItem[-1]:
				# focus returned on menu
				self.isCollapsingSubmenu = False
				self.collapsingMenuItem.pop()
				debugLog("Set %s as userObj"%obj.name)
				self.userObj = obj
				nextHandler()
				debugLog("Successfully closing submenu without exploration exit")
			return
		elif not isRibbonInAncestors() or obj.role in (roles.EDITABLETEXT,):
			self.explorationEnd()
		else:
			debugLog("Set %s as userObj"%obj.name)
			self.userObj = obj
		debugLog("Process event")
		nextHandler()

	def event_loseFocus(self, obj, nextHandler):
		if not self.exploring:
			nextHandler()
			return
		debugLog("Running %s for obj %s,%s"%(inspect.currentframe().f_code.co_name,obj.name,obj.role))
		if not obj.name and obj.role == roles.MENUITEM:
			debugLog("Collapsing submenu. Ignore event to avoid exploration ending")
			return
		nextHandler()

	def explorationStart(self):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		self.exploring = True
		self.startReviewMode = review.getCurrentMode()
		review.setCurrentMode("object", updateReviewPosition=False)
		self.userObj = api.getFocusObject()
		self.userObjHasFocus = True
		# simple gestures
		for gesture in ("tab", "escape", "enter", "downArrow", "leftArrow", "rightArrow", "upArrow"):
			self.bindGesture("kb:%s"%gesture, gesture)
		self.bindGesture("kb:shift+tab", "shiftTab")
		self.bindGesture("kb:alt+upArrow", "altUpArrow")
		self.bindGesture("kb:alt+downArrow", "altDownArrow")
		self.bindGesture("kb:NVDA+space", "toggleExploration")
		# for debug
#		self.bindGesture("kb:i", "debug")

	def explorationEnd(self):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		self.exploring = False
		review.setCurrentMode(self.startReviewMode, updateReviewPosition=False)
		self.startReviewMode = None
		self.userObj = None
		self.userObjHasFocus = False
		self.menubar.clear()
		self.layoutableObj.clear()
		self.isExpandingMenu = False
		self.expandedMenu.clear()
		self.isExpandingSubmenu = False
		self.expandedSubmenu.clear()
		self.isCollapsingSubmenu = False
		self.collapsingMenuItem.clear()
		self.clearGestureBindings()

	def script_tab(self, gesture):
		self.nextItem()

	def script_shiftTab(self, gesture):
		self.prevItem()

	def script_escape(self, gesture):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		superEsc = True
		try:
			if isRibbonRoot(self.userObj.simpleParent) or isSubtab(self.userObj.parent):
				debugLog("isRibbon or isTab, pass escape")
				gesture.send()
			elif superEsc:
				debugLog("SuperEsc, go to main menu")
				self.collapseMenu()
			elif self.userObj.parent in self.expandedMenu:
				debugLog("Escaping from expandedMenu, focus menu in menubar")
				self.collapseMenu()
			else:
				debugLog("Go to parent")
				self.parentItem()
		except: # ensure end in case of problems
			debugLog("Exception, terminate exploration")
			self.explorationEnd()

	def script_downArrow(self, gesture):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		if isRibbonRoot(self.userObj.simpleParent):
			self.expandMenu(self.userObj)
		else:
			self.nextItem()

	def script_upArrow(self, gesture):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		if isRibbonRoot(self.userObj.simpleParent):
			self.expandMenu(self.userObj)
		else:
			self.prevItem()

	def script_leftArrow(self, gesture):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		if isRibbonRoot(self.userObj.simpleParent):
			self.prevMenu()
		else:
			self.parentItem()

	def script_rightArrow(self, gesture):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		if isRibbonRoot(self.userObj.simpleParent):
			self.nextMenu()
		else:
			self.childItem()

	def script_enter(self, gesture):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		if states.OFFSCREEN in self.userObj.states:
			self.userObj.doAction()
		elif not self.userObjHasFocus or states.UNAVAILABLE in self.userObj.states:
			ui.message(NVDALocale("No action"))
		elif isRibbonRoot(self.userObj.simpleParent):
			self.expandMenu(self.userObj)
		# splitbutton must perform default action on enter
		elif states.COLLAPSED in self.userObj.states and self.userObj.role != roles.SPLITBUTTON:
			self.expandSubmenu(self.userObj)
		else:
			gesture.send()

	def script_altUpArrow(self, gesture):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		if self.expandedSubmenu:
			self.collapseSubmenu()
		elif self.expandedMenu:
			self.collapseMenu()
		else:
			ui.message(NVDALocale("No action"))

	def script_altDownArrow(self, gesture):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		if not self.userObjHasFocus or states.UNAVAILABLE in self.userObj.states:
			ui.message(NVDALocale("No action"))
		elif isRibbonRoot(self.userObj.simpleParent):
			self.expandMenu(self.userObj)
		elif states.COLLAPSED in self.userObj.states:
			self.expandSubmenu(self.userObj)

	def script_toggleExploration(self, gesture):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		if self.exploring:
			# Translators: a message when user manually disable exploration (NVDA+space)
			ui.message(_("Exploration end"))
			self.explorationEnd()

	def script_debug(self, gesture):
		ui.message("Performing debug script")
		obj = api.getNavigatorObject()
		obj.UIALegacyIAccessiblePattern.Select(2)

	def reportUser(self, obj):
		# it should not happen, but anyway...
		if obj is None:
			return
		self.userObj = obj
		self.userObjHasFocus = False
		debugLog("Report obj: %s,%s"%(obj.name,obj.role))
		# unconditionally set as navigator object
		api.setNavigatorObject(obj)
		if not obj.isFocusable:
			speech.speakObject(obj, reason=REASON_FOCUS)
			return
		api.setFocusObject(obj)
		if obj.hasFocus:
			self.userObjHasFocus = True
			speech.speakObject(obj, reason=REASON_FOCUS)
			braille.handler.handleGainFocus(obj)
		if not self.userObjHasFocus:
			try:
				obj.setFocus()
				self.userObjHasFocus = True
			except:
				debugLog("SetFocus failed, try forcing")
				self.forceFocus(obj)

	def forceFocus(self, obj):
		debugLog("Forcing focus on %s,%s"%(obj.name, obj.role))
		# offscreen obj can be reported only
		if states.OFFSCREEN in obj.states:
			speech.speakObject(obj, reason=REASON_FOCUS)
			braille.handler.handleGainFocus(obj)
			return
		# workaround: find previous/next focusable obj,
		# focus it, then simulate a tab/shift+tab
		# to get focus on obj we want
		tryAgain = True
		if obj.positionInfo: #and obj.positionInfo["indexInGroup"] not in (1, obj.positionInfo["similarItemsInGroup"]):
			scriptRef = scriptHandler._lastScriptRef
			if scriptRef and scriptRef().__name__ in ("script_downArrow",): #"script_rightArrow"):
				debugLog("Send downArrow in list")
				InputGesture.fromName("downArrow").send()
				tryAgain = False
			elif scriptRef and scriptRef().__name__ in ("script_upArrow",):
				debugLog("Send upArrow in list")
				InputGesture.fromName("upArrow").send()
				speech.speakObject(obj, reason=REASON_FOCUS)
				braille.handler.handleGainFocus(obj)
				tryAgain = False
		if tryAgain:
			prevObj = findFocusablePrevious(obj)
			if moveFocusTo(prevObj):
				debugLog("prevObj.setFocus() success; send tab")
				InputGesture.fromName("tab").send()
				tryAgain = False
		if tryAgain:
			nextObj = findFocusableNext(obj)
			if moveFocusTo(nextObj):
				debugLog("nextObj.setFocus() success; send shift+tab")
				InputGesture.fromName("shift+tab").send()
				tryAgain = False
		if tryAgain:
			# the last hope: send tab/shift+tab without knowing where the focus is
			# but first, understand what direction we're moving to
			scriptRef = scriptHandler._lastScriptRef
			if scriptRef and scriptRef().__name__ in ("script_tab", "script_downArrow", "script_rightArrow"):
				debugLog("Send tab blindly")
				InputGesture.fromName("tab").send()
			elif scriptRef:
				debugLog("Send shift+tab blindly")
				InputGesture.fromName("shift+tab").send()
		# and now, see where focus is
		curFocus = api.getFocusObject()
		if (curFocus.name, curFocus.role) == (obj.name, obj.role):
			debugLog("Focus moved successfully")
			self.userObj = curFocus
			self.userObjHasFocus = True
		else:
			# guarantee an output
			debugLog("Moving focus definitely failed")
			self.userObj = obj
			self.userObjHasFocus = False
			speech.speakObject(obj, reason=REASON_FOCUS)
			braille.handler.handleGainFocus(obj)

	def expandMenu(self, menu):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		if menu.UIAElement.cachedClassName != "NetUIRibbonTab" and menu.role != roles.MENUITEM and states.COLLAPSED not in menu.states:
			return
		# consider expandable menuitem in main menubar as submenu
		# (like View options)
		if menu.role == roles.MENUITEM:
			self.isExpandingSubmenu = True
			self.collapsingMenuItem.append(menu)
		# no post action for File tab
		elif menu.role != roles.BUTTON:
			self.isExpandingMenu = True
			self.menubar.append(menu)
		debugLog("List %s as in menubar"%menu.name)
		if menu.role == roles.MENUITEM:
			InputGesture.fromName("alt+downArrow").send()
		elif states.SELECTED not in menu.states:
			menu.doAction()
		else:
			InputGesture.fromName("downArrow").send()

	def expandedMenuAction(self):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		try:
			groupMenu = self.expandedMenu[-1]
		except IndexError:
			# Translators: a message when something goes wrong and exploration ends
			ui.message(_("Exploration end"))
			self.explorationEnd()
			return
		newObj = groupMenu.simpleFirstChild
		if newObj.name in groupMenu.name:
			self.layoutableObj.append(newObj)
			newObj = newObj.simpleFirstChild
		self.isExpandingMenu = False
		self.reportUser(newObj)

	def collapseMenu(self):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		newObj = self.menubar.pop()
		self.expandedMenu.clear()
		self.expandedSubmenu.clear()
		self.collapsingMenuItem.clear()
		self.reportUser(newObj)

	def expandSubmenu(self, submenu):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		if states.UNAVAILABLE in submenu.states:
			# submenu cannot be expanded
			return
		self.collapsingMenuItem.append(submenu)
		if not self.userObjHasFocus and submenu.role == roles.COMBOBOX:
			debugLog("Try to focus a child")
			tryObj = findFirstFocusable(submenu)
			tryObj.setFocus()
		else:
			debugLog("Expanding submenu")
			self.isExpandingSubmenu = True
			InputGesture.fromName("alt+downArrow").send()

	def expandedSubmenuAction(self):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		try:
			groupMenu = self.expandedSubmenu[-1]
		except IndexError:
			# Translators: a message when something goes wrong and exploration ends
			ui.message(_("Exploration end"))
			self.explorationEnd()
			return
		newObj = groupMenu.simpleFirstChild
		if not newObj.isFocusable and newObj.name in groupMenu.name:
			self.layoutableObj.append(newObj)
			newObj = newObj.simpleFirstChild
		self.isExpandingSubmenu = False
		self.reportUser(newObj)

	def collapseSubmenu(self):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		self.expandedSubmenu.pop()
		self.isCollapsingSubmenu = True
		InputGesture.fromName("alt+upArrow").send()

	def nextItem(self):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		nextObj = self.userObj.simpleNext
		# for circular scrolling
		if not nextObj:
			nextObj = self.userObj.simpleParent.simpleFirstChild
		self.reportUser(nextObj)

	def nextMenu(self):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		curMenu = self.userObj
		nextMenu = curMenu.simpleNext
		if nextMenu and isSubtab(nextMenu):
			debugLog("subtab case")
			nextMenu = nextMenu.simpleNext
		if not nextMenu:
			if curMenu.parent.role == roles.UNKNOWN:
				# in v1
				nextMenu = curMenu.parent.simpleFirstChild
			else:
				nextMenu = curMenu.simpleParent.simpleFirstChild
		self.reportUser(nextMenu)

	def prevItem(self):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		prevObj = self.userObj.simplePrevious
		# for circular scrolling
		if not prevObj:
			prevObj = self.userObj.simpleParent.simpleLastChild
		self.reportUser(prevObj)

	def prevMenu(self):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		curMenu = self.userObj
		prevMenu = curMenu.simplePrevious
		if prevMenu:
			if isSubtab(prevMenu):
				debugLog("menu under PanViewer, go previous")
				prevMenu = prevMenu.simplePrevious
			elif prevMenu.parent.role == roles.MENUITEM:
				# in v1
				debugLog("Full screen, too up! Refer to old parent")
				prevMenu = curMenu.parent.simpleLastChild
		if not prevMenu:
			debugLog("No prevMenu, go to simpleParent.simpleLastChild")
			prevMenu = curMenu.simpleParent.simpleLastChild
		self.reportUser(prevMenu)

	def parentItem(self):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		if self.userObj.parent in self.expandedMenu or isSubtab(self.userObj.parent):
			debugLog("Avoid expanded menu")
			return
		parObj = self.userObj.simpleParent
		if isSubtab(parObj):
			return
		try:
			curSubmenu = self.expandedSubmenu[-1]
			debugLog("curSubmenu is %s"%curSubmenu.name)
		except IndexError:
			debugLog("No curSubmenu")
			curSubmenu = None
		if parObj == curSubmenu:
			debugLog("Submenu condition")
			self.collapseSubmenu()
			return
		self.reportUser(parObj)

	def childItem(self):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		if states.COLLAPSED in self.userObj.states:
			self.expandSubmenu(self.userObj)
			return
		else:
			childObj = self.userObj.simpleFirstChild
		if childObj:
			self.reportUser(childObj)

class EditWithoutSelection(IAccessible):

	def script_caret_moveByLine(self, gesture):
		# avoid "selected" announcement
		self.terminateAutoSelectDetection()
		super(EditWithoutSelection, self).script_caret_moveByLine(gesture)

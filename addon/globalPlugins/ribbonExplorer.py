# -*- coding: UTF-8 -*-
# Author: Alberto Buffolino
# License: GPLv3
import globalPluginHandler
import addonHandler
import controlTypes as ct
import globalVars
import ui
import api
import speech
from keyboardHandler import KeyboardInputGesture as InputGesture
import braille
import inspect
from NVDAObjects.IAccessible import IAccessible
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
# todo: consider to check appModule.productName for Office
# todo: fix alt+downArrow in v1 and v3 expanding menubar items

# to enable logging
DEBUG = False

def debugLog(message):
	from logHandler import log
	if DEBUG:
		log.info(message)

def isOffice(obj):
	try:
		if obj.appModule.productName == "Microsoft Office":
			return True
	except:
		pass
	return False

def isRibbonRoot(obj):
	debugLog("Running %s"%inspect.currentframe().f_code.co_name)
	if obj.name == "Ribbon" and obj.role == roles.PANE:
		return True
	return False

def isSubtab(obj):
	debugLog("Running %s"%inspect.currentframe().f_code.co_name)
	if obj.role == roles.PANE and hasattr(obj, "UIAElement") and obj.UIAElement.cachedClassName == "NetUIPanViewer" and (
			# in v1 and v2
			(obj.name)
			or
			# in v3
			# parent check avoid problems in submenu
			(not obj.name and obj.parent.name and obj.parent.role == roles.GROUPING and hasattr(obj.parent, "UIAElement") and obj.parent.UIAElement.cachedClassName == "NetUIElement")
		):
		return True
	return False

def isRibbonInAncestors():
	debugLog("Running %s"%inspect.currentframe().f_code.co_name)
	for obj in reversed(globalVars.focusAncestors):
		if isRibbonRoot(obj):
			return True
	return False

def allObjPassCheck(check, objects):
	for obj in objects:
		if not check(obj):
			return False
	return True

def findFirstFocusable(obj):
	for descendant in obj.recursiveDescendants:
		if descendant.isFocusable:
			return descendant

def findFocusablePrevious(obj):
	res = obj.simplePrevious
	if not res:
		res = obj.simpleParent.simplePrevious
	if res:
		if res.isFocusable:
			return res
		elif res.childCount:
			return res.simpleLastChild

class EditWithoutSelection(IAccessible):

	def script_caret_moveByLine(self, gesture):
		self.terminateAutoSelectDetection()
		super(EditWithoutSelection, self).script_caret_moveByLine(gesture)

class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	# starting variables
	exploring = False
	userObj = None
	userObjHasFocus = False
	menubar = []
	layoutableObj = []
	isExpandingMenu = False
	expandedMenu = []
	isExpandingSubmenu = False
	expandedSubmenu = []
	isCollapsingSubmenu = False
	collapsingMenuItem = []

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		if not self.exploring:
			return
		debugLog("Running %s for obj %s,%s"%(inspect.currentframe().f_code.co_name,obj.name,obj.role))
		if obj.role == roles.EDITABLETEXT and isRibbonRoot(obj.simpleParent):
			clsList.insert(0, EditWithoutSelection)
		if isRibbonRoot(obj):
			# to simplify check of menubar items
			debugLog("Root, set content")
			obj.presentationType = "content"
			return
		elif obj in self.layoutableObj:
			debugLog("Redundant obj whose we want children of, set layout")
			obj.presentationType = "layout"
		elif not obj.name and obj.role == roles.MENUITEM:
			debugLog("Anonymous menuitem, set layout")
			obj.presentationType = "layout"
			return
		elif obj.role == roles.TABCONTROL:
			# to simplify menubar exploration (enforcing)
			debugLog("Role tabcontrol, set layout")
			obj.presentationType = "layout"
			return
		elif isSubtab(obj):
			debugLog("Subtab, set content")
			obj.presentationType = "content"
			return
		elif obj in self.expandedMenu:
			# for expanded menu
			debugLog("ExpandedMenu, set layout")
			obj.presentationType = "layout"
			return
		elif obj.role == roles.POPUPMENU and not obj.states:
			# to select this as simpleParent when closing submenu
			debugLog("Role popupmenu w/ states, set content")
			obj.presentationType = "content"
			return
		elif hasattr(obj, "UIAElement") and obj.UIAElement.cachedClassName in ("NetUIRepeatButton", "NetUIScrollBar", "NetUIAppFrameHelper"):
			# to hide scrolling and window-action buttons
			debugLog("CachedClassName %s, set unavailable"%obj.UIAElement.cachedClassName)
			obj.presentationType = "unavailable"
			return
		elif not obj.name:
			# generic
			debugLog("Anonymous obj, set layout")
			obj.presentationType = "layout"
			return
		elif obj.presentationType == "unavailable":
			# for menu items not currently available,
			# but which we want to show to users
			debugLog("PresType unavailable, set content")
			obj.presentationType = "content"
			return
		elif obj.role == roles.DATAGRID and obj.description:
			# to hide in grouping (it should be a grid associated to a visible button)
			debugLog("Role datagrid has description, set unavailable")
			obj.presentationType = "unavailable"
			return
		elif obj.role == roles.DATAGRID and not obj.description:
			# to show in submenu
			if allObjPassCheck(lambda i: i.role == roles.GROUPING and i.presentationType == "content", obj.children):
				obj.presentationType = "layout"
				return
			else:
				# ...manage other cases
				debugLog("...set content")
				obj.presentationType = "content"
				return
		elif obj.role == roles.LIST:
			# to explore children only
			debugLog("Role list, set layout")
			obj.presentationType = "layout"
		elif obj.role == roles.GROUPING:
			debugLog("Role grouping, set content")
			obj.presentationType = "content"
		elif obj.role in (roles.GRAPHIC, roles.STATICTEXT):
			# rare and useless, hide
			debugLog("Role %s, set unavailable"%obj.role)
			obj.presentationType = "unavailable"

	def event_focusEntered(self, obj, nextHandler):
		if not isOffice(obj):
			# don't process, for now
			nextHandler()
			return
		elif not self.exploring and isRibbonRoot(obj):
			debugLog("Exploration starts")
			obj.presentationType = "content"
			self.explorationStart()
			return
		elif not self.exploring:
			# stop immediately
			nextHandler()
			return
		debugLog("Running %s for obj %s,%s"%(inspect.currentframe().f_code.co_name,obj.name,obj.role))
		if obj.role == roles.MENUITEM and isRibbonRoot(obj.parent):
			# in v1
			obj.presentationType = "layout"
			return
		elif obj.role == roles.TABCONTROL:
			debugLog("Mute %s"%obj.role)
			obj.presentationType = "layout"
			self.subtab = obj
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
		self.userObj = api.getFocusObject()
		self.userObjHasFocus = True
		# simple gestures
		for gesture in ("tab", "escape", "enter", "downArrow", "leftArrow", "rightArrow", "upArrow"):
			self.bindGesture("kb:%s"%gesture, gesture)
		self.bindGesture("kb:shift+tab", "shiftTab")
		self.bindGesture("kb:alt+upArrow", "altUpArrow")
		self.bindGesture("kb:alt+downArrow", "altDownArrow")
		# for debugging
		self.bindGesture("kb:NVDA+space", "toggleExploration")

	def explorationEnd(self):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		self.exploring = False
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
			ui.message(_("Action not available here"))
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
			ui.message(_("Action not available here"))

	def script_altDownArrow(self, gesture):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		if not self.userObjHasFocus or states.UNAVAILABLE in self.userObj.states:
			ui.message(_("Action not available here"))
		elif isRibbonRoot(self.userObj.simpleParent):
			self.expandMenu(self.userObj)
		elif states.COLLAPSED in self.userObj.states:
			self.expandSubmenu(self.userObj)

	def script_toggleExploration(self, gesture):
		debugLog("Running %s"%inspect.currentframe().f_code.co_name)
		if self.exploring:
			self.explorationEnd()
		else:
			self.explorationStart()

	def script_objInfo(self, gesture):
		obj = api.getNavigatorObject()
		prop = obj.presentationType
		parObj = obj.simpleParent
		ui.message(''.join([prop, parObj.name, str(parObj.role)]))

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
		else:
			try:
				obj.setFocus()
				self.userObjHasFocus = True
			except:
				debugLog("SetFocus failed, try forcing")
				self.forceFocus(obj)

	def forceFocus(self, obj):
		# workaround: find previous focusable obj,
		# focus it, then simulate a tab
		# to get focus on obj we want
		# (not applicable to offscreen obj)
		if states.OFFSCREEN in obj.states:
			prevObj = None
		else:
			prevObj = findFocusablePrevious(obj)
		if prevObj:
			debugLog("Found prevObj %s"%prevObj.name)
			prevObj.setFocus()
			InputGesture.fromName("tab").send()
			# remark after prevObj.setFocus()
			self.userObj = api.getFocusObject()
			self.userObjHasFocus = True
		else:
			# guarantee an output, userObj vars should be untouched
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
		groupMenu = self.expandedMenu[-1]
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
		groupMenu = self.expandedSubmenu[-1]
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
			debugLog("Submenu condition!")
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

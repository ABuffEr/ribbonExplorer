# -*- coding: UTF-8 -*-
# Author: Alberto Buffolino
# License: GPLv3
import controlTypes as ct
import globalVars
import inspect
from logHandler import log
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
# to enable logging
DEBUG = False

def debugLog(message):
	if DEBUG:
		log.info(message)

def isOfficeApp(obj):
	try:
		# check from nvda\source\UIAHandler\__init__.py
		if obj.appModule.productName.startswith(("Microsoft Office", "Microsoft Outlook")):
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

def findFocusableNext(obj):
	res = obj.simpleNext
	if not res:
		res = obj.simpleParent.simpleNext
	if not res:
		res = obj.simpleParent.simpleFirstChild
	if res:
		if res.isFocusable:
			return res
		elif res.childCount:
			return res.simpleFirstChild

def moveFocusTo(obj):
	if not obj:
		return False
	debugLog("Try focus on obj %s"%obj.name)
	try:
		obj.setFocus()
		return True
	except:
		debugLog("Failed setFocus() on %s"%obj.name)
	return False

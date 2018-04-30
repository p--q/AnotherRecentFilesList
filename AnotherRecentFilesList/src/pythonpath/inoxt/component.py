#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import uno
import unohelper

from os import sep
if sep == '\\':
	from urllib.parse import unquote
	
	
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.util import XStringWidth
from com.sun.star.awt import XMenuListener
from com.sun.star.frame import XPopupMenuController, XDispatchProvider, XStatusListener, XDispatch
from com.sun.star.lang import XInitialization, XServiceInfo
from com.sun.star.container import XContainerListener
from com.sun.star.util import XStringAbbreviation
from com.sun.star.util import URL  # Struct
PROTOCOL = 'mytools.frame:'
# Menu_Path = 'ContextSpecificRecentFileList'
# IMPL_NAME = 'mytools.frame.ContextSpecificRecentFileList'
# SERVICE_NAME = 'com.sun.star.frame.PopupMenuController'

Node_History = '/org.openoffice.Office.Histories/Histories'
Node_Common_History = '/org.openoffice.Office.Common/History'

Mod_StartModule = 'com.sun.star.frame.StartModule'
Mod_BasicIDE = 'com.sun.star.script.BasicIDE'
# Mod_Chart2 = 'com.sun.star.chart2.ChartDocument'
Mod_Global = 'com.sun.star.text.GlobalDocument'
Mod_Text = 'com.sun.star.text.TextDocument'
Mod_Database = 'com.sun.star.sdb.OfficeDatabaseDocument'
# Mod_Spreadsheet = 'com.sun.star.SpreadsheetDocument'
#Mod_Formular = "com.sun.star.formula.FormularProperties"
#Mod_Formula = "com.sun.star.formula.FormulaProperties"

# Mod_sdb_prefix = 'com.sun.star.sdb'


IMPLE_NAME = None
SERVICE_NAME = None
def create(ctx, *args, imple_name, service_name):
	global IMPLE_NAME
	global SERVICE_NAME
	if IMPLE_NAME is None:
		IMPLE_NAME = imple_name 
	if SERVICE_NAME is None:
		SERVICE_NAME = service_name
	return AnotherRecentFilesPopupMenuController(ctx, *args)
class AnotherRecentFilesPopupMenuController(unohelper.Base, XPopupMenuController, XDispatchProvider, XMenuListener, XContainerListener, XServiceInfo):
	def __init__(self, ctx, *propertyvalues):  # argsはPropertyValueのタプル。
		self.ctx = ctx
		self.frame = None # frame of the document
		self.modname = "" # module name
		self.command = ""
		self.list_changed = False
		self.file_list = []
		self.menu = None
		self.history_list = None
		if propertyvalues:
			self.initialize(propertyvalues)
		if self.frame:
			self.frame.addEventListener(self)
	# XInitialization
	def initialize(self, propertyvalues):
		for propertyvalue in propertyvalues:
			name, value = propertyvalue.Name, propertyvalue.Value
			if name=='Frame':
				self.frame = value
			elif name=='ModuleIdentifier':
				if value.startswith('com.sun.star.sdb'):
					self.modname = Mod_Database
				elif value=='com.sun.star.chart2.ChartDocument':
					self.modname = 'com.sun.star.SpreadsheetDocument'
				else:
					self.modname = value
			elif name=='CommandURL':
				self.command = value
	# XServiceInfo
	def getImplementationName(self):
		return IMPLE_NAME
	def supportsService(self, servicename):
		return servicename==SERVICE_NAME
	def getSupportedServiceNames(self):
		return (SERVICE_NAME,)		
	# XDispatchProvider コマンドURLを受け取ってXDispatchを備えたオブジェクトを返す。今回は自身を返している。
	def queryDispatch(self, url, targetframename, searchflags):
		if url.Protocol == PROTOCOL:
			if url.Path == 'ContextSpecificRecentFileList':
				return self
		return None
	def queryDispatches(self, requests):
		return tuple(self.queryDispatch(i.FeatureURL, i.FrameName, i.SearchFlags) for i in requests)
	# XContainerListener
	def elementInserted(self, ev):
		self.list_changed = True
		self.unregister_listener()
	
	def elementRemoved(self, ev):
		self.list_changed = True
		self.unregister_listener()
	
	def elementReplaced(self, ev):
		self.list_changed = True
		self.unregister_listener()
	
	# XStatusListener
	def statusChanged(self, state):
		"""This menu is always enabled."""
		pass
	
	def disposing(self, ev):
		if ev.Source == self.frame:
			#print "pmc: disposing"
			try:
				self.unregister_listener()
				##self.menu = None # crash
				self.ctx = None
				self.frame = None
				self.file_list = []
				self.history_list = None
			except Exception as e:
				print(e)
	
	# XPopupMenuController
	def setPopupMenu(self, menu):
		"""Set content of the popup menu passed by the factory."""
		if not self.frame: return
		if not menu: return
		
		self.menu = menu # keep the menu
		try:
			self.fill_menu()
		except Exception as e:
			print(e)
			return
		# add menu listener
		if self.menu:
			self.menu.addMenuListener(self)
		# adds container listener
		#self.register_listener()
	
	
	def updatePopupMenu(self):
		"""updatePopupMenu call."""
		#print "update pm"
		try:
			if self.list_changed:
				self.clear_menu()
				self.fill_menu()
				self.register_listener()
			self.list_changed = False
		except Exception as e:
			print(e)
	
	
	def _get_pick_list_size(self):
		reader = get_configreader(self.ctx, Node_Common_History)
		return reader.getByName('PickListSize')
	
	def __get_history_reader(self):
		"""Get ConfigurationAccess of the History nodepath."""
		return get_configreader(self.ctx, Node_History)
	
	# if the listener added to the "List" container, 
	# inserted event called "Size" times.
	def register_listener(self):
		"""Adds listener to the list."""
		reader = self.__get_history_reader()
		
		hist_list = reader.getPropertyValue('PickList')
		hist_list.addContainerListener(self)
		self.history_list = hist_list
	
	def unregister_listener(self):
		"""Remove listener from the list."""
		if self.history_list:
			self.history_list.removeContainerListener(self)
	
	def clear_menu(self):
		"""Remove all items from the menu."""
		self.menu.removeItem(0,self.menu.getItemCount())
	
	def fill_menu(self):
		"""Fill menu with entries."""
		self.file_list = []
		
		reader = self.__get_history_reader()
		
		# create history list according to the module name
		if self.modname in (Mod_StartModule,Mod_BasicIDE,Mod_Database):
			#self.create_general_history(reader)
			self.file_list = create_general_history(reader)
		else:
			n = self._get_pick_list_size()
			#print(n)
			self.file_list = create_context_spacific_history(
					self.ctx, reader, self.modname, n)
		
		if not self.file_list:
			self.menu.insertItem(1, '~No Documents.',0,1)
			self.menu.enableItem(1, False)
			return
		
		
		ua = self.ctx.ServiceManager.createInstanceWithContext(
			'com.sun.star.util.UriAbbreviation', self.ctx)
		sw = string_width()
		
		urlStr = 'URL'
# 		entries = []
		#print"..."
		try:
			if sep == "\\":

				for i, v in enumerate(self.file_list):

					if v[urlStr].startswith('file:///'):
						syspath = self.abbreviation(str(unquote(v[urlStr].encode('ascii')),'utf8')[8:].replace('/','\\'), 46, '\\')
					else:

						syspath = ua.abbreviateString(sw,46,v[urlStr])
					label = '~%s: %s' % (i+1, syspath)

					self.menu.insertItem(i+1,label,0,i)
					self.menu.setTipHelpText(i+1, v[urlStr])

			else:
				for i, v in enumerate(self.file_list):
					if v[urlStr].startswith('file:///'):
						syspath = uno.fileUrlToSystemPath(ua.abbreviateString(sw, 46, v[urlStr]))
					else:
						syspath = ua.abbreviateString(sw,46,v[urlStr])
					label = '~%s: %s' % (i+1, syspath)

					self.menu.insertItem(i+1, label, 0, i)
					self.menu.setTipHelpText(i+1, v[urlStr])
		except Exception as e:
			print(e)
	
	# static method
	def abbreviation(url, length, pathsep="/"):
		"""Abbreviate file path."""
		if len(url) <= length: return url
		
		parts = url.split(pathsep)
		
		while len(parts) > 3:
			if len(parts) <= 3:
				return pathsep.join(parts)
			
			l = len(parts) / 2
			del parts[l]
			if sum([len(p) for p in parts]) <= length:
				l = len(parts) / 2
				parts.insert(l+1,"...")
				return pathsep.join(parts)
		
		if len(parts) == 3 and sum([len(p) for p in parts]) > length:
			parts[1] = ''.join((parts[1][0:7], '...'))
		
		return pathsep.join(parts)
	
	abbreviation = staticmethod(abbreviation)
	

	# XMenuListener
	def itemHighlighted(self, ev):
		pass
	def itemActivated(self, ev):
		pass
	def itemDeactivated(self, ev):
		pass
	def itemSelected(self, ev):
		menu_id = ev.MenuId
		if menu_id <= 0: return
		try:
			if self.file_list:# and len(self.file_list) < menu_id -1:
				self.open_file( self.file_list[menu_id -1] )
		except Exception as e:
			print(e)
	
	def open_file(self,entry):
		"""Open file with dispatch."""
		if not self.frame: return
		url = URL()
		url.Complete = '.uno:Open'
		#print entry["URL"],entry["Filter"]
		transformer = self.ctx.ServiceManager.createInstanceWithContext('com.sun.star.util.URLTransformer', self.ctx)
		dummy, url = transformer.parseStrict(url)
		
		arg1 = create_PropertyValue('Referer', 'private:user')
		arg2 = create_PropertyValue('AsTemplate',False)
		arg3 = create_PropertyValue('FilterName',entry['Filter'])
		arg4 = create_PropertyValue('SynchronMode',False)
		arg5 = create_PropertyValue('URL',entry['URL'])
		arg6 = create_PropertyValue('FrameName','_default')
		args = (arg1,arg2,arg3,arg4,arg5,arg6)
		
		desktop = self.ctx.ServiceManager.createInstanceWithContext(
			'com.sun.star.frame.Desktop', self.ctx)
		
		disp = desktop.queryDispatch(url,'_self',0)
		
		if disp:
			disp.dispatch(url,args)


class string_width(unohelper.Base, XStringWidth):
	def queryStringWidth(self,string):
		return len(string)


def create_general_history(reader):
	"""History list as normal list."""
	pk_list = reader.getPropertyValue('PickList')
	file_list = []
	
	filterName = 'Filter'
	urlStr = 'URL'
	
	if pk_list.hasElements():
		for name in pk_list.getElementNames():
			element = {}
			p = pk_list.getByName(name)
			
			element[urlStr] = p.getPropertyValue(urlStr)
			element[filterName] = p.getPropertyValue(filterName)
			
			file_list.append(element)
	return file_list


def create_context_spacific_history(ctx, reader, modname, pick_size):
	"""Context specific history."""
	# make a module specific list
	iFlag = 0x1
	eFlag = 0x8 + 0x1000 + 0x40000
	
	filter_list = get_filter_list(ctx, modname, iFlag, eFlag)
	
	if modname not in filter_list:
		return create_general_history(reader)	# not found
	
	file_list = []
	
	# module specific filter names
	if modname == Mod_Global:
		# GlobalDocument: Global + Text
		Global_filters = set(filter_list.get(Mod_Global,[]))
		Text_filters = set(filter_list.get(Mod_Text,[]))
		mod_filters = Global_filters | Text_filters
	else:
		mod_filters = set(filter_list[modname])
	
	filterName = 'Filter'
	urlStr = 'URL'
	
	# to check file exists without from url to path conversion
	#sfa = ctx.ServiceManager.createInstanceWithContext(
	#	u'com.sun.star.ucb.SimpleFileAccess', ctx)
	#file_exists = sfa.exists
	
	# "List" does not list stored file. So, get files form "PickList" and 
	# complement remains from "List".
	
	# PickList keeps Recent Files entries.
# 	pk_list = reader.getByName('URLHistory')#('PickList')
	
	pk_list = reader.getByName('PickList')
	
	
# 	pk_size = pick_size#reader.getPropertyValue(u'PickListSize')
	pk_items = pk_list.getByName('ItemList')
	pk_order = pk_list.getByName('OrderList')
	
	# convert order list to real number and sort it
	order_list = [int(i) for i in pk_order.getElementNames()]
	order_list.sort()
	
	if pk_list.hasElements():
		#for name in pk_list.getElementNames(): # not sorted in the ElementNames
		for i in order_list:
			#print(i)
			if pk_order.hasByName(str(i)):
				name = pk_order.getByName(str(i)).HistoryItemRef
				#print(name)
				if pk_items.hasByName(name):
					pk = pk_items.getByName(name)
					fl = pk.getPropertyValue(filterName)
					#print(fl)
					if fl in mod_filters:
						element = {}
						element[urlStr] = name#pk.getPropertyValue(urlStr) #url
						element[filterName] = fl
						file_list.append(element)

	return file_list



def get_filter_info(descs):
	docService = 'DocumentService'
	filterName = 'Name'
	service = ""
	name = ""
	
	for desc in descs:
		if desc.Name == docService:
			service = desc.Value
		elif desc.Name == filterName:
			name = desc.Value
	return service,name


def get_filter_list(ctx, mod="", iFlag=0, eFlag=0):
	"""Get filter list from FilterFactory.
	
	Returned filters are categorized in their module.
	filer_list["com.sun.star.sheet.SpreadsheetDocument": ["calc8", "..."], "": ... ]
	"""
	ff = ctx.ServiceManager.createInstanceWithContext(
		'com.sun.star.document.FilterFactory' ,ctx)
	que = 'getSortedFilterList():module=%s:iflags=%s:eflags=%s' % (mod,iFlag,eFlag)
	filters = ff.createSubSetEnumerationByQuery(que)
	
	filter_list = {}
	
	get_info = get_filter_info
	
	# categorizing
	while filters.hasMoreElements():
		fl = filters.nextElement()
		service,name = get_info(fl)
		if service in filter_list:
			filter_list[service].append(name)
		else:
			filter_list[service] = [name]
	
	return filter_list


def create_PropertyValue(name,value):
	p = PropertyValue()
	p.Name = name
	p.Value = value
	return p


def get_configreader(ctx,node):
	"""Get specified configuration reader of the nodepath."""
	# just only read access
	try:
		cp = ctx.ServiceManager.createInstanceWithContext(
			'com.sun.star.configuration.ConfigurationProvider', ctx)
		props = PropertyValue()
		props.Name = 'nodepath'
		props.Value = node
		cra = cp.createInstanceWithArguments(
			'com.sun.star.configuration.ConfigurationAccess', (props,))
	except:
		return None
	return cra


# g_ImplementationHelper = unohelper.ImplementationHelper()
# g_ImplementationHelper.addImplementation(
# 	AnotherRecentFilesPopupMenuController,
# 	IMPL_NAME,
# 	(SERVICE_NAME,),)



#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper
# import re, os
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.lang import XInitialization, XServiceInfo
from com.sun.star.awt import XMenuListener
from com.sun.star.container import XContainerListener
from com.sun.star.frame import XPopupMenuController, XDispatchProvider, XStatusListener, XDispatchProvider
from com.sun.star.util import URL  # Struct
# from com.sun.star.awt import XContainerWindowEventHandler
# from com.sun.star.uno.TypeClass import ENUM, TYPEDEF, STRUCT, EXCEPTION, INTERFACE, CONSTANTS  # enum
# from pq import XTcu  # 拡張機能で定義したインターフェイスをインポート。
# from .optiondialog import dilaogHandler
# from .wsgi import Wsgi, createHTMLfile
# from .wcompare import wCompare
IMPLE_NAME = None
SERVICE_NAME = None
HANDLED_PROTOCOL = 'mytools.frame:'
def create(ctx, *args, imple_name, service_name):
	global IMPLE_NAME
	global SERVICE_NAME
	if IMPLE_NAME is None:
		IMPLE_NAME = imple_name 
	if SERVICE_NAME is None:
		SERVICE_NAME = service_name
	return AnotherRecentFilesPopupMenuController(ctx, *args)
class AnotherRecentFilesPopupMenuController(unohelper.Base, XPopupMenuController, XInitialization, XDispatchProvider, XServiceInfo): 	 
	def __init__(self, ctx, *args):  #argsはPropertyValue|。	NameアトリビュートはFrame, CommandURL, ModuleName
		self.ctx = ctx
		self.frame = None # frame of the document
		self.modname = "" # module name
		self.command = ""
		self.menu = None
		self.file_list = []
		self.list_changed = False
		args and self.initialize(args)
# 		if self.frame:
# 			self.frame.addEventListener(self)			
	# XServiceInfo
	def getImplementationName(self):
		return IMPLE_NAME
	def supportsService(self, name):
		return name == SERVICE_NAME
	def getSupportedServiceNames(self):
		return (SERVICE_NAME,)		
	# XInitialization
	def initialize(self, args):
		for arg in args:
			if arg.Name == 'Frame':
				self.frame = arg.Value
			elif arg.Name == 'ModuleName':
				if arg.Value.startswith('com.sun.star.sdb'):
					self.modname = 'com.sun.star.sdb.OfficeDatabaseDocument'
				elif arg.Value == 'com.sun.star.chart2.ChartDocument':
					self.modname = 'com.sun.star.SpreadsheetDocument'
				else:
					self.modname = arg.Value
			elif arg.Name == 'CommandURL':
				self.command = arg.Value	
	# XDispatchProviderの実装 コマンドURLを受け取ってXDispatchを備えたオブジェクトを返す。今回は自身を返している。
	def queryDispatch(self, url, targetframename, searchflags): 
		if url.Protocol==HANDLED_PROTOCOL:
			if url.Path in ['ContextSpecificRecentFileList']:
				return self
		return None
	def queryDispatches(self, requests):
		return tuple(self.queryDispatch(request.FeatureURL, request.FrameName, request.SearchFlags) for request in requests)   
	# XPopupMenuController
	def setPopupMenu(self, popupmenu):
		if self.frame and popupmenu:
			self.menu = popupmenu # keep the menu
			fill_menu()
			self.menu and self.menu.addMenuListener(MenuListener(self.ctx, self.frame, self.file_list))
	def updatePopupMenu(self):
		if self.list_changed:
			self.menu.removeItem(0, self.menu.getItemCount())
			fill_menu()
			
			
			self.register_listener()
		self.list_changed = False		
def fill_menu():
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
	entries = []
	#print"..."
	try:
		if sep == "\\":
		
			for i, v in enumerate(self.file_list):
			
				if v[urlStr].startswith(u'file:///'):
					syspath = self.abbreviation(unicode(unquote(v[urlStr].encode('ascii')),'utf8')[8:].replace('/','\\'), 46, '\\')
				else:
				
					syspath = ua.abbreviateString(sw,46,v[urlStr])
				label = u'~%s: %s' % (i+1, syspath)
			
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
				
				
				
class ContainerListener(unohelper.Base, XContainerListener):
	def __init__(self): 
		pass
	def elementInserted(self, containerevent):
		self.list_changed = True
		self.unregister_listener()
	def elementRemoved(self, containerevent):
		self.list_changed = True
		self.unregister_listener()
	def elementReplaced(self, containerevent):
		self.list_changed = True
		self.unregister_listener()	
	def disposing(self, eventobject):
		eventobject.Source.removeMenuListener(self)	
class MenuListener(unohelper.Base, XMenuListener):
	def __init__(self, ctx, frame, file_list): 
		self.file_list = file_list
		self.frame = frame
		self.ctx = ctx
	def itemHighlighted(self, menuevent):
		pass
	def itemSelected(self, menuevent):
		menu_id = menuevent.MenuId
		if menu_id>0 and self.file_list and self.frame:
			open_file(self.ctx, self.file_list[menu_id-1])
	def itemActivated(self, menuevent):
		pass
	def itemDeactivated(self, menuevent):
		pass   
	def disposing(self, eventobject):
		eventobject.Source.removeMenuListener(self)	
class StatusListener(unohelper.Base, XStatusListener):
	def statusChanged(self, state):
		pass	
	def disposing(self, eventobject):
		eventobject.Source.removeMenuListener(self)		
def open_file(ctx, entry):
	url = URL(Complete='.uno:Open')
	transformer = ctx.ServiceManager.createInstanceWithContext('com.sun.star.util.URLTransformer', ctx)
	dummy, url = transformer.parseStrict(url)
	desktop = ctx.getByName('/singletons/com.sun.star.frame.theDesktop') 
	dispatch = desktop.queryDispatch(url, '_self', 0)
	if dispatch:
		args = PropertyValue(Name='Referer', Value='private:user'),\
			PropertyValue(Name='AsTemplate', Value=False),\
			PropertyValue(Name='FilterName', Value=entry['Filter']),\
			PropertyValue(Name='SynchronMode', Value=False),\
			PropertyValue(Name='URL', Value=entry['URL']),\
			PropertyValue(Name='FrameName', Value='_default')
		dispatch.dispatch(url, args)	
	
	
	


# def getConfigs(consts):
# 	ctx, smgr, configurationprovider, css, properties, nodepath, simplefileaccess = consts
# 	fns_keys = "SERVICE", "INTERFACE", "PROPERTY", "INTERFACE_METHOD", "INTERFACE_ATTRIBUTE", "NOLINK"  # fnsのキーのタプル。
# 	node = PropertyValue(Name="nodepath", Value="{}OptionDialog".format(nodepath))
# 	root = configurationprovider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
# 	offline, refurl, refdir, idlstext = root.getPropertyValues(properties)  # コンポーネントデータノードから値を取得する。		
# 	prefix = "https://{}".format(refurl)
# 	if offline:  # ローカルリファレンスを使うときはprefixを置換する。
# 		pathsubstservice = smgr.createInstanceWithContext("com.sun.star.comp.framework.PathSubstitution", ctx)
# 		fileurl = pathsubstservice.substituteVariables(refdir, True)  # $(inst)を変換する。fileurlが返ってくる。
# 		if simplefileaccess.exists(fileurl):
# 			systempath = os.path.normpath(unohelper.fileUrlToSystemPath(fileurl))  # fileurlをシステムパスに変換して正規化する。
# 			prefix = "file://{}".format(systempath)
# 		else:
# 			raise RuntimeError("Local API Reference does not exists.")
# 	if not prefix.endswith("/"):
# 		prefix = "{}/".format(prefix)
# 	idls = "".join(idlstext.split()).split(",")  # xmlがフォーマットされていると空白やタブが入ってくるのでそれを除去してリストにする。
# 	idlsset = set("{}{}".format(css, i) if i.startswith(".") else i for i in idls)  # "com.sun.star"が略されていれば付ける。
# 	return ctx, configurationprovider, css, fns_keys, offline, prefix, idlsset	


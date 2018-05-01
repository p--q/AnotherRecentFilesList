#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import uno
import unohelper
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.util import XStringWidth
from com.sun.star.awt import XMenuListener
from com.sun.star.frame import XPopupMenuController, XStatusListener
from com.sun.star.lang import XServiceInfo
from com.sun.star.container import XContainerListener
from com.sun.star.util import XStringAbbreviation
from com.sun.star.util import URL  # Struct
from com.sun.star.frame import XDispatchProviderInterceptor

# PROTOCOL = 'mytools.frame:'
# Menu_Path = 'ContextSpecificRecentFileList'
# IMPL_NAME = 'mytools.frame.ContextSpecificRecentFileList'
# SERVICE_NAME = 'com.sun.star.frame.PopupMenuController'

# Node_History = '/org.openoffice.Office.Histories/Histories'
# Node_Common_History = '/org.openoffice.Office.Common/History'

# Mod_StartModule = 'com.sun.star.frame.StartModule'
# Mod_BasicIDE = 'com.sun.star.script.BasicIDE'
# Mod_Chart2 = 'com.sun.star.chart2.ChartDocument'
# Mod_Global = 'com.sun.star.text.GlobalDocument'
# Mod_Text = 'com.sun.star.text.TextDocument'
# Mod_Database = 'com.sun.star.sdb.OfficeDatabaseDocument'
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
class AnotherRecentFilesPopupMenuController(unohelper.Base, XPopupMenuController, XServiceInfo, XStatusListener):  # import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	def __init__(self, ctx, *args):  # argsはPropertyValueのタプルを受け取る。
		self.frame = None
		moduleidentifier = "" 	
		for propertyvalue in args:
			name, value = propertyvalue.Name, propertyvalue.Value
			if name=='Frame':
				self.frame = value
			elif name=='ModuleIdentifier':
				moduleidentifier = 'com.sun.star.SpreadsheetDocument' if value=='com.sun.star.chart2.ChartDocument' else value
		smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
		self.configreader = createConfigReader(ctx, smgr)  # 読み込み専用の関数を取得。
		self.filterlist = [] if moduleidentifier in ('com.sun.star.frame.StartModule', 'com.sun.star.script.BasicIDE') or moduleidentifier.startswith('com.sun.star.sdb') else getFilterList(self.configreader, moduleidentifier)
		self.uriabbreviation = smgr.createInstanceWithContext('com.sun.star.util.UriAbbreviation', ctx)					
		self.picklistchangeflg = []

					
					
# 		self.ctx = ctx
# 		self.list_changed = False
# 		self.file_list = []
# 		self.history_list = None
	# XServiceInfo
	def getImplementationName(self):
		return IMPLE_NAME
	def supportsService(self, servicename):
		return servicename==SERVICE_NAME
	def getSupportedServiceNames(self):
		return (SERVICE_NAME,)		
	# XStatusListener
	def statusChanged(self, state):  # メニュー項目のチェックボックスなどの把握のため?
		pass
	def disposing(self, eventobject):
		eventobject.Source.removeMenuListener(self)		
	
	# XPopupMenuController
	def setPopupMenu(self, popupmenu):  # ポップアップメニューを作成。引数はcom.sun.star.awt.PopupMenu。
		self._fillPopupMenu(popupmenu)	
		self.popupmenu = popupmenu
	def updatePopupMenu(self):
		if self.picklistchangeflg:
			self.popupmenu.clear()
			self._fillPopupMenu(self.popupmenu)
			self.picklistchangeflg.clear()

	
	

	

	
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
			
	def _fillPopupMenu(self, popupmenu):
		filterlist = self.filterlist
		itempos = 0
		if filterlist:
			uriabbreviation = self.uriabbreviation
			stringwidth = StringWidth()
			picklist = self.configreader('/org.openoffice.Office.Histories/Histories/PickList')
			itemlist, orderlist = picklist.getPropertyValues(("ItemList", "OrderList"))  # ItemListからTitleとFilterが取得できるが順序は保存されていない。順序はOrderListから取得する。
			for i in orderlist:  # oor:name="HistoryOrder"には古い順から番号が入っている。
				fileurl = orderlist[i]  # fileurlが返る。
				filtername = itemlist[fileurl]["Filter"]
				if filtername in filterlist:
					abbreviatefileurl = uriabbreviation.abbreviateString(stringwidth, 46, fileurl)  # 46文字に切り詰める。
					systempath = unohelper.fileUrlToSystemPath(abbreviatefileurl)
					popupmenu.insertItem(itempos+1, '~{}: {}'.format(itempos+1, systempath), 0, itempos)  # ItemIdは1から始まり区切り線は含まない。ItemPosは0から始まり区切り線を含む。
					popupmenu.setTipHelpText(itempos+1, systempath)
					itempos += 1
			picklist.addContainerListener(ContainerListener(self.picklistchangeflg))		
		if itempos:  
			popupmenu.addMenuListener(MenuListener())
		else:  # フィルターリストが取得できなかったときやポップアップメニューの項目がない時。	
			popupmenu.clear()  # 親メニューがグレイアウトするが>は消えない。	
			
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
			
			
					
class string_width(unohelper.Base, XStringWidth):
	def queryStringWidth(self, string):
		return len(string)
class ContainerListener(unohelper.Base, XContainerListener):
	def __init__(self, picklistchangeflg): 
		self.picklistchangeflg = picklistchangeflg
	def elementInserted(self, containerevent):
		self.picklistchangeflg.append(True)
		containerevent.Source.removeContainerListener(self)	
	def elementRemoved(self, containerevent):
		self.picklistchangeflg.append(True)
		containerevent.Source.removeContainerListener(self)	
	def elementReplaced(self, containerevent):
		self.picklistchangeflg.append(True)
		containerevent.Source.removeContainerListener(self)	
	def disposing(self, eventobject):
		eventobject.Source.removeContainerListener(self)	
class MenuListener(unohelper.Base, XMenuListener):
	def __init__(self, ctx, frame, file_list): 
		self.file_list = file_list
		self.frame = frame
		self.ctx = ctx
	def itemHighlighted(self, menuevent):
		pass
	def itemSelected(self, menuevent):
		menu_id = menuevent.MenuId
# 		if menu_id>0 and self.file_list and self.frame:
# 			open_file(self.ctx, self.file_list[menu_id-1])
	def itemActivated(self, menuevent):
		pass
	def itemDeactivated(self, menuevent):
		pass   
	def disposing(self, eventobject):
		eventobject.Source.removeMenuListener(self)	
# class StatusListener(unohelper.Base, XStatusListener):
# 	def statusChanged(self, state):
# 		pass	
# 	def disposing(self, eventobject):
# 		eventobject.Source.removeStatusListener(self)		
class StringWidth(unohelper.Base, XStringWidth):
	def queryStringWidth(self,string):
		return len(string)
def createConfigReader(ctx, smgr):  # ConfigurationProviderサービスのインスタンスを受け取る高階関数。
	configurationprovider = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)  # ConfigurationProviderの取得。
	def configReader(path):  # ConfigurationAccessサービスのインスタンスを返す関数。
		node = PropertyValue(Name="nodepath", Value=path)
		return configurationprovider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
	return configReader
def getFilterList(configreader, moduleidentifier):
	filterlist = []
	filters = configreader("/org.openoffice.TypeDetection.Filter/Filters")  # org.openoffice.TypeDetectionパンケージのTypesコンポーネントのTypesノードを根ノードにする。
	for filtername in filters:  # 各子ノード名について。
		filternode = filters[filtername]  # 子ノードを取得。
		if "DocumentService" in filternode:
			if filternode["DocumentService"]==moduleidentifier:
				filterlist.append(filtername)
	return filterlist		

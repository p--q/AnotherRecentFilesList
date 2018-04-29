#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import xml.etree.ElementTree as ET
from helper import Elem
from config import getConfig
import os
#  拡張機能のxcsファイルでノードを書き込み後にテストするときは、registrymodifications.xcuにあるノードの値を忘れずに消しておかないと、そちらの値が使われて、編集後のxcsファイルの動作確認ができません。
def createXcs(c):
	os.chdir(c["src_path"])  # srcフォルダに移動。
	filename =  "config.xcs"
	c["backup"](filename)  # すでにあるファイルをバックアップ	
	with open(filename, "w", encoding="utf-8") as f:  # config.xcsファイルの作成
		root = Elem("oor:component-schema", {"oor:name": "ExtensionData", "oor:package": c["ExtentionID"], "xmlns:oor": "http://openoffice.org/2001/registry", "xmlns:xs": "http://www.w3.org/2001/XMLSchema", "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance", "xml:lang": "en-US"})  # 根の要素を作成。
		# テンプレートノードの作成
		root.append(Elem("templates"))
		templates = root[-1]
		# ノードタイプでデフォルト値を設定しておく。
		templates.append(Elem("group", {"oor:name": "nodetypes"}))
		nodetypes = templates[-1]
		nodetypes.append(Elem("info"))
		nodetypes[-1].append(Elem("desc", text="Do you use local SDK reference?"))	
		nodetypes.append(Elem("prop", {"oor:name": "OffLine", "oor:type": "xs:boolean"}))
		nodetypes[-1].append(Elem("value", text="false"))	
		nodetypes.append(Elem("info"))
		nodetypes[-1].append(Elem("desc", text="URL to online API reference"))	
		nodetypes.append(Elem("prop", {"oor:name": "RefURL", "oor:type": "xs:string"}))
		nodetypes[-1].append(Elem("value", text="api.libreoffice.org/docs/idl/ref/"))
		nodetypes.append(Elem("info"))
		nodetypes[-1].append(Elem("desc", text="Fileurl to offline API reference"))	
		nodetypes.append(Elem("prop", {"oor:name": "RefDir", "oor:type": "xs:string"}))
		nodetypes[-1].append(Elem("value", text="$(inst)/sdk/docs/idl/ref/"))
		nodetypes.append(Elem("info"))
		nodetypes[-1].append(Elem("desc", text="Save IgnoredIDLs as text. 'com.sun.star' is omitted."))	
		nodetypes.append(Elem("prop", {"oor:name": "IgnoredIDLs", "oor:type": "xs:string"}))
		nodetypes[-1].append(Elem("value", text=".uno.XInterface,.lang.XTypeProvider,.lang.XServiceInfo,.uno.XWeak,.lang.XMain,.uno.XAggregation,.lang.XUnoTunnel,X\w*Listener"))		
		# コンポーネントノードの作成
		root.append(Elem("component"))
		root[-1].append(Elem("group", {"oor:name": "Leaves"}))
		root[-1][-1].append(Elem("group", {"oor:name": "XUnoTreeCommandSettings"}))
		xunotreecommandsettings = root[-1][-1][-1]
		# デフォルト値として書き換えないノード。
		xunotreecommandsettings.append(Elem("info"))
		xunotreecommandsettings[-1].append(Elem("desc", text="These nodes should not be rewritten with ConfigurationProvider for default."))	
		xunotreecommandsettings.append(Elem("node-ref", {"oor:name": "Defaults", "oor:node-type": "nodetypes"}))
		# 設定値として書き換えるノード。
		xunotreecommandsettings.append(Elem("info"))
		xunotreecommandsettings[-1].append(Elem("desc", text="Rewrite these nodes with ConfigurationProvider."))	
		xunotreecommandsettings.append(Elem("node-ref", {"oor:name": "OptionDialog", "oor:node-type": "nodetypes"}))
		tree = ET.ElementTree(root)  # 根要素からxml.etree.ElementTree.ElementTreeオブジェクトにする。
		tree.write(f.name, "utf-8", True)  # xml_declarationを有効にしてutf-8でファイルに出力する。   
	print("{} has been created.".format(filename))	
if __name__ == "__main__":
	createXcs(getConfig(False))  # バックアップをとるときはTrue

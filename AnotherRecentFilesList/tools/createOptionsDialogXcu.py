#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import os
# import xml.etree.ElementTree as ET
# from helper import Elem, ElemProp, ElemPropLoc
from config import getConfig
from helper import createElem
def createOptionsDialogXcu(c):  #Creation of OptionsDialog.xcu
	os.chdir(c["src_path"])  # srcフォルダに移動。
	filename =  "OptionsDialog.xcu"
	c["backup"](filename)  # すでにあるファイルをバックアップ
	with open(filename, "w", encoding="utf-8") as f:  # OptionsDialog.xcuファイルの作成
		
		
		
		attrib = {"oor:name": "OptionsDialog", "oor:package": "org.openoffice.Office", "xmlns:oor": "http://openoffice.org/2001/registry", "xmlns:xs": "http://www.w3.org/2001/XMLSchema", "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"}
		root = createElem("oor:component-data", attrib)  # 根の要素を作成。

		
# 		root = Elem("oor:component-data", {"oor:name": "OptionsDialog", "oor:package": "org.openoffice.Office", "xmlns:oor": "http://openoffice.org/2001/registry", "xmlns:xs": "http://www.w3.org/2001/XMLSchema", "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"})  # 根の要素を作成。
		
		#オプションページの作成。拡張機能IDのc["ExtentionID"]をNameに使った時はLeafを一つにしないと拡張機能マネージャーでフリーズする。
		root.append(Elem("node", {'oor:name': "Nodes"}))  # セットノードNodesを追加。
		leaf = ElemLeaf(c, {"Name": c["ExtentionID"], "Label": {"en-US": "TreeCommandforUNO"}})  # node-type=Leaf。小項目の設定。Idを拡張機能Idと同じc["ExtentionID"]にする。
		name = "{}.Node1".format(c["ExtentionID"])  # ロードマップコントロールに表示させる大項目のID。ユニークの必要があると考えるので拡張機能IDにくっつける。
		node = ElemNode(c, {"Name": name, "Label": {"en-US": "Extensions", "ja-JP": "拡張機能"}, "AllModules": "true"}, leaves=(leaf,))  # node-type=Node。大項目の設定。
		root[-1].append(node)  #  node-type=NodeをセットノードNodesに追加。
		
		tree = ET.ElementTree(root)  # 根要素からxml.etree.ElementTree.ElementTreeオブジェクトにする。
		tree.write(f.name, "utf-8", True)  # xml_declarationを有効にしてutf-8でファイルに出力する。   
	print("{} has been created.".format(filename))	

def createLeaf(c, attrs):  # node-type=Leafのノードを返す。オプションページのロードマップコントロールの小項目になる。辞書attrのkeyは"Name", "Label", "GroupId", "GroupIndex", "Id"。 valueはすべて文字列でいれる。
	if not "Id" in attrs:  # IdがなければNameを代用する。Idを拡張機能Idと同じにすると拡張機能マネージャーでオプションボタンが表示される。
		attrs["Id"] = attrs["Name"]		
	attrs["OptionsPage"] = "%origin%/dialogs/optionsdialog.xdl"  # オプションダイアログxdlファイル名。UnoControlDialogオブジェクトになる。
	attrs["EventHandlerService"] = c["ExtentionID"]  # オプションダイアログの呼び出しと共に呼び出す実装サービス名。
	elem = createElem("node", {'oor:name': attrs.pop("Name"), "oor:op": "fuse"})
	for key, val in attrs.items():
		if key=="Label":
			elem.append(ElemPropLoc(key, val))
		else:
			elem.append(ElemProp(key, val))	


# class ElemLeaf(Elem):  # node-type=Leaf。オプションページのロードマップコントロールの小項目になる。
# 	def __init__(self, c, attrs):
# 		if not "Id" in attrs:  # IdがなければNameを代用する。Idを拡張機能Idと同じにすると拡張機能マネージャーでオプションボタンが表示される。
# 			attrs["Id"] = attrs["Name"]		
# 		attrs["OptionsPage"] = "%origin%/dialogs/optionsdialog.xdl"  # オプションダイアログxdlファイル名。UnoControlDialogオブジェクトになる。
# 		attrs["EventHandlerService"] = c["ExtentionID"]  # オプションダイアログの呼び出しと共に呼び出す実装サービス名。
# 		super().__init__("node", {'oor:name': attrs.pop("Name"), "oor:op": "fuse"})
# 		for key, val in attrs.items():
# 			if key=="Label":
# 				self.append(ElemPropLoc(key, val))
# 			else:
# 				self.append(ElemProp(key, val))
class ElemNode(Elem):  # node-type=Node
	def __init__(self, c, attrs, *, leaves=None):  # 辞書attrのkeyはName, Label, OptionsPage, AllModules, GroupId, GroupIndex。 valueはすべて文字列でいれる。
		super().__init__("node", {'oor:name': attrs.pop("Name"), "oor:op": "fuse"})
		for key, val in attrs.items():
			if key=="Label":
				self.append(ElemPropLoc(key, val))
			else:
				self.append(ElemProp(key, val))		
		if leaves is not None:
			self.append(Elem("node", {'oor:name': "Leaves"}))  # セットノードLeaves
			for leaf in leaves:
				self[-1].append(leaf)  #  node-type=LeafをセットノードLeavesに追加。						

if __name__ == "__main__":
# 	createOptionsDialogXcu(getConfig(False))  # バックアップをとるときはTrue
	createOptionsDialogXcu(getConfig(True))  # バックアップをとるときはTrue
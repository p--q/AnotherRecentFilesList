#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
# import xml.etree.ElementTree as ET
# class Elem(ET.Element):  # xml.etree.ElementTree.Element派生クラス。
# 	def __init__(self, tag, attrib={},  **kwargs):  # ET.Elementのアトリビュートのtextとtailはkwargsで渡す。
# 		txt = kwargs.pop("text", None)
# 		tail = kwargs.pop("tail", None)
# 		super().__init__(tag, attrib, **kwargs)
# 		if txt:
# 			self.text = txt
# 		if tail:
# 			self.tail = tail
# class ElemProp(Elem):  # 値を持つpropノード
# 	def __init__(self, name, txt):
# 		super().__init__("prop", {'oor:name': name}) 
# 		self.append(Elem("value", text=txt))
# class ElemPropLoc(Elem):  # 多言語化された値を持つpropノード。
# 	def __init__(self, name, langs):
# 		super().__init__("prop", {'oor:name': name})
# 		for lang, value in langs.items():
# 			self.append(Elem("value", {"xml:lang": lang}, text=value))

			
from xml.etree.ElementTree import Element			
def createElem(tag, attrib={},  **kwargs):  # ET.Elementのアトリビュートのtextとtailはkwargsで渡す。		
	txt = kwargs.pop("text", None)
	tail = kwargs.pop("tail", None)
	sub = kwargs.pop("sub", None)  # サブノードにするET.Elementを取得。
	subs = kwargs.pop("subs", None)  # サブノードにするET.Elementのタプルを取得。
	elem = Element(tag, attrib, **kwargs)
	if txt is not None:
		elem.text = txt
	if tail is not None:
		elem.tail = tail	
	if sub is not None:  # ET.Elementが入っていてもsubだけだとFalseになる。
		elem.append(sub)
	if subs is not None:
		elem.extend(subs)	
	return elem	 # Elementオブジェクトを返す。

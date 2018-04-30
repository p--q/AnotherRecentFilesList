#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
from config import getConfig
import os
import glob	
def defineIDLs():  # IDLの定義を設定する。継承しているUNOIDLのidlファイルは自動include。

	# サービスの定義。new-styleサービスはひとつのインターフェイスしか継承しない。
	service = UNOIDL("pq.Tcu")  # 定義するUNOIDLのフルパスでインスタンス化。
	service.setSuper("XTcu")  # 継承するUNOIDLを定義。同じパスのものはフルパスで書く必要はない。idlファイルは自動include。
	yield service

	# インターフェイスの定義。多重継承するインターフェイスにはメソッドをつけない(慣習?)。
	interface = UNOIDL("pq.XTcu")  # 定義するUNOIDLのフルパスでインスタンス化。XInterfaceは自動include。
	interface.setSubs(  # 属性、メソッド、多重継承のインターフェイスなど
		"sequence <string> treelines([in] any Object)",		
		"sequence <string> wtreelines([in] any Object)",		
		"sequence <string> wcomparelines([in] any Object1, [in] any Object2)",		
		"void wtree([in] any Object)",
		"void wcompare([in] any Object1, [in] any Object2)"
		)
	yield interface

class UNOIDL:
	def __init__(self, name):  # UNOIDLのフルパスを引数にする。
		self.name = name  # UNOIDLのフルパス。
		self.includes = []  # インクルードするidlファイル。
		self.subs = tuple()  # 属性、メソッドなど。
		self.super = ""  # 継承するUNOIDLのフルパス。
	def setIncludes(self, *args):  # includeするidlファイルを取得。継承するUNOIDLは自動取得。
		self._args(args, "includes")
	def setSubs(self, *args):  # 定義したUNOIDLがもつ属性、メソッドなどを取得。 
		for arg in args:
			if arg.startswith("interface"):  # インターフェイスのときはidlファイルをインクルードする。
				self.includes.append(arg.replace("interface", "").strip())
		self._args(args, "subs")	 
	def setSuper(self, super):  # 定義したUNOIDLが継承するUNOIDLを取得。
		self.super = super
	def _args(self, args, attr):  # 長さ1以上のタプルをプロパティに取得。
		if len(args) > 0:
			setattr(self, attr, args)  
	def getVal(self):  # アンダースコア名、モジュールのリスト、出力するidlファイル名を返す。
		tab = "	"  # インデント
		name_underscore = "_" + self.name.replace(".", "_") + "_idl_"  # アンダースコア名
		ms = self.name.split(".")  # UNOIDLのパスを分割。
		name_base = ms.pop()  # パスのないUNOIDL名を取得。
		interface = name_base.startswith("X")  # インターフェイスであればTrue
		ms = list(map(lambda x: "module " + x, ms))  # モジュール名を付ける。
		s = "interface " if interface else "service "  # UNOIDLの接頭語。
		ms.append("\n" + tab + s + name_base)
		if self.subs:  # 属性やメソッドなどがあるとき
			s = self._superInclude(ms.pop(), interface)  # 最後の要素を置換して処理する。
			ms.append(s + " {\n" + ";\n".join(map(lambda x:tab * 2 + x, self.subs)) + ";\n" + tab + "};\n")			   
		else:
			s = self._superInclude("", interface)
			ms[-1] += s + ";\n"	
		return name_underscore, self._createNested(ms), name_base + ".idl"  
	def _superInclude(self, s, interface):  # 継承しているUNOIDLやXInterfaceのidlファイルを自動include。
		xinterface = "com.sun.star.uno.XInterface"
		if interface:  # インターフェイスのとき
			if not self.includes:  # インクルードするファイルがあるときはインターフェイスをインクルードしていると決めつける。
				if not self.super.rsplit(".")[-1].startswith("X"):  # インターフェイスを継承していないとき
					self.includes.append(xinterface)  # XInterfaceをインクルードする。
		if self.super:  # 継承している時
			if not self.super in self.includes:  # 継承したUNOIDLをincludeしてなければ
				self.includes.append(self.super)  # includeに追加。
			s += " : " + self.super.replace(".", "::")  # 継承しているUNOIDLを表記。
		return s
	def _createNested(self, ms):  # モジュールの入れ子の表記。
		s = ""
		ms.reverse()
		out = ms.pop()  # 最外層は{};でくくらない。
		for m in ms:
			s = " {" + m + s + "};"
		s = out + s
		return s
def createIDLs(c):
	myidl_path = os.path.join(c["src_path"], "idl")  # PyDevプロジェクトのidlフォルダへの絶対パスを取得。
	if not os.path.exists(myidl_path):  # idlフォルダがないとき
		os.mkdir(myidl_path)  # idlフォルダを作成
	os.chdir(myidl_path)  # idlフォルダに移動	 
	for f in glob.iglob("*.idl"):
		c["backup"](f)  # 既存のidlファイルを削除。
	for idl in defineIDLs():  # 各UNOIDLの定義について。
		name_underscore, m, fn = idl.getVal()  # アンダースコア名、モジュールのリスト、出力するidlファイル名を取得。
		with open(fn, "w", encoding="utf-8") as f:
			lines = list()
			lines.append("#ifndef " + name_underscore)
			lines.append("#define " + name_underscore)
			lines.append("")
			for inc in idl.includes:  # includeするidlファイル。
				s = inc.replace(".", "/") + ".idl"
				lines.append("#include <" + s + ">")  
			lines.extend(["", m, "", "#endif"])
			f.write("\n".join(lines))
			print(fn + " has been created from the contents defined with defineIDLs() in createIDLs.py.")
if __name__ == "__main__":
	createIDLs(getConfig(False))		
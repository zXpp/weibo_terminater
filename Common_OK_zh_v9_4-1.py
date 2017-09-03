# -*- coding: utf-8 -*-
#
# Copyright © 2009-2010 CEA
# Pierre Raybaut
# Licensed under the terms of the CECILL License
# (see guidata/__init__.py for details)

"""
All guidata DataItem objects demo
A DataSet object is a set of parameters of various types (integer, float,
boolean, string, etc.) which may be edited in a dialog box thanks to the
'edit' method. Parameters are defined by assigning DataItem objects to a
DataSet class definition: each parameter type has its own DataItem class
(IntItem for integers, FloatItem for floats, StringItem for strings, etc.)
"""

from __future__ import print_function

SHOW = True # Show test in GUI-based test launcher

import os
from guidata.dataset.qtwidgets import  DataSetEditGroupBox
from guidata.qt.QtCore import SIGNAL
from guidata.qt.QtGui import QDialog,QGridLayout
from guidata.dataset.datatypes import DataSet,GetAttrProp, FuncProp,ValueProp, BeginGroup, EndGroup
from guidata.dataset.dataitems import (IntItem,BoolItem,MultipleChoiceItem,FilesOpenItem,StringItem,ChoiceItem,DirectoryItem,TextItem)
prop1 = ValueProp(True)
prop2 = ValueProp(False)

class InoutSet(DataSet):
	def updatedir(self, item, value):
		func=lambda x:os.path.basename(x).rsplit('.')[0]
		fun2=lambda x:os.path.basename(x)
		if value and len(value) :
			self.outpath = os.path.split(value[0])[0]
			self.results = u'您选择的文件是：'.encode("gbk")
			self.outprefix = ";".join(map(func, value))
			self.results = "\n\n".join(map(fun2,value))

			self.WebTitle = self.outprefix
			self.FilesDirName = self.outprefix
			print(u"pptfiles: \n", self.results)
	PPTnames = FilesOpenItem(u" 输入", ("ppt", "pptx"),help=u"选择要处理的PPT文件（必选、同路径下可批量处理）",all_files_first=True).set_prop('display',callback=updatedir,icon=u"ppt.png",active=prop1,store=ValueProp(len("PPTnames")))
	outpath = DirectoryItem(u"输出路径",
				help=u"(可选)选择存放输出文件夹的位置，默认为输入文件所在路径").set_prop('display',active=prop1)
	# FilesDirName=StringItem(u"(根)目录名称 ",
		# 			help=u"(可选)输入要替换显示的目录名,默认为当前处理的PPT文件名\n批处理，对所有文件应用此名称").set_prop('display',active=True)
	WebTitle=StringItem(u"网页标题\n(Title)",
					help=u"(可选)输入要替换显示的HTM、XML标题，默认为当前处理的PPT文件名\n批处理，对所有文件应用此名称").set_prop('display',active=prop1)
	outprefix=StringItem(u"文件名称",help=u"(可选)输入要替换显示的文件名(应用于当前所有输出文件)，默认为当前处理的PPT文件名\n批处理，对所有输入应用此名称").set_prop('display',active=prop1)
	results=TextItem(u"已输入文件 ",default=u"已选择输入的PPT文件:\n").set_prop('display',active=False)

class ImageSet(DataSet):
	def update(self, item, value):#更新缩放图名称
		self.ImagesDirName = unicode(value)
	imgtype = ChoiceItem(u"目标格式", ["PNG","JPG","GIF"]).set_prop("display")#.set_pos(col=1)
	#imgtype = ImageChoiceItem(u"图片格式", ["PNG", "JPG"],help=u"选择输出图片格式").vertical(2),callback=update
	raw=BoolItem(u"原图",default=False,help=u"(可选)直接从PPT中导出原尺寸的PNG图片，无缩放").set_prop('display',active=prop1).set_pos(col=0)
	_prop = GetAttrProp("resize")
	#choice = ChoiceItem('Choice', choices).set_prop("display", store=_prop)
	resize = BoolItem(u"缩放图",default=True,help=u"在原尺寸的基础上对宽、高进行按比例缩放,默认勾选").set_pos(col=1).set_prop("display", active=prop1,store=_prop)
	newsize = IntItem(u"重新调整\n宽度", default=709, min=0, help=u"(可选)输出的缩放图宽度，单位dpi,，默认709，上限不得超过原图宽度。",max=2160, slider=True).set_prop("display",active=FuncProp(_prop, lambda x: x))
	ImagesDirName=StringItem(u"缩放图集名称",default="image",
		help=u"(可选)输入存放调整后图片的目录名，默认image").set_prop('display',active=FuncProp(_prop, lambda x: x))

class FormatSet(DataSet):
	Langue=(("zh-cn",u"zh-cn"),("en",u"en"))#加入发布的中英文选择
	g0=BeginGroup(u"① 网络文件")
	webFormat = MultipleChoiceItem(u"",[ "HTM   "," XML  "],help=u"(可选)默认全选",
								  default=(0,1)).set_prop('display').vertical(2).set_pos(col=0)

	_prop3=GetAttrProp("langue")
	langue=ChoiceItem(u"\r\n发布语言", Langue,help=u"选择要发布的语言类型").set_prop("display",store=_prop3).set_pos(col=0,colspan=2)
	_g0=EndGroup(u"网络文件")
	txpdf=MultipleChoiceItem(u"② 文本",["PDF "," TXT"],default=(0,),help=u"默认输出pdf").set_pos(col=1)
	singimg=MultipleChoiceItem(u"③ 图片",[u"原图(集)",u"缩放图(集)",u"长图(目标格式)"],default=(0,1,2),help=u"默认输出原图集、缩放图集\n如要生成htm、xml，请保持勾选")#.set_pos(col=1)
	g1=BeginGroup(u"④ PPT(单张)")
	_prop1 = GetAttrProp("singppt")#g3=,"PPT(X)"
	singppt=BoolItem("PPT(X)",default=False,help=u"发布单张幻灯片").set_prop('display',store=_prop1).set_pos(col=0)
	expind=StringItem(u"幻灯片编号",help=u"输入要单张发布成PPT的幻灯片编号,逗号分割，连续编号短线相连.1,2,3,5-7,8",
					  default="1,2,3").set_prop('display',active=FuncProp(_prop1, lambda x: x)).set_pos(col=1)
	_g1=EndGroup("PPT")
#class Dlg(QDialog):
class MainWindow(QDialog):
	"""
u"EasyPPT"""
	def __init__(self):
		layout = QGridLayout()
		self.setLayout(layout)

		self.setWindowTitle(u"EasyPPT")
		self.setGeometry(50, 50, 300, 250)
		# Instantiate dataset-related widgets:
		self.groupbox1 = DataSetEditGroupBox(u"【名称设置】",InoutSet,show_button=False, comment='')
		self.groupbox2 = DataSetEditGroupBox(u"【图片设置】",ImageSet,show_button=False, comment='')
		self.groupbox3 = DataSetEditGroupBox(u"【格式设置】",FormatSet,comment='')
		#self.btn = QPushButton("Quit", self)
		# self.btn.clicked.connect(self.close)show_button=False

		layout.addWidget(self.groupbox1)
		layout.addWidget(self.groupbox2)
		layout.addWidget(self.groupbox3)
		self.home()

	def home(self):
		self.show()
		#self.update_groupboxes()
		self.connect(self.groupbox3, SIGNAL("apply_button_clicked()"),self.close)

	def print(self):
		print("hellp")
	def update_groupboxes(self):
		#self.groupbox1.dataset.set_readonly() # This is an activable dataset
		self.groupbox1.get()
		#self.show()
		self.groupbox2.get()
		self.groupbox3.get()

if __name__ == "__main__":
	from guidata.qt.QtGui import QApplication
	import sys
	app = QApplication(sys.argv)
	window = MainWindow()
	window.home()
	print(window.__dict__)
	sys.exit(app.exec_())
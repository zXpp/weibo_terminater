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

from guidata.dataset.datatypes import DataSet,BeginGroup, EndGroup,GetAttrProp, FuncProp,ValueProp
from guidata.dataset.dataitems import (IntItem,BoolItem,MultipleChoiceItem,FilesOpenItem,StringItem,ChoiceItem,DirectoryItem,TextItem)
prop1 = ValueProp(True)
prop2 = ValueProp(False)
choices=(("all",u"全部"),("range",u"选择范围"))
class TestParameters(DataSet):
    """
    If you have any question,please keep your
z81022868：高亮区域不允许留空
    """
    def updatedir(self, item, value):
        func=lambda x:os.path.basename(x).rsplit('.')[0]
        if value and len(value) :
        #func=lambda x:os.path.basename(x).encode('gbk').rsplit('.')[0]
            self.outpath = os.path.split(value[0])[0]
            #self.outprefix=''
            #if self.results is None:
            self.results = u'您选择的文件是：'.encode("gbk")
            self.outprefix = ";".join(map(func, value))
            self.results = "\n".join(value)

            self.WebTitle = self.outprefix
            self.FilesDirName = self.outprefix
            print(u"pptfiles: \n", self.results)

        #print("\nitem: ", item, "\nvalue:", value)encode('gbk').
        # if value:
        #     self.outpath = os.path.split(value[0])[0]
        #     if len(value) ==1:#如果只是选择了单个文件
        #         self.outprefix=os.path.basename(value[0]).split('.')[0]
        #     elif len(value)>1:#选了多个文件
        #         #self.outpath = os.getcwdu()
        #         self.outprefix= ";".join([os.path.basename(x).partition('.')[0] for x in self.PPTnames])
        #     self.WebTitle=self.outprefix
        #     self.FilesDirName=self.outprefix
        #     self.ImagesDirName=str(self.newsize)
    def update(self, item, value):#更新缩放图名称
        self.ImagesDirName = unicode(value)

    g0=BeginGroup(u"                                       ① 输出名称、路径自定义")
    #enable0=BoolItem(u"仅从PPT处理",default=False).set_prop("display", store=prop1)
    PPTnames = FilesOpenItem(u"输入PPT", ("ppt", "pptx"),help=u"选择要处理的PPT文件（必选、同路径下可批量处理）",all_files_first=True).set_prop('display',callback=updatedir,icon=u"ppt.png",active=prop1,store=ValueProp(len("PPTnames")))
    outpath = DirectoryItem(u"输出路径",
				help=u"(可选)选择存放输出文件夹的位置，默认为输入文件所在路径").set_prop('display',active=prop1)
    # FilesDirName=StringItem(u"(根)目录名称 ",
		# 			help=u"(可选)输入要替换显示的目录名,默认为当前处理的PPT文件名\n批处理，对所有文件应用此名称").set_prop('display',active=True)
    WebTitle=StringItem(u"网页标题\n(Title)",
					help=u"(可选)输入要替换显示的HTM、XML标题，默认为当前处理的PPT文件名\n批处理，对所有文件应用此名称").set_prop('display',active=prop1)
    outprefix=StringItem(u"文件名称",help=u"(可选)输入要替换显示的文件名(应用于当前所有输出文件)，默认为当前处理的PPT文件名\n批处理，对所有输入应用此名称").set_prop('display',active=prop1)
    #_prop3=GetAttrProp("selecttype")
    #choices=(("all",u"全部"),("range",u"选择范围"))
    #selecttype=ChoiceItem(u"应用范围", choices).set_prop("display",store=_prop3).set_pos(col=0)
    #pptrange=StringItem(u"输入编号",help=u"输入要应用的幻灯片编号范围，如1,2,3,5-7,9\n"
                                                #u"批量处理时，该范围应用于批量处理时的全部文件，慎重选择",default="").set_prop('display',active=FuncProp(_prop3, lambda x: x=="range")).set_pos(col=1)
    results=TextItem(u"Result ",default=u"已选择输入的PPT文件:\n").set_prop('display',active=False)
    _g0=EndGroup(u"                                       ① 输出名称、路径自定义")
    #enable=BoolItem("enable",default=False)
    g1=BeginGroup(u"                                             ② 输出图片自定义")
    #enable1=BoolItem("hah",default=False).set_prop('hide')
    #Imgin=DirectoryItem(u"图片输入").set_prop('hide')

    imgtype = ChoiceItem(u"目标图片格式", ["PNG","JPG","GIF"]).set_prop("display")#.set_pos(col=1)

    #imgtype = ImageChoiceItem(u"图片格式", ["PNG", "JPG"],help=u"选择输出图片格式").vertical(2)
    raw=BoolItem(u"原图模式",default=False,help=u"(可选)直接从PPT中导出原尺寸的PNG图片，无缩放").set_prop('display',active=prop1).set_pos(col=0)
    _prop = GetAttrProp("resize")
    #choice = ChoiceItem('Choice', choices).set_prop("display", store=_prop)
    resize = BoolItem(u"缩放图模式",default=True,help=u"在原尺寸的基础上对宽、高进行按比例缩放,默认勾选").set_pos(col=1).set_prop("display", active=prop1,store=_prop)
    newsize = IntItem(u"图片宽度", default=709, min=0, help=u"(可选)输出的缩放图宽度，单位dpi,，默认709，上限不得超过原图宽度。",max=2160, slider=True).set_prop("display",active=FuncProp(_prop, lambda x: x),callback=update)
    ImagesDirName=StringItem(u"缩放图集名称",
		help=u"(可选)输入要替换显示的缩放图文件夹名称，默认以当前尺寸命名").set_prop('display',active=FuncProp(_prop, lambda x: x))
    _g1 = EndGroup(u"                                             ② 输出图片自定义")

    g2=BeginGroup(u"                                              ③ 输出格式自定义")
    #must=BoolItem(u" 图集\n(必选)",default=True,help=u"始终输出原图集、缩放图集").set_prop("display",active=0).set_pos(col=0)
    singimg=MultipleChoiceItem(u"单图",[u"原图集",u"缩略图集"],default=(0,1),help=u"默认输出原图集、缩略图集\n如要生成html、xml，请保持勾选").horizontal(2).set_pos(col=0)

    outFormat = MultipleChoiceItem(u"合成",
                                  ["PDF","HTML","XML","TXT",u"长图(目标格式)"],help=u"(可选)默认输出前4种类型",
                                  default=(0,1,2,4)).set_prop('display').vertical(3).set_pos(col=1)
    _prop1 = GetAttrProp("singppt")#g3=,"PPT(X)"
    singppt=BoolItem("PPT(X)",default=False).set_prop('display',store=_prop1).set_pos(col=1,colspan=2)
    expind=StringItem(u"幻灯片编号",help=u"输入要单张发布成PPT的幻灯片编号,逗号分割",default="1,2,3").set_prop('display',active=FuncProp(_prop1, lambda x: x)).set_pos(col=1)
    _g2=EndGroup(u"                                              ③ 输出格式自定义")

#TestParameters._active_prop()
#TestParameters.active_setup()
#TestParameters.items_list
ppt_diago = TestParameters(u"EasyPPTools_v1")


ppt_qg=TestParameters("QuickGuide")

if __name__ == "__main__":
#     #Create QApplication
     import guidata#,subprocess
     _app = guidata.qapplication()
     #print([x for x in TestParameters.items_list()])
     #ppt_diago.set_writeable()
     ppt_diago.edit()
     #ppt_diago.set_readonly()
     # ppt_diago.set_writeable()
     #ppt_diago.view()
    #如果面板被编辑过。# if ppt_diago.edit()
     # args=ppt_diago.__dict__
     # pptname=args["_PPTnames"][0]#+chr(34)chr(34)+
     # newsize=args["_newsize"]
     # outpath=args["_outpath"]
     # pdfout=1
     # if 2 not in args["_outFormat"]:
     #     pdfout=0

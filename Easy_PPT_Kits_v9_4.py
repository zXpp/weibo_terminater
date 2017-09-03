# -*- coding: utf-8 -*-
"""
Created on Tue Jul 25 15:23:13 2017
@author: z81022868
"""
import sys,ctypes
from zx_module.ppt_v10 import easyPPT
from zx_module.Common_OK_zh_v9_4_1 import MainWindow
from zx_module.Webppt_v9_3 import *
global pyrealpath
def Mbox(title, text, style):
	return ctypes.windll.user32.MessageBoxW(0, text, title, style)

if __name__ == "__main__":
	#global pyrealpath
	if getattr(sys,'frozen',False):
		pyrealpath = sys._MEIPASS
	else:
		pyrealpath = os.path.split(os.path.realpath(__file__))[0]

	from guidata.qt.QtGui import QApplication
	args={}
	app = QApplication(sys.argv)
	ppt_diago = MainWindow()
	ppt_diago.show()
	state=app.exec_()
	a,b,c=map(lambda x:x.dataset.__dict__,[ppt_diago.groupbox1,ppt_diago.groupbox2,ppt_diago.groupbox3])
	args.update(a)
	args.update(b)
	args.update(c)
	#sys.exit(status=state)


	singing,language=list(args["_singimg"]),args["_langue"]#singimg 是图集，不光是单图
	try:
		if args["_imgtype"]==0:
			ImgTYPE="PNG"
		elif args["_imgtype"] == 1:
			ImgTYPE="JPG"
		elif args["_imgtype"]==2:
			ImgTYPE="GIF"
		Num,outdir,imageDir,width=len(args["_PPTnames"]),args["_outpath"],args["_ImagesDirName"],args["_newsize"]
		func=lambda x:[args["_outprefix"],args["_WebTitle"]] if Num == 1 else [os.path.basename(x).rsplit('.')[0]]*2
		labAtit=map(func,args["_PPTnames"])#label and title
		pptx=easyPPT()
		for numb,pptname in enumerate(args["_PPTnames"]):
			#######################################################
			label,newtitle=labAtit[numb]
			try:
				#pptx=easyPPT()#singlePPT()
				try:
					pptx.open(pptname,outdir,label)#pptx.selppt已经是有序 唯一 合法的编号
				except Exception,e:
					Mbox(u"Error",u"当前文件打开失败，请重试",0)
					sys.exit(-1)
				#######PPT PDF PPTX########不依赖图片的输出
				if 1 in args["_txpdf"]:#txttxpdf=[pdf,txt]
					pptx.saveAs(Format="TXT")
				if 0 in args["_txpdf"]:#导出PDF
					pptx.saveAs(Format="PDF")
				if args["_singppt"] and args["_expind"] not in [""," "]:#PPTX
					pptx.slid2PPT(substr=args["_expind"])
					####原图 缩放图（单图）#################### singimg=原 缩 长
				try:
					Merge=1 if 2 in singing else 0#2代表要拼长图
					kwarg={"newtit":newtitle,"outpre":pptx.outfile_prefix,
					"imgtype":ImgTYPE.lower(),"merge":Merge,"langue":args["_langue"],"choice":args["_webFormat"]
					}#此时choice是网页的，因为只剩他们没处理了。
					if args["_resize"] or 1 in singing:#输出要求1代输出缩放图，是单图中的原图或者勾选了缩放图模式
						pptx.pngExport(imageDir,width,imgtype=ImgTYPE.lower(),merge=Merge)
						lstdir,imageDir=pptx.outresizedir2,imageDir
						webmod(lstdir,imageDir,kwarg)

					if args["_raw"] or 0 in singing:#输出代表选择了原图模式 并且要求输出中选择要生成原图
						pptx.pngExport(imgtype=ImgTYPE.lower(),merge=Merge)
						lstdir,imageDir=pptx.outslidir2,"Slides"
						kwarg["outpre"]=pptx.outfile_prefix#+u'_raw'
						webmod(lstdir,imageDir,kwarg)

				except Exception, e:
					#traceback.print_exc(f)
					Mbox(u"Error",u"当前步骤发生错误，跳过该步骤，执行下一步?",1)
				##########长图 HTML  XML###########
			except Exception, e:
				#traceback.print_exc(f)
				Mbox(u"Error",u"当前步骤发生错误，跳过该步骤，执行下一步?",1)
			pptx.closepres()
			delflag=1 not in args["_singimg"] and 0 not in args["_singimg"]
			if 0 not in singing and delflag :#删除Slides
				pptx.delslides()
			if 1 not in singing and delflag:#删除resize
				pptx.delresize()
			######################
		#pptx.closeppt()#无路如何都不关闭，防止正在工作的文件被退出。
		Mbox(u"Finished",u"运行完毕! 共处理 "+unicode(Num)+u"个文件",0)
		sys.exit(0)
	except Exception ,e:
		Mbox(u"Error",u"最外层发生错误,程序退出",0)
		sys.exit(-1)
""""""







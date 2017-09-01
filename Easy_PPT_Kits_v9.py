# -*- coding: utf-8 -*-
"""
Created on Tue Jul 25 15:23:13 2017
@author: z81022868
"""
import sys,os,ctypes,traceback
from zx_module.merge_img_PIL_v4 import png2pngs
from zx_module.Common_OK_zh_v8 import ppt_diago
from zx_module.ppt_v6 import easyPPT
global pyrealpath
ppSaveAsGIF = 16
ppSaveAsJPG = 17
ppSaveAsPNG = 18
ppSaveAsPDF = 32
XMLTemplate=u"""
<?xml version="1.0" encoding="gb2312"?>
<!--Arbortext, Inc., 1988-2008, v.4002-->
<!DOCTYPE concept PUBLIC "-//OASIS//DTD DITA Concept//EN" "concept.dtd">
<?Pub Inc?>
<concept xml:lang="zh-cn">
<title>##NAME##</title>
<prolog>
<metadata><keywords><keyword>##NAME##</keyword></keywords></metadata>
</prolog>
<conbody>
<section><image href="{toReP}" placement="break"><?Pub Caret1?></image></section>
</conbody>
</concept>
<?Pub *0000001274?>
"""
HTMLTEMPLTE=u"""<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="zh-cn" xml:lang="zh-cn" charset="gbk"><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"><title>test_10.png</title></head><body style="margin: 0px;"><div align="center" font-size="0" display="block"><img src="test_10.png" display = "inline" ></img></br></div></body></html>"""
def Mbox(title, text, style):
	return ctypes.windll.user32.MessageBoxW(0, text, title, style)
def pngs2web(mod,reldirlist=None,outpre=u"",newtit=u""):#reldirlist=relativelist.相对路径的列表,outpre是加了点的
	if mod in ["html","HTML","Html"]:
		#file=u"Htmltemp.html"
		template=HTMLTEMPLTE
		begflag,endflag,oldflag=u'<img src="',u'</br>',u'" display'#pyrealpath=u"D:\\untar\\Panel\\Gudie"f1.read()
	elif mod in ["XML","xml","Xml"]:
		#file=u"VAN_Feature_Glance.xml"
		template=XMLTemplate
		begflag,endflag,oldflag=u'<image href="',u'</image>',u'" placement='
#def readweb()
	#with codecs.open(os.path.join(pyrealpath,u"guidata\\Tmp",file),'rb') as f1:
	content,titbeg,titend,new=template,u'<title>',u'</title>',[]
	matchtag=content[content.find(begflag):content.rfind(endflag)+len(endflag)]
	# if not matchtag:
	# 	content=template
	# 	matchtag=content[content.find(begflag):content.rfind(endflag)+len(endflag)]
	oldstr=matchtag[matchtag.find(begflag)+len(begflag):matchtag.rfind(oldflag)]
	new.extend([matchtag.replace(oldstr,ele) for ele in reldirlist])
	titrep=content[:content.find(titbeg)+len(titbeg)]+newtit+content[content.rfind(titend):content.index(begflag)]
	webout=titrep+''.join(new)+content[content.rfind(endflag)+len(endflag):]
	outname="".join([outpre,mod.lower()])
	out1=open(outname,'wb')
	out1.write(webout.encode('gbk'))
	out1.close()
def modeOut(imgpath,mode,choice=None,out_pre=None,webtitle=None,imgtype="png"):
				imgbase = os.path.basename(imgpath)
				#files = glob.glob(os.path.join(imgpath,'*'+r"."+imgtype))#abspath:files,全部的图片，不分格式
				allimg=os.listdir(imgpath)
				if mode=="Resize":
					suffix=r"."
					realimg=map(lambda x:"".join([r"..","\\",imgbase,"\\",x]),allimg)#glob.glob0("..\\" + imgbase,'*')
				elif mode == "Raw":
					suffix = r"_raw."
					realimg = map(lambda x:"".join([r"..","\\Slides\\",x]),allimg)
				mergeprefix = "".join([out_pre,suffix])
				kwarg1={"reldirlist":realimg,"outpre":mergeprefix,"newtit":webtitle}
				if 1 in choice:#html
					pngs2web("html",**kwarg1)
				if 2 in choice:#xml
					pngs2web("xml",**kwarg1)
				if 4 in choice:#PNG--changtu
					merge_output = "".join([mergeprefix,imgtype]) #absfile
					png2pngs(imgpath,merge_output,flag=imgtype)
def str2pptind(strinput=""):##将输入处理范围的幻灯片编号转为顺序的、不重复的list列表
	selppt=[]
	if strinput and strinput not in [""," "]:
		pptind=strinput.split(",")
		tmp1=map(lambda x:x.split(r"-"),pptind)
		for each in tmp1:
			each1=map(int,each)
			selppt+=range(each1[0],each1[-1]+1)
		selppt=set(sorted(selppt))
	selppt=list(selppt)
	return selppt
if __name__ == "__main__":
	#global pyrealpath
	if getattr(sys,'frozen',False):
		pyrealpath = sys._MEIPASS
	else:
		pyrealpath = os.path.split(os.path.realpath(__file__))[0]
	import guidata,subprocess
	_app = guidata.qapplication()
	if ppt_diago.edit():#如果面板被编辑过。
#print(ppt_diago)
		pass
		#ppt_diago.view()
	args=ppt_diago.__dict__
	singing=list(args["_singimg"])
	f=open(os.path.join(args["_outpath"],r'error.log'),'wb+')
	try:
		if args["_imgtype"]==0:
			ImgTYPE="PNG"
		elif args["_imgtype"] == 1:
			ImgTYPE="JPG"
		elif args["_imgtype"]==2:
			ImgTYPE="GIF"
		Num,outdir,imageDir=len(args["_PPTnames"]),args["_outpath"],args["_ImagesDirName"]
		if args["_selecttype"]=="range" and args["_pptrange"] not in [""," "]:
			ppttoexp=str2pptind(args["_pptrange"])#list,有序 唯一 但不一定合法
		width=args["_newsize"]
		func=lambda x:[args["_outprefix"],args["_WebTitle"]] if Num == 1 else [os.path.basename(x).rsplit('.')[0]]*2
		labAtit=map(func,args["_PPTnames"])#label and title
		#newtitles=map(lambda x:args["_WebTitle"] if Num == 1 else os.path.basename(x).rsplit('.')[0],args["_PPTnames"])
		pptx=easyPPT()
		for numb,pptname in enumerate(args["_PPTnames"]):
			#######################################################
			label,newtitle=labAtit[numb]#map(lambda x:x["_PPTnames"].index(pptname),[labels,newtitles])
			#"PDF","HTML","XML","TXT",u"长图(目标格式)"
			try:
				#pptx=easyPPT()#singlePPT()
				try:
					pptx.open(pptname,outdir,label,pptrang=ppttoexp)#pptx.selppt已经是有序 唯一 合法的编号
					#pptx.delslid()
					#pptx.selppt=ppttoexp#范围还没好
				except Exception,e:
					print(e)
					Mbox(u"Error",u"当前文件打开失败，请重试",0)
					sys.exit(-1)

				#######PPT PDF PPTX########不依赖图片的输出
				if 3 in args["_outFormat"]:#txt
					pptx.saveAs(Format="TXT")
				if 0 in args["_outFormat"]:#导出PDF
					pptx.saveAs(Format="PDF")
				if args["_singppt"] and args["_expind"] not in [""," "]:#PPTX
					pptind=args["_expind"].split(",")
					ppind2=map(lambda x:x.split(r"-"),pptind)

					ppttoexp=map(int,args["_expind"].split(','))
					pptx.slid2PPT(sublst=ppttoexp)
					####原图 缩放图（单图）####################
				try:
					pptx.pngExport(imageDir,width,imgtype=ImgTYPE.lower())
					kwarg={"webtitle":newtitle,"out_pre":pptx.outfile_prefix,
					"choice":args["_outFormat"], "imgtype":ImgTYPE.lower()
					}

					if args["_resize"]:

						#imgedirnaem ,newidth
						if os.path.exists(pptx.outresizedir2):
							modeOut(pptx.outresizedir2,"Resize",**kwarg)
					if 0 in singing:
						pptx.pngExport(imgtype=ImgTYPE.lower())#留空为原尺寸导出
					if args["_raw"]:
						if os.path.exists(pptx.outslidir2):
							pass
						else:
							pptx.pngExport(imgtype=ImgTYPE.lower())#留空为原尺寸导出
						modeOut(pptx.outslidir2,"Raw",**kwarg)
				except Exception, e:
					traceback.print_exc(f)
					Mbox(u"Error",u"当前步骤发生错误，跳过该步骤，执行下一步",0)
				##########长图 HTML  XML###########

				# if args["_resize"]:
				# 	modeOut(pptx.outresizedir2,"Resize",**kwarg)
				# if args["_raw"]:
				# 	modeOut(pptx.outslidir2,"Raw",**kwarg)
			except Exception, e:
				traceback.print_exc(f)
				Mbox(u"Error",u"当前步骤发生错误，跳过该步骤，执行下一步",0)
			pptx.ppt.PresentationBeforeClose(pptx.pres, False)
			#pptx.closepres()
			delflag=1 not in args["_outFormat"] and 2 not in args["_outFormat"]
			if 0 not in singing and delflag :#删除Slides
				pptx.delslides()
			if 1 not in singing and delflag:#删除resize
				pptx.delresize()
			######################
		#pptx.closeppt()#无路如何都不关闭，防止正在工作的文件被退出。
		Mbox(u"Finished",u"运行成功! 共处理 "+unicode(Num)+u"个文件",0)
	except Exception ,e:
		#pptx.closeppt()
		traceback.print_exc(f)
		Mbox(u"Error",u"最外层发生错误",0)
	f.close()
""""""







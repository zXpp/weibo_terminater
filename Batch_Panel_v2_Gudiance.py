# -*- coding: utf-8 -*-
"""
Created on Tue Jul 25 15:23:13 2017
@author: z81022868
"""
#from __future__ import unicode_literals
#show=True
#from __future__ import print_function
import sys,os,codecs,shutil,ctypes,time
import numpy as np
from PIL import Image,ImageFilter
from zx_module.Common_OK import ppt_qg
from zx_module.win32PPT_new import extractTit_PPT
from zx_module.DelDirNoPrompt import removeDir as rmdir

def Mbox(title, text, style):
	return ctypes.windll.user32.MessageBoxW(0, text, title, style)
def VbName(name):
	if name.find(" ") < 0 :#if filebam have jongg
		vbname=name
	else:
		vbname=chr(34)+name+chr(34)
	return vbname
def renamePngs(path):#tmp/tmp2
	#dock=[]
	for root,dirs,files in os.walk(path):	#遍历统计
		for each in files:
			fileabs=os.path.join(root,each)#dirs=[]
			pre=os.path.splitext(fileabs)[0]
			nochang=pre.rpartition("_")#
			if len(nochang[-1])< len(nochang[-1].zfill(3)):
				shutil.move(fileabs,"_".join([nochang[0],nochang[-1].zfill(3)+r".png"]))
		break

def png2pngs(indir,outname):#indir :abso path;outname:abo
	abolist=[os.path.join(indir,x) for x in os.listdir(indir)]
	#print(abolist[0])
	baseimg=Image.open(abolist[0])
	baseimg=baseimg.convert('RGB')
	baseimg=baseimg.filter(ImageFilter.DETAIL)#细节滤波，细节更明显
	sz=baseimg.size
	basemat=np.atleast_2d(baseimg)

	for i in range(1,len(abolist)):
		file=abolist[i]
		im=Image.open(file)
		if im.mode != "RGB":#转化成统一通道数RGB
			conmode=im.convert("RGB")
			im=conmode
		im=im.resize(sz,Image.ANTIALIAS)#每一张图都切割成第一张大小
		mat=np.atleast_2d(im)
		#print(file)
		basemat=np.append(basemat,mat,axis=0)
	final_img=Image.fromarray(basemat)
	final_img.save(outname,"PNG")

def pngs2xml(reldirlist,outname,newtit):#reldirlist=relativelist.相对路径的列表
	with codecs.open(os.path.join(pyrealpath,
	r"guidata\Tmp\VAN_Feature_Glance.xml"),'rb') as f1:
		content,begflag,endflag,new=f1.read(),'<image href="','</image>',[]
		titbeg,titend='<title>','</title>'
		#content=content.repalce("title>##NAME##</title","title>"+newtit+"</title")
		#matchline=[line for line in content if begflag in line][0]
		matchtag=content[content.find(begflag):content.rfind(endflag)+len(endflag)]
		oldstr=matchtag[matchtag.find(begflag)+len(begflag):matchtag.rfind('" hwFormat=')]
		#repeach=matchtag.replace(oldstr,ele)content[:content.index(begflag)]
		new.extend([matchtag.replace(oldstr,ele) for ele in reldirlist])
		titrep=content[:content.find(titbeg)+len(titbeg)]+\
						newtit+content[content.rfind(titend):content.index(begflag)]
		xmlout=titrep+''.join(new)+content[content.rfind(endflag)+len(endflag):]
		out1=open(outname,'w')
		out1.write(xmlout.encode('gbk'))
		##print xmlout >out1,outname
		out1.close()
def pngs2html(reldirlist,outname,newtit):
	with codecs.open(os.path.join('',r"guidata\Tmp\Htmltemp.html"),'rb') as f2:
		content,begflag,endflag,new=f2.read(),'<img src="','</br>',[]
		titbeg,titend='<title>','</title>'
		#matchline=[line for line in content if begflag in line][0]
		#content=content.repalce("title>##NAME##</title","title>"+newtit+"</title")
		matchtag=content[content.find(begflag):content.rfind(endflag)+len(endflag)]
		oldstr=matchtag[matchtag.find(begflag)+len(begflag):matchtag.rfind('" display')]
		#repeach=matchtag.replace(oldstr,ele)content[:content.index(begflag)]
		#new.extend([matchtag.replace(oldstr,ele)])
		new.extend([matchtag.replace(oldstr,ele) for ele in reldirlist])
		titrep=content[:content.find(titbeg)+len(titbeg)]+newtit\
				+content[content.rfind(titend):content.index(begflag)]
		#没有闭合操作符，用反斜杠将太长的一行换位多行
		writeout=titrep+"".join(new)+content[content.rfind(endflag)+len(endflag):]
		out=open(outname,'wb')
		out.write(writeout.encode('gbk'))
		out.close()

if __name__ == "__main__":

	if getattr(sys,'frozen',False):
		pyrealpath = sys._MEIPASS
	else:
		pyrealpath = os.path.split(os.path.realpath(__file__))[0]
	import guidata,subprocess
	_app = guidata.qapplication()
	if ppt_qg.edit():ppt_qg.view()#pass#如果面板被编辑过。
#print(ppt_qg)
	args=ppt_qg.__dict__
	Num=len(args["_PPTnames"])
	width,raw=args["_newsize"],args["_outpath"]
	imageDir=args["_ImagesDirName"] if args["_ImagesDirName"] not in ['',' '] else str(width)

	for pptname in args["_PPTnames"]:
		#pptname = args["_PPTnames"][0]
		#label=os.path.basename(pptname).split('.')[0]
		#print label--文件的名字，后缀之前的部分
		label=args["_outprefix"] if Num == 1 else os.path.basename(pptname).split('.')[0]
		vbpptname,vbout,vblabel,vbimagedir=map(VbName,[pptname,raw,label,imageDir])
		eachtit=extractTit_PPT(pptname,3)##从目录中读出的二级标题

		#print raw---选择的输出路径，在该路径下生成以label 为名称的文件夹 0--longpng;1--pdf,2--html,3--cml
		#pyrealpath=os.path.split(os.path.realpath(__file__))[0]#'os.path.split(sys.argv[0])[0]
		pdfout,tmp= 1 if 1 in args["_outFormat"] else 0,os.path.join(raw,label,imageDir)
		#print pyrealpath
		tmp2 =os.path.join(raw,label,"Slides")
		cmd=" ".join([os.path.join(pyrealpath,r"guidata\Tmp\wscript.exe ")+os.path.join(pyrealpath,r"guidata\Tmp\PPT_Panel_imgdir.vbs"),vbpptname,str(width),vbout,str(pdfout),vblabel,vbimagedir]).encode('gbk')
		#print cmd
	 #cmd=" ".join([r"C:\Windows\system32\wscript.exe "+os.path.join(pyrealpath,r"guidata\Tmp\PPT_Panel.vbs"),vbpptname,imageDir,vbout,str(pdfout)])
		child=subprocess.Popen(cmd)#
		child.wait()
		#print ('sucess')
		with open(os.path.join(raw,label,"Content.txt"),'wb') as c:
			c.write("\r\n".join(eachtit))
		Nextflag=2 in args["_outFormat"] or 0 in args["_outFormat"] or 3 in args["_outFormat"]

		if Nextflag and os.path.exists(tmp) and os.path.exists(tmp2):#710cunzai
			renamePngs(tmp);renamePngs(tmp2)
			#xml html 中图片文件的相对路径
			relativepath=["".join([r"..","\\",imageDir,"\\",i]) for i in os.listdir(tmp)]#str(i).zfill(2)
			relativerawpath=["".join([r"..\Slides","\\",k]) for k in os.listdir(tmp2)]
			#新定义并创建分别存放HTML、xml的文件夹
			outdir_web=os.path.join(raw, label,"HtmlAndXml")
			rmdir(outdir_web);os.makedirs(outdir_web)
#			outdir_HTML,outdir_XML=[os.path.join(raw, label,suffix) for suffix \
#			in [label+r"_HTML",label+r"_XML"]]
#			rmdir(outdir_HTML);os.makedirs(outdir_HTML)
#			rmdir(outdir_XML);os.makedirs(outdir_XML)
			#=os.path.join(raw, label, "XML")
			mergename_Sets=[r".pdf",r".png",r'_raw.png']
			pdfname,lpngname,rlpngname,=[os.path.join(raw,label,label+name) for name in mergename_Sets]
			# if os.path.isfile("".join([raw,"\\",label,"\\",label,r".pdf"])):
				# shutil.move("".join([raw,"\\",label,"\\",label,r".pdf"]),outdir)
			try:
				begnum=3
				for each in range(begnum,len(relativepath)):
					newtitle=eachtit[each-3]
					xmlname,htmlname=[os.path.join(outdir_web,newtitle+x) for x\
					in [r".xml",r".html"]]
					#htmlname=
					commonlist = [relativepath[each]]#fast_guidance mode
					#commonlist=relativepath#glance mode
				# #Mbox(u"iam runing...",u'不要着急..我还在努力的拼...完成的时候还会看见我..',0)
					if 3 in args["_outFormat"]and args["_resize"]:
						pngs2xml(commonlist,xmlname,newtitle)
					if 3 in args["_outFormat"]and args["_raw"]:
						png2xml(commonlist,rawxmlname,newtitle)
					if 2 in args["_outFormat"] and args["_resize"]:
						pngs2html(commonlist,htmlname,newtitle)
					if 2 in args["_outFormat"] and args["_raw"]:
						pngs2html(commonlist,rawhtmlname,newtitle)
				if 0 in args["_outFormat"] and args["_resize"]:
					png2pngs(tmp,lpngname)#Imgarr,lpng=
					time.sleep(3)
				if 0 in args["_outFormat"] and args["_raw"]:
					png2pngs(tmp2,rlpngname)
					time.sleep(3)
			except:
				print sys.exc_info()
				Mbox(u'Sorry', u'Something Error,Change filename to English and retry', 0)
				#raw_input('press anykey to exit')
	Mbox(u'HEllo', u'可以啦 \n', 0)

"""
if os.path.isfile(pdfname) and os.path.isfile(lpngname) and os.path.isfile(htmlname) and os.path.isfile(xmlname):
	user32.MessageBoxA(0,"Finished Sucessdully! GO and Check Your file",'Ctypes',0)"""
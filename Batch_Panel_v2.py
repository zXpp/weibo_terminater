# -*- coding: utf-8 -*-
"""
Created on Tue Jul 25 15:23:13 2017
@author: z81022868
"""
#from __future__ import unicode_literals
#show=True
#from __future__ import print_function
import sys,os,codecs,shutil,tempfile,ctypes
import numpy as np
from PIL import Image,ImageFilter
from Common_OK import ppt_diago
from zx_module.DelDirNoPrompt import removeDir as rmdir
def VbName(name):
	if name.find(" ") < 0 :#if filebam have jongg
		vbname=name
	else:
		vbname=chr(34)+name+chr(34)
	return vbname
def Mbox(title, text, style):
	return ctypes.windll.user32.MessageBoxW(0, text, title, style)
#ctypes,
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
	with codecs.open(os.path.join(pyrealpath,r"guidata\Tmp\VAN_Feature_Glance.xml"),'rb') as f1:
		content,begflag,endflag,new=f1.read(),'<image href="','</image>',[]
		titbeg,titend='<title>','</title>'
		#content=content.repalce("title>##NAME##</title","title>"+newtit+"</title")
		#matchline=[line for line in content if begflag in line][0]
		matchtag=content[content.find(begflag):content.rfind(endflag)+len(endflag)]
		oldstr=matchtag[matchtag.find(begflag)+len(begflag):matchtag.rfind('" hwFormat=')]
		#repeach=matchtag.replace(oldstr,ele)content[:content.index(begflag)]
		new.extend([matchtag.replace(oldstr,ele) for ele in reldirlist])
		titrep=content[:content.find(titbeg)+len(titbeg)]+newtit+content[content.rfind(titend):content.index(begflag)]
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
		new.extend([matchtag.replace(oldstr,ele) for ele in reldirlist])
		titrep=content[:content.find(titbeg)+len(titbeg)]+newtit+content[content.rfind(titend):content.index(begflag)]
		writeout=titrep+"".join(new)+content[content.rfind(endflag)+len(endflag):]
		out=open(outname,'w')
		out.write(writeout.encode('gbk'))
		out.close()
if __name__ == "__main__":

	if getattr(sys,'frozen',False):
		pyrealpath = sys._MEIPASS
	else:
		pyrealpath = os.path.split(os.path.realpath(__file__))[0]
	import guidata,subprocess
	_app = guidata.qapplication()
	if ppt_diago.edit():pass#如果面板被编辑过。
#print(ppt_diago)
	ppt_diago.view()
	args=ppt_diago.__dict__
	Num=len(args["_PPTnames"])
	collectname=args["_FilesDirName"]
	width,raw=args["_newsize"],args["_outpath"]
	imageDir=args["_ImagesDirName"] if args["_ImagesDirName"] not in ['',' '] else str(width)
	try:
		for pptname in args["_PPTnames"]:
			#pptname = args["_PPTnames"][0]
			#label=os.path.basename(pptname).split('.')[0]
			#print label--文件的名字，后缀之前的部分
			label=args["_outprefix"] if Num == 1 else os.path.basename(pptname).split('.')[0]
			vbpptname,vbout,vblabel,vbimagedir=map(VbName,[pptname,raw,label,imageDir])
			#eachtit=extractTit_PPT(pptname,3)##从目录中读出的二级标题

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
			Nextflag=2 in args["_outFormat"] or 0 in args["_outFormat"] or 3 in args["_outFormat"]
			if os.path.exists(tmp) and os.path.exists(tmp2) and Nextflag:#710cunzai
				renamePngs(tmp)
				renamePngs(tmp2)
				relativepath=["".join([r"..","\\",str(width),"\\",i]) for i in os.listdir(tmp)]#str(i).zfill(2)
				relativerawpath=["".join([r"..\Slides","\\",k]) for k in os.listdir(tmp2)]
				outdir_web=os.path.join(raw, label,collectionname)
				rmdir(outdir_web);os.makedirs(outdir_web)
				#outdir=tempfile.mkdtemp(prefix=str(width)+"PPTMerged_",dir=os.path.join(raw,label))
				mergename_Sets=[r".pdf",r".png",r".html",r".xml",r'_raw.png',r"_raw.html",r"_raw.xml"]
				pdfname,lpngname,htmlname,xmlname,rlpngname,rawhtmlname,rawxmlname=[os.path.join(outdir,label+name) for name in mergename_Sets]
				if os.path.isfile("".join([raw,"\\",label,"\\",label,r".pdf"])):
					shutil.move("".join([raw,"\\",label,"\\",label,r".pdf"]),outdir)
				try:
					#Mbox(u"iam runing...",u'不要着急..我还在努力的拼...完成的时候还会看见我..',0)
					if args["_resize"]:
						if 3 in args["_outFormat"]:
							pngs2xml(relativerawpath,xmlname,newtitle)
						if 2 in args["_outFormat"]:
							pngs2html(relativerawpath,htmlname,newtitle)
					#pngs2html(relativepath,htmlname,newtitle)
						if 0 in args["_outFormat"]:
							png2pngs(tmp2,lpngname)
					if args["_raw"]:#pngs2xml(relativepath,xmlname,newtitle)
						if 3 in args["_outFormat"]:
							pngs2xml(relativerawpath,rawxmlname,newtitle)
						if 2 in args["_outFormat"]:
							pngs2html(relativerawpath,rawhtmlname,newtitle)
						if 0 in args["_outFormat"]:
							png2pngs(tmp,rlpngname)#Imgarr,lpng=

					#print ('ok')
				except:
					Mbox(u'Error', u'Something Failed During Concentrate,Check Your File and Retry', 0)
					#print sys.exc_info()

		Mbox(u'Finished', u'可以啦\n', 0)			#raw_input('press anykey to exit')
	except:
		Mbox(u'Error', u'Something Failed,Check Your PPT Format and Retry', 0)


"""
if os.path.isfile(pdfname) and os.path.isfile(lpngname) and os.path.isfile(htmlname) and os.path.isfile(xmlname):
	user32.MessageBoxA(0,"Finished Sucessdully! GO and Check Your file",'Ctypes',0)"""
# -*- coding: utf-8 -*-
#import win32com.client
from win32com.client import Dispatch
from zx_module.rmdjr import rmdAndmkd as rmdir#import timeremove and makedir
import os

OFFICE=["PowerPoint.Application","Word.Application","Excel.Application"]
ppFixedFormatTypePDF = 2
class easyPPT():
    def __init__(self):
        self.ppt = Dispatch('PowerPoint.Application')

        # try:
        #     if filename and os.path.isfile(filename):
        #         self.filename = filename#给定文件名，chuangeilei
        #         self.ppt_prefix=os.path.basename(self.filename).rsplit('.')[0]
        #         self.outdir0=outdir0 if outdir0 else os.path.pardir(self.filename)#输出跟路径
        #         self.label = label if label else self.ppt_prefix
        #         self.outdir1=os.path.join(self.outdir0,self.label)#输出跟文件夹路径+用户定义-同ppt文件名
        #     else:
        #         self.pres = self.ppt.Presentations.Add()
        #         self.filename = ''
        # except Exception, e:
        #     print e
        #self.ppt.Visible = 1，label生成根目录名称，,filename=None,outdir0=None,label=None
    def open(self,filename=None,outdir0=None,label=None,pptrange=None):
        #self.__init__(self,filename=None,outdir0=None,label=None)
        try:
            if filename and os.path.isfile(filename):
                self.filename = filename#给定文件名，chuangeilei
                self.ppt_prefix=os.path.basename(self.filename).rsplit('.')[0]
                self.outdir0=outdir0 if outdir0 else os.path.pardir(self.filename)#输出跟路径
                self.label = label if label else self.ppt_prefix
                self.outdir1=os.path.join(self.outdir0,self.label)#输出跟文件夹路径+用户定义-同ppt文件名, WithWindow=False

                self.pres = self.ppt.Presentations.Open(self.filename)
            else:
                self.pres = self.ppt.Presentations.Add()
                self.filename = ''
            self.Slides=self.pres.Slides
            self.count_Slid = self.Slides.Count
            self.allppt=range(1,self.count_Slid+1)#所有的PPT编号
            ###########################################
            if pptrange and len(pptrange):
                self.selppt=filter(lambda x:x>0 and x<self.count_Slid+1,pptrange)
            else:
                self.selppt=self.allppt
            ##################################################
            self.width,self.height = [self.pres.PageSetup.SlideWidth*3,self.pres.PageSetup.SlideHeight*3]
            rmdir(self.outdir1)
            self.outfordir2=os.path.join(self.outdir1,"OutFormats")#输出放多格式的二级目录,默认为"OurFormats"
            rmdir(self.outfordir2)
            self.outfile_prefix=os.path.join(self.outfordir2,self.label)#具体到文件格式的文件前缀（除了生成格式外的）

            #os.makedirs(self.outdir1)
            # if self.outdir1:
            #     #os.path.splitext(self.filename)[0]#self.filename.split('.')[0]#路径+前缀
            #     self.outfordir2=os.path.join(self.outdir1,"OutFormats")#输出放多格式的二级目录,默认为"OurFormats"
            #     rmdir(self.outfordir2)
                #os.makedirs(self.outfordir2)
                # self.outslidir2=os.path.join(self.outdir1,"Slides")#存放原图的二级目录
                # rmdir(self.outslidir2)
                #os.makedirs(self.outslidir2)
            #self.outresizedir2=os.path.join(self.outdir1,)            #self.dpi_size=map(lambda x: (x/72)+1,[self.width,self.height])
        except Exception,e:
            return e
            #print ("sorry")
    #def saveAsPDF(self,):
    def closepres(self):
        self.pres.Close()
    def closeppt(self):
        self.ppt.Quit()
    def saveAs(self,Format=None,flag="f"):
        # if newfilename:
        #     self.filename = newfilename
        # if not newfilename and not Format:
        #     self.pres.Save()
        try:
            if flag=="f" and Format:
                #self.outfile_prefix=os.path.join(self.outfordir2,self.label)#具体到文件格式的文件前缀（除了生成格式外的）

            #if :
                newname=u".".join([self.outfile_prefix,Format])#目标格式的完整路径
                #rmdir(newname)
                if Format in ["PDF","pdf","Pdf"]:
                    #self.pres.ExportAsFixedFormat(Path=newname,FixedFormatType=2)
                    self.pres.SaveAs(FileName=newname,FileFormat=32)
                if Format in ["txt", "TXT", "Txt"]:
                    self.exTXT()
                # if Format =="xml":
                #     self.pres.SaveAs(FileName=newname, FileFormat=34)
                # if Format in ["HTML", "htm", "html"]:
                #     self.pres.SaveAs(FileName=newname, FileFormat=20,EmbedTrueTypeFonts=-2)
                # #self.close()
        except Exception,e:
            return e

    def pngExport(self,imgDir=None,newsize=709,imgtype="png"):#导出原图用这个
        radio=round(newsize/self.width,0)
        if self.outdir1:
            self.outslidir2=os.path.join(self.outdir1,"Slides")#存放原图的二级目录
            if not imgDir:
                rmdir(self.outslidir2)
                #在Presentation的层面上全面导出
                self.pres.Export(self.outslidir2,imgtype,self.width,self.height)#可设参数导出幻灯片的宽度（以像素为单位）
                self.renameFiles(self.outslidir2,torep=u"幻灯片")
            else:
                self.imDir=imgDir
                self.outresizedir2=os.path.join(self.outdir1,self.imDir)
                rmdir(self.outresizedir2)
                #os.makedirs(self.outresizedir2)
                self.pres.Export(self.outresizedir2, imgtype,newsize,round(self.height*radio,0))  # 可设参数
                self.renameFiles(self.outresizedir2,torep=u"幻灯片")
    def renameFiles(self,imgdir_name,torep=None):#torep:带替换的字符
        srcfiles = os.listdir(imgdir_name)
        for srcfile in srcfiles:
            inde = srcfile.split(torep)[-1].split('.')[0]#haha1.png
            suffix=srcfile[srcfile.index(inde)+len(inde):].lower()
            #sufix = os.path.splitext(srcfile)[1]
            # 根据目录下具体的文件数修改%号后的值，"%04d"最多支持9999
            newname="".join([self.label,"_",inde.zfill(3),suffix])
            destfile = os.path.join(imgdir_name,newname)
            srcfile = os.path.join(imgdir_name, srcfile)
            os.rename(srcfile, destfile)
            #index += 1
#        for each in os.listdir(imgdir_name):
    def slid2PPT(self,outdir=None,sublst=None):#选择的编号的幻灯片单独发布为ppt文件，
        try:
            if not outdir:
                outdir=os.path.join(self.outdir1 ,"Slide2PPT")
            if not os.path.isdir(outdir):#只需要判断文件夹是否存在，默认是覆盖生成。
                os.makedirs(outdir)
            if not sublst or not len(sublst):#如果不指定，列表为空或者是空列表
                #return#sublst=range(1,self.count_Slid+1)
            #else:
                sublst=self.selppt
                #sublst=filter(lambda x:x>0 and x<self.count_Slid+1,sublst)#筛选出小于幻灯片总数的序号
                map(lambda x:self.Slides.Item(x).PublishSlides(outdir),sublst)
                self.renameFiles(outdir,torep=self.ppt_prefix)
        except Exception,e:
            return e
    def exTXT(self):
        s=[]
        for x in self.selppt:#range(1,self.count_Slid+1):#
            shape_count = self.Slides(x).Shapes.Count
            for j in range(1, shape_count + 1):
                if self.Slides(x).Shapes(j).HasTextFrame:
                    s.append(self.Slides(x).Shapes(j).TextFrame.TextRange.Text)#
        if len(s):
            with open("".join([self.outfile_prefix,r".txt"]),"wb") as f:#  = "\\r\\n".join([s,])
                f.write(u"\r\n".join(s).encode("utf-8"))
        else:#如果没有字的情况
            return
    def delslides(self):
        rmdir(self.outslidir2,mksd=0)
    def delresize(self):
        rmdir(self.outresizedir2,mksd=0)
    def weboptions(self):
        web=self.pres.WebOptions
        web.FrameColors = 4 #ppFrameColorsWhiteTextOnBlack
        web.AllowPNG = True

        '''
     '"""
     enum {
        ppPublishAll = 1,
        ppPublishSlideRange = 2,
        ppPublishNamedSlideShow = 3
    } PpPublishSourceType;
    def SlideResize(self,newidth=None,imgDir=None):#用在导出原图，留空吧
        #传6参数：vbpptname,str(width),vbout,str(pdfout),vblabel,vbimagedir
        #imgDir不带路径，out0是指整个输出文件夹的存放路径，必须，label是指文件冲命名输出。
        if newidth :
            newidth=int(newidth) if newidth else self.width
        else:
            newidth=self.width
        self.outresizedir2=os.path.join(self.outdir1,imgDir)
        if self.width>newidth:
            #radio=newidth/(self.width) +1
            newheight=round(newidth*(self.height)/(self.width),0)
            for ind in range(1,self.count_Slid+1):
                self.imgDir=os.path.join(self.prefix,self.label)
                rmdir(self.imgDir);os.makedirs(self.imgDir)
                outname="".join([self.imgDir+"\\"+self.label,"_",str(ind).zfill(3),".png"])
                rmdir(outname)
                self.Slides.Item(ind).Export(outname, "png",newidth,newheight)
        #else:"""
'''
    def str2pptind(self,strinput=""):
        #self.selppt=[]
        if strinput and strinput not in [""," "]:
            pptind=strinput.split(",")
            tmp1=map(lambda x:x.split(r"-"),pptind)
            for each in tmp1:
                each1=map(int,each)
                self.selppt+=range(each1[0],each1[-1]+1)
            self.selppt=set(sorted(self.selppt))
        self.selppt=list(self.selppt)
        # else:
        #     final=range(1,self.count_Slid+1)
        return self.selppt##fan会用户编号
    # def delslid(self,exclude=None):
    #     if not exclude:
    #         exclude=self.selppt
    #     todel=filter(lambda x:x not in exclude,self.allppt)
    #     if len(todel):
    #         map(lambda x:self.Slides.Item(x).Select(),todel)#you can shanchu
    #         self.ppt.Selection.Delete()


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
     label=u"bug 的zx 啊"
     imdir=u"反对果"
     outdir = u"D:\\HW\\张雪\\中文测试"#u"G:\\简历"+u"\\"+u'张雪--兼辅PPT.ppt'
     inpu=u"1,2,4,6-9,5,8-11,22-25"
     pprange=str2pptind(inpu)
     pptx=easyPPT()

     pptx.open(u"C:\\Users\\z81022868\\Desktop\\hahadel.pptx",outdir,label,pptrange=pprange)
     pptx.delslid()
     # print pptx.width,pptx.height,pptx.count_Slid,pptx.pres.PageSetup.SlideHeight
     # print pptx.pres.PageSetup.SlideWidth

     #pptind=pptx.str2pptind(inpu)
     #pptind2=pptx.str2pptind("")
     format=["png","html","xml","ppt/pptx","txt","pdf"]#formatpptx.saveAs(Format="PNG")
     pptx.saveAs(Format=format[-1])#txt
     # pptx.saveAs(Format=format[-1])#导出pdf
     pptx.pngExport()#留空为原尺寸导出
     pptx.pngExport(u"啦啦 啦",700)#imgedirnaem ,newidth.
     pptx.slid2PPT()
     #pptx.slid2PPT(sublst=pptind)
     #pptx.weboptions()
     #print "dpi:",pptx.dpi_size[0],pptx.dpi_size[1]
     pptx.ppt.PresentationBeforeClose(pptx.pres, False)
     #pptx.close()

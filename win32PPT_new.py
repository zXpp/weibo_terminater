# -*- coding: utf-8 -*-
"""
Created on Mon Aug 07 01:25:07 2017

@author: z81022868
"""
import win32com
from win32com.client import Dispatch
def Cut(sigin):
      if sigin:
        ind1,ind2=sigin.find(u"…"),sigin.find('...')
        if ind1>0 and ind2>0:#<ind2:
            c=sigin[:min(ind1,ind2)]#.partition(u"…")[0].rstrip()
        else:
            c=sigin[:max(ind1,ind2)]#.partition('..')[0].rstrip()
        return c
def catalog(strin):
    #if isinstance(strin,unicode):(?:[^\W\d_]*))---
    lines=strin.splitlines()
    lines=[x for x in lines if len(x) and not x.isspace()]
    a=map(Cut,lines)
    m,p=0,[]
    while m <len(a)-1:
        if m <len(a)-1 and a[m][0] !=' ' and a[m+1][0] != ' ':
            p.append(a[m])
            m+=1
            if m == len(a)-1:
                p.append(a[m])
                return p
        elif m<len(a)-1:
            b=[]
            for n in range(m+1,len(a)):
                if m <len(a) and a[n][0] == ' ':
                    #print "**"+a[n]+"**"+a[n].strip()+"aa"
                    if not (a[n].strip() == ' ' or a[n].strip() ==''):
                        b.append(a[n].strip())
                    else:
                        p.append(a[m])
                        return p
                else:
                    m=n
                    p.extend(b)
                    break
        #return p

def extractTit_PPT(pfile,index):
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    ppt.Visible = 1
    pptSel = ppt.Presentations.Open(pfile)
    win32com.client.gencache.EnsureDispatch('PowerPoint.Application')
    slide_count = pptSel.Slides.Count
    each,mulu=pptSel.Slides(index),pptSel.Slides(index).Shapes.Count
    content,flag=[],False
    for j in range(1,mulu + 1):
        if each.Shapes(j).HasTextFrame :
              s = each.Shapes(j).TextFrame.TextRange.Text
              if u"目录" in s or u"Content" in s:
                  flag=True
                  continue
              if flag and s!= '':
                  content.append(s)
                  break
    #cata=catalog(s)#cata is the title
    pptSel.Close()
    ppt.Quit()
    return catalog(s)
if __name__=="__main__":
    pn=u'D:\\untar\\Panel\\Gudie\\BatchProcessTestPPT\\ea5800pinst\\EA5800-X17(N63E-22) Quick Installation Guide 01.pptx'
    x=extractTit_PPT(pn,3)
    #dd=catalog(x)
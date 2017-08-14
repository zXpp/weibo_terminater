'Date: 2017/7/31 
'NO BLANKSPACE!
Dim ppt_path,newWidth,outdir,pdfout,AbsolutePath,exportPreName
 '= ""
Set objArgs = WScript.Arguments '命令行参数 
'if objArgs.Count<1 then 
'return 
'end if 
'ls_File = objArgs(0) '第一个参数通常就是文件名 
'Set vbpath = 
'Dim appPath 
'appPath = Application.StartupPath()
'selfWidth的值设为比需求大1，如709设为710.

'DIm strWidth
'strWidth=cstr(selfWidth)
'WScript.Echo 'selfWidth
'ppt_path = BrowseForFile()
ppt_path=objArgs(0)'ppt abs path
'msgbox "ppt_path"& ppt_path
newWidth = cint(objArgs(1))'710'val()'newsize
outdir=objArgs(2)
'msgbox "outdir"& outdir
pdfout=objArgs(3)' the flag wheather to output pdf
'msgbox "pdfout"& pdfout
exportPreName=objArgs(4)
imageDir=objArgs(5)
'if exportPreName = "" then
'exportPreName=Mid(ppt_path,InStrRev(ppt_path, "\")+1,InStrRev(ppt_path, ".")-InStrRev(ppt_path, "\")-1)
'msgbox "exportPreName" & exportPreName
'end if
AbsolutePath =outdir &"\"& exportPreName'"Tmp"
' Const ForWriting = 2
' Set objFSO = CreateObject("Scripting.FileSystemObject")
' Set objFile = objFSO.OpenTextFile("var.log", ForWriting)
' objFile.WriteLine Now
' objFile.WriteLine "ppt_path: "& ppt_path
' objFile.WriteLine "newWidth: "& newWidth
' objFile.WriteLine "pdfout: "& pdfout
' objFile.WriteLine " outdir: " & outdir
' objFile.WriteLine "AbsolutePath: "& AbsolutePath
' objFile.WriteLine "exportPreName: "& exportPreName
' objFile.Close

post_name=Mid(ppt_path,InStrRev(ppt_path, ".")+1) 
If ppt_path <> ""  Then
 
 if instr(ucase(Mid(ppt_path,InStrRev(ppt_path, ".")+1)),"PPT" ) =0 then  '判断文件是否为ppt 
	msgbox ppt_path & " is  a not PPT file" 
	else 
	Form_Load ppt_path,"png",newWidth,outdir,pdfout,AbsolutePath,imageDir
	msgbox "Split Into pngs successfully! Begin to Convert,Please Wait"
	'Callpy(ppt_path)
	'msgbox "All Done!"
 end if
end if
'objFile.Close

Private Sub Form_Load(Filepath,format,selfWidth,outdad,pdf,Get_AbsolutePath,imageDir)
'msgbox Filepath
If format = "" then
format = "png"
end if
'获取ppt名称，用于输出文件夹的命名
' if exportPreName = "" then
' exportPreName=Mid(ppt_path,InStrRev(ppt_path, "\")+1,InStrRev(ppt_path, ".")-InStrRev(ppt_path, "\")-1)
' msgbox "exportPreName" & exportPreName
' end if
'获取路径,以PPT文件名命名的文件夹--eg,特性Glance格式转换工具需求分析
'Get_AbsolutePath = Left(ppt_path, InStrRev(ppt_path, "\")) &  exportPreName
'Get_AbsolutePath =outdad &"\"& exportPreName'outpath 
msgbox "get:"& Get_AbsolutePath	
'Dim pdfout




'创建输出目录

	set fso=createobject("scripting.filesystemobject")
		str="\"
		getpath=split(Get_AbsolutePath,str)
		for i= 1 to ubound(getpath)
		path=path & str &getpath(i)
		if not fso.folderexists(getpath(0)& str &path)then
		fso.createfolder(getpath(0)& str &path)
		end if
		next
', -1, 0, 0
Set ppApp = CreateObject("PowerPoint.Application")   
Set ppPresentations = ppApp.Presentations      
Set ppPres = ppPresentations.Open(Filepath,-1,0,0)

if pdf then
ppPres.SaveAs Get_AbsolutePath &"\"&exportPreName &".pdf",32, false
end if
'ppPres.ExportAsFixedFormat , ppFixedFormatTypePDF, ppFixedFormatIntentScreen, msoCTrue, ppPrintHandoutHorizontalFirst, ppPrintOutputBuildSlides, msoFalse, , , , False, False, False, False, False 

Set ppSlides = ppPres.Slides
	 Dim Width
	 Dim Height
      With ppPres.PageSetup
         Width = .SlideWidth*3
		 Height = .SlideHeight*3
      End With
	
	  
 '     MsgBox "幻灯片的宽度是： " & (Width / 72) & "英寸，高度是： " & (Height / 72) & "英寸。"1 to ppSildes.Count
savePath_for_raw=Get_AbsolutePath & "\Slides"
' dim k

' for k = 1 To ppSlides.Count
	' rang(k)=k.tostring("D3")
' Next

For  i = 1 To ppSlides.Count
	'range(i)=i
	'Dim k
	'k=i
	iname = exportPreName&"_"& i'k.tostring("D3")
	set fso_raw=createobject("scripting.filesystemobject")
			str="\"
			path_raw=""
			getrawpath=split(savePath_for_raw,str)
			for n= 1 to ubound(getrawpath)
				path_raw=path_raw & str &getrawpath(n)
				if not fso_raw.folderexists(getrawpath(0)& str &path_raw)then
					fso_raw.createfolder(getrawpath(0)& str &path_raw)		
				end if
			next


	Call ppSlides.Item(i).Export(savePath_for_raw &"\"&iname&"."&format, format,Width ,Height)
'call ppSlides.Item(i).SaveAs(Get_AbsolutePath &"\"&iname&"."&format, "ppSaveAsPNG",msoFalse)
'ppSlides.Item(i).SaveAs FileName=Get_AbsolutePath &"\"&iname&"."&format, FileFormat=ppSaveAsPNG, EmbedTrueTypeFonts=msoFalse

'最终图片大小大于selfWidth的使用selfWidth进行导出
	if Width>selfWidth then
		Conversion_ratio=selfWidth/Width+1'表示向上取整
	'	msgbox Conversion_ratio
		Width_new=selfWidth
		Height_new=Height*Conversion_ratio
		savePath_for_selfWidth=Get_AbsolutePath & "\"& imageDir
		'msgbox savePath_for_selfWidth
			set fso_selfWidth=createobject("scripting.filesystemobject")
			str="\"
			path_selfWidth=""
			getpath=split(savePath_for_selfWidth,str)
			for n= 1 to ubound(getpath)
				path_selfWidth=path_selfWidth & str &getpath(n)
				if not fso_selfWidth.folderexists(getpath(0)& str &path_selfWidth)then
					fso_selfWidth.createfolder(getpath(0)& str &path_selfWidth)		
				end if
			next
		
			Dim Img 'As ImageFile
			Dim IP 'As ImageProcess
			Set Img = CreateObject("WIA.ImageFile")
			Set IP = CreateObject("WIA.ImageProcess")
			Img.LoadFile savePath_for_raw &"\"&iname&"."&format
			IP.Filters.Add IP.FilterInfos("Scale").FilterID
			IP.Filters(1).Properties("MaximumWidth") = Width_new
			IP.Filters(1).Properties("MaximumHeight") = Height_new
			Set Img = IP.Apply(Img)
			DeleteAFile savePath_for_selfWidth &"\"&iname&"."&format
			Img.SaveFile savePath_for_selfWidth &"\"&iname&"."&format
		  end if

Next 



'/WScript.echo ppSildes.Count > Folderpath&"\log.txt"
Set ppApp = Nothing

Set ppPres = Nothing
End Sub

'Dim vbpath
'vbpath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)



Function BrowseForFile() 
Dim shell : Set shell = CreateObject("WScript.Shell") 
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject") 

Dim tempFolder : Set tempFolder = fso.GetSpecialFolder(2) 
Dim tempName : tempName = fso.GetTempName() 
Dim tempFile : Set tempFile = tempFolder.CreateTextFile(tempName & ".hta") 

tempFile.Write _ 
"<html>" & _ 
"<head>" & _ 
"<title>Browse</title>" & _ 
"</head>" & _ 
"<body>" & _ 
"<input type='file' id='f' />" & _ 
"<script type='text/javascript'>" & _ 
"var f = document.getElementById('f');" & _ 
"f.click();" & _ 
"var shell = new ActiveXObject('WScript.Shell');" & _ 
"shell.RegWrite('HKEY_CURRENT_USER\\Volatile Environment\\MsgResp', f.value);" & _ 
"window.close();" & _ 
"</script>" & _ 
"</body>" & _ 
"</html>" 

tempFile.Close 
shell.Run tempFolder & "\" & tempName & ".hta", 0, True 
BrowseForFile = shell.RegRead("HKEY_CURRENT_USER\Volatile Environment\MsgResp") 

shell.RegDelete "HKEY_CURRENT_USER\Volatile Environment\MsgResp" 
End Function 

Sub DeleteAFile(filespec)
	Dim fso_del
	Set fso_del =CreateObject("Scripting.FileSystemObject")
		If fso_del.fileExists(filespec) Then 
		fso_del.DeleteFile(filespec)
		end If
End Sub
Sub Callpy(pptpath)
	Dim wshShell, fso, loc, cmd1,cmd,vbspath',newWidth',Get_AbsolutePath
	Set fso = CreateObject("Scripting.FileSystemObject")
	loc = fso.GetAbsolutePathName(".")
	'WScript.Echo loc' 你在什么地方点击右键菜单的export pptisuite
	'~ cmd = "%ComSpec% /k python.exe " + loc + "\all.py"
	Set wshShell = CreateObject("WScript.Shell")
	vbspath = fso.GetParentFolderName(Wscript.ScriptFullName)'vbs 程序的放的地方。就是d:\untar\latest
	'WScript.Echo vbspath "%ComSpec% /k "&"D:\untar\util\python.exe "&vbspath & "\PPTiSuiteSelfWidth " noprompt_PPTiSuiteSelfWidthLatest\noprompt_PPTiSuiteSelfWidthLatest.exe " 
	'cmd = "%ComSpec% /k "+vbspath+"\PPTiSuite-imp3\PPTiSuite-imp3.exe "+ ppt_path +" "+vbspath+" "+selfWidth
	cmd1 = "%ComSpec% /k "&vbspath & "\util\dist\2noprompt_PPTiSuiteSelfWidthLatest2.exe "
	cmd2=cmd1 & pptpath & " " & vbspath &" "& newWidth
	cmd=cmd2 &" "& AbsolutePath 
	'objFile.WriteLine "cmd: "& cmd
	WScript.Echo cmd1
	WScript.Echo cmd
	wshShell.Run cmd,0
End Sub
'ppPres.SaveAs Get_AbsolutePath & exportPreName &".pdf",32,false
'WScript.Echo selfWidth
'Callpy(ppt_path)
'WScript.Echo "All Done!"
'ppPresentations.Quit
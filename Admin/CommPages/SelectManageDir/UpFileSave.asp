<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<% Option Explicit %>
<%Session.CodePage=936%> 
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp"-->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../../FS_InterFace/NS_Function.asp" -->
<!--#include file="News_Upfile.asp" -->
<!--#include file="../../../FS_Inc/WaterPrint_Function.asp" -->
<%
Dim Conn
Dim TypeSql,RsTypeObj,LableSql,RsLableObj
Dim FilePath,MaxFileSize,AllowFileExtStr,AutoReName,RsConfigObj,IsAddWaterMark
Dim FormName,Path,UpFileObj
Dim ReturnValue,str_MF_UpFile_Type,str_MF_UpFile_File_Size
MF_Default_Conn
if not MF_Check_Pop_TF("MF025") then Err_Show
'=========================
'读参数
NS_GetMT_SysParm
'==========================
if str_MF_UpFile_File_Size <>"" then:MaxFileSize = clng(str_MF_UpFile_File_Size):Else:MaxFileSize =1024:End if
if str_MF_UpFile_Type <>"" then:AllowFileExtStr = str_MF_UpFile_Type:Else:AllowFileExtStr = "jpg,gif,jpeg,png,bmp,txt,doc":End if
Set UpFileObj = New UpFileClass
UpFileObj.GetData
Dim dirpath
dirpath = UpFileObj.Form("Dirlist")
If dirpath <> "" And dirpath<>"dir" Then
	FilePath = Server.MapPath(UpFileObj.Form("Dirlist"))& "\"
Else
	FilePath=Server.MapPath(UpFileObj.Form("Path")) & "\"
End if
AutoReName = UpFileObj.Form("AutoRename")
IsAddWaterMark = UpFileObj.Form("chkAddWaterMark")
If IsAddWaterMark <> "1" Then	'生成是否要添加水印标记
	IsAddWaterMark = "9"
End if

ReturnValue = CheckUpFile(FilePath,MaxFileSize,AllowFileExtStr,AutoReName,IsAddWaterMark)
if instr(ReturnValue,"Error")>0  then
%>
<script language="JavaScript">
	alert('<% = "以下文件上传失败，错误信息：\n" & ReturnValue %>');
	dialogArguments.location.reload();
	close();
</script>
<%
else
if not isNull(ReturnValue) and instr(ReturnValue,"*")>0 then
	Session("upfiles")=split(ReturnValue,"*")(1)
End if
%>
<script language="JavaScript">
	dialogArguments.location.reload();
	close();
</script>
<%
end if
Set UpFileObj=Nothing


Function CheckUpFile(Path,FileSize,AllowExtStr,AutoReName,IsAddWaterMark)
	Dim ErrStr,NoUpFileTF,FsoObj,FileName,FileExtName,FileContent,SameFileExistTF
	NoUpFileTF = True
	ErrStr = ""
	Set FsoObj = Server.CreateObject(G_FS_FSO)
	For Each FormName in UpFileObj.File
		SameFileExistTF = False
		FileName = UpFileObj.File(FormName).FileName
		If NoIllegalStr(FileName)=False Then
			ErrStr=ErrStr&"文件：上传被禁止！\n"
		End If
		FileExtName = UpFileObj.File(FormName).FileExt
		FileContent = UpFileObj.File(FormName).FileData
		'是否存在重名文件
		if UpFileObj.File(FormName).FileSize > 1 then
			NoUpFileTF = False
			ErrStr = ""
			if UpFileObj.File(FormName).FileSize > CLng(FileSize)*1024 then
				ErrStr = ErrStr & FileName & "文件:超过了限制，最大只能上传" & FileSize & "K的文件\n"
			end if
			if AutoRename = "0" then
				If FsoObj.FileExists(Path & FileName) = True  then
					ErrStr = ErrStr & FileName & "文件:存在同名文件\n"
				else
					SameFileExistTF = True
				end if
			else
				SameFileExistTF = True
			End If
			if CheckFileType(AllowExtStr,FileExtName) = False then
				ErrStr = ErrStr & FileName & "文件:不允许上传,上传文件类型有" + AllowExtStr + "\n"
			end if
			if ErrStr = "" then
				if SameFileExistTF = True then
					CheckUpFile=CheckUpFile&"*"&SaveFile(Path,FormName,AutoReName,IsAddWaterMark)
				else
					CheckUpFile=CheckUpFile&"*"&SaveFile(Path,FormName,"",IsAddWaterMark)
				end if
			else
				CheckUpFile ="Error:"& CheckUpFile & ErrStr
			end if
		end if
	Next
	Set FsoObj = Nothing
	if NoUpFileTF = True then
		CheckUpFile = "没有上传文件"
	end if
End Function

Function CheckFileType(AllowExtStr,FileExtName)
	Dim i,AllowArray
	AllowArray = Split(AllowExtStr,",")
	FileExtName = LCase(FileExtName)
	CheckFileType = False
	For i = LBound(AllowArray) to UBound(AllowArray)
		if LCase(AllowArray(i)) = LCase(FileExtName) then
			CheckFileType = True
		end if
	Next
	if FileExtName="asp" or FileExtName="asa" or FileExtName="aspx" or  FileExtName="php" or  FileExtName="php3" or  FileExtName="php4"  or  FileExtName="php5"then
		CheckFileType = False
	end if
End Function
Function DealExtName(Byval UpFileExt)
		If IsEmpty(UpFileExt) Then Exit Function
		DealExtName = Lcase(UpFileExt)
		DealExtName = Replace(DealExtName,Chr(0),"")
		DealExtName = Replace(DealExtName,".","")
		DealExtName = Replace(DealExtName,"'","")
		DealExtName = Replace(DealExtName,"asp","")
		DealExtName = Replace(DealExtName,"asa","")
		DealExtName = Replace(DealExtName,"aspx","")
		DealExtName = Replace(DealExtName,"cer","")
		DealExtName = Replace(DealExtName,"cdx","")
		DealExtName = Replace(DealExtName,"htr","")
		DealExtName = Replace(DealExtName,"php","")
End Function

Function NoIllegalStr(Byval FileNameStr)
	Dim Str_Len,Str_Pos
	Str_Len=Len(FileNameStr)
	Str_Pos=InStr(FileNameStr,Chr(0))
	If Str_Pos=0 or Str_Pos=Str_Len then
	 	NoIllegalStr=True
	Else
	 	NoIllegalStr=False
	End If
End function

Function SaveFile(FilePath,FormNameItem,AutoNameType,IsAddWaterMark)
	Dim FileName,FileExtName,FileContent,FormName,RandomFigure,SPicPath,TimeNameStr
	Randomize 
	RandomFigure = CStr(Int((99999 * Rnd) + 1))
	FileName = UpFileObj.File(FormNameItem).FileName
	FileExtName = UpFileObj.File(FormNameItem).FileExt
	FileExtName=DealExtName(FileExtName)
	FileContent = UpFileObj.File(FormNameItem).FileData
	If AutoNameType = "1" Then
		SPicPath = FilePath & "S_" & "副件" & FileName
		FileName = FilePath & "副件" & FileName
	elseif AutoNameType = "2" Then
		SPicPath = FilePath & "S_" & "1" & FileName 
		FileName = FilePath & "1" & FileName 
	elseif AutoNameType = "3" Then
		TimeNameStr = Year(Now())&Right("0"&Month(Now()),2)&Right("0"&Day(Now()),2)&Right("0"&Hour(Now()),2)&Right("0"&Minute(Now()),2)&Right("0"&Second(Now()),2)&RandomFigure
		SPicPath = FilePath & "S_" & TimeNameStr & "."&FileExtName
		FileName = FilePath & TimeNameStr & "."&FileExtName
	Else
		SPicPath = FilePath & "S_" & FileName
		FileName = FilePath&FileName
	End If
	UpFileObj.File(FormNameItem).SaveToFile FileName
	CreateThumbnailEx FileName,SPicPath  '生成缩略图
	If IsAddWaterMark = "1" Then   '在保存好的图片上添加水印
		AddWaterMark FileName
	End if
	SaveFile=mid(FileName,instrrev(FileName,"\")+1,len(FileName))
End Function
Set Conn = Nothing
%>
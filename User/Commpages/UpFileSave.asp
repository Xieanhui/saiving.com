<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="News_Upfile.asp" -->
<!--#include file="../../FS_Inc/WaterPrint_Function.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
'==========================
User_GetParm
getGroupIDinfo
dim p_FSO_,str_ShowPath_,UserFileSpace_s
str_ShowPath_ = Replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_USERFILES_DIR & "/" &Fs_User.UserNumber,"//","/")
Set p_FSO_ = Server.CreateObject(G_FS_FSO)
set UserFileSpace_s=p_FSO_.GetFolder(Server.MapPath(str_ShowPath_))
if UpfileSize="" or Isnull(UpfileSize) Then 'UpfileSize是从数据库里读出来的
		  	UpfileSize=2                          '默认是2M
		  else
		  	UpfileSize=Clng(UpfileSize)
		  end if
if UserFileSpace_s.size>=UpfileSize*1024*1024 then response.Write("您的空间不足"):response.end
set p_FSO_=nothing
if p_UpfileType="" or isnull(p_UpfileType) then response.Write("没有开放上传权限"):response.end
if p_UpfileSize <>"" then:MaxFileSize = clng(p_UpfileSize):Else:MaxFileSize =100:End if
if p_UpfileType <>"" then:AllowFileExtStr = p_UpfileType:Else:AllowFileExtStr = "jpg,gif,jpeg,png,bmp,txt,doc,rar":End if
Set UpFileObj = New UpFileClass
UpFileObj.GetData
FilePath=Server.MapPath(UpFileObj.Form("Path")) & "\"
AutoReName = UpFileObj.Form("AutoRename")
IsAddWaterMark = UpFileObj.Form("chkAddWaterMark")
If IsAddWaterMark <> "1" Then	'生成是否要添加水印标记
	IsAddWaterMark = "0"
End if
ReturnValue = CheckUpFile(FilePath,MaxFileSize,AllowFileExtStr,AutoReName,IsAddWaterMark)
if ReturnValue <> "" then
%>
<script language="JavaScript">
	alert('<% = "以下文件上传失败，错误信息：\n" & ReturnValue %>');
	dialogArguments.location.reload();
	close();
</script>
<%
else
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
					SaveFile Path,FormName,AutoReName,IsAddWaterMark
				else
					SaveFile Path,FormName,"",IsAddWaterMark
				end if
			else
				CheckUpFile = CheckUpFile & ErrStr
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
	if FileExtName="asp" or FileExtName="asa" or FileExtName="aspx" or  FileExtName="php" or  FileExtName="php3" or  FileExtName="php4"  or  FileExtName="php5" then
		CheckFileType = False
	end if
End Function
Function DealExtName(Byval UpFileExt)
		If IsEmpty(UpFileExt) Then Exit Function
		DealExtName = Lcase(UpFileExt)
		DealExtName = Replace(DealExtName,chr(0),"")
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
	Dim FileName,FileExtName,FileContent,FormName,RandomFigure
	Randomize 
	RandomFigure = CStr(Int((99999 * Rnd) + 1))
	FileName = DealExtName(UpFileObj.File(FormNameItem).FileName)
	FileExtName = UpFileObj.File(FormNameItem).FileExt
	FileExtName=DealExtName(FileExtName)
	FileContent = UpFileObj.File(FormNameItem).FileData
	'If AutoNameType = "1" Then
		'FileName = FilePath & "副件" & FileName
	If AutoNameType = "2" Then
		'FileName = FilePath & "1" & FileName 
		FileName = FilePath & Year(Now())&"_"&Right("0"&Month(Now()),2)&"_"&Right("0"&Day(Now()),2)&"_"&Right("0"&Hour(Now()),2)&"_"&Right("0"&Minute(Now()),2)&"_"&Right("0"&Second(Now()),2)&RandomFigure&"."&FileExtName
		
	Elseif AutoNameType = "3" Then
		FileName = FilePath & Year(Now())&Right("0"&Month(Now()),2)&Right("0"&Day(Now()),2)&Right("0"&Hour(Now()),2)&Right("0"&Minute(Now()),2)&Right("0"&Second(Now()),2)&RandomFigure&"."&FileExtName
	Else
		FileName = FilePath&FileName
	End If
	UpFileObj.File(FormNameItem).SaveToFile FileName
	
	If IsAddWaterMark = "1" Then   '在保存好的图片上添加水印
		AddWaterMark FileName
	End if
	if Lcase(FileExtName)="jpg" or Lcase(FileExtName)="gif" or Lcase(FileExtName)="png" or Lcase(FileExtName)="jpeg" or Lcase(FileExtName)="bmp" then
		If CheckFileTypeimage(Lcase(trim(FileName)))=false then
			response.write "错误的图像格式,请不要上传图片木马!"
			Set fsos = CreateObject(G_FS_FSO)
			if fsos.FileExists(FileName) = true then fsos.DeleteFile FileName
			set ficn=nothing
			set fsos=nothing
			response.end
		end if
	end if
End Function


function CheckFileTypeimage(filename)
	const adTypeBinary=1
	dim jpg(1):jpg(0)=CByte(&HFF):jpg(1)=CByte(&HD8)
	dim bmp(1):bmp(0)=CByte(&H42):bmp(1)=CByte(&H4D)
	dim png(3):png(0)=CByte(&H89):png(1)=CByte(&H50):png(2)=CByte(&H4E):png(3)=CByte(&H47)
	dim gif(5):gif(0)=CByte(&H47):gif(1)=CByte(&H49):gif(2)=CByte(&H46):gif(3)=CByte(&H39):gif(4)=CByte(&H38):gif(5)=CByte(&H61)
	on error resume next
	CheckFileTypeimage=false
	filename=LCase(filename)
	dim fstream,fileExt,stamp,i
	fileExt=mid(filename,InStrRev(filename,".")+1)
	set fstream=Server.createobject(G_FS_STREAM)
	fstream.Open
	fstream.Type=adTypeBinary
	fstream.LoadFromFile filename
	fstream.position=0
	select case fileExt
		case "jpg","jpeg"
			stamp=fstream.read(2)
			for i=0 to 1
				if ascB(MidB(stamp,i+1,1))=jpg(i) then CheckFileTypeimage=true else CheckFileTypeimage=false
			next
		case "gif"
			stamp=fstream.read(6)
			for i=0 to 5
				if ascB(MidB(stamp,i+1,1))=gif(i) then CheckFileTypeimage=true else CheckFileTypeimage=false
			next
		case "png"
			stamp=fstream.read(4)
			for i=0 to 3
				if ascB(MidB(stamp,i+1,1))=png(i) then CheckFileTypeimage=true else CheckFileTypeimage=false
			next
		case "bmp"
			stamp=fstream.read(2)
			for i=0 to 1
				if ascB(MidB(stamp,i+1,1))=bmp(i) then CheckFileTypeimage=true else CheckFileTypeimage=false
			next
	end select
	fstream.Close
	set fseteam=nothing
	if err.number<>0 then CheckFileTypeimage=false
end function

Set Conn = Nothing
%>
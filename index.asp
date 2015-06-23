<% Option Explicit %>
<!--#include file="FS_Inc/Const.asp" -->
<!--#include file="FS_Inc/Function.asp" -->
<!--#include file="FS_InterFace/MF_Function.asp" -->
<!--#include file="FS_InterFace/Dynamic_Function.asp" -->
<%
''动态的首页，整站通用，可以用于人才，供求，房产以及一些需要自动更新的系统。
''使用方法,IIS中将index.asp的优先级别设置得高一些,这样会自动更新。
''或者访问/index.asp?isNews=1 则可以立即手动更新。
''或者手工删除前台的静态文件如index.html，根据后台设置的来
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
'''更新单位秒 RefreshTime = 12*60*60 为1天12小时
Dim RefreshTime,isNew
RefreshTime = 180
''是否不生成静态 0 为生成 1 为不生成
isNew = request.QueryString("isNew")
if not isnumeric(isNew) then isNew = 0  else isNew=cint(isNew) end if

Dim conn,Str_SD_Templet,User_Conn,tmpRS_SD,MF_Index_File_Name,Dynamic_HTML

MF_Default_Conn

set tmpRS_SD = Conn.execute("select top 1 MF_Index_Templet,MF_Index_File_Name,MF_Index_Refresh from FS_MF_Config")
if not tmpRS_SD.eof then
	Str_SD_Templet = tmpRS_SD("MF_Index_Templet")
	If G_VIRTUAL_ROOT_DIR<>"" Then
		Str_SD_Templet="/"&G_VIRTUAL_ROOT_DIR&Str_SD_Templet
	End If
	MF_Index_File_Name = tmpRS_SD("MF_Index_File_Name")
	RefreshTime = CInt(tmpRS_SD("MF_Index_Refresh"))*60
end if
tmpRS_SD.close : set tmpRS_SD=nothing

if MF_Index_File_Name="" or isnull(MF_Index_File_Name) then
	MF_Index_File_Name = "index.html"   ''首页静态名称 为空则不生成静态
end if

if isNew=0 then
	''===================================================
	 ''判断MF_Index_File_Name是否合法，如果合法则跳到MF_Index_File_Name
	Dim IsExistsHtmPage_
	IsExistsHtmPage_ = IsExistsHtmPage(RefreshTime) ''12个小时更新一次
	if IsExistsHtmPage_ Then
		Conn.close
		response.Redirect(MF_Index_File_Name)
		response.End()
	end if
end If
	''===================================================

if Str_SD_Templet="" then
	Str_SD_Templet="/Job/index.htm"
	If G_VIRTUAL_ROOT_DIR<>"" Then
		If G_TEMPLETS_DIR<>"" Then
			Str_SD_Templet="/"&G_VIRTUAL_ROOT_DIR&"/"&G_TEMPLETS_DIR&Str_SD_Templet
		Else
			Str_SD_Templet="/"&G_VIRTUAL_ROOT_DIR&Str_SD_Templet
		End If
	Else
		If G_TEMPLETS_DIR<>"" Then
			Str_SD_Templet="/"&G_TEMPLETS_DIR&Str_SD_Templet
		Else
			Str_SD_Templet=Str_SD_Templet
		End If
	End If
end if

MF_User_Conn

''读取动态信息

Dynamic_HTML = Get_Dynamic_Refresh_Content(Str_SD_Templet,"","MF",0,"")

if isNew = 0 then
	''生成静态
	Call WriteLocalFile(MF_Index_File_Name,Dynamic_HTML)
	''跳转到静态
	response.Redirect(MF_Index_File_Name)
else
	response.Write(	Dynamic_HTML )
end if

''-----------------
'检查MF_Index_File_Name是否存在 并且有内容，和最近几天的日期，为true则跳转到MF_Index_File_Name，否则则返回false.
Function IsExistsHtmPage(HowSecond)
	'HowSecond = 12*60*60 为1天12小时
	if MF_Index_File_Name = "" then IsExistsHtmPage=False : exit function
	if HowSecond="" then HowSecond=180
	Dim Fso,MyFile,PhFileName,isTrue
	PhFileName = MF_Index_File_Name
	isTrue = False
	Set Fso = CreateObject(G_FS_FSO)
	If Fso.FileExists(server.MapPath(PhFileName)) Then
		set MyFile = Fso.GetFile(server.MapPath(PhFileName))
		if (MyFile.Size > 10 and datediff("s",MyFile.DateLastModified,now()) < HowSecond) Or HowSecond<0 then
			if request.QueryString("isNew")="1" then
				MyFile.Delete(True)
				isTrue = False
			else
				isTrue = True
			end if
		else
			MyFile.Delete(True)
			isTrue = False
		end if
		set MyFile = nothing
	End If
	Set Fso = Nothing
	IsExistsHtmPage = isTrue
End Function

'将内容写入本地文件
Sub WriteLocalFile(PhFileName,strFile)
	Dim Fso, MyFile
	Set Fso = CreateObject(G_FS_FSO)
	Set MyFile = fso.CreateTextFile(server.MapPath(PhFileName), True)
	MyFile.Write strFile
	Set MyFile = Nothing
	Set Fso = Nothing
End Sub

%>






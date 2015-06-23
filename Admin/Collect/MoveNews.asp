<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../News/lib/cls_main.asp" -->
<%
Dim Conn,CollectConn,p_CName,p_ID,p_Action,p_ClassID,sRootDir,str_CurrPath,p_Templet
Dim p_File_Ext_Name,p_Save_Path
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
MF_Default_Conn
MF_Collect_Conn
MF_Session_TF

p_CName = Request.QueryString("CName")
p_ID = Request("ID")
p_Action = Request.QueryString("Action")

if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/" & G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table height="120" width="60%" border="0" align="center">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<%
	Dim p_Return_Str,p_Delete_ID
	if NewsToSystem("FS_NS_News",p_ID) then
		p_Return_Str = "转移成功"
		p_Delete_ID = Replace(p_ID,"***",",")
		if p_ID = "all" then
			CollectConn.Execute("Update FS_News Set History=1 where 1=1")
		else
			CollectConn.Execute("Update FS_News Set History=1 where ID in (" & FormatIntArr(p_Delete_ID) & ")")
		end if
	else
		p_Return_Str = "转移失败"
	end if
%>
<table width="60%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback">
    <td><% = p_Return_Str %></td>
  </tr>
  <tr class="hback">
    <td height="30"><div align="center">
        <input type="button" name="Submit" onClick="location='Check.asp';" value=" 返 回 ">
      </div></td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script> <!--用于摸版选择-->
<%
'参数f_Object_Table为目标数据库中的表名
'参数f_Source_ID为采集库的新闻表(FS_News)中的ID集合，ID之间以***分割
Function NewsToSystem(f_Object_Table,f_Source_ID)
	Dim f_Field_Array,f_Source_Sql,f_Object_Sql,f_Collect_RS,f_System_RS,f_i,TempNewsID,f_System_RS_Pop,OldID,Fs_news
	Dim NewsSql_Arr(),Str_Temp_Flag,temp_j,StrSql
	if f_Source_ID = "" then Exit Function
	Set Fs_news = New Cls_News
	Fs_News.GetSysParam()
	If Not Fs_news.IsSelfRefer Then
		p_File_Ext_Name = "html"
		p_Save_Path = "/" & Year(Now) & "-" & Month(Now) & "-" & Day(Now)
	else
		p_File_Ext_Name = Fs_News.fileExtName
		p_Save_Path = Fs_news.SaveNewsPath(Fs_news.fileDirRule)
	end if
	'-----2006-12-07 by ken 采集数据转移到主数据库时候，设置生成静态文件扩展名
	If p_File_Ext_Name <> "html" Then
		If CInt(p_File_Ext_Name) = 0 then
			p_File_Ext_Name = "html"
		ElseIf CInt(p_File_Ext_Name) = 1 then
			p_File_Ext_Name = "htm"
		ElseIf CInt(p_File_Ext_Name) = 2 then
			p_File_Ext_Name = "shtml"
		ElseIf CInt(p_File_Ext_Name) = 3 then
			p_File_Ext_Name = "shtm"
		ElseIf CInt(p_File_Ext_Name) = 4 then
			p_File_Ext_Name = "asp"
		Else
			p_File_Ext_Name = "html"
		End If				
	End If
	'------End-------------------------------------------------------------	
	f_Source_ID = Replace(f_Source_ID,"***",",")
	'-----2007-01-25 Edit By ken For CollectNews To SysTem
	Set f_Collect_RS = Server.CreateObject(G_FS_RS)
	if f_Source_ID = "all" then
		f_Source_Sql = "Select * from FS_News where 1=1 Order By ID Desc"
	else
		f_Source_Sql = "Select * from FS_News where ID in (" & FormatIntArr(f_Source_ID) & ") Order By ID Desc"
	end if
	f_Collect_RS.Open f_Source_Sql,CollectConn,1,1
	'--------------------------------------------
	ReDim NewsSql_Arr(0)
	Str_Temp_Flag=True
	While Not f_Collect_RS.Eof	
		StrSql="INSERT INTO FS_NS_News([NewsID],[PopId],[ClassID],[NewsTitle],[isShowReview],[Content],[Templet],[Source],[Author],[SaveNewsPath],[FileName],[FileExtName],[NewsProperty],[isLock],[addtime],[isPicNews],[NewsPicFile],[NewsSmallPicFile]) VALUES ("
		TempNewsID=GetRamCode(15)
		StrSql=StrSql & "'" & NoSqlHack(TempNewsID) & "'"
		StrSql=StrSql & ",0"
		StrSql=StrSql & ",'" & NoSqlHack(GetNewsInfoBySiteID(f_Collect_RS("SiteID"),"ClassID")) & "'"
		StrSql=StrSql & ",'"&NoSqlHack(f_Collect_RS("Title"))&"'"
		StrSql=StrSql & ","&NoSqlHack(NUllToStr(f_Collect_RS("ReviewTF")))&""
		StrSql=StrSql & ",'"&Replace(f_Collect_RS("Content"),"'","''")&"'"
		StrSql=StrSql & ",'"&NoSqlHack(GetNewsInfoBySiteID(f_Collect_RS("SiteID"),"Temp"))&"'"
		StrSql=StrSql & ",'"&NoSqlHack(left(f_Collect_RS("Source"),50))&"'"
		StrSql=StrSql & ",'"&NoSqlHack(Left(f_Collect_RS("Author"),50))&"'"
		StrSql=StrSql & ",'"&NoSqlHack(Fs_news.SaveNewsPath(Fs_news.fileDirRule))&"'"
		'------
		OldID = Fs_News.strFileNameRule(Fs_News.fileNameRule,0,0)
		if instr(OldID,"自动编号ID") > 0 then OldID = Replace(OldID,"自动编号ID",TempNewsID)
		if instr(OldID,"唯一NewsID") > 0 then OldID = Replace(OldID,"唯一NewsID",TempNewsID)
		'------
		StrSql=StrSql & ",'"&OldID&"'"
		StrSql=StrSql & ",'"&p_File_Ext_Name&"'"
		StrSql=StrSql & ",'0,1,1,0,1,0,0,0,1,0,0'"
		If f_Collect_RS("isLock") Then
			StrSql=StrSql & ",1"
		Else
			StrSql=StrSql & ",0"
		End If	
		StrSql=StrSql & ",'"&f_Collect_RS("AddDate")&"'"
		'===2007-02-25 Edit By Ken======
		If GetCeSitePicTF(f_Collect_RS("SiteID")) = True Then
			If ContentInnerPicTF(Replace(f_Collect_RS("Content"),"'","''"),"TF") = True Then
				StrSql = StrSql & ",1"
				StrSql = StrSql & ",'" & ContentInnerPicTF(Replace(f_Collect_RS("Content"),"'","''"),"PicUrl") & "'"
				StrSql = StrSql & ",'" & ContentInnerPicTF(Replace(f_Collect_RS("Content"),"'","''"),"PicUrl") & "'"
			Else
				StrSql = StrSql & ",0"
				StrSql = StrSql & ",''"
				StrSql = StrSql & ",''"
			End If
		Else
			StrSql = StrSql & ",0"
			StrSql = StrSql & ",''"
			StrSql = StrSql & ",''"
		End If
		'====End=====================			
		StrSql=StrSql & ")"
		If Str_Temp_Flag Then
			NewsSql_Arr(Ubound(NewsSql_Arr))=StrSql
			Str_Temp_Flag=False
		Else
			ReDim Preserve NewsSql_Arr(Ubound(NewsSql_Arr)+1)
			NewsSql_Arr(Ubound(NewsSql_Arr))=StrSql
		End If
		f_Collect_RS.movenext
	Wend
	On Error Resume Next
	For temp_j=Lbound(NewsSql_Arr) to Ubound(NewsSql_Arr)
		If NewsSql_Arr(temp_j) <> "" Then
			Conn.Execute(NewsSql_Arr(temp_j))
		End If
	Next
	f_Collect_RS.Close
	Set f_Collect_RS = Nothing
	Set Fs_news = Nothing
	NewsToSystem = True
End Function

'----
Function NUllToStr(num)
	If IsNull(num) Or num = "" Then
		NUllToStr = 0
	Else
		If Not IsNumeric(num) Then
			NUllToStr = 0
		Else
			NUllToStr = Cint(num)
		End If	
	End if
End Function

'===========================================================
'判断传入的字符传中是否包含本地图片并取得此图片地址
'===========================================================
Function ContentInnerPicTF(StrCon,ReturnTF)
	Dim ConStr,Re,InnerPicAll,FistPicUrl,PicUrlStr
	ConStr = StrCon & ""
	Set Re = New RegExp
	Re.IgnoreCase = True
	Re.Global = True
	Re.Pattern = "(src\S+\.{1}(gif|jpg|png)(""|\'|>|\s)?)"
	InnerPicAll = ""
	Set InnerPicAll = Re.Execute(ConStr)
	Set Re = Nothing
	FistPicUrl = ""
	For Each PicUrlStr in InnerPicAll
		FistPicUrl = Replace(Replace(Replace(PicUrlStr,"src=",""),"'",""),"""","")
		If LCase(Left(FistPicUrl,Len(sRootDir))) = LCase(sRootDir) Then
			FistPicUrl = Mid(FistPicUrl,Len(sRootDir)+1)
		End If
		Exit For
	Next
	If ReturnTF = "TF" Then
		If FistPicUrl <> "" And (Not IsNull(FistPicUrl)) then
			ContentInnerPicTF = True
		Else
			ContentInnerPicTF = False	
		End If
	ElseIf ReturnTF = "PicUrl" Then
		If FistPicUrl <> "" And (Not IsNull(FistPicUrl)) then
			ContentInnerPicTF = FistPicUrl
		End If
	End If					
End Function

'===========================================================
'判断传入的采集站点设置属性
'===========================================================
Function GetCeSitePicTF(SiteID)
	Dim GetSiteRs
	IF SiteID = "" Then : GetCeSitePicTF = False : Exit Function
	SiteID = Clng(SiteID)
	Set GetSiteRs = CollectConn.ExeCute("Select IsAutoPicNews From FS_Site Where ID = " & CintStr(SiteID) & " And IsLock = 0")
	If GetSiteRs.Eof Then
		GetCeSitePicTF = False
	Else
		If GetSiteRs(0) = 1 Then
			GetCeSitePicTF = True
		Else
			GetCeSitePicTF = False
		End If
	End If
	GetSiteRs.Close : Set GetSiteRs = NoThing			
End Function

Function GetNewsInfoBySiteID(SiteID,Act)
	Dim GetSiteRs
	IF SiteID = "" Or IsNull(SiteID) Or NOt IsNumeric(SiteID) Then
		If Act = "ClassID" Then
			GetNewsInfoBySiteID = 0
		Else
			GetNewsInfoBySiteID = "/" & G_TEMPLETS_DIR & "/NewsClass/new.htm"
		End IF	
	End If
	Set GetSiteRs = CollectConn.ExeCute("Select ToClassID,NewsTemplets From FS_Site Where ID = " & CintStr(SiteID) & " And IsLock = 0")
	If GetSiteRs.Eof Then
		If Act = "ClassID" Then
			GetNewsInfoBySiteID = 0
		Else
			GetNewsInfoBySiteID = "/" & G_TEMPLETS_DIR & "/NewsClass/new.htm"
		End IF
	Else
		If Act = "ClassID" Then
			GetNewsInfoBySiteID = GetSiteRs(0)
		Else
			GetNewsInfoBySiteID = GetSiteRs(1)
		End IF
	End If
	GetSiteRs.Close : Set GetSiteRs = NoThing
End Function


Set CollectConn = Nothing
Set Conn = Nothing
%>







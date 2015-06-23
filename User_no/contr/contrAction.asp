<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
'-----------------------------------------
response.expires=0 
response.addHeader "pragma" , "no-cache" 
response.addHeader "cache-control" , "private" 
response.Charset="gb2312"
'---------------------------------------------
Dim action,id,contrRs,sqlstatement,contrPoint,contrMoney,Rs
Dim ContID,ContSytle,ContTitle,SubTitle,ContContent,AddTime,PassTime,ClassID,MainID,KeyWords,IsPublic,InfoType,UserNumber,OtherContent,IsLock,isTF,Hits,AdminLock,PicFile,TempletID,FileName,FileExeName,AuditTF,Untread,contr_type,ConstrCheck
action=request.QueryString("action")
id=DelHeadAndEndDot(CintStr(Request.QueryString("id")))
Set Rs=User_Conn.execute("Select top 1 contrPoint,contrMoney from FS_ME_SysPara")
if Rs.eof then
	Response.Redirect("../error.asp?ErrCodes=<li>请先设置会员参数</li>&ErrorUrl=")
ELse
	contrPoint=Rs("contrPoint")
	if not isnumeric(contrPoint) then contrPoint=0
	contrMoney=Rs("contrMoney")
	if not isnumeric(contrMoney) then contrMoney=0
End if
Set Rs=Conn.execute("Select top 1 isConstrCheck from FS_NS_SysParam")
if Rs.eof then
	ConstrCheck=1
ELse
	ConstrCheck=Rs("isConstrCheck")
End if
Rs.close
Set Rs=nothing
if action="lock" then
	User_Conn.execute("Update FS_ME_InfoContribution set isLock=1 where contID in("& FormatIntArr(id) &") and UserNumber='"&Session("FS_UserNumber")&"'")
	Response.Write("<a href='解锁' onClick=""contrAction('unlock','"&id&"','span_lock_"&id&"');return false;""><font color='red'>解锁</font></a>")
	response.End()
Elseif action="unlock" then
	User_Conn.execute("Update FS_ME_InfoContribution set isLock=0 where contID in("& FormatIntArr(id) &") and UserNumber='"&Session("FS_UserNumber")&"'")
	Response.Write("<a href='锁定' onClick=""contrAction('lock','"&id&"','span_lock_"&id&"');return false;"">锁定</a>")
	response.End()
Elseif action="delete" then
	User_Conn.execute("Delete from FS_ME_InfoContribution where contID in("& FormatIntArr(id) &") and UserNumber='"&Session("FS_UserNumber")&"'")
	Response.Write("ok")
	response.End()
Else
	Set contrRs=server.CreateObject(G_FS_RS)
	if id<>"" And action="edit" then
		sqlstatement="select ContID,ContSytle,ContTitle,SubTitle,ContContent,AddTime,PassTime,ClassID,MainID,KeyWords,IsPublic,InfoType,UserNumber,OtherContent,IsLock,isTF,Hits,AdminLock,PicFile,TempletID,FileName,FileExeName,AuditTF,Untread,type,NewsID from FS_ME_InfoContribution where ContID="&CintStr(ID)&" and UserNumber='"&Session("FS_UserNumber")&"'"
	Elseif action="add" then
		sqlstatement="Select * From FS_ME_InfoContribution where 1=2"
	End if
	contrRs.open sqlstatement,User_Conn,1,3
	if action="add" or contrRs.eof then
		contrRs.addnew
	End if
	ContSytle=NoSqlHack(request.Form("sel_ContSytle"))
	ContTitle=FiltBad(NoSqlHack(request.Form("txt_ContTitle")))
	SubTitle=FiltBad(NoSqlHack(request.Form("txt_subTitle")))
	ContContent=FiltBad(NoSqlHack(request.Form("txt_Content")))
	AddTime=dateValue(Now)
	ClassID=NoSqlHack(request.Form("txt_ClassID"))
	if trim(ClassID)="" then ClassID=0
	MainID=NoSqlHack(request.Form("txt_mainClassID"))
	if trim(MainID)="" then MainID=0
	KeyWords=NoSqlHack(request.Form("txt_KeyWords"))
	IsPublic=NoSqlHack(request.Form("rad_IsPublic"))
	InfoType=NoSqlHack(request.Form("sel_InfoType"))
	UserNumber=Session("FS_UserNumber")
	OtherContent=NoSqlHack(request.Form("txt_OtherContent"))
	IsLock=NoSqlHack(request.Form("rad_IsLock"))
	isTF=NoSqlHack(request.Form("rad_isTF"))
	PicFile=NoSqlHack(request.Form("txt_img"))
	If ConstrCheck=1 Then
		AuditTF=0
	Else
		AuditTF=1
	End If
	contr_type=request.Form("sel_type")
	'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	'是否需要审核
	If ConstrCheck=0 and IsPublic="1" and IsLock="0" and ID="0" Then
		Dim NewsTemplet,sql_news_cmd,news_Rs,Str_NS_Config

		if trim(MainID)="" then
			Response.Redirect("lib/error.asp?ErrCodes=<li>请选择栏目</li>")
			Response.End()
		End if
		Response.Write "SELECT NewsTemplet FROM FS_NS_NewsClass WHERE ClassID='"&NoSqlHack(MainID)&"'"
		Set Str_NS_Config=Conn.Execute("SELECT NewsTemplet FROM FS_NS_NewsClass WHERE ClassID='"&NoSqlHack(MainID)&"'")
		NewsTemplet=Str_NS_Config(0)
		Str_NS_Config.Close:Set Str_NS_Config=Nothing
		Set news_Rs= Server.CreateObject(G_FS_RS)
		sql_news_cmd="select  ID,NewsID,NewsTitle,CurtTitle,Content,Keywords,Author,classID,isPicNews,NewsPicFile,SaveNewsPath,FileName,FileExtName,addtime,Templet,NewsProperty,Editor from FS_NS_News"
		news_Rs.open sql_news_cmd,Conn,1,3
		news_Rs.addNew
		news_Rs("NewsID")=GetRamCode(15)
		news_Rs("NewsTitle")=ContTitle
		news_Rs("CurtTitle")=SubTitle
		news_Rs("Content")=ContContent
		news_Rs("Keywords")=Keywords
		news_Rs("Author")=session("FS_UserName")
		news_Rs("Editor")="" 
		news_Rs("NewsProperty")="0,1,1,1,0,0,0,0,1,0,0"
		news_Rs("Templet")=NewsTemplet
		news_Rs("classID")=MainID
		if trim(PicFile)<>"" then
			news_Rs("NewsPicFile")=PicFile
			news_Rs("isPicNews")=1
			news_Rs("NewsSmallPicFile")=PicFile
		End if
		news_Rs("SaveNewsPath") = "/Constr_"&year(now)&month(now)&day(now)&""
		news_Rs("FileName")=GetRamCode(8)
		news_Rs("FileExtName")="html"
		news_Rs("AddTime")=Now()
		news_Rs.update
		if G_IS_SQL_DB = 0 then'是否是sqlserver
			contrRs("NewsId")=news_Rs("ID")
		Else
			contrRs("NewsId")=Conn.Execute("SELECT IDENT_CURRENT('FS_NS_News')")(0)
		End if
		news_Rs.close
		Set news_Rs=Nothing
	End If
	'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	contrRs("ContSytle")=ContSytle
	contrRs("ContTitle")=ContTitle
	contrRs("SubTitle")=SubTitle
	contrRs("ContContent")=ContContent
	contrRs("AddTime")=AddTime
	if ClassID<>"" then contrRs("ClassID")=ClassID
	if MainID<>"" then
		Dim tmpRs
		Set tmpRs=Conn.execute("select id from FS_NS_NewsClass where classid='"&NoSqlHack(MainID)&"'") 
		if not tmpRs.eof then
			contrRs("MainID")=tmpRs("id")
		End if
		tmpRs.close
		Set tmpRs=nothing
	End if
	contrRs("KeyWords")=KeyWords
	contrRs("IsPublic")=IsPublic
	contrRs("InfoType")=InfoType
	contrRs("UserNumber")=UserNumber
	contrRs("OtherContent")=OtherContent
	contrRs("IsLock")=IsLock
	contrRs("isTF")=isTF
	contrRs("PicFile")=PicFile
	contrRs("AuditTF")=AuditTF
	contrRs("type")=contr_type
	contrRs.update
	contrRs.close
	if action="add" then
		User_Conn.execute("Update FS_ME_Users set Integral=Integral+"&Clng(contrPoint)&" , FS_Money=FS_Money+"&Clng(contrMoney)&",ConNumber=ConNumber+1 where usernumber='"&usernumber&"'")
	End if
	if err.number<>0 then
		Response.Redirect("../lib/error.asp?ErrCodes="&err.description&"&ErrorUrl=")
	Else
		Response.Redirect("../lib/success.asp?ErrCodes=操作成功！&ErrorUrl=../contr/contrManage.asp")
	End if
End if
Conn.close
User_Conn.close
Set Conn=nothing
Set contrRs=nothing

%>

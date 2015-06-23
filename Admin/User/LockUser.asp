<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
	Dim Conn,User_Conn,strShowErr,GetUser,str_tmp
	MF_Default_Conn
	MF_User_Conn
	'session判断
	MF_Session_TF 
	set GetUser= Server.CreateObject(G_FS_RS)
	GetUser.open "select UserNumber,isLock From FS_ME_Users where UserNumber='"&NoSqlHack(Request.QueryString("UserNumber"))&"'",User_Conn,1,3
	if GetUser.eof then
		GetUser.close:set GetUser=nothing
		set conn=nothing
		set user_conn=nothing
		Response.Write("<script>alert(""找不到此用户"");window.close();</script>")
		Response.End
	else
		if NoSqlHack(Request.QueryString("action"))="Lock" then
			GetUser("isLock")=1
			str_tmp="锁定"
			Call MF_Insert_oper_Log("解锁用户","用户编号:("& NoSqlHack(Request.QueryString("UserNumber"))&")",now,session("admin_name"),"ME")
		elseif NoSqlHack(Request.QueryString("action"))="UnLock" then
			GetUser("isLock")=0
			str_tmp="解锁"
			Call MF_Insert_oper_Log("锁定用户","用户编号:("& NoSqlHack(Request.QueryString("UserNumber"))&")",now,session("admin_name"),"ME")
		end if
		GetUser.update
		GetUser.close:set GetUser=nothing
		set conn=nothing
		set user_conn=nothing
		Response.Write("<script>alert(""用户["&str_tmp&"]操作成功"");window.close();</script>")
		Response.End
	end if
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






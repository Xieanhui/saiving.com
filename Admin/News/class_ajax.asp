
<% Option Explicit %>
<%
Response.Charset="GB2312"
%>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%


	dim classID,Fs_news,Conn,User_Conn
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF 
	set Fs_news = new Cls_News
	
	classID=request.QueryString("classid")
	Response.Write(Fs_news.GetChildNewsList(classID,request.QueryString("cc")))
	response.End()
%>

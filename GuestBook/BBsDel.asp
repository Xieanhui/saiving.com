<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp"-->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
MF_Default_Conn
Dim Conn,User_Conn
Dim Configobj,PageS,sql,MSTitle,ShowIP
Set Configobj=server.CreateObject (G_FS_RS)
sql="select ID,Title,IPShow,IsUser,IsAut,PageSize,Style From FS_WS_Config"
configobj.open sql,Conn,1,1
if not configobj.eof then
PageS=configobj("PageSize")
MSTitle=configobj("Title")
ShowIP=configobj("IPShow")
end if
set configobj=nothing
Dim BBsID,NoteID,ClassName,ClassID,Pag,NoteTilte,ShowRs,url
if NoSqlHack(Request.querystring("Act"))="SinglDel" then
	NoteTilte=NoSqlHack(Request.QueryString("NoteTilte"))
	BBsID=NoSqlHack(Request.QueryString("BBSID"))
	NoteID=NoSqlHack(Request.QueryString("NoteID"))
	ClassName=NoSqlHack(Request.QueryString("ClassName"))
	ClassID=NoSqlHack(Request.QueryString("ClassID"))
	Pag=NoSqlHack(Request.QueryString("Page"))
	Set ShowRs=Conn.execute("Delete  From FS_WS_BBS Where ID="&CintStr(BBSID)&" and User='"&session("FS_UserName")&"'")
	Conn.execute("Delete  From FS_WS_BBS Where ParentID='"&NoSqlHack(BBSID)&"' and User='"&session("FS_UserName")&"'")
	url="ShowNote.asp?ClassName="&ClassName&"&NoteID="&NoteID&"&ClassID="&ClassID&"&Page="&Pag&""
	Response.write("<script>location.href='"&url&"'</script>")
	Response.end
end if
Set Conn=nothing
%>







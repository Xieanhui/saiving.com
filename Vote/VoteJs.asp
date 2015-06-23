<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<%session.CodePage="936"%>
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
response.Charset = "gb2312"
Dim Steps,TID,Cookie_Domain,OutHtmlID,PicW
''前台页面,由JS调用得到 调用该文件必须给定一些参数.
TID = NoSqlHack(request.QueryString("TID"))
OutHtmlID = NoSqlHack(request.QueryString("InfoID"))
PicW = NoSqlHack(request.QueryString("PicW"))
if OutHtmlID = "" then OutHtmlID = "Vote_HTML_ID_" & TID
if TID = "" or not isnumeric(TID) then response.Write("document.writeln('内部错误:调用时,TID必须提供.\n');"&vbNewLine) : response.End()
if PicW = "" or not isnumeric(PicW) then PicW = 60

Dim Conn
MF_Default_Conn

Cookie_Domain = Get_MF_Domain()
if Cookie_Domain="" then 
	Cookie_Domain = "localhost"
else
	if left(lcase(Cookie_Domain),len("http://"))="http://" then Cookie_Domain = mid(Cookie_Domain,len("http://")+1)
	if right(Cookie_Domain,1)="/" then Cookie_Domain = mid(Cookie_Domain,1,len(Cookie_Domain) - 1)
end if	
 
response.Write("function f_FirstGetVote_"&OutHtmlID&"() {new Ajax.Updater('"&OutHtmlID&"', 'http://"&Cookie_Domain&"/Vote/Index.asp?no-cache='+Math.random() , {method: 'get', parameters: 'TID="&TID&"&InfoID="&OutHtmlID&"&PicW="&PicW&"' });} "&vbNewLine)
response.Write("setTimeout('f_FirstGetVote_"&OutHtmlID&"()',200);"&vbNewLine)
Conn.close
Set Conn=Nothing
%>
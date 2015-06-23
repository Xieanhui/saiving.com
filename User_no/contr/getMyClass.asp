<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
Response.Charset="GB2312"
Dim classRs,classid,classIndex
classid=CintStr(request.QueryString("classid"))
classIndex=CintStr(trim(request.QueryString("index")))
if classid="" then response.End()
Set classRs=User_Conn.execute("select ClassID,ClassCName,UserNumber from FS_ME_InfoClass where UserNumber='"&session("FS_UserNumber")&"' and parentid="&classid)'ID or ClassID?
Response.Write("<select name='sel_myclass_"&(Cint(classIndex)+1)&"' onChange=""getMyChildClass(this,"&(CintStr(classIndex)+1)&")"">"&vbcrlf)
Response.Write("<option value='myclass'>"&(Cint(classIndex)+1)&"¼¶×¨À¸</option>"&vbcrlf)
while not classRs.eof
	Response.Write("<option value='"&classRs("Classid")&"'>"&classRs("ClassCName")&"</option>"&vbcrlf) 
	classRs.movenext
Wend	
Response.Write("</select>"&vbcrlf)
Conn.close
User_Conn.close
Set User_Conn=nothing
Set Conn=nothing
%>
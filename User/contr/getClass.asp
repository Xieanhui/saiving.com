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
classIndex=CintStr(request.QueryString("index"))
if classid="" then response.End()
Set classRs=Conn.execute("select id,ClassName,classid from FS_NS_NewsClass where isConstr=1 and ParentID=(select ClassID from FS_NS_NewsClass where id="&CintStr(classid)&")")'ID or ClassID?
Response.Write("<select name='sel_class_"&(Cint(classIndex)+1)&"' onChange=""getChildClass(this,"&(Cint(classIndex)+1)&")"">"&vbcrlf)
Response.Write("<option value='class'>"&(Cint(classIndex)+1)&"����Ŀ</option>"&vbcrlf)
while not classRs.eof
	Response.Write("<option value='"&classRs("id")&"'>"&classRs("ClassName")&"</option>"&vbcrlf) 
	classRs.movenext
Wend	
Response.Write("</select>"&vbcrlf)
Conn.close
User_Conn.close
Set User_Conn=nothing
Set Conn=nothing
%>

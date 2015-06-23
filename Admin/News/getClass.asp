<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Response.Charset="GB2312"
Dim classRs,classid,classIndex,Conn
MF_Default_Conn
classid=trim(request.QueryString("classid"))
classIndex=trim(request.QueryString("index"))
if classid="" then response.End()
Set classRs=Conn.execute("select id,ClassName,classid from FS_NS_NewsClass where isConstr=1 and ParentID=(select ClassID from FS_NS_NewsClass where id="&classid&")")'ID or ClassID?
Response.Write("<select name='sel_class_"&(Cint(classIndex)+1)&"' onChange=""getChildClass(this,"&(Cint(classIndex)+1)&")"">"&vbcrlf)
Response.Write("<option value='class'>"&(Cint(classIndex)+1)&"¼¶À¸Ä¿</option>"&vbcrlf)
while not classRs.eof
	Response.Write("<option value='"&classRs("id")&"'>"&classRs("ClassName")&"</option>"&vbcrlf) 
	classRs.movenext
Wend	
Response.Write("</select>"&vbcrlf)
Conn.close
Set Conn=nothing
%>







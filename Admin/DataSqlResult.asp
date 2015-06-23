<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%'Copyright (c) 2006 Foosun Inc.
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Sql,Result,ExeResult,ExeResultNum,ExeSelectTF,ErrorTF,FiledObj
Dim I,ErrObj,conn,User_Conn
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("MF017") then Err_Show
Result = Request.Form("Result")
if Result = "Submit" then
	Sql = Request.Form("Sql")
	Call MF_Insert_oper_Log("数据库维护","执行SQL语句:"& Request.Form("Sql") &"",now,session("admin_name"),"MF")
	if (Sql <> "") then
		If Instr(1,lcase(Sql),"delete from fs_ns_oper_log")<>0 then
			Sql="Select top 10 * from fs_ns_oper_log order by id desc"
		End If
		ExeSelectTF = (LCase(Left(Trim(Sql),6)) = "select")
		Conn.Errors.Clear
		On Error Resume Next
		if ExeSelectTF = True then
			if instr(Ucase(Sql)," FS_ME_")>0 then 
				Set ExeResult = User_Conn.ExeCute(Sql,ExeResultNum)
			else
				Set ExeResult = Conn.ExeCute(Sql,ExeResultNum)
			end if	
		else
			if instr(Ucase(Sql)," FS_ME_")>0 then 
				User_Conn.ExeCute Sql,ExeResultNum
			else
				Conn.ExeCute Sql,ExeResultNum
			end if				
		end if
		If Conn.Errors.Count<>0 Then
			ErrorTF = True
			Set ExeResult = Conn.Errors
		Else
			ErrorTF = False
		End If
	end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>执行结果</title>
</head>

<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<body topmargin="2" leftmargin="2">
<%
if Result = "Submit" then
if ErrorTF = True then
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="table">
  <tr class="xingmu"> 
    <td height="20" nowrap class="xingmu"> 
      <div align="center">错误号</div></td>
    <td height="20" nowrap class="xingmu"> 
      <div align="center">来源</div></td>
    <td height="20" nowrap class="xingmu"> 
      <div align="center">描述</div></td>
    <td height="20" nowrap class="xingmu"> 
      <div align="center">帮助</div></td>
    <td height="20" nowrap class="xingmu"> 
      <div align="center">帮助文档</div></td>
  </tr>
  <%
	For I=1 To Conn.Errors.Count
		Set ErrObj=Conn.Errors(I-1)
%>
  <tr class="hback"> 
    <td nowrap class="hback"> 
      <% = ErrObj.Number %> </td>
    <td nowrap class="hback"> 
      <% = ErrObj.Description %> </td>
    <td nowrap class="hback"> 
      <% = ErrObj.Source %> </td>
    <td nowrap class="hback"> 
      <% = ErrObj.Helpcontext %> </td>
    <td nowrap class="hback"> 
      <% = ErrObj.HelpFile %> </td>
  </tr>
  <%
	next
%>
</table>
<%
else
%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><table border="0" cellpadding="4" cellspacing="1" class="table">
        <%
	if ExeSelectTF = True then
%>
        <tr class="habck"> 
          <%
		For Each FiledObj In ExeResult.Fields
%>
          <td nowrap  height="26" class="xingmu"><div align="center"> 
              <% = FiledObj.name %>
            </div></td>
          <%
		next
%>
        </tr>
        <%
		do while Not ExeResult.Eof
%>
        <tr class="hback"> 
          <%
			For Each FiledObj In ExeResult.Fields
%>
          <td nowrap class="hback"> <div align="center"> 
              <%
		 if IsNull(FiledObj.value) then
		 	Response.Write("&nbsp;")
		 else
		 	Response.Write(FiledObj.value)
		 end if
		 %>
            </div></td>
          <%
			next
%>
        </tr>
        <%
			ExeResult.MoveNext
		loop
	else
%>
        <tr class="xingmu"> 
          <td class="xingmu" height="26"> <div align="center">执行结果</div></td>
        </tr>
        <tr class="hback"> 
          <td class="hback"> <div align="center"> 
              <% = ExeResultNum & "条纪录被影响"%>
            </div></td>
        </tr>
        <%
	end if
%>
      </table></td>
  </tr>
</table>
<%
end if
end if
%>
<form name="ExecuteForm" method="post" action="">
  <input type="hidden" name="Sql">
  <input type="hidden" name="Result">
</form>
</body>
</html>
<%
Set Conn = Nothing
Set ExeResult = Nothing
%>







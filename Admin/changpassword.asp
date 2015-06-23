<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Md5.asp" -->
<!--#include file="../FS_Inc/Cls_Cache.asp" -->
<%
Dim Conn
MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("MF024") then Err_Show
if Request.Form("Action") = "Save" then
	Dim str_o_pass,str_n_pass,str_cn_pass,strShowErr,f_Sql,obj_Cpwd_rs
	str_o_pass = Request.Form("pwd")
	str_n_pass =  Request.Form("pwd_new")
	str_cn_pass =  Request.Form("Confi_pwd_new")
	if Trim(str_n_pass)<>Trim(str_cn_pass) then
		strShowErr = "<li>两次输入密码不一致</li>"
		Response.Redirect("Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	if str_o_pass="" or str_n_pass="" or str_cn_pass="" then
		strShowErr = "<li>所有的必须填写</li>"
		Response.Redirect("Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
		Set  obj_Cpwd_rs = server.CreateObject(G_FS_RS)
		f_Sql = "Select Admin_Pass_Word from FS_MF_Admin Where Admin_Name='"& NoSqlHack(session("Admin_Name"))&"' and Admin_Pass_Word='"& MD5(str_o_pass,16)&"'"
		obj_Cpwd_rs.Open f_Sql,Conn,1,3
		If obj_Cpwd_rs.eof then
			strShowErr = "<li>原密码不正确</li>"
			Response.Redirect("Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		Else
			obj_Cpwd_rs("Admin_Pass_Word") = md5(str_n_pass,16)
			obj_Cpwd_rs.update
			strShowErr = "<li>密码修改成功</li>"
			Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
	End if
	obj_Cpwd_rs.close:set obj_Cpwd_rs = nothing
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form action="" method="post" name="newadmin" id="newadmin">
    <tr class="hback"> 
      <td colspan="2" class="xingmu">修改密码</td>
    </tr>
    <tr class="hback"> 
      <td width="140" align="right">管理员帐号</td>
      <td> <input type="text" name="name" value="<% = session("Admin_Name")%>" size="60" readonly="true"/> 
      </td>
    </tr>
    <tr class="hback"> 
      <td width="140" height="23" align="right">原密码</td>
      <td> <input type="password" name="pwd" value="" size="60" /> </td>
    </tr>
    <tr class="hback"> 
      <td align="right"><div align="right">新密码</div></td>
      <td><input name="pwd_new" type="password" id="pwd_new" value="" size="60" /></td>
    </tr>
    <tr class="hback"> 
      <td align="right"><div align="right">确认新密码</div></td>
      <td><input name="Confi_pwd_new" type="password" id="Confi_pwd_new" value="" size="60" /></td>
    </tr>
    <tr class="hback"> 
      <td align="right">&nbsp;</td>
      <td><input type="submit" name="Submit3" value=" 保存 "> <input type="reset" name="Submit4" value=" 重置 "> 
        <input name="Action" type="hidden" id="Action" value="Save"></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
Conn.Close
Set Conn = Nothing
%>






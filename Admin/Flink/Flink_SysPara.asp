<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,obj_sys_Rs,tmp_type,strShowErr
Dim IsOpen,IsRegister,ArrSize,Content,isLock,Update_type
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("FL002") then Err_Show	
set obj_sys_Rs = Server.CreateObject(G_FS_RS)
obj_sys_Rs.open "select top 1 id,IsOpen,IsRegister,ArrSize,Content,isLock from FS_FL_SysPara",Conn,1,3
if not obj_sys_Rs.eof then
	IsOpen = obj_sys_Rs("IsOpen")
	IsRegister = obj_sys_Rs("IsRegister")
	ArrSize = obj_sys_Rs("ArrSize")
	Content = obj_sys_Rs("Content")
	isLock = obj_sys_Rs("isLock")
	Update_type = "edit"
else
	IsOpen = 1
	IsRegister =0
	ArrSize = "88,31"
	Content = "申请注册"
	isLock = 1
	Update_type = "add"
end if
if Request.Form("Edit_save")<>"" then
	dim obj_sys_Rs_1,arr_tmp
	if trim(Request.Form("ArrSize"))<>"" then
		if instr(Request.Form("ArrSize"),",")=0 then
			strShowErr = "<li>Logo尺寸格式不对</li>"
			Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Flink/Flink_SysPara.asp")
			Response.end
		end if
	else
		strShowErr = "<li>请填写LOGO尺寸</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Flink/Flink_SysPara.asp")
		Response.end
	end if
	set obj_sys_Rs_1 = Server.CreateObject(G_FS_RS)
	obj_sys_Rs_1.open "select top 1 id,IsOpen,IsRegister,ArrSize,Content,isLock from FS_FL_SysPara",Conn,1,3
	if trim(Request.Form("Edit_save"))="add" then
		obj_sys_Rs_1.addnew
	end if
	if trim(Request.Form("isopen"))<>"" then:obj_sys_Rs_1("isOpen") = 1:else:obj_sys_Rs_1("isOpen") = 0:end if
	if trim(Request.Form("IsRegister"))<>"" then:obj_sys_Rs_1("IsRegister") = 1:else:obj_sys_Rs_1("IsRegister") = 0:end if
	if trim(Request.Form("isLock"))<>"" then:obj_sys_Rs_1("isLock") = 1:else:obj_sys_Rs_1("isLock") = 0:end if
	obj_sys_Rs_1("ArrSize") = NoSqlHack(Request.Form("ArrSize"))
	obj_sys_Rs_1("Content") = NoSqlHack(Request.Form("Content"))
	obj_sys_Rs_1.update
	obj_sys_Rs_1.close:set obj_sys_Rs_1=nothing
	strShowErr = "<li>修改成功</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Flink/Flink_SysPara.asp")
	Response.end
end if
%>
<html>
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr> 
    <td align="left" colspan="2" class="xingmu">参数设置</td>
  </tr>
</table>
  
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="form1" method="post" action="">
    <tr> 
      <td width="26%" class="hback"> 
        <div align="right">是否开放友情连接申请</div></td><td width="74%" class="hback"><input name="IsOpen" type="checkbox" id="IsOpen" value="1" <%if IsOpen=1 then response.Write("checked")%>>
        是</td>
    </tr>
    <tr> 
      <td class="hback"> 
        <div align="right">注册会员才能申请</div></td><td class="hback"><input name="IsRegister" type="checkbox" id="IsRegister" value="1" <%if IsRegister=1 then response.Write("checked")%>>
        是</td>
    </tr>
    <tr> 
      <td class="hback"> 
        <div align="right">友情连接图片最大及最小尺寸</div></td><td class="hback"><input name="ArrSize" type="text" id="ArrSize" value="<%= arrsize%>">
        格式:88,31</td>
    </tr>
    <tr> 
      <td class="hback"> 
        <div align="right">申请连接须知</div></td><td class="hback"><textarea name="Content" rows="5" id="Content"  style="width:60%"><% = content %></textarea></td>
    </tr>
    <tr> 
      <td class="hback"> 
        <div align="right">申请的连接是否要审核</div></td><td class="hback"><input name="isLock" type="checkbox" id="isLock" value="1" <%if isLock=1 then response.Write("checked")%>>
        是</td>
    </tr>
    <tr>
      <td class="hback"> 
        <div align="right"></div></td><td class="hback"><input type="submit" name="Submit" value="更新参数">
        <input name="Edit_save" type="hidden" id="Edit_save" value="<% = Update_type %>">
        <input type="reset" name="Submit2" value="重置"></td>
    </tr>
  </form>
</table>
</body>
</html>







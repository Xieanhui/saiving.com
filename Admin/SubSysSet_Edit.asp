<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Cls_Cache.asp" -->
<%
Dim Conn
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("MF_SubSite") then Err_Show
if not MF_Check_Pop_TF("MF013") then Err_Show
Dim p_Sub_Sys_Name,p_Sub_Sys_Path,p_Sub_Sys_Index,p_Sub_Sys_Installed,p_Sub_Sys_Link
DIm p_Save_Rs,p_Sub_ID
p_Sub_ID = NoSqlHack(Request.QueryString("Sub_ID"))
If Request.QueryString("Act")="Save" Then
	p_Sub_ID			= NoSqlHack(Request.form("Sub_Sys_ID"))
	p_Sub_Sys_Name		= NoSqlHack(Request.form("Sub_Sys_Name"))
	p_Sub_Sys_Path		= NoSqlHack(Request.form("Sub_Sys_Path"))
	p_Sub_Sys_Index		= NoSqlHack(Request.form("Sub_Sys_Index"))
	p_Sub_Sys_Installed	= NoSqlHack(Request.form("Sub_Sys_Installed"))
	p_Sub_Sys_Link	= NoSqlHack(Request.form("Sub_Sys_Link"))
	
	Set p_Save_Rs	= CreateObject(G_FS_RS)
	p_Save_Rs.Open "select * from FS_MF_Sub_Sys Where Sub_Sys_ID='" & p_Sub_ID & "'",Conn,3,3
	p_Save_Rs("Sub_Sys_Name") 		= p_Sub_Sys_Name
	p_Save_Rs("Sub_Sys_Path")		= p_Sub_Sys_Path
	p_Save_Rs("Sub_Sys_Index")		= p_Sub_Sys_Index
	p_Save_Rs("Sub_Sys_Link")		= p_Sub_Sys_Link
	If p_Sub_Sys_Installed <> "" Then
		p_Save_Rs("Sub_Sys_Installed")	= p_Sub_Sys_Installed
	End If
	p_Save_Rs.Update
	p_Save_Rs.Close
	Set p_Save_Rs = Nothing
	'Call Load_Install_Flag(p_Sub_ID)
	SubSys_Cookies
	Call MF_Insert_oper_Log("子系统维护","修改了子系统为:"& Request.form("Sub_Sys_Name") &","& Request.form("Sub_Sys_Path") &",,"& Request.form("Sub_Sys_Index") &",",now,session("admin_name"),"MF")
	Response.Redirect("SubSysSet_List.asp")
End If
Set p_Save_Rs	= CreateObject(G_FS_RS)
p_Save_Rs.Open "select * from FS_MF_Sub_Sys Where Sub_Sys_ID='" & p_Sub_ID & "'",Conn,1,1
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes >
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form action="?Act=Save" method="post" name="SubSysSet_Add" id="SubSysSet_Add">
    <tr class="hback"> 
      <td colspan="2" align="right" class="xingmu"><div align="left">子系统维护</div></td>
    </tr>
    <tr class="hback"> 
      <td width="180" align="right">子系统中文名称：</td>
      <td> <input type="text" name="Sub_Sys_Name" value="<%=p_Save_Rs("Sub_Sys_Name")%>" size="50" /> 
        <input type="hidden" name="Sub_Sys_ID" value="<%=p_Save_Rs("Sub_Sys_ID")%>" />      </td>
    </tr>
    <tr class="hback"> 
      <td align="right">子系统标识符：</td>
      <td> <input type="text" name="Sub_Sys_ID" value="<%=p_Save_Rs("Sub_Sys_ID")%>" size="50" disabled/>
        (不可修改！) </td>
    </tr>
    <tr class="hback"> 
      <td align="right">子系统安装目录：</td>
      <td> <input type="text" name="Sub_Sys_Path" value="<%=p_Save_Rs("Sub_Sys_Path")%>" size="50"  readonly />
        如果要修改，请及时修改子目录名称,目前不允许修改安装目录 </td>
    </tr>
    <tr class="hback"> 
      <td align="right">子系统首页文件：</td>
      <td> <input type="text" name="Sub_Sys_Index" value="<%=p_Save_Rs("Sub_Sys_Index")%>" size="50" />      </td>
    </tr>
    <tr class="hback">
      <td align="right">子系统导航连接地址</td>
      <td><input name="Sub_Sys_Link" type="text" id="Sub_Sys_Link" value="<%=p_Save_Rs("Sub_Sys_Link")%>" size="50" />
      <br>
      <span class="tx">(各个子系统的导航连接请在系统参数--子系统维护里设置)，特别注意：各个子站系统的前台导航连接必须与子系统的参数配置里的域名相同.格式：www.foosun.cms/news，前面请不要带&quot;http://&quot;</span></td>
    </tr>
    <tr class="hback"> 
      <td align="right">是否启用：</td>
      <td> <input name="Sub_Sys_Installed" type="radio" value="0" <% If p_Save_Rs("Sub_Sys_Installed")="0" Then Response.Write("checked")%> />
        关闭 
        <input type="radio" name="Sub_Sys_Installed" value="1" <% If p_Save_Rs("Sub_Sys_Installed")="1" Then Response.Write("checked")%> />
        启用</td>
    </tr>
    <tr class="hback"> 
      <td align="center" colspan="2"> <input type="button" name="tijiao" value=" 保存 " onClick="checkvalue();"/>
        　 
        <input type="reset" name="Submit2" value=" 重置 " /> </td>
    </tr>
  </form>
</table>
</body>
</html>
<%
p_Save_Rs.Close
Set p_Save_Rs = Nothing
Conn.Close
Set Conn = Nothing
%>
<script language="javascript">
function checkvalue()
{
	if(document.SubSysSet_Add.Sub_Sys_Name.value=='')
	{
		alert('子系统的中文名称不可为空！');
		return;
	}
	if(document.SubSysSet_Add.Sub_Sys_Path.value=='')
	{
		alert('子系统的安装目录不可为空！');
		return;
	}
	if(document.SubSysSet_Add.Sub_Sys_Index.value=='')
	{
		alert('子系统的首页文件不可为空！');
		return;
	}
	document.SubSysSet_Add.submit();
}
</script>






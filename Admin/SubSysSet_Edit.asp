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
	Call MF_Insert_oper_Log("��ϵͳά��","�޸�����ϵͳΪ:"& Request.form("Sub_Sys_Name") &","& Request.form("Sub_Sys_Path") &",,"& Request.form("Sub_Sys_Index") &",",now,session("admin_name"),"MF")
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
      <td colspan="2" align="right" class="xingmu"><div align="left">��ϵͳά��</div></td>
    </tr>
    <tr class="hback"> 
      <td width="180" align="right">��ϵͳ�������ƣ�</td>
      <td> <input type="text" name="Sub_Sys_Name" value="<%=p_Save_Rs("Sub_Sys_Name")%>" size="50" /> 
        <input type="hidden" name="Sub_Sys_ID" value="<%=p_Save_Rs("Sub_Sys_ID")%>" />      </td>
    </tr>
    <tr class="hback"> 
      <td align="right">��ϵͳ��ʶ����</td>
      <td> <input type="text" name="Sub_Sys_ID" value="<%=p_Save_Rs("Sub_Sys_ID")%>" size="50" disabled/>
        (�����޸ģ�) </td>
    </tr>
    <tr class="hback"> 
      <td align="right">��ϵͳ��װĿ¼��</td>
      <td> <input type="text" name="Sub_Sys_Path" value="<%=p_Save_Rs("Sub_Sys_Path")%>" size="50"  readonly />
        ���Ҫ�޸ģ��뼰ʱ�޸���Ŀ¼����,Ŀǰ�������޸İ�װĿ¼ </td>
    </tr>
    <tr class="hback"> 
      <td align="right">��ϵͳ��ҳ�ļ���</td>
      <td> <input type="text" name="Sub_Sys_Index" value="<%=p_Save_Rs("Sub_Sys_Index")%>" size="50" />      </td>
    </tr>
    <tr class="hback">
      <td align="right">��ϵͳ�������ӵ�ַ</td>
      <td><input name="Sub_Sys_Link" type="text" id="Sub_Sys_Link" value="<%=p_Save_Rs("Sub_Sys_Link")%>" size="50" />
      <br>
      <span class="tx">(������ϵͳ�ĵ�����������ϵͳ����--��ϵͳά��������)���ر�ע�⣺������վϵͳ��ǰ̨�������ӱ�������ϵͳ�Ĳ����������������ͬ.��ʽ��www.foosun.cms/news��ǰ���벻Ҫ��&quot;http://&quot;</span></td>
    </tr>
    <tr class="hback"> 
      <td align="right">�Ƿ����ã�</td>
      <td> <input name="Sub_Sys_Installed" type="radio" value="0" <% If p_Save_Rs("Sub_Sys_Installed")="0" Then Response.Write("checked")%> />
        �ر� 
        <input type="radio" name="Sub_Sys_Installed" value="1" <% If p_Save_Rs("Sub_Sys_Installed")="1" Then Response.Write("checked")%> />
        ����</td>
    </tr>
    <tr class="hback"> 
      <td align="center" colspan="2"> <input type="button" name="tijiao" value=" ���� " onClick="checkvalue();"/>
        �� 
        <input type="reset" name="Submit2" value=" ���� " /> </td>
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
		alert('��ϵͳ���������Ʋ���Ϊ�գ�');
		return;
	}
	if(document.SubSysSet_Add.Sub_Sys_Path.value=='')
	{
		alert('��ϵͳ�İ�װĿ¼����Ϊ�գ�');
		return;
	}
	if(document.SubSysSet_Add.Sub_Sys_Index.value=='')
	{
		alert('��ϵͳ����ҳ�ļ�����Ϊ�գ�');
		return;
	}
	document.SubSysSet_Add.submit();
}
</script>






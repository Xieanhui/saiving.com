<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<%
Dim Conn,User_Conn
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("AS_site") then Err_Show 
%>
<script language="JavaScript">
var isInternetExplorer = navigator.appName.indexOf("Microsoft") != -1;
function Menu_DoFSCommand(command, args) {
	var MenuObj = isInternetExplorer?document.all.Menu:document.Menu;
	top.ChangeURL(args);
}
</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>[site] �����̨ -- ��Ѷ���ݹ���ϵͳ FoosunCMS V5.0</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body scroll=yes style="margin:0px;" class="Leftback">
<!--#include file="../CommPages/Main_Navi.asp" -->
<table width="95%" border="0" align="center" cellpadding="2" cellspacing="2" class="leftframetable">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td><img src="../Images/Folder/sub_sys.gif" width="15" height="17"></td>
    <td class="titledaohang">���ϵͳ����</td>
  </tr>
<%if MF_Check_Pop_TF("AS001") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td width="97%"><a href="Ads_Manage.asp" target="ContentFrame">������</a></td>
  </tr>
  <%
end if%>
  <%if MF_Check_Pop_TF("AS002") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Ads_Count.asp" target="ContentFrame">���ͳ��</a></td>
    </tr>
  <%end if%>
  <%if MF_Check_Pop_TF("AS003") then%>
	<tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Ads_ClassManage.asp" target="ContentFrame">�������</a></td>
  </tr>
  <%end if%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="../shortCutMenu.asp" target="_self" class="admintx">������һ��</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="25">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>







<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<%
Dim Conn,User_Conn
MF_Default_Conn
MF_Session_TF
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
    <td class="titledaohang">վ��ͳ�Ƶ���</td>
  </tr>
  <%if MF_Check_Pop_TF("SS_site") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td width="97%"><a href="Visit_ObtainCode.asp" target="ContentFrame" class="lefttop">��ȡ����</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td width="97%"><a href="Visit_DataStatistic.asp" target="ContentFrame" class="lefttop">��Ҫ����</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td width="97%"><a href="Visit_HoursStatistic.asp" target="ContentFrame" class="lefttop">24Сʱͳ��</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_DaysStatistic.asp" target="ContentFrame">��ͳ��</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="16"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_MonthsStatistic.asp" target="ContentFrame">��ͳ��</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="16"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_SystemStatistic.asp" target="ContentFrame">ϵͳ/�����</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_AreaStatistic.asp" target="ContentFrame">����ͳ��</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_SourceStatistic.asp" target="ContentFrame">��Դͳ��</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_VisitorList.asp" target="ContentFrame">��������ͳ��</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:none;"> 
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_SearchStatistic.asp" target="ContentFrame">��������</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:none;"> 
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_fbStatistic.asp" target="ContentFrame">�ֱ���ͳ��</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="../SystemCheckplus.asp" target="ContentFrame">ϵͳ̽��</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_sysPara.asp" target="ContentFrame">��������</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="../shortCutMenu.asp" target="_self" class="admintx">������һ��</a></td>
  </tr>
  <%end if%>
</table>
</body>
</html>







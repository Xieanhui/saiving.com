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
<title>[site] 管理后台 -- 风讯内容管理系统 FoosunCMS V5.0</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body scroll=yes style="margin:0px;" class="Leftback">
<!--#include file="../CommPages/Main_Navi.asp" -->
<table width="95%" border="0" align="center" cellpadding="2" cellspacing="2" class="leftframetable">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td><img src="../Images/Folder/sub_sys.gif" width="15" height="17"></td>
    <td class="titledaohang">站点统计导航</td>
  </tr>
  <%if MF_Check_Pop_TF("SS_site") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td width="97%"><a href="Visit_ObtainCode.asp" target="ContentFrame" class="lefttop">获取代码</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td width="97%"><a href="Visit_DataStatistic.asp" target="ContentFrame" class="lefttop">简要数据</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td width="97%"><a href="Visit_HoursStatistic.asp" target="ContentFrame" class="lefttop">24小时统计</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_DaysStatistic.asp" target="ContentFrame">日统计</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="16"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_MonthsStatistic.asp" target="ContentFrame">月统计</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="16"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_SystemStatistic.asp" target="ContentFrame">系统/浏览器</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_AreaStatistic.asp" target="ContentFrame">地区统计</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_SourceStatistic.asp" target="ContentFrame">来源统计</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_VisitorList.asp" target="ContentFrame">来访者者统计</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:none;"> 
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_SearchStatistic.asp" target="ContentFrame">搜索引擎</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:none;"> 
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_fbStatistic.asp" target="ContentFrame">分辨率统计</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="../SystemCheckplus.asp" target="ContentFrame">系统探针</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Visit_sysPara.asp" target="ContentFrame">参数设置</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="25"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="../shortCutMenu.asp" target="_self" class="admintx">返回上一级</a></td>
  </tr>
  <%end if%>
</table>
</body>
</html>







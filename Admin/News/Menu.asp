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
    <td class="titledaohang">新闻系统导航</td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td width="97%"><a href="News_manage.asp" target="ContentFrame" class="lefttop">新闻管理</a>┆<a href="News_add.asp?ClassID=" target="ContentFrame">添加</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News_MyFolder.asp" target="ContentFrame">我的工作目录</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Class_Manage.asp" target="ContentFrame" class="lefttop">栏目管理</a>┆<a href="Class_add.asp?ClassID=&Action=add" target="ContentFrame">添加</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td> <div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Special_Manage.asp" target="ContentFrame">专题管理</a>┆<a href="Special_Add.asp?Action=add" target="ContentFrame">添加</a></td>
  </tr>
  <%if MF_Check_Pop_TF("NS_Constr") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Constr_Manage.asp" target="ContentFrame">投稿管理</a>┆<a href="Constr_stat.asp" target="ContentFrame">统计</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Templet") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="Class_ToTemplet.asp" target="ContentFrame">捆绑模板</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Freejs") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="JS_Free_Manage.asp" target="ContentFrame">自由JS管理</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Sysjs") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="JS_sys_Manage.asp" target="ContentFrame">系统JS管理</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Recyle") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News_Recyle.asp" target="ContentFrame" class="lefttop">回收站管理</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_UnRl") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="DefineNews_Manage.asp" target="ContentFrame">不规则新闻</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Genal") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="other_manage.asp" target="ContentFrame">常规管理</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Param") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="SysParaSet.asp" target="ContentFrame" class="lefttop">系统参数设置</a></td>
  </tr>
 <%end if%>
	<tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
    <td><a href="../shortCutMenu.asp" target="_self" class="admintx">返回上一级</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="25">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>







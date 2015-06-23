<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Cls_Cache.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn
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
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="2" class="Leftback" style="margin:0px;" scroll=yes>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="2"></td>
  </tr>
</table>
<%if MF_Check_Pop_TF("MF_Pop") then %>
<table width="95%" border="0" align="center" cellpadding="2" cellspacing="1" class="leftframetable">
  <tr classid="VoteManage"> 
    <td width="2%" height="21"><img src="Images/Folder/main_sys.gif" width="15" height="17"></td>
    <td width="98%"><table width="120" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="titledaohang" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(Menuid_main);"  language=javascript><font style="font-size:12px">系统参数</font></td>
        </tr>
      </table></td>
  </tr>
</table>
<table width="95%" border="0" align="center" cellpadding="2" cellspacing="1" class="leftframetable" Id="Menuid_main" style="display:none">
  <tr> 
    <td colspan="2"><table width="95" border="0" align="left" cellpadding="2" cellspacing="0">
    
		<%if MF_Check_Pop_TF("MF_SysSet") then %>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
          <td width="22%" valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
          <td width="78%"><a href="SysParaSet.asp" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>通用参数设置<br> </div>')" target="ContentFrame" class="lefttop">参数设置</a></td>
        </tr>
		<%end if%>
		<%if MF_Check_Pop_TF("MF_Const") then %>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="AdsManage" style="display:;"> 
          <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
          <td><a href="SysConstSet.asp"  onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>系统配置设置<br> </div>')"
		  target="ContentFrame">配置文件</a></td>
        </tr>
		<%end if%>
		<%if MF_Check_Pop_TF("MF_SubSite") then %>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="AdsManage" style="display:;"> 
          <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
          <td><a href="SubSysSet_List.asp" 	  onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>子系统维护<br> </div>')"
		  target="ContentFrame" class="lefttop">子系统维护</a></td>
        </tr>
		<%end if%>
		<%if MF_Check_Pop_TF("MF029") then %>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
          <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
          <td><a href="SysAdmin_list.asp"  onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>管理员管理<br> </div>')"
		  target="ContentFrame" class="lefttop">管理员管理</a></td>
        </tr>
		<%end if%>
		<%if MF_Check_Pop_TF("MF_DataFix") then %>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
          <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
          <td><a href="DataManage.asp"  onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>数据库维护<br> </div>')"
		  target="ContentFrame">数据库维护</a></td>
        </tr>
		<%end if%>
		<%if MF_Check_Pop_TF("MF_Log") then %>
        <tr classid="VoteManage" style="display:;"> 
          <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
          <td><a href="Sys_Oper_Log.asp"  onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>日志管理<br> </div>')"
		  target="ContentFrame">日志管理</a></td>
        </tr>
		<%end if%>
		<%if MF_Check_Pop_TF("MF_Define") then %>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
          <td height="20" valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
          <td><a href="DefineTable_Manage.asp" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>自定义字段<br> </div>')"
		  target="ContentFrame" class="lefttop">自定义字段</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
          <td height="20" valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
          <td><a href="CustomForm/FormManage.asp" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>自定义字段<br> </div>')"
		  target="ContentFrame" class="lefttop">自定义表单</a></td>
        </tr>
		<%end if%>
      </table></td>
  </tr>
</table>
<%end if%>
<%if MF_Check_Pop_TF("NS_Pop") then %>
<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="leftframetable">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td width="2%"><img src="Images/Folder/sub_sys.gif" width="15" height="17"></td>
    <td width="98%" class="titledaohang"  id=item$pval[CatID]) style="CURSOR: hand"  
	onmouseup="opencat(sub_news);"  language=javascript>新闻系统</td>
  </tr>
  <tr>
  <td colspan="2">
  <table width="100%" border="0" cellspacing="0" cellpadding="2" id="sub_news">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="2%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td width="98%"><a href="News/News_manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>新闻管理<br> </div>')"
	target="ContentFrame" class="lefttop">新闻管理</a>|<a href="News/News_add.asp?ClassID=" target="ContentFrame" title="添加新闻:高级模式">高级</a>┆<a href="News/News_add_Conc.asp?ClassID=" target="ContentFrame" title="添加新闻:简洁模式">简洁</a></td>
  </tr>
  
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/News_MyFolder.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>我的工作目录<br> </div>')"
	target="ContentFrame">我的工作目录</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/Class_Manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>新闻栏目管理<br> </div>')"
	target="ContentFrame" class="lefttop">栏目管理</a>┆<a href="News/Class_add.asp?ClassID=&Action=add" 
	target="ContentFrame">添加</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/Special_Manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>新闻专题管理<br> </div>')"
	target="ContentFrame">专题管理</a>┆<a href="News/Special_Add.asp?Action=add" target="ContentFrame">添加</a></td>
  </tr>
  <%if MF_Check_Pop_TF("NS_Constr") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/Constr_Manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>新闻投稿管理<br> </div>')"
	target="ContentFrame">投稿管理</a>┆<a href="News/Constr_stat.asp" target="ContentFrame">统计</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Templet") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/Class_ToTemplet.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>捆绑模板<br> </div>')"
	target="ContentFrame">捆绑模板</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Freejs") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/JS_Free_Manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>自由JS管理<br> </div>')"
	target="ContentFrame">自由JS管理</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Sysjs") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/JS_sys_Manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>系统JS管理<br> </div>')"
	target="ContentFrame">系统JS管理</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Recyle") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/News_Recyle.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>回收站管理<br> </div>')"
	target="ContentFrame" class="lefttop">回收站管理</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_UnRl") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/DefineNews_Manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>不规则新闻<br> </div>')"
	target="ContentFrame">不规则新闻</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Genal") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/other_manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>常规管理<br> </div>')"
	target="ContentFrame">常规管理</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Param") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/SysParaSet.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>系统参数设置<br> </div>')"
	target="ContentFrame" class="lefttop">系统参数设置</a></td>
  </tr>
 <%end if%>
  </table></td>
  </tr>
</table>
<%end if%>
<%
if Request.Cookies("FoosunSUBCookie")("FoosunSUBDS")=1 Then
	if MF_Check_Pop_TF("DS_Pop") then
%>
<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="leftframetable">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
    <td width="2%"><img src="Images/Folder/sub_sys.gif" width="15" height="17"></td>
    <td width="98%" class="titledaohang"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_down);"  language=javascript>下载系统</td>
  </tr>
  <tr>
    <td colspan="2">
	<table width="100%" border="0" cellspacing="0" cellpadding="2" id="sub_down" style="display:none">
  <%if MF_Check_Pop_TF("DS_Param") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="down/FS_DS_SysPara.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>参数设置<br> </div>')"
	target="ContentFrame" class="lefttop">参数设置</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("DS_Class") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="down/Class_Manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>栏目管理<br> </div>')"
	target="ContentFrame" class="lefttop">栏目管理</a>┆<a href="down/Class_add.asp?ClassID=&Action=add" target="ContentFrame">添加</a></td>
  </tr>
  <%end if
	  if MF_Check_Pop_TF("DS_speical") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
		<td> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="down/Down_Special_Manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>专区管理<br> </div>')"
		target="ContentFrame">专区管理</a>┆<a href="down/Down_Special_Edit_Add.asp?act=add" target="ContentFrame">添加</a></td>
	  </tr>
  <%
  end if
  if MF_Check_Pop_TF("DS_KunBang") then
   %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="down/Class_ToTemplet.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>捆绑模板<br> </div>')"
	target="ContentFrame" class="lefttop">捆绑模板</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("Down_List") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="down/DownloadList.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>下载管理<br> </div>')"
	target="ContentFrame" class="lefttop">下载管理</a></td>
  </tr>
  <%end if%>
    </table></td>
  </tr>
</table>
<%
	end if
End if
if Request.Cookies("FoosunSUBCookie")("FoosunSUBMS")=1 Then
	if MF_Check_Pop_TF("MS_Pop") then
%>
<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="leftframetable">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
    <td width="2%"><img src="Images/Folder/sub_sys.gif" width="15" height="17"></td>
    <td width="98%" class="titledaohang"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(mall_news);"  language=javascript> 商城B2C</td>
  </tr>
  <tr>
    <td colspan="2">
	<table width="100%" border="0" cellpadding="2" cellspacing="0"  id="mall_news" style="display:none;">
	  <%if MF_Check_Pop_TF("MS_Products") then%>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage"> 
		<td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/product/Product_Manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>商品管理<br> </div>')"
		target="ContentFrame" class="lefttop">商品管理</a></td>
	  </tr>
	  <%end if
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage"> 
		<td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/product/my_Product_Manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>我的目录<br> </div>')"
		target="ContentFrame">我的目录</a></td>
	  </tr>
	  <%
	  if MF_Check_Pop_TF("MS_Class") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage"> 
		<td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/Product/Product_Class_Set.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>栏目管理<br> </div>')"
		target="ContentFrame" class="lefttop">栏目管理</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("MS_Special") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
		<td> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/product/Product_Special_Manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>专区管理<br> </div>')"
		target="ContentFrame">专区管理</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("MS_order") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="23"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/order/Order_Manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>定单管理<br> </div>')"
		target="ContentFrame">定单管理</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("MS_WrOut") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="21"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/WithDraw/WithDraw_Deal.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>退货管理<br> </div>')"
		target="ContentFrame">退货管理</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("MS_Company") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/Company/Company_Express_Manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>物流公司管理<br> </div>')"
		target="ContentFrame">物流公司</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("MS_Recycle") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/Product_Recyle.asp" target="ContentFrame"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>回收站管理<br> </div>')"
		class="lefttop">回收站管理</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("MS_bind") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/templet/Class_ToTemplet.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>回收站管理<br> </div>')"
		target="ContentFrame">捆绑模板</a></td>
	  </tr>
	  <%
	  end if
	  if MF_Check_Pop_TF("MS_Param") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="#" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>回收站管理<br> </div>')">系统参数</a></td>
	  </tr>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="20" colspan="3"> <table width="100" border="0" cellpadding="0" cellspacing="0">
			<tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
			  <td width="25" height="20"><div align="right"><img src="Images/L.gif" width="17" height="16"></div></td>
			  <td width="75"><a href="mall/Sys_Para_Set.asp" target="ContentFrame">参数设置</a></td>
			</tr>
			<tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
			  <td height="20"><div align="right"><img src="Images/L.gif" width="17" height="16"></div></td>
			  <td><a href="mall/Help/Help_Mail_Infor.asp" target="ContentFrame">邮寄资料</a></td>
			</tr>
			<tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
			  <td height="20"><div align="right"><img src="Images/L.gif" width="17" height="16"></div></td>
			  <td><a href="mall/Help/Help_Express_Infor.asp" target="ContentFrame">配送须知</a></td>
			</tr>
			<tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
			  <td height="20"><div align="right"><img src="Images/L.gif" width="17" height="16"></div></td>
			  <td><a href="mall/Help/Help_bank_Infor.asp" target="ContentFrame">银行资料</a></td>
			</tr>
		  </table></td>
	  </tr>
	  <%end if%>
	</table></td>
  </tr>
</table>
<%
	end if
End if
if Request.Cookies("FoosunSUBCookie")("FoosunSUBME")=1 Then
	if MF_Check_Pop_TF("ME_Pop") then
%>
<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="leftframetable">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
    <td width="2%"><img src="Images/Folder/sub_sys.gif" width="15" height="17"></td>
    <td width="98%" class="titledaohang"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_user);"  language=javascript>会员系统</td>
  </tr>
  <tr>
    <td colspan="2">
	<table width="100%" border="0" cellpadding="2" cellspacing="0" id="sub_user" style="display:none;">
  <%if MF_Check_Pop_TF("ME_List") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/User_manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>个人会员管理<br> </div>')"
	target="ContentFrame" class="lefttop">个人会员管理</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/UserCorp.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>企业会员管理<br> </div>')"
	target="ContentFrame" class="lefttop">企业会员管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Intergel") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/Integral.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>积分管理<br> </div>')"
	target="ContentFrame" class="lefttop">积分管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Card") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/Card.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>点卡管理<br> </div>')"
	target="ContentFrame" class="lefttop">点卡管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_News") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/news_manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>公告管理<br> </div>')"
	target="ContentFrame" class="lefttop">公告管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Form") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/GroupDebate_manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>社群管理<br> </div>')"
	target="ContentFrame" class="lefttop">社群管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_HY") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/VocationClass.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>行业管理<br> </div>')"
	target="ContentFrame" class="lefttop">行业管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_award") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/award.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>抽奖管理<br> </div>')"
	target="ContentFrame" class="lefttop">抽奖管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Order") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/Order_Pay.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>定单管理(在线支付)<br> </div>')"
	target="ContentFrame">定单管理(在线支付)</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Mproducts") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
    <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/get_Thing.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>添加会员商品<br> </div>')"
	target="ContentFrame">添加会员商品</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Horder") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/History_order.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>交易明晰<br> </div>')"
	target="ContentFrame" class="lefttop">交易明晰</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_GUser") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/Group_manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>会员组<br> </div>')"
	target="ContentFrame" class="lefttop">会员组</a>┆<a href="user/Group_Add.asp" target="ContentFrame" class="lefttop">添加</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Jubao") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/UserReport.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>举报管理<br> </div>')"
	target="ContentFrame">举报管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Review") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/Review.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>评论管理<br> </div>')"
	target="ContentFrame">评论管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Log") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/iLog.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>日记网摘<br> </div>')"
	target="ContentFrame">日记网摘</a></td>
  </tr>
   <% end if %>
  <%if MF_Check_Pop_TF("ME_Photo") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
    <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/Photo.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>相册管理<br> </div>')"
	target="ContentFrame">相册管理</a></td>
  </tr>
    <% end if %>
  <%if MF_Check_Pop_TF("ME_Param") then%>
<tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td width="3%" valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td width="97%"><a href="user/UserParam.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>参数设置<br> </div>')"
	target="ContentFrame" class="lefttop">参数设置</a></td>
  </tr>
   <% end if %>
  <%if MF_Check_Pop_TF("ME_Pay") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20" valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/PayParam.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>在线支付<br> </div>')"
	target="ContentFrame">在线支付</a></td>
  </tr>
   <% end if %>
</table></td>
  </tr>
</table>
<%
	End if
End if
if Request.Cookies("FoosunSUBCookie")("FoosunSUBSD")=1 Then
	if MF_Check_Pop_TF("SD_Pop") then
%>
<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="leftframetable">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
    <td width="2%"><img src="Images/Folder/sub_sys.gif" width="15" height="17"></td>
    <td width="98%" class="titledaohang"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_sd);"  language=javascript>供求系统</td>
  </tr>
  <tr>
    <td colspan="2">
	<table width="100%" border="0" cellpadding="2" cellspacing="0" id="sub_sd" style="display:none;">
	<%if MF_Check_Pop_TF("SD_List") then%>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22"><a  href="supply/News.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>供求信息<br> </div>')"
		target="ContentFrame">供求信息</a></td>
	  </tr>
	<%End IF%>
	<%if  MF_Check_Pop_TF("SD_Class") then %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22"><a  href="supply/Class.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>类别栏目<br> </div>')"
		target="ContentFrame">类别栏目</a></td>
	  </tr>
	<%End IF%>
	<%if MF_Check_Pop_TF("AP_area") then %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22"><a href="supply/Area.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>区域管理<br> </div>')"
		target="ContentFrame">区域管理</a></td>
	  </tr>
	 <%End IF%>
	 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:none;"> 
		<td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22"><a  href="supply/Lable.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>标签管理<br> </div>')"
		target="ContentFrame">标签管理</a></td>
	  </tr>
	
	<%if MF_Check_Pop_TF("AP_param") then %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
		<td> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22"><a  href="supply/Config.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>系统设置<br> </div>')"
		target="ContentFrame">系统设置</a></td>
	  </tr>
	 <%End IF%>
	</table>
</td>
  </tr>
</table>
<%
	end if
End if
if Request.Cookies("FoosunSUBCookie")("FoosunSUBAP")=1 Then
	if MF_Check_Pop_TF("AP_Pop") then
%>
<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="leftframetable">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
    <td width="2%"><img src="Images/Folder/sub_sys.gif" width="15" height="17"></td>
    <td width="98%" class="titledaohang"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_ap);"  language=javascript>人才招聘</td>
  </tr>
  <tr>
    <td colspan="2"><table width="100%" border="0" cellpadding="2" cellspacing="0" id="sub_ap" style="display:none;">
  <%if MF_Check_Pop_TF("AP_Param") then%>
   <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="job/FS_AP_SysPara.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>系统参数设置<br> </div>')"
	target="ContentFrame" class="lefttop">系统参数设置</a></td>
  </tr>
  <%
  End if
  if MF_Check_Pop_TF("AP_Trade") then
  %>
   <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="job/AP_Trade.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>省份设置<br> </div>')"
	target="ContentFrame" class="lefttop">人才行业设置</a></td>
  </tr>
  <%
  End if
  if MF_Check_Pop_TF("AP_Job") then
  %>
   <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="job/AP_Job.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>省份设置<br> </div>')"
	target="ContentFrame" class="lefttop">行业职位设置</a></td>
  </tr>
  <%
  End if
  if MF_Check_Pop_TF("AP_Province") then
  %>
   <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="job/AP_Province.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>省份设置<br> </div>')"
	target="ContentFrame" class="lefttop">省份设置</a></td>
  </tr>
  <%
  End if
  if MF_Check_Pop_TF("AP_city") then
  %>
   <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="job/AP_City.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>城市设置<br> </div>')"
	target="ContentFrame" class="lefttop">城市设置</a></td>
  </tr>
  <%
  End if
  if MF_Check_Pop_TF("AP_Search") then
  %>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="job/AP_Payment_Search.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>会员记录查询<br> </div>')"
	target="ContentFrame" class="lefttop">会员记录查询</a></td>
  </tr>
  <%
  End if
  if MF_Check_Pop_TF("AP_check") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="job/AP_Register_List.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>注册信息审核<br> </div>')"
	target="ContentFrame" class="lefttop">注册信息审核</a></td>
  </tr>
  <%End if%>
</table></td>
  </tr>
</table>
<%
	end if
End if
if Request.Cookies("FoosunSUBCookie")("FoosunSUBHS")=1 Then
	if MF_Check_Pop_TF("HS_Pop") then
%>
<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="leftframetable">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
    <td width="2%"><img src="Images/Folder/sub_sys.gif" width="15" height="17"></td>
    <td width="98%" class="titledaohang"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_hs);"  language=javascript>房产楼盘</td>
  </tr>
  <tr>
    <td colspan="2">
	<table width="95%" border="0" align="center" cellpadding="2" cellspacing="0" id="sub_hs" style="display:none">
	  <%if MF_Check_Pop_TF("HS_Loup") then%>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td width="2%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td width="196%" height="22" colspan="2"><a href="house/HS_Quotation_Open.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>楼盘信息<br> </div>')"
		target="ContentFrame" class="lefttop">楼盘信息</a></td>
	  </tr>
	  <%
	  End if
	  if MF_Check_Pop_TF("HS_Ero") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="house/HS_Second.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>二手房发布<br> </div>')"
		target="ContentFrame" class="lefttop">二手房发布</a></td>
	  </tr>
	  <%
	  End if
	  if MF_Check_Pop_TF("HS_Zu") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
		<td> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="house/HS_Tenancy.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>租赁信息<br> </div>')"
		target="ContentFrame" class="lefttop">租赁信息</a></td>
	  </tr>
	  <%
	  End if
	  if MF_Check_Pop_TF("HS_param") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="23"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="23" colspan="2"><a href="house/Sys_Para_Set.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>系统参数设置<br> </div>')"
		target="ContentFrame">系统参数设置</a></td>
	  </tr>
	  <%
	  End if
	  %>
	  <%
	  if MF_Check_Pop_TF("HS013") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="22"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="house/HS_Recyle.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>回收站管理<br> </div>')"
		class="lefttop" target="ContentFrame">回收站管理</a></td>
	  </tr>
	  <%
	  End if
	  if MF_Check_Pop_TF("HS014") or MF_Check_Pop_TF("HS_Search") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td colspan="2"><a href="house/HS_Clear.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>清理过期房源<br> </div>')"
		target="ContentFrame" class="lefttop">清理过期房源</a></td>
	  </tr>
	  <%End if%>
	</table>
	</td>
  </tr>
</table>
<%
	end if
End if%>
<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="leftframetable">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
    <td width="2%"><img src="Images/Folder/sub_sys.gif" width="15" height="17"></td>
    <td width="98%" class="titledaohang"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_other);"  language=javascript>其他系统</td>
  </tr>
  <tr style="display:none" id="sub_other">
    <td colspan="2">
	<%
	if Request.Cookies("FoosunSUBCookie")("FoosunSUBCS")=1 Then
	%>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0">
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
          <td width="100%" height="19"   id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_cs);"  onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>采集导航<br> </div>')"
		  language=javascript><img src="Images/+.gif" width="15" height="15">采集导航</td>
        </tr>
      </table>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0" id="sub_cs" style="display:none">
         <%if MF_Check_Pop_TF("CS_site") then%>
		 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td><div align="center"><img src="Images/L-.gif" width="17" height="16"></div></td>
          <td width="96%" height="16"><a href="collect/Site.asp"
		  target="ContentFrame">站点设置</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="4%"><div align="center"><img src="Images/L-.gif" width="17" height="16"></div></td>
          <td height="16"><a href="collect/Rule.asp" target="ContentFrame">关键字过滤</a></td>
        </tr>
		<%end if%>
		<%if MF_Check_Pop_TF("CS_Ink") then%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="16"><div align="center"><img src="Images/L-.gif" width="17" height="16"></div></td>
          <td height="16"><a href="collect/Check.asp" target="ContentFrame">新闻处理</a></td>
        </tr>
		<%end if%>
		 <%if MF_Check_Pop_TF("CS_site") then%>
		 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td><div align="center"><img src="Images/L-.gif" width="17" height="16"></div></td>
          <td width="96%" height="16"><a href="collect/AutoCollect.asp" target="ContentFrame">定时采集</a></td>
        </tr>
		<%end if%>
      </table>
		<%
		End if
		if Request.Cookies("FoosunSUBCookie")("FoosunSUBSS")=1 Then
		%>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0">
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
          <td width="100%" height="19"   id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_ss);" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>站点统计<br> </div>')"
		  language=javascript><img src="Images/+.gif" width="15" height="15">站点统计</td>
        </tr>
      </table>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0" style="display:none" id="sub_ss">
        <%if MF_Check_Pop_TF("SS_site") then%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="4%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="96%"><a href="Stat/Visit_ObtainCode.asp" target="ContentFrame" class="lefttop">获取代码</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="4%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="96%"><a href="Stat/Visit_DataStatistic.asp" target="ContentFrame" class="lefttop">简要数据</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="4%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="96%"><a href="Stat/Visit_HoursStatistic.asp" target="ContentFrame" class="lefttop">24小时统计</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="20"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_DaysStatistic.asp" target="ContentFrame">日统计</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="16"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_MonthsStatistic.asp" target="ContentFrame">月统计</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="16"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_SystemStatistic.asp" target="ContentFrame">系统/浏览器</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="25"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_AreaStatistic.asp" target="ContentFrame">地区统计</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="25"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_SourceStatistic.asp" target="ContentFrame">来源统计</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="25"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_VisitorList.asp" target="ContentFrame">来访者者统计</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:none;">
          <td height="25"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_SearchStatistic.asp" target="ContentFrame">搜索引擎</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:none;">
          <td height="25"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_fbStatistic.asp" target="ContentFrame">分辨率统计</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="25"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="SystemCheckplus.asp" target="ContentFrame">系统探针</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="25"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_sysPara.asp" target="ContentFrame">参数设置</a></td>
        </tr>
        <%end if%>
      </table>
		<%
		End if
		if Request.Cookies("FoosunSUBCookie")("FoosunSUBVS")=1 Then
		%>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0">
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
          <td width="100%" height="19"   id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_vs);" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>投票管理<br> </div>')"
		  language=javascript><img src="Images/+.gif" width="15" height="15">投票管理</td>
        </tr>
      </table>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="1" id="sub_vs" style="display:none">
        <%if MF_Check_Pop_TF("VS_site") Then%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td height="22" colspan="2"><a href="vote/FS_VS_SysPara.asp" target="ContentFrame" class="lefttop">系统参数设置</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td height="22" colspan="2"><a href="vote/VS_Class.asp" target="ContentFrame" class="lefttop">投票分类设置</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td height="22" colspan="2"><a href="vote/VS_Theme.asp" target="ContentFrame" class="lefttop">投票主题设置</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td height="22" colspan="2"><a href="vote/VS_Items.asp" target="ContentFrame" class="lefttop">投票选项设置</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td height="22" colspan="2"><a href="vote/VS_Steps.asp" target="ContentFrame" class="lefttop">多步投票管理</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td height="22" colspan="2"><a href="vote/VS_Items_Result.asp" target="ContentFrame" class="lefttop">投票情况管理</a></td>
        </tr>
        <%end if%>
      </table>
		<%
		End if
		if Request.Cookies("FoosunSUBCookie")("FoosunSUBAS")=1 Then
		%>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0">
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
          <td width="100%" height="19"   id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_as);" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>广告管理<br> </div>')"
		  language=javascript><img src="Images/+.gif" width="15" height="15">广告管理</td>
        </tr>
      </table>
	  <table width="95%" border="0" align="center" cellpadding="2" cellspacing="2" id="sub_as" style="display:none">
        <%if MF_Check_Pop_TF("AS001") then%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="97%"><a href="ads/Ads_Manage.asp" target="ContentFrame">广告管理</a></td>
        </tr>
        <%
		end if%>
        <%if MF_Check_Pop_TF("AS002") then%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="ads/Ads_Count.asp" target="ContentFrame">广告统计</a></td>
        </tr>
        <%end if%>
        <%if MF_Check_Pop_TF("AS003") then%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="ads/Ads_ClassManage.asp" target="ContentFrame">分类管理</a></td>
        </tr>
        <%end if%>
      </table>
		<%
		End if
		if Request.Cookies("FoosunSUBCookie")("FoosunSUBWS")=1 Then
		%>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0">
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
          <td width="100%" height="19"   id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_ws);" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>留言管理<br> </div>')"
		  language=javascript><img src="Images/+.gif" width="15" height="15">留言管理</td>
        </tr>
      </table>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0" id="sub_ws" style="display:none">
        <%if MF_Check_Pop_TF("WS_site") then%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="97%"><a href="bbs/SysConfig.asp" target="ContentFrame" class="lefttop">参数设置</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="97%"><a href="bbs/NewTell.asp" target="ContentFrame" class="lefttop">公告管理</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="97%"><a href="bbs/ClassManager.asp" target="ContentFrame" class="lefttop">系统分类</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="97%"><a href="bbs/ClassMessageManager.asp" target="ContentFrame" class="lefttop">留言管理</a></td>
        </tr>
        <%end if%>
      </table>
		<%
		End if
		if Request.Cookies("FoosunSUBCookie")("FoosunSUBFL")=1 Then
		%>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0">
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
          <td width="100%" height="19"   id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_fl);" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>友情连接<br> </div>')"
		  language=javascript><img src="Images/+.gif" width="15" height="15">友情连接</td>
        </tr>
      </table>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="1" id="sub_fl" style="display:none">
  
	  <%if MF_Check_Pop_TF("FL_site") then%>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
		<td width="97%"><a href="Flink/Flink_Manage.asp" target="ContentFrame" class="lefttop">添加┆管理</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("FL_site") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
		<td width="97%"><a href="Flink/Flik_Class.asp" target="ContentFrame" class="lefttop">分类管理</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("FL_site") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
		<td width="97%"><a href="Flink/Flink_SysPara.asp" target="ContentFrame" class="lefttop">参数设置</a></td>
	  </tr>
	  <%end if%>
	</table>
	<%END IF%>
</td>
  </tr>
</table>

<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="leftframetable">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td width="2%"><img src="Images/Folder/sub_sys.gif" width="15" height="17"></td>
    <td width="98%" class="titledaohang">系统信息</td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="105" colspan="2">版权所有：<a href="http://www.foosun.net" target="_blank">风讯在线</a><br>
      设计制作：<a href="http://www.foosun.cn" target="_blank">Foosun Inc.</a><br>
      技术支持：<a href="http://bbs.foosun.net" target="_blank">风讯论坛</a><br>
      帮助中心：<a href="http://help.foosun.net" target="_blank">风讯帮助</a> <br>
      系统版本：5.0.0</td>
  </tr>
</table>
<table width="95%" height="115" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
<script language="javascript" type="text/javascript" src="../FS_Inc/wz_tooltip.js"></script>
<script language="JavaScript" type="text/JavaScript">
function opencat(cat)
{
	  if(cat.style.display=="none")
	  {
		 cat.style.display="";
	  }
	  else
	  {
		 cat.style.display="none"; 
	  }
}
</script>






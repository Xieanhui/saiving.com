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
<title>[site] �����̨ -- ��Ѷ���ݹ���ϵͳ FoosunCMS V5.0</title>
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
          <td class="titledaohang" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(Menuid_main);"  language=javascript><font style="font-size:12px">ϵͳ����</font></td>
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
          <td width="78%"><a href="SysParaSet.asp" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>ͨ�ò�������<br> </div>')" target="ContentFrame" class="lefttop">��������</a></td>
        </tr>
		<%end if%>
		<%if MF_Check_Pop_TF("MF_Const") then %>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="AdsManage" style="display:;"> 
          <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
          <td><a href="SysConstSet.asp"  onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>ϵͳ��������<br> </div>')"
		  target="ContentFrame">�����ļ�</a></td>
        </tr>
		<%end if%>
		<%if MF_Check_Pop_TF("MF_SubSite") then %>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="AdsManage" style="display:;"> 
          <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
          <td><a href="SubSysSet_List.asp" 	  onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��ϵͳά��<br> </div>')"
		  target="ContentFrame" class="lefttop">��ϵͳά��</a></td>
        </tr>
		<%end if%>
		<%if MF_Check_Pop_TF("MF029") then %>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
          <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
          <td><a href="SysAdmin_list.asp"  onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>����Ա����<br> </div>')"
		  target="ContentFrame" class="lefttop">����Ա����</a></td>
        </tr>
		<%end if%>
		<%if MF_Check_Pop_TF("MF_DataFix") then %>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
          <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
          <td><a href="DataManage.asp"  onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>���ݿ�ά��<br> </div>')"
		  target="ContentFrame">���ݿ�ά��</a></td>
        </tr>
		<%end if%>
		<%if MF_Check_Pop_TF("MF_Log") then %>
        <tr classid="VoteManage" style="display:;"> 
          <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
          <td><a href="Sys_Oper_Log.asp"  onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��־����<br> </div>')"
		  target="ContentFrame">��־����</a></td>
        </tr>
		<%end if%>
		<%if MF_Check_Pop_TF("MF_Define") then %>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
          <td height="20" valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
          <td><a href="DefineTable_Manage.asp" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>�Զ����ֶ�<br> </div>')"
		  target="ContentFrame" class="lefttop">�Զ����ֶ�</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
          <td height="20" valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
          <td><a href="CustomForm/FormManage.asp" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>�Զ����ֶ�<br> </div>')"
		  target="ContentFrame" class="lefttop">�Զ����</a></td>
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
	onmouseup="opencat(sub_news);"  language=javascript>����ϵͳ</td>
  </tr>
  <tr>
  <td colspan="2">
  <table width="100%" border="0" cellspacing="0" cellpadding="2" id="sub_news">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="2%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td width="98%"><a href="News/News_manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>���Ź���<br> </div>')"
	target="ContentFrame" class="lefttop">���Ź���</a>|<a href="News/News_add.asp?ClassID=" target="ContentFrame" title="�������:�߼�ģʽ">�߼�</a>��<a href="News/News_add_Conc.asp?ClassID=" target="ContentFrame" title="�������:���ģʽ">���</a></td>
  </tr>
  
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/News_MyFolder.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>�ҵĹ���Ŀ¼<br> </div>')"
	target="ContentFrame">�ҵĹ���Ŀ¼</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/Class_Manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>������Ŀ����<br> </div>')"
	target="ContentFrame" class="lefttop">��Ŀ����</a>��<a href="News/Class_add.asp?ClassID=&Action=add" 
	target="ContentFrame">���</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/Special_Manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>����ר�����<br> </div>')"
	target="ContentFrame">ר�����</a>��<a href="News/Special_Add.asp?Action=add" target="ContentFrame">���</a></td>
  </tr>
  <%if MF_Check_Pop_TF("NS_Constr") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/Constr_Manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>����Ͷ�����<br> </div>')"
	target="ContentFrame">Ͷ�����</a>��<a href="News/Constr_stat.asp" target="ContentFrame">ͳ��</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Templet") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/Class_ToTemplet.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>����ģ��<br> </div>')"
	target="ContentFrame">����ģ��</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Freejs") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/JS_Free_Manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>����JS����<br> </div>')"
	target="ContentFrame">����JS����</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Sysjs") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/JS_sys_Manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>ϵͳJS����<br> </div>')"
	target="ContentFrame">ϵͳJS����</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Recyle") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/News_Recyle.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>����վ����<br> </div>')"
	target="ContentFrame" class="lefttop">����վ����</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_UnRl") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/DefineNews_Manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>����������<br> </div>')"
	target="ContentFrame">����������</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Genal") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/other_manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>�������<br> </div>')"
	target="ContentFrame">�������</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("NS_Param") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td><a href="News/SysParaSet.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>ϵͳ��������<br> </div>')"
	target="ContentFrame" class="lefttop">ϵͳ��������</a></td>
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
    <td width="98%" class="titledaohang"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_down);"  language=javascript>����ϵͳ</td>
  </tr>
  <tr>
    <td colspan="2">
	<table width="100%" border="0" cellspacing="0" cellpadding="2" id="sub_down" style="display:none">
  <%if MF_Check_Pop_TF("DS_Param") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="down/FS_DS_SysPara.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��������<br> </div>')"
	target="ContentFrame" class="lefttop">��������</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("DS_Class") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="down/Class_Manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��Ŀ����<br> </div>')"
	target="ContentFrame" class="lefttop">��Ŀ����</a>��<a href="down/Class_add.asp?ClassID=&Action=add" target="ContentFrame">���</a></td>
  </tr>
  <%end if
	  if MF_Check_Pop_TF("DS_speical") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
		<td> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="down/Down_Special_Manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>ר������<br> </div>')"
		target="ContentFrame">ר������</a>��<a href="down/Down_Special_Edit_Add.asp?act=add" target="ContentFrame">���</a></td>
	  </tr>
  <%
  end if
  if MF_Check_Pop_TF("DS_KunBang") then
   %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="down/Class_ToTemplet.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>����ģ��<br> </div>')"
	target="ContentFrame" class="lefttop">����ģ��</a></td>
  </tr>
  <%
  end if
  if MF_Check_Pop_TF("Down_List") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="down/DownloadList.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>���ع���<br> </div>')"
	target="ContentFrame" class="lefttop">���ع���</a></td>
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
    <td width="98%" class="titledaohang"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(mall_news);"  language=javascript> �̳�B2C</td>
  </tr>
  <tr>
    <td colspan="2">
	<table width="100%" border="0" cellpadding="2" cellspacing="0"  id="mall_news" style="display:none;">
	  <%if MF_Check_Pop_TF("MS_Products") then%>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage"> 
		<td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/product/Product_Manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��Ʒ����<br> </div>')"
		target="ContentFrame" class="lefttop">��Ʒ����</a></td>
	  </tr>
	  <%end if
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage"> 
		<td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/product/my_Product_Manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>�ҵ�Ŀ¼<br> </div>')"
		target="ContentFrame">�ҵ�Ŀ¼</a></td>
	  </tr>
	  <%
	  if MF_Check_Pop_TF("MS_Class") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage"> 
		<td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/Product/Product_Class_Set.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��Ŀ����<br> </div>')"
		target="ContentFrame" class="lefttop">��Ŀ����</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("MS_Special") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
		<td> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/product/Product_Special_Manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>ר������<br> </div>')"
		target="ContentFrame">ר������</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("MS_order") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="23"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/order/Order_Manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��������<br> </div>')"
		target="ContentFrame">��������</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("MS_WrOut") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="21"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/WithDraw/WithDraw_Deal.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>�˻�����<br> </div>')"
		target="ContentFrame">�˻�����</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("MS_Company") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/Company/Company_Express_Manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>������˾����<br> </div>')"
		target="ContentFrame">������˾</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("MS_Recycle") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/Product_Recyle.asp" target="ContentFrame"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>����վ����<br> </div>')"
		class="lefttop">����վ����</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("MS_bind") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="mall/templet/Class_ToTemplet.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>����վ����<br> </div>')"
		target="ContentFrame">����ģ��</a></td>
	  </tr>
	  <%
	  end if
	  if MF_Check_Pop_TF("MS_Param") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="20"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="#" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>����վ����<br> </div>')">ϵͳ����</a></td>
	  </tr>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="20" colspan="3"> <table width="100" border="0" cellpadding="0" cellspacing="0">
			<tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
			  <td width="25" height="20"><div align="right"><img src="Images/L.gif" width="17" height="16"></div></td>
			  <td width="75"><a href="mall/Sys_Para_Set.asp" target="ContentFrame">��������</a></td>
			</tr>
			<tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
			  <td height="20"><div align="right"><img src="Images/L.gif" width="17" height="16"></div></td>
			  <td><a href="mall/Help/Help_Mail_Infor.asp" target="ContentFrame">�ʼ�����</a></td>
			</tr>
			<tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
			  <td height="20"><div align="right"><img src="Images/L.gif" width="17" height="16"></div></td>
			  <td><a href="mall/Help/Help_Express_Infor.asp" target="ContentFrame">������֪</a></td>
			</tr>
			<tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
			  <td height="20"><div align="right"><img src="Images/L.gif" width="17" height="16"></div></td>
			  <td><a href="mall/Help/Help_bank_Infor.asp" target="ContentFrame">��������</a></td>
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
    <td width="98%" class="titledaohang"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_user);"  language=javascript>��Աϵͳ</td>
  </tr>
  <tr>
    <td colspan="2">
	<table width="100%" border="0" cellpadding="2" cellspacing="0" id="sub_user" style="display:none;">
  <%if MF_Check_Pop_TF("ME_List") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/User_manage.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>���˻�Ա����<br> </div>')"
	target="ContentFrame" class="lefttop">���˻�Ա����</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/UserCorp.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��ҵ��Ա����<br> </div>')"
	target="ContentFrame" class="lefttop">��ҵ��Ա����</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Intergel") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/Integral.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>���ֹ���<br> </div>')"
	target="ContentFrame" class="lefttop">���ֹ���</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Card") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/Card.asp" 	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>�㿨����<br> </div>')"
	target="ContentFrame" class="lefttop">�㿨����</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_News") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/news_manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>�������<br> </div>')"
	target="ContentFrame" class="lefttop">�������</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Form") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/GroupDebate_manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��Ⱥ����<br> </div>')"
	target="ContentFrame" class="lefttop">��Ⱥ����</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_HY") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/VocationClass.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��ҵ����<br> </div>')"
	target="ContentFrame" class="lefttop">��ҵ����</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_award") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/award.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>�齱����<br> </div>')"
	target="ContentFrame" class="lefttop">�齱����</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Order") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/Order_Pay.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��������(����֧��)<br> </div>')"
	target="ContentFrame">��������(����֧��)</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Mproducts") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
    <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/get_Thing.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��ӻ�Ա��Ʒ<br> </div>')"
	target="ContentFrame">��ӻ�Ա��Ʒ</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Horder") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/History_order.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��������<br> </div>')"
	target="ContentFrame" class="lefttop">��������</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_GUser") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/Group_manage.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��Ա��<br> </div>')"
	target="ContentFrame" class="lefttop">��Ա��</a>��<a href="user/Group_Add.asp" target="ContentFrame" class="lefttop">���</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Jubao") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/UserReport.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>�ٱ�����<br> </div>')"
	target="ContentFrame">�ٱ�����</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Review") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/Review.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>���۹���<br> </div>')"
	target="ContentFrame">���۹���</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Log") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/iLog.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>�ռ���ժ<br> </div>')"
	target="ContentFrame">�ռ���ժ</a></td>
  </tr>
   <% end if %>
  <%if MF_Check_Pop_TF("ME_Photo") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
    <td valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/Photo.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>������<br> </div>')"
	target="ContentFrame">������</a></td>
  </tr>
    <% end if %>
  <%if MF_Check_Pop_TF("ME_Param") then%>
<tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td width="3%" valign="top"> <div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td width="97%"><a href="user/UserParam.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��������<br> </div>')"
	target="ContentFrame" class="lefttop">��������</a></td>
  </tr>
   <% end if %>
  <%if MF_Check_Pop_TF("ME_Pay") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20" valign="top"><div align="center"><img src="Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="user/PayParam.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>����֧��<br> </div>')"
	target="ContentFrame">����֧��</a></td>
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
    <td width="98%" class="titledaohang"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_sd);"  language=javascript>����ϵͳ</td>
  </tr>
  <tr>
    <td colspan="2">
	<table width="100%" border="0" cellpadding="2" cellspacing="0" id="sub_sd" style="display:none;">
	<%if MF_Check_Pop_TF("SD_List") then%>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22"><a  href="supply/News.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>������Ϣ<br> </div>')"
		target="ContentFrame">������Ϣ</a></td>
	  </tr>
	<%End IF%>
	<%if  MF_Check_Pop_TF("SD_Class") then %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22"><a  href="supply/Class.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>�����Ŀ<br> </div>')"
		target="ContentFrame">�����Ŀ</a></td>
	  </tr>
	<%End IF%>
	<%if MF_Check_Pop_TF("AP_area") then %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22"><a href="supply/Area.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>�������<br> </div>')"
		target="ContentFrame">�������</a></td>
	  </tr>
	 <%End IF%>
	 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:none;"> 
		<td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22"><a  href="supply/Lable.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��ǩ����<br> </div>')"
		target="ContentFrame">��ǩ����</a></td>
	  </tr>
	
	<%if MF_Check_Pop_TF("AP_param") then %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
		<td> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22"><a  href="supply/Config.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>ϵͳ����<br> </div>')"
		target="ContentFrame">ϵͳ����</a></td>
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
    <td width="98%" class="titledaohang"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_ap);"  language=javascript>�˲���Ƹ</td>
  </tr>
  <tr>
    <td colspan="2"><table width="100%" border="0" cellpadding="2" cellspacing="0" id="sub_ap" style="display:none;">
  <%if MF_Check_Pop_TF("AP_Param") then%>
   <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="job/FS_AP_SysPara.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>ϵͳ��������<br> </div>')"
	target="ContentFrame" class="lefttop">ϵͳ��������</a></td>
  </tr>
  <%
  End if
  if MF_Check_Pop_TF("AP_Trade") then
  %>
   <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="job/AP_Trade.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>ʡ������<br> </div>')"
	target="ContentFrame" class="lefttop">�˲���ҵ����</a></td>
  </tr>
  <%
  End if
  if MF_Check_Pop_TF("AP_Job") then
  %>
   <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="job/AP_Job.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>ʡ������<br> </div>')"
	target="ContentFrame" class="lefttop">��ҵְλ����</a></td>
  </tr>
  <%
  End if
  if MF_Check_Pop_TF("AP_Province") then
  %>
   <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="job/AP_Province.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>ʡ������<br> </div>')"
	target="ContentFrame" class="lefttop">ʡ������</a></td>
  </tr>
  <%
  End if
  if MF_Check_Pop_TF("AP_city") then
  %>
   <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="job/AP_City.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��������<br> </div>')"
	target="ContentFrame" class="lefttop">��������</a></td>
  </tr>
  <%
  End if
  if MF_Check_Pop_TF("AP_Search") then
  %>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td width="3%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="job/AP_Payment_Search.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��Ա��¼��ѯ<br> </div>')"
	target="ContentFrame" class="lefttop">��Ա��¼��ѯ</a></td>
  </tr>
  <%
  End if
  if MF_Check_Pop_TF("AP_check") then
  %>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
    <td height="22" colspan="2"><a href="job/AP_Register_List.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>ע����Ϣ���<br> </div>')"
	target="ContentFrame" class="lefttop">ע����Ϣ���</a></td>
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
    <td width="98%" class="titledaohang"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_hs);"  language=javascript>����¥��</td>
  </tr>
  <tr>
    <td colspan="2">
	<table width="95%" border="0" align="center" cellpadding="2" cellspacing="0" id="sub_hs" style="display:none">
	  <%if MF_Check_Pop_TF("HS_Loup") then%>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td width="2%"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td width="196%" height="22" colspan="2"><a href="house/HS_Quotation_Open.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>¥����Ϣ<br> </div>')"
		target="ContentFrame" class="lefttop">¥����Ϣ</a></td>
	  </tr>
	  <%
	  End if
	  if MF_Check_Pop_TF("HS_Ero") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="house/HS_Second.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>���ַ�����<br> </div>')"
		target="ContentFrame" class="lefttop">���ַ�����</a></td>
	  </tr>
	  <%
	  End if
	  if MF_Check_Pop_TF("HS_Zu") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
		<td> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="house/HS_Tenancy.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>������Ϣ<br> </div>')"
		target="ContentFrame" class="lefttop">������Ϣ</a></td>
	  </tr>
	  <%
	  End if
	  if MF_Check_Pop_TF("HS_param") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="23"> <div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="23" colspan="2"><a href="house/Sys_Para_Set.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>ϵͳ��������<br> </div>')"
		target="ContentFrame">ϵͳ��������</a></td>
	  </tr>
	  <%
	  End if
	  %>
	  <%
	  if MF_Check_Pop_TF("HS013") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td height="22"><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td height="22" colspan="2"><a href="house/HS_Recyle.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>����վ����<br> </div>')"
		class="lefttop" target="ContentFrame">����վ����</a></td>
	  </tr>
	  <%
	  End if
	  if MF_Check_Pop_TF("HS014") or MF_Check_Pop_TF("HS_Search") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td><div align="center"><img src="Images/Folder/folderclosed.gif"></div></td>
		<td colspan="2"><a href="house/HS_Clear.asp"	onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>������ڷ�Դ<br> </div>')"
		target="ContentFrame" class="lefttop">������ڷ�Դ</a></td>
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
    <td width="98%" class="titledaohang"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_other);"  language=javascript>����ϵͳ</td>
  </tr>
  <tr style="display:none" id="sub_other">
    <td colspan="2">
	<%
	if Request.Cookies("FoosunSUBCookie")("FoosunSUBCS")=1 Then
	%>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0">
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
          <td width="100%" height="19"   id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_cs);"  onmouseover="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>�ɼ�����<br> </div>')"
		  language=javascript><img src="Images/+.gif" width="15" height="15">�ɼ�����</td>
        </tr>
      </table>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0" id="sub_cs" style="display:none">
         <%if MF_Check_Pop_TF("CS_site") then%>
		 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td><div align="center"><img src="Images/L-.gif" width="17" height="16"></div></td>
          <td width="96%" height="16"><a href="collect/Site.asp"
		  target="ContentFrame">վ������</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="4%"><div align="center"><img src="Images/L-.gif" width="17" height="16"></div></td>
          <td height="16"><a href="collect/Rule.asp" target="ContentFrame">�ؼ��ֹ���</a></td>
        </tr>
		<%end if%>
		<%if MF_Check_Pop_TF("CS_Ink") then%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="16"><div align="center"><img src="Images/L-.gif" width="17" height="16"></div></td>
          <td height="16"><a href="collect/Check.asp" target="ContentFrame">���Ŵ���</a></td>
        </tr>
		<%end if%>
		 <%if MF_Check_Pop_TF("CS_site") then%>
		 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td><div align="center"><img src="Images/L-.gif" width="17" height="16"></div></td>
          <td width="96%" height="16"><a href="collect/AutoCollect.asp" target="ContentFrame">��ʱ�ɼ�</a></td>
        </tr>
		<%end if%>
      </table>
		<%
		End if
		if Request.Cookies("FoosunSUBCookie")("FoosunSUBSS")=1 Then
		%>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0">
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
          <td width="100%" height="19"   id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_ss);" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>վ��ͳ��<br> </div>')"
		  language=javascript><img src="Images/+.gif" width="15" height="15">վ��ͳ��</td>
        </tr>
      </table>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0" style="display:none" id="sub_ss">
        <%if MF_Check_Pop_TF("SS_site") then%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="4%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="96%"><a href="Stat/Visit_ObtainCode.asp" target="ContentFrame" class="lefttop">��ȡ����</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="4%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="96%"><a href="Stat/Visit_DataStatistic.asp" target="ContentFrame" class="lefttop">��Ҫ����</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="4%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="96%"><a href="Stat/Visit_HoursStatistic.asp" target="ContentFrame" class="lefttop">24Сʱͳ��</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="20"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_DaysStatistic.asp" target="ContentFrame">��ͳ��</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="16"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_MonthsStatistic.asp" target="ContentFrame">��ͳ��</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="16"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_SystemStatistic.asp" target="ContentFrame">ϵͳ/�����</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="25"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_AreaStatistic.asp" target="ContentFrame">����ͳ��</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="25"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_SourceStatistic.asp" target="ContentFrame">��Դͳ��</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="25"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_VisitorList.asp" target="ContentFrame">��������ͳ��</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:none;">
          <td height="25"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_SearchStatistic.asp" target="ContentFrame">��������</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:none;">
          <td height="25"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_fbStatistic.asp" target="ContentFrame">�ֱ���ͳ��</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="25"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="SystemCheckplus.asp" target="ContentFrame">ϵͳ̽��</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td height="25"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="Stat/Visit_sysPara.asp" target="ContentFrame">��������</a></td>
        </tr>
        <%end if%>
      </table>
		<%
		End if
		if Request.Cookies("FoosunSUBCookie")("FoosunSUBVS")=1 Then
		%>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0">
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
          <td width="100%" height="19"   id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_vs);" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>ͶƱ����<br> </div>')"
		  language=javascript><img src="Images/+.gif" width="15" height="15">ͶƱ����</td>
        </tr>
      </table>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="1" id="sub_vs" style="display:none">
        <%if MF_Check_Pop_TF("VS_site") Then%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td height="22" colspan="2"><a href="vote/FS_VS_SysPara.asp" target="ContentFrame" class="lefttop">ϵͳ��������</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td height="22" colspan="2"><a href="vote/VS_Class.asp" target="ContentFrame" class="lefttop">ͶƱ��������</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td height="22" colspan="2"><a href="vote/VS_Theme.asp" target="ContentFrame" class="lefttop">ͶƱ��������</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td height="22" colspan="2"><a href="vote/VS_Items.asp" target="ContentFrame" class="lefttop">ͶƱѡ������</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td height="22" colspan="2"><a href="vote/VS_Steps.asp" target="ContentFrame" class="lefttop">�ಽͶƱ����</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td height="22" colspan="2"><a href="vote/VS_Items_Result.asp" target="ContentFrame" class="lefttop">ͶƱ�������</a></td>
        </tr>
        <%end if%>
      </table>
		<%
		End if
		if Request.Cookies("FoosunSUBCookie")("FoosunSUBAS")=1 Then
		%>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0">
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
          <td width="100%" height="19"   id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_as);" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>������<br> </div>')"
		  language=javascript><img src="Images/+.gif" width="15" height="15">������</td>
        </tr>
      </table>
	  <table width="95%" border="0" align="center" cellpadding="2" cellspacing="2" id="sub_as" style="display:none">
        <%if MF_Check_Pop_TF("AS001") then%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="97%"><a href="ads/Ads_Manage.asp" target="ContentFrame">������</a></td>
        </tr>
        <%
		end if%>
        <%if MF_Check_Pop_TF("AS002") then%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="ads/Ads_Count.asp" target="ContentFrame">���ͳ��</a></td>
        </tr>
        <%end if%>
        <%if MF_Check_Pop_TF("AS003") then%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td><a href="ads/Ads_ClassManage.asp" target="ContentFrame">�������</a></td>
        </tr>
        <%end if%>
      </table>
		<%
		End if
		if Request.Cookies("FoosunSUBCookie")("FoosunSUBWS")=1 Then
		%>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0">
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
          <td width="100%" height="19"   id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_ws);" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>���Թ���<br> </div>')"
		  language=javascript><img src="Images/+.gif" width="15" height="15">���Թ���</td>
        </tr>
      </table>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0" id="sub_ws" style="display:none">
        <%if MF_Check_Pop_TF("WS_site") then%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="97%"><a href="bbs/SysConfig.asp" target="ContentFrame" class="lefttop">��������</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="97%"><a href="bbs/NewTell.asp" target="ContentFrame" class="lefttop">�������</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="97%"><a href="bbs/ClassManager.asp" target="ContentFrame" class="lefttop">ϵͳ����</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
          <td width="97%"><a href="bbs/ClassMessageManager.asp" target="ContentFrame" class="lefttop">���Թ���</a></td>
        </tr>
        <%end if%>
      </table>
		<%
		End if
		if Request.Cookies("FoosunSUBCookie")("FoosunSUBFL")=1 Then
		%>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0">
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
          <td width="100%" height="19"   id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sub_fl);" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';return escape('<div align=\'left\'>��������<br> </div>')"
		  language=javascript><img src="Images/+.gif" width="15" height="15">��������</td>
        </tr>
      </table>
	  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="1" id="sub_fl" style="display:none">
  
	  <%if MF_Check_Pop_TF("FL_site") then%>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
		<td width="97%"><a href="Flink/Flink_Manage.asp" target="ContentFrame" class="lefttop">��ө�����</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("FL_site") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
		<td width="97%"><a href="Flink/Flik_Class.asp" target="ContentFrame" class="lefttop">�������</a></td>
	  </tr>
	  <%end if
	  if MF_Check_Pop_TF("FL_site") then
	  %>
	  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
		<td width="3%"><div align="center"><img src="Images/L.gif" width="17" height="16"></div></td>
		<td width="97%"><a href="Flink/Flink_SysPara.asp" target="ContentFrame" class="lefttop">��������</a></td>
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
    <td width="98%" class="titledaohang">ϵͳ��Ϣ</td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="105" colspan="2">��Ȩ���У�<a href="http://www.foosun.net" target="_blank">��Ѷ����</a><br>
      ���������<a href="http://www.foosun.cn" target="_blank">Foosun Inc.</a><br>
      ����֧�֣�<a href="http://bbs.foosun.net" target="_blank">��Ѷ��̳</a><br>
      �������ģ�<a href="http://help.foosun.net" target="_blank">��Ѷ����</a> <br>
      ϵͳ�汾��5.0.0</td>
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






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
<body scroll=yes style="margin:0px;" onselectstart="return false;" class="Leftback">
<!--#include file="../CommPages/Main_Navi.asp" -->
<table width="95%" border="0" align="center" cellpadding="2" cellspacing="2" class="leftframetable">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td><div align="center"><img src="../Images/Folder/sub_sys.gif" width="15" height="17"></div></td>
    <td class="titledaohang">��Աϵͳ����</td>
  </tr>
  <%if MF_Check_Pop_TF("ME_List") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="User_manage.asp" target="ContentFrame" class="lefttop">���˻�Ա����</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="UserCorp.asp" target="ContentFrame" class="lefttop">��ҵ��Ա����</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Intergel") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="Integral.asp" target="ContentFrame" class="lefttop">���ֹ���</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Card") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="Card.asp" target="ContentFrame" class="lefttop">�㿨����</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_News") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="news_manage.asp" target="ContentFrame" class="lefttop">�������</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Form") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="GroupDebate_manage.asp" target="ContentFrame" class="lefttop">��Ⱥ����</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_HY") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="VocationClass.asp" target="ContentFrame" class="lefttop">��ҵ����</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_award") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="award.asp" target="ContentFrame" class="lefttop">�齱����</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Order") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="Order_Pay.asp" target="ContentFrame">��������(����֧��)</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Mproducts") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
    <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="get_Thing.asp" target="ContentFrame">��ӻ�Ա��Ʒ</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Horder") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="History_order.asp" target="ContentFrame" class="lefttop">��������</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_GUser") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="Group_manage.asp" target="ContentFrame" class="lefttop">��Ա��</a>��<a href="Group_Add.asp" target="ContentFrame" class="lefttop">���</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Jubao") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="UserReport.asp" target="ContentFrame">�ٱ�����</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Review") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="Review.asp" target="ContentFrame">���۹���</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Log") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="iLog.asp" target="ContentFrame">�ռ���ժ</a></td>
  </tr>
   <% end if %>
  <%if MF_Check_Pop_TF("ME_Photo") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
    <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="Photo.asp" target="ContentFrame">������</a></td>
  </tr>
    <% end if %>
  <%if MF_Check_Pop_TF("ME_Param") then%>
<tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td width="3%" valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td width="97%"><a href="UserParam.asp" target="ContentFrame" class="lefttop">��������</a></td>
  </tr>
   <% end if %>
  <%if MF_Check_Pop_TF("ME_Pay") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20" valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="PayParam.asp" target="ContentFrame">����֧��</a></td>
  </tr>
   <% end if %>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20" valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="../shortCutMenu.asp" target="_self" class="admintx">������һ��</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="133">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>






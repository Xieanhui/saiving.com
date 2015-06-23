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
<body scroll=yes style="margin:0px;" onselectstart="return false;" class="Leftback">
<!--#include file="../CommPages/Main_Navi.asp" -->
<table width="95%" border="0" align="center" cellpadding="2" cellspacing="2" class="leftframetable">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td><div align="center"><img src="../Images/Folder/sub_sys.gif" width="15" height="17"></div></td>
    <td class="titledaohang">会员系统导航</td>
  </tr>
  <%if MF_Check_Pop_TF("ME_List") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="User_manage.asp" target="ContentFrame" class="lefttop">个人会员管理</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="UserCorp.asp" target="ContentFrame" class="lefttop">企业会员管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Intergel") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="Integral.asp" target="ContentFrame" class="lefttop">积分管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Card") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="Card.asp" target="ContentFrame" class="lefttop">点卡管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_News") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="news_manage.asp" target="ContentFrame" class="lefttop">公告管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Form") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="GroupDebate_manage.asp" target="ContentFrame" class="lefttop">社群管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_HY") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="VocationClass.asp" target="ContentFrame" class="lefttop">行业管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_award") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="award.asp" target="ContentFrame" class="lefttop">抽奖管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Order") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="Order_Pay.asp" target="ContentFrame">定单管理(在线支付)</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Mproducts") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
    <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="get_Thing.asp" target="ContentFrame">添加会员商品</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Horder") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="History_order.asp" target="ContentFrame" class="lefttop">交易明晰</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_GUser") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="Group_manage.asp" target="ContentFrame" class="lefttop">会员组</a>┆<a href="Group_Add.asp" target="ContentFrame" class="lefttop">添加</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Jubao") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="UserReport.asp" target="ContentFrame">举报管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Review") then%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="Review.asp" target="ContentFrame">评论管理</a></td>
  </tr>
  <% end if %>
  <%if MF_Check_Pop_TF("ME_Log") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="iLog.asp" target="ContentFrame">日记网摘</a></td>
  </tr>
   <% end if %>
  <%if MF_Check_Pop_TF("ME_Photo") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage">
    <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="Photo.asp" target="ContentFrame">相册管理</a></td>
  </tr>
    <% end if %>
  <%if MF_Check_Pop_TF("ME_Param") then%>
<tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td width="3%" valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td width="97%"><a href="UserParam.asp" target="ContentFrame" class="lefttop">参数设置</a></td>
  </tr>
   <% end if %>
  <%if MF_Check_Pop_TF("ME_Pay") then%>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20" valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="PayParam.asp" target="ContentFrame">在线支付</a></td>
  </tr>
   <% end if %>
 <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="20" valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif" ></div></td>
    <td><a href="../shortCutMenu.asp" target="_self" class="admintx">返回上一级</a></td>
  </tr>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td height="133">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>






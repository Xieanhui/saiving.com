<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
Dim Conn,special_rs
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,i
'---------------------------------分页定义
int_RPP=15 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"			'尾页
'-----------------------------------------
MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("DS018") then Err_Show
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>CMS5.0</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/prototype.js"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
<tr>
<td class="xingmu">下载系统--专区管理</td>
</tr>
<tr>
<td class="hback">
<a href="Down_Special_manage.asp">管理首页</a>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
<a href="Down_Special_Edit_Add.asp?act=add">新增专区</a>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
<a href="#" onClick="javascript:history.back()">后退</a>&nbsp;&nbsp;&nbsp;|
<a href="../../help?Lable=DS_Special" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a>
</td>
</tr>
</table>

<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr> 
    <td class="xingmu" width="4%" align="center">ID</td>
    <td width="39%" class="xingmu">专区名称[英文名称]</td>
    <td class="xingmu" align="center" width="10%">是否锁定</td>
    <td class="xingmu" align="center" width="19%">创建时间</td>
    <td class="xingmu" align="center" width="19%">操作</td>
    <td class="xingmu" align="center" width="9%"><input type="checkbox" name="specialList" value="" onClick="selectAll(document.all('specialList'))"/></td>
  </tr>
  <%
	Set special_rs=Server.CreateObject(G_FS_RS)
	special_rs.open "Select specialID,SpecialEName,SpecialCName,IsUrl,Addtime,[Domain],isLock from FS_DS_Special order by addTime desc",Conn,1,1
	If Not special_rs.eof then
	'分页使用-----------------------------------
		special_rs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo<=0 Then cPageNo=1
		If cPageNo>special_rs.PageCount Then cPageNo=special_rs.PageCount 
		special_rs.AbsolutePage=cPageNo
	End if
	for i=0 to int_RPP
		if special_rs.eof then exit for
		Response.Write("<tr>"&vbcrlf)
			Response.Write("<td class='hback' align='center'>"&special_rs("specialID")&"</td>")
		Response.Write("<td class='hback'><a href='Down_Special_Edit_Add.asp?act=edit&specialID="&special_rs("specialID")&"'>"&special_rs("SpecialCName")&"<span style='font-size:10px'>["&special_rs("SpecialEName")&"]</span></a></td>"&vbcrlf)
		if special_rs("isLock")=1 then
			Response.Write("<td class='hback' align='center' id='lockTD_"&special_rs("specialID")&"'><a href='#' onClick=""javascript:changeLockState(false,'"&special_rs("specialID")&"')"" style='color:red'>解锁</a></td>")
		else
			Response.Write("<td class='hback' align='center' id='lockTD_"&special_rs("specialID")&"'><a href='#' onClick=""javascript:changeLockState(true,'"&special_rs("specialID")&"')"">锁定</a></td>")
		End if
			Response.Write("<td class='hback' align='center'>"&special_rs("Addtime")&"</td>")
			Response.Write("<td class='hback' align='center'><a href='Down_Special_Edit_Add.asp?act=edit&specialID="&special_rs("specialID")&"'>修改</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href='#' onclick=""del('"&special_rs("specialID")&"')"">删除</a></td>")
			Response.Write("<td class='hback' align='center'><input type='checkbox' name='specialList' value='"&special_rs("specialid")&"'/></td>")
		special_rs.movenext
	next
%>
  <tr> 
    <td colspan="6" align="right" class="hback"> <button onClick="chageLockStateBat(true)">批量锁定</button>
      <button onClick="chageLockStateBat(false)">批量解锁</button>
      <button onClick="del('')">批量删除</button></td>
  </tr>
  <%
	Response.Write("<tr>"&vbcrlf)
	Response.Write("<td align='right' colspan='8'  class=""hback"">"&fPageCount(special_rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td>"&vbcrlf)
	Response.Write("</tr>"&vbcrlf)
%>
</table>
</body>
<script language="JavaScript">
//修改锁状态
function changeLockState(tf,specialid)
{
	if (isNaN(specialid)) return;
	var container=$('lockTD_'+specialid)
	if(tf)
	{
		var param="act=lock&specialid="+specialid
		var ajax=new Ajax.Updater(container,'Down_Special_Action.asp',{method:'get',parameters:param})
	}else
	{
		var param="act=unlock&specialid="+specialid
		var ajax=new Ajax.Updater(container,'Down_Special_Action.asp',{method:'get',parameters:param})
	}
}
//批量改变锁定状态
function chageLockStateBat(tf)
{
	var element=document.all("specialList")
	for(var i=1;i<element.length;i++)
	{	
		if(element[i].checked)
		{
				changeLockState(tf,element[i].value)
		}
	}
}
//
function del(specialID)
{
	if(specialID!="")
	{
		if(confirm("该操作不可恢复，确认删除该专区？"))
		{
			location="Down_Special_Action.asp?act=del&specialID="+specialID
		}
	}else
	{
		element=document.all("specialList");
		var specialID；
		for(var i=1;i<element.length;i++)
		{	
			if(element[i].checked)
			{
				specialID+=element[i].value+","
			}
		}
		if(specialID.length>1)
		{
			if(confirm("该操作不可恢复，确认删除该专区？"))
			{
				location="Down_Special_Action.asp?act=del&specialID="+specialID
			}
		}else
		{
			alert("请选择要删除的专区")
		}
	}
}
</script>
</html>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->







<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp"-->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../../FS_InterFace/NS_Function.asp" -->
<%
Dim Conn
Dim TypeSql,RsTypeObj,LableSql,RsLableObj
MF_Default_Conn
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>标签插入</title>
</head>
<style>
.LableSelectItem {
	background-color:highlight;
	cursor: hand;
	color: white;
	text-decoration: underline;
}
.LableItem {
	cursor: hand;
}
</style>
<link href="../../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<body ondragstart="return false;" onselectstart="return false;" topmargin="0" leftmargin="0">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="0">
  <%
TypeSql = "Select SpecialCName,SpecialEName from FS_NS_Special where isLock=0"
Set RsTypeObj = Conn.Execute(TypeSql)
if Not RsTypeObj.Eof then
	do while Not RsTypeObj.Eof
%>
  <tr>
	<td><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td valign="top">
		<%
		Dim str_action,obj_news_rs_1
		str_action=  "<img src=""../images/+.gif""></img>"
		Response.Write str_action
		  %></td>
          <td valign="bottom"><span onDblClick="SubmitLable(this);" Extend="False" class="LableItem" TypeID="<% = RsTypeObj("SpecialEName") %>" onClick="SelectFolder(this)">
		<% = RsTypeObj("SpecialCName") %></span></td>
        </tr>
      </table>
    </td>
</tr>
<%
		RsTypeObj.MoveNext
	loop
end if
%>
</table>
</body>
</html>
<%
Set Conn = Nothing
%>
<script language="JavaScript">
function SelectLable(Obj)
{
	for (var i=0;i<document.all.length;i++)
	{
		if (document.all(i).className=='LableSelectItem') document.all(i).className='LableItem';
	}
	Obj.className='LableSelectItem';
}
function SelectFolder(Obj)
{
	var CurrObj;
	for (var i=0;i<document.all.length;i++)
	{
		if (document.all(i).className=='LableSelectItem') document.all(i).className='LableItem';
	}
	Obj.className='LableSelectItem';
	if (Obj.Extend=='True')
	{
		ShowOrDisplay(Obj,'none',true);
		Obj.Extend='False';
	}
	else
	{
		ShowOrDisplay(Obj,'',false);
		Obj.Extend='True';
	}
}
function ShowOrDisplay(Obj,Flag,Tag)
{
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if (CurrObj.ParentID==Obj.TypeID)
		{
			CurrObj.style.display=Flag;
			if (Tag) 
			if (CurrObj.TypeFlag=='Class') ShowOrDisplay(CurrObj.children(0).children(0).children(0).children(0).children(1).children(0),Flag,Tag);
		}
	}
}
function SubmitLable(Obj)
{
	var LableName=Obj.TypeID+'***'+Obj.innerText;
	window.returnValue=LableName;
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>






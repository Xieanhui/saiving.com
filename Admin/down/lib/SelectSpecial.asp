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
	background-color:#333333;
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
TypeSql = "Select specialID,SpecialCName,ParentID,isUrl from FS_DS_Special where ParentID=0 and isUrl=0 order by specialID desc"
Set RsTypeObj = Conn.Execute(TypeSql)
if Not RsTypeObj.Eof then
	do while Not RsTypeObj.Eof
%>
  <tr ParentID="<% = RsTypeObj("ParentID") %>">
	<td><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td valign="top">
		  <%
		Dim str_action,obj_news_rs_1
		Set obj_news_rs_1 = server.CreateObject(G_FS_RS)
		obj_news_rs_1.Open "Select Count(specialID) from FS_DS_Special where ParentID="& RsTypeObj("specialID"),Conn,1,1
		if obj_news_rs_1(0)>0 then
			str_action=  "<img src=""../../images/+.gif""></img>"
		Else
			str_action=  "<img src=""../../images/-.gif""></img>"
		End if
		obj_news_rs_1.close:set obj_news_rs_1 =nothing
		Response.Write str_action
		  %></td>
          <td valign="bottom"><span onDblClick="SubmitLable(this);" Extend="False" class="LableItem" TypeID="<% = RsTypeObj("specialID") %>" onClick="SelectFolder(this)">
<% = RsTypeObj("SpecialCName") %></span></td>
        </tr>
      </table>
    </td>
</tr>
<%
		Response.Write(GetChildTypeList(RsTypeObj("specialID"),""," style=""display:none;"" "))
		RsTypeObj.MoveNext
	loop
end if
%>
</table>
</body>
</html>
<%
Set Conn = Nothing
Function GetChildTypeList(TypeID,CompatStr,ShowStr)
	Dim ChildTypeListRs,ChildTypeListStr,TempStr
	Set ChildTypeListRs = Conn.Execute("Select * from FS_DS_Special where ParentID=" & CintStr(TypeID) & " and IsUrl=0" )
	TempStr = CompatStr & "&nbsp;&nbsp;&nbsp;&nbsp;"
	do while Not ChildTypeListRs.Eof
		Dim str_action_1,obj_news_rs_1s
		Set str_action_1 = server.CreateObject(G_FS_RS)
		str_action_1.Open "Select Count(SpecialID) from FS_DS_Special where ParentID="& CintStr(ChildTypeListRs("specialID")),Conn,1,1
		if str_action_1(0)>0 then
			str_action_1=  "<img src=""../img/+.gif""></img>"
		Else
			str_action_1=  "<img src=""../img/-.gif""></img>"
		End if
		set obj_news_rs_1s =nothing
	  	GetChildTypeList = GetChildTypeList & "<tr TypeFlag=""Class"" ParentID=""" & ChildTypeListRs("ParentID") & """ " & ShowStr & ">" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & "<td>" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & Chr(13) & Chr(10) & "<tr>"  & Chr(13) & Chr(10) & "<td>" & TempStr & ""& str_action_1 &"</td>"
		GetChildTypeList = GetChildTypeList & "<td><span onDblClick=""SubmitLable(this);"" class=""LableItem"" TypeID=""" & ChildTypeListRs("specialID") & """ Extend=""False"" onClick=""SelectFolder(this)"">" & ChildTypeListRs("SpecialCName") & "</span></td>" & Chr(13) & Chr(10) & "</tr>" & Chr(13) & Chr(10) & "</table>"
		GetChildTypeList = GetChildTypeList & "</td>" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & "</tr>" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & GetChildTypeList(ChildTypeListRs("specialID"),TempStr,ShowStr)
		ChildTypeListRs.MoveNext
	loop
	ChildTypeListRs.Close
	Set ChildTypeListRs = Nothing
End Function
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






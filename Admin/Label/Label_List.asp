<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Dim Conn
	MF_Default_Conn
	MF_Session_TF 
Dim TypeSql,RsTypeObj,LableSql,RsLableObj
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
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<body ondragstart="return false;" onselectstart="return false;" topmargin="0" leftmargin="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
<%
TypeSql = "Select Id,ClassName,ParentID from FS_MF_LableClass where ParentID=0"
Set RsTypeObj = Conn.Execute(TypeSql)
if Not RsTypeObj.Eof then
	do while Not RsTypeObj.Eof
%>
  <tr ParentID="<% = RsTypeObj("ParentID") %>">
	<td><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="21" valign="top"><img src="../Images/Folder/folder.gif" width="20" height="16"></td>
          <td width="954" valign="bottom"><span Extend="False" class="LableItem" TypeID="<% = RsTypeObj("ID") %>" onClick="SelectFolder(this)">
          <% = RsTypeObj("ClassName") %></span></td>
        </tr>
      </table>
    </td>
</tr>
<%
		Response.Write(GetLableList(RsTypeObj("ID"),"&nbsp;&nbsp;&nbsp;&nbsp;"," style=""display:none;"" "))
		Response.Write(GetChildTypeList(RsTypeObj("ID"),""," style=""display:none;"" "))
		RsTypeObj.MoveNext
	loop
end if
Response.Write(GetLableList(0,"",""))
%>
</table>
</body>
</html>
<%
Set Conn = Nothing
Function GetLableList(TypeID,CompatStr,ShowStr)
	Dim ListSql,RsListObj,TempStr
	ListSql = "Select ID,LableName,LableClassID from FS_MF_Lable where LableClassID=" & CintStr(TypeID)
	Set RsListObj = Conn.Execute(ListSql)
	TempStr = CompatStr & "&nbsp;"
	do while Not RsListObj.Eof
	  	GetLableList = GetLableList & "<tr ParentID=""" & RsListObj("LableClassID") & """ " & ShowStr & ">" & Chr(13) & Chr(10)
		GetLableList = GetLableList & "<td>" & Chr(13) & Chr(10)
		GetLableList = GetLableList & "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & Chr(13) & Chr(10) & "<tr>"  & Chr(13) & Chr(10) & "<td valign=""top"" align=""right"">" & CompatStr & "<img src=""../Images/L.gif""></td>"
		GetLableList = GetLableList & "<td  valign=""bottom""><span class=""LableItem"" LableName=""" & RsListObj("LableName") & """ onclick=""SelectLable(this);"" onDblClick=""SubmitLable(this)"">" & Replace(Replace(RsListObj("LableName"),"{FS400_",""),"}","") & "</span></td>" & Chr(13) & Chr(10) & "</tr>" & Chr(13) & Chr(10) & "</table>"
		GetLableList = GetLableList & "</td>" & Chr(13) & Chr(10)
		GetLableList = GetLableList & "</tr>" & Chr(13) & Chr(10)
		RsListObj.MoveNext
	Loop
	Set RsListObj = Nothing
End Function
Function GetChildTypeList(TypeID,CompatStr,ShowStr)
	Dim ChildTypeListRs,ChildTypeListStr,TempStr
	Set ChildTypeListRs = Conn.Execute("Select ID,ClassName,ParentID from FS_MF_LableClass where ParentID=" & CintStr(TypeID))
	TempStr = CompatStr & "&nbsp;"
	do while Not ChildTypeListRs.Eof
	  	GetChildTypeList = GetChildTypeList & "<tr TypeFlag=""Class"" ParentID=""" & ChildTypeListRs("ParentID") & """ " & ShowStr & ">" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & "<td>" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & Chr(13) & Chr(10) & "<tr>"  & Chr(13) & Chr(10) & "<td>" & TempStr & "<img src=""../Images/Folder/folder.gif""></td>"
		GetChildTypeList = GetChildTypeList & "<td><span class=""LableItem"" TypeID=""" & ChildTypeListRs("ID") & """ Extend=""False"" onClick=""SelectFolder(this)"">" & ChildTypeListRs("ClassName") & "</span></td>" & Chr(13) & Chr(10) & "</tr>" & Chr(13) & Chr(10) & "</table>"
		GetChildTypeList = GetChildTypeList & "</td>" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & "</tr>" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & Chr(13) & Chr(10) & GetLableList(ChildTypeListRs("ID"),"&nbsp;&nbsp;&nbsp;&nbsp;" & TempStr,ShowStr)
		GetChildTypeList = GetChildTypeList & GetChildTypeList(ChildTypeListRs("ID"),TempStr,ShowStr)
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
	var LableName=Obj.LableName;
	window.returnValue=LableName;
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>






<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp"-->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<%
Dim Conn
MF_Default_Conn
MF_Session_TF 
Dim CurrPath,FsoObj,FolderObj,SubFolderObj,FileObj,i,FsoItem
Dim ParentPath,FileExtName,AllowShowExtNameStr
AllowShowExtNameStr = "htm,html,shtml"
CurrPath = Request("CurrPath")
if CurrPath = "" then
	CurrPath = "/"
end if
Set FsoObj = Server.CreateObject(G_FS_FSO)
Set FolderObj = FsoObj.GetFolder(Server.MapPath(CurrPath))
Set SubFolderObj = FolderObj.SubFolders
Set FileObj = FolderObj.Files
Function CheckFileShowTF(AllowShowExtNameStr,ExtName)
	if ExtName="" then
		CheckFileShowTF = False
	else
		if InStr(1,AllowShowExtNameStr,ExtName) = 0 then
			CheckFileShowTF = False
		else
			CheckFileShowTF = True
		end if
	end if
End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><% = CurrPath %>目录文件列表</title>
</head>
<link href="../../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<style>
.TempletItem {
	cursor: default;
}
.TempletSelectItem {
	background-color:highlight;
	cursor: default;
	color: white;
}</style>
<body topmargin="0" leftmargin="0" scroll=yes>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr class="xingmu"> 
    <td width="43%" height="20" class="xingmu"> <div align="left">名称</div></td>
    <td width="26%" height="20" class="xingmu"> <div align="center">类型</div></td>
    <td width="31%" height="20" class="xingmu"> <div align="center">修改日期</div></td>
  </tr>
<%
for Each FsoItem In SubFolderObj
%>
  <tr> 
    <td height="20"> 
        <table width="95%" height="19" border="0" cellpadding="0" cellspacing="0">
        <tr title="双击鼠标进入此目录" style="cursor:hand;"> 
          <td width="6">&nbsp;</td>
          <td width="20"> <span class="TempletItem" Path="<% = FsoItem.name %>" onDblClick="OpenFolder(this);" onClick="SelectFolder(this);"><img src="../../Images/Folder/folder.gif" height="16"> 
            </span> </td>
          <td width="433"><span class="TempletItem" Path="<% = FsoItem.name %>" onDblClick="OpenFolder(this);" onClick="SelectFolder(this);">
            <font style="font-size:12px">
            <% = FsoItem.name %></font>
          </span></td>
        </tr>
      </table>
      </div></td>
    <td height="20"> 
      <div align="center">文件夹</div></td>
    <td height="20"> 
      <div align="center"><font style="font-size:12px"><% = FsoItem.Size %></font></div></td>
  </tr>
<font style="font-size:12px"><%
Next
for each FsoItem In FileObj
	FileExtName = LCase(Mid(FsoItem.name,InstrRev(FsoItem.name,".")+1))
	if CheckFileShowTF(AllowShowExtNameStr,FileExtName) = True then
%></font>
  <tr title="单击选择文件"> 
    <td height="20"> 
      <table width="99%" border="0" cellspacing="0" cellpadding="0">
        <tr style="cursor:hand;">
          <td width="3%"><img src="../../Images/Folder/folder_1.gif" width="20" height="16"></td>
          <td width="97%"><span class="TempletItem" File="<% = FsoItem.name %>" onClick="SelectFile(this);">
            <font style="font-size:12px"><% = FsoItem.name %></font>
            </span></td>
        </tr>
      </table>
    </td>
    <td height="20"> <div align="center"> 
        <font style="font-size:12px"><% = FsoItem.Type %></font>
      </div></td>
    <td height="20"> <div align="center"> 
        <font style="font-size:12px"><% = FsoItem.DateLastModified %></font>
      </div></td>
  </tr>
  <%
  	end if
next
%>
</table>
</body>
</html>
<%
Set FsoObj = Nothing
Set FolderObj = Nothing
Set FileObj = Nothing
%>
<script language="JavaScript">
var CurrPath='<% = CurrPath %>';
var FileName='';
function SelectFile(Obj)
{
	for (var i=0;i<document.all.length;i++)
	{
		if (document.all(i).className=='TempletSelectItem') document.all(i).className='TempletItem';
	}
	Obj.className='TempletSelectItem';
	FileName=Obj.File;
}
function SelectFolder(Obj)
{
	for (var i=0;i<document.all.length;i++)
	{
		if (document.all(i).className=='TempletSelectItem') document.all(i).className='TempletItem';
	}
	Obj.className='TempletSelectItem';
}
function OpenFolder(Obj)
{
	var SubmitPath='';
	if (CurrPath=='/') SubmitPath=CurrPath+Obj.Path;
	else SubmitPath=CurrPath+'/'+Obj.Path;
	location.href='FolderList.asp?CurrPath='+SubmitPath;
	AddFolderList(parent.document.all.FolderSelectList,SubmitPath,SubmitPath);
}
function AddFolderList(SelectObj,Lable,LableContent)
{
	var i=0,AddOption;
	if (!SearchOptionExists(SelectObj,Lable))
	{
		AddOption = document.createElement("OPTION");
		AddOption.text=Lable;
		AddOption.value=LableContent;
		SelectObj.add(AddOption);
		SelectObj.options(SelectObj.length-1).selected=true;

	}
}

function SearchOptionExists(Obj,SearchText)
{
	var i;
	for(i=0;i<Obj.length;i++)
	{
		if (Obj.options(i).text==SearchText)
		{
			Obj.options(i).selected=true;
			return true;
		}
	}
	return false;
}
</script>






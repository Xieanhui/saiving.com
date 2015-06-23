<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<%
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
<link href="../../FS_Inc/FS_css.css" rel="stylesheet">
<body topmargin="0" leftmargin="0" scroll=yes>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="20" class="ButtonListLeft"> <div align="center"><font color="#000000">名称</font></div></td>
    <td height="20" class="ButtonList"> <div align="center"><font color="#000000">类型</font></div></td>
    <td height="20" class="ButtonList"> <div align="center"><font color="#000000">修改日期</font></div></td>
  </tr>
<%
for Each FsoItem In SubFolderObj
%>
  <tr> 
    <td height="20"> 
        <table border="0" cellspacing="0" cellpadding="0">
          <tr title="双击鼠标进入此目录"> 
            
          <td><img src="../Images/Folder/folderclosed.gif"></td>
            <td> <span class="TempletItem" Path="<% = FsoItem.name %>" onDblClick="OpenFolder(this);" onClick="SelectFolder(this);"> 
              <% = FsoItem.name %>
              </span> </td>
          </tr>
        </table>
      </div></td>
    <td height="20"> 
      <div align="center">文件夹</div></td>
    <td height="20"> 
      <div align="center"><% = FsoItem.Size %></div></td>
  </tr>
    <%
Next
for each FsoItem In FileObj
	FileExtName = LCase(Mid(FsoItem.name,InstrRev(FsoItem.name,".")+1))
	if CheckFileShowTF(AllowShowExtNameStr,FileExtName) = True then
%>
  <tr title="单击选择文件"> 
    <td height="20"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="3%">&nbsp;</td>
          <td width="97%"><span class="TempletItem" File="<% = FsoItem.name %>" onClick="SelectFile(this);">
            <% = FsoItem.name %>
            </span></td>
        </tr>
      </table>
    </td>
    <td height="20"> <div align="center"> 
        <% = FsoItem.Type %>
      </div></td>
    <td height="20"> <div align="center"> 
        <% = FsoItem.DateLastModified %>
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






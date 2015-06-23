<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp"-->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<%
Dim Conn
MF_Default_Conn
MF_Session_TF
Dim FsoObj,OType
Set FsoObj = Server.CreateObject(G_FS_FSO)
OType = Request("Type")
Dim CurrPath,SubFolderObj,FolderObj,i,FsoItem,sRootDir
Dim ParentPath
if OType <> "" then
	Dim Path
	if OType = "Del" then
		if not MF_Check_Pop_TF("MF027") then Err_Show
		Path = Request("Path") 
		Path = Replace(Path,"//","/")
		If Instr(Path,".") > 0 Then
			Response.Write "<script>alert('需要删除的目录路径不正确哦');</script>"
			Response.End
		End If
		if Path <> "" then
			Path = Server.MapPath(Path)
			if FsoObj.FolderExists(Path) = true then FsoObj.DeleteFolder Path
		end if
	elseif OType = "AddFolder" then
		if not MF_Check_Pop_TF("MF028") then Err_Show
		Path = Request("Path")
		Path = Replace(Path,"//","/")
		if Path <> "" then
			Dim FildNameStr
			If Right(Path,1) = "/" Then
				Path = Left(Path,Len(Path) - 1)
			End If
			FildNameStr = Split(Path,"/")(Ubound(Split(Path,"/")))
			If ReplaceExpChar(FildNameStr) = False Then
				Response.Write "<script>alert('新的目录名不规范，请重设');window.location.href='FolderImageList.asp';</script>"
				Response.End
			End If
			Path = Server.MapPath(Path)
			if FsoObj.FolderExists(Path) = True then
				Response.Write("<script>alert('目录已经存在');</script>")
			else
				FsoObj.CreateFolder Path
			end if
		end if
	elseif OType = "FolderReName" then
		if not MF_Check_Pop_TF("MF026") then Err_Show
		Dim NewPathName,OldPathName,PhysicalPath,FileObj
		Path = Request("Path")
		Path = Replace(Path,"//","/")
		if Path <> "" then
			NewPathName = Request("NewPathName")
			OldPathName = Request("OldPathName")
			If ReplaceExpChar(NewPathName) = False Then
				Response.Write "<script>alert('新的目录名不规范，请重设');window.location.href='FolderImageList.asp';</script>"
				Response.End
			End If	
			if (NewPathName <> "") And (OldPathName <> "") then
				PhysicalPath = Server.MapPath(Path) & "\" & OldPathName
				if FsoObj.FolderExists(PhysicalPath) = True then
					PhysicalPath = Server.MapPath(Path) & "\" & NewPathName
					if FsoObj.FolderExists(PhysicalPath) = False then
						Set FileObj = FsoObj.GetFolder(Server.MapPath(Path) & "\" & OldPathName)
						FileObj.Name = NewPathName
						Set FileObj = Nothing
					end if
				end if
			end if
		end if
	end if
end if

CurrPath = Replace(trim(Request("CurrPath")),"//","/")
if CurrPath = "" then:response.Write("错误的参数"):response.end:end if
if CurrPath = "" then
	CurrPath = "/"
	ParentPath = ""
else
	if instr(CurrPath,"/") then ParentPath = Mid(CurrPath,1,InstrRev(CurrPath,"/")-1)
	if ParentPath = "" then
		ParentPath = "/"
	end if
end if
On Error Resume Next

if FsoObj.FolderExists(Server.MapPath(CurrPath))=false then FsoObj.CreateFolder(Server.MapPath(CurrPath))
Set FolderObj = FsoObj.GetFolder(Server.MapPath(CurrPath))
Set SubFolderObj = FolderObj.SubFolders
if err.number>0 then
	response.Write "错误，可能是找不到路径，或者你的服务器不支持FSO组件"
	response.end
end if


Function ReplaceExpChar(Str)
	Dim RegEx,StrRs,S_Str,ReturnV,HaveV
	S_Str = Str & ""
	ReturnV = True
	Set RegEx = New RegExp
	RegEx.IgnoreCase = True
	RegEx.Global=True
	RegEx.Pattern = "([^a-zA-Z0-9])"
	Set StrRs = RegEx.ExeCute(S_Str)
	Set RegEx = Nothing
	For Each HaveV In StrRs
		If Instr(S_Str,HaveV) <> 0 Then
			ReturnV = False
			Exit For
		End IF	
	Next
	ReplaceExpChar = ReturnV
End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>文件和目录列表</title>
</head>
<style>
.LableItem {
	cursor: default;
}
.LableSelectItem {
	background-color:highlight;
	cursor: default;
	color: white;
	text-decoration: underline;
}

</style>
<link href="../../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../../FS_Inc/PublicJS.js"></script>
<script language="JavaScript">
var ObjPopupMenu=window.createPopup();
document.oncontextmenu=new Function("return ShowMouseRightMenu(window.event);");
var DocumentReadyTF=false;
var ContentMenuArray=new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	InitialClassListContentMenu();
	DocumentReadyTF=true;
}
function ShowMouseRightMenu(event)
{
	ContentMenuShowEvent();
	var width=100;
	var height=0;
	var lefter=event.clientX;
	var topper=event.clientY;
	var ObjPopDocument=ObjPopupMenu.document;
	var ObjPopBody=ObjPopupMenu.document.body;
	var MenuStr='';
	if('<%=G_VIRTUAL_ROOT_DIR%>'=='') sRootDir='';else sRootDir='/<%=G_VIRTUAL_ROOT_DIR%>';
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (ContentMenuArray[i].ExeFunction=='seperator')
		{
			MenuStr+=FormatSeperator();
			height+=16;
		}
		else
		{
			MenuStr+=FormatMenuRow(ContentMenuArray[i].ExeFunction,ContentMenuArray[i].Description,ContentMenuArray[i].EnabledStr);
			height+=20;
		}
	}
	MenuStr="<TABLE border=0 cellpadding=0 cellspacing=0 class=Menu width=100>"+MenuStr
	MenuStr=MenuStr+"<\/TABLE>";
	ObjPopDocument.open();
	ObjPopDocument.write("<head><link href=\"select_css.css\" type=\"text/css\" rel=\"stylesheet\"></head><body scroll=\"no\" onConTextMenu=\"event.returnValue=false;\" onselectstart=\"event.returnValue=false;\">"+MenuStr);
	ObjPopDocument.close();
	height+=4;
	if(lefter+width > document.body.clientWidth) lefter=lefter-width;
	ObjPopupMenu.show(lefter, topper, width, height, document.body);
	return false;
}
function FormatSeperator()
{
	var MenuRowStr="<tr><td height=16 valign=middle><hr><\/td><\/tr>";
	return MenuRowStr;
}
function FormatMenuRow(MenuOperation,MenuDescription,EnabledStr)
{
	var MenuRowStr="<tr "+EnabledStr+"><td align=left height=20 class=MouseOut onMouseOver=this.className='MouseOver'; onMouseOut=this.className='MouseOut'; valign=middle"
	if (EnabledStr=='') MenuRowStr+=" onclick=\""+MenuOperation+"parent.ObjPopupMenu.hide();\">&nbsp;&nbsp;&nbsp;&nbsp;";
	else MenuRowStr+=">&nbsp;&nbsp;&nbsp;&nbsp;";
	MenuRowStr=MenuRowStr+MenuDescription+"<\/td><\/tr>";
	return MenuRowStr;
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function ContentMenuShowEvent()
{
	SelectFolder();
}
function InitialClassListContentMenu()
{
	<%Response.write "ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction(""parent.AddFolderOperation();"",'新建','');"%>
	<%Response.write "ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction(""if (confirm('确定要删除吗？')==true) parent.DelFolderFile();"",'删除','disabled');"%>
	<%Response.write "ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction(""parent.EditFolder();"",'重命名','disabled');"%>
}
</script>
<body topmargin="0" leftmargin="0" onClick="SelectFolder();">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <%
if Err.Number = 0 then
	for each FsoItem In SubFolderObj
%>
  <tr> 
    <td width="30%"><table border="0" cellspacing="0" cellpadding="0">
        <tr title="双击鼠标进入此目录"> 
          <td><img src="../../Images/Folder/folder.gif" width="20" height="16"></td>
          <td nowrap> <span class="LableItem" Path="<% = Replace(FsoItem.name,"//","/") %>" onDblClick="OpenFolder(this);"> 
            <% = Replace(FsoItem.name,"//","/") %>
            </span> </td>
        </tr>
      </table></td>
  </tr>
  <%
	next
else
%>
  <tr> 
    <td height="20"> <div align="center"> 
        <% = "路径不存在" %>
      </div></td>
  </tr>
  <%
end if
%>
</table>
</body>
</html>
<%
Set FsoObj = Nothing
Set SubFolderObj = Nothing
%>
<script language="JavaScript">
var CurrPath='<% = CurrPath %>';
var SelectedFolder='';
function SelectFolder()
{
	Obj=event.srcElement,DisabledContentMenuStr='';
	for (var i=0;i<document.all.length;i++)
	{
		if (document.all(i).className=='LableSelectItem') document.all(i).className='LableItem';
	}
	if (Obj.Path!=null)
	{
		Obj.className='LableSelectItem';
		SelectedFolder=Obj.Path;
	}
	else
	{
		SelectedFolder='';
	}
	if (SelectedFolder!='')
		DisabledContentMenuStr='';
	else
		DisabledContentMenuStr=',删除,重命名,';
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function OpenParentFolder(Obj)
{
	location.href='SelectPath.asp?CurrPath='+Obj.Path;
	SearchOptionExists(parent.document.all.FolderSelectList,Obj.Path);
}

function OpenFolder(Obj)
{
	var SubmitPath='';
	if (CurrPath=='/') SubmitPath=CurrPath+Obj.Path;
	else SubmitPath=CurrPath+'/'+Obj.Path;
	location.href='SelectPath.asp?CurrPath='+SubmitPath;
	AddFolderList(parent.document.all.FolderSelectList,SubmitPath,SubmitPath);
}

function SelectUpFolder(Obj)
{
	for (var i=0;i<document.all.length;i++)
	{
		if (document.all(i).className=='LableSelectItem') document.all(i).className='LableItem';
	}
	Obj.className='LableSelectItem';
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


function AddFolderOperation()
{
	var ReturnValue=prompt('新建目录名：','');
	if ((ReturnValue!='') && (ReturnValue!=null))
	{
		var patrn =/([^a-zA-Z0-9])/; 
		if (patrn.exec(ReturnValue))
		{
			alert('创建目录名不规范，请重设');
			return false;
		}
		else
		{
			window.location.href='?Type=AddFolder&Path='+CurrPath+'/'+ReturnValue+'&CurrPath='+CurrPath;
		}
	}
}
function DelFolderFile()
{
	var ReturnValue='';
	if (SelectedFolder!='') 
		window.location.href='?Type=Del&Path='+CurrPath+'/'+SelectedFolder+'&CurrPath='+CurrPath;
	else alert('请选择要删除的目录');
}
function EditFolder()
{
	if (SelectedFolder!='')
	{
		var ReturnValue=prompt('修改的目录名：',SelectedFolder);
		if ((ReturnValue!='') && (ReturnValue!=null))
		{
			var patrn =/([^a-zA-Z0-9])/; 
			if (patrn.exec(ReturnValue))
			{
				alert('修改目录名不规范，请重设');
				return false;
			}
			else
			{
				window.location.href='?Type=FolderReName&Path='+CurrPath+'&CurrPath='+CurrPath+'&OldPathName='+SelectedFolder+'&NewPathName='+ReturnValue;
			}
		}	
	}
	else alert('请填写要更名的目录名称');
}
</script>






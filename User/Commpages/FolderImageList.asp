<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

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

'if trim(Request("f_UserNumber"))="" then  response.Write("登陆过期,或者参数不正确"):response.end
'if trim(Request("f_UserNumber"))<>Fs_User.UserNumber then  response.Write("错误参数"):response.end
Dim TypeSql,RsTypeObj,LableSql,RsLableObj
Dim CurrPath,FsoObj,SubFolderObj,FolderObj,FileObj,i,FsoItem,OType
Dim ParentPath,FileExtName,AllowShowExtNameStr,str_CurrPath,sRootDir 
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/" & G_VIRTUAL_ROOT_DIR else sRootDir=""
CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
If ReplaceExpChar(Replace(CurrPath,"/","")) = False Then
	Response.Write "<script>alert('发生错误，为了分析错误原因?\n您的IP已经被记录!\n我们将收集您部分计算机资料');window.location.href='javascript:history.back();';</script>"
	Response.End
End If	
Set FsoObj = Server.CreateObject(G_FS_FSO)
OType = Request("Type")
if OType <> "" then
	Dim Path,PhysicalPath
	if OType = "DelFolder" then
		Path = Request("Path") 
		If Instr(Path,".") <> 0 Or Instr(Path,"/") = 0 Or Left(Path,1) <> "/" Then
			Response.Write "<script>alert('需要删除的目录路径不正确!');window.location.href='FolderImageList.asp';</script>"
			Response.End
		End If
		Path = Right(Path,Len(Path) - 1)
		If UBound(Split(Path,"/")) < 2 Then
			Response.Write "<script>alert('需要删除的目录路径不正确!');window.location.href='FolderImageList.asp';</script>"
			Response.End
		End If
		If Cstr(Split(Path,"/")(1)) <> Cstr(Session("FS_UserNumber")) Then
			Response.Write "<script>alert('不能删除别人的目录!');window.location.href='FolderImageList.asp';</script>"
			Response.End
		End If 	
		Path = "/" & Path
		Path = Server.MapPath(Path)
		if FsoObj.FolderExists(Path) = true then FsoObj.DeleteFolder Path
	elseif OType = "DelFile" then
		Dim DelFileName
		Path = Request("Path") 
		If Instr(Path,".") <> 0 Or Instr(Path,"/") = 0 Or Left(Path,1) <> "/" Then
			Response.Write "<script>alert('需要删除的文件路径不正确!');window.location.href='FolderImageList.asp';</script>"
			Response.End
		End If
		Path = Right(Path,Len(Path) - 1)
		If UBound(Split(Path,"/")) < 1 Then
			Response.Write "<script>alert('需要删除的文件路径不正确!');window.location.href='FolderImageList.asp';</script>"
			Response.End
		End If
		If Cstr(Split(Path,"/")(1)) <> Cstr(Fs_User.UserNumber) Then
			Response.Write "<script>alert('不能删除别人的文件!');window.location.href='FolderImageList.asp';</script>"
			Response.End
		End If 	
		Path = "/" & Path
		DelFileName = Request("FileName")
		If Instr(DelFileName,"/") <> 0 Or Instr(DelFileName,"\") <> 0 Or Left(DelFileName,1) = "." Then
			Response.Write "<script>alert('需要删除的文件路径不正确!');window.location.href='FolderImageList.asp';</script>"
			Response.End
		End If	
		if (DelFileName <> "") then
			Path = Server.MapPath(Path)
			if FsoObj.FileExists(Path & "\" & DelFileName) = true then FsoObj.DeleteFile Path & "\" & DelFileName
		end if
	elseif OType = "AddFolder" then
		Path = Request("Path")
		if Path <> "" then
			Dim FildNameStr
			If Right(Path,1) = "/" Then
				Path = Left(Path,Len(Path) - 1)
			End If
			FildNameStr = Split(Path,"/")(Ubound(Split(Path,"/")))
			IF ReplaceExpChar(FildNameStr) = False Then
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
	End If	
end if
AllowShowExtNameStr = "jpg,txt,gif,bmp,png"
CurrPath = Replace(Request("CurrPath"),"//","/")

If ReplaceExpChar(Replace(CurrPath,"/","")) = False Then
	Response.Write "<script>alert('您想干什么?\n您的IP已经被记录!\n我们将收集您部分计算机资料');window.location.href='javascript:history.back();';</script>"
	Response.End
End If	

if CurrPath = "" then
	CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
	ParentPath = ""
else
	ParentPath = Mid(CurrPath,1,InstrRev(CurrPath,"/")-1)
	if ParentPath = "" then
		CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
	end if
end if
dim CurrPath1
CurrPath1 = Server.MapPath(replace(Replace(CurrPath,"\\","\"),"/","\"))
if FsoObj.FolderExists(CurrPath1) = false then FsoObj.CreateFolder CurrPath1
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
.TempletItem {
	cursor: default;
}
.TempletSelectItem {
	background-color:highlight;
	cursor: default;
	color: white;
}
</style>
<link href="../../<% = G_USER_DIR %>/images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
var ObjPopupMenu=window.createPopup();
document.oncontextmenu=new Function("return ShowMouseRightMenu(window.event);");
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
</script>
<body topmargin="0" leftmargin="0" onClick="SelectFolder();"  class="hback">
<table width="99%" border="0" align="center" cellpadding="2" cellspacing="0"  class="hback">
  <%
if  lcase(Trim(CurrPath))<>  lcase(Trim(str_CurrPath)) then  
%>
  <tr title="上级目录<% = ParentPath %>" onClick="SelectUpFolder(this);" Path="<% = ParentPath %>" onDblClick="OpenParentFolder(this);"> 
    <td colspan="2"> <table width="62" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="21"><font color="#FFFFFF"><img src="../Images/folder.gif" width="20" height="16"></font></td>
          <td width="41">...</td>
        </tr>
      </table></td>
    <td width="35%"><div align="center"><font color="#FFFFFF">-</font></div></td>
    <td width="17%"><div align="center"><font color="#FFFFFF">-</font></div></td>
  </tr>
  <%
end if
for each FsoItem In SubFolderObj
%>
  <tr> 
    <td colspan="2"><table border="0" cellspacing="0" cellpadding="0">
        <tr title="双击鼠标进入此目录"> 
          <td valign="top"><img src="../Images/folder.gif" width="20" height="16"></td>
          <td valign="bottom"> <span class="TempletItem" Path="<% = FsoItem.name %>" onClick="ClearPicUrl()"; onDblClick="OpenFolder(this);"> 
            <% = FsoItem.name %>
            </span> </td>
        </tr>
      </table></td>
    <td><div align="left">文件夹</div></td>
    <td><div align="center"> 
       0 <!--<% = FsoItem.Size %>-->
      </div></td>
  </tr>
  <%
next

for each FsoItem In FileObj
	FileExtName = LCase(Mid(FsoItem.name,InstrRev(FsoItem.name,".")+1))
	if True then 'CheckFileShowTF(AllowShowExtNameStr,FileExtName) = 
%>
  <tr title="双击鼠标选择此文件"> 
    <td width="2%"> <span class="TempletItem" File="<% = FsoItem.name %>" onDblClick="SetFile(this);" onClick="SelectFile(this);"> 
      <img src="../Images/files.gif" width="16" height="16"> </span> </td>
    <td width="46%"><span class="TempletItem" File="<% = FsoItem.name %>" onDblClick="SetFile(this);" onClick="SelectFile(this);">
      <% = FsoItem.name %>
      </span></td>
    <td> <div align="left"> 
        <%if len(FsoItem.Type)>18 then:response.Write left(FsoItem.Type,18)&"...":else:response.Write FsoItem.Type:end if%>
      </div></td>
    <td><div align="center"> 
        <%
		if FsoItem.Size>1000 then
			Response.Write FormatNumber(FsoItem.Size/1024,1,-1) &"KB"
        Else
			Response.Write FsoItem.Size &"字节"
		End if
		%>
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
Set SubFolderObj = Nothing
Set FileObj = Nothing
%>
<script language="JavaScript">
var CurrPath='<% = CurrPath %>';
var G_VIRTUAL_ROOT_DIR='<% = G_VIRTUAL_ROOT_DIR %>';
var SelectedObj=null;
var ContentMenuArray=new Array();
DocumentReadyTF=false;
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	InitialClassListContentMenu();
	DocumentReadyTF=true;
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
	<%Response.write "ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction(""parent.AddFolderOperation();"",'新建目录','');"%>
	<%Response.write "ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction(""if (confirm('确定要删除吗？')==true) parent.DelFolderFile();"",'删除','disabled');"%>
}
function SelectFolder()
{
	Obj=event.srcElement,DisabledContentMenuStr='';
	if (SelectedObj!=null) SelectedObj.className='TempletItem';
	if ((Obj.Path!=null)||(Obj.File!=null))
	{
		Obj.className='TempletSelectItem';
		SelectedObj=Obj;
	}
	else SelectedObj=null;
	if (SelectedObj!=null)	DisabledContentMenuStr='';
	else DisabledContentMenuStr=',删除,';
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function SelectFile(Obj)
{
	//for (var i=0;i<document.all.length;i++)
	//{
		//if (document.all(i).className=='TempletSelectItem') document.all(i).className='TempletItem';
	//}
	Obj.className='TempletSelectItem';
	PreviewFile(Obj);
}
function OpenParentFolder(Obj)
{
	location.href='FolderImageList.asp?f_UserNumber=<%=session("FS_UserNumber")%>&CurrPath='+Obj.Path;
	SearchOptionExists(parent.document.all.FolderSelectList,Obj.Path);
}

function OpenFolder(Obj)
{
	var SubmitPath='';
	if (CurrPath=='/') SubmitPath=CurrPath+Obj.Path;
	else SubmitPath=CurrPath+'/'+Obj.Path;
	location.href='FolderImageList.asp?f_UserNumber=<%=session("FS_UserNumber")%>&CurrPath='+SubmitPath;
	AddFolderList(parent.document.all.FolderSelectList,SubmitPath,SubmitPath);
}

function SelectUpFolder(Obj)
{
	//for (var i=0;i<document.all.length;i++)
	//{
		//if (document.all(i).className=='TempletSelectItem') document.all(i).className='TempletItem';
	//}
	Obj.className='TempletSelectItem';
	parent.UserUrl.value='';
}
function PreviewFile(Obj)
{
	var Url='';
	var Path=escape();
	if (CurrPath=='/') Path=escape(CurrPath+Obj.File);
	else Path=escape(CurrPath+'/'+Obj.File);
	Url='PreviewImage.asp?FilePath='+Path;
	if (G_VIRTUAL_ROOT_DIR!='')
	Path=Path.slice(G_VIRTUAL_ROOT_DIR.length+1)
	parent.UserUrl.value=Path;
	parent.frames["PreviewArea"].location=Url.toLowerCase();
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
function SetFile(Obj)
{
	var PathInfo='',TempPath='';
	if (G_VIRTUAL_ROOT_DIR!='')
	{
		TempPath=CurrPath;
		PathInfo=TempPath.substr(TempPath.indexOf(G_VIRTUAL_ROOT_DIR)+G_VIRTUAL_ROOT_DIR.length);
	}
	else
	{
		PathInfo=CurrPath;
	}
	if (CurrPath=='/')	window.returnValue=PathInfo+Obj.File;
	else window.returnValue=PathInfo+'/'+Obj.File;
	window.close();
}
window.onunload=CheckReturnValue;
function CheckReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
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
			window.location.href='?f_UserNumber=<%=session("FS_UserNumber")%>&Type=AddFolder&Path='+CurrPath+'/'+ReturnValue+'&CurrPath='+CurrPath;
		}
	}	
}
function DelFolderFile()
{
	if (SelectedObj!=null)
	{
		if (SelectedObj.Path!=null) window.location.href='?f_UserNumber=<%=session("FS_UserNumber")%>&Type=DelFolder&Path='+CurrPath+'/'+SelectedObj.Path+'&CurrPath='+CurrPath;
		if (SelectedObj.File!=null) window.location.href='?f_UserNumber=<%=session("FS_UserNumber")%>&Type=DelFile&Path='+CurrPath+'&FileName='+SelectedObj.File+'&CurrPath='+CurrPath;
	}
	else alert('请选择要删除的目录');
}
/*function EditFolder()
{
	var ReturnValue='';
	if (SelectedObj!=null)
	{
		if (SelectedObj.Path!=null)
		{
			ReturnValue=prompt('修改的名称：',SelectedObj.Path);
			if ((ReturnValue!='') && (ReturnValue!=null)) window.location.href='?f_UserNumber=<%=session("FS_UserNumber")%>&Type=FolderReName&Path='+CurrPath+'&CurrPath='+CurrPath+'&OldPathName='+SelectedObj.Path+'&NewPathName='+ReturnValue;
		}
		if (SelectedObj.File!=null)
		{
			ReturnValue=prompt('修改的名称：',SelectedObj.File);
			if ((ReturnValue!='') && (ReturnValue!=null)) window.location.href='?f_UserNumber=<%=session("FS_UserNumber")%>&Type=FileReName&Path='+CurrPath+'&CurrPath='+CurrPath+'&OldFileName='+SelectedObj.File+'&NewFileName='+ReturnValue;
		}
	}
	else alert('请填写要更名的目录名称');
}*/
function ClearPicUrl()
{
parent.UserUrl.value='';
}
</script>
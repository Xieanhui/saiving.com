<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp"-->
<!--#include file="../../../FS_Inc/md5.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<%
Dim Conn
MF_Default_Conn
MF_Session_TF
Dim TypeSql,RsTypeObj,LableSql,RsLableObj
Set FsoObj = Server.CreateObject(G_FS_FSO)
Dim CurrPath,FsoObj,SubFolderObj,FolderObj,FileObj,i,FsoItem,OType
Dim ParentPath,FileExtName,AllowShowExtNameStr,str_CurrPath,sRootDir
Dim UpLoadFileNameStr
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
UpLoadFileNameStr = Conn.ExeCute("Select MF_UpFile_Type From FS_MF_Config Where ID > 0")(0)
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/" & G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if

OType = Request("Type")
if OType <> "" then
	Dim Path,PhysicalPath
	if OType = "DelFolder" then
		if not MF_Check_Pop_TF("MF027") then Err_Show
		Path = Trim(Request("Path")) 
		If Instr(Path,".") > 0 Then
			Response.Write "<script>alert('需要删除的目录路径不正确哦');window.location.href='FolderImageList.asp';</script>"
			Response.End
		End If
		If LCase(Cstr(Left(Right(Path,Len(Path) - 1),Len(G_UP_FILES_DIR)))) <> LCase(Cstr(G_UP_FILES_DIR)) Then
			Response.Write "<script>alert('需要删除的目录路径不正确哦');window.location.href='FolderImageList.asp';</script>"
			Response.End
		End If	
		if Path <> "" then
			Path = Server.MapPath(Path)
			if FsoObj.FolderExists(Path) = true then FsoObj.DeleteFolder Path
		end if
	elseif OType = "DelFile" then
		if not MF_Check_Pop_TF("MF027") then Err_Show
		Dim DelFileName
		Path = Trim(Request("Path"))
		If Instr(Path,".") > 0 Then
			Response.Write "<script>alert('需要删除的文件路径不正确哦');window.location.href='FolderImageList.asp';</script>"
			Response.End
		End If
		If LCase(Cstr(Left(Right(Path,Len(Path) - 1),Len(G_UP_FILES_DIR)))) <> LCase(Cstr(G_UP_FILES_DIR)) Then
			Response.Write "<script>alert('需要删除的文件路径不正确哦');window.location.href='FolderImageList.asp';</script>"
			Response.End
		End If 
		DelFileName = Request("FileName")
		If Instr(DelFileName,"/") <> 0 Or Instr(DelFileName,"\") <> 0 Or Left(DelFileName,1) = "." Then
			Response.Write "<script>alert('需要删除的文件路径不正确哦');window.location.href='FolderImageList.asp';</script>"
			Response.End
		End If
		if (DelFileName <> "") And (Path <> "") then
			Path = Server.MapPath(Path)
			if FsoObj.FileExists(Path & "\" & DelFileName) = true then FsoObj.DeleteFile Path & "\" & DelFileName
		end if
	elseif OType = "AddFolder" then
		if not MF_Check_Pop_TF("MF028") then Err_Show
		Path = Request("Path")
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
	elseif OType = "FileReName" then
		if not MF_Check_Pop_TF("MF026") then Err_Show
		Dim NewFileName,OldFileName,NewArr,Arr_i,NameTF
		Path = Request("Path")
		if Path <> "" then
			NewFileName = Request("NewFileName")
			OldFileName = Request("OldFileName")
			IF Ubound(Split(NewFileName,".")) <> 1 Then
				Response.Write "<script>alert('新的文件名不规范，请重设');window.location.href='FolderImageList.asp';</script>"
				Response.End
			End If
			NewArr = Split(UpLoadFileNameStr,",")
			NameTF = False
			For Arr_i = LBound(NewArr) To Ubound(NewArr)
				IF NewArr(Arr_i) = Split(NewFileName,".")(1) Then
					NameTF = True
					Exit For
				End If	
			NExt
			IF NameTF = False Then
				Response.Write "<script>alert('新的文件扩展名不被允许，请重设');window.location.href='FolderImageList.asp';</script>"
				Response.End
			End If	
			if (NewFileName <> "") And (OldFileName <> "") then
				PhysicalPath = Server.MapPath(Path) & "\" & OldFileName
				if FsoObj.FileExists(PhysicalPath) = True then
					PhysicalPath = Server.MapPath(Path) & "\" & NewFileName
					if FsoObj.FileExists(PhysicalPath) = False then
						Set FileObj = FsoObj.GetFile(Server.MapPath(Path) & "\" & OldFileName)
						FileObj.Name = NewFileName
						Set FileObj = Nothing
					end if
				end if
			end if
		end if
	elseif OType = "FolderReName" then
		if not MF_Check_Pop_TF("MF026") then Err_Show
		Dim NewPathName,OldPathName
		Path = Request("Path")
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
AllowShowExtNameStr = "jpg,txt,gif,bmp,png"
CurrPath = Replace(Request("CurrPath"),"//","/")
if G_VIRTUAL_ROOT_DIR <>"" then
	if Trim(CurrPath) = "/"  or  Trim(CurrPath) =G_UP_FILES_DIR &"/"  or lcase(Trim(CurrPath)) = lcase(G_UP_FILES_DIR &"/Foosun_Data") then
		Response.Write("非法参数")
		Response.end
	End if
Else
	if Trim(CurrPath) = "/"  or lcase(Trim(CurrPath)) = lcase("/Foosun_Data") then
		Response.Write("非法参数")
		Response.end
	End if
End if
if CurrPath = "" then
	CurrPath = str_CurrPath
	ParentPath = ""
else
	ParentPath = Mid(CurrPath,1,InstrRev(CurrPath,"/")-1)
	if ParentPath = "" then
		ParentPath =sRootDir &"/adminfiles/"&Temp_Admin_Name
	End If 
End If 

On Error Resume Next
Set FolderObj = FsoObj.GetFolder(Server.MapPath(CurrPath))
Set SubFolderObj = FolderObj.SubFolders
Set FileObj = FolderObj.Files
If Err Then
	Response.write "路径错误"
	Response.End
End If 
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
<link href="../../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../../../FS_Inc/prototype.js"></script>
<script language="JavaScript">
var CurrPath='<% = CurrPath %>';
var MFDomain='<%= Request.Cookies("FoosunMFCookies")("FoosunMFDomain") %>';
var G_VIRTUAL_ROOT_DIR='<% = G_VIRTUAL_ROOT_DIR %>';
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
function AutoSelectFile(Obj)
{	for (var i=0;i<document.all.length;i++)
	{
		if (document.all(i).className=='TempletSelectItem') document.all(i).className='TempletItem';
	}
	Obj.className='TempletSelectItem';
}
function AutoSetFile(Obj)
{
		var PathInfo='',TempPath='';
		//if (G_VIRTUAL_ROOT_DIR!='')
		//{
			//TempPath=CurrPath;
			//PathInfo=TempPath.substr(TempPath.indexOf(G_VIRTUAL_ROOT_DIR)+G_VIRTUAL_ROOT_DIR.length);
		//}
		//else
		//{
			PathInfo=CurrPath;
		//}
	//PathInfo="Http://"+MFDomain+PathInfo;
	if (CurrPath=='/')	window.returnValue=PathInfo+Obj.File;
	else window.returnValue=PathInfo+'/'+Obj.File;
	window.parent.document.all.UserUrl.value=window.returnValue;
}


function Order_AllFiles(Obj)
{
	var Order_ID = Obj.id;
	var Order_Text = 0;
	if (Order_ID.indexOf("_") != -1)
	{
		var TempArr = Order_ID.split("_");
		Order_Text = TempArr[1];
	}
	var Order_Type = Obj.value;
	if (Order_Type == "" || Order_Type == null)
	{
		Order_Type = 0;
	}
	else
	{
		if (Order_Type == "0")
		{
			Order_Type = "1";
		}
		else
		{
			Order_Type = "0";
		}
	}
	var Request = window.location.href;
	Request = Request.replace(/(\&Order_Text).*$/img,"");
	Request = Request.replace(/(\&Order_Type).*$/img,"");
	if (Request.indexOf("?") != -1)
		{
			window.location.href = Request + "&Order_Text=" + Order_Text + "&Order_Type=" + Order_Type;
		}
		else
		{
			window.location.href = Request + "?Order_Text=" + Order_Text + "&Order_Type=" + Order_Type;
		}
}


</script>
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
	<%if MF_Check_Pop_TF("MF028") then Response.Write "ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction(""parent.AddFolderOperation();"",'新建目录','');"%>
	<%if MF_Check_Pop_TF("MF027") then Response.write "ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction(""if (confirm('确定要删除吗？')==true) parent.DelFolderFile();"",'删除','disabled');"%>
	//此处是否允许管理员重命名文件名，如果开启重命名有风险，请小心使用。
	<%if MF_Check_Pop_TF("MF026") then Response.write "ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction(""parent.EditFolder();"",'重命名','disabled');"%>
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
	else DisabledContentMenuStr=',删除,重命名,';
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
	location.href='FolderImageList.asp?CurrPath='+Obj.Path;
	SearchOptionExists(parent.document.all.FolderSelectList,Obj.Path);
}

function OpenFolder(Obj)
{
	var SubmitPath='';
	if (CurrPath=='/') SubmitPath=CurrPath+Obj.Path;
	else SubmitPath=CurrPath+'/'+Obj.Path;
	location.href='FolderImageList.asp?CurrPath='+SubmitPath;
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
	//if (G_VIRTUAL_ROOT_DIR!='')
	//Path=Path.slice(G_VIRTUAL_ROOT_DIR.length+1)
	//parent.UserUrl.value="Http://"+MFDomain+Path;
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

		//var PathInfo='',TempPath='';
		//if (G_VIRTUAL_ROOT_DIR!='')
		//{
			//TempPath=CurrPath;
			//PathInfo=TempPath.substr(TempPath.indexOf(G_VIRTUAL_ROOT_DIR)+G_VIRTUAL_ROOT_DIR.length);
		//}
		//else
		//{
			PathInfo=CurrPath;
		//}
	//PathInfo="Http://"+MFDomain+PathInfo;
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
			window.location.href='?Type=AddFolder&Path='+CurrPath+'/'+ReturnValue+'&CurrPath='+CurrPath;
		}
	}
}
function DelFolderFile()
{
	if (SelectedObj!=null)
	{
		if (SelectedObj.Path!=null) window.location.href='?Type=DelFolder&Path='+CurrPath+'/'+SelectedObj.Path+'&CurrPath='+CurrPath;
		if (SelectedObj.File!=null) window.location.href='?Type=DelFile&Path='+CurrPath+'&FileName='+SelectedObj.File+'&CurrPath='+CurrPath;
	}
	else alert('请选择要删除的目录');
}
function EditFolder()
{
	var ReturnValue='';
	if (SelectedObj!=null)
	{
		if (SelectedObj.Path!=null)
		{
			ReturnValue=prompt('修改的名称：',SelectedObj.Path);
			var patrn =/([^a-zA-Z0-9])/; 
			if (patrn.exec(ReturnValue))
			{
				alert('修改目录名不规范，请重设');
				return false;
			}
			else
			{
				window.location.href='?Type=FolderReName&Path='+CurrPath+'&CurrPath='+CurrPath+'&OldPathName='+SelectedObj.Path+'&NewPathName='+ReturnValue;
			}
		}
		if (SelectedObj.File!=null)
		{
			ReturnValue=prompt('修改的名称：',SelectedObj.File);
			if ((ReturnValue!='') && (ReturnValue!=null)) 
			{
				var oldArr = SelectedObj.File.split('.');
				var NewArr = ReturnValue.split('.');
				var upload_Str = '<% = UpLoadFileNameStr %>';
				var Up_Arr = upload_Str.split(',');
				if (oldArr.length != NewArr.length)
				{
					alert('新的文件名不规范，请重设');
					return false;
				}
				window.location.href='?Type=FileReName&Path='+CurrPath+'&CurrPath='+CurrPath+'&OldFileName='+SelectedObj.File+'&NewFileName='+ReturnValue;
			}	
		}
	}
	else alert('请填写要更名的目录名称');
}
function ClearPicUrl()
{
parent.UserUrl.value='';
}
</script>
<body topmargin="0" leftmargin="0" onLoad="AddFolderList(parent.document.all.FolderSelectList,'<% = CurrPath %>','<% = CurrPath %>');" onClick="SelectFolder();">
<%
'----2007-07-19 图片存入数组，排序显示
Dim All_Files_Arr(),str_Files_Name,str_Files_Size,str_Files_Date,str_Files_Type,Files_obj
Dim int_AllFilesNum,List_i,List_j,Ruler_Txt,Ruler_ID,Temp_Obj,Arr_List_i
Dim Order_Text,Order_Type,Order_Kind,Order_Mark
'排序字段：0文件名，1文件类型，2文件大小，3文件日期
Order_Text = Trim(Request.QueryString("Order_Text"))
If Order_Text = "" Or IsNull(Order_Text) Or Not IsNumeric(Order_Text) Then
	Order_Text = 3
Else
	Order_Text = Cint(Order_Text)
End If
'排序方式，0降序，1升序
Order_Type = Trim(Request.QueryString("Order_Type"))
If Order_Type = "" Or IsNull(Order_Type) Or Not IsNumeric(Order_Type) Then
	Order_Type = 0
Else
	Order_Type = Cint(Order_Type)
End If
'---
if Order_Text = 0 Or Order_Text = 1 Then
	If Order_Type = 0 Then
		Order_Kind = 1
	Else
		Order_Kind = 2
	End if
Else
	If Order_Type = 0 Then
		Order_Kind = 3
	Else
		Order_Kind = 4
	End if
End if	
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
<tr>
    <td width="37%" align="left" valign="middle" class="xingmu" title="点击以文件名排序"><span value="<% = Order_Type %>" id="Order_0" style="cursor:pointer;" onClick="Javascript:Order_AllFiles(this);">文 件 名</span></td>
    <td width="18%" align="left" valign="middle" class="xingmu" title="点击以文件类型排序"><span value="<% = Order_Type %>" id="Order_1" style="cursor:pointer;" onClick="Javascript:Order_AllFiles(this);">文件类型</span></td>
    <td width="15%" align="left" valign="middle" class="xingmu" title="点击以文件大小排序"><span value="<% = Order_Type %>" id="Order_2" style="cursor:pointer;" onClick="Javascript:Order_AllFiles(this);">文件大小</span></td>
    <td width="30%" align="left" valign="middle" class="xingmu" title="点击以上传日期排序"><span value="<% = Order_Type %>" id="Order_3" style="cursor:pointer;" onClick="Javascript:Order_AllFiles(this);">上传日期</span></td>
  </tr> 
  <%
if  lcase(Trim(CurrPath))<>  lcase(Trim(str_CurrPath)) then  
%>
  <tr title="上级目录<% = ParentPath %>" onClick="SelectUpFolder(this);" Path="<% = ParentPath %>" onDblClick="OpenParentFolder(this);"> 
    <td width="37%" class="hback"> <table width="62" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="21"><font color="#FFFFFF"><img src="../../Images/Folder/folder.gif" width="20" height="16"></font></td>
          <td width="41">...</td>
        </tr>
    </table></td>
    <td width="18%" class="hback"><div align="center"><font color="#FFFFFF">-</font></div></td>
    <td width="15%" class="hback"><div align="center"><font color="#FFFFFF">-</font></div></td>
    <td width="30%" class="hback">&nbsp;</td>
  </tr>
  <%
end if
for each FsoItem In SubFolderObj
%>
  <tr> 
    <td width="37%"  class="hback"><table border="0" cellspacing="0" cellpadding="0">
        <tr title="双击鼠标进入此目录"> 
          <td valign="top"><img src="../../Images/Folder/folder.gif" width="20" height="16"></td>
          <td> <span class="TempletItem" Path="<% = FsoItem.name %>" onClick="ClearPicUrl()"; onDblClick="OpenFolder(this);"> 
            <% = FsoItem.name %>
            </span> </td>
        </tr>
    </table></td>
    <td width="18%" class="hback"><div align="left">文件夹</div></td>
    <td width="15%" align="left" class="hback"><div align="left"> 
        0
		<!--<% = FsoItem.Size %>-->
    </div></td>
    <td width="30%" align="left" class="hback">&nbsp;</td>
  </tr>
  <%
next
%>
 
<%
'---存入数组
int_AllFilesNum = -1
For Each Files_obj In FileObj
	int_AllFilesNum = int_AllFilesNum + 1
	str_Files_Name = Files_obj.Name
	str_Files_Size = Files_obj.Size
	str_Files_Date = Files_obj.DateCreated
	str_Files_Type = Files_obj.Type
	ReDim Preserve All_Files_Arr(int_AllFilesNum)
	All_Files_Arr(int_AllFilesNum) = Array(str_Files_Name,str_Files_Type,str_Files_Size,str_Files_Date)
Next

Set FsoObj = Nothing
Set SubFolderObj = Nothing
Set FileObj = Nothing

'---按照排序重排数组
For List_i = int_AllFilesNum To 0 Step - 1
	Ruler_Txt = All_Files_Arr(0)(Order_Text)
	Ruler_ID = 0
	For List_j = 1 To List_i
		Select Case Order_Kind
			Case 1
				Order_Mark = (StrComp(All_Files_Arr(List_j)(Order_Text),Ruler_Txt,vbTextCompare) < 0)
			Case 2
				Order_Mark = (StrComp(All_Files_Arr(List_j)(Order_Text),Ruler_Txt,vbTextCompare) > 0)
			Case 3
				Order_Mark = (All_Files_Arr(List_j)(Order_Text) < Ruler_Txt)
			Case 4
				Order_Mark = (All_Files_Arr(List_j)(Order_Text) > Ruler_Txt)			
		End Select
		If Order_Mark = True Then
			Ruler_Txt = All_Files_Arr(List_j)(Order_Text)
			Ruler_ID = List_j
		End If
	Next
	If Ruler_ID <> List_i Then
		Temp_Obj = All_Files_Arr(Ruler_ID)
		All_Files_Arr(Ruler_ID) = All_Files_Arr(List_i)
		All_Files_Arr(List_i) = Temp_Obj
	End If	
Next

For Arr_List_i = LBound(All_Files_Arr) To UBound(All_Files_Arr)
	Dim Arr_Files_Type
	Arr_Files_Type = All_Files_Arr(Arr_List_i)(1)
	Arr_Files_Type = Replace(Arr_Files_Type," ","")
%>
  <tr title="双击鼠标选择此文件"> 
    <td width="37%"  class="hback">
	<%
	if Session("upfiles")=All_Files_Arr(Arr_List_i)(0) then %>
	<span  id="span_<% = All_Files_Arr(Arr_List_i)(0) %>" File="<% = All_Files_Arr(Arr_List_i)(0) %>" onDblClick="SetFile(this);" onClick="SelectFile(this);"> 
	<script language="javascript">
		AutoSelectFile($("span_<% = All_Files_Arr(Arr_List_i)(0) %>"));
		AutoSetFile($("span_<% = All_Files_Arr(Arr_List_i)(0) %>"));
		SelectedObj = $("span_<% = All_Files_Arr(Arr_List_i)(0) %>");
	</script>
	<%Else%>
	<span  id="<% = All_Files_Arr(Arr_List_i)(0) %>" File="<% = All_Files_Arr(Arr_List_i)(0) %>" onDblClick="SetFile(this);" onClick="SelectFile(this);"> 
	<%End if%>
      <img src="../../Images/FileIcon/doc.gif" width="16" height="16"> 
      <% = All_Files_Arr(Arr_List_i)(0) %>
      </span> </td>
    <td width="18%" class="hback"> <div align="left"> 
        <%if len(Arr_Files_Type)>18 then:response.Write left(Arr_Files_Type,18)&"...":else:response.Write Arr_Files_Type:end if%>
      </div></td>
    <td width="15%" class="hback"><div align="left"> 
        <%
		if All_Files_Arr(Arr_List_i)(2)>1000 then
			Response.Write FormatNumber(All_Files_Arr(Arr_List_i)(2)/1024,1,-1) &"KB"
        Else
			Response.Write All_Files_Arr(Arr_List_i)(2) &"字节"
		End if
		%>
      </div></td>
    <td width="30%" class="hback"><% = All_Files_Arr(Arr_List_i)(3) %></td>
  </tr>
  <%
  	'end if
next
%>
</table>
</body>
</html>








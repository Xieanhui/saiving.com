<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Cls_Cache.asp" -->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("MF_Templet") then Err_Show
Dim p_ParentPath,p_ShowPath,str_ShowPath,strShowErr
p_ShowPath=replace(replace(Request("ShowPath"),"'",""),"""","")
If p_ShowPath="" Then
	p_ShowPath=Add_Root_Dir("/") & G_TEMPLETS_DIR
	str_ShowPath = p_ShowPath
Else
	str_ShowPath =p_ShowPath
End If
p_ParentPath = Mid(p_ShowPath,1,InstrRev(p_ShowPath,"/")-1)
Dim p_FSO,p_FolderObj,p_SubFolderObj,p_FileObj,p_FileIconDic
Dim p_FileItem
Set p_FSO = Server.CreateObject(G_FS_FSO)

Set p_FolderObj = p_FSO.GetFolder(Server.MapPath(p_ShowPath))
Set p_SubFolderObj = p_FolderObj.SubFolders
Set p_FileObj = p_FolderObj.Files
Set p_FileIconDic = CreateObject(G_FS_DICT)
p_FileIconDic.Add "txt","Images/FileIcon/txt.gif"
p_FileIconDic.Add "gif","Images/FileIcon/gif.gif"
p_FileIconDic.Add "exe","Images/FileIcon/exe.gif"
p_FileIconDic.Add "asp","Images/FileIcon/asp.gif"
p_FileIconDic.Add "html","Images/FileIcon/html.gif"
p_FileIconDic.Add "htm","Images/FileIcon/html.gif"
p_FileIconDic.Add "jpg","Images/FileIcon/jpg.gif"
p_FileIconDic.Add "jpeg","Images/FileIcon/jpg.gif"
p_FileIconDic.Add "pl","Images/FileIcon/perl.gif"
p_FileIconDic.Add "perl","Images/FileIcon/perl.gif"
p_FileIconDic.Add "zip","Images/FileIcon/zip.gif"
p_FileIconDic.Add "rar","Images/FileIcon/zip.gif"
p_FileIconDic.Add "gz","Images/FileIcon/zip.gif"
p_FileIconDic.Add "doc","Images/FileIcon/doc.gif"
p_FileIconDic.Add "xml","Images/FileIcon/xml.gif"
p_FileIconDic.Add "xsl","Images/FileIcon/xml.gif"
p_FileIconDic.Add "dtd","Images/FileIcon/xml.gif"
p_FileIconDic.Add "vbs","Images/FileIcon/vbs.gif"
p_FileIconDic.Add "js","Images/FileIcon/vbs.gif"
p_FileIconDic.Add "wsh","Images/FileIcon/vbs.gif"
p_FileIconDic.Add "sql","Images/FileIcon/script.gif"
p_FileIconDic.Add "bat","Images/FileIcon/script.gif"
p_FileIconDic.Add "tcl","Images/FileIcon/script.gif"
p_FileIconDic.Add "eml","Images/FileIcon/mail.gif"
p_FileIconDic.Add "swf","Images/FileIcon/flash.gif"
if Request.QueryString("Type") = "FileReName" then
		if not MF_Check_Pop_TF("MF002") then Err_Show
		Dim NewFileName,OldFileName,Path,PhysicalPath,FileObj
		Path = replace(replace(Request("Path"),"'",""),"""","")
		if Path <> "" then
			NewFileName = Request("NewFileName")
			OldFileName = Request("OldFileName")
			if (NewFileName <> "") And (OldFileName <> "") then
				PhysicalPath = Server.MapPath(Path) & "\" & OldFileName
				if p_FSO.FileExists(PhysicalPath) = True then
					PhysicalPath = Server.MapPath(Path) & "\" & NewFileName
					if p_FSO.FileExists(PhysicalPath) = False then
						Set FileObj = p_FSO.GetFile(Server.MapPath(Path) & "\" & OldFileName)
						FileObj.Name = NewFileName
						Set FileObj = Nothing
					end if
				end if
			end if
		end if
		Call MF_Insert_oper_Log("模板管理","改名了模板文件，名称："& OldFileName &",路径："& Path &"",now,session("admin_name"),"MF")
		strShowErr = "<li>修改文件成功</li>"
		Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
End if

if Request.QueryString("Type") = "FolderReName" then
		if not MF_Check_Pop_TF("MF002") then Err_Show
		Dim NewPathName,OldPathName
		Path = replace(replace(Request("Path"),"'",""),"""","")
		if Path <> "" then
			NewPathName = replace(replace(Request("NewFileName"),"'",""),"""","")
			OldPathName = replace(replace(Request("OldFileName"),"'",""),"""","")
			if (NewPathName <> "") And (OldPathName <> "") then
				PhysicalPath = Server.MapPath(Path) & "\" & OldPathName
				if p_FSO.FolderExists(PhysicalPath) = True then
					PhysicalPath = Server.MapPath(Path) & "\" & NewPathName
					if p_FSO.FolderExists(PhysicalPath) = False then
						Set FileObj = p_FSO.GetFolder(Server.MapPath(Path) & "\" & OldPathName)
						FileObj.Name = NewPathName
						Set FileObj = Nothing
					end if
				end if
			end if
		end if
		Call MF_Insert_oper_Log("模板管理","改名了模板目录，目录："& OldPathName &",路径："& Path &"",now,session("admin_name"),"MF")
		strShowErr = "<li>修改目录成功</li>"
		Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
End if

if Request.QueryString("Type") = "AddFolder" then
		if not MF_Check_Pop_TF("MF004") then Err_Show
		Path = replace(replace(Request("Path"),"'",""),"""","")
		if Path <> "" then
			Path = Server.MapPath(Replace(Path,"//","/"))
			if p_FSO.FolderExists(Path) = True then
				strShowErr = "<li>目录已经存在</li>"
				Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			else
				p_FSO.CreateFolder Path
			end if
		end if
		strShowErr = "<li>创建目录成功</li>"
		Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
end if

if Request.QueryString("Action") = "Delfolder" then
	if not MF_Check_Pop_TF("MF003") then Err_Show
	Path = replace(replace(Request.QueryString("Dir"),"'",""),"""","") &"/"&replace(replace(Request.QueryString("File"),"'",""),"""","")
	if Path <> "" then
		Path = Server.MapPath(Path)
		if p_FSO.FolderExists(Path) = true then p_FSO.DeleteFolder Path
	end if
	Call MF_Insert_oper_Log("模板管理","删除了目录，目录："& replace(replace(Request.QueryString("File"),"'",""),"""","") &",路径："& replace(replace(Request.QueryString("Dir"),"'",""),"""","") &"",now,session("admin_name"),"MF")
	strShowErr = "<li>删除目录成功！</li>"
	Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
if Request.QueryString("Action") = "Delfile" then
	if not MF_Check_Pop_TF("MF003") then Err_Show
	Dim DelFileName
	Path = replace(replace(Request.QueryString("Dir"),"'",""),"""","")
	DelFileName = replace(replace(Request.QueryString("File"),"'",""),"""","")
	if (DelFileName <> "") And (Path <> "") then
		Path = Server.MapPath(Path)
		if p_FSO.FileExists(Path & "\" & DelFileName) = true then p_FSO.DeleteFile Path & "\" & DelFileName
	end if
	Call MF_Insert_oper_Log("模板管理","删除了文件，文件："& replace(replace(Request.QueryString("File"),"'",""),"""","") &",路径："& replace(replace(Request.QueryString("Dir"),"'",""),"""","") &"",now,session("admin_name"),"MF")
	strShowErr = "<li>删除文件成功！</li>"
	Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr> 
    <td class="xingmu">模板管理</td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table" >
  <form name="form1" method="post" action="">
    <tr class="hback"> 
      <td width="204" height="20"  class="xingmu"><div align="left">名称</div></td>
      <td height="20" class="xingmu"><div align="center">类型</div></td>
      <td height="20" class="xingmu"><div align="center">大小(kb)</div></td>
      <td height="20" class="xingmu"><div align="center">最后修改时间</div></td>
      <td class="xingmu"><div align="center">操作</div></td>
    </tr>
    <% If p_ParentPath <> "/" & G_VIRTUAL_ROOT_DIR And p_ParentPath <> "/" Then %>
    <tr class="hback"> 
      <td  style="cursor:hand"><table border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td><img src="Images/arrow.gif" width="18" height="18"></td>
            <td><span class="TempletItem" title="上级目录<% = p_ParentPath %>" onDblClick="OpenParentFolder(this);" Path="<% = p_ParentPath %>">上级目录</span></td>
          </tr>
        </table></td>
      <td height="20" width="98">&nbsp;</td>
      <td height="20" width="87">&nbsp;</td>
      <td height="20" width="143">&nbsp;</td>
      <td width="218">&nbsp;</td>
    </tr>
    <%
	End If
	For Each p_FileItem In p_SubFolderObj
		%>
    <tr class="hback"> 
      <td height=""> <table border="0" cellspacing="0" cellpadding="0">
          <tr title="双击鼠标进入此目录" style="cursor:hand"> 
            <td><img src="Images/Folder/folder.gif" width="20" height="16"></td>
            <td ><span class="TempletItem" Path="<% = p_FileItem.name %>" onDblClick="OpenFolder(this);"> 
              <% = p_FileItem.name %>
              </span> </td>
          </tr>
        </table></td>
      <td> <div align="center">文件夹</div></td>
      <td> <div align="center"> 
          <%
		  if p_FileItem.Size<100 then
			 Response.Write p_FileItem.Size &"Byte"
		  Else
			 Response.Write FormatNumber(p_FileItem.Size/1024,1,-1) &"KB"
		  End if
		  %>
        </div></td>
      <td> <div align="center"> 
          <% = p_FileItem.DateLastModified %>
        </div></td>
      <td><div align="left"><a href="#" onClick="EditFolder('<% = p_FileItem.name %>','<%=str_ShowPath%>')">改名</a>｜<a href="Templets_List.asp?Action=Delfolder&File=<% = p_FileItem.name %>&Dir=<%=str_ShowPath%>" onClick="{if(confirm('确定删除此目录吗？')){return true;}return false;}">删除</a> 
        </div></td>
    </tr>
    <%
		Next
		For Each p_FileItem In p_FileObj
			Dim p_FileIcon,p_FileExtName
			p_FileExtName = Mid(CStr(p_FileItem.Name),Instr(CStr(p_FileItem.Name),".")+1)
			If lcase(p_FileExtName)="html" Or lcase(p_FileExtName)="htm" Or lcase(p_FileExtName)="css" Then 
				p_FileIcon = p_FileIconDic.Item(LCase(p_FileExtName))
				If p_FileIcon = "" Then
					p_FileIcon = "Images/FileIcon/unknown.gif"
				End If
		%>
    <tr class="hback"> 
      <td  style="cursor:hand"> <table border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><img src="<% = p_FileIcon %>"></td>
            <td><span File="<% = p_FileItem.Name %>"> 
              <% = p_FileItem.Name %>
              </span></td>
          </tr>
        </table></td>
      <td width="98"> <div align="center"> 
          <% = p_FileItem.Type %>
        </div></td>
      <td width="87"> <div align="center"> 
          <%
		  if p_FileItem.Size<100 then
			 Response.Write p_FileItem.Size &"字节"
		  Else
			 Response.Write FormatNumber(p_FileItem.Size/1024,1,-1) &"KB"
		  End if
		  %>
        </div></td>
      <td width="143"> <div align="center"> 
          <% = p_FileItem.DateLastModified %>
        </div></td>
      <td width="218"><div align="left"><a href="Templets_Edit_.asp?File=<% = p_FileItem.name %>&Dir=<%=str_ShowPath%>">在线编辑</a>｜<a href="Templets_Edit_text.asp?File=<% = p_FileItem.name %>&Dir=<%=str_ShowPath%>">文本编辑</a>｜<a href="#" onClick="Editfile('<% = replace(p_FileItem.name,"'","\'") %>','<%=str_ShowPath%>')">改名</a>｜<a href="<%=str_ShowPath%>/<% = p_FileItem.name %>" target="_blank">预览</a>｜<a href="Templets_List.asp?Action=Delfile&File=<% = p_FileItem.name %>&Dir=<%=str_ShowPath%>" onClick="{if(confirm('确定删除此文件吗？')){return true;}return false;}">删除</a> 
        </div></td>
    </tr>
    <%
	  else
	  end if
	next
	%>
    <tr class="hback"> 
      <td colspan="5" height="32"><div align="right"><span  class="tx">小提示：双击目录进入下一级目录</span>
          <input type="button" name="Submit2" value="创建目录" onClick="AddFolder();">
          <input type="button" name="Submit" value="导入文件" onClick="ImportTempletFile();" <%if not MF_Check_Pop_TF("MF005") then response.Write " disabled"%>>
        </div></td>
    </tr>
  </form>
</table>

</body>
</html>
<%
'Conn.Close
'Set Conn = Nothing
Set p_FSO = Nothing
Set p_FolderObj = Nothing
Set p_FileObj = Nothing
Set p_SubFolderObj = Nothing
Set p_FileIconDic = Nothing
%>
<script>
function AddFolder()
{
	var CurrPath='<% = p_ShowPath %>';
	var ReturnValue=prompt('新建目录名：','');
	if ((ReturnValue!='') && (ReturnValue!=null))
	window.location.href='?Path=<%=p_ShowPath%>/'+ReturnValue+'&Type=AddFolder&CurrPath=<%=p_ShowPath%>';
}
function ImportTempletFile()
{
	var CurrPath='<% = p_ShowPath %>';
	OpenWindow('CommPages/SelectManageDir/Frame.asp?FileName=UpFileForm.asp&PageTitle=上传文件&Path='+CurrPath,350,170,window);
}

function OpenFolder(Obj)
{    
	var SubmitPath='';
	var CurrPath='<% = p_ShowPath %>';
	if (CurrPath=='/') SubmitPath=CurrPath+Obj.Path;
	else SubmitPath=CurrPath+'/'+Obj.Path;
	location.href='Templets_List.asp?ShowPath='+SubmitPath;
}
function OpenParentFolder(Obj)
{
	location.href='Templets_List.asp?ShowPath='+Obj.Path;
}
function Editfile(filename,path)
{
		var ReturnValue='';
		ReturnValue=prompt('修改的名称：',filename.replace(/'|"/g,''));
		if ((ReturnValue!='') && (ReturnValue!=null)) window.location.href='?Type=FileReName&Path='+path+'&OldFileName='+filename+'&NewFileName='+ReturnValue;
		else if(ReturnValue!=null){alert('请填写要更名的名称');}
}
function EditFolder(filename,path)   
{
	var ReturnValue='';
	ReturnValue=prompt('修改的名称：',filename.replace(/'|"/g,''));
	if ((ReturnValue!='') && (ReturnValue!=null)) window.location.href='?Type=FolderReName&Path='+path+'&OldFileName='+filename+'&NewFileName='+ReturnValue;
		else if(ReturnValue!=null){alert('请填写要更名的名称');}
}
</script>







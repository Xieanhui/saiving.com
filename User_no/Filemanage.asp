<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim p_ParentPath,UserPath,str_ShowPath,PathName
Dim p_FSO,p_FolderObj,p_SubFolderObj,p_FileObj,p_FileIconDic
Dim p_FileItem,UserFileSpace,FildNameStr

User_GetParm
p_ParentPath=""
UserPath = Add_Root_Dir("/") & G_USERFILES_DIR & "/" &Fs_User.UserNumber

PathName = Request("ShowPath")

if ReplaceExpChar(PathName) = False Then
	Response.Write "<script>alert('目录名不规范');window.location.href='FileManage.asp';</script>"
	Response.End
End If

If InstrRev(PathName,"/") >0 Then
	p_ParentPath = Mid(PathName,1,InstrRev(PathName,"/")-1)
End If
If PathName="" Then
	str_ShowPath = UserPath
Else
	str_ShowPath = UserPath & "/" & PathName
End If

Set p_FSO = Server.CreateObject(G_FS_FSO)
If p_FSO.FolderExists(Server.MapPath(UserPath))=false then
	p_FSO.CreateFolder Server.MapPath(UserPath)
End If
if p_FSO.FolderExists(Server.MapPath(str_ShowPath))=false then
	Response.Write "<script>alert('目录不存在');window.location.href='FileManage.asp';</script>"
	Response.End()
end if

'获得空间大小
set UserFileSpace=p_FSO.GetFolder(Server.MapPath(UserPath))
Set p_FolderObj = p_FSO.GetFolder(Server.MapPath(str_ShowPath))
Set p_SubFolderObj = p_FolderObj.SubFolders
Set p_FileObj = p_FolderObj.Files
Set p_FileIconDic = CreateObject(G_FS_DICT)
p_FileIconDic.Add "txt","../Sys_Images/FileIcon/txt.gif"
p_FileIconDic.Add "gif","../Sys_Images/FileIcon/gif.gif"
p_FileIconDic.Add "exe","../Sys_Images/FileIcon/exe.gif"
p_FileIconDic.Add "asp","../Sys_Images/FileIcon/asp.gif"
p_FileIconDic.Add "html","../Sys_Images/FileIcon/html.gif"
p_FileIconDic.Add "htm","../Sys_Images/FileIcon/html.gif"
p_FileIconDic.Add "jpg","../Sys_Images/FileIcon/jpg.gif"
p_FileIconDic.Add "jpeg","../Sys_Images/FileIcon/jpg.gif"
p_FileIconDic.Add "pl","../Sys_Images/FileIcon/perl.gif"
p_FileIconDic.Add "perl","../Sys_Images/FileIcon/perl.gif"
p_FileIconDic.Add "zip","../Sys_Images/FileIcon/zip.gif"
p_FileIconDic.Add "rar","../Sys_Images/FileIcon/zip.gif"
p_FileIconDic.Add "gz","../Sys_Images/FileIcon/zip.gif"
p_FileIconDic.Add "doc","../Sys_Images/FileIcon/doc.gif"
p_FileIconDic.Add "xml","../Sys_Images/FileIcon/xml.gif"
p_FileIconDic.Add "xsl","../Sys_Images/FileIcon/xml.gif"
p_FileIconDic.Add "dtd","../Sys_Images/FileIcon/xml.gif"
p_FileIconDic.Add "vbs","../Sys_Images/FileIcon/vbs.gif"
p_FileIconDic.Add "js","../Sys_Images/FileIcon/vbs.gif"
p_FileIconDic.Add "wsh","../Sys_Images/FileIcon/vbs.gif"
p_FileIconDic.Add "sql","../Sys_Images/FileIcon/script.gif"
p_FileIconDic.Add "bat","../Sys_Images/FileIcon/script.gif"
p_FileIconDic.Add "tcl","../Sys_Images/FileIcon/script.gif"
p_FileIconDic.Add "eml","../Sys_Images/FileIcon/mail.gif"
p_FileIconDic.Add "swf","../Sys_Images/FileIcon/flash.gif"
if Request.QueryString("Type") = "AddFolder" then
		FildNameStr = Request("FolderName")
		If FildNameStr <> "" then
			If ReplaceExpChar(FildNameStr) = False Then
				strShowErr = "<li>新的目录名不规范，请重设.</li>"
			Else
				Path = UserPath & "/" & FildNameStr
				Path = Server.MapPath(Path)
				If p_FSO.FolderExists(Path) = True Then
					strShowErr = "<li>目录已经存在</li>"
				Else
					p_FSO.CreateFolder Path
					strShowErr = "<li>创建目录成功</li>"
				End If
			End If
		End If
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../FileManage.asp")
		Response.End()
end if
if Request.QueryString("Type") = "FolderReName" then
		Dim NewPathName,OldPathName,PhysicalPath,FileObj
		NewPathName = Request("NewFileName")
		OldPathName = Request("OldFileName")
		If ReplaceExpChar(NewPathName) = False Then
			Response.Write "<script>alert('目录名不规范，请重设');window.location.href='FileManage.asp';</script>"
			Response.End
		End If	
		if (NewPathName <> "") And (OldPathName <> "") then
			PhysicalPath = Server.MapPath(UserPath &"/"&OldPathName)
			if p_FSO.FolderExists(PhysicalPath) = True then
				PhysicalPath = Server.MapPath(UserPath &"/"&NewPathName)
				if p_FSO.FolderExists(PhysicalPath) = False then
					Set FileObj = p_FSO.GetFolder(Server.MapPath(UserPath &"/"&OldPathName))
					FileObj.Name = NewPathName
					Set FileObj = Nothing
				end if
			end if
		end if
		strShowErr = "<li>修改目录成功</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../FileManage.asp?ShowPath="& Path &"")
		Response.end
End if

if Request.QueryString("Action") = "Delfolder" then
	dim Path
	If Request.QueryString("Dir") = "" Then
		Path = UserPath &"/"& NoSqlHack(Request.QueryString("File"))
	Else
		Path = UserPath &"/"& NoSqlHack(Request.QueryString("Dir")) &"/"& NoSqlHack(Request.QueryString("File"))
	End If
	If Instr(Path,".") <> 0 Or Instr(Path,"/") = 0 Or Left(Path,1) <> "/" Then
		Response.Write "<script>alert('需要删除的目录路径不正确哦');window.location.href='FileManage.asp';</script>"
		Response.End
	End If
	If UBound(Split(Path,"/")) < 3 Then
		Response.Write "<script>alert('需要删除的目录路径不正确哦');window.location.href='FileManage.asp';</script>"
		Response.End
	End If
	If Cstr(Split(Path,"/")(2)) <> Cstr(Fs_User.UserNumber) Then
		Response.Write "<script>alert('不能删除别人的目录哦');window.location.href='FileManage.asp';</script>"
		Response.End
	End If 	
	Path = Server.MapPath(Path)
	if p_FSO.FolderExists(Path) = true then p_FSO.DeleteFolder Path
	strShowErr = "<li>删除目录成功！</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../FileManage.asp?ShowPath="& NoSqlHack(Request.QueryString("Dir")) &"")
	Response.end
End if
if Request.QueryString("Action") = "Delfile" then
	Dim DelFileName
	Path = replace(Request.QueryString("Dir"),"//","/")
	Path = UserPath &"/"& NoSqlHack(Request.QueryString("Dir"))
	If Instr(Path,".") <> 0 Or Instr(Path,"/") = 0 Or Left(Path,1) <> "/" Then
		Response.Write "<script>alert('需要删除的文件路径不正确!');window.location.href='FileManage.asp';</script>"
		Response.End
	End If
	Path = Right(Path,Len(Path) - 1)
	If UBound(Split(Path,"/")) < 1 Then
		Response.Write "<script>alert('需要删除的文件路径不正确!');window.location.href='FileManage.asp';</script>"
		Response.End
	End If
	If Cstr(Split(Path,"/")(1)) <> Cstr(Fs_User.UserNumber) Then
		Response.Write "<script>alert('不能删除别人的文件!');window.location.href='FileManage.asp';</script>"
		Response.End
	End If 	
	Path = "/" & Path
	DelFileName = Request.QueryString("File")
	If Instr(DelFileName,"/") <> 0 Or Instr(DelFileName,"\") <> 0 Or Left(DelFileName,1) = "." Then
		Response.Write "<script>alert('需要删除的文件路径不正确!');window.location.href='FileManage.asp';</script>"
		Response.End
	End If
	if (DelFileName <> "") then
		Path = Server.MapPath(Path)
		if p_FSO.FileExists(Path & "\" & DelFileName) = true then p_FSO.DeleteFile Path & "\" & DelFileName
	end if
	strShowErr = "<li>删除文件成功！</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../FileManage.asp?ShowPath="& NoSqlHack(Request.QueryString("Dir")) &"")
	Response.end
End if
if Request.Form("Action")="Saves" then
	if Request.Form("title")="" or Request.Form("PicSavePath")="" then
		strShowErr = "<li>图片标题和图片地址不能为空！</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	dim rs
	set rs= Server.CreateObject(G_FS_RS)
	rs.open "select * From FS_ME_Photo where PicSavePath='"& NoSqlHack(request.Form("PicSavePath")) &"' and UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
	if Not rs.eof then
		strShowErr = "<li>相册中已经有此图片！</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		rs.addnew
		rs("title")=NoSqlHack(request.Form("title"))
		rs("PicSavePath")=NoSqlHack(request.Form("PicSavePath"))
		rs("Content")=NoSqlHack(request.Form("Content"))
		rs("ClassID")=CintStr(request.Form("ClassID"))
		rs("Addtime")=now
		rs("UserNumber")=Fs_User.UserNumber
		if Request.Form("s")<>"" or not isnull(trim(Request.Form("s"))) then
			rs("PicSize") =CintStr(Request.Form("s"))
		else
			rs("PicSize") =0
		end if
		rs.update
	end if
	rs.close:set rs=nothing
	strShowErr = "<li>添加图片进相册成功！</li><li><a href=../FileManage.asp>返回文件管理</a>&nbsp;&nbsp;<a href=../PhotoManage.asp>返回相册</a></li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../FileManage.asp?ShowPath="& NoSqlHack(Request.QueryString("Dir")) &"")
	Response.end
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
<html xmlns="http://www.w3.org/1999/xhtml">
<title>欢迎用户<%=Fs_User.UserName%>来到<%=GetUserSystemTitle%>-文件管理</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head><body oncontextmenu="return false;">
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
	<tr>
		<td>
			<!--#include file="top.asp" -->
		</td>
	</tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
	<tr class="back">
		<td   colspan="2" class="xingmu" height="26">
			<!--#include file="Top_navi.asp" -->
		</td>
	</tr>
	<tr class="back">
		<td width="18%" valign="top" class="hback">
			<div align="left">
				<!--#include file="menu.asp" -->
			</div>
		</td>
		<td width="82%" valign="top" class="hback">
			<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
				<tr class="hback">
					<td class="hback"><strong>位置：</strong><a href="../">网站首页</a> &gt;&gt; <a href="main.asp">会员首页</a> &gt;&gt; <a href="FileManage.asp">文件管理</a> &gt;&gt;</td>
				</tr>
			</table>
			<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
				<tr class="hback">
					<td class="hback"><a href="FileManage.asp">所有文件</a>┆<a href="FileManage.asp?type=pic">图片文件</a>┆<a href="FileManage.asp?type=rar">压缩文件</a>┆<a href="FileManage.asp?type=doc">文档文件</a>
						<%if p_UpfileType="" or isnull(p_UpfileType) then%>
						┆<a href="#">上传图片没开放</a>
						<%else%>
						┆<a href="#" onClick="UpFileo();">上传图片</a>
						<%end if%>
						┆<a href="#" onClick="AddFolderOperation();">创建目录</a></td>
				</tr>
			</table>
			<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
				<tr>
					<td width="24%" class="hback">
						<div align="center">空间占用
							<%
		  dim tmp_use,tmp_space
		  getGroupIDinfo
		  tmp_use= (UserFileSpace.size)/1024
		  if UpfileSize="" or Isnull(UpfileSize) Then 'UpfileSize是从数据库里读出来的
		  	UpfileSize=2                          '默认是2M
		  else
		  	UpfileSize=Clng(UpfileSize)
		  end if
		  tmp_space = UpfileSize*1024
		  Response.Write FormatNumber(tmp_use/tmp_space,2,-1)*100
		  %>
							%，共<strong><%=Cint(UpfileSize/1024)%></strong>M</div>
					</td>
					<td width="76%" class="hback"><img src="images/space_pic_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.gif" width="<%=FormatNumber(tmp_use/tmp_space,2,-1)*100%>%" height="17"></td>
				</tr>
			</table>
			<%
	  if NoSqlHack(Request.QueryString("Action"))="joinphoto" then
	  %>
			<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
				<form name="form1" method="post" action="">
					<tr>
						<td colspan="2" class="xingmu">添加图片到相册</td>
					</tr>
					<tr class="hback">
						<td width="22%">
							<div align="right">标题</div>
						</td>
						<td width="78%">
							<input name="title" type="text" id="title" size="40" maxlength="50" value="<%=NoSqlHack(Request.QueryString("File"))%>">
						</td>
					</tr>
					<tr class="hback">
						<td>
							<div align="right">相册类别</div>
						</td>
						<td>
							<div id="selclass">
								<select name="Classid">
									<option value="0">选择相册分类</option>
									<%
				set rs=User_Conn.execute("select id,title From FS_ME_PhotoClass where UserNumber='"&session("FS_UserNumber")&"' order by id desc")
				do while not rs.eof
						response.Write"		<option value="""&rs("id")&""">"&rs("title")&"</option>"&chr(13)
					rs.movenext
				loop
				rs.close:set rs=nothing
				%>
								</select>
							</div>
						</td>
					</tr>
					<tr class="hback">
						<td>
							<div align="right">图片地址</div>
						</td>
						<td>
							<input name="PicSavePath" type="text" id="PicSavePath" size="40" maxlength="225" value="<%=NoSqlHack(Request.QueryString("Dir"))&"/"&NoSqlHack(Request.QueryString("File"))%>">
						</td>
					</tr>
					<tr class="hback">
						<td>
							<div align="right">描述</div>
						</td>
						<td>
							<textarea name="Content" rows="6" id="Content" style="width:80%"><%=Request.QueryString("File")%></textarea>
						</td>
					</tr>
					<tr class="hback">
						<td>&nbsp;</td>
						<td>
							<input type="submit" name="Submit" value="保存图片到相册中">
							<input type="reset" name="Submit2" value="重置">
							<input name="Action" type="hidden" id="Action" value="Saves">
							<input name="s" type="hidden" id="s" value="<%=request.QueryString("s")%>">
						</td>
					</tr>
				</form>
			</table>
			<%end if%>
			<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
				<tr>
					<td class="xingmu">ICO</td>
					<td class="xingmu">文件名</td>
					<td class="xingmu">大小</td>
					<td class="xingmu">最后修改时间</td>
					<td class="xingmu">
						<div align="center">类型</div>
					</td>
					<td class="xingmu">
						<div align="center">操作</div>
					</td>
				</tr>
				<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
					<td valign="top" class="hback">
						<div align="center"><img src="../sys_images/arrow.gif" width="24" height="22"></div>
					</td>
					<td  class="hback">
						<table border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td width="90"><span  style="cursor:hand" class="TempletItem" title="上级目录<% = p_ParentPath %>" onDblClick="OpenParentFolder(this);" Path="<% = p_ParentPath %>">返回上级目录</span></td>
							</tr>
						</table>
					</td>
					<td class="hback">&nbsp;</td>
					<td class="hback">&nbsp;</td>
					<td class="hback">&nbsp;</td>
					<td class="hback">&nbsp;</td>
				</tr>
				<%	For Each p_FileItem In p_SubFolderObj%>
				<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
					<td valign="top" class="hback">
						<div align="center"><img src="../sys_images/folder.gif" width="20" height="16"></div>
					</td>
					<td valign="bottom" class="hback"><span class="TempletItem" Path="<% = p_FileItem.name %>" onDblClick="OpenFolder(this);" style="cursor:hand;"> <strong>
						<% = p_FileItem.name %>
						</strong></span></td>
					<td class="hback">
						<%
		  if p_FileItem.Size<100 then
			 Response.Write p_FileItem.Size &"Byte"
		  Else
			 Response.Write FormatNumber(p_FileItem.Size/1024,1,-1) &"KB"
		  End if
		  %>
					</td>
					<td class="hback">
						<% = p_FileItem.DateLastModified %>
					</td>
					<td class="hback">
						<div align="center">文件夹</div>
					</td>
					<td class="hback">
						<div align="left"><a href="改名" onClick="EditFolder('<% = p_FileItem.name %>');return false;">改名</a>┆<a href="FileManage.asp?Action=Delfolder&File=<% = p_FileItem.name %>" onClick="{if(confirm('确定删除此目录吗？')){return true;}return false;}">删除</a> </div>
					</td>
				</tr>
				<%
		Next
		For Each p_FileItem In p_FileObj
			Dim p_FileIcon,p_FileExtName,o_types
			p_FileExtName = Mid(CStr(p_FileItem.Name),Instr(CStr(p_FileItem.Name),".")+1)
			select case Request.QueryString("type")
				case "pic"
					o_types =  lcase(p_FileExtName)="jpg" Or lcase(p_FileExtName)="gif" Or lcase(p_FileExtName)="png"  Or lcase(p_FileExtName)="bmp" Or lcase(p_FileExtName)="jpeg"
				case "rar"
					o_types =  lcase(p_FileExtName)="rar" Or lcase(p_FileExtName)="zip"
				case "doc"
					o_types =  lcase(p_FileExtName)="doc"
				case else
					o_types =  lcase(p_FileExtName)="jpg" Or lcase(p_FileExtName)="gif" Or lcase(p_FileExtName)="png"  Or lcase(p_FileExtName)="bmp" Or lcase(p_FileExtName)="jpeg" or lcase(p_FileExtName)="rar" Or lcase(p_FileExtName)="zip" or lcase(p_FileExtName)="doc"
			end select
			If o_types Then 
				p_FileIcon = p_FileIconDic.Item(LCase(p_FileExtName))
				If p_FileIcon = "" Then
					p_FileIcon = "Images/FileIcon/unknown.gif"
				End If
		%>
				<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
					<td width="6%" class="hback">
						<div align="center"><img src="<% = p_FileIcon %>"></div>
					</td>
					<td width="36%" class="hback">
						<% = p_FileItem.Name %>
					</td>
					<td width="12%" class="hback"> <font style="font-family:Verdana;font-size:7pt">
						<%
		  if p_FileItem.Size<100 then
			 Response.Write p_FileItem.Size &"Byte"
		  Else
			 Response.Write FormatNumber(p_FileItem.Size/1024,1,-1) &"KB"
		  End if
		  %>
						</font>&nbsp;</td>
					<td width="19%" class="hback">
						<% = p_FileItem.DateLastModified %>
					</td>
					<td width="8%" class="hback">
						<div align="center">文件</div>
					</td>
					<td width="19%" class="hback">
						<div align="left"><a href="FileManage.asp?Action=Delfile&File=<% = p_FileItem.name %>&Dir=<%=PathName%>" onClick="{if(confirm('确定删除此文件吗？')){return true;}return false;}">删除</a>┆
							<%if lcase(p_FileExtName)="jpg" or lcase(p_FileExtName)="gif"or lcase(p_FileExtName)="png" or lcase(p_FileExtName)="bmp" or lcase(p_FileExtName)="jpeg" then%>
							<a href="FileManage.asp?Action=joinphoto&File=<% = p_FileItem.name %>&Dir=<%=str_ShowPath%>&s=<% = p_FileItem.size %>">加入相册</a>┆
							<%end if%>
							<a href="<%=str_ShowPath%>/<% = p_FileItem.name %>" target="_blank">预览</a></div>
					</td>
				</tr>
				<%
	  else
	  end if
	next
	%>
			</table>
			<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
				<tr class="hback">
					<td class="hback">提醒：双击目录进入下级目录</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr class="back">
		<td height="20"  colspan="2" class="xingmu">
			<div align="left">
				<!--#include file="Copyright.asp" -->
			</div>
		</td>
	</tr>
</table>
</body>
</html>
<%
dim CurrPath
CurrPath=str_ShowPath
Set Fs_User = Nothing
%>
<script>
function OpenFolder(Obj)
{
	location.href='FileManage.asp?ShowPath='+Obj.Path;
}
function OpenParentFolder(Obj)
{
	location.href='FileManage.asp?ShowPath='+Obj.Path;
}
function EditFolder(filename)
{
	var ReturnValue='';
	ReturnValue=prompt('修改的名称：\n修改后，可能对您前台的图片文件路径有所影响！请小心修改',filename);
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
			window.location.href='?Type=FolderReName&OldFileName='+filename+'&NewFileName='+ReturnValue;
		}
	}
	else
	{
		alert('请填写要更名的名称');
	}
}
function UpFileo()
{
	OpenWindow('Commpages/Frame.asp?FileName=User_UpFileForm.asp&Path=<%= Request("ShowPath") %>',350,170,window);
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
			window.location.href='?Type=AddFolder&FolderName='+ReturnValue;
		}
	}
}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->

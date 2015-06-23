<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Cls_Cache.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("MF_Const") then Err_Show
if not MF_Check_Pop_TF("MF012") then Err_Show
Dim Path,FileName,EditFile,FileContent,Result,strShowErr
Result = Request.Form("Action")
Path = "../FS_InterFace/Public_Log"
FileName = "RefreshTimeSet.ini"
EditFile = Server.MapPath(Path) & "\" & FileName
Dim FsoObj,FileObj,FileStreamObj
Set FsoObj = Server.CreateObject(G_FS_FSO)
Set FileObj = FsoObj.GetFile(EditFile)
if Result = "" then
	Set FileStreamObj = FileObj.OpenAsTextStream(1)
	if Not FileStreamObj.AtEndOfStream then
		FileContent = FileStreamObj.ReadAll
	else
		FileContent = ""
	end if
else
	on error resume next
	Set FileStreamObj = FileObj.OpenAsTextStream(2)
	FileContent = Request.Form("ConstContent")
	FileStreamObj.Write FileContent
	if Err.Number <> 0 then
		strShowErr = "<li>保存失败</li><li>"& Err.Description &"</li><li>可能是您的刷新文件没有开启读写功能</li>"
		Response.Redirect("Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		strShowErr = "<li>自动刷新保存成功</li>"
		Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
end if
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
    <td class="xingmu">配置文件</td>
  </tr>
  <tr> 
    <td class="hback"><a href="SysConstSet.asp">全局变量设置</a>┆<a href="SysRefreshset.asp">自动刷新配置文件</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form action="" method="post" name="SysPara" id="SysPara">
    <tr class="hback"> 
      <td align="right"> <div align="center"> 
          <textarea name="ConstContent" rows="22" style="width:98%;"><% = FileContent %></textarea>
        </div></td>
    </tr>
    <tr class="hback"> 
      <td align="right"><div align="left"><span class="tx">特别说明:此设置为自动配置刷新服务而设定，此文件注释文件以&quot;;&quot;开始，而不是以&quot;'&quot;开始，请把注释写在&quot;;&quot;后面.如果你的是虚拟主机空间，请把[IsTF]设置为0，设置为1无效</span></div></td>
    </tr>
    <tr class="hback"> 
      <td align="right"><div align="left"> 
          <input type="Button" name="Submit" value=" 保存 " onClick="{if(confirm('确认保存吗？\n请确认您修改的确认无误!!!\n否则整个站点将无法正常运行!!!')){this.document.SysPara.submit();return true;}return false;}"/>
          <input type="reset" name="Submit2" value=" 重置 " />
          <input name="Action" type="hidden" id="Action" value="Save">
        </div></td>
    </tr>
  </form>
</table>
<br />
</body>
</html>







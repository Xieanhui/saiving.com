<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->

<%'Copyright (c) 2006 Foosun Inc.   UeUo.cn
Dim Conn,sRootDir
Function getCodePath(selectedJS)
	Dim Path,tmprs,MF_Domain
	MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
	MF_Default_Conn
	set tmprs = Conn.execute("Select FileSavePath,FileName from FS_NS_Sysjs where ID="&CintStr(selectedJS)&"")
	if  tmprs.eof then 
		getCodePath = "无效的ID"
	else
		path = tmprs("FileSavePath")&"/"&tmprs("FileName")&".js"
		path = replace(path,"//","/")
		if left(path,1) = "/" then path = mid(path,2)
		path = "/"&path
		getCodePath="<script src='"&path&"' language=""javascript""></script>"
	end if	
	Conn.close
	Set Conn=nothing
End function
%>
<HTML>
<HEAD>
<TITLE>CMS5.0</TITLE>
<link href="../../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>
<table width="75%" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr> 
    <td>该JS调用代码为:</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="2"> 
	<div align="center"> 
        <textarea name="textfield" cols="60" id="codepath" rows="4"><%=getCodePath(NoSqlHack(request.queryString("jsid")))%></textarea>
      </div></td>
  </tr>
  <tr> 
    <td> <div align="center"> 
        <input type="button" name="Submit" value=" 关 闭 " onClick="window.close();">
      </div></td>
	  <td> <div align="center"> 
        <input type="button" name="copy" value=" 复 制 " onClick="copyToClipBoard();">
      </div></td>
  </tr>
  <tr> 
    <td height="10" colspan="2">&nbsp;</td>
  </tr>
</table>
</BODY>
<script  language="JavaScript">  
<!--  
function copyToClipBoard()
{
	var clipBoardContent=document.getElementById("codepath").value
	window.clipboardData.setData("Text",clipBoardContent);
	alert("复制成功\n<%=G_COPYRIGHT%>");
}

//-->  
</script>  
</HTML>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






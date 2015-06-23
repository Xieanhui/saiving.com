<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%'Copyright (c) 2006 Foosun Inc. 
Function getCodePath(selectedJS)
	Dim Path
	Path = "/Vote/VoteJs.asp?TID="&NoSqlHack(selectedJS)&"&InfoID=Vote_HTML_ID_"&NoSqlHack(selectedJS)&"&PicW=60"
	getCodePath=getCodePath & "<script src='"&path&"' language=""javascript""></script>"&vbNewLine
	getCodePath=getCodePath & "<span id=""Vote_HTML_ID_"&selectedJS&""">正在加载...</span>"
End function
%>
<HTML>
<HEAD>
<TITLE>CMS5.0</TITLE>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>
<table width="75%" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr> 
    <td colspan="2">该JS调用代码为:</td>
  </tr>
  <tr> 
    <td colspan="2"> 
	<div align="center"> 
        <textarea name="textfield" style="width:90%;height:140" id="codepath"><%=getCodePath(Trim(NoSqlHack(request.queryString("jsid"))))%></textarea>
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






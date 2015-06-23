<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../../FS_InterFace/CLS_Foosun.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="cls_main.asp" -->
<!--#include file="Cls_Js.asp"-->
<%
Dim Conn
Function getCodePath(selectedJS)
	Dim FS_Js,Path,DoMain,DomainStr,DoMainRs
	MF_Default_Conn
	Set FS_Js=New Cls_Js
	FS_Js.getFreeJsParam(selectedJS)
		path=Conn.execute("SELECT MF_Domain FROM FS_MF_Config")(0)&"/"&Conn.execute("Select NewsDir from FS_NS_SysParam")(0)&"/js/FreeJs/"&FS_Js.eName&".js"
		path=Replace(Replace(path,"///","/"),"//","/")
	getCodePath="<script src='http://"&path&"' language=""javascript""></script>"
	Set DoMain=Nothing
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
        <textarea name="textfield" cols="60" id="codepath" rows="4"><%=getCodePath(Trim(NoSqlHack(request.queryString("jsid"))))%></textarea>
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
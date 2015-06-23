<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%
Dim Conn
MF_Default_Conn

Dim CID,ClassID
CID = CintStr(Request.QueryString("CID"))
ClassID = CintStr(Request.QueryString("ClassID"))
If CID = "" Or Not IsNumeric(CID) Or ClassID = "" Or Not IsNumeric(ClassID) Then
	Response.Write "<script>alert('参数错误，页面即将关闭');window.close();</script>" : Response.End
End If

Dim Com_Obj
Set Com_Obj = Server.CreateObject(G_FS_RS)
Com_Obj.Open "Select ComName,ComDescryption,ComAddress,ComWebSite,ComPrice,ComContact from FS_MS_Company where IsLock=0 and ComClass=" & CintStr(ClassID) & " And ComID = " & CintStr(CID),Conn,1,1
If Com_Obj.Bof And Com_Obj.Eof then
	Response.Write "<script>alert('参数错误，页面即将关闭');window.close();</script>" : Response.End
End If	 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>CMS5.0</title>
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" >
	  查看物流公司信息
	  </td>
	</tr>
    <tr  class="hback"> 
      <td width="20%" align="right" valign="middle">公司名称：</td>
      <td width="80%" align="left" valign="middle"><% = Com_Obj(0) %></td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right" valign="middle">公司介绍：</td>
      <td width="80%" align="left" valign="middle">
	  <%
		If Com_Obj(1) = "" Or IsNull(Com_Obj(1)) Then
			Response.Write "无"
		Else
			Response.Write Replace(Replace(Replace(Server.HTMLEncode(Com_Obj(1))," ","&nbsp;"),Chr(13),"<br />"),Chr(10),VbNewline)
		End iF
	   %>
	   </td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right" valign="middle">公司地址：</td>
      <td width="80%" align="left" valign="middle"><% = Com_Obj(2) %></td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right" valign="middle">公司网站：</td>
      <td width="80%" align="left" valign="middle"><% = Com_Obj(3) %></td>
    </tr>
    <tr  class="hback" id="TR_ComPrice" style="display:"> 
      <td width="20%" align="right" valign="middle">收费标准：</td>
      <td width="80%" align="left" valign="middle"><% = Com_Obj(4) %></td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right" valign="middle">联系方式：</td>
      <td width="80%" align="left" valign="middle">
	  <%
		If Com_Obj(5) = "" Or IsNull(Com_Obj(5)) Then
			Response.Write "无"
		Else
			Response.Write Replace(Replace(Replace(Server.HTMLEncode(Com_Obj(5))," ","&nbsp;"),Chr(13),"<br />"),Chr(10),VbNewline)
		End iF
	   %>
	  </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4" height="30" align="center">
	  	<span onclick="Javascript:window.close();" style="cursor:pointer;">【关闭】</span>
	  </td>
    </tr>	
  <td width="20%">
</table>
<%
Com_Obj.CLose : Set Com_Obj = Nothing
COnn.CLose : Set Conn = Nothing
%>
</body>
</html>
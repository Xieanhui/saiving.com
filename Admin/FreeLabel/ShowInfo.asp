<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>自由标签管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="100">
<%
Dim InfoTitle,InfoType,InfoContent,ReturnUrl
InfoTitle = Request.QueryString("Str_T")
InfoType = Request.QueryString("Str_C")
InfoContent = Request.QueryString("Con_Str")
ReturnUrl = Request.QueryString("Str_U")

If InfoTitle <> "" Then
	InfoTitle = "<span class=""tx"" style=""font-size:14px; font-weight:bold"">" & InfoTitle & "</span>"
Else
	InfoTitle = "<span class=""tx"" style=""font-size:14px; font-weight:bold"">未知错误</span>"
End If
If InfoType = "ER" Then
	If InfoContent <> "" Then
		InfoContent = "<span class=""tx"">" & InfoContent & "</span>"
	Else
		InfoContent = "<span class=""tx"">未知错误</span>"
	End IF
ElseIf InfoType = "OK" Then
	InfoContent = InfoContent & "<li><a href=""AddFreeOne.asp?Act=Add"">继续添加</a></li>"
End If
If ReturnUrl <> "" then
	ReturnUrl = "<li><a href=""" & ReturnUrl & """>返回</a></li>"
Else
	ReturnUrl = "<li><span onClick=""javascript:history.back();"" style=""cursor:hand;"">返回</span></li>"
End IF			
%>
<table width="50%" height="50" border="0" align="center" cellpadding="4" cellspacing="1" class="table" style="margin-top:100px;">
  <tr class="hback" >
    <td width="100%" height="25"  align="Left" class="xingmu" valign="middle"><% = InfoTitle %></td>
  </tr>
  <tr class="hback" >
    <td height="25" align="left" class="hback" valign="middle">
	<% = InfoContent & ReturnUrl %>
	</td>
  </tr>
</table>
</body>
</html>







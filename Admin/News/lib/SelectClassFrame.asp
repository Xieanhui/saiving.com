<%
dim query
query = Request.QueryString
if query<>"" then
	query="?"&query
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>选择栏目</title>
</head>
<link href="../../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<body scroll="no" ondragstart="return false;" onselectstart="return false;" topmargin="0" leftmargin="0">
<table width="100%" height="100" border="0" cellpadding="0" cellspacing="0">
  <tr>
	<td colspan="2"><iframe id="LableList" src="SelectClass.asp<%= query %>" scrolling="yes" width="100%" height="330" frameborder="1"></iframe></td>
</tr>
<tr>
    <td width="50%" align="center" valign="middle">
		<span class="tx">双击栏目名称进行选择</span>
	</td>
	<td width="50%" height="30" align="center" valign="middle">
        <input name="Submitasd" onClick="window.close();" type="button" id="Submitasd" value=" 确 定 ">
    </td>
</tr>
</table>
</body>
</html>






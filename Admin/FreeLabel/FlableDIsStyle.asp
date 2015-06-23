<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<title>自由标签管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<div style="width:100%; height:100%; text-align:center;">
<%
Dim DisStr
DisStr = Request.QueryString("ConStr")
Response.Write DisStr
%>
</div>
</body>
</html>







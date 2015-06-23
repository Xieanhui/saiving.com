<% Option Explicit %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><% = Request("PageTitle") %></title>
<%
Dim RequestItem,ParaList,FileName,Url
ParaList = ""
For Each RequestItem In Request.QueryString
	if RequestItem <> "FileName" And RequestItem <> "PageTitle" then
		if ParaList = "" then
			ParaList = RequestItem & "=" & Server.URLEncode(Request.QueryString(RequestItem))
		else
			ParaList = ParaList & "&" & RequestItem & "=" & Server.URLEncode(Request.QueryString(RequestItem))
		end if
	end if
Next
FileName = Request("FileName")
if FileName <> "" then
	Url = FileName & "?" & ParaList
else
	%>
	<script language="JavaScript">
		alert('文件不存在');
		window.close();
	</script>
	<%
	Response.End
end if
%>
</head>
<body scrolling=no bgcolor="#E6E6E6" topmargin="0" leftmargin="0">
<iframe src=<% = Url %> style="width:100%;height:100%;"  frameborder=0 scrolling="auto"></iframe>
</body>
</html>







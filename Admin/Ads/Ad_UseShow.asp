<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�����ô���___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
Dim AdID,str_TempStr
AdID=Request.QueryString("ID")
If AdID="" or IsNull(AdID) Then
	str_TempStr="<textarea name=""code"" cols=""60"" rows=""5"">��������</textarea>"
Else
	If IsNumeric(AdID)=False Then
		str_TempStr="<textarea name=""code"" cols=""60"" rows=""5"">��������</textarea>"
	Else
		str_TempStr="<textarea name=""code"" cols=""60"" rows=""5""><script language=""javascript"" src=""/Ads/"&AdID&".js""></script></textarea>"
	End If
End If
str_TempStr="<table width=""100%"" border=""0""><tr align=""center""><td>"&str_TempStr&"</td</tr><tr><td height=""10""></td></tr><tr align=""center""><td><input type=""button"" value=""���ƴ���"" onclick=""javascript:copyToClipBoard();"">  <input type=""button"" value=""�رմ���"" onclick=""javascript:window.close();""></td></tr></table>"
Response.Write(str_TempStr)
%>
</body>
</html>
<script language="javascript">
function copyToClipBoard()
{
	var clipBoardContent=document.getElementById("code").value
	window.clipboardData.setData("Text",clipBoardContent);
	alert("���Ƴɹ�");
}
</script>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->







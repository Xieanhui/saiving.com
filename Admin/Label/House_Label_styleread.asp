<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/HS_Function.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Dim Conn,obj_label_Rs,SQL
	MF_Default_Conn
	'session�ж�
	MF_Session_TF 
	If Request("id")="" Then
		Response.Write "<script language=""JavaScript"" type=""text/javascript"" src=""../../FS_Inc/PublicJs.js""></script>"&vbcrlf
		Response.Write "<SCRIPT LANGUAGE=""JavaScript"">ShowErr(""����\n��������\n�Ҳ���id"");</SCRIPT>"
		Response.End
	End If
%>
<html>
<head>
<title>���ű�ǩ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<base target=self>
</head>
<body>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<table width="98%" height="81" border="0" align="center" cellpadding="4" cellspacing="0">
  <tr > 
    <td width="100%" align="center">  
	<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
        <%
			dim regEx,result
			Set obj_label_Rs = server.CreateObject(G_FS_RS)
			SQL = "Select  StyleName,Content,AddDate from FS_MF_Labestyle where id=" & CintStr(Request.QueryString("id"))
			obj_label_Rs.Open SQL,Conn,1,3
			if not obj_label_Rs.eof then
				Set regEx = New RegExp '
				regEx.Pattern = "<img(.+?)>" '  
				regEx.IgnoreCase = false ' 
				regEx.Global = True '  
			    result = regEx.replace(obj_label_Rs("Content"),"<img src='../images/default.png'/>") 
			%>
        <tr style="display:"  class="hback"> 
          <td height="42" colspan="2"  class="hback"><div><%= result%></div></td>
        </tr>
        <%
		obj_label_Rs.close:set obj_label_Rs = nothing
		  End if
		%>
      </table>
</body>
</html>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,tmp_type,strShowErr
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("SS_site") then Err_Show
if not MF_Check_Pop_TF("SS001") then Err_Show
%>
<html>
<head>
<title>访问者地区统计</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script src="../../FS_Inc/PublicJS.js" language="JavaScript"></script>
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2">
<%
Dim RsAreaObj,Sql
Set RsAreaObj = Server.CreateObject(G_FS_RS)
Sql="Select Area From FS_SS_Stat"
RsAreaObj.Open Sql,Conn,3,3
Dim AreaType
Dim NumIn,NumOut,NumOther
NumIn=0
NumOut=0
NumOther=0
Do While not RsAreaObj.Eof
	AreaType= RsAreaObj("Area")
	Select Case AreaType
	Case "局域网内部网"
		NumIn=NumIn+1
	Case "未知区域"
		NumOther=NumOther+1
	Case Else
		NumOut=NumOut+1
	End Select
	RsAreaObj.MoveNext
Loop
%>
<%
Dim AllNum
AllNum=NumIn+NumOut+NumOther
%>
<table width=98% border=0 align="center" cellpadding=5 cellspacing="1" class="table">
  <tr>
    <td align=center class="xingmu"><div align="left"><strong>访问者地区统计</strong></div></td>
	</tr>
	<tr class="hback">
		<td align=center>
			<table align=left>
        <tr valign=bottom >
					
          <td nowap>内部网</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../Images/bar2.gif height=15></td>
					<td nowap><% =NumIn %></td>
				</tr>
				<tr valign=bottom >
					
          <td nowap>外部网</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../Images/bar2.gif height=15></td>
					<td nowap><% =NumOut %></td>
				</tr>
				<tr valign=bottom >
					
          <td align="right" nowap>未知</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../Images/bar2.gif height=15></td>
					<td nowap><% =NumOther %></td>
				</tr>
				<tr valign=cente>
					<td align=right nowap>共</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../Images/bar2.gif width="150" height=15></td>
					<td nowap><% = AllNum %></td>
			</table><br>
		</td>
	</tr>
</table>






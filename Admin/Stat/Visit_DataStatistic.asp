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

Dim Days,Months,Years,TempObj,SunObj,SunNum,VisitTodayNum,VisitMonthNum,VisitAllNums,TempObjs,TeempTimestr
Days = Day(Now())
Months = Month(Now())
Years = Year(Now())
SunNum = 0
Set TempObj = Conn.Execute("Select WebCountTime from FS_SS_SysPara")
If TempObj.eof then
	TeempTimestr = Now()
Else
	TeempTimestr = TempObj("WebCountTime")
End If
If G_IS_SQL_DB=0 then
	Set SunObj = Conn.Execute("Select LoginNum from FS_SS_Stat where VisitTime>#"&TeempTimestr&"#")
Else
	Set SunObj = Conn.Execute("Select LoginNum from FS_SS_Stat where VisitTime>'"&TeempTimestr&"'")
End If
If Not SunObj.eof then
Do while not SunObj.eof
	SunNum = SunNum + clng(SunObj("LoginNum"))
	SunObj.MoveNext
Loop
End If
SunObj.Close
Set SunObj = Nothing
Set TempObjs = Conn.Execute("Select Count(ID) from FS_SS_Stat where day(VisitTime) = '"&Days&"' and month(VisitTime)='"&Months&"' and year(VisitTime)='"&Years&"'")
	VisitTodayNum = Clng(TempObjs(0))
Set TempObjs = Conn.Execute("Select Count(ID) from FS_SS_Stat where month(VisitTime)='"&Months&"' and year(VisitTime)='"&Years&"'")
	VisitMonthNum = Clng(TempObjs(0))
If G_IS_SQL_DB=0 then
	Set TempObjs = Conn.Execute("Select Count(ID) from FS_SS_Stat where VisitTime>#"&TeempTimestr&"#")
Else
	Set TempObjs = Conn.Execute("Select Count(ID) from FS_SS_Stat where VisitTime>'"&TeempTimestr&"'")
End If
	VisitAllNums = Clng(TempObjs(0))
TempObjs.Close
Set TempObjs = Nothing
TempObj.Close
Set TempObj = Nothing
Conn.Execute("Update FS_SS_SysPara Set VisitToday="&NoSqlHack(VisitTodayNum)&",VisitMonth="&NoSqlHack(VisitMonthNum)&",VisitAllNum="&NoSqlHack(VisitAllNums)&",RefreashNum = "&Clng(SunNum)&"")
Dim Sql,RsWebObj,WebName,WebUrl,WebIntro,WebEmail,WebAdmin,WebCountTime,VisitAllNum,VisitToday,VisitMonth,RefreashNum
Sql ="Select * from FS_SS_SysPara"
Set RsWebObj = Server.CreateObject(G_FS_RS)
	RsWebObj.Open Sql,Conn,3,3
if Not RsWebObj.Eof then
	WebName = RsWebObj("WebName")
	WebUrl = RsWebObj("WebUrl")
	WebIntro = RsWebObj("WebIntro")
	WebEmail = RsWebObj("WebEmail")
	WebAdmin = RsWebObj("WebAdmin")
	WebCountTime = RsWebObj("WebCountTime")
	VisitAllNum = CLng(RsWebObj("VisitAllNum"))
	VisitToday = Clng(RsWebObj("VisitToday"))
	VisitMonth = RsWebObj("VisitMonth")
	RefreashNum = RsWebObj("RefreashNum")
else
	Response.Redirect"Visit_sysPara.asp"
	Response.end
end if
%>
<%
	Dim ForseeVisitToday,I,NumofDays,AverageNum,Tnum,TNumStr
	NumofDays=DATEDIFF("d",WebCountTime,Date())
	If NumofDays = 0 then
		NumofDays = 1
	End If
	AverageNum=CLng(VisitAllNum/NumofDays*1000)/1000
	I= Now()-Date()
	ForseeVisitToday = Round(VisitToday/I*(1-I) + VisitToday)
%>
<html>
<head>
<title>网站简要信息统计</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script src="../../FS_Inc/PublicJS.js" language="JavaScript"></script>
<body class="hback" topmargin="2" leftmargin="2">
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1"  class="table">
  <tr class="xingmu"> 
    <td height="30" colspan="2" class="xingmu"><strong>网站简要信息统计</strong></td>
  </tr>
  <tr class="hback"> 
    <td width="19%" height="30"> <div align="center">网站名称</div></td>
    <td width="81%"> <% = WebName %></td>
  </tr>
  <tr class="hback"> 
    <td height="30"> <div align="center">管 理 员</div></td>
    <td height="30"> <% = WebAdmin %></td>
  </tr>
  <tr class="hback"> 
    <td height="30"> <div align="center">网站地址</div></td>
    <td height="30"> <% = WebUrl %></td>
  </tr>
  <tr class="hback"> 
    <td height="30"> <div align="center">网站信箱</div></td>
    <td height="30"> <% = WebEmail %></td>
  </tr>
  <tr class="hback"> 
    <td height="30"> <div align="center">网站简介</div></td>
    <td height="30"> <% = WebIntro %></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="hback"> 
    <td width="20%"> 
      <div align="center">总 访问人数</div></td>
    <td width="30%"> 
      <% = VisitAllNum %></td>
    <td width="20%"> 
      <div align="center">开始统计日期</div></td>
    <td width="30%"> 
      <% = WebCountTime %></td>
  </tr>
  <tr class="hback"> 
    <td> 
      <div align="center">今日 访问量</div></td>
    <td> 
      <% =VisitToday %></td>
    <td> 
      <div align="center">本月 访 问量</div></td>
    <td> 
      <% =VisitMonth %></td>
  </tr>
  <tr class="hback"> 
    <td> 
      <div align="center">统 计 天 数</div></td>
    <td> 
      <% =NumofDays %></td>
    <td> 
      <div align="center">平均日访问量</div></td>
    <td> 
      <% = AverageNum%></td>
  </tr>
  <tr class="hback"> 
    <td> 
      <div align="center">整站页面刷新</div></td>
    <td> 
      <% =RefreashNum %></td>
    <td> 
      <div align="center">预计本日访问</div></td>
    <td> 
      <% =ForseeVisitToday %></td>
  </tr>
</table>
</body>
</html>
<%
	RsWebObj.Close
	Conn.Close
	Set RsWebObj=nothing
	Set Conn=nothing
%>






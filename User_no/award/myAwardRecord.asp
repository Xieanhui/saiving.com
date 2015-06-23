<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="lib/cls_award.asp"-->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
'no cache
response.expires=0 
response.addHeader "pragma" , "no-cache" 
response.addHeader "cache-control" , "private" 
'-------------------------------------------
Dim aType,sql_cmd,awardPeriod,awardObj
Dim awardRs,Rs,tr_count,eofTF,j
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,i
'---------------------------------分页定义
int_RPP=10 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"			'尾页
'-----------------------------------------

aType=NoSqlHack(request.QueryString("type"))
set awardObj=New cls_award
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<title><%=GetUserSystemTitle%></title>
<script language="javascript" src="../../FS_Inc/prototype.js"></script>
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
<tr>
<td class="xingmu" height="20"><img src="../images/award.gif" />我的中奖记录</td>
</tr>
<tr>
<td class="hback"><a href="myAwardRecord.asp?type=1">积分抽奖</a>&nbsp;|&nbsp;<a href="myAwardRecord.asp?type=2">积分兑换</a></td>
</tr>
</table>
<%if aType<>"2" then%>
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
	<tr><td class="hback">
	<%
		eofTF=false
		tr_count=0
		Set awardRs=Server.CreateObject(G_FS_RS)
		if G_IS_SQL_DB=0 then
			sql_cmd="Select awardID from FS_ME_award where opened=1"
		Else
			sql_cmd="Select awardID from FS_ME_award where opened=1"
		End if
		awardRs.open sql_cmd,User_Conn,1,1
		If Not awardRs.eof then
	'分页使用-----------------------------------
			awardRs.PageSize=int_RPP
			cPageNo=NoSqlHack(Request.QueryString("page"))
			If cPageNo="" Then cPageNo = 1
			If not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo<=0 Then cPageNo=1
			If cPageNo>awardRs.PageCount Then cPageNo=awardRs.PageCount 
			awardRs.AbsolutePage=cPageNo
		End if
		if awardRs.eof then
			eofTF=true
		End if
		for i=0 to int_RPP
			if not awardRs.eof then
				Set Rs=User_Conn.execute("Select awardID,prizeID from FS_ME_User_Prize where awardID="&awardRs("awardID")&" And userNumber='"&session("FS_UserNumber")&"' and winner=1")
				Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""table"">"&vbcrlf)
				Response.Write("<tr>"&vbcrlf)
				Response.Write("<td class=""hback"" colspan=""8""><img src=""../images/award.gif""/>第"&awardRs("awardID")&"期</td>"&vbcrlf)
				Response.Write("</tr>"&vbcrlf)
				while not Rs.eof
					awardObj.getPrizeInfo(Rs("PrizeID"))
					if tr_count Mod 4=0 then
						Response.Write("<tr>"&vbcrlf)
					End if
					Response.Write("<td class=""hback"" width=""10%"" height=""80""><img src="""&awardObj.PrizePic&""" width=""80"" height=""80""/></td>"&vbcrlf)
					Response.Write("<td class=""hback"" width=""90""><Div><img src=""../images/award.gif""/><font color=""red"">"&awardObj.PrizeName&"</font><img src=""../images/award.gif""/></div><div>"&awardObj.PrizeGrade&"等奖</div></td>"&vbcrlf)
					tr_count=tr_count+1
					if tr_count Mod 4=0 then
						Response.Write("</tr>"&vbcrlf)
					End if
					Rs.movenext
				Wend
				if tr_count Mod 4<>0 then
					for j=(tr_count Mod 4)+1 to 4 
						Response.Write("<td class=""hback""></td>"&vbcrlf)
					next
					Response.Write("</tr>"&vbcrlf)
				End if
				awardRs.movenext
				Response.Write("</table>"&vbcrlf)
			End if
		next
		if eofTF then
			Response.Write("<tr>"&vbcrlf)
			Response.Write("<td class=""hback""><img src=""../images/alert.gif""/>暂无中奖记录</td>"&vbcrlf)
			Response.Write("</tr>"&vbcrlf)
		End if
	%>
	</td>
	</tr>	 
	<%
	Response.Write("<tr>"&vbcrlf)
	Response.Write("<td align='right' colspan=""8""  class=""hback"">"&fPageCount(awardRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td>"&vbcrlf)
	Response.Write("</tr>"&vbcrlf)
	%>
	</table>
<%Elseif aType="2" Then%>
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
	<%
		int_RPP=20
		eofTF=false
		tr_count=0
		Set awardRs=Server.CreateObject(G_FS_RS)
		sql_cmd="Select distinct up.PrizeID from FS_ME_User_Prize as up,FS_ME_Prize as p where up.UserNumber='"&session("FS_UserNumber")&"' and up.winner=1 and isChange=1 and up.PrizeID=p.PrizeId"
		awardRs.open sql_cmd,User_Conn,1,3
		If Not awardRs.eof then
	'分页使用-----------------------------------
			awardRs.PageSize=int_RPP
			cPageNo=NoSqlHack(Request.QueryString("page"))
			If cPageNo="" Then cPageNo = 1
			If not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo<=0 Then cPageNo=1
			If cPageNo>awardRs.PageCount Then cPageNo=awardRs.PageCount
			on error resume next 
			awardRs.AbsolutePage=cPageNo
		End if
		if awardRs.eof then
			eofTF=true
		End if
		for i=0 to int_RPP
			if awardRs.eof then exit for
			awardObj.getPrizeInfo(awardRs("PrizeID"))
			if tr_count Mod 4=0 then
				Response.Write("<tr>"&vbcrlf)
			End if
			tr_count=tr_count+1
			Response.Write("<td class=""hback"" width=""10%"" height=""80""><img src="""&awardObj.PrizePic&""" width=""80"" height=""80""/></td>"&vbcrlf)
			Response.Write("<td class=""hback"" width=""90""><Div><img src=""../images/award.gif""/><font color=""red"">"&awardObj.PrizeName&"</font><img src=""../images/award.gif""/></div><div>积分消费：<img src=""../images/moneyOrPoint.gif""/> "&awardObj.prize_NeedPoint&"</div></td>"&vbcrlf)
			if tr_count Mod 4=0 then
				Response.Write("<tr>"&vbcrlf)
			End if
			awardRs.movenext
		next
		if tr_count Mod 4<>0 then
			for j=(tr_count Mod 4)+1 to 4 
				Response.Write("<td class=""hback""></td>"&vbcrlf)
				Response.Write("<td class=""hback""></td>"&vbcrlf)
			next
			Response.Write("</tr>"&vbcrlf)
		End if
	%>
	<%
	Response.Write("<tr>"&vbcrlf)
	Response.Write("<td align='right' colspan=""8""  class=""hback"">"&fPageCount(awardRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td>"&vbcrlf)
	Response.Write("</tr>"&vbcrlf)
	%>
	</table>
<%End if%>
</body>
</html>







<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="lib/cls_main.asp"-->
<%
Dim Conn,User_Conn,Constr_user_Rs,Constr_stat_Rs,sql_user_cmd,classid,ConstrObj
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,i
int_RPP=15 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"				'βҳ
'------------------------------------------------
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("NS034") then Err_Show
Set ConstrObj=New Cls_News
classid=NoSqlHack(Request.QueryString("classid"))
Set Constr_user_Rs=server.CreateObject(G_FS_RS)
sql_user_cmd="Select distinct UserNumber from FS_ME_InfoContribution"
Constr_user_Rs.open sql_user_cmd,User_Conn,1,1
function MonthStat(usernumber,Str_month,Str_year)
	Dim beginDate,endDate,ConstrNumber,ConstrAuditedNumber,lastMonthRs,lastMonthAuditedRs
	select case Str_month
		case "1" beginDate="-1-1"
				 endDate="-1-31"
		case "2" beginDate="-2-1"
				 endDate="-2-28"
		case "3" beginDate="-3-1"
				 endDate="-3-31"
		case "4" beginDate="-4-1"
				 endDate="-4-30"
		case "5" beginDate="-5-1"
				 endDate="-5-31"
		case "6" beginDate="-6-1"
				 endDate="-6-30"
		case "7" beginDate="-7-1"
				 endDate="-7-31"
		case "8" beginDate="-8-1"
				 endDate="-8-31"
		case "9" beginDate="-9-1"
				 endDate="-9-30"
		case "10" beginDate="-10-1"
				 endDate="-10-31"
		case "11" beginDate="-11-1"
				 endDate="-11-30"
		case "12" beginDate="-12-1"
				 endDate="-12-31"	
		case "0" beginDate="-12-1"
				 endDate="-12-31"	
	end select
	if Str_month=0 then
		Str_year=Str_year-1
	End if
	beginDate=Str_year&beginDate
	endDate=Str_year&endDate
	if  G_IS_SQL_DB=0 then
		Set lastMonthRs=User_Conn.execute("select count(contID) from FS_ME_InfoContribution where UserNumber='"&NoSqlHack(usernumber)&"' and addtime<#"&endDate&"# and addtime>#"&beginDate&"#")
		Set lastMonthAuditedRs=User_Conn.execute("select count(contID) from FS_ME_InfoContribution where UserNumber='"&NoSqlHack(usernumber)&"' and audittf=1 and addtime<#"&endDate&"# and addtime>#"&beginDate&"#")
	else
		Set lastMonthRs=User_Conn.execute("select count(contID) from FS_ME_InfoContribution where UserNumber='"&NoSqlHack(usernumber)&"' and addtime<'"&endDate&"' and addtime>'"&beginDate&"'")
		Set lastMonthAuditedRs=User_Conn.execute("select count(contID) from FS_ME_InfoContribution where UserNumber='"&NoSqlHack(usernumber)&"' and audittf=1 and addtime<'"&endDate&"' and addtime>'"&beginDate&"'")
	End if
	ConstrNumber=lastMonthRs(0)
	ConstrAuditedNumber=lastMonthAuditedRs(0)
	lastMonthRs.close
	lastMonthAuditedRs.close
	set lastMonthRs=nothing
	set lastMonthAuditedRs=nothing
	MonthStat=ConstrNumber&"/<font color=""red"">"&ConstrAuditedNumber&"</font>"
End function
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>CMS5.0</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language='javascript' src="../../FS_Inc/prototype.js"></script>
</head>

<body class="hback">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
<tr>
<td align="center" class="xingmu">�û���</td>
<td align="center" class="xingmu">Ͷ����</td>
<td align="center" class="xingmu">�������</td>
<td align="center" width="20%" class="xingmu">����Ͷ����</td>
</tr>
<%
	If Not Constr_user_Rs.eof then
		Constr_user_Rs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo<=0 Then cPageNo=1
		If cPageNo>Constr_user_Rs.PageCount Then cPageNo=Constr_user_Rs.PageCount 
		Constr_user_Rs.AbsolutePage=cPageNo
	End if
	for i=0 to int_RPP
		if Constr_user_Rs.eof then exit for
		Response.Write("<tr class=""hback"">"&vbcrlf)
		response.Write("<td align=""center"" class=""hback""><a href=""../../"& G_USER_DIR &"/showuser.asp?UserNumber="&Constr_user_Rs("UserNumber")&""" target=""_blank"">"&ConstrObj.GetUserName(Constr_user_Rs("UserNumber"))&"</a></td>"&vbcrlf)
		response.Write("<td align=""center"" class=""hback"">"&ConstrObj.newsStat(Constr_user_Rs("UserNumber"),0)&"</td>"&vbcrlf)
		response.Write("<td align=""center"" class=""hback"">"&ConstrObj.newsStat(Constr_user_Rs("UserNumber"),1)&"</td>"&vbcrlf)
		response.Write("<td align=""center"" class=""hback"">"&MonthStat(Constr_user_Rs("UserNumber"),month(now)-1,year(now))&"</td>"&vbcrlf)
		Constr_user_Rs.movenext
	next
%>
<%
Response.Write("<tr>"&vbcrlf)
Response.Write("<td align='right' colspan='4'  class=""hback"">"&fPageCount(Constr_user_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td>"&vbcrlf)
Response.Write("</tr>"&vbcrlf)
%>
</table>
</body>
</html>







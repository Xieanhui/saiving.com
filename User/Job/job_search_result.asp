<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--Copyright (c) 2006 Foosun Inc. Code by Einstein.liu-->
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<script language="javascript" src="../../FS_Inc/CheckJs.js"></script>
<script language="javascript" src="../../FS_Inc/prototype.js"></script>
<script language="javascript" src="../../FS_Inc/PublicJS.js"></script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td><!--#include file="../top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr class="back">
    <td   colspan="2" class="xingmu" height="26"><!--#include file="../Top_navi.asp" -->
    </td>
  </tr>
  <tr class="back">
    <td width="18%" valign="top" class="hback"><div align="left">
        <!--#include file="../menu.asp" -->
      </div></td>
    <td width="82%" valign="top" class="hback">
	<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
	<tr>
	<td colspan="20" class="xingmu">职位搜索结果</td>
	</tr>
	<tr>
	<td class="hback_1">职位名称</td>
	<td align="center" class="hback_1">公司名称</td>
	<td align="center" class="hback_1">工作地点</td>
	<td align="center" class="hback_1">简历语言</td>
	<td align="center" class="hback_1">招聘人数</td>
	<td align="center" class="hback_1">发布日期</td>
	<td align="center" class="hback_1">结束日期</td>
	<td align="center" class="hback_1">发送简历</td>
	</tr>
	<%
	
		''得到相关表的值。
		Function Get_OtherTable_Value(This_Fun_Sql)
			Dim This_Fun_Rs
			if instr(This_Fun_Sql," FS_ME_")>0 then 
				set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
			else
				set This_Fun_Rs = Conn.execute(This_Fun_Sql)
			end if			
			if not This_Fun_Rs.eof then 
				Get_OtherTable_Value = This_Fun_Rs(0)
			else
				Get_OtherTable_Value = ""
			end if
			if Err.Number>0 then 
				response.Redirect("../lib/error.asp?ErrCodes=<li>Get_OtherTable_Value未能得到相关数据。错误描述："&Err.Description&"</li>") : response.End()
			end if
			set This_Fun_Rs=nothing 
		End Function
	
	
	
		Dim Rs,job,workcity,publicdate,condition
		job = NoSqlHack(request.Form("JobName"))
		workcity = NoSqlHack(request.Form("txt_workcity"))
		publicdate = NoSqlHack(request.Form("hd_publicdate"))
		if job<>"" then
			condition=" And JobName like '"&job&"'"
		End if
		
		'-----------------------------------------
		if workcity<>"" then
			condition=condition&" And workcity ='"&workcity&"'"
		End if
		'------------------------------------------------------
		if publicdate<>"" then
			Dim dateCondition
			dateCondition=DateValue(Now())-Cint(publicdate)
			if G_IS_SQL_DB=0 then
				condition=condition&" And publicdate>#"&dateCondition&"#"
			Else
				condition=condition&" And publicdate>'"&dateCondition&"'"
			End if
		End if
		if G_IS_SQL_DB=0 then
			Set Rs=Conn.execute("Select PID,UserNumber,JobName,JobDescription,ResumeLang,WorkCity,PublicDate,EndDate,NeedNum,jlmode from FS_AP_Job_Public where 1=1 "&condition&" And (EndDate>#"&DateValue(Now())&"# or EndDate=#"&DateValue(Now())&"#)")
		else
			Set Rs=Conn.execute("Select PID,UserNumber,JobName,JobDescription,ResumeLang,WorkCity,PublicDate,EndDate,NeedNum,jlmode from FS_AP_Job_Public where 1=1 "&condition&" And (EndDate>'"&DateValue(Now())&"' or EndDate='"&DateValue(Now())&"')")
		End if
		Dim meRs,corpName,Introduct,Email,Phone
		while not Rs.eof
			Response.Write("<tr height='25' onMouseOver=overColor(this) onMouseOut=outColor(this) onClick=""javascript:Element.toggle('descripe_"&Rs("PID")&"')"">"&vbcrlf)
			Response.Write("<td class='hback'>"&Rs("JobName")&"</td>"&vbcrlf)
			Set meRs=Conn.execute("Select CompanyName,Introduct,Email,Phone from FS_AP_UserList where UserNumber='"&Rs("UserNumber")&"'")
			if not meRs.eof then
				corpName=meRs("CompanyName")
				Introduct = meRs("Introduct")
				Email = meRs("Email")
				Phone = meRs("Phone")
			Else
				corpName="无"
				Introduct = ""
				Phone = ""
				Email = ""
			End if
			meRs.close
			if Email="" then 
				Set meRs=User_Conn.execute("Select Email from FS_ME_Users where UserNumber='"&Rs("UserNumber")&"'")
				if not meRs.eof then Email = meRs(0)
				meRs.close
			end if	
			if Email<>"" then Email = "<a href=""mailto:"&Email&"?Subject=应聘贵公司:"&Rs("JobName")&"&Body="&f_Do_Thief("http://localhost/User/job/Pserson.asp")&""">"&Email&"</a>"
			
			
			Response.Write("<td class='hback' align='center'>"&corpName&"</td>"&vbcrlf)
			Response.Write("<td class='hback' align='center'>"&Rs("WorkCity")&"</td>"&vbcrlf)
			Response.Write("<td class='hback' align='center'>"&Rs("ResumeLang")&"</td>"&vbcrlf)
			Response.Write("<td class='hback' align='center'>"&Rs("NeedNum")&"</td>")
			Response.Write("<td class='hback' align='center'>"&Rs("PublicDate")&"</td>")
			Response.Write("<td class='hback' align='center'>"&Rs("EndDate")&"</td>")
			Response.Write("<td class='hback' align='center'>"&Email&"</td>")
			
			Response.Write("<tr id='descripe_"&Rs("PID")&"' style=""display:'none'""><td colspan=20>"&vbcrlf)

			response.Write("<table width=""98%"" border=""0"" align=""center"" cellpadding=""4"" cellspacing=""1"" class=""table"">"&vbcrlf)
			
			Response.Write("<tr><td class='hback' >职位说明："&Rs("JobDescription")&"</td></tr>"&vbcrlf)
			Response.Write("<tr><td class='hback' >应聘说明："&Rs("jlmode")&"</td></tr>"&vbcrlf)
			Response.Write("<tr><td class='hback' >联系电话："&Phone&"</td></tr>"&vbcrlf)
			
			response.Write("</table>"&vbcrlf)
			
			Response.Write("</td></tr>"&vbcrlf)
			Rs.movenext
		Wend
	%>
	</table>
	</td>
  </tr>
  <tr class="back">
    <td height="20"  colspan="2" class="xingmu"><div align="left">
        <!--#include file="../Copyright.asp" -->
      </div></td>
  </tr>
</table>
</body>
</html>
<%
''++++++++++++++++++++++++++++++++++++++++++++  
''根据网站地址采集网站HTML
Function f_Do_Thief(URL)
	Dim Thief_,myValue
	''++++++++++++++++++++++++++++++++++++++++++++  
	if URL="" then f_Do_Thief="":exit function	
	'Set Thief_ = New Cls_Thief
'	
'	Thief_.Source=URL
'	
'	Thief_.Method="GET"
'	Thief_.steal	
'	Thief_.noReturn
'	myValue=Thief_.Value
'	''                  Thief_.DeBug
'	Set Thief_=nothing
	myValue = Replacestr(myValue,"0:,else:"&myValue)
	if instr(myValue,"无法找到该页") then myValue=""
	f_Do_Thief = myValue
End Function

Set User_Conn=nothing
Set Conn=nothing
%>
<script language="javascript">
<!--
-->
</script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->







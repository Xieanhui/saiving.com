<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%'Copyright (c) 2006 Foosun Inc. 
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
if not Session("FS_UserNumber")<>"" then response.Redirect("../lib/error.asp?ErrCodes=<li>你尚未登陆,或过期.</li>&ErrorUrl=../login.asp") : response.End()

Dim Ap_Rs,Ap_Rs1
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=15 '设置每页显示数目
int_showNumberLink_=10 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings>9</font>"   			'首页 
toP10_=" <font face=webdings>7</font>"			'上十
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"				'尾页

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
	if Err<>0 then 
		response.Redirect("../lib/error.asp?ErrCodes=<li>Get_OtherTable_Value未能得到相关数据。错误描述："&Err.Description&"</li>") : response.End()
	end if
	set This_Fun_Rs=nothing 
End Function
  
Function Get_While_Info(Add_Sql)
	Dim Get_Html,This_Fun_Sql,ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp
	Dim PublicDate,EndDate,lastTime,Trade,Job,Province,City
	This_Fun_Sql = "select distinct A.UserNumber from FS_AP_Resume_BaseInfo A,FS_AP_Resume_Position B,FS_AP_Resume_WorkCity C where A.UserNumber=B.UserNumber AND B.UserNumber=C.UserNumber "
	PublicDate = NoSqlHack(request.QueryString("PublicDate"))
	EndDate = NoSqlHack(request.QueryString("EndDate"))
	if G_IS_SQL_DB = 1 then PublicDate = replace(PublicDate,"#","'"):EndDate = replace(EndDate,"#","'")
	Trade = CintStr(request.QueryString("Trade"))
	Job = NoSqlHack(request.QueryString("Job"))
	Province = CintStr(request.QueryString("Province"))
	City = NoSqlHack(request.QueryString("City"))
	if Trade<>"" then
		if isnumeric(Trade) then Trade = Get_OtherTable_Value("select Trade from FS_AP_Trade where TID="&Trade)
		New_Search_Str = and_where(New_Search_Str) & "B.Trade = '"&Trade&"'"
	end if	
	if Job<>"" then New_Search_Str = and_where(New_Search_Str) & "B.Job = '"&Job&"'"
	if Province<>"" then 
		if isnumeric(Province) then Province = Get_OtherTable_Value("select Province from FS_AP_Province where PID="&Province)
		New_Search_Str = and_where(New_Search_Str) & "C.Province = '"&Province&"'"
	end if
	if City<>"" then New_Search_Str = and_where(New_Search_Str) & "C.City = '"&City&"'"
	''时间段
	if PublicDate<>"" then New_Search_Str = and_where(New_Search_Str) & "lastTime "&PublicDate
	if EndDate<>"" then New_Search_Str = and_where(New_Search_Str) & "lastTime "&EndDate
	if New_Search_Str<>"" then This_Fun_Sql = and_where(This_Fun_Sql) & replace(New_Search_Str," where ","")
	if instr(Add_Sql,"order by")>0 then 
		if instr(Add_Sql,"Addr desc") then 
			Add_Sql = replace(Add_Sql,"Addr","C.Province desc,C.City")
		else	
			if instr(Add_Sql,"Addr") then Add_Sql = replace(Add_Sql,"Addr","C.Province,C.City")
		end if
		This_Fun_Sql = This_Fun_Sql &"  "& Add_Sql
	end if
	Str_Tmp = ""
	'response.Write(This_Fun_Sql)
	
	Dim UserNumberList
	UserNumberList = ""
	On Error Resume Next
	set Ap_Rs1 = Conn.execute(This_Fun_Sql)
	if Err<>0 then 
		response.Redirect("../lib/error.asp?ErrCodes=<li>查询错误："&Err.Description&"</li><li>请确保您填写的数据类型是否匹配</li>") : response.End()
	end if
	if Ap_Rs1.eof then 
	
		Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 height=30 align=""center"" class=""ischeck"">未查询到符合条件的信息"
		Get_Html = Get_Html &"</td></tr>"
	
	else
		do while not Ap_Rs1.eof 
			UserNumberList = UserNumberList&"','"&Ap_Rs1("UserNumber")	
			Ap_Rs1.movenext
		loop
		UserNumberList = mid(UserNumberList,4)
		'------------------
		This_Fun_Sql="select  *  from FS_AP_Resume_BaseInfo where UserNumber In ('"&FormatStrArr(UserNumberList)&"')"

		'response.Write(This_Fun_Sql)
		On Error Resume Next
		Set Ap_Rs = CreateObject(G_FS_RS)
		Ap_Rs.Open This_Fun_Sql,Conn,1,1	
		if Err<>0 then 
			response.Redirect("../lib/error.asp?ErrCodes=<li>查询错误："&Err.Description&"</li><li>请确保您填写的数据类型是否匹配</li>") : response.End()
		end if
		IF Ap_Rs.eof THEN
			Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 height=30 align=""center"" class=""ischeck"">未查询到符合条件的信息"
			Get_Html = Get_Html &"</td></tr>"
		else	
		dim sysaprs,int_PerCount
		int_PerCount = 0 
		set sysaprs = Conn.execute("select top 1 PerCount,InitCount from FS_AP_SysPara")
		if not sysaprs.eof then 
			if not isnull(sysaprs(0)) then int_PerCount = clng(sysaprs(0))
		end if
	
		Ap_Rs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("Page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo<=0 Then cPageNo=1
		If cPageNo>Ap_Rs.PageCount Then cPageNo=Ap_Rs.PageCount 
		Ap_Rs.AbsolutePage=cPageNo
		
		  FOR int_Start=1 TO int_RPP 
			Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
			Get_Html = Get_Html & "<td align=""center""><a href=""Person.asp?UID="&Ap_Rs("UserNumber")&""" title=""点击查看简历（普通会员需扣除相应点数"&int_PerCount&"点）"" target=_blank>"&Ap_Rs("Uname")&"</td>" & vbcrlf
			Get_Html = Get_Html & "<td align=""center"">"&Replacestr(Ap_Rs("sex"),"0:男,1:女")&"</td>" & vbcrlf
			Get_Html = Get_Html & "<td align=""center"">"&Ap_Rs("XueLi")&"</td>" & vbcrlf
			Get_Html = Get_Html & "<td align=""center"">"&Ap_Rs("Province")&" "&Ap_Rs("City")&"</td>" & vbcrlf
			Get_Html = Get_Html & "<td align=""center"">"&Replacestr(Ap_Rs("WorkAge"),"1:在读学生,2:应届毕业生,3:一年以上,4:两年以上,5:三年以上,6:五年以上,7:八年以上,8:十年以上")&"</td>" & vbcrlf
			Get_Html = Get_Html & "<td align=""center"">"&Ap_Rs("lastTime")&"</td>" & vbcrlf
			Get_Html = Get_Html & "</tr>" & vbcrlf
			Ap_Rs.MoveNext
			if Ap_Rs.eof or Ap_Rs.bof then exit for
		  NEXT
		END IF
		Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
		Get_Html = Get_Html & fPageCount(Ap_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
		Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
		Get_Html = Get_Html &"</td></tr>"
		Ap_Rs.close
	end if	
	Ap_Rs1.close:set Ap_Rs1=nothing

	Get_While_Info = Get_Html
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-已发布的招聘</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
//点击标题排序
var Old_Sql = document.URL;

function OrderByName(FildName)
{
	//alert(document.URL);	
	var New_Sql;
	if(Old_Sql.indexOf('Add_Sql')<0)
	{
		if(Old_Sql.indexOf('?')<0)
			New_Sql = Old_Sql + "?Add_Sql=order by " + FildName;	
		else
			New_Sql = Old_Sql + "&Add_Sql=order by " + FildName;	
	}
	else
	{
		if(Old_Sql.indexOf("Add_Sql=order by " + FildName + " desc")>-1)
		{
			New_Sql = Old_Sql.substring(0,Old_Sql.indexOf("Add_Sql=")) + "Add_Sql=order by " + FildName;
		}
		else
		{
			New_Sql = Old_Sql.substring(0,Old_Sql.indexOf("Add_Sql=")) + "Add_Sql=order by " + FildName + " desc";	
		}	
	}
	//alert(New_Sql);	
	location = New_Sql;
}
-->
</script>
<head>
<body>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="../top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="../Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="../menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr class="hback"> 
            
          <td class="hback"><strong>位置：</strong><a href="../../">网站首页</a> &gt;&gt; 
            <a href="../main.asp">会员首页</a> &gt;&gt; <a href="job_applications.asp">招聘首页</a>－人才查询</td>
          </tr>
        </table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	  <tr  class="hback"> 
		<td colspan="10" height="25">
		 <a href="AP_Person_Search.asp">首页</a>
		</td>
	  </tr>
</table>
<%
'******************************************************************
Call View
'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
          <tr  class="hback"> 
            <td align="center" class="xingmu" >姓名</td>
            <td align="center" class="xingmu">性别</td>
            <td align="center" class="xingmu">学历</td>
             <td align="center" class="xingmu">所在地/籍贯</td>
             <td align="center" class="xingmu">工作年限</td>
           <td align="center" class="xingmu">发布日期</td>
          </tr>
          <%
		response.Write( Get_While_Info( request.QueryString("Add_Sql") ) )
	%>
      </table>
<%End Sub%>
       </td>
    </tr>
    <tr class="back"> 
      <td height="20" colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="../Copyright.asp" -->
        </div></td>
    </tr>
 
</table>
<%
Set Ap_Rs=nothing
Set Fs_User = Nothing
%>
<script language="JavaScript">
<!--//判断后将排序完善.字段名后面显示指示
var Req_FildName;
var New_FildName='';
if (Old_Sql.indexOf("Add_Sql=order by ")>-1)
{
	if(Old_Sql.indexOf(" desc")>-1)
		Req_FildName = Old_Sql.substring(Old_Sql.indexOf("Add_Sql=order by ") + "Add_Sql=order by ".length , Old_Sql.indexOf(" desc"));
	else
		Req_FildName = Old_Sql.substring(Old_Sql.indexOf("Add_Sql=order by ") + "Add_Sql=order by ".length , Old_Sql.length);	
	
	if (document.getElementById('Show_Oder_'+Req_FildName)!=null)  
	{
		if(Old_Sql.indexOf(Req_FildName + " desc")>-1)
		{
			eval('Show_Oder_'+Req_FildName).innerText = '↓';
		}
		else
		{
			eval('Show_Oder_'+Req_FildName).innerText = '↑';
		}
	}	
}
-->
</script>

<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






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

Dim Ap_Rs
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

Function and_where(sql)
	if instr(lcase(sql)," where ")>0 then 
		and_where = sql & " and "
	else
		and_where = sql & " where "	
	end if
End Function
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
	Str_Tmp = "PID,JobName,JobDescription,ResumeLang,WorkCity,PublicDate,EndDate,NeedNum,EducateExp,Sex,WorkAge,Age,JobType,OtherJobDes,MoneyMonth,FreeMoney,OtherMoneyDes,HolleType"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_AP_Job_Public where UserNumber = '"&Session("FS_UserNumber")&"'"
	Arr_Tmp = split(Str_Tmp,",")
	for each Str_Tmp in Arr_Tmp
		Req_Str = NoSqlHack(Trim(request.QueryString(Str_Tmp)))
		if Req_Str<>"" then 				
			select case Str_Tmp
			case "PID","PublicDate","EndDate","NeedNum"
			''数字,日期
				regxp = "|<|>|=|<=|>=|<>|"
				if instr(regxp,"|"&left(Req_Str,1)&"|")>0 or instr(regxp,"|"&left(Req_Str,2)&"|")>0 then 
					New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" "& Req_Str
				elseif instr(Req_Str,"*")>0 then 
					if left(Req_Str,1)="*" then Req_Str = "%"&mid(Req_Str,2)
					if right(Req_Str,1)="*" then Req_Str = mid(Req_Str,1,len(Req_Str) - 1) & "%"							
					New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" like '"& Req_Str &"'"							
				else	
					New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" = "& Req_Str
				end if		
			case else
			''字符
				if left(Req_Str,1)="" then
					New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &left(Req_Str,1)&"'"& mid(Req_Str,2) &"'"
				elseif left(Req_Str,2)="<>" then
					New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" not like '%"& mid(Req_Str,3) &"%'"
				elseif instr(Req_Str,"*")>0 then 
					if left(Req_Str,1)="*" then Req_Str = "%"&mid(Req_Str,2)
					if right(Req_Str,1)="*" then Req_Str = mid(Req_Str,1,len(Req_Str) - 1) & "%"							
					New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" like '"& Req_Str &"'"							
				else	
					New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" like '%"& Req_Str &"%'"
				end if		
			end select 		
		end if
	next
	if New_Search_Str<>"" then This_Fun_Sql = and_where(This_Fun_Sql) & replace(New_Search_Str," where ","")
	if instr(Add_Sql,"order by")>0 then 
		This_Fun_Sql = This_Fun_Sql &"  "& Add_Sql
	end if
	Str_Tmp = ""
	On Error Resume Next
	Set Ap_Rs = CreateObject(G_FS_RS)
	Ap_Rs.Open This_Fun_Sql,Conn,1,1	

	if Err<>0 then 
		response.Redirect("../lib/error.asp?ErrCodes=<li>查询错误："&Err.Description&"</li><li>请确保您填写的数据类型是否匹配</li>") : response.End()
	end if

	IF  Ap_Rs.eof THEN
		Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 height=30 align=""center"" class=""ischeck"">未查询到符合条件的信息"
		Get_Html = Get_Html &"</td></tr>"
	else
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
		Get_Html = Get_Html & "<td align=""center"" style=""cursor:hand"" onclick=""javascript:if(TD_U_"&Ap_Rs("PID")&".style.display=='') TD_U_"&Ap_Rs("PID")&".style.display='none'; else TD_U_"&Ap_Rs("PID")&".style.display='';"" title='点击查看更多信息'>"&Ap_Rs("PID")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Ap_Rs("JobName")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Ap_Rs("PublicDate")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Ap_Rs("EndDate")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""AP_Job_Public_AddUpdate.asp?Act=Edit&PID="&Ap_Rs("PID")&""">编辑</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" name=""PID"" id=""PID"" value="""&Ap_Rs("PID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		''++++++++++++++++++++++++++++++++++++++点开时显示详细信息。
		Get_Html = Get_Html & "<tr class=""hback"" id=""TD_U_"& Ap_Rs("PID") &""" style=""display:none""><td colspan=20>" & vbcrlf
		Get_Html = Get_Html & "<table width=""100%"" height=""30"" border=""0"" cellspacing=""1"" cellpadding=""2"" class=""table"">" & vbcrlf 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td  colspan=10>职位描述:"&Ap_Rs("JobDescription")& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>接受简历语言:"&Ap_Rs("ResumeLang") & "</td><td>工作地点:"&Ap_Rs("WorkCity") & "</td><td>招聘人数:"&Ap_Rs("NeedNum")& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "</table>" & vbcrlf
		Get_Html = Get_Html &"</td></tr>" & vbcrlf
		''+++++++++++++++++++++++++++++++++++++++		
		Ap_Rs.MoveNext
 		if Ap_Rs.eof or Ap_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(Ap_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""submit"" name=""submit"" value="" 删除 "" onclick=""javascript:return confirm('确定要删除所选项目吗?');""></td>"
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	Get_Html = Get_Html &"</td></tr>"
	Ap_Rs.close
	Get_While_Info = Get_Html
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>已发布的招聘---网站内容管理系统</title>
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
            <a href="../main.asp">会员首页</a> &gt;&gt; <a href="job_applications.asp">招聘首页</a>－已发布的招聘</td>
          </tr>
        </table>

<%
'******************************************************************
Call View
'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" method="post" action="AP_Job_Public_Action.asp?Act=Del">
     <tr  class="hback"> 
      <td colspan="10" align="left" class="xingmu" >已发布的信息</td>
	</tr>
  <tr  class="hback"> 
    <td colspan="10" height="25">
	 <a href="AP_Job_Public_List.asp">已发布的信息</a> | <a href="AP_Job_Public_AddUpdate.asp?Act=Add">添加招聘信息</a> | <a href="AP_Job_Public_AddUpdate.asp?Act=Search">查询招聘信息</a>	
	</td>
  </tr>

   <tr  class="hback"> 
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('PID')" class="sd">序号</a> 
        <span id="Show_Oder_PID"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('JobName')" class="sd">职位名称</a> 
        <span id="Show_Oder_JobName"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('PublicDate')" class="sd">发布日期</a> 
        <span id="Show_Oder_PublicDate"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('EndDate')" class="sd">有效日期</a> 
        <span id="Show_Oder_EndDate"></span></td>
      <td align="center" class="xingmu">编辑</td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_While_Info( request.QueryString("Add_Sql") ) )
	%>
  </form>
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






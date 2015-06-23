<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<% 
Dim Ap_Rs,Ap_Sql
if not Session("FS_UserNumber")<>"" then response.Redirect("../lib/error.asp?ErrCodes=<li>你尚未登陆,或过期.</li>&ErrorUrl=../login.asp") : response.End()
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-已发布的招聘</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS_YanZheng.js"></script>
<script language="JavaScript" src="../../FS_Inc/coolWindowsCalendar.js"></script>
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
select case request.QueryString("Act")
	case "Add","Edit","Search"
	Add_Edit_Search
	case else
	response.Redirect("job_applications.asp")
end select
'******************************************************************
Sub Add_Edit_Search()
Dim PID,Bol_IsEdit,ResumeLang,jlmode
Dim EducateExp,Sex,WorkAge,Age,JobType,OtherJobDes,MoneyMonth,FreeMoney,OtherMoneyDes,HolleType
Bol_IsEdit = false
if request.QueryString("Act")="Edit" then 
	PID = CintStr(request.QueryString("PID"))
	if PID="" then response.Redirect("../lib/error.asp?ErrorUrl=&ErrCodes=<li>必要的ID没有提供</li>") : response.End()
	Ap_Sql = "select PID,UserNumber,JobName,JobDescription,ResumeLang,WorkCity,PublicDate,EndDate,NeedNum,jlmode,EducateExp,Sex,WorkAge,Age,JobType,OtherJobDes,MoneyMonth,FreeMoney,OtherMoneyDes,HolleType from FS_AP_Job_Public where UserNumber = '"&Session("FS_UserNumber")&"' and PID="&CintStr(PID)
	Set Ap_Rs	= CreateObject(G_FS_RS)
	'response.Write(Ap_Sql)
'	response.End()
	Ap_Rs.Open Ap_Sql,Conn,1,1
	if Ap_Rs.eof then response.Redirect("../lib/error.asp?ErrorUrl=&ErrCodes=<li>没有相关的内容,或该内容已不存在.</li>") : response.End()
	Bol_IsEdit = True
	ResumeLang = Ap_Rs("ResumeLang")
	jlmode= Ap_Rs("jlmode")
	EducateExp = Ap_Rs("EducateExp")
	Sex = Ap_Rs("Sex")
	WorkAge = Ap_Rs("WorkAge")
	Age = Ap_Rs("Age")
	JobType = Ap_Rs("JobType")
	HolleType = Ap_Rs("HolleType")
	FreeMoney = Ap_Rs("FreeMoney")
	OtherJobDes = Ap_Rs("OtherJobDes")
	MoneyMonth = Ap_Rs("MoneyMonth")
	OtherMoneyDes = Ap_Rs("OtherMoneyDes")
end if	
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" id="form1" onSubmit="return Validator.Validate(this,3);" <%if request.QueryString("Act")<>"Search" then response.Write(" method=""post"" action=""AP_Job_Public_Action.asp?Act=Save""") else response.Write(" method=""get"" action=""AP_Job_Public_List.asp""") end if%>>
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" ><%if Bol_IsEdit then
	  response.Write("修改招聘信息<input type=""hidden"" name=""PID"" ID=""PID"" value="""&Ap_Rs("PID")&""">")
	  elseif request.QueryString("Act")="Search" then
	  response.Write("查询已发布的招聘信息") 
	  else
	  response.Write("发布招聘信息")
	  end if%></td>
	</tr>
  <tr  class="hback"> 
    <td colspan="3" height="25">
	 <a href="AP_Job_Public_List.asp">已发布的信息</a> | <a href="AP_Job_Public_AddUpdate.asp?Act=Add">添加招聘信息</a> | <a href="AP_Job_Public_AddUpdate.asp?Act=Search">查询招聘信息</a>	
	</td>
  </tr>
<%if request.QueryString("Act")="Search" then%>  
    <tr  class="hback"> 
      <td align="right">ID序号</td>
      <td>
	  <input type="text" name="PID" id="PID" size="10" maxlength="10" value="">
      </td>
    </tr>  
<%end if%>  
    <tr  class="hback"> 
      <td align="right">职位名称</td>
      <td>
	  <input type="text" name="JobName" id="JobName" <%if request.QueryString("Act")<>"Search" then%> dataType="LimitB" min="2" max="30" msg="长度必须在2-30之间"<%end if%> size="60" maxlength="30" value="<%if Bol_IsEdit then response.Write(Ap_Rs("JobName")) end if%>">
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">职位描述</td>
      <td>
	  <textarea name="JobDescription" id="JobDescription" rows="10" cols="60"<%if request.QueryString("Act")<>"Search" then%> dataType="LimitB" min="10" max="3000" msg="长度必须在 10-3000之间"<%end if%>><%if Bol_IsEdit then response.Write(Ap_Rs("JobDescription")) end if%></textarea>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">接受简历语言</td>
      <td>
		<select name="ResumeLang" id="ResumeLang">
		<%=PrintOption(ResumeLang,"无:无,英语:英语,粤语:粤语,日语:日语,法语:法语,德语:德语,俄语:俄语,韩语:韩语,西班牙语:西班牙语,葡萄牙语:葡萄牙语")%>
		</select><span id="ResumeLang_Alt"></span>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">工作地点</td>
      <td>
	  <input type="text" name="WorkCity" id="WorkCity" size="30" maxlength="20" value="<%if Bol_IsEdit then response.Write(Ap_Rs("WorkCity")) end if%>"<%if request.QueryString("Act")<>"Search" then%> dataType="LimitB" min="2" max="20" msg="长度必须在 2-20之间"<%end if%>>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">学历要求</td>
      <td>
	  	<select name="EducateExp" id="EducateExp">
			<% = PrintOption(EducateExp,"0:不限,1:中专以下,2:大专,3:本科,4:硕士以上") %>
		</select>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">性别要求</td>
      <td>
	  <select name="Sex" id="Sex">
			<%=PrintOption(Sex,"0:男,1:女")%>
		</select>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">工作经验</td>
      <td>
	  	<select name="WorkAge" id="WorkAge">
			<%=PrintOption(WorkAge,"0:不限,1:在读学生,2:应届毕业生,3:一年以上,4:二年以上,5:三年以上,6:五年以上,7:八年以上,8:十年以上")%>
		</select>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">年龄要求</td>
      <td>
	  	<select name="Age" id="Age">
			<%=PrintOption(Age,"0:不限,1:18-22岁,2:20-25岁,3:22-28岁,4:22-30岁,5:25-35岁,6:35岁以上")%>
		</select>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">工作性质</td>
      <td>
	  	<select name="JobType" id="JobType">
			<%=PrintOption(JobType,"0:不限,1:兼职,2:全职,3:实习")%>
		</select>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">其他职位要求</td>
      <td>
	  	<textarea name="OtherJobDes" id="OtherJobDes" rows="5" cols="60"<%if request.QueryString("Act")<>"Search" then%> dataType="LimitB" min="0" max="3000" msg="长度必须在 0-3000之间"<%end if%>><% If Bol_IsEdit Then : Response.Write OtherJobDes : End If %></textarea>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">招聘人数</td>
      <td>
	  <input type="text" name="NeedNum" id="NeedNum" size="4" maxlength="3"<%if request.QueryString("Act")<>"Search" then%> dataType="Integer" msg="人数必须数字"<%end if%> value="<%if Bol_IsEdit then response.Write(Ap_Rs("NeedNum")) end if%>">
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">薪金待遇</td>
      <td>
	  	<input type="text" name="MoneyMonth" id="MoneyMonth" size="10" maxlength="7" value="<% If Bol_IsEdit Then : Response.Write  MoneyMonth : End If %>"<%if request.QueryString("Act")<>"Search" then%> dataType="Integer" msg="必须数字"<%end if%>>
		<input type="checkbox" name="FreeMoney" id="FreeMoney" value="1" <% If FreeMoney = "1" Then : Response.Write "checked" : Else : Response.Write "" : End If %>>面议
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">其他待遇说明</td>
      <td>
		<textarea name="OtherMoneyDes" id="OtherMoneyDes" rows="5" cols="60"<%if request.QueryString("Act")<>"Search" then%> dataType="LimitB" min="0" max="3000" msg="长度必须在 0-3000之间"<%end if%>><% If Bol_IsEdit Then : Response.Write OtherMoneyDes  : End If %></textarea>      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">应聘方式</td>
      <td>
	  	<select name="HolleType" id="HolleType">
			<%=PrintOption(HolleType,"0:不限,1:投递简历,2:电话咨询,3:直接面试")%>
		</select>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">应聘说明</td>
      <td>
	  <input type="text" name="jlmode" id="jlmode" size="60" maxlength="300" value="<%if Bol_IsEdit then response.Write(Ap_Rs("jlmode")) end if%>"<%if request.QueryString("Act")<>"Search" then%> dataType="LimitB" min="2" max="300" msg="长度必须在 2-300之间"<%end if%>>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">发布日期</td>
      <td>
	  <input type="text" name="PublicDate" id="PublicDate" onFocus="setday(this)" style="WIDTH: 100px; HEIGHT: 22px" maskType="shortDate" value="<%if Bol_IsEdit then response.Write(Ap_Rs("PublicDate")) else if request.QueryString("Act")<>"Search" then  response.Write(date()) end if end if%>">
	  <IMG id="img3" onClick="PublicDate.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom">
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">有效日期</td>
      <td>
	  <input type="text" name="EndDate" id="EndDate"  onfocus="setday(this)" style="WIDTH: 100px; HEIGHT: 22px" maskType="shortDate" value="<%if Bol_IsEdit then response.Write(Ap_Rs("EndDate")) end if%>" <%if request.QueryString("Act")<>"Search" then%>dataType="Require" msg="有效日期必须填写"<%end if%>>
	  <IMG id="img4" onClick="EndDate.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom">
      </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit"value=" <% If request.QueryString("Act")="Search" Then : Response.Write "查询" : Else : Response.Write "确定发布" : End If %> " /> 
              &nbsp; <input type="reset" value=" 重置 " />
  			  &nbsp; <input type="button" name="btn_todel" value=" 删除 " onClick="if(confirm('确定删除该项目吗？')) location='<%="AP_Job_Public_AddUpdate.asp?Act=Del&PID="&PID%>'">
            </td>
          </tr>
        </table>
      </td>
    </tr>	
  </form>
</table>
<%
End Sub
%>
       </td>
    </tr>
    <tr class="back"> 
      <td height="20" colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="../Copyright.asp" -->
        </div></td>
    </tr>
 
</table>
<%
Set Fs_User = Nothing

%>

<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






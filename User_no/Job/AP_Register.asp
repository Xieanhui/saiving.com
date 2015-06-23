<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<% 
Dim Ap_Rs,Ap_Sql
Dim UserNumber,C_Name
UserNumber = Session("FS_UserNumber")
if not UserNumber<>"" then response.Redirect("../lib/error.asp?ErrCodes=<li>你尚未登陆,或过期.</li>&ErrorUrl=../login.asp") : response.End()
C_Name = Get_OtherTable_Value("select C_Name from FS_ME_CorpUser where UserNumber='"&UserNumber&"'")
if C_Name = "" then response.Redirect("../lib/error.asp?ErrCodes=<li>你不是公司会员你没有资格进入.</li>&ErrorUrl=../login.asp") : response.End()
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
		Err.Clear
		response.Redirect("../lib/error.asp?ErrCodes=<li>Get_OtherTable_Value未能得到相关数据。错误描述："&Err.Description&"</li>") : response.End()
	end if
	set This_Fun_Rs=nothing 
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>注册公司---网站内容管理系统</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS_YanZheng.js"></script>
<script language="JavaScript" src="../../FS_Inc/coolWindowsCalendar.js"></script>
<script language="JavaScript">
<!--
setTimeout('showhide()',500);
function showhide()
{
	if(document.all.GroupLevel==null||document.all.GroupLevel=='undefined') return;
	if(document.all.GroupLevel.value=='1'){document.all.PuT.style.display='';document.all.BaoYue.style.display='none';document.all.BaoYue1.style.display='none'} else {document.all.PuT.style.display='none';document.all.BaoYue.style.display='';document.all.BaoYue1.style.display=''}
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
            <a href="../main.asp">会员首页</a> &gt;&gt; <a href="job_applications.asp">招聘首页</a>－注册公司扩展信息</td>
          </tr>
        </table>

<%
'******************************************************************
select case request.QueryString("Act")
	case "","Add","Edit","Search"
	Add_Edit_Search
	case else
	response.Redirect("job_applications.asp")
end select
'******************************************************************
Sub Add_Edit_Search()
Dim Bol_IsEdit,UserClass,GroupLevel,DocExt,Audited
Audited = 0
Bol_IsEdit = false
if UserNumber="" then response.Redirect("../lib/error.asp?ErrorUrl=&ErrCodes=<li>你没有登陆，或过期</li>") : response.End()
Ap_Sql = "select top 1 UserClass,GroupLevel,LeftCount,BeginDate,EndDate,CompanyName,Introduct,DocExt,DescriptApply,DescriptAudit,Audited,Phone,Email from FS_AP_UserList where UserNumber = '"&UserNumber&"'"
Set Ap_Rs	= CreateObject(G_FS_RS)
Ap_Rs.Open Ap_Sql,Conn,1,1
if not Ap_Rs.eof then 
	Bol_IsEdit = True
	UserClass = Ap_Rs("UserClass")
	GroupLevel = Ap_Rs("GroupLevel")
	DocExt = Ap_Rs("DocExt")
	Audited = Ap_Rs("Audited")
end if	
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" id="form1" onSubmit="return chkinput();" method="post" action="AP_Register_Action.asp?Act=Save">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" ><%if Bol_IsEdit then
	  response.Write("修改注册信息")
	  else
	  response.Write("填写注册信息（审核通过后就可以发布招聘信息啦）")
	  end if%></td>
	</tr>
  <tr  class="hback"> 
    <td colspan="3" height="25">
	 <a href="AP_Register.asp">首页</a>	
	</td>
  </tr>
    <tr  class="hback"> 
      <td align="right">会员类型</td>
      <td>
		<select name="UserClass" id="UserClass"<%if Audited=1 then response.Write(" disabled") end if%>>
		<%=PrintOption(UserClass,"1:用人单位,2:猎头公司")%>
		</select><span id="UserClass_Alt"></span>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">等级名称</td>
      <td>
		<select name="GroupLevel" id="GroupLevel" onChange="showhide()"<%if Audited=1 then response.Write(" disabled") end if%>>
		<%=PrintOption(GroupLevel,"1:普通会员,2:包月会员,3:VIP会员")%>
		</select><span id="UserClass_Alt"></span>
      </td>
    </tr>
    <tr class="hback" id="PuT" style="display:none"> 
      <td align="right">剩余点数</td>
      <td>
	  	<%if Bol_IsEdit then response.Write(Ap_Rs("LeftCount")) else response.Write("0") end if%>
      </td>
    </tr>
    <tr class="hback" id="BaoYue" style="display:none"> 
      <td align="right">生效日期</td>
      <td>
	  <input type="text" name="BeginDate" id="BeginDate" onFocus="setday(this)" style="WIDTH: 100px; HEIGHT: 22px" maskType="shortDate" value="<%if Bol_IsEdit then response.Write(Ap_Rs("BeginDate")) else response.Write(date()) end if%>"<%if Audited=1 then response.Write(" disabled") end if%>>
	  <IMG id="img3" onClick="BeginDate.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom"><span id="BeginDate_Alt"></span>
      </td>
    </tr>
    <tr  class="hback" id="BaoYue1" style="display:none"> 
      <td align="right">结束日期</td>
      <td>
	  <input type="text" name="EndDate" id="EndDate"  onfocus="setday(this)" style="WIDTH: 100px; HEIGHT: 22px" maskType="shortDate" value="<%if Bol_IsEdit then response.Write(Ap_Rs("EndDate")) end if%>"<%if Audited=1 then response.Write(" disabled") end if%>>
	  <IMG id="img4" onClick="EndDate.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom"><span id="EndDate_Alt"></span>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">公司名称</td>
      <td>
	  <input type="text" name="CompanyName" id="CompanyName" size="60" maxlength="30" value="<%if Bol_IsEdit then response.Write(Ap_Rs("CompanyName")) else response.Write(C_Name) end if%>"<%if request.QueryString("Act")<>"Search" then%> dataType="LimitB" min="4" max="30" msg="长度必须在4-30之间"<%end if%><%if Audited=1 then response.Write(" disabled") end if%>>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">公司简介</td>
      <td>
	  <textarea name="Introduct" cols="60" rows="20" id="Introduct"<%if Audited=1 then response.Write(" disabled") end if%>><%if Bol_IsEdit then response.Write(Ap_Rs("Introduct")) end if%></textarea>
      <span id="Introduct_Alt"></span>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">应聘表格</td>
      <td>
		<select name="DocExt" id="DocExt">
		<%=PrintOption(DocExt,"1:.doc,2:.rar,0:无表格")%>
		</select><span id="DocExt_Alt"></span>
      </td>
    </tr>
	
    <tr  class="hback"> 
      <td align="right">公司电话</td>
      <td>
	  <input type="text" name="phone" id="phone" size="40" maxlength="50"<%if request.QueryString("Act")<>"Search" then%> dataType="Require" msg="电话必须填写"<%end if%> value="<%if Bol_IsEdit then response.Write(Ap_Rs("phone")) end if%>">
      </td>
    </tr>

    <tr  class="hback"> 
      <td align="right">Email</td>
      <td>
	  <input type="text" name="Email" id="Email" size="15" maxlength="20"<%if request.QueryString("Act")<>"Search" then%> dataType="Require" msg="邮件必须填写"<%end if%> value="<%if Bol_IsEdit then response.Write(Ap_Rs("Email")) end if%>">
      </td>
    </tr>
	

    <tr  class="hback"> 
      <td align="right">申请备注</td>
      <td>
	  <textarea name="DescriptApply" cols="60" rows="5" id="DescriptApply"<%if Audited=1 then response.Write(" disabled") end if%>><%if Bol_IsEdit then response.Write(Ap_Rs("DescriptApply")) end if%></textarea>
      <span id="DescriptApply_Alt"></span>
      </td>
    </tr>
     <tr  class="hback"> 
      <td align="right">审核备注</td>
      <td>
	  <textarea name="DescriptAudit" disabled cols="60" rows="5" id="DescriptAudit"><%if Bol_IsEdit then response.Write(Ap_Rs("DescriptAudit")) end if%></textarea>
      <span id="DescriptAudit_Alt"></span>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">是否通过审核?</td>
      <td>
			<%=Replacestr(Audited,"1:已通过,0:<span class=tx>未通过</span>,else:未提交")%>
      </td>
    </tr>
   <tr  class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit"value=" 确定提交 "/> 
              &nbsp; <input type="reset" value=" 重置 " />
  			  &nbsp; <input type="button" name="btn_todel" value=" 撤消注册 " onClick="if(confirm('确定删除该项目吗？')) location='<%=server.URLEncode("AP_Register_Action.asp?Act=Del")%>'">
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
set AP_Rs = Nothing
%>


<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






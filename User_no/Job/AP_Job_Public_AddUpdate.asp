<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<% 
Dim Ap_Rs,Ap_Sql
if not Session("FS_UserNumber")<>"" then response.Redirect("../lib/error.asp?ErrCodes=<li>����δ��½,�����.</li>&ErrorUrl=../login.asp") : response.End()
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-�ѷ�������Ƹ</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
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
            
          <td class="hback"><strong>λ�ã�</strong><a href="../../">��վ��ҳ</a> &gt;&gt; 
            <a href="../main.asp">��Ա��ҳ</a> &gt;&gt; <a href="job_applications.asp">��Ƹ��ҳ</a>���ѷ�������Ƹ</td>
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
	if PID="" then response.Redirect("../lib/error.asp?ErrorUrl=&ErrCodes=<li>��Ҫ��IDû���ṩ</li>") : response.End()
	Ap_Sql = "select PID,UserNumber,JobName,JobDescription,ResumeLang,WorkCity,PublicDate,EndDate,NeedNum,jlmode,EducateExp,Sex,WorkAge,Age,JobType,OtherJobDes,MoneyMonth,FreeMoney,OtherMoneyDes,HolleType from FS_AP_Job_Public where UserNumber = '"&Session("FS_UserNumber")&"' and PID="&CintStr(PID)
	Set Ap_Rs	= CreateObject(G_FS_RS)
	'response.Write(Ap_Sql)
'	response.End()
	Ap_Rs.Open Ap_Sql,Conn,1,1
	if Ap_Rs.eof then response.Redirect("../lib/error.asp?ErrorUrl=&ErrCodes=<li>û����ص�����,��������Ѳ�����.</li>") : response.End()
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
	  response.Write("�޸���Ƹ��Ϣ<input type=""hidden"" name=""PID"" ID=""PID"" value="""&Ap_Rs("PID")&""">")
	  elseif request.QueryString("Act")="Search" then
	  response.Write("��ѯ�ѷ�������Ƹ��Ϣ") 
	  else
	  response.Write("������Ƹ��Ϣ")
	  end if%></td>
	</tr>
  <tr  class="hback"> 
    <td colspan="3" height="25">
	 <a href="AP_Job_Public_List.asp">�ѷ�������Ϣ</a> | <a href="AP_Job_Public_AddUpdate.asp?Act=Add">�����Ƹ��Ϣ</a> | <a href="AP_Job_Public_AddUpdate.asp?Act=Search">��ѯ��Ƹ��Ϣ</a>	
	</td>
  </tr>
<%if request.QueryString("Act")="Search" then%>  
    <tr  class="hback"> 
      <td align="right">ID���</td>
      <td>
	  <input type="text" name="PID" id="PID" size="10" maxlength="10" value="">
      </td>
    </tr>  
<%end if%>  
    <tr  class="hback"> 
      <td align="right">ְλ����</td>
      <td>
	  <input type="text" name="JobName" id="JobName" <%if request.QueryString("Act")<>"Search" then%> dataType="LimitB" min="2" max="30" msg="���ȱ�����2-30֮��"<%end if%> size="60" maxlength="30" value="<%if Bol_IsEdit then response.Write(Ap_Rs("JobName")) end if%>">
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">ְλ����</td>
      <td>
	  <textarea name="JobDescription" id="JobDescription" rows="10" cols="60"<%if request.QueryString("Act")<>"Search" then%> dataType="LimitB" min="10" max="3000" msg="���ȱ����� 10-3000֮��"<%end if%>><%if Bol_IsEdit then response.Write(Ap_Rs("JobDescription")) end if%></textarea>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">���ܼ�������</td>
      <td>
		<select name="ResumeLang" id="ResumeLang">
		<%=PrintOption(ResumeLang,"��:��,Ӣ��:Ӣ��,����:����,����:����,����:����,����:����,����:����,����:����,��������:��������,��������:��������")%>
		</select><span id="ResumeLang_Alt"></span>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�����ص�</td>
      <td>
	  <input type="text" name="WorkCity" id="WorkCity" size="30" maxlength="20" value="<%if Bol_IsEdit then response.Write(Ap_Rs("WorkCity")) end if%>"<%if request.QueryString("Act")<>"Search" then%> dataType="LimitB" min="2" max="20" msg="���ȱ����� 2-20֮��"<%end if%>>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">ѧ��Ҫ��</td>
      <td>
	  	<select name="EducateExp" id="EducateExp">
			<% = PrintOption(EducateExp,"0:����,1:��ר����,2:��ר,3:����,4:˶ʿ����") %>
		</select>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">�Ա�Ҫ��</td>
      <td>
	  <select name="Sex" id="Sex">
			<%=PrintOption(Sex,"0:��,1:Ů")%>
		</select>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">��������</td>
      <td>
	  	<select name="WorkAge" id="WorkAge">
			<%=PrintOption(WorkAge,"0:����,1:�ڶ�ѧ��,2:Ӧ���ҵ��,3:һ������,4:��������,5:��������,6:��������,7:��������,8:ʮ������")%>
		</select>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">����Ҫ��</td>
      <td>
	  	<select name="Age" id="Age">
			<%=PrintOption(Age,"0:����,1:18-22��,2:20-25��,3:22-28��,4:22-30��,5:25-35��,6:35������")%>
		</select>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">��������</td>
      <td>
	  	<select name="JobType" id="JobType">
			<%=PrintOption(JobType,"0:����,1:��ְ,2:ȫְ,3:ʵϰ")%>
		</select>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">����ְλҪ��</td>
      <td>
	  	<textarea name="OtherJobDes" id="OtherJobDes" rows="5" cols="60"<%if request.QueryString("Act")<>"Search" then%> dataType="LimitB" min="0" max="3000" msg="���ȱ����� 0-3000֮��"<%end if%>><% If Bol_IsEdit Then : Response.Write OtherJobDes : End If %></textarea>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">��Ƹ����</td>
      <td>
	  <input type="text" name="NeedNum" id="NeedNum" size="4" maxlength="3"<%if request.QueryString("Act")<>"Search" then%> dataType="Integer" msg="������������"<%end if%> value="<%if Bol_IsEdit then response.Write(Ap_Rs("NeedNum")) end if%>">
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">н�����</td>
      <td>
	  	<input type="text" name="MoneyMonth" id="MoneyMonth" size="10" maxlength="7" value="<% If Bol_IsEdit Then : Response.Write  MoneyMonth : End If %>"<%if request.QueryString("Act")<>"Search" then%> dataType="Integer" msg="��������"<%end if%>>
		<input type="checkbox" name="FreeMoney" id="FreeMoney" value="1" <% If FreeMoney = "1" Then : Response.Write "checked" : Else : Response.Write "" : End If %>>����
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">��������˵��</td>
      <td>
		<textarea name="OtherMoneyDes" id="OtherMoneyDes" rows="5" cols="60"<%if request.QueryString("Act")<>"Search" then%> dataType="LimitB" min="0" max="3000" msg="���ȱ����� 0-3000֮��"<%end if%>><% If Bol_IsEdit Then : Response.Write OtherMoneyDes  : End If %></textarea>      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">ӦƸ��ʽ</td>
      <td>
	  	<select name="HolleType" id="HolleType">
			<%=PrintOption(HolleType,"0:����,1:Ͷ�ݼ���,2:�绰��ѯ,3:ֱ������")%>
		</select>
      </td>
    </tr>
	<tr  class="hback"> 
      <td align="right">ӦƸ˵��</td>
      <td>
	  <input type="text" name="jlmode" id="jlmode" size="60" maxlength="300" value="<%if Bol_IsEdit then response.Write(Ap_Rs("jlmode")) end if%>"<%if request.QueryString("Act")<>"Search" then%> dataType="LimitB" min="2" max="300" msg="���ȱ����� 2-300֮��"<%end if%>>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��������</td>
      <td>
	  <input type="text" name="PublicDate" id="PublicDate" onFocus="setday(this)" style="WIDTH: 100px; HEIGHT: 22px" maskType="shortDate" value="<%if Bol_IsEdit then response.Write(Ap_Rs("PublicDate")) else if request.QueryString("Act")<>"Search" then  response.Write(date()) end if end if%>">
	  <IMG id="img3" onClick="PublicDate.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom">
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ч����</td>
      <td>
	  <input type="text" name="EndDate" id="EndDate"  onfocus="setday(this)" style="WIDTH: 100px; HEIGHT: 22px" maskType="shortDate" value="<%if Bol_IsEdit then response.Write(Ap_Rs("EndDate")) end if%>" <%if request.QueryString("Act")<>"Search" then%>dataType="Require" msg="��Ч���ڱ�����д"<%end if%>>
	  <IMG id="img4" onClick="EndDate.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom">
      </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit"value=" <% If request.QueryString("Act")="Search" Then : Response.Write "��ѯ" : Else : Response.Write "ȷ������" : End If %> " /> 
              &nbsp; <input type="reset" value=" ���� " />
  			  &nbsp; <input type="button" name="btn_todel" value=" ɾ�� " onClick="if(confirm('ȷ��ɾ������Ŀ��')) location='<%="AP_Job_Public_AddUpdate.asp?Act=Del&PID="&PID%>'">
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

<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






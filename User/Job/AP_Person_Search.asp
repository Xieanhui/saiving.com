<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%'Copyright (c) 2006 Foosun Inc. 
if not Session("FS_UserNumber")<>"" then response.Redirect("../lib/error.asp?ErrCodes=<li>����δ��½,�����.</li>&ErrorUrl=../login.asp") : response.End()
Function Get_FildValue_List(This_Fun_Sql,EquValue,Get_Type)
'''This_Fun_Sql ����sql���,EquValue�����ݿ���ͬ��ֵ�����<option>�����selected,Get_Type=1Ϊ<option>
Dim Get_Html,This_Fun_Rs,Text
On Error Resume Next
if instr(This_Fun_Sql,"FS_ME_")>0 then 
	set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
else	
	set This_Fun_Rs = Conn.execute(This_Fun_Sql)
end if	
If Err.Number <> 0 then Err.clear : response.Redirect("../error.asp?ErrCodes=<li>��Ǹ,�����Sql���������.�����ֶβ�����.</li>")
do while not This_Fun_Rs.eof 
	select case Get_Type
	  case 1
		''<option>		
		if instr(This_Fun_Sql,",") >0 then 
			Text = This_Fun_Rs(1)
		else
			Text = This_Fun_Rs(0)
		end if	
		if EquValue = This_Fun_Rs(0) then 
			Get_Html = Get_Html & "<option value="""&This_Fun_Rs(0)&"""  style=""color:#0000FF"" selected>"&Text&"</option>"&vbNewLine
		else
			Get_Html = Get_Html & "<option value="""&This_Fun_Rs(0)&""">"&Text&"</option>"&vbNewLine
		end if		
	  case else
		exit do : Get_FildValue_List = "Get_Typeֵ�������" : exit Function
	end select
	This_Fun_Rs.movenext
loop
This_Fun_Rs.close
Get_FildValue_List = Get_Html
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>�ѷ�������Ƹ---��վ���ݹ���ϵͳ</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/Prototype.js"></script>
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
            <a href="../main.asp">��Ա��ҳ</a> &gt;&gt; <a href="job_applications.asp">��Ƹ��ҳ</a>���˲Ų�ѯ</td>
          </tr>
        </table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	  <tr  class="hback"> 
		<td colspan="10" height="25">
		 <a href="AP_Person_Search.asp">��ҳ</a>
		</td>
	  </tr>
</table>
<%
'******************************************************************
Call Search
'******************************************************************
Sub Search()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" method="get" action="AP_Person_Search_Result.asp" onSubmit="return chkinput();">
    <tr  class="hback" id="TradeTR"> 
      <td align="right">��ҵ</td>
      <td width="600">
	  <select name="Trade" id="Trade" onChange="GetJobOptions(this.options[this.selectedIndex].value)">
	   <option value="">����</option>
	   <%=Get_FildValue_List("select TID,Trade from FS_AP_Trade","",1)%>
	  </select>
       ְλ <span id="JobSelect">��ѡ��</span>
	   </td>
    </tr>
    <tr  class="hback" id="ProvinceTR"> 
      <td align="right">�����ص�</td>
      <td>
	  ʡ<select name="Province" id="Province" onChange="GetCityOptions(this.options[this.selectedIndex].value)">
	   <option value="">����</option>
	   <%=Get_FildValue_List("select PID,Province from FS_AP_Province","",1)%>
	  </select>
	  ���� <span id="CitySelect">��ѡ��</span>
      </td>
    </tr>
	<tr  class="hback"> 
      <td width="133" align="right">������������ʼ����</td>
      <td>
	  <input type="text" name="PublicDate" id="PublicDate" readonly="" onFocus="setday(this)" style="WIDTH: 100px; HEIGHT: 22px" maskType="shortDate">
	  <IMG id="img3" onClick="PublicDate.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom"><span id="PublicDate_Alt"></span>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">���������Ľ�ֹ����</td>
      <td>
	  <input type="text" name="EndDate" id="EndDate" readonly=""  onfocus="setday(this)" style="WIDTH: 100px; HEIGHT: 22px" maskType="shortDate" value="<%=date+1%>">
	  <IMG id="img4" onClick="EndDate.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom"><span id="EndDate_Alt"></span>
      </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4"> <input type="submit" value=" �顤ѯ " /> 
              &nbsp; <input type="reset" value=" ���� " />
      </td>
    </tr>	
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
Set Fs_User = Nothing
%>
<script language="JavaScript">
<!--
function chkinput()
{
	if(document.all.Trade.value=='')
	{alert('��ҵ����ѡ��');document.all.Trade.focus();return false}
	if(document.all.PublicDate.value) if (document.all.PublicDate.value.indexOf('>=')<0) {document.all.PublicDate.value='>=#'+document.all.PublicDate.value+'#'};
	if(document.all.EndDate.value) if (document.all.EndDate.value.indexOf('<=')<0) {document.all.EndDate.value='<=#'+document.all.EndDate.value+'#'};
}
//�õ������˵�����
function GetJobOptions(Parentvalue)
{
	if (Parentvalue=='')
		$('JobSelect').innerHTML='��ѡ��';
	else	
		new Ajax.Updater('JobSelect', '../SetNextOptions.asp?SelectName=Job&EquValue='+Parentvalue+'&sType=1&no-cache='+Math.random() , 
		{method: 'post', onLoading:function(request) {$('Trade').disabled=true;$('TradeTR').style.cursor='wait'; $('JobSelect').innerHTML='<b>���ڶ�ȡ����,��ȴ�...</b>' },onComplete:function(request) {$('Trade').disabled=false; $('TradeTR').style.cursor='auto';  }});
		//fucxiע��, parameters:'ReqSql=<%=server.URLEncode("select Job from FS_AP_Job where TID=")%>'+Parentvalue});
}
//�õ������˵�����
function GetCityOptions(Parentvalue)
{
	if (Parentvalue=='')
		$('CitySelect').innerHTML='��ѡ��';
	else	
		new Ajax.Updater('CitySelect', '../SetNextOptions.asp?SelectName=City&EquValue='+Parentvalue+'&sType=1&no-cache='+Math.random() , {method: 'post', onLoading:function(request) {$('Province').disabled=true;$('ProvinceTR').style.cursor='wait'; $('CitySelect').innerHTML='<b>���ڶ�ȡ����,��ȴ�...</b>' },onComplete:function(request) {$('Province').disabled=false; $('ProvinceTR').style.cursor='auto';  }});
		//fucxiע��, parameters: 'ReqSql=<%=server.URLEncode("select City from FS_AP_City where PID=")%>'+Parentvalue}); 
}
-->
</script>

<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






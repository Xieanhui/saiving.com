<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="cls_resume.asp"-->
<%
Response.Charset="GB2312"
Dim resumeObj,id,Conn,action
MF_Default_Conn
id=NoSqlHack(request.QueryString("id"))
action=NoSqlHack(request.QueryString("action"))
Set resumeObj=New cls_resume
if id<>"" then call resumeObj.getResumeInfo("certificate",id)
Conn.close
Set Conn=nothing
%>
<form name="CertificateForm" action="AP_Resume_Action.asp?action=certificate&id=" method="post">
  <table width="100%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td width="19%" align="right" class="hback">获证时间：</td>
    <td width="81%" class="hback">
<select name="txt_year">
	<%dim ii,dbyear,dbmonth,dbday
	dbyear = resumeObj.FetchDate
	if isdate(dbyear) then 
		dbyear = year(dbyear)
		dbmonth = month(resumeObj.FetchDate)
		dbday  = day(resumeObj.FetchDate)
	else
		dbyear = 1980
		dbmonth = 1
	end if
	for ii= 1960 to 2010
	if cstr(ii)=cstr(dbyear) then 
		response.Write("<option value="""&ii&""" selected>"&ii&"</option>"&vbcrlf)
	else
		response.Write("<option value="""&ii&""">"&ii&"</option>"&vbcrlf)
	end if	
	next
	%>
	</select>
	<select name="txt_month">
	<%
	for ii= 1 to 12
		if cstr(ii)=cstr(dbmonth) then 
			response.Write("<option value="""&ii&""" selected>"&ii&"</option>"&vbcrlf)
		else
			response.Write("<option value="""&ii&""">"&ii&"</option>"&vbcrlf)
		end if	
	next
	%>
	</select>
	<select name="txt_day">
	<%
	for ii= 1 to 31
		if cstr(ii)=cstr(dbday) then 
			response.Write("<option value="""&ii&""" selected>"&ii&"</option>"&vbcrlf)
		else
			response.Write("<option value="""&ii&""">"&ii&"</option>"&vbcrlf)
		end if	
	next
	%>
	</select>	</td>
  </tr>
  <tr> 
    <td align="right" class="hback">证书名称：</td>
    <td class="hback">
	<input name="txt_Certificate" type="text" id="txt_Certificate" value="<%=resumeObj.Certificate%>" style="width:60%"
	onfocus="Do.these('txt_Certificate',function(){return isEmpty('txt_Certificate','span_Certificate')})"
	onKeyUp="Do.these('txt_Certificate',function(){return isEmpty('txt_Certificate','span_Certificate')})"
	/><span id="span_Certificate"></span></td>
  </tr>
  <tr>
    <td align="right" class="hback">分数（等级）：</td>
    <td class="hback">
	<input name="txt_Score" type="text" id="txt_Score" value="<%=resumeObj.Score%>" style="width:60%"
	onfocus="Do.these('txt_Score',function(){return isEmpty('txt_Score','span_Score')})"
	onKeyUp="Do.these('txt_Score',function(){return isEmpty('txt_Score','span_Score')})"
	/><span id="span_Score"></span></td>
  </tr>
  <tr>
    <td class="hback">&nbsp;</td>
    <td class="hback">
	<input type="hidden" name="txt_FetchDate" value="" />
	<input type="button" name="SubmitButton" value="保存/下一步" onclick="txt_FetchDate.value=txt_year.value+'-'+txt_month.value+'-'+txt_day.value ;ajaxPost('AP_Resume_Action.asp', Form.serialize('CertificateForm'),'CertificateForm','<%=action%>','<%=id%>');"/>
	&nbsp;&nbsp;
	<input type="reset" name="resetButton" value="重 设" onClick="javascript:if(confirm('确认清空所有表单输入？')){$('IntentionForm').reset();}else{return false;}" />	</td>
  </tr>
</table>
</form>







<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="cls_resume.asp"-->
<%
Response.Charset="GB2312"
Dim resumeObj,id,Conn,action
MF_Default_Conn
id=trim(NoSqlHack(request.QueryString("id")))
action=trim(NoSqlHack(request.QueryString("action")))
Set resumeObj=New cls_resume
if id<>"" then call resumeObj.getResumeInfo("educateexp",id)
Conn.close
Set Conn=nothing
%>
<form name="EducateExpForm" action="" method="post">
<table width="100%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td width="19%" align="right" class="hback">开始时间：</td>
    <td width="81%" class="hback">
	<select name="txt_year">
	<%dim ii,dbyear,dbmonth,dbday
	dbyear = resumeObj.eBeginDate
	if isdate(dbyear) then 
		dbyear = year(dbyear)
		dbmonth = month(resumeObj.eBeginDate)
		dbday  = day(resumeObj.eBeginDate)
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
	</select>
	
	</td>
  </tr>
  <tr>
    <td align="right" class="hback">结束时间：</td>
    <td class="hback">
	<select name="txt_year1">
	<%dim dbyear1,dbmonth1,dbday1
	dbyear1 = resumeObj.eEndDate
	if isdate(dbyear1) then 
		dbyear1 = year(dbyear1)
		dbmonth1 = month(resumeObj.eEndDate)
		dbday1  = day(resumeObj.eEndDate)
	else
		dbyear1 = 1980
		dbmonth1 = 1
	end if
	for ii= 1960 to 2010
	if cstr(ii)=cstr(dbyear1) then 
		response.Write("<option value="""&ii&""" selected>"&ii&"</option>"&vbcrlf)
	else
		response.Write("<option value="""&ii&""">"&ii&"</option>"&vbcrlf)
	end if	
	next
	%>   
	</select>
	<select name="txt_month1">
	<%
	for ii= 1 to 12
		if cstr(ii)=cstr(dbmonth1) then 
			response.Write("<option value="""&ii&""" selected>"&ii&"</option>"&vbcrlf)
		else
			response.Write("<option value="""&ii&""">"&ii&"</option>"&vbcrlf)
		end if	
	next
	%>
	</select>
	<select name="txt_day1">
	<%
	for ii= 1 to 31
		if cstr(ii)=cstr(dbday1) then 
			response.Write("<option value="""&ii&""" selected>"&ii&"</option>"&vbcrlf)
		else
			response.Write("<option value="""&ii&""">"&ii&"</option>"&vbcrlf)
		end if	
	next
	%>
	</select>

</td></tr>
  <tr>
    <td align="right" class="hback">学校：</td>
    <td class="hback">
	<input name="txt_SchoolName" type="text" id="txt_SchoolName" style="width:60%"
	onfocus="Do.these('txt_SchoolName',function(){return isEmpty('txt_SchoolName','span_school')})"
	onKeyUp="Do.these('txt_SchoolName',function(){return isEmpty('txt_SchoolName','span_school')})" value="<%=resumeObj.SchoolName%>" maxlength="40"
	/>
	<span id="span_school"></span></td>
  </tr>
  <tr>
    <td align="right" class="hback">专业：</td>
    <td class="hback">
	<input name="txt_Specialty" type="text" id="txt_Specialty" style="width:60%" onfocus="this.className='RightInput'"  value="<%=resumeObj.Specialty%>" maxlength="40"/></td>
  </tr>
  <tr>
    <td align="right" class="hback">学历：</td>
    <td class="hback"><input name="txt_Diploma" type="text" id="txt_Diploma" style="width:60%" onfocus="this.className='RightInput'"  value="<%=resumeObj.Diploma%>" maxlength="30"/></td>
  </tr>
  <tr>
    <td align="right" class="hback">专业描述：</td>
    <td class="hback"><textarea name="txt_Description" rows="10" id="txt_Description" style="width:60%" onfocus="this.className='RightInput'"><%=resumeObj.eDescription%></textarea></td>
  </tr>
  <tr>
    <td class="hback">&nbsp;</td>
    <td class="hback">
		<input type="hidden" name="txt_BeginDate" value="" />
	<input type="hidden" name="txt_EndDate" value="" />

	<input type="Button" name="SubmitButton" value="保存/下一步" onclick="txt_Description.value=txt_Description.value.substring(0,3000);txt_BeginDate.value=txt_year.value+'-'+txt_month.value+'-'+txt_day.value ;txt_EndDate.value=txt_year1.value+'-'+txt_month1.value+'-'+txt_day1.value ;ajaxPost('AP_Resume_Action.asp', Form.serialize('EducateExpForm'),'EducateExpForm','<%=action%>','<%=id%>');"/>
	&nbsp;&nbsp;
	<input type="reset" name="resetButton" value="重 设" onClick="javascript:if(confirm('确认清空所有表单输入？')){$('IntentionForm').reset();}else{return false;}" />	</td>
  </tr>
</table>
</form>






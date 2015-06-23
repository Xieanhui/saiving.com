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
if id<>"" then call resumeObj.getResumeInfo("projectexp",id)
Conn.close
Set Conn=nothing
%>
<form name="ProjectExpForm" action="" method="post">
<table width="100%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td width="19%" align="right" class="hback">开始时间：</td>
    <td width="81%" class="hback">
	<select name="txt_year">
	<%dim ii,dbyear,dbmonth,dbday
	dbyear = resumeObj.pBeginDate
	if isdate(dbyear) then 
		dbyear = year(dbyear)
		dbmonth = month(resumeObj.pBeginDate)
		dbday  = day(resumeObj.pBeginDate)
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
	dbyear1 = resumeObj.pEndDate
	if isdate(dbyear1) then 
		dbyear1 = year(dbyear1)
		dbmonth1 = month(resumeObj.pEndDate)
		dbday1  = day(resumeObj.pEndDate)
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
</td>
  </tr>
  <tr>
    <td align="right" class="hback">项目名称：</td>
    <td class="hback">
	<input name="txt_Project" type="text" id="txt_Project" value="<%=resumeObj.Project%>" style="width:60%"
	onfocus="Do.these('txt_Project',function(){return isEmpty('txt_Project','span_project')})"
	onKeyUp="Do.these('txt_Project',function(){return isEmpty('txt_Project','span_project')})"
	/><span id="span_project"></span></td>
  </tr>
  <tr>
    <td align="right" class="hback">软件环境：</td>
    <td class="hback"><input name="txt_SoftSettings" type="text" id="txt_SoftSettings" value="<%=resumeObj.SoftSettings%>" style="width:60%" onfocus="this.className='RightInput'"/></td>
  </tr>
  <tr>
    <td align="right" class="hback">硬件环境：</td>
    <td class="hback"><input name="txt_HardSettings" type="text" id="txt_HardSettings" value="<%=resumeObj.HardSettings%>" style="width:60%" onfocus="this.className='RightInput'"/></td>
  </tr>
  <tr>
    <td align="right" class="hback">开发工具：</td>
    <td class="hback"><input name="txt_Tools" type="text" id="txt_Tools" value="<%=resumeObj.Tools%>" style="width:60%" onfocus="this.className='RightInput'"/></td>
  </tr>
  <tr>
    <td align="right" class="hback">项目描述：</td>
    <td class="hback"><textarea name="txt_ProjectDescript" rows="10" id="txt_ProjectDescript" style="width:60%" 
	onfocus="Do.these('txt_ProjectDescript',function(){return isEmpty('txt_ProjectDescript','span_ProjectDescript')})"
	onKeyUp="Do.these('txt_ProjectDescript',function(){return isEmpty('txt_ProjectDescript','span_ProjectDescript')})"
><%=resumeObj.ProjectDescript%></textarea><span id="span_ProjectDescript"></span></td>
  </tr>
  <tr>
    <td align="right" class="hback">责任描述：</td>
    <td class="hback"><textarea name="txt_Duty" rows="8" id="txt_Duty" style="width:60%" 
	onfocus="Do.these('txt_Duty',function(){return isEmpty('txt_Duty','span_Duty')})"
	onKeyUp="Do.these('txt_Duty',function(){return isEmpty('txt_Duty','span_Duty')})"
	><%=resumeObj.Duty%></textarea><span id="span_Duty"></span></td>
  </tr>
  <tr>
    <td class="hback">&nbsp;</td>
    <td class="hback">
	<input type="hidden" name="txt_BeginDate" value="" />
	<input type="hidden" name="txt_EndDate" value="" />
	<input type="submit" name="SubmitButton" value="保存/下一步"  onclick="txt_BeginDate.value=txt_year.value+'-'+txt_month.value+'-'+txt_day.value ;txt_EndDate.value=txt_year1.value+'-'+txt_month1.value+'-'+txt_day1.value ; ajaxPost('AP_Resume_Action.asp', Form.serialize('ProjectExpForm'),'ProjectExpForm','<%=action%>','<%=id%>');"/>
	&nbsp;&nbsp;
	<input type="reset" name="resetButton" value="重 设" onClick="javascript:if(confirm('确认清空所有表单输入？')){$('IntentionForm').reset();}else{return false;}" />	</td>
  </tr>
</table>
</form>






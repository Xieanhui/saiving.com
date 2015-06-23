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
if id<>"" then call resumeObj.getResumeInfo("trainexp",id)
Conn.close
Set Conn=nothing
%>
<form name="TrainExpForm" action="" method="post">
  <table width="100%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td width="19%" align="right" class="hback">开始时间：</td>
    <td width="81%" class="hback">
	<select name="txt_year">
	<%dim ii,dbyear,dbmonth,dbday
	dbyear = resumeObj.tBeginDate
	if isdate(dbyear) then 
		dbyear = year(dbyear)
		dbmonth = month(resumeObj.tBeginDate)
		dbday  = day(resumeObj.tBeginDate)
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
	dbyear1 = resumeObj.tEndDate
	if isdate(dbyear1) then 
		dbyear1 = year(dbyear1)
		dbmonth1 = month(resumeObj.tEndDate)
		dbday1  = day(resumeObj.tEndDate)
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
    <td align="right" class="hback">培训机构：</td>
    <td class="hback">
	<input name="txt_TrainOrgan" type="text" id="txt_TrainOrgan" value="<%=resumeObj.TrainOrgan%>" style="width:60%"
	onFocus="Do.these('txt_TrainOrgan',function(){return isEmpty('txt_TrainOrgan','span_trainOrgan')})"
	onKeyUp="Do.these('txt_TrainOrgan',function(){return isEmpty('txt_TrainOrgan','span_trainOrgan')})"
	/><span id="span_trainOrgan"></span></td>
  </tr>
  <tr>
    <td align="right" class="hback">培训地点：</td>
    <td class="hback"><input name="txt_TrainAdress" type="text" id="txt_TrainAdress" style="width:60%" onfocus="this.className='RightInput'" value="<%=resumeObj.TrainAdress%>"/></td>
  </tr>
  <tr>
    <td align="right" class="hback">培训内容：</td>
    <td class="hback"><textarea name="txt_TrainContent" rows="10" id="txt_TrainContent" style="width:60%" onfocus="this.className='RightInput'"><%=resumeObj.TrainContent%></textarea></td>
  </tr>
  <tr>
    <td align="right" class="hback">证书：</td>
    <td class="hback"><input name="txt_Certificate" type="text" id="txt_Certificate" style="width:60%" onfocus="this.className='RightInput'" value="<%=resumeObj.tCertificate%>"/></td>
  </tr>
  <tr>
    <td class="hback">&nbsp;</td>
    <td class="hback">
		<input type="hidden" name="txt_BeginDate" value="" />
	<input type="hidden" name="txt_EndDate" value="" />
	<input type="submit" name="SubmitButton" value="保存/下一步" onclick="txt_BeginDate.value=txt_year.value+'-'+txt_month.value+'-'+txt_day.value ;txt_EndDate.value=txt_year1.value+'-'+txt_month1.value+'-'+txt_day1.value ;ajaxPost('AP_Resume_Action.asp', Form.serialize('TrainExpForm'),'TrainExpForm','<%=action%>','<%=id%>');"/>
	&nbsp;
	<input type="reset" name="resetButton" value="重 设" onClick="javascript:if(confirm('确认清空所有表单输入？')){$('IntentionForm').reset();}else{return false;}" />	</td>
  </tr>
</table>
</form>






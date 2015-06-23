<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="cls_resume.asp"-->
<%
Response.Charset="GB2312"
session("resumeStep")="baseinfo"
Dim resumeObj,id,Conn
MF_Default_Conn
id=trim(NoSqlHack(request.QueryString("id")))
Set resumeObj=New cls_resume
if id<>"" then call resumeObj.getResumeInfo("intention",id)
Conn.close
Set Conn=nothing
%>
<form name="IntentionForm" action="" method="post">
	<table width="100%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
		<tr>
			<td width="19%" align="right" class="hback">工作类型：</td>
			<td width="81%" class="hback">
				<select name="sel_WorkType">
					<option value="1" <%if resumeObj.WorkTypee="1" then Response.Write("selected")%>>全 职</option>
					<option value="2" <%if resumeObj.WorkTypee="2" then Response.Write("selected")%>>兼 职</option>
					<option value="3" <%if resumeObj.WorkTypee="3" then Response.Write("selected")%>>实 习</option>
					<option value="4" <%if resumeObj.WorkTypee="4" then Response.Write("selected")%>>全职/兼职</option>
				</select>
			</td>
		</tr>
		<tr>
			<td align="right" class="hback">期望工资：</td>
			<td class="hback">
				<select name="sel_Salary" id="sel_Salary">
					<option value="1"  <%if resumeObj.Salary="1" then Response.Write("selected")%>>1500以下</option>
					<option value="2"  <%if resumeObj.Salary="2" then Response.Write("selected")%>>1500-1999</option>
					<option value="3"  <%if resumeObj.Salary="3" then Response.Write("selected")%>>2000-2999</option>
					<option value="4"  <%if resumeObj.Salary="4" then Response.Write("selected")%>>3000-4499</option>
					<option value="5"  <%if resumeObj.Salary="5" then Response.Write("selected")%>>4500-5999</option>
					<option value="6"  <%if resumeObj.Salary="6" then Response.Write("selected")%>>6000-7999</option>
					<option value="7"  <%if resumeObj.Salary="7" then Response.Write("selected")%>>8000-9999</option>
					<option value="8"  <%if resumeObj.Salary="8" then Response.Write("selected")%>>10000-14999</option>
					<option value="9"  <%if resumeObj.Salary="9" then Response.Write("selected")%>>15000-19999</option>
					<option value="10" <%if resumeObj.Salary="10" then Response.Write("selected")%>>20000-29999</option>
					<option value="11" <%if resumeObj.Salary="11" then Response.Write("selected")%>>30000-49999</option>
					<option value="12" <%if resumeObj.Salary="12" then Response.Write("selected")%>>50000以上</option>
				</select>
			</td>
		</tr>
		<tr>
			<td align="right" class="hback">自我评价：</td>
			<td class="hback">
				<textarea name="txt_SelfAppraise" rows="20" style="width:80%" onfocus="this.className='RightInput'"><%=resumeObj.SelfAppraise%></textarea>
			</td>
		</tr>
		<tr>
			<td class="hback">&nbsp;</td>
			<td class="hback">
				<input type="button" name="SubmitButton" value="保存/下一步" onclick="txt_SelfAppraise.value=txt_SelfAppraise.value.substring(0,3000);ajaxPost('AP_Resume_Action.asp', Form.serialize('IntentionForm'),'IntentionForm','edit','<%=id%>');"/>
				&nbsp;
				<input type="reset" name="resetButton" value="重 设" onClick="javascript:if(confirm('确认清空所有表单输入？')){$('IntentionForm').reset();}else{return false;}" />
			</td>
		</tr>
	</table>
</form>






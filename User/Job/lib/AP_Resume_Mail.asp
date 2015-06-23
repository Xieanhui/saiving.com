<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<%session.CodePage="936"%>
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
if id<>"" then call resumeObj.getResumeInfo("mail",id)
Conn.close
Set Conn=nothing
%>
<form name="MailForm" action="AP_Resume_Action.asp?action=mail&id=" method="post">
  <table width="100%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td width="19%" align="right" class="hback">求职信主题：</td>
    <td width="81%" class="hback">
	<input name="txt_MailName" type="text" id="txt_MailName" style="width:60%" 
	onfocus="Do.these('txt_MailName',function(){return isEmpty('txt_MailName','span_mailName')})"
	onKeyUp="Do.these('txt_MailName',function(){return isEmpty('txt_MailName','span_mailName')})" value="<%=resumeObj.mtitle%>" maxlength="30"
	/>
	<span id="span_mailName"></span></td>
  </tr>
  <tr>
    <td align="right" class="hback">内容：</td>
    <td class="hback"><textarea name="txt_Content" rows="10" style="width:60%"
	onfocus="Do.these('txt_Content',function(){return isEmpty('txt_Content','span_content')})"
	onKeyUp="Do.these('txt_Content',function(){return isEmpty('txt_Content','span_content')})"
	><%=resumeObj.mContent%></textarea><span id="span_content"></span></td>
  </tr>
  <tr>
    <td class="hback">&nbsp;</td>
    <td class="hback">
	<input type="button" name="SubmitButton" value="保存" onclick="txt_Content.value=txt_Content.value.substring(0,3000);ajaxPost('AP_Resume_Action.asp', Form.serialize('MailForm'),'MailForm','<%=action%>','<%=id%>');"/>
	&nbsp;&nbsp;
	<input type="reset" name="resetButton" value="重 设" onClick="javascript:if(confirm('确认清空所有表单输入？')){$('IntentionForm').reset();}else{return false;}" />	</td>
  </tr>
</table>
</form>







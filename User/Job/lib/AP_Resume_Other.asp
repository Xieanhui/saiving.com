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
if id<>"" then call resumeObj.getResumeInfo("other",id)
Conn.close
Set Conn=nothing
%>
<form name="OtherForm" action="" method="post">
  <table width="100%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td width="19%" align="right" class="hback">主题：</td>
    <td width="81%" class="hback">
	<input name="txt_Title" type="text" id="txt_Title" value="<%=resumeObj.title%>" style="width:60%" 
	onfocus="Do.these('txt_Title',function(){return isEmpty('txt_Title','span_title')})"
	onKeyUp="Do.these('txt_Title',function(){return isEmpty('txt_Title','span_title')})"
	/><span id="span_title"></span></td>
  </tr>
  <tr>
    <td align="right" class="hback">内容：</td>
    <td class="hback"><textarea name="txt_Content" rows="10"  style="width:60%"
	onfocus="Do.these('txt_Content',function(){return isEmpty('txt_Content','span_content')})"
	onKeyUp="Do.these('txt_Content',function(){return isEmpty('txt_Content','span_content')})"
	><%=resumeObj.Content%></textarea><span id="span_content"></span></td>
  </tr>
  <tr>
    <td class="hback">&nbsp;</td>
    <td class="hback">
	<input type="button" name="SubmitButton" value="保存/下一步" onclick="txt_Content.value=txt_Content.value.substring(0,3000);ajaxPost('AP_Resume_Action.asp', Form.serialize('OtherForm'),'OtherForm','<%=action%>','<%=id%>');"/>
	&nbsp;&nbsp;
	<input type="reset" name="resetButton" value="重 设" onClick="javascript:if(confirm('确认清空所有表单输入？')){$('IntentionForm').reset();}else{return false;}" />	</td>
  </tr>
</table>
</form>






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
if id<>"" then call resumeObj.getResumeInfo("language",id)
Conn.close
Set Conn=nothing
%>
<form name="LanuageForm" action="AP_Resume_Action.asp?action=language&id=" method="post">
  <table width="100%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td width="19%" align="right" class="hback">语种：</td>
    <td width="81%" class="hback">
	<input name="txt_Language" type="text" id="txt_Language" style="width:60%" 
	onfocus="Do.these('txt_Language',function(){return isEmpty('txt_Language','span_language')})"
	onKeyUp="Do.these('txt_Language',function(){return isEmpty('txt_Language','span_language')})" value="<%=resumeObj.language%>" maxlength="30"
	/>
	<span id="span_language"></span></td>
  </tr>
  <tr>
    <td align="right" class="hback">掌握程度：</td>
    <td class="hback">
	<input name="txt_Degree" type="text" id="txt_Degree" style="width:60%"
	onfocus="Do.these('txt_Degree',function(){return isEmpty('txt_Degree','span_degree')})"
	onKeyUp="Do.these('txt_Degree',function(){return isEmpty('txt_Degree','span_degree')})" value="<%=resumeObj.Degree%>" maxlength="30"
	/>
	<span id="span_degree"></span></td>
  </tr>
  <tr>
    <td class="hback">&nbsp;</td>
    <td class="hback">
	<input type="submit" name="SubmitButton" value="保存/下一步"  onclick="ajaxPost('AP_Resume_Action.asp', Form.serialize('LanuageForm'),'LanuageForm','<%=action%>','<%=id%>');"/>
	&nbsp;&nbsp;
	<input type="reset" name="resetButton" value="重 设" onClick="javascript:if(confirm('确认清空所有表单输入？')){$('IntentionForm').reset();}else{return false;}" />	</td>
  </tr>
</table>
</form>






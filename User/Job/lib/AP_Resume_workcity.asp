<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<%session.CodePage="936"%>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="cls_resume.asp"-->
<%
response.Charset="GB2312"
Dim resumeObj,id,Conn,action
MF_Default_Conn
id=trim(NoSqlHack(request.QueryString("id")))
action=trim(NoSqlHack(request.QueryString("action")))
Set resumeObj=New cls_resume
if id<>"" then call resumeObj.getResumeInfo("workcity",id)

%>
<form name="WorkCityForm" action="AP_Resume_Action.asp?action=workcity" method="post">
  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td width="19%" align="right" class="hback">工作地点：</td>
    <td width="81%" class="hback">
	<%
		dim tradeRs
		set tradeRs=Conn.execute("Select PID,Province From FS_AP_Province")
		response.Write("<select name=""sel_trade"" onChange=""setValue(this,$('hid_province'));getCity('city_container',this.value)"">"&vbcrlf)
		response.Write("<option vaule="""">请选择省</option>"&vbcrlf)
		while not tradeRs.eof
			response.Write("<option value="""&tradeRs("PID")&""">"&tradeRs("Province")&"</option>"&vbcrlf)
			tradeRs.movenext
		wend
		response.Write("</select>"&vbcrlf)
	%>
	<span id="span_province"></span>
	<span id="city_container"></span></td>
  </tr>
  <tr>
    <td class="hback">&nbsp;</td>
    <td class="hback">
	<input type="hidden" name="hid_province" value="" />
	<input type="hidden" name="hid_city"  value=""/>
	<input type="button" name="SubmitButton" value="保存" onclick="ajaxPost('AP_Resume_Action.asp', Form.serialize('WorkCityForm'),'WorkCityForm','<%=action%>','<%=id%>');"/>
  </tr>
</table>
</form>
<%
Conn.close
Set Conn=nothing
%>
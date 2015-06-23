<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<%session.CodePage="936"%>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="cls_resume.asp"-->
<%
Response.Charset="GB2312"
session("resumeStep")="baseinfo"
Dim resumeObj,id,Conn,action
MF_Default_Conn

id=trim(NoSqlHack(request.QueryString("id")))
action=trim(NoSqlHack(request.QueryString("action")))
Set resumeObj=New cls_resume
if id<>"" then call resumeObj.getResumeInfo("workexp",id)


Function Get_FildValue_List(This_Fun_Sql,EquValue,Get_Type)
'''This_Fun_Sql 传入sql语句,EquValue与数据库相同的值如果是<option>则加上selected,Get_Type=1为<option>
Dim Get_Html,This_Fun_Rs,Text
'On Error Resume Next
if instr(This_Fun_Sql,"FS_ME_")>0 then 
	set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
else	
	set This_Fun_Rs = Conn.execute(This_Fun_Sql)
end if	
If Err.Number <> 0 then response.Write("<option value="""">"&Err.Number&":"&Err.Description&"</option>"&vbNewLine):Exit Function
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
		exit do : Get_FildValue_List = "<option value="""">Get_Type值传入错误</option>" : exit Function
	end select
	This_Fun_Rs.movenext
loop
This_Fun_Rs.close
Get_FildValue_List = Get_Html
End Function


%>
<form name="WorkExpForm" action="" method="post">
<table width="100%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td width="19%" align="right" class="hback">开始时间：</td>
    <td width="81%" class="hback">
		<select name="txt_year">
	<%dim ii,dbyear,dbmonth,dbday
	dbyear = resumeObj.wBeginDate
	if isdate(dbyear) then 
		dbyear = year(dbyear)
		dbmonth = month(resumeObj.wBeginDate)
		dbday  = day(resumeObj.wBeginDate)
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
	dbyear1 = resumeObj.wEndDate
	if isdate(dbyear1) then 
		dbyear1 = year(dbyear1)
		dbmonth1 = month(resumeObj.wEndDate)
		dbday1  = day(resumeObj.wEndDate)
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
</td>
  </tr>
  <tr>
    <td align="right" class="hback">公司名称：</td>
    <td class="hback"><input name="txt_CompanyName" type="text" id="txt_CompanyName" style="width:60%" 
	onfocus="Do.these('txt_CompanyName',function(){return isEmpty('txt_CompanyName','companyAlert')})"
	onKeyUp="Do.these('txt_CompanyName',function(){return isEmpty('txt_CompanyName','companyAlert')})" value="<%=resumeObj.CompanyName%>" maxlength="50"/>
    <span id="companyAlert"></span></td>
  </tr>
  <tr>
    <td align="right" class="hback">公司性质：</td>
    <td class="hback"><select name="sel_CompanyKind" id="sel_CompanyKind">
      <option value="1" <%if resumeObj.CompanyKind="1" then Response.Write("selected")%>>外商独资（欧美企业）</option>
      <option value="2" <%if resumeObj.CompanyKind="2" then Response.Write("selected")%>>外商独资（非欧美企业）</option>
      <option value="3" <%if resumeObj.CompanyKind="3" then Response.Write("selected")%>>合资/合作（欧美企业）</option>
      <option value="4" <%if resumeObj.CompanyKind="4" then Response.Write("selected")%>>合资/合作（非欧美企业</option>
      <option value="5" <%if resumeObj.CompanyKind="5" then Response.Write("selected")%>>国营企业（上市公司）</option>
      <option value="6" <%if resumeObj.CompanyKind="6" then Response.Write("selected")%>>国营企业（非上市公司）</option>
      <option value="7" <%if resumeObj.CompanyKind="7" then Response.Write("selected")%>>民营/私营企业/非上市公司</option>
      <option value="8" <%if resumeObj.CompanyKind="8" then Response.Write("selected")%>>外企代表处</option>
      <option value="9" <%if resumeObj.CompanyKind="9" then Response.Write("selected")%>>其它</option>
    </select>
    </td>
  </tr>
  
    <tr  class="hback" id="TradeTR"> 
      <td align="right">行业</td>
      <td width="600">
	  <select name="txt_Trade1" id="txt_Trade1" onChange="setValue(this,$('txt_Trade'));getJob('JobSelect',this.options[this.selectedIndex].value)">
	   <option value="">所有</option>
	   <%=Get_FildValue_List("select TID,Trade from FS_AP_Trade",resumeObj.Trade,1)%>
	  </select>
       职位 <span id="JobSelect">请选择</span>
	   旧的：<%=resumeObj.Trade &" "& resumeObj.Job&" 变更请选择"%>
	   </td>
    </tr>
  
  <tr>
    <td align="right" class="hback">部门：</td>
    <td class="hback"><input name="txt_Department" type="text" id="txt_Department"  value="<%=resumeObj.Department%>" style="width:60%" onfocus="this.className='RightInput'"/></td>
  </tr>
  <tr>
    <td align="right" class="hback">工作描述：</td>
    <td class="hback"><textarea name="txt_Description" rows="10" id="txt_Description" style="width:60%" onfocus="this.className='RightInput'"><%=resumeObj.workDescription%></textarea></td>
  </tr>
  <tr>
    <td align="right" class="hback">证明人：</td>
    <td class="hback">
	<input name="txt_Certifier" type="text" id="txt_Certifier" value="<%=resumeObj.Certifier%>" style="width:60%"
	onfocus="Do.these('txt_Certifier',function(){return isEmpty('txt_Certifier','span_certifier')})"
	onKeyUp="Do.these('txt_Certifier',function(){return isEmpty('txt_Certifier','span_certifier')})"
	/><span id="span_certifier"></span></td>
  </tr>
  <tr>
    <td align="right" class="hback">证明人联系方式：</td>
    <td class="hback">
	<input name="txt_CertifierTel" type="text" id="txt_CertifierTel" value="<%=resumeObj.CertifierTel%>" style="width:60%"
	onfocus="Do.these('txt_CertifierTel',function(){return isEmpty('txt_CertifierTel','span_certifiertel')})"
	onKeyUp="Do.these('txt_CertifierTel',function(){return isEmpty('txt_CertifierTel','span_certifiertel')})"
	/><span id="span_certifiertel"></span>
	</td>
  </tr>
  <tr>
    <td class="hback">&nbsp;</td>
    <td class="hback">
		<input type="hidden" name="txt_BeginDate" value="" />
	<input type="hidden" name="txt_EndDate" value="" />
	<input type="hidden" name="txt_Trade" value="<%=resumeObj.Trade%>" />
	<input type="hidden" name="hid_job" value="" />
	<input type="hidden" name="txt_job"  value="<%=resumeObj.Job%>"/>
	
	
	<input type="submit" name="SubmitButton" value="保存/下一步" onclick="txt_job.value=hid_job.value;txt_BeginDate.value=txt_year.value+'-'+txt_month.value+'-'+txt_day.value ;txt_EndDate.value=txt_year1.value+'-'+txt_month1.value+'-'+txt_day1.value ;ajaxPost('AP_Resume_Action.asp', Form.serialize('WorkExpForm'),'WorkExpForm','<%=action%>','<%=id%>');"/>
	&nbsp;&nbsp;
	<input type="reset" name="resetButton" value="重 设" onClick="javascript:if(confirm('确认清空所有表单输入？')){$('IntentionForm').reset();}else{return false;}" />	</td>
  </tr>
</table>
</form>
<%
Conn.close
Set Conn=nothing
%>






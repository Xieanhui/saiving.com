<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<%session.CodePage="936"%>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="cls_resume.asp"-->
<%
Response.Charset="GB2312"
Dim resumeObj,id,Conn
MF_Default_Conn
id=trim(NoSqlHack(request.QueryString("id")))
Set resumeObj=New cls_resume
if id<>"" then call resumeObj.getResumeInfo("baseinfo",id)
Conn.close
Set Conn=nothing

Dim str_CurrPath
str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
%>
<form action="" method="post" name="BaseInfoForm">
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr>
    <td width="20%" align="right" class="hback">姓名：</td>
    <td class="hback"><input type="text" name="txt_Uname" id="txt_Uname" style="width:60%" onfocus="Do.these('txt_Uname',function(){return isEmpty('txt_Uname','span_Uname')})" onkeyup="Do.these('txt_Uname',function(){return isEmpty('txt_Uname','span_Uname')})" onChange="Do.these('txt_Uname',function(){return isEmpty('txt_Uname','span_Uname')})" value="<%=resumeObj.bs_Uname%>"><span id="span_Uname"></span></td>
  </tr>
  <tr>
    <td align="right" class="hback">性别：</td>
    <td class='hback'><select name="sel_sex">
      <option value="0" <%if resumeObj.bs_sex="0" then Response.Write("selected")%>>先 生</option>
      <option value="1" <%if resumeObj.bs_sex="1" then Response.Write("selected")%>>女 士</option>
    </select>    </td>
  </tr>
  <tr>
    <td align="right" class="hback">照片地址：</td>
    <td class='hback'>
	<input name="sel_PictureExt" type="text" id="sel_PictureExt" style="width:40%" onfocus="this.className='RightInput'" value="<%=resumeObj.bs_PictureExt%>" readonly />
	<input type="button" value="选择照片" onClick="OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.BaseInfoForm.sel_PictureExt);" style="cursor:hand;" />
	
	</td>
  </tr>
  <tr>
    <td align="right" class="hback">出生日期：</td>
    <td class='hback'><select name="txt_year">
	<%dim ii,dbyear,dbmonth
	dbyear = resumeObj.bs_Birthday
	if isdate(dbyear) then 
		dbyear = year(dbyear)
		dbmonth = month(resumeObj.bs_Birthday)
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
	</select></td>
  </tr>
  <tr>
    <td align="right" class="hback">证件类型：</td>
    <td class='hback'><select name="sel_CertificateClass" id="sel_CertificateClass">
      <option value="1" <%if resumeObj.bs_CertificateClass="1" then Response.Write("Selected")%>>身份证</option>
      <option value="2" <%if resumeObj.bs_CertificateClass="2" then Response.Write("Selected")%>>护照</option>
      <option value="3" <%if resumeObj.bs_CertificateClass="3" then Response.Write("Selected")%>>军人证</option>
      <option value="4" <%if resumeObj.bs_CertificateClass="4" then Response.Write("Selected")%>>其他</option>
    </select>     </td>
  </tr>
  <tr>
    <td align="right" class="hback">证件号码：</td>
    <td class='hback'><input name="txt_CertificateNo" type="text" id="txt_CertificateNo" style="width:60%" onfocus="this.className='RightInput'" value="<%=resumeObj.bs_CertificateNo%>" maxlength="20"></td>
  </tr>
  <tr>
    <td align="right" class="hback">目前月薪：</td>
    <td class='hback'>
	<select name="sel_CurrentWage" id="sel_CurrentWage">
      <option value="1"  <%if resumeObj.bs_CurrentWage="1" then response.Write("selected")%>>1500以下</option>
      <option value="2"  <%if resumeObj.bs_CurrentWage="2" then response.Write("selected")%>>1500-1999</option>
      <option value="3"  <%if resumeObj.bs_CurrentWage="3" then response.Write("selected")%>>2000-2999</option>
      <option value="4"  <%if resumeObj.bs_CurrentWage="4" then response.Write("selected")%>>3000-4499</option>
      <option value="5"  <%if resumeObj.bs_CurrentWage="5" then response.Write("selected")%>>4500-5999</option>
      <option value="6"  <%if resumeObj.bs_CurrentWage="6" then response.Write("selected")%>>6000-7999</option>
      <option value="7"  <%if resumeObj.bs_CurrentWage="7" then response.Write("selected")%>>8000-9999</option>
      <option value="8"  <%if resumeObj.bs_CurrentWage="8" then response.Write("selected")%>>10000-14999</option>
      <option value="9"  <%if resumeObj.bs_CurrentWage="9" then response.Write("selected")%>>15000-19999</option>
      <option value="10" <%if resumeObj.bs_CurrentWage="10" then response.Write("selected")%>>20000-29999</option>
      <option value="11" <%if resumeObj.bs_CurrentWage="11" then response.Write("selected")%>>30000-49999</option>
      <option value="12" <%if resumeObj.bs_CurrentWage="12" then response.Write("selected")%>>50000以上</option>
        </select>
      <select name="sel_CurrencyType" id="sel_CurrencyType">
        <option value="1" <%if resumeObj.bs_CurrencyType="1" then Response.Write("selected")%>>人民币</option>
        <option value="2" <%if resumeObj.bs_CurrencyType="2" then Response.Write("selected")%>>港元</option>
        <option value="3" <%if resumeObj.bs_CurrencyType="3" then Response.Write("selected")%>>美元</option>
        <option value="4" <%if resumeObj.bs_CurrencyType="4" then Response.Write("selected")%>>日元</option>
        <option value="5" <%if resumeObj.bs_CurrencyType="5" then Response.Write("selected")%>>欧元</option>
        <option value="6" <%if resumeObj.bs_CurrencyType="6" then Response.Write("selected")%>>其他</option>
      </select>      </td>
  </tr>
  <tr>
    <td align="right" class="hback">工作年限：</td>
    <td class='hback'><select name="sel_WorkAge" id="select">
      <option value="1" <%if resumeObj.bs_WorkAge="1" then Response.Write("Selected")%>>在读学生</option>
      <option value="2" <%if resumeObj.bs_WorkAge="2" then Response.Write("Selected")%>>应届毕业生</option>
      <option value="3" <%if resumeObj.bs_WorkAge="3" then Response.Write("Selected")%>>一年以上</option>
      <option value="4" <%if resumeObj.bs_WorkAge="4" then Response.Write("Selected")%>>两年以上</option>
      <option value="5" <%if resumeObj.bs_WorkAge="5" then Response.Write("Selected")%>>三年以上</option>
      <option value="6" <%if resumeObj.bs_WorkAge="6" then Response.Write("Selected")%>>五年以上</option>
      <option value="7" <%if resumeObj.bs_WorkAge="7" then Response.Write("Selected")%>>八年以上</option>
      <option value="8" <%if resumeObj.bs_WorkAge="8" then Response.Write("Selected")%>>十年以上</option>
        </select></td>
  </tr>
  <tr>
    <td align="right" class="hback">所在省：</td>
    <td class="hback"><input name="txt_Province" type="text" id="txt_Province" style="width:60%" onfocus="this.className='RightInput'" value="<%=resumeObj.bs_Province%>" maxlength="15"></td>
  </tr>
  <tr>
    <td align="right" class="hback">所在城市：</td>
    <td class="hback"><input name="txt_City" type="text" id="txt_City" style="width:60%" onfocus="this.className='RightInput'" value="<%=resumeObj.bs_City%>" maxlength="15"></td>
  </tr>
  <tr>
    <td align="right" class="hback">家庭电话：</td>
    <td class="hback"><input name="txt_HomeTel" type="text" id="txt_HomeTel" style="width:60%" onfocus="this.className='RightInput'" value="<%=resumeObj.bs_HomeTel%>" maxlength="15"></td>
  </tr>
  <tr>
    <td align="right" class="hback">公司电话：</td>
    <td class="hback"><input name="txt_CompanyTel" type="text" id="txt_CompanyTel" style="width:60%" onfocus="this.className='RightInput'" value="<%=resumeObj.bs_CompanyTel%>" maxlength="15"></td>
  </tr>
  <tr>
    <td align="right" class="hback">移动电话：</td>
    <td class="hback"><input name="txt_Mobile" type="text" id="txt_Mobile" style="width:60%" onfocus="this.className='RightInput'" value="<%=resumeObj.bs_Mobile%>" maxlength="15"></td>
  </tr>
  <tr>
    <td align="right" class="hback">E-Mail:</td>
    <td class="hback"><input name="txt_Email" type="text" id="txt_Email" style="width:60%" onfocus="Do.these('txt_Email',function(){return checkMail('txt_Email','span_mail')})" onblur="Do.these('txt_Email',function(){return checkMail('txt_Email','span_mail')})" onkeyup="Do.these('txt_Email',function(){return checkMail('txt_Email','span_mail')})" value="<%=resumeObj.bs_Email%>" maxlength="20">
    <span id="span_mail"></span></td>
  </tr>
  <tr>
    <td align="right" class="hback">OICQ:</td>
    <td class="hback"><input name="txt_QQ" type="text" id="txt_QQ" style="width:60%" onfocus="Do.these('txt_QQ',function(){return isNumber('txt_QQ','span_qq','请检查你的qq号码！','ture')})" onblur="Do.these('txt_QQ',function(){return isNumber('txt_QQ','span_qq','请检查你的qq号码！','ture')})" onkeyup="Do.these('txt_QQ',function(){return isNumber('txt_QQ','span_qq','请检查你的qq号码！','ture')})" value="<%=resumeObj.bs_QQ%>" maxlength="10" >
    <span id="span_qq"></span></td>
  </tr>
  
  <tr>
    <td align="right" class="hback">地址：</td>
    <td class="hback"><input name="txt_address" type="text" id="txt_address" style="width:60%" onfocus="this.className='RightInput'" value="<%=resumeObj.bs_Address%>" maxlength="80"></td>
  </tr>
  <tr>
    <td align="right" class="hback">身高：</td>
    <td class="hback"><input name="txt_ShenGao" type="text" id="txt_ShenGao" style="width:60%" onfocus="Do.these('txt_ShenGao',function(){return isNumber('txt_ShenGao','span_ShenGao','身高必须是数字！','ture')})" onblur="Do.these('txt_ShenGao',function(){return isNumber('txt_ShenGao','span_ShenGao','身高必须是数字！','ture')})" onkeyup="Do.these('txt_ShenGao',function(){return isNumber('txt_ShenGao','span_ShenGao','身高必须是数字！','ture')})" value="<%=resumeObj.bs_ShenGao%>" maxlength="4">CM
    <span id="span_ShenGao"></span></td>
  </tr>
  <tr>
    <td align="right" class="hback">学历：</td>
    <td class="hback"><input name="txt_XueLi" type="text" id="txt_XueLi" style="width:60%" onfocus="this.className='RightInput'" value="<%=resumeObj.bs_XueLi%>" maxlength="30"></td>
  </tr>
  <tr>
    <td align="right" class="hback">多久可以上岗：</td>
    <td class="hback"><input name="txt_HowDay" type="text" id="txt_HowDay" style="width:60%" onfocus="this.className='RightInput'" value="<%=resumeObj.bs_HowDay%>" maxlength="10"></td>
  </tr>
  
  <tr>
    <td align="right" class="hback">是否公开：</td>
    <td class="hback"><select name="sel_isPublic" id="sel_isPublic">
      <option value="0" <%if resumeObj.bs_isPublic="0" then Response.Write("selected")%>>公开</option>
      <option value="1" <%if resumeObj.bs_isPublic="1" then Response.Write("selected")%>>不公开</option>
    </select>
    </td>
  </tr>
  <tr>
    <td align="right" class="hback">&nbsp;</td>
    <td class="hback"><input type="hidden" name="txt_Birthday" value="" />
	<input name="SubmitButton" type="Button" id="SubmitButton" value="保存/下一步" onclick="document.all.txt_Birthday.value=document.all.txt_year.value+'-'+document.all.txt_month.value;ajaxPost('AP_Resume_Action.asp', Form.serialize('BaseInfoForm'),'BaseInfoForm','edit','<%=id%>');">
      <input name="resetButton" type="reset" id="resetButton" value="重置"></td>
  </tr>
</table>
</form>
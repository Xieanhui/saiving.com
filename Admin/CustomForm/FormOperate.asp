<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,CharIndexStr,strShowErr
Dim obj_form_rs,form_sql,userGroup_Sql,obj_userGroup_Rs,VerifyLogin,DataInitStatus
Dim act,formid,formName,tableName,upfileSaveUrl,upfileSize,stateSet,TimeLimited,StartTime,EndTime,SubmitType,GoldFactor,PointFactor,UserGroup,UserOnce,Validate,remark,ArrUserGroup,i
MF_Default_Conn
MF_User_Conn
MF_Session_TF 
act=NoSqlHack(Request.QueryString("act"))
formid=NoSqlHack(Request.QueryString("id"))
if act="edit" then
	if not MF_Check_Pop_TF("MF098") then Err_Show
	form_sql="select formName,tableName,upfileSaveUrl,upfileSize,state,TimeLimited,StartTime,EndTime,SubmitType,GoldFactor,PointFactor,UserGroup,UserOnce,Validate,VerifyLogin,DataInitStatus,remark from FS_MF_CustomForm where id="&formID
	set obj_form_rs=conn.execute(form_sql)
	if obj_form_rs.eof then 
		strShowErr = "<li>操作的数据不正确！</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	formName=obj_form_rs("formName")
	tableName=Replace(obj_form_rs("tableName"),"FS_MF_CustomForm_","")
	upfileSaveUrl=obj_form_rs("upfileSaveUrl")
	upfileSize=obj_form_rs("upfileSize")
	stateSet=obj_form_rs("state")
	TimeLimited=obj_form_rs("TimeLimited")
	StartTime=obj_form_rs("StartTime")
	EndTime=obj_form_rs("EndTime")
	SubmitType=obj_form_rs("SubmitType")
	GoldFactor=obj_form_rs("GoldFactor")
	PointFactor=obj_form_rs("PointFactor")
	UserGroup=obj_form_rs("UserGroup")
	arrUserGroup=split(UserGroup,",")
	UserOnce=obj_form_rs("UserOnce")
	Validate=obj_form_rs("Validate")
	VerifyLogin=obj_form_rs("VerifyLogin")
	DataInitStatus=obj_form_rs("DataInitStatus")
	remark=obj_form_rs("remark")
elseif act="del" then
	if not MF_Check_Pop_TF("MF097") then Err_Show
	form_sql="select formName,tableName from FS_MF_CustomForm where id="&formID
	set obj_form_rs=conn.execute(form_sql)
	if obj_form_rs.eof then 
		obj_form_rs.Close
		Set obj_form_rs = Nothing
		strShowErr = "<li>操作的数据不正确！</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		formName=obj_form_rs("formName")
		tableName=obj_form_rs("tableName")
	end if
	obj_form_rs.Close
	Set obj_form_rs = Nothing
	'删除表
	form_sql="DROP TABLE " & tableName
	on error resume next
	conn.execute(form_sql)
	if err.number <> 0 then err.clear
	'删除表单项数据
	form_sql="delete from FS_MF_CustomForm_Item where formid="&formid
	conn.execute(form_sql)
	'删除表单
	form_sql="delete from FS_MF_CustomForm where id="&formid
	conn.execute(form_sql)
	
	strShowErr = "<li>恭喜，删除自定义表单 "&formName&" 成功!</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../FormManage.asp")
	Response.end
else
	if not MF_Check_Pop_TF("MF099") then Err_Show
	formName=""
	tableName=""
	upfileSaveUrl=""
	upfileSize=""
	stateSet=0
	TimeLimited=1
	StartTime=""
	EndTime=""
	SubmitType=0
	GoldFactor=0
	PointFactor=0
	UserGroup=0
	UserOnce=0
	Validate=0
	VerifyLogin = 0
	DataInitStatus = 1
	remark=""
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自定义表单管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
</head>
<body onLoad='changetm(<%=TimeLimited%>);GetSelect(document.getElementById("SubmitType"));'>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="hback"> 
    <td width="100%" class="xingmu"><a href="#" class="sd"><strong>自定义表单</strong></a><a href="../../help?Lable=NS_Form_Custom" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a>　<a href="FormOperate.asp?act=add"></a>　　　　　　　　　　　　　    </td>
  </tr>
  <tr>
      <td height="18" class="hback"><a href="FormManage.asp">表单管理</a></td>
  </tr>
</table>
  <table  width="98%" border="0" cellspacing="1" cellpadding="5" align="center" class="table">
<form name="form1" method="post" onSubmit="return CheckData(this);" action="FormSave.asp?act=<%=act%>">
	<tr>
      <td width="23%" align="right" class="hback">表单名称：</td>
      <td width="77%" class="hback"><input name="formName" <% if act="edit" then Response.Write("readonly") %> type="text" id="formName" value="<%=formName%>"></td>
    </tr>
    <tr>
      <td align="right" class="hback">表名称：</td>
      <td class="hback">FS_MF_CustomForm_
      <input name="tableName" type="text" id="tableName" <% if act="edit" then Response.Write("readonly") %> value="<%=tableName%>" <%if act="edit" then response.Write("readonly")%>></td>
    </tr>
    <tr>
      <td align="right" class="hback">上传附件保存地址：</td>
      <td class="hback"><input name="upfileSaveUrl" type="text" id="upfileSaveUrl" value="<%=upfileSaveUrl%>" size="40" maxlength="255" readonly>
        <%if act="add" then%>
        <INPUT type="button"  name="Submit4" value="选择路径" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPathFrame.asp?CurrPath=<%= Replace("/" & G_VIRTUAL_ROOT_DIR & "/Userfiles/form","//","/")%>',320,280,window,document.form1.upfileSaveUrl);document.form1.upfileSaveUrl.focus();">
        <%End if%></td>
    </tr>
    <tr>
      <td align="right" class="hback">上传文件大小：</td>
      <td class="hback">最大值
        <input type="text" name="upfileSize" value="<% if upfileSize = "" then Response.Write("1") else Response.Write(upfileSize) %>">
      KB</td>
    </tr>
    <tr>
      <td align="right" class="hback">状态：</td>
      <td class="hback"><input type="radio" name="stateSet" value="0" <%if stateSet=0 then response.Write("checked")%>>
        正常
        <input type="radio" name="stateSet" value="1" <%if stateSet=1 then response.Write("checked")%>>
      锁定</td>
    </tr>
    <tr>
      <td align="right" class="hback">启用时间限制：</td>
      <td class="hback"><INPUT onClick="changetm(0);" type="radio" value="0" name="TimeLimited" <%if TimeLimited=0 then response.Write("checked")%>>启用
        <INPUT onClick="changetm(1);" type="radio" value="1" name="TimeLimited" <%if TimeLimited=1 then response.Write("checked")%>>不启用		</td>
    </tr>
    <TR id="tr_tms">
      <TD align="right" class="hback">开始时间：</TD>
      <TD align="left" class="hback"><INPUT id="StartTime" readOnly name="StartTime" value="<% = StartTime %>">
          <input name="SelectDate" type="button" id="SelectDate" value="选择时间" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.StartTime);" > </TD>
    </TR>
    <TR id="tr_tme">
      <TD align="right" class="hback">结束时间：</TD>
      <TD align="left" class="hback"><INPUT id="EndTime" readOnly name="EndTime" value="<% = EndTime %>">
         <input name="SelectDate" type="button" id="SelectDate" value="选择时间" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.EndTime);" ></TD>
    </TR>
    
    <tr>
      <td align="right" class="hback">提交权限：</td>
      <td class="hback"><SELECT id="SubmitType" onChange="GetSelect(this)" name="SubmitType">
        <OPTION value="0" <%if SubmitType=0 then response.Write("selected")%>>未设置</OPTION>
        <OPTION value="1" <%if SubmitType=1 then response.Write("selected")%>>扣除金币</OPTION>
        <OPTION value="2" <%if SubmitType=2 then response.Write("selected")%>>扣除积分</OPTION>
        <OPTION value="3" <%if SubmitType=3 then response.Write("selected")%>>扣除金币和积分</OPTION>
        <OPTION value="4" <%if SubmitType=4 then response.Write("selected")%>>达到金币</OPTION>
        <OPTION value="5" <%if SubmitType=5 then response.Write("selected")%>>达到积分</OPTION>
        <OPTION value="6" <%if SubmitType=6 then response.Write("selected")%>>达到金币和积分</OPTION>
      </SELECT>
	  <div id="Div_Gold" style="display:inline">
		&nbsp;&nbsp;&nbsp;
		金币：
		<input name="GoldFactor" type="text" value="<%=GoldFactor%>" style="width:58px;" />
	  </div>
	  <div id="Div_Point" style="display:inline">
		&nbsp;&nbsp;&nbsp;&nbsp;
		积分:
		<input name="PointFactor" type="text" value="<%=PointFactor%>" style="width:58px;" />
	  </div>
	  <div id="Div_userGroup">
		请选择会员组:<br />
		<select size="4" name="UserGroup" multiple="multiple" style="height:160px;width:154px;">
			<%
			dim isSelect
			userGroup_sql="select GroupID,GroupName,GroupType From FS_ME_Group order by GroupID desc"
			set obj_userGroup_rs=User_Conn.execute(userGroup_sql)
			do while not obj_userGroup_rs.eof
				if isarray(arrUserGroup) then				
					isSelect=true
					for i=0 to ubound(arrUserGroup)
						if trim(cstr(obj_userGroup_rs(0)))=trim(cstr(arrUserGroup(i))) then
							response.write("<option value="""&obj_userGroup_rs(0)&""" selected>"&obj_userGroup_rs(1)&"</option>")
							isSelect=false
							exit for
						end if
					next
					if isSelect=true then
						response.write("<option value="""&obj_userGroup_rs(0)&""">"&obj_userGroup_rs(1)&"</option>")
					end if
				else
					response.write("<option value="""&obj_userGroup_rs(0)&""">"&obj_userGroup_rs(1)&"</option>")
				end if
			  obj_userGroup_rs.movenext
			loop			  
			%>		
		</select>
		</div></td>
    </tr>
    <tr>
      <td align="right" class="hback">提交次数限制：</td>
      <td class="hback"><INPUT id="UserOnce" type="checkbox" name="UserOnce" value="0" <%if UserOnce=0 then response.Write("checked")%>>
      <LABEL for="ChbOnce">每个用户只能提交一次</LABEL></td>
    </tr>
    <tr>
      <td align="right" class="hback">验证码设置：</td>
      <td class="hback"><INPUT id="Validate" type="checkbox" name="Validate" value="1" <%if Validate=1 then response.Write("checked")%>>
      <LABEL for="ChbShowValidate">显示验证码</LABEL></td>
    </tr>
    <tr>
      <td align="right" class="hback">是否验证登陆：</td>
      <td class="hback"><input name="VerifyLogin" type="checkbox" id="VerifyLogin" value="1" <%if VerifyLogin=1 then response.Write("checked")%>>
        用户是否必须登陆后才能够发布数据</td>
    </tr>
    <tr>
      <td align="right" class="hback">数据初始状态：</td>
      <td class="hback"><label>
        <input name="DataInitStatus" type="checkbox" id="DataInitStatus" value="1"  <%if DataInitStatus=1 then response.Write("checked")%>>
      用户发布的数据，初始状态是否锁定,选中为锁定</label></td>
    </tr>
    <tr>
      <td align="right" class="hback">表单说明：</td>
      <td class="hback"><TEXTAREA name="remark" cols="40" rows="8" id="remark"><%=remark%></TEXTAREA>
(255个字符以内有效)</td>
    </tr>
    <tr>
      <td align="right" class="hback">&nbsp;</td>
      <td class="hback"><input type="hidden" name="formid" value="<%=formid%>">
	  <INPUT type="submit" value=" 确定 " name="BtnOK">
        <INPUT name="reset" type="reset" value=" 重写 "></td>
    </tr>
</form>
  </table>
</body>
</html>
<%
Set Conn = Nothing
Set User_Conn = Nothing
%>
<script language="javascript">
function CheckData(theForm)
{
	if(theForm.formName.value=='')
	{
		alert('请填写表单名！');
		theForm.formName.focus();
		return false;
	}
	if(theForm.tableName.value=='')
	{
		alert('请填写表名！');
		theForm.tableName.focus();
		return false;
	}
	if(theForm.upfileSaveUrl.value=='')
	{
		alert('请选择上传附件保存地址！');
		theForm.upfileSaveUrl.focus();
		return false;
	}
	if(theForm.upfileSize.value=='')
	{
		alert('请填写上传文件大小！');
		theForm.upfileSize.focus();
		return false;
	}
	if (theForm.upfileSize.value!='' && (isNaN(theForm.upfileSize.value) || theForm.upfileSize.value<0))
	{
		alert("上传文件大小应填有效数字！");
		theForm.upfileSize.value="";
		theForm.upfileSize.focus();
		return false;
	}
	return true;
}

function changetm(flag)
{
	var f = 'none';
	if(flag == 0)
	{
		f = '';
	}
	document.getElementById('tr_tms').style.display = f;
	document.getElementById('tr_tme').style.display = f;            
}
function GetSelect(obj)
{
	var selval = parseInt(obj.options[obj.selectedIndex].value);
	var divgroup = document.getElementById("Div_userGroup");
	var divpoint = document.getElementById("Div_Point");
	var divgold = document.getElementById("Div_Gold");
	
	switch(selval)
	{
		case 0:
			divgroup.style.display = "none";
			divpoint.style.display = "none";
			divgold.style.display = "none";
			break;
		case 1:
		case 4:
			divgroup.style.display = "";
			divgold.style.display = "inline";
			divpoint.style.display = "none";
			document.getElementById("userGroup").value = "0";
			break;
		case 2:
		case 5:
			divgroup.style.display = "";
			divpoint.style.display = "inline";
			divgold.style.display = "none";
			document.getElementById("userGroup").value = "0";
			break;
		case 3:
		case 6:
			divgroup.style.display = "";
			divpoint.style.display = "inline";
			divgold.style.display = "inline";
			break;
	}
}
</script>
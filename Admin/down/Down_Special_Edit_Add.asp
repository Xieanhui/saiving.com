<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%
Dim Conn,special_rs,sRootDir,MallDir,str_CurrPath,str_SavePath,Fs_Down,action,specialID,str_Templet,str_DownDir
Dim SpecialEName,SpecialCName,SpecialTemplet,IsUrl,Domain,IsLimited,naviPic,isLock,naviText,FileExtName,Savepath
set Fs_Down = new Cls_News
MF_Default_Conn
'session判断
MF_Session_TF
if not MF_Check_Pop_TF("DS018") then Err_Show
Fs_Down.GetSysParam()
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
if Fs_Down.DownDir<>"" then str_DownDir = "/"+Fs_Down.DownDir else str_DownDir=""
str_SavePath=replace("/"&Fs_Down.DownDir,"//","/")
str_Templet=Replace("/"&G_TEMPLETS_DIR&"/down/Special.htm","//","/")
'Call MF_Check_Pop_TF("NS_Class_000001")
specialID=trim(NoSqLHack(request.QueryString("specialID")))
if specialID<>"" then
	Set special_rs=Conn.execute("select SpecialEName,SpecialCName,SpecialTemplet,IsUrl,Addtime,[Domain],IsLimited,naviPic,isLock,naviText,FileExtName,SavePath from FS_DS_Special where specialID="&CintStr(specialID))
	if not special_rs.eof then
		SpecialEName=special_rs("SpecialEName")
		SpecialCName=special_rs("SpecialCName")
		SpecialTemplet=special_rs("SpecialTemplet")
		IsUrl=special_rs("IsUrl")
		Domain=special_rs("Domain")
		IsLimited=special_rs("IsLimited")
		naviPic=special_rs("naviPic")
		isLock=special_rs("isLock")
		naviText=special_rs("naviText")
		FileExtName=special_rs("FileExtName")
		
	End if
End if

Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")

If G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if
action=request.QueryString("act")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>专区管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<script language="JavaScript" src="../../FS_Inc/prototype.js"></script>
<script language="JavaScript" src="../../FS_Inc/CheckJs.js"></script>
<script language="JavaScript" src="../../FS_Inc/GetLettersByChinese.js"></script>
</head>
 <body>
<% if action="add" then%>
<form name="SpecialForm" method="post" action="Down_Special_Action.asp?act=addaction">
<%else%>
<form name="SpecialForm" method="post" action="Down_Special_Action.asp?act=editaction&specialID=<%=specialID%>">
<%end if%>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr class="hback"> 
      <td class="xingmu">下载系统--专区管理</td>
    </tr>
    <tr> 
      <td height="18" class="hback"><a href="Down_Special_Manage.asp">管理首页</a> 
		| <a href="Down_Special_Edit_Add.asp?act=add">新增专区</a> | <a href="../../help?Lable=DS_Special_Add" target="_blank" style="cursor:help;"><img src="../Images/_help.gif" width="50" height="17" border="0"></a></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	<tr class="hback"> 
	  <td colspan="3" class="xingmu"><%if request.QueryString("Action")="edit" then response.Write("修改专区") else response.Write("添加专区") end if%></td>
	</tr>
	<tr> 
	  <td width="23%" class="hback"><div align="right">专区中文名称：</div></td>
	  <td width="78%" colspan="2" class="hback"><input onBlur="<% if action="add" then %>SetClassEName(this.value,document.SpecialForm.SpecialEName);<% end if %>" name="SpecialCName" type="text" id="SpecialCName"  maxlength="100" style="width:50%" value="<%=SpecialCName%>"> 
		<span class="tx" id="alert_SpecialCName"> *3-100个字符</span></td>
	</tr>
	<tr> 
	  <td width="23%" class="hback"><div align="right">专区英文名称：</div></td>
	  <td width="78%" class="hback"><input name="SpecialEName" type="text" id="SpecialEName"  maxlength="50" <%if Request.QueryString("Act")<>"add" then response.Write("Readonly  disabled")%> style="width:50%" value="<%=SpecialEName%>"><span class="tx"> *</span><button onClick="checkEname()" <%if Request.QueryString("Act")<>"add" then response.Write("disabled")%>>检查英文名是否可用</button>&nbsp;&nbsp;<span style="color:red" id="checkResult"></span>
		<br>
		3-50个字符,必须是字母，数字，中划线，下划线,@,.，一旦确认,将不能修改</span></td>
	</tr>
	<tr> 
	  <td class="hback"><div align="right">二级域名捆绑：</div></td>
	  <td class="hback"><input name="Domain" type="text" id="Domain"  maxlength="150" style="width:60%" value="<%=Domain%>">
	  </td>
	</tr>
	<tr> 
	  <td class="hback"><div align="right">专区导航说明：</div></td>
	  <td class="hback"><textarea name="naviText" cols="35" rows="6" id="naviText" style="width:60%"></textarea> 
	  </td>
	</tr>
	<tr> 
	  <td class="hback"><div align="right">专区保存路径：</div></td>
	  <td class="hback"> <input name="SavePath" type="text" id="SavePath" style="width:60%" value="<%=str_SavePath%>"> <INPUT type="button"  name="Submit4" value="选择路径" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPathFrame.asp?CurrPath=<%=sRootDir & str_DownDir%>',300,250,window,document.SpecialForm.SavePath);document.SpecialForm.SavePath.focus();"> 
		<span class="tx"> *</span><span id="alert_SavePath"></span></td>
	</tr>
	<tr> 
	  <td class="hback"><div align="right">专区模板地址：</div></td>
	  <td class="hback"><input name="SpecialTemplet" type="text" id="SpecialTemplet" maxlength="250" readonly style="width:60%" value="<%if trim(SpecialTemplet)="" then response.Write(str_Templet) else response.Write(SpecialTemplet)%>"> 
		<input type="button" name="Submit" value="选择模板" onClick="OpenWindowAndSetValue('../Commpages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir %>/<% = G_TEMPLETS_DIR %>',400,300,window,document.SpecialForm.SpecialTemplet);document.SpecialForm.SpecialTemplet.focus();"> 
		<span class="tx"> *250个字符</span></td>
	</tr>
	<tr>
	  <td align="right" class="hback">专区导航图片：</td>
	<td class="hback"><input name="NaviPic" type="text" id="NaviPic" value="<%=naviPic%>" style="width:60%">
        <input type="button" name="bnt_Choose"  value="选择图片" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath%>',500,300,window,document.SpecialForm.NaviPic);">	
	</td>
	</tr>
	<tr> 
	  <td class="hback"><div align="right">专区扩展名：</div></td>
	  <td class="hback"> <select name="FileExtName" id="FileExtName">
		  <option value="html"<%if FileExtName="html" then Response.Write("checked")%>>.html</option>
		  <option value="htm" <%if FileExtName="htm" then Response.Write("checked")%>>.htm</option>
		  <option value="shtml" <%if FileExtName="shtml" then Response.Write("checked")%>>.shtml</option>
		  <option value="shtm" <%if FileExtName="shtm" then Response.Write("checked")%>>.shtm</option>
		  <option value="asp" <%if FileExtName="asp" then Response.Write("checked")%>>.asp</option>
		</select> <span class="tx"> *如果需要阅读权限，必须设置为.asp</span></td>
	</tr>
	<tr> 
	  <td height="22" align="right" class="hback">是否锁定：</td>
	  <td height="22" class="hback"><input name="isLock" type="checkbox" id="isLock" value="1" <%if isLock=1 then Response.Write("checked") %>></td>
	</tr>
	<tr> 
	  <td height="22" align="right" class="hback">浏览权限是否限制：</td>
	  <td height="22" class="hback"><input name="IsLimited" type="checkbox" id="IsLimited" value="1" <%if IsLimited="1" then response.Write("checked")%>></td>
	</tr>
	<tr> 
	  <td height="22" align="right" class="hback">是否外部栏目：</td>
	  <td height="22" class="hback"><input name="isUrl" type="checkbox" id="IsUrl" value="1" <%if IsUrl="1" then Response.Write("checked")%>></td>
	</tr>	
	  <td height="21" class="hback"></td>
	<td height="21" class="hback"><input type="button" name="Submit4222" value="保存专区" onClick="checkInput()"> 
	  <input type="reset" name="Submit5222" value="重置" onClick="javascript:if(confirm('是否重置表单？')){	$('SpecialForm').reset()}"> 
	  </td>
	</tr>
  </table>
</form>
</body>
</html>
<script language="JavaScript">
function SetClassEName(Str,Obj)
{
	Obj.value=ConvertToLetters(Str,1);
}
//输入有效性检查
function checkInput()
{
	 var flag1=CheckContentLen('SpecialCName','alert_SpecialCName',"3-100");
	 var flag2=isEmpty('SpecialEName','checkResult');
	 var flag3=isEmpty('SavePath','alert_SavePath');
	 if(flag1&&flag2&&flag3)
	 	$('SpecialForm').submit()
}
//检查英文名是否重复
 function checkEname()
 {
 	if($('SpecialEName').value=="")
	{
		$('checkResult').innerHTML="英文名不能为空";
		$('checkResult').focus();
		return; 
	}
	var param="act=checkename&ename="+$('SpecialEName').value; 
	var jax=new Ajax.Updater('checkResult','Down_Special_Action.asp',{method:'get',parameters:param});
 }
</script>
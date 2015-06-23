<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,tmp_type,strShowErr
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("FL_site") then Err_Show
if not MF_Check_Pop_TF("FL002") then Err_Show

if Request("Edit") ="del" then
	if not MF_Check_Pop_TF("FL002") then Err_Show
	if request("Id")="" then
		strShowErr = "<li>请选择一项</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		Conn.execute("Delete from FS_FL_Class where id in ("& FormatIntArr(request("Id")) &")")
		strShowErr = "<li>删除成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
end if
if Request.Form("Edit_save")<>"" then
	if not MF_Check_Pop_TF("FL002") then Err_Show
	dim obj_fl_Rs_1,SQL_1
	if Len(Request.Form("F_Content"))>200 then
		strShowErr = "<li>说明不能超过200个字符</li>"
		Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	Set obj_fl_Rs_1 = server.CreateObject(G_FS_RS)
	if trim(Request.Form("Edit_save"))="add" then
		Dim CheckRs
		Set CheckRs = Conn.ExeCute("Select F_ClassCName,F_ClassEName from FS_FL_Class Where F_ClassCName = '" & NoSqlHack(Request.Form("F_ClassCName")) & "' Or F_ClassEName = '" & NoSqlHack(Request.Form("F_ClassEName")) & "'")
		If Not CheckRs.Eof Then
			strShowErr = "<li><font color=red>分类中文名或者英文名重复</font></li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
		CheckRs.Close : Set CheckRs = Nothing
		SQL_1= "Select  ID,F_ClassCName,F_ClassEName,F_Content,ParentID  from FS_FL_Class"
		obj_fl_Rs_1.Open SQL_1,Conn,1,3
		obj_fl_Rs_1.addnew
	else
		SQL_1= "Select  ID,F_ClassCName,F_ClassEName,F_Content,ParentID  from FS_FL_Class where id="&CintStr(Request.Form("id"))&""
		obj_fl_Rs_1.Open SQL_1,Conn,1,3
	end if
	obj_fl_Rs_1("F_ClassCName") = NoSqlHack(Request.Form("F_ClassCName"))
	obj_fl_Rs_1("F_ClassEName") = NoSqlHack(Request.Form("F_ClassEName"))
	obj_fl_Rs_1("F_Content") = NoSqlHack(Request.Form("F_Content"))
	obj_fl_Rs_1("ParentID") =0
	obj_fl_Rs_1.update
	obj_fl_Rs_1.close:set obj_fl_Rs_1 = nothing
	strShowErr = "<li>操作成功</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
end if 
%>
<html>
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/GetLettersByChinese.js"></script>
<BODY>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr> 
    <td align="left" colspan="2" class="xingmu">友情连接管理</td>
  </tr>
  <tr> 
    <td align="left" colspan="2" class="hback"><a href="Flink_Manage.asp">管理首页</a>┆<a href="Flink_Edit.asp?Action=Add">添加连接</a>┆<a href="Flik_Class.asp?Action=Add&Edit=true">添加分类</a>┆<a href="Flink_Manage.asp?Type=0">图片连接</a>┆<a href="Flink_Manage.asp?Type=1">文字连接</a>┆<a href="Flink_Manage.asp?Lock=1">已锁定</a>┆<a href="Flink_Manage.asp?Lock=0">未锁定</a></td>
  </tr>
</table>
  
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="ListForms" id="ListForms" method="post" action="">
    <tr> 
      <td width="26%" class="xingmu"><div align="center">分类名称</div></td>
      <td width="26%" class="xingmu"><div align="center">操作</div></td>
    </tr>
    <%
	dim obj_fl_Rs,SQL
	Set obj_fl_Rs = server.CreateObject(G_FS_RS)
	SQL = "Select  ID,F_ClassCName,F_ClassEName,F_Content,ParentID  from FS_FL_Class where ParentID=0 Order by ID desc"
	obj_fl_Rs.Open SQL,Conn,1,3
	do while not obj_fl_Rs.eof 
	%>
    <tr> 
      <td class="hback"><a href="Flik_Class.asp?Edit=true&id=<% = obj_fl_Rs("ID") %>"> 
        <% = obj_fl_Rs("F_ClassCName") %>
        </a></td>
      <td class="hback"><div align="center"><a href="Flik_Class.asp?id=<% = obj_fl_Rs("id") %>&Edit=true">修改</a>┆<a href="Flik_Class.asp?id=<% = obj_fl_Rs("id") %>&Edit=del" onClick="{if(confirm('确定清除您所选择的记录吗？')){return true;}return false;}">删除</a> 
          <input name="Id" type="checkbox" id="Id" value="<% = obj_fl_Rs("id") %>">
        </div></td>
    </tr>
    <%
			obj_fl_Rs.movenext
		Loop
	 %>
    <tr> 
      <td colspan="2" class="hback"><div align="right"> 
          <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form);">
          选中/取消所有 
          <input name="Edit" type="hidden" id="Edit">
          <input type="button" name="Submit" value="删除"  onClick="document.ListForms.Edit.value='del';{if(confirm('确定清除您所选择的记录吗？')){this.document.ListForms.submit();return true;}return false;}">
        </div></td>
    </tr>
    <tr> 
      <td colspan="2" class="hback"> </td>
    </tr>
  </form>
</table>
<%if Request.QueryString("Edit")="true" then%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="myForm" method="post" action="" onSubmit="javascript:return checkdata();">
    <tr> 
      <td colspan="2" class="xingmu"> 
	  <%
	  dim tmp_edit,tmp_id,edit_rs,tmp_ClassCName,tmp_ClassEName,tmp_F_Content,tmp_s
	  if Request.QueryString("Action")="Add" then
		  response.Write("添加分类")
		  tmp_edit="add"
		  tmp_id = ""
		  tmp_ClassCName = ""
		  tmp_ClassEName = ""
		  tmp_F_Content = ""
		  tmp_s = "添加"
	 else
		  response.Write("修改分类")
		  set edit_rs = conn.execute("select F_ClassCName,F_ClassEName,id,F_Content From FS_FL_Class where id="&clng(Request.QueryString("id")))
		  tmp_edit="edit"
		  tmp_id = NoSqlHack(Request.QueryString("id"))
		  tmp_ClassCName = NoSqlHack(edit_rs("F_ClassCName"))
		  tmp_ClassEName = NoSqlHack(edit_rs("F_ClassEName"))
		  tmp_F_Content = NoSqlHack(edit_rs("F_Content"))
		  tmp_s = "修改"
	 end if
	  %>
	  </td>
    </tr>
    <tr> 
      <td width="14%" class="hback"><div align="right">分类名称</div></td>
      <td width="86%" class="hback"> <input onBlur="SetClassEName(this.value,document.myForm.F_ClassEName);" name="F_ClassCName" type="text" value="<% = tmp_ClassCName %>"> </td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">英文名称</div></td>
      <td class="hback"> <input name="F_ClassEName" type="text" value="<% = tmp_ClassEName %>"> </td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">分类说明</div></td>
      <td class="hback"><div align="left"> 
          <textarea name="F_Content" cols="60" rows="6"><% = tmp_F_Content %></textarea>
          200个字符</div></td>
    </tr>
    <tr> 
      <td class="hback"> <div align="center"> </div></td>
      <td class="hback"><input type="submit" name="Submit2" value="<% = tmp_s %>" > <input type="reset" name="Submit3" value="重置">
        <input name="Edit_save" type="hidden" id="Edit_save" value="<% = tmp_edit %>">
        <input name="id" type="hidden" id="id" value="<% = tmp_id %>"></td>
    </tr>
    <tr> 
      <td colspan="2" class="hback"> </td>
    </tr>
  </form>
</table>
<%end if%>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = ListForms.elements[i];  
    if (e.name != 'chkall')  
       e.checked = ListForms.chkall.checked;  
    }  
}
function checkdata()
{
	if (f_trim(document.myForm.F_ClassCName.value)=='')
	{
		alert('分类名称不能为空');
		document.myForm.F_ClassCName.focus();
		return false;
	}
	if (f_trim(document.myForm.F_ClassEName.value)=='')
	{
		alert('英文名称不能为空');
		document.myForm.F_ClassEName.focus();
		return false;
	}
	return true;
}
//去掉字串左边的空格 
function lTrim(str) 
{ 
	if (str.charAt(0) == " ") 
	{ 
		//如果字串左边第一个字符为空格 
		str = str.slice(1);//将空格从字串中去掉 
		//这一句也可改成 str = str.substring(1, str.length); 
		str = lTrim(str); //递归调用 
	} 
	return str; 
} 

//去掉字串右边的空格 
function rTrim(str) 
{ 
	var iLength; 

	iLength = str.length; 
	if (str.charAt(iLength - 1) == " ") 
	{ 
		//如果字串右边第一个字符为空格 
		str = str.slice(0, iLength - 1);//将空格从字串中去掉 
		//这一句也可改成 str = str.substring(0, iLength - 1); 
		str = rTrim(str); //递归调用 
	} 
	return str; 
} 
//去除左右空格
/*
返回值:去除后的值
参数说明:_str,原值
*/
function f_trim(_str)
{
	return lTrim(rTrim(_str)); 
}
function SetClassEName(Str,Obj)
{
	Obj.value=ConvertToLetters(Str,1);
}
</script>
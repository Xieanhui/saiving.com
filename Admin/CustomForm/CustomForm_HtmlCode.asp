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
Dim Conn
Dim CharIndexStr,strShowErr
Dim obj_form_rs,form_sql,userGroup_Sql,obj_userGroup_Rs
dim selectItemArr,i,jsShow
dim id,formName,tableName,ShowStr
MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("MF096") then Err_Show

id=NoSqlHack(request.QueryString("id"))
ShowStr=""

form_sql="select id,formName,tableName from FS_MF_CustomForm where state=0 and id="&id
set obj_form_rs=conn.execute(form_sql)
if obj_form_rs.eof then 
	strShowErr = "<li>操作的数据不正确！</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
else
	jsShow="<script language=""javascript"" charset=""utf-8"" type=""text/javascript"" src="""&replace("/"&G_VIRTUAL_ROOT_DIR&"/customform/CustomFormJS.asp?CustomFormId="&id&"","//","/")&"""></script>"
	formName=obj_form_rs(1)
	tableName=obj_form_rs(2)
end if
form_sql="select formitemid,ItemName,FieldName,IsNull,ItemType,MaxSize,DefaultValue,SelectItem,Remark from FS_MF_CustomForm_Item where formid="&id&" and State=0 order by orderby"
set obj_form_rs=conn.execute(form_sql)
ShowStr=ShowStr&"<table width=""98%"" align=""center"">"&vbcrlf
ShowStr=ShowStr&"<form name="""&formName&""" method=""post"">"&vbcrlf
ShowStr=ShowStr&"<input type=""hidden"" name=""id"" value="""&id&""" >"&vbcrlf
ShowStr=ShowStr&"<input type=""hidden"" name=""submitdata"" value=""1"" >"&vbcrlf
do while not obj_form_rs.eof
	selectItemArr=split(obj_form_rs("SelectItem"),Chr(13)&Chr(10))
	ShowStr=ShowStr&"  <tr>"&vbcrlf
	ShowStr=ShowStr&"    <td width=""25%"" align=""right"">"&obj_form_rs("ItemName")&"</td>"&vbcrlf
	ShowStr=ShowStr&"    <td width=""75%"">"&vbcrlf
	select case obj_form_rs("ItemType")
		case "SingleLineText" '单行文本
			ShowStr=ShowStr&"<input type=""text"" name="""&obj_form_rs("FieldName")&""" value="""&obj_form_rs("DefaultValue")&""""
			if obj_form_rs("MaxSize")<>0 and cstr(obj_form_rs("MaxSize"))<>"" then
				ShowStr=ShowStr&" maxsize="""&obj_form_rs("MaxSize")&""""
			end if
			ShowStr=ShowStr&">"&obj_form_rs("Remark")&vbcrlf
			
		case "MultiLineText" '多行文本
			ShowStr=ShowStr&"<textarea name="""&obj_form_rs("FieldName")&""" cols=""40"" rows=""8"">"&obj_form_rs("DefaultValue")&"</textarea>"&obj_form_rs("Remark")&vbcrlf
			
		case "PassWordText" '密码
			ShowStr=ShowStr&"<input type=""password"" name="""&obj_form_rs("FieldName")&""" value="""&obj_form_rs("DefaultValue")&""""
			if obj_form_rs("MaxSize")<>0 and cstr(obj_form_rs("MaxSize"))<>"" then
				ShowStr=ShowStr&" maxsize="""&obj_form_rs("MaxSize")&""""
			end if
			ShowStr=ShowStr&">"&obj_form_rs("Remark")&vbcrlf
			
		case "DateTime" '日期时间
				ShowStr=ShowStr&"<input type=""text"" name="""&obj_form_rs("FieldName")&""" readOnly value="""&obj_form_rs("DefaultValue")&"""> <input name="""&obj_form_rs("FieldName")&"btn"" type=""button"" value=""选择时间"" onClick=""OpenWindowAndSetValue('"&replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_ADMIN_DIR&"/CommPages/SelectDate.asp","//","/")&"',300,130,window,document.all."&obj_form_rs("FieldName")&");"" >"&obj_form_rs("Remark")&vbcrlf
							
		case "RadioBox" '单选项
			if isarray(selectItemArr) then 
				if obj_form_rs("Remark")<>"" then
					ShowStr=ShowStr&obj_form_rs("Remark")&"<br>"&vbcrlf
				end if
			end if
			for i=0 to ubound(selectItemArr)
				if trim(selectItemArr(i))=obj_form_rs("DefaultValue") then
					ShowStr=ShowStr&"<input type=""radio"" name="""&obj_form_rs("FieldName")&""" value="""&selectItemArr(i)&""" checked>"&selectItemArr(i)&vbcrlf
				else
					ShowStr=ShowStr&"<input type=""radio"" name="""&obj_form_rs("FieldName")&""" value="""&selectItemArr(i)&""" >"&selectItemArr(i)&vbcrlf
				end if
			next
			
		case "CheckBox" '多选项
			if isarray(selectItemArr) then 
				if obj_form_rs("Remark")<>"" then
					ShowStr=ShowStr&obj_form_rs("Remark")&"<br>"&vbcrlf
				end if
			end if
			for i=0 to ubound(selectItemArr)
				if trim(selectItemArr(i))=obj_form_rs("DefaultValue") then
					ShowStr=ShowStr&"<input type=""checkbox"" name="""&obj_form_rs("FieldName")&""" value="""&selectItemArr(i)&""" checked>"&selectItemArr(i)&vbcrlf
				else
					ShowStr=ShowStr&"<input type=""checkbox"" name="""&obj_form_rs("FieldName")&""" value="""&selectItemArr(i)&""" >"&selectItemArr(i)&vbcrlf
				end if
			next
				
		case "Numberic" '数字
			ShowStr=ShowStr&"<input type=""text"" name="""&obj_form_rs("FieldName")&""" value="""&obj_form_rs("DefaultValue")&""""
			if obj_form_rs("MaxSize")<>0 and cstr(obj_form_rs("MaxSize"))<>"" then
				ShowStr=ShowStr&" maxsize="""&obj_form_rs("MaxSize")&""""
			end if
			ShowStr=ShowStr&" onKeyUp=""value=value.replace(/[^\d]/g,'') "">"&obj_form_rs("Remark")&vbcrlf
			
		case "UploadFile" '附件
			ShowStr=ShowStr&"<input type=""file"" name="""&obj_form_rs("FieldName")&""" value="""&obj_form_rs("DefaultValue")&""">"&obj_form_rs("Remark")&vbcrlf
		
		case "DropList" '下拉框
			if isarray(selectItemArr) then 
				if obj_form_rs("Remark")<>"" then
					ShowStr=ShowStr&obj_form_rs("Remark")&"<br>"&vbcrlf
				end if
				ShowStr=ShowStr&"<select name="""&obj_form_rs("FieldName")&""">"&vbcrlf
			end if
			for i=0 to ubound(selectItemArr)
				if trim(selectItemArr(i))=obj_form_rs("DefaultValue") then
					ShowStr=ShowStr&"<option  value="""&selectItemArr(i)&""" checked>"&selectItemArr(i)&"</option>"&vbcrlf
				else
					ShowStr=ShowStr&"<option  value="""&selectItemArr(i)&""" >"&selectItemArr(i)&"</option>"&vbcrlf
				end if
			next
			if isarray(selectItemArr) then 
				ShowStr=ShowStr&"</select>"&vbcrlf
			end if
		
		case "List" '列表框
			if isarray(selectItemArr) then 
				if obj_form_rs("Remark")<>"" then
					ShowStr=ShowStr&obj_form_rs("Remark")&"<br>"&vbcrlf
				end if
				ShowStr=ShowStr&"<select name="""&obj_form_rs("FieldName")&""" size=""8"" style=""height:150px"">"&vbcrlf
			end if
			for i=0 to ubound(selectItemArr)
				if trim(selectItemArr(i))=obj_form_rs("DefaultValue") then
					ShowStr=ShowStr&"<option  value="""&selectItemArr(i)&""" checked>"&selectItemArr(i)&"</option>"&vbcrlf
				else
					ShowStr=ShowStr&"<option  value="""&selectItemArr(i)&""" >"&selectItemArr(i)&"</option>"&vbcrlf
				end if
			next
			if isarray(selectItemArr) then 
				ShowStr=ShowStr&"</select>"&vbcrlf
			end if
	end select
	ShowStr=ShowStr&"  </td>"&vbcrlf
	ShowStr=ShowStr&"</tr>"&vbcrlf
	obj_form_rs.movenext
loop
ShowStr=ShowStr&"<tr>"&vbcrlf
ShowStr=ShowStr&"  <td colspan=""2"" align=""center"">"&vbcrlf
ShowStr=ShowStr&"  <input type=""submit"" onclick=""alert('管理员后台不能够添加数据！');return false;"" value=""提交表单"">"&vbcrlf
ShowStr=ShowStr&"  </td>"&vbcrlf
ShowStr=ShowStr&"</tr>"&vbcrlf
ShowStr=ShowStr&"</form>"&vbcrlf
ShowStr=ShowStr&"</table>"&vbcrlf
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自定义表单管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu">
    <td class="xingmu"><a href="#" class="sd"><strong>预览表单</strong></a><a href="../../help?Lable=NS_Form_Custom" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr class="hback">
    <td class="hback"><a href="FormManage.asp">表单管理</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr id="tr_Show">
    <td height="18" class="hback"><div align="center">
	<%=ShowStr%>
	</div></td>
  </tr>
</table>
</body>
</html>
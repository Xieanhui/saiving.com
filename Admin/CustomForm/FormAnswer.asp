<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Sql,conn,FormData,formid,action,obj_form_rs,tableName,ExeResult,DataField,strShowErr,form_answer
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("MF087") then Err_Show

FormData=NoSqlHack(Request("FormData"))
formid=NoSqlHack(Request("formid"))
if FormData = "" OR formid = "" then	
	strShowErr = "<li>�Ƿ����ݴ���</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
end if
sql="select tableName from FS_MF_CustomForm where id="&formid
set obj_form_rs=conn.execute(sql)
if obj_form_rs.eof then 
	obj_form_rs.Close
	Set obj_form_rs = Nothing
	strShowErr = "<li>�����ı�����ȷ��</li>"
    Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
   	Response.end
else
	tableName=obj_form_rs(0)
end if
obj_form_rs.Close
Set obj_form_rs = Nothing

action = NoSqlHack(Request("action"))
form_answer = NoSqlHack(Request("form_answer"))
if action = "1" then
	Sql = "Update " & tableName & " Set form_answer='" & form_answer & "' Where ID=" & FormData
	Conn.Execute(Sql)
	strShowErr = "<li>���ظ��ɹ�!</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Form_Data.asp?formid="&formid)
	Response.end
end if
sql="select * from " & tableName & " Where ID=" & FormData
Set ExeResult = Server.CreateObject(G_FS_RS)
ExeResult.Open Sql,Conn,1,1
if ExeResult.eof then 
	ExeResult.Close
	Set ExeResult = Nothing
    Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode("<li>�����ı����ݲ����ڣ�</li>")&"&ErrorUrl=")
   	Response.end
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ظ�</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<body topmargin="2" leftmargin="2">
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu">
    <td class="xingmu"><a href="#" class="sd"><strong>���ظ�</strong></a><a href="../../help?Lable=NS_Form_Custom" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr class="hback">
    <td class="hback"><a href="FormManage.asp">������</a></td>
  </tr>
</table>
<table align="center" width="98%" border="0" cellpadding="3" cellspacing="1" class="table">
  <form name="DataList" method="post" onSubmit="return CheckData(this);" action="?action=1">
  <%
	Dim f_Dict
	Set f_Dict = Server.CreateObject(G_FS_DICT)
	f_Dict.Add "id","���"
	f_Dict.Add "form_usernum","�û���"
	f_Dict.Add "form_username","�û���"
	f_Dict.Add "form_ip","IP��ַ"
	f_Dict.Add "form_time","���ʱ��"
	f_Dict.Add "form_lock","����"
	f_Dict.Add "form_answer","�ظ�����"
  For Each DataField In ExeResult.Fields
  %>
  <tr>
    <td class="hback" width="10%" align="center" height="30"><%
	if f_Dict.Exists(DataField.Name) then
		Response.Write(f_Dict.Item(DataField.Name))
	else
		Response.Write(DataField.Name)
	end if	
	%></td>
    <td class="hback">
	<% if DataField.Name = "form_answer" then %>
	<textarea name="form_answer" style="width:98%;" rows="6"><% = DataField.Value %></textarea>
	<% elseif DataField.Name = "form_lock" then %>
	<% if DataField.Value = 1 then Response.Write("��") else Response.Write("��") %>
	<% else %>
		<% = DataField.Value %>
	<% end if %>
	</td>
  </tr>
  <%
  Next
  %>
  <tr>
    <td height="30" colspan="2" align="center" class="hback">
      <input type="submit" name="Submit" value=" �� �� "><input type="hidden" name="FormData" value="<% = FormData %>">
      <input type="hidden" name="formid" value="<% = formid %>">&nbsp;&nbsp;
      <input type="reset" name="Submit2" value=" �� �� ">&nbsp;&nbsp;
      <input type="button" onClick="history.back();" name="Submit3" value=" �� �� ">    </td>
  </tr>
  </form>
  </table>
</body>
</html>
<%
Set Conn = Nothing
ExeResult.Close
Set ExeResult = Nothing
f_Dict.RemoveAll
Set f_Dict = Nothing
%>
<script language="javascript">
function CheckData(theForm){
	if(theForm.form_answer.value==''){
		alert('����д�ظ������ݣ�');
		theForm.form_answer.focus();
		return false;
	}
	return true;
}
</script>
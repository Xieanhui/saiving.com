<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
'session�ж�
MF_Session_TF 
'Ȩ���ж�
'Call MF_Check_Pop_TF("NS_Class_000001")
'�õ���Ա���б�
dim Fs_news
set Fs_news = new Cls_News
Fs_News.GetSysParam()
Dim str_StyleName,txt_Content,dmt_time,strShowErr,Lableclass_SQL,obj_Lableclass_rs
str_StyleName = NoSqlHack(Request.Form("StyleName"))
txt_Content = NoSqlHack(Request.Form("TxtFileds"))
if Request.Form("Action") = "add_save" then
		if str_StyleName ="" or txt_Content ="" then
			strShowErr = "<li>���ж��Ǳ�����д��</li><li>��������д</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		Lableclass_SQL = "Select ClassName,ClassContent from FS_NS_LableClass where ClassName ='"& NoSqlHack(str_StyleName) &"'"
		Set obj_Lableclass_rs = server.CreateObject(G_FS_RS)
		obj_Lableclass_rs.Open Lableclass_SQL,Conn,1,3
		if obj_Lableclass_rs.eof then
			obj_Lableclass_rs.addnew
			obj_Lableclass_rs("ClassName") = str_StyleName
			obj_Lableclass_rs("ClassContent") = txt_Content
			obj_Lableclass_rs.update
		else
			strShowErr = "<li>�˷��������ظ�,����������</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		obj_Lableclass_rs.close:set obj_Lableclass_rs =nothing
		strShowErr = "<li>������ӳɹ�</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Lable_Dir_add.asp")
		Response.end
Elseif Request.Form("Action") = "edit_save" then
		Lableclass_SQL = "Select ClassName,ClassContent from FS_NS_LableClass where id ="& CintStr(Request.Form("ID")) 
		Set obj_Lableclass_rs = server.CreateObject(G_FS_RS)
		obj_Lableclass_rs.Open Lableclass_SQL,Conn,1,3
		if not obj_Lableclass_rs.eof then
			obj_Lableclass_rs("ClassName") = str_StyleName
			obj_Lableclass_rs("ClassContent") = txt_Content
			'obj_Lableclass_rs("AddDate") =now
			obj_Lableclass_rs.update
		End if
		obj_Lableclass_rs.close:set obj_Lableclass_rs =nothing
		strShowErr = "<li>�����޸ĳɹ�</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Lable_Dir_add.asp")
		Response.end
End if
if Request.QueryString("Action") = "del" then
	if Request.QueryString("id") = "" or isnumeric(Request.QueryString("id"))=false then
		strShowErr = "<li>����Ĳ���</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Lable_Dir_add.asp")
		Response.end
	Else
		Conn.execute("Delete from FS_NS_LableClass where id ="&CintStr(Request.QueryString("id")))
		Conn.execute("Update  FS_NS_Lable set LabeClassID=0 where LabeClassID="&CintStr(Request.QueryString("id")))
	End if
		strShowErr = "<li>ɾ������ɹ�</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Lable_Dir_add.asp")
		Response.end
End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ǩ����___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu">��ǩ����<a href="../../help?Lable=NS_Lable_Manage" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr> 
    <td height="18" class="hback"><div align="left"><a href="Lable_Manage.asp">������ҳ</a> 
        &nbsp;|&nbsp; <a href="Lable_Create.asp?ClassID=<%=Request.QueryString("ClassID")%>">������ǩ</a> 
        &nbsp;|&nbsp; <a href="Lable_Dir_add.asp#Add">��ӱ�ǩ��Ŀ</a> &nbsp;|&nbsp; 
        <a href="Lable_style.asp">��ǩ��ʽ����</a> &nbsp;|&nbsp; <a href="Lable_Manage.asp?Action=del_lable"  onClick="{if(confirm('ȷ��ɾ����!')){return true;}return false;}">ɾ�����б�ǩ</a> 
        &nbsp;|&nbsp; <a href="Lable_Manage.asp?Action=del_lable_dir"  onClick="{if(confirm('ȷ��ɾ����!')){return true;}return false;}">ɾ�����б�ǩĿ¼</a> 
        | <a href="../../help?Lable=NS_Lable_Manage_1" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></div></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="hback"> 
    <td width="30%" class="xingmu"> <div align="center">�������� </div></td>
    <td width="44%" class="xingmu"><div align="center">˵��</div></td>
    <td width="26%" class="xingmu"><div align="center">����</div></td>
  </tr>
  <%
		Dim list_SQL,obj_List_rs
		list_SQL = "Select top 50 id,ClassName,ClassContent,ParentID from FS_NS_LableClass Order by Id desc"
		Set obj_List_rs = server.CreateObject(G_FS_RS)
		obj_List_rs.Open list_SQL,Conn,1,3
		do while not obj_List_rs.eof 
	%>
  <tr class="hback"> 
    <td> �� <a href="Lable_Dir_add.asp?id=<% = obj_List_rs("id")%>&Action=edit#add"><% = obj_List_rs("ClassName")%></a></td>
    <td><% = obj_List_rs("ClassContent")%></td>
    <td><div align="center"><a href="Lable_Dir_add.asp?id=<% = obj_List_rs("id")%>&Action=edit#add">�޸�</a>��<a href="Lable_Dir_add.asp?id=<% = obj_List_rs("id")%>&Action=del" onClick="{if(confirm('ȷ��ɾ����!!')){return true;}return false;}">ɾ��</a></div></td>
  </tr>
  <%
	  obj_List_rs.movenext
  Loop
  obj_List_rs.close
  set  obj_List_rs = nothing
  %>
</table>
<%
if Request.QueryString("Action")="edit" then
	Dim tmp_obj,str_ClassName_e,str_Content_e,str_add,str_id
	set tmp_obj = Conn.execute("select id,ClassName,ClassContent from FS_NS_LableClass where id="&CintStr(Request.QueryString("id")))
	if Not tmp_obj.eof then
		str_ClassName_e = tmp_obj("ClassName")
		str_Content_e = tmp_obj("ClassContent")
		str_id = tmp_obj("id")
	End if
	str_add = "edit_save"
Else
	str_add = "add_save"
End if
%>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="xingmu"> 
    <td colspan="2" class="xingmu"><a name="Add" id="Add"></a>������ǩ���� ���������50������</td>
  </tr>
  <form name="Lable_Form" method="post" action="">
    <tr class="hback"> 
      <td width="13%"> <div align="right"> Ŀ¼����</div></td>
      <td width="87%"><input name="StyleName" type="text" id="StyleName" value="<% = str_ClassName_e %>" size="40"> 
        <input name="id" type="hidden" id="id" value="<% = str_id %>"></td>
    </tr>
    <tr class="hback"> 
      <td><div align="right">˵������</div></td>
      <td><textarea name="TxtFileds" rows="15" id="TxtFileds" style="width:90%"><% = str_Content_e %></textarea></td>
    </tr>
    <tr class="hback"> 
      <td>&nbsp;</td>
      <td><input type="submit" name="Submit" value="������ʽ"> <input name="Action" type="hidden" id="Action" value="<% = str_add %>"> 
        <input type="reset" name="Submit2" value="����"></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
set Fs_news = nothing
%>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_InterFace/HS_Function.asp" -->
<!--#include file="../../FS_InterFace/AP_Function.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Dim Conn,obj_Label_Rs,SQL,strShowErr
	MF_Default_Conn
	'session�ж�
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show
	Dim str_StyleName,txt_Content,Labelclass_SQL,obj_Labelclass_rs,obj_Count_rs
	if Request.QueryString("action")="del" then
		Conn.execute("Delete From FS_MF_LableClass where id="& CintStr(Request.QueryString("id"))&"")
		Conn.execute("Update FS_MF_Lable set LableClassID=0 where LableClassID="&CintStr(Request.QueryString("id")))
		strShowErr = "<li>ɾ���ɹ�</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Label/Label_Class.asp")
		Response.end
	end if
%>
<html>
<head>
<title>��ǩ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<body>
<table width="98%" height="76" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="hback" > 
    <td width="100%" height="20"  align="Left" class="xingmu">��ǩ��</td>
  </tr>
  <tr class="hback" > 
   <td height="27" align="center" class="hback"><div align="left"><a href="All_Label_Stock.asp">���б�ǩ</a>��<a href="../FreeLabel/FreeLabelList.asp"><font color="#FF0000">���ɱ�ǩ</font></a>��<a href="All_Label_Stock.asp?isDel=1">���ݿ�</a>��<a href="label_creat.asp">������ǩ</a>��<a href="Label_Class.asp" target="_self">��ǩ����</a>&nbsp;��<a href="All_label_style.asp">��ʽ����</a>&nbsp;<a href="../../help?Label=MF_Label_Stock" target="_blank" style="cursor:help;"><img src="../Images/_help.gif" width="50" height="17" border="0"></a></div></td>
  </tr>
  <tr class="hback" > 
    <td align="center" class="hback">
	<div align="left">
	<%
	dim rs_class,i
	set rs_class=Conn.execute("select id,ClassName,ClassContent,ParentID From FS_MF_LableClass where ParentID=0 order by id desc")
	do while not rs_class.EOF
		response.Write "��"&rs_class("ClassName")&"&nbsp;&nbsp;<a href=Label_Class.asp?id="&rs_class("id")&"&action=edit>[�޸�]</a><a href=Label_Class.asp?id="&rs_class("id")&"&action=del onClick=""{if(confirm('ȷ���������ѡ��ļ�¼��\nɾ���󣬴���Ŀ�µı�ǩ���ŵ���Ŀ¼�£�')){return true;}return false;}"">[ɾ��]</a><br />"
		response.Write get_childList1(rs_class("id"),"")
		rs_class.movenext
	loop
	rs_class.close:set rs_class=nothing
	%>
	
	</div>
	</td>
  </tr>
  <form name="form1" method="post" action="">
  </form>
</table>
<%
dim str_action,str_id,str_ClassName,str_ClassContent,rs_edit
if Request.QueryString("Action")="edit" then
	str_action = "Edit_Save"
	str_id= Request.QueryString("id")
	if not isnumeric(str_id) then
		strShowErr = "<li>����Ĳ���</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	set rs_edit = Conn.execute("select ClassName,ClassContent,id from FS_MF_LableClass where id="&CintStr(request.QueryString("id")))
	if rs_edit.eof then
		rs_edit.close:set rs_edit = nothing
		strShowErr = "<li>����Ĳ���</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		str_ClassName = rs_edit("ClassName")
		str_ClassContent = rs_edit("ClassContent")
		rs_edit.close:set rs_edit = nothing
	end if
else
	str_action = "Add_Save"
	str_id = ""
end if
%>
<table width="98%" height="76" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback" >
    <td colspan="2" align="center" class="xingmu"><div align="left">����/�޸ķ���</div></td>
  </tr>
  <form name="form1" method="post" action=""><tr class="hback" >
    <td width="16%" align="center" class="hback">
      ��ǩ��������</td>
      <td width="84%" align="center" class="hback"><div align="left">
        <label>
        <input name="ClassName" type="text" id="ClassName" value="<%=str_ClassName%>">
        </label>
      </div>	</td>
  </tr>
    <tr class="hback" >
      <td align="center" class="hback">������Ŀ</td>
      <td align="center" class="hback"><div align="left">
        <label>
        <select name="ParentID" id="ParentID">
          <option value="0">ѡ��������Ŀ</option>
		  <%
		  dim class_rs_obj,tmp_obj,tmp_ParentId
		  set tmp_obj = Conn.execute("select ParentID From FS_MF_LableClass where id="&CintStr(Request.QueryString("id")))
		  if tmp_obj.eof then
			  tmp_ParentId = 0
			  tmp_obj.close:set tmp_obj = nothing
		  else
			  tmp_ParentId = tmp_obj(0)
			  tmp_obj.close:set tmp_obj = nothing
		  end if
		  set class_rs_obj=Conn.execute("select id,ParentID,ClassName From FS_MF_LableClass where ParentID=0 order by id desc")
		  do while not class_rs_obj.eof 
		  		if tmp_ParentId = class_rs_obj("id") then
					response.Write "<option value="""&class_rs_obj("id")&""" selected>"&class_rs_obj("ClassName")&"</option>"
				else
					response.Write "<option value="""&class_rs_obj("id")&""">"&class_rs_obj("ClassName")&"</option>"
				end if
				response.Write get_childList(class_rs_obj("id"),"",tmp_ParentId)
		  	class_rs_obj.movenext
		  loop
		  class_rs_obj.close:set class_rs_obj=nothing
		  %>
        </select>
        </label>
      </div></td>
    </tr>
    <tr class="hback" >
    <td align="center" class="hback">˵��</td>
    <td align="center" class="hback"><div align="left">
      <label>
      <textarea name="ClassContent" cols="50" rows="6" id="ClassContent"><%=str_ClassContent%></textarea>
      </label>
    </div></td>
  </tr>
  <tr class="hback" >
    <td align="center" class="hback">&nbsp;</td>
    <td align="center" class="hback"><div align="left">
      <label>
      <input type="submit" name="Submit" value="������ǩ����">
      </label>
      <label>
      <input type="reset" name="Submit2" value="������д">
      </label>
      <input name="Action" type="hidden" id="Action" value="<%=str_action%>">
      <input name="ID" type="hidden" id="ID" value="<%=str_id%>">
    </div></td>
  </tr>
  </form>
</table>
</body>
<% 
If request.Form("Action")<>"" then
	dim rs,wheresql
	if Request.Form("Action")="Add_Save" then
		wheresql = " where 1=0"
		strShowErr = "<li>��������ɹ�</li><li><a href=Label/Label_Class.asp>��������</a></li><li><a href=Label/All_Label_Stock.asp>���ع���</A></li>"
	elseif Request.Form("Action")="Edit_Save" then
		strShowErr = "<li>�޸ķ���ɹ�</li><li><a href=Label/Label_Class.asp?id="&Request.Form("Id")&"&action=edit>�����޸�</a></li><li><a href=Label/All_Label_Stock.asp>���ع���</A></li>"
		wheresql = " where id="&CintStr(Request.Form("Id"))&""
	end if
	set rs = Server.CreateObject(G_FS_RS)
	rs.open "select ClassName,ClassContent,ParentID From FS_MF_LableClass "& wheresql &"",Conn,1,3
	if Request.Form("Action")="Add_Save" then
		rs.addnew
	end if
	rs("ClassName")=Request.Form("ClassName")
	rs("ClassContent")=Request.Form("ClassContent")
	rs("ParentID")=Request.Form("ParentID")
	rs.update
	rs.close
	set rs = nothing
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Label/All_Label_Stock.asp")
	Response.end
End if
Function get_childList(TypeID,f_CompatStr,tmp_ParentId)  
	Dim f_ChildNewsRs,ChildTypeListStr,f_TempStr,f_isUrlStr,lng_GetCount
	Set f_ChildNewsRs = Conn.Execute("Select id,ParentID,ClassName from FS_MF_LableClass where ParentID=" & CintStr(TypeID) & " order by id desc" )
	f_TempStr =f_CompatStr & "��"
	do while Not f_ChildNewsRs.Eof
			if tmp_ParentId = f_ChildNewsRs("id") then
				get_childList = get_childList & "<option value="""& f_ChildNewsRs("id") &""" selected>"
			else
				get_childList = get_childList & "<option value="""& f_ChildNewsRs("id") &""">"
			end if
			get_childList = get_childList & "��" &  f_TempStr & f_ChildNewsRs("ClassName") 
			get_childList = get_childList & "</option>" & Chr(13) & Chr(10)
			get_childList = get_childList &get_childList(f_ChildNewsRs("id"),f_TempStr,tmp_ParentId)
		f_ChildNewsRs.MoveNext
	loop
	f_ChildNewsRs.Close
	Set f_ChildNewsRs = Nothing
End Function
Function get_childList1(TypeID,f_CompatStr)  
	Dim f_ChildNewsRs,ChildTypeListStr,f_TempStr,f_isUrlStr,lng_GetCount
	Set f_ChildNewsRs = Conn.Execute("Select id,ParentID,ClassName from FS_MF_LableClass where ParentID=" & CintStr(TypeID) & " order by id desc" )
	f_TempStr =f_CompatStr & "��"
	do while Not f_ChildNewsRs.Eof
			get_childList1 = get_childList1 & "��" &  f_TempStr &f_ChildNewsRs("ClassName") &"&nbsp;&nbsp;<a href=Label_Class.asp?id="&f_ChildNewsRs("id")&"&action=edit>[�޸�]</a><a href=Label_Class.asp?id="&f_ChildNewsRs("id")&"&action=del onClick=""{if(confirm('ȷ���������ѡ��ļ�¼��\nɾ���󣬴���Ŀ�µı�ǩ���ŵ���Ŀ¼�£�')){return true;}return false;}"">[ɾ��]</a><br />"
			get_childList1 = get_childList1 &get_childList1(f_ChildNewsRs("id"),f_TempStr)
		f_ChildNewsRs.MoveNext
	loop
	f_ChildNewsRs.Close
	Set f_ChildNewsRs = Nothing
End Function
Set Conn=nothing
%>
</html>
<script language="JavaScript" type="text/JavaScript">
function insert(insertContent)
{
		obj=window.frames.item('NewsContent').EditArea.document.body;
		obj.focus();
	if(document.selection==null)
	{
		var iStart = obj.selectionStart
		var iEnd = obj.selectionEnd;
		obj.value = obj.value.substring(0, iEnd) +insertContent+ obj.value.substring(iEnd, obj.value.length);
	}else
	{
		var range = document.selection.createRange();
		range.text=insertContent;
	}
}
function opencat(cat)
{
  if(cat.style.display=="none"){
     cat.style.display="";
  } else {
     cat.style.display="none"; 
  }
}
</script>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Md5.asp" -->
<!--#include file="../FS_Inc/Cls_Cache.asp" -->
<%
Dim Conn,strShowErr
MF_Default_Conn
MF_Session_TF
Dim Temp_Admin_Name,Temp_Admin_Is_Super
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_Name = Session("Admin_Name")
if not MF_Check_Pop_TF("MF_Pop") then Err_Show
Dim p_name_str,p_pwd_str,p_truename_str,p_email_str,p_homepage_str,p_qq_str,p_sex_num,p_lock_num,p_child_num,p_style_num,p_selfintro_str,p_Admin_OnlyLogin,p_Admin_FilesTF
if Trim(Request.Form("act")) <>"" then
	p_name_str        = NoSqlHack(Request.Form("name"))
	p_pwd_str         = NoSqlHack(Request.Form("pwd"))
	p_truename_str    = NoSqlHack(Request.Form("truename"))
	p_email_str       = NoSqlHack(Request.Form("email"))
	p_homepage_str    = NoSqlHack(Request.Form("homepage"))
	p_qq_str          = NoSqlHack(Request.Form("qq"))
	p_sex_num         = NoSqlHack(Request.Form("sex"))
	p_lock_num        = NoSqlHack(Request.Form("lock"))
	p_child_num       = NoSqlHack(Request.Form("createchild"))
	p_style_num       = CintStr(Request.Form("style"))
	p_selfintro_str   = NoSqlHack(Request.Form("selfintro"))
	p_Admin_OnlyLogin = NoSqlHack(Request.Form("Admin_OnlyLogin"))
	p_Admin_FilesTF   = NoSqlHack(Request.Form("Admin_FilesTF"))
	Dim p_RsAdmin_add
	Set p_RsAdmin_add = CreateObject(G_FS_RS)
	if Request.Form("act")="add" then
		p_RsAdmin_add.open "select * from FS_MF_Admin where Admin_Name ='"& p_name_str&"'",Conn,3,3
		if Not p_RsAdmin_add.eof then
			strShowErr = "<li>����Ա����������������</li>"
			Response.Redirect("Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		Else
			p_RsAdmin_add.AddNew
			p_RsAdmin_add("Admin_Name") = p_name_str
			p_RsAdmin_add("Admin_Pass_Word") = md5(p_pwd_str,16)
			p_RsAdmin_add("Admin_Parent_Admin") = Temp_Admin_Name
			p_RsAdmin_add("Admin_Is_Super") = 0
		End if
	Else
		p_RsAdmin_add.open "select * from FS_MF_Admin where id ="&CintStr(Request.Form("id")),Conn,3,3
		if Trim(p_pwd_str)<>"" then
			p_RsAdmin_add("Admin_Pass_Word") = md5(p_pwd_str,16)
		End if
	End if
		p_RsAdmin_add("Admin_Real_Name") = p_truename_str
		p_RsAdmin_add("Admin_Email") = p_email_str
		p_RsAdmin_add("Admin_Home_Page") = p_homepage_str
		p_RsAdmin_add("Admin_Self_Intro") = p_selfintro_str
		p_RsAdmin_add("Admin_QQ") = p_qq_str
		p_RsAdmin_add("Admin_Sex") = p_sex_num
		p_RsAdmin_add("Admin_Is_Locked") = p_lock_num
		p_RsAdmin_add("Admin_Add_Admin") = p_child_num
		p_RsAdmin_add("Admin_Style_Num") = p_style_num
		if p_Admin_OnlyLogin <>"" then
			p_RsAdmin_add("Admin_OnlyLogin") = 1
		Else
			p_RsAdmin_add("Admin_OnlyLogin") = 0
		End if
		p_RsAdmin_add("Admin_FilesTF") = p_Admin_FilesTF
	p_RsAdmin_add.Update
	'��������ԱͼƬĿ¼
	if Request.Form("act")="add" then
		Dim p_FSO,tmps_path,Temps_AdminPath
		Set p_FSO = Server.CreateObject(G_FS_FSO)
			Temps_AdminPath = "..\"& G_UP_FILES_DIR &"\adminFiles"
			if p_FSO.FolderExists(Server.MapPath(Temps_AdminPath)) = false then p_FSO.CreateFolder(Server.MapPath(Temps_AdminPath))
			tmps_path = Temps_AdminPath & "\" & UCase(md5(p_name_str,16))
			if p_FSO.FolderExists(Server.MapPath(tmps_path)) = false then p_FSO.CreateFolder(Server.MapPath(tmps_path))
		set p_FSO = nothing
    End if
	p_RsAdmin_add.close : Set p_RsAdmin_add=Nothing
	strShowErr = "<li>�����ɹ�!</li>"
	Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_List.asp")
	Response.end
End if
Dim obj_admin_Rs,p_act
Dim p_Admin_Name,p_Admin_Pass_Word,p_Admin_Parent_Admin,p_Admin_Is_Super,p_Admin_Real_Name,p_Admin_Is_Locked,p_Admin_Pop_List,p_Admin_Email
Dim p_Admin_Home_Page,p_Admin_Self_Intro,p_Admin_QQ,p_Admin_Sex,p_Admin_Add_Admin,p_Admin_Login_Num,p_Admin_Reg_Time,p_Admin_Style_Num,p_id
if Request.QueryString("Action") = "edit" then
	Set obj_admin_Rs = server.CreateObject(G_FS_RS)
	obj_admin_Rs.Open "Select ID,Admin_Name,Admin_Pass_Word,Admin_Parent_Admin,Admin_Is_Super,Admin_Real_Name,Admin_Is_Locked,Admin_Pop_List,Admin_Email,Admin_Home_Page,Admin_Self_Intro,Admin_QQ,Admin_Sex,Admin_Add_Admin,Admin_Login_Num,Admin_Reg_Time,Admin_Style_Num,Admin_OnlyLogin,Admin_FilesTF from FS_MF_Admin where id="&CintStr(Request.QueryString("AdminID")),Conn,1,3
	if obj_admin_Rs("Admin_Name")<>Temp_Admin_Name then
		if Temp_Admin_Is_Super<>1 then
			if obj_admin_rs("Admin_Is_Super")=1 then
				strShowErr = "<li>�������޸�ϵͳ����Ա!</li>"
				Response.Redirect("Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_List.asp")
				Response.end
			end if
			if obj_admin_rs("Admin_Parent_Admin")<>Temp_Admin_Name then
				strShowErr = "<li>�˹���Ա���ϼ�����Ա���������������޸Ĵ˹���Ա!</li>"
				Response.Redirect("Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_List.asp")
				Response.end
			end if
		end if
	end if
	p_act = "Edit"
	p_id = obj_admin_Rs("ID")
	p_Admin_Name = obj_admin_Rs("Admin_Name")
	p_Admin_Pass_Word = obj_admin_Rs("Admin_Pass_Word")
	p_Admin_Parent_Admin = obj_admin_Rs("Admin_Parent_Admin")
	p_Admin_Real_Name = obj_admin_Rs("Admin_Real_Name")
	p_Admin_Is_Locked = obj_admin_Rs("Admin_Is_Locked")
	p_Admin_Pop_List = obj_admin_Rs("Admin_Pop_List")
	p_Admin_Email = obj_admin_Rs("Admin_Email")
	p_Admin_Home_Page = obj_admin_Rs("Admin_Home_Page")
	p_Admin_Self_Intro = obj_admin_Rs("Admin_Self_Intro")
	p_Admin_QQ = obj_admin_Rs("Admin_QQ")
	p_Admin_Sex = obj_admin_Rs("Admin_Sex")
	p_Admin_Add_Admin = obj_admin_Rs("Admin_Add_Admin")
	p_Admin_Login_Num = obj_admin_Rs("Admin_Login_Num")
	p_Admin_Reg_Time = obj_admin_Rs("Admin_Reg_Time")
	p_Admin_Style_Num = obj_admin_Rs("Admin_Style_Num")
	p_Admin_OnlyLogin = obj_admin_Rs("Admin_OnlyLogin")
	p_Admin_FilesTF = obj_admin_Rs("Admin_FilesTF")
	obj_admin_Rs.close:set obj_admin_Rs = nothing
Else
	p_act = "add"
	p_Admin_Parent_Admin = Temp_Admin_Name
	Dim obj_Add_Admin_Pop_rs
	Set obj_Add_Admin_Pop_rs = server.CreateObject(G_FS_RS)
	obj_Add_Admin_Pop_rs.open "select Admin_Add_Admin,Admin_Is_Super from FS_MF_Admin where Admin_Name ='"&  Temp_Admin_Name &"'",Conn,1,3
	if obj_Add_Admin_Pop_rs("Admin_Is_Super")=0 then
		if obj_Add_Admin_Pop_rs("Admin_Add_Admin") = 0 then
			strShowErr = "<li>��ûȨ�޽�������Ա</li>"
			Response.Redirect("Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
	End if	
	obj_Add_Admin_Pop_rs.close:set obj_Add_Admin_Pop_rs = nothing
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/Prototype.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/CheckJs.js" type="text/JavaScript"></script>
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="xingmu">
    <td class="xingmu">����Ա����</td>
  </tr>
  <tr class="hback">
    <td><a href="SysAdmin_List.asp">����Ա��ҳ</a> | <a href="SysAdmin_List.asp?Is_Super=1">��������Ա</a> | <a href="SysAdmin_List.asp?islock=1">�����Ĺ���Ա</a> | <a href="SysAdmin_List.asp?islock=0">���ŵĹ���Ա</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form action="" method="post" name="newadmin" id="newadmin" onSubmit="return checkinput();">
    <tr class="hback">
      <td colspan="2" class="xingmu">���\�޸Ĺ���Ա</td>
    </tr>
    <tr class="hback">
      <td align="right">������Ա</td>
      <td><input name="Admin_Parent_Admin" type="text" id="Admin_Parent_Admin" value="<%=p_Admin_Parent_Admin%>" size="60" maxlength="16" readonly>
        0��ʾ��������Ա���ϼ�û����Ա</td>
    </tr>
    <tr class="hback">
      <td width="140" align="right">����Ա�ʺ�</td>
      <td><input name="name" type="text" onFocus="Do.these('name',function(){return CheckContentLen('name','span_name','3-20')})" onKeyUp="Do.these('name',function(){return CheckContentLen('name','span_name','3-20')})" value="<% = p_Admin_Name%>" size="60" maxlength="16" <%if Request.QueryString("Action") = "edit" then response.Write("Readonly")%> />
        <span id="span_name"></span>
        <input name="ID" type="hidden" id="ID" value="<% = p_id %>">
      </td>
    </tr>
    <tr class="hback">
      <td width="140" align="right">����</td>
      <td><input name="pwd" type="password" onFocus="Do.these('pwd',function(){return CheckContentLen('pwd','span_pwd','6-16')})" onKeyUp="Do.these('pwd',function(){return CheckContentLen('pwd','span_pwd','6-16')})" value="" size="60" maxlength="16" />
        <span id="span_pwd"></span> ���޸��뱣��Ϊ�� </td>
    </tr>
	<tr class="hback">
      <td width="140" align="right">�ظ�����</td>
      <td><input name="Rpwd" type="password" onFocus="Do.these('Rpwd',function(){return CheckContentLen('Rpwd','span_Rpwd','6-16')})" onKeyUp="Do.these('Rpwd',function(){return CheckContentLen('Rpwd','span_Rpwd','6-16')})" value="" size="60" maxlength="16" />
        <span id="span_Rpwd"></span> ���޸��뱣��Ϊ�� </td>
    </tr>
    <tr class="hback">
      <td width="140" align="right">��ʵ����</td>
      <td><input name="truename" type="text" onFocus="Do.these('truename',function(){return isEmpty('truename','span_truename')})" onKeyUp="Do.these('truename',function(){return isEmpty('truename','span_truename')})" value="<% = p_Admin_Real_Name%>" size="60" maxlength="10" />
        &nbsp;<span id="span_truename"></span> </td>
    </tr>
    <tr class="hback">
      <td width="140" align="right">�����ַ</td>
      <td><input name="email" type="text" onFocus="Do.these('email',function(){return checkMail('email','span_email')})" onKeyUp="Do.these('email',function(){return checkMail('email','span_email')})" value="<% = p_Admin_Email%>" size="60" maxlength="50" />&nbsp;<span id="span_email"></span>
      </td>
    </tr>
    <tr class="hback">
      <td width="140" align="right">��ҳ</td>
      <td><input name="homepage" type="text" value="<% = p_Admin_Home_Page%>" size="60" maxlength="100" />
      </td>
    </tr>
    <tr class="hback">
      <td width="140" align="right">����ԱQQ</td>
      <td><input name="qq" type="text" value="<% = p_Admin_QQ%>" size="60" maxlength="16" />
      </td>
    </tr>
	 <tr class="hback">
      <td width="140" align="right">�ļ�Ȩ��</td>
      <td><select name="Admin_FilesTF" id="Admin_FilesTF" style="width:100">
          <option value="1" <%if p_Admin_FilesTF = 1 then response.Write("selected")%>>��</option>
          <option value="0" <%if p_Admin_FilesTF = 0 then response.Write("selected")%>>��</option>
        </select>
		<span class="tx">*�Ƿ�����ù���Ա����"<% = G_UP_FILES_DIR %>"Ŀ¼������Ŀ¼�ļ���</span>
      </td>
    </tr>
    <tr class="hback">
      <td width="140" align="right">�Ա�</td>
      <td><select name="sex" style="width:100">
          <option value="1" <%if p_Admin_Sex=1 then response.Write("selected")%>>��</option>
          <option value="0" <%if p_Admin_Sex=0 then response.Write("selected")%>>Ů</option>
        </select>
      </td>
    </tr>
    <tr class="hback">
      <td width="140" align="right">�Ƿ�����</td>
      <td><select name="lock" style="width:100">
          <option value="0" <%if p_Admin_Is_Locked=0 then response.Write("selected")%>>������</option>
          <option value="1"<%if p_Admin_Is_Locked=1 then response.Write("selected")%>>����</option>
        </select>
      </td>
    </tr>
    <tr class="hback">
      <td width="140" align="right">�½��¼�����Ա</td>
      <td><select name="createchild" style="width:100">
          <option value="1" <%if p_Admin_Add_Admin=1 then response.Write("selected")%>>����</option>
          <option value="0" <%if p_Admin_Add_Admin=0 then response.Write("selected")%>>������</option>
        </select>
      </td>
    </tr>
    <tr class="hback">
      <td width="140" align="right">��̨ʹ�÷��</td>
      <td><select name="style" style="width:100">
          <option value="3" <%if p_Admin_Style_Num=3 then response.Write("selected")%>>��ɫ����</option>
          <option value="1" <%if p_Admin_Style_Num=1 then response.Write("selected")%>>Ĭ�Ϸ��</option>
          <option value="2" <%if p_Admin_Style_Num=2 then response.Write("selected")%>>��ɫ���</option>
          <option value="4" <%if p_Admin_Style_Num=4 then response.Write("selected")%>>��������</option>
          <option value="5" <%if p_Admin_Style_Num=5 then response.Write("selected")%>>����Ӳ�</option>
        </select>
      </td>
    </tr>
    <tr class="hback">
      <td align="right">ֻ����һ���˵�½</td>
      <td><input name="Admin_OnlyLogin" type="checkbox" id="Admin_OnlyLogin" value="1" <%if p_Admin_OnlyLogin=1 then response.Write("checked")%>>
        ��</td>
    </tr>
    <tr class="hback">
      <td width="140" align="right">���ҽ���</td>
      <td><textarea name="selfintro" cols="60" rows="6"><% = p_Admin_Self_Intro%>
</textarea>
      </td>
    </tr>
    <tr class="hback">
      <td align="right">&nbsp;</td>
      <td><input type="submit" name="Submit3" value=" ���� ">
        <input type="reset" name="Submit4" value=" ���� ">
        <input name="act" type="hidden" id="act" value="<% = p_act %>"></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
Conn.Close
Set Conn = Nothing
%>
<script language="JavaScript" type="text/JavaScript">
function checkinput(){
	if($("name").value=='')
		{
			alert('����д����Ա�ʺ�');
			newadmin.name.focus();
			return false;
		}
	if($("name").value.length>20||$("name").value.length<3)
		{
			alert('�ʻ�����Ϊ3-20');
			$("name").focus();
			return false;
		}
	<% If Session("Admin_Is_Super")=1 Then %>
	if ($("pwd").value!=""){
		if ($("pwd").value.length>18||$("pwd").value.length<6)
			{
				alert('���볤��Ϊ6-18');
				$("pwd").focus();
				return false;
			}else if ($("Rpwd").value.length>18||$("Rpwd").value.length<6){
				alert('���볤��Ϊ6-18');
				$("Rpwd").focus();
				return false;
			}else if ($("pwd").value!=$("Rpwd").value){
				alert('�������벻һ��');
				$("Rpwd").focus();
				return false;
			}
		}
	<% End If %>
	if($("email").value!=''){
		if (!checkMail('email','')){
			alert('����д��ȷ��Email��ַ');
			$("email").focus();
			return false;
		}
	}
	if($("truename").value=='')
		{
			alert('����д��ʵ����');
			$("truename").focus();
			return false;
		}
	}
</script>
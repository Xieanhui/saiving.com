<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
'session�ж�
MF_Session_TF 
if not MF_Check_Pop_TF("WS002") then Err_Show

Dim i,Style
dim ConfigRs,sql,IsUser,PageS,Title,UserMember,strShowErr,RepUserMember,IsAudit
Set ConfigRs= server.CreateObject (G_FS_RS)
sql="select top 1 Title,isUser,PageSize,Style,UserMember,RepUserMember,IsAut From FS_WS_Config"
ConfigRs.open sql,conn,1,1
	if not ConfigRS.eof then
		IsUser=ConfigRs("IsUser")
		PageS=Cint(ConfigRs("PageSize"))
		Style=ConfigRs("Style")
		Title=ConfigRs("Title")
		UserMember=ConfigRs("UserMember")
		RepUserMember = ConfigRs("RepUserMember")
		IsAudit =  ConfigRs("IsAut")&""
	End if
Set ConfigRs=nothing
%>
<html>
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</HEAD>
<body>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form action="?Action=save" method="post" name="SysParaForm" id="SysParaForm">
    <tr class="hback">
      <td class="xingmu" colspan="4">��������&nbsp;&nbsp;&nbsp; <a href="../../help?Lable=Message" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
    </tr>
    <tr class="hback">
      <td width="17%" align="right">����ϵͳ����:</td>
      <td width="83%"><label>
        <input name="Title" type="text" id="Title" value="<%=Title%>" size="40">
        <font color="#FF0000">*</font></label></td>
    </tr>
    <tr class="hback">
      <td align="right">����δע���û�����:</td>
      <td width="85%"><%if IsUser="0" then%>
        <input type="radio" name="IsUser" value="0" checked>
        ����&nbsp;&nbsp;
        <input type="radio" name="IsUser" value="1">
        ��ֹ
        <%else%>
        <input type="radio" name="IsUser" value="0" >
        ����&nbsp;&nbsp;
        <input type="radio" name="IsUser" value="1" checked>
        ��ֹ
        <%End if %>
      </td>
    </tr>
     <tr class="hback">
      <td align="right">�����Ƿ���Ҫ���:</td>
      <td width="85%">
        <input type="radio" name="IsAudit" value="1" <% If IsAudit = "1" Then Response.Write "checked" %>>
        ��&nbsp;&nbsp;
        <input type="radio" name="IsAudit" <% If IsAudit = "0" Then Response.Write "checked" %> value="0">
        ��
      </td>
    </tr>
	<tr>
      <td class="hback" align="right">�����ɹ����ӻ�Ա����:</td>
      <td class="hback"><input name="UserMember" type="text" onKeyUp="checknum(this)" value="<%=UserMember%>" size="15" maxlength="3">
        <label>��д������</label>
    </tr>
    <tr class="hback">
      <td align="right">�ظ����ӻ��֣�</td>
      <td><label>
        <input name="RepUserMember" type="text" id="RepUserMember" value="<% = RepUserMember %>" size="15" maxlength="3">
        </label>
        ��д������ 
    </tr>
    <tr class="hback">
      <td align="right">ÿҳ��ʾ��������:</td>
      <td width="85%"><select name="PageS">
          <%
		  for i=10 to 50 step 2
		  	if i=PageS then
		  %>
          <option value="<%=i%>" selected><%=i%></option>
          <% 
		  	else
			%>
          <option value="<%=i%>"><%=i%></option>
          <%
		  	end if
		  Next 
		  %>
        </select>
    </tr>
    <tr class="hback">
      <td align="right">ѡ����ʽ:</td>
      <td><select name="Style">
          <option value="1" <%if Style="1" then response.Write"selected"%>>Ĭ����ʽ</option>
          <option value="2" <%if Style="2" then response.Write"selected"%>>��ɫ���</option>
          <option value="3" <%if Style="3" then response.Write"selected"%>>��ɫ����</option>
          <option value="4" <%if Style="4" then response.Write"selected"%>>��ɫ����</option>
          <option value="5" <%if Style="5" then response.Write"selected"%>>��ɫ����</option>
        </select>
      </td>
    </tr>
    <tr class="hback">
      <td align="right">&nbsp;</td>
      <td><input type="Submit" name="btn_SetSysParam" value=" ���� " />
        <input type="reset" name="sub_rest" value=" ���� " /></td>
    </tr>
  </form>
</table>
<script language="javascript">
function checknum(obj){
if (isNaN(obj.value)|| obj.value<0){
alert("��ʽ����!");
obj.focus();
}
}
</script>
<%
if NoSqlHack(Request.QueryString("Action"))="save" Then
	IsAudit = Request.Form("IsAudit")
	If IsAudit = "" then IsAudit = "0"
	IsUser=Request.Form("IsUser")
	PageS=Request.Form("PageS")
	Style=Request.Form("Style")
	Title=NoSqlHack(trim(Replace(Request.Form("Title"),"'","")))
	UserMember=NoSqlHack(Trim(Request.form("UserMember")))
	RepUserMember = NoSqlHack(Trim(Request.form("RepUserMember")))
	if Title="" then
	strShowErr="<li>����ϵͳ���Ʋ���Ϊ��</li>"
	Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
	Response.end
	end if
	if IsNumeric(UserMember) then
		UserMember=int(abs(UserMember))
	else
		strShowErr = "<li>���ӻ�Ա���ֲ���</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	if IsNumeric(RepUserMember) then
		RepUserMember=int(abs(RepUserMember))
	else
		strShowErr = "<li>���ӻ�Ա���ֲ���</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	if IsUser<>"" and PageS<>"" then
		Conn.Execute("Update FS_WS_Config set Title='"&NoSqlHack(Title)&"',IsUser='"&NoSqlHack(IsUser)&"',PageSize="&CintStr(PageS)&",Style='"&NoSqlHack(Style)&"',UserMember="&CintStr(UserMember)&",RepUserMember="&CintStr(RepUserMember)&",IsAut="&CintStr(IsAudit)&"")
		strShowErr = "<li>�޸ĳɹ�</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	else
		strShowErr = "<li>��������</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	End if
End if
Set Conn=Nothing
%>
</body>
</html>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->







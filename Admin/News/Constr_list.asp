<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
Dim User_Conn, info_Rs,sql_cmd,sql_info_cmd,ClassID,AuditTF,SuperClass_Str,ClassPath,Conn
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
Dim ContID,title,addtime,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,isTF,untread '�����Ϣ
Response.CacheControl = "no-cache"

MF_Default_Conn
MF_User_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("NS_Constr") then Err_Show
Function GetFriendName(f_strNumber)
	Dim RsGetFriendName
	Set RsGetFriendName = User_Conn.Execute("Select UserName From FS_ME_Users Where UserNumber = '"& NoSqlHack(f_strNumber) &"'")
	If  Not RsGetFriendName.eof  Then 
		GetFriendName = RsGetFriendName("UserName") 
	Else
		GetFriendName = 0
	End If 
	set RsGetFriendName = nothing
End Function 
'MF_Check_Pop_TF�����������ļ�������+����
'if Not MF_Check_Pop_TF("NS_News_01") then Top_Go_To_Error_Page 
int_RPP=15 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"				'βҳ
'------------------------------------------------
ClassID=NoSqlHack(Request.QueryString("classid"))
AuditTF=NoSqlHack(Request.QueryString("audittf"))
if ClassID="" then
ClassID=0
End if
if AuditTF="" then
AuditTF=0
End if
User_Conn.CursorLocation = 3
Set info_Rs=Server.CreateObject(G_FS_RS)
if ClassID<>0 then'���classid��Ϊ��Ŀ¼
	if  AuditTF="1" then '����Ѿ����
		sql_info_cmd="Select ContID,ContTitle,addtime,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,isTF,MainID,untread from FS_ME_InfoContribution where IsPublic=1 and IsLock=0 and AuditTF=1 and MainID="&ClassID
	else'���δ���
		sql_info_cmd="Select ContID,ContTitle,addtime,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,isTF,MainID,untread from FS_ME_InfoContribution where IsPublic=1 and IsLock=0 and AuditTF=0 and MainID="&ClassID
	End if
else'���classidΪ��
	if AuditTF="1" then'����Ѿ����  
		sql_info_cmd="Select ContID,ContTitle,addtime,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,MainID,isTF,untread from FS_ME_InfoContribution where IsPublic=1 and IsLock=0 and AuditTF=1"
	else'���δ��� 
		sql_info_cmd="Select ContID,ContTitle,addtime,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,MainID,isTF,untread from FS_ME_InfoContribution where IsPublic=1 and IsLock=0 and AuditTF=0"
	End if
End IF
info_Rs.open sql_info_cmd,User_Conn,1,1
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language='javascript' src="../../FS_Inc/prototype.js"></script>
<script language='javascript' src="../../FS_Inc/publicjs.js"></script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>

<body>
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu">
	<td class="xingmu">�������</td>
	<td align='center'  class="xingmu">Ͷ��ʱ��</td>
  	<td align='center' class="xingmu">����</td>
	<td align='center' class="xingmu">��Ϣ��</td>
	<td align='center' class="xingmu">������</td>
	<td align='center' class="xingmu">��վ��ʾ</td>
	<td align='center' class="xingmu">���˸�</td>
  	<td align='center' class="xingmu">�Ƽ�</td>
	<td align='center' class="xingmu">����</td>
	<td align='center'><input type="checkbox" name="chk_constr" onClick="selectAll(document.all('chk_constr'))" value=""/></td>
  </tr>
  <%
	If Not info_Rs.eof then
		info_Rs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo<=0 Then cPageNo=1
		If cPageNo>info_Rs.PageCount Then cPageNo=info_Rs.PageCount 
		info_Rs.AbsolutePage=cPageNo
	End if

	FOR int_Start=1 TO int_RPP 
		if info_Rs.eof then exit for
		ContID=info_Rs("ContID")
		title=info_Rs("ContTitle")
		addtime=info_Rs("AddTime")
		ContSytle=info_Rs("ContSytle")
		if ContSytle=0 then
			ContSytle="ԭ��"
		elseif ContSytle=1 then
			ContSytle="ת��"
		elseif ContSytle=2 then
			ContSytle="����"
		End IF
		
		InfoType=info_Rs("InfoType")
		if InfoType=0 then
			InfoType="��ͨ"
		elseif InfoType=1 then
			InfoType="����"
		elseif InfoType=2 then
			InfoType="�Ӽ�"
		End IF
		UserNumber=info_Rs("UserNumber")
		isPublic=info_Rs("isPublic")
		if isPublic=1 then
			isPublic=info_Rs("MainID")
		else
			isPublic="��"
		End IF
		isTF=info_Rs("isTF")
		if isTF=1 then
			isTF="�Ƽ�"
		else
			isTF="��"
		End IF
		AdminLock=info_Rs("AdminLock")
		if AdminLock=1 then
			AdminLock="<font color=""red"">����</font>"
		else
			AdminLock="��"
		End IF
		untread=info_Rs("untread")
		if untread=1 then
			untread="<font color='red'>��</font>"
		else
			untread="-"
		End if
		Response.Write("<tr class='hback'>"&chr(10)&chr(13))
		Response.Write("<td><a href='Audit_Edit.asp?contid="&ContID&"'>"&title&"</a></td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'>"&addtime&"</td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'>"&ContSytle&"</td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'>"&InfoType&"</td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'><a href=../../"&G_USER_DIR&"/ShowUser.asp?UserNumber="&UserNumber&" target=_blank>"&GetFriendName(UserNumber)&"</a></td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'>"&isPublic&"</td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'>"&untread&"</td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'>"&isTF&"</td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'>"&AdminLock&"</td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'><input type=""checkbox"" name=""chk_constr"" id=""chk_constr"" value="""&ContID&"""></td>"&Chr(10)&Chr(13))
		Response.write("</tr>")
		info_Rs.movenext
	Next
  %>
  <tr>
  <td class="hback" colspan="10" align="right">
	<%if AuditTF="0" then Response.Write("<button onClick=""Operator('audit')"" >��  ��</button>&nbsp;")%>
	<button onClick="Operator('lock')">��  ��</button>&nbsp;
	<button onClick="Operator('unlock')">�������</button>&nbsp;
	<button onClick="Operator('tf')">��  ��</button>&nbsp;
	<button onClick="Operator('untf')">ȡ���Ƽ�</button>&nbsp;
	<button onClick="Operator('delete')">ɾ  ��</button>&nbsp;
	<button onClick="Operator('deleteAll')">ɾ������</button>&nbsp;
	<%if AuditTF="0" then Response.Write("<button onClick=""Operator('untread')"">��  ��</button>")%>
  </td>
  </tr>
  <%
  	Response.Write("<tr>"&vbcrlf)
	Response.Write("<td align='right' colspan='10'  class=""hback"">"&fPageCount(info_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td>"&vbcrlf)
	Response.Write("</tr>"&vbcrlf)
  %>
</table>
</body>
</html>
<script language="javascript">
//ִ���ύ������
//param������Ϊѡ��ִ�ж����ı�־

function Operator(param)
{
	var values=getCheckBoxValues();
	if(param!="deleteAll"&&values.length<2)
	{	
		alert("��ѡ���������");
		return;
	}
	switch(param)
	{
		case "audit":location="Constr_Action.asp?act=audit&values="+values;break;//���
		case "recall":location="Constr_Action.asp?act=recall&values="+values;break//�������
		case "lock":location="Constr_Action.asp?act=lock&values="+values;break//����
		case "unlock":location="Constr_Action.asp?act=unlock&values="+values;break//�������
		case "tf":location="Constr_Action.asp?act=tf&values="+values;break//�Ƽ�
		case "untf":location="Constr_Action.asp?act=untf&values="+values;break//ȡ���Ƽ�
		case "delete": if(confirm("ȷ��ɾ����"))location="Constr_Action.asp?act=delete&values="+values;break//ɾ��
		case "deleteAll": if(confirm("ȷ��ɾ�����м�¼��"))location="Constr_Action.asp?act=deleteAll";break//ɾ������
		case "untread": if(confirm("ȷ���˸壿"))location="Constr_Action.asp?act=untread&values="+values;break//�˸�
	}
	
	
}
//��ñ�ѡ�ֵ�checkbox��ֵ�ļ�
function getCheckBoxValues()
{
	var checkBoxValues=""
	for(var i=1;i<document.all("chk_constr").length;i++)
	{
		if(document.all("chk_constr")[i].checked)
		{
			checkBoxValues=checkBoxValues+","+document.all("chk_constr")[i].value;
		}
	}
	return checkBoxValues;
}
parent.document.all.hd_classid.value="<%=Classid%>"
parent.document.all.hd_audit.value="<%=AuditTF%>"
</script>
<%
info_Rs.close
User_Conn.close
Set info_Rs=nothing
Set User_Conn=nothing
%>






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
Dim Conn,CharIndexStr,strShowErr,obj_form_rs,form_sql
Dim act,itemID,formid,ItemName,FieldName,orderby,StateSet,IsNullset,ItemType,MaxSize,DefaultValue,SelectItem,Remark,formName
MF_Default_Conn
MF_Session_TF 
act=NoSqlHack(Request.QueryString("act"))
formid=NoSqlHack(Request.QueryString("formid"))
itemID=NoSqlHack(Request.QueryString("FormItemID"))
if act="edit" then
	if not MF_Check_Pop_TF("MF093") then Err_Show
	if itemID="" then	
		strShowErr = "<li>�Ƿ����ݴ���</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	form_sql="select FormItemID,formid,ItemName,FieldName,orderby,State,IsNull,ItemType,MaxSize,DefaultValue,SelectItem,Remark from FS_MF_CustomForm_Item where FormItemID="&itemID
	set obj_form_rs=conn.execute(form_sql)
	if obj_form_rs.eof then 
		strShowErr = "<li>���������ݲ���ȷ��</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	itemID=obj_form_rs("FormItemID")
	formid=obj_form_rs("formid")
	ItemName=obj_form_rs("ItemName")
	FieldName=obj_form_rs("FieldName")
	orderby=obj_form_rs("orderby")
	StateSet=obj_form_rs("State")
	IsNullset=obj_form_rs("IsNull")
	ItemType=obj_form_rs("ItemType")
	MaxSize=obj_form_rs("MaxSize")
	DefaultValue=obj_form_rs("DefaultValue")
	SelectItem=obj_form_rs("SelectItem")
	Remark=obj_form_rs("Remark")
	form_sql="select formName from FS_MF_CustomForm where id="&formid
	set obj_form_rs=conn.execute(form_sql)
	if obj_form_rs.eof then 
		strShowErr = "<li>���������ݲ���ȷ��</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	formName=obj_form_rs(0)
elseif act="del" then
	if not MF_Check_Pop_TF("MF092") then Err_Show
	if itemID="" then
		strShowErr = "<li>�Ƿ����ݴ���</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	dim tableName
	'ȡ�ֶ����ͱ�ID
	form_sql="select FormID,FieldName from FS_MF_CustomForm_Item where FormItemID="&itemID
	set obj_form_rs=conn.execute(form_sql)
	if obj_form_rs.eof then 
		strShowErr = "<li>���������ݲ���ȷ��</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		formid=obj_form_rs(0)
		FieldName=obj_form_rs(1)
	end if
	'ȡ����
	form_sql="select tableName from FS_MF_CustomForm where id="&formid
	set obj_form_rs=conn.execute(form_sql)
	if obj_form_rs.eof then 
		strShowErr = "<li>���������ݲ���ȷ��</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		tableName=obj_form_rs(0)
	end if
	'�޸ı�
	form_sql="ALTER TABLE "&tableName&" DROP COLUMN "&FieldName
	conn.execute(form_sql)
	form_sql="delete * from FS_MF_CustomForm_Item where FormItemID="&itemID
	conn.execute(form_sql)
	strShowErr = "<li>��ϲ��ɾ������ɹ�!</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../FormItem.asp?formid="&formid)
	Response.end
else
	if not MF_Check_Pop_TF("MF094") then Err_Show
	if formID="" then	
		strShowErr = "<li>�Ƿ����ݴ���</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	itemID=""
	ItemName=""
	FieldName=""
	orderby="1"
	StateSet=0
	IsNullset=1
	ItemType="SingleLineText"
	MaxSize="0"
	DefaultValue=""
	SelectItem=""
	Remark=""
	form_sql="select formName from FS_MF_CustomForm where id="&formid
	set obj_form_rs=conn.execute(form_sql)
	if obj_form_rs.eof then 
		strShowErr = "<li>���������ݲ���ȷ��</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	formName=obj_form_rs(0)
end if
obj_form_rs.close()
set obj_form_rs=nothing
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�Զ��������___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
</head>
<body onLoad="changetype('<% = ItemType %>');">
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="hback"> 
    <td width="100%" class="xingmu"><a href="#" onMouseOver="this.T_BGCOLOR='#404040';this.T_FONTCOLOR='#FFFFFF';return escape('<div align=\'center\'>FoosunCMS5.0<br> Code by Samjun <BR>Copyright (c) 2006 Foosun Inc</div>')" class="sd"><strong>�Զ������</strong></a><a href="../../help?Lable=NS_Form_Custom" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a>��<a href="FormOperate.asp?act=add"></a>��������������������������    </td>
  </tr>
  <tr>
      <td height="18" class="hback"><div align="left"><a href="FormManage.asp">�Զ��������</a></div></td>
    </tr>
</table>
  <table  width="98%" border="0" cellspacing="1" cellpadding="5" align="center" class="table">
<form name="form1" method="post" onSubmit="return CheckData(this);" action="FormItemSave.asp?act=<%=act%>">
    <tr>
      <td width="23%" align="right" class="hback">�����ƣ�</td>
      <td width="77%" class="hback"><%=formName%></td>
    </tr>
    <tr>
      <td align="right" class="hback">�������ƣ�</td>
      <td class="hback"><INPUT id="ItemName" maxLength="50" <% if act="edit" then response.write("readonly") %> name="ItemName" value="<%=ItemName%>"></td>
    </tr>
    <tr>
      <td align="right" class="hback">�ֶ�����</td>
      <td class="hback"><INPUT id="FieldName" maxLength="50" <% if act="edit" then response.write("readonly") %> name="FieldName" value="<%=FieldName%>" <%if act="edit" then response.Write(" readonly")%>></td>
    </tr>
    <tr>
      <td align="right" class="hback">������ţ�</td>
      <td class="hback"><INPUT id="orderby" maxLength="50" name="orderby" value="<%=orderby%>"> 
      ���ԽС������Խǰ��</td>
    </tr>
    <tr>
      <td align="right" class="hback">�Ƿ����ã�</td>
      <td class="hback"><INPUT type="radio" value="0" name="StateSet" <%if StateSet=0 then response.Write("checked")%>>
        ��
        <INPUT  type="radio" value="1" name="StateSet" <%if StateSet=1 then response.Write("checked")%>>
        ��</td>
    </tr>

    <tr>
      <td align="right" class="hback">�Ƿ���</td>
      <td class="hback"><INPUT <% if act <> "add" then Response.Write("disabled") %> type="radio" value="0" name="IsNullset" <%if IsNullset=0 then response.Write("checked")%>>
        ��
        <INPUT  type="radio" value="1" name="IsNullset" <%if IsNullset=1 then response.Write("checked")%>>
        ��</td>
    </tr>
    <TR id="tr_tms">
      <TD align="right" class="hback">�������ͣ�</TD>
      <TD align="left" class="hback"><SELECT <% if act <> "add" then Response.Write("disabled") %> onChange="changetype(this.value);" name="ItemType">
        <OPTION value="SingleLineText" <%if ItemType="SingleLineText" then response.Write("selected")%>>�����ı�</OPTION>
        <OPTION value="MultiLineText" <%if ItemType="MultiLineText" then response.Write("selected")%>>�����ı�</OPTION>
        <OPTION value="PassWordText" <%if ItemType="PassWordText" then response.Write("selected")%>>����</OPTION>
        <OPTION value="DateTime" <%if ItemType="DateTime" then response.Write("selected")%>>����ʱ��</OPTION>
        <OPTION value="RadioBox" <%if ItemType="RadioBox" then response.Write("selected")%>>��ѡ��</OPTION>
        <OPTION value="CheckBox" <%if ItemType="CheckBox" then response.Write("selected")%>>��ѡ��</OPTION>
        <OPTION value="Numberic" <%if ItemType="Numberic" then response.Write("selected")%>>����</OPTION>
        <OPTION value="UploadFile" <%if ItemType="UploadFile" then response.Write("selected")%>>����</OPTION>
        <OPTION value="DropList" <%if ItemType="DropList" then response.Write("selected")%>>������</OPTION>
        <OPTION value="List" <%if ItemType="List" then response.Write("selected")%>>�б��</OPTION>
      </SELECT></TD>
    </TR>
    <TR id="tr_tme">
      <TD align="right" class="hback">�ı�����</TD>
      <TD align="left" class="hback"><input <% if act <> "add" then Response.Write("disabled") %> name="MaxSize" value="<%=MaxSize%>">
      0��ʾ���޳���</TD>
    </TR>
    <tr>
      <td align="right" class="hback">Ĭ��ֵ��</td>
      <td class="hback"><INPUT <% if act <> "add" then Response.Write("disabled") %> maxLength="50" name="DefaultValue" value="<%=DefaultValue%>"></td>
    </tr>
    <tr id="tr_SelectItem">
      <td align="right" class="hback">ѡ�</td>
      <td class="hback"><TEXTAREA name="SelectItem" cols="40" rows="8" id="SelectItem"><%=SelectItem%></TEXTAREA>
ÿһ��Ϊһ���б�ѡ��</td>
    </tr>
    <tr>
      <td align="right" class="hback">������ʾ��</td>
      <td class="hback"><TEXTAREA name="Remark" cols="40" rows="8" id="Remark"><%=Remark%></TEXTAREA>
(�������Ե���ʾ��Ϣ��255���ַ�������Ч)</td>
    </tr>
    <tr>
      <td align="right" class="hback">&nbsp;</td>
      <td class="hback"><input type="hidden" name="formid" value="<%=formid%>">
	  <input type="hidden" name="itemID" value="<%=itemID%>">
	  <INPUT type="submit" value=" ȷ�� " name="BtnOK">
        <INPUT name="reset" type="reset" value=" ��д "></td>
    </tr>
</form>
  </table>
</body>
</html>
<script language="javascript" type="text/javascript">
function changetype(val)
{
	var f = 'none';
	if(val == 'RadioBox' || val == 'CheckBox' || val == 'DropList' || val == 'List') f = '';
	document.getElementById('tr_SelectItem').style.display = f;
}
function CheckData(theForm){
<% if act="add" then %>
	if(theForm.ItemName.value==''){
		alert('�����������');
		theForm.ItemName.focus();
		return false;
	}
	if(theForm.FieldName.value==''){
		alert('����д�ֶ�����');
		theForm.FieldName.focus();
		return false;
	}
<% end if %>
	if(theForm.orderby.value==''){
		alert('����д������ţ�');
		theForm.orderby.focus();
		return false;
	}
	if (theForm.orderby.value!='' && (isNaN(theForm.orderby.value) || theForm.orderby.value<0)){
		alert("�������Ӧ����Ч���֣�");
		theForm.orderby.value="";
		theForm.orderby.focus();
		return false;
	}
	if(theForm.MaxSize.value==''){
		alert('����д�ı����ȣ�');
		theForm.MaxSize.focus();
		return false;
	}
	if (theForm.MaxSize.value!='' && (isNaN(theForm.MaxSize.value) || theForm.MaxSize.value<0)){
		alert("�ı�����Ӧ����Ч���֣�");
		theForm.MaxSize.value="";
		theForm.MaxSize.focus();
		return false;
	}
	return true;
}
</script>
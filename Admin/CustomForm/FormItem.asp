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
Dim Conn,strShowErr,obj_form_rs,formlist_sql,showMorePageGo_Type_,cPageNo,i
Dim strpage,int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_

int_RPP=20'����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"

MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("MF095") then Err_Show
dim FormID
FormID=NoSqlHack(request.QueryString("FormID"))
if formID="" then	
	strShowErr = "<li>�Ƿ����ݴ���</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�Զ��������___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="hback"> 
    <td width="100%" class="xingmu"><a href="FormManage.asp" onMouseOver="this.T_BGCOLOR='#404040';this.T_FONTCOLOR='#FFFFFF';return escape('<div align=\'center\'>FoosunCMS5.0<br> Code by Samjun <BR>Copyright (c) 2006 Foosun Inc</div>')" class="sd"><strong>�������</strong></a><a href="../../help?Lable=NS_Form_Custom" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a>����������������������������    </td>
  </tr>
  <tr>
      <td height="18" class="hback"><a href="FormManage.asp">������</a>&nbsp;&nbsp;<a href="FormItemOperate.asp?act=add&formid=<%=formID%>">�½�����</a></td>
  </tr>
</table>
<%
strpage=request("page")
if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
formlist_sql = "Select FormItemID,orderby,ItemName,FieldName,ItemType,IsNull from FS_MF_CustomForm_Item where formID="&formID&" order by orderby"
Set obj_form_rs = Server.CreateObject(G_FS_RS)
obj_form_rs.Open formlist_sql,Conn,1,1
%>

<table width="98%" border="0" cellspacing="1" cellpadding="5" align="center" class="table">
  <tr class="xingmu">
    <td width="6%" align="center" class="xingmu">˳��</td>
    <td width="25%" align="center" class="xingmu">������</td>
    <td width="23%" align="center" class="xingmu">�ֶ���</td>
    <td width="20%" align="center" class="xingmu">�ֶ�����</td>
    <td width="13%" align="center" class="xingmu">�Ƿ����</td>
    <td width="13%" align="center" class="xingmu">����</td>
  </tr>
  <%
  for i=1 to obj_form_rs.pagesize
  	if obj_form_rs.eof then exit for
  %>
  <tr class="hback">
    <td align="center" class="hback"><%=obj_form_rs("orderby")%></td>
    <td class="hback"><%=obj_form_rs("ItemName")%></td>
    <td class="hback"><%=obj_form_rs("FieldName")%></td>
    <td class="hback"><%
	select case obj_form_rs("ItemType")
		case "SingleLineText"
			response.Write("�����ı�")
		case "MultiLineText"
			response.Write("�����ı�")
		case "PassWordText"
			response.Write("����")
		case "DateTime"
			response.Write("����ʱ��")
		case "RadioBox"
			response.Write("��ѡ��")
		case "CheckBox"
			response.Write("��ѡ��")
		case "Numberic"
			response.Write("����")
		case "UploadFile"
			response.Write("����")
		case "DropList"
			response.Write("������")
		case "List"
			response.Write("�б��")
	end select%></td>
    <td align="center" class="hback"><%if obj_form_rs("IsNull")=0 then response.Write("������") else response.Write("�Ǳ�����") end if%></td>
    <td align="center" class="hback"><A href="FormItemOperate.asp?act=edit&FormItemID=<%=obj_form_rs("FormItemID")%>">�޸�</A> <A onClick="if(confirm('ȷ��Ҫɾ����'))location='FormItemOperate.asp?act=del&FormItemID=<%=obj_form_rs("FormItemID")%>';" href="javascript:void(0);">ɾ��</A> </td>
  </tr>
  <%
  	obj_form_rs.movenext
  next
  %>
  <tr>
    <td colspan="6"class="hback">
	  <table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td width="79%" colspan="2" align="right" class="hback"> <%
			response.Write "<p>"&  fPageCount(obj_form_rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
	%> </td>
        </tr>
      </table>
	 </td>
  </tr>
</table>
</body>
</html>
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
Dim Sql,Result,ExeResult,ExeResultNum,ExeSelectTF,ErrorTF,FiledObj
Dim I,j,ErrObj,conn,formid,obj_form_rs,tableName,action,FormData,strShowErr
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("MF091") then Err_Show

Dim int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_showNumberLink_=8
showMorePageGo_Type_ = 1
str_nonLinkColor_="#999999"
toF_="<font face=webdings title=""��ҳ"">9</font>"
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"
toP1_=" <font face=webdings title=""��һҳ"">3</font>"
toN1_=" <font face=webdings title=""��һҳ"">4</font>"
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"
toL_="<font face=webdings title=""���һҳ"">:</font>"

cPageNo = NoSqlHack(Request("Page"))
formid=NoSqlHack(Request("formid"))
if formID="" then	
	strShowErr = "<li>�Ƿ����ݴ���</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
end if
sql="select tableName from FS_MF_CustomForm where id="&formid
set obj_form_rs=conn.execute(sql)
if obj_form_rs.eof then 
	obj_form_rs.Close
	Set obj_form_rs = Nothing
	strShowErr = "<li>���������ݲ���ȷ��</li>"
    Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
   	Response.end
else
	tableName=obj_form_rs(0)
end if
obj_form_rs.Close
Set obj_form_rs = Nothing

Action = NoSqlHack(Request("Action"))
FormData = RemovePlace(NoSqlHack(Request.Form("FormData")))
if action = "DelAll" then
	if not MF_Check_Pop_TF("MF090") then Err_Show
	Sql = "Delete from " & tableName
	Conn.Execute(Sql)
	strShowErr = "<li>��ϲ��ɾ��ȫ���ɹ�!</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=" & Server.URLEncode("../Form_Data.asp?formid=" & formid & "&Page=" & cPageNo))
	Response.end
end if
if FormData <> "" then
	if action = "Del" then
		if not MF_Check_Pop_TF("MF090") then Err_Show
		Sql = "Delete from " & tableName & " Where ID In (" & FormData & ")"
		Conn.Execute(Sql)
		strShowErr = "<li>��ϲ��ɾ���ɹ�!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=" & Server.URLEncode("../Form_Data.asp?formid=" & formid & "&Page=" & cPageNo))
		Response.end
	end if
	if action = "Lock" then
		if not MF_Check_Pop_TF("MF089") then Err_Show
		Sql = "Update " & tableName & " Set form_lock=1 Where ID In (" & FormData & ")"
		Conn.Execute(Sql)
		strShowErr = "<li>��ϲ�������ɹ�!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=" & Server.URLEncode("../Form_Data.asp?formid=" & formid & "&Page=" & cPageNo))
		Response.end
	end if
	if action = "UNLock" then
		if not MF_Check_Pop_TF("MF088") then Err_Show
		Sql = "Update " & tableName & " Set form_lock=0 Where ID In (" & FormData & ")"
		Conn.Execute(Sql)
		strShowErr = "<li>��ϲ�������ɹ�!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=" & Server.URLEncode("../Form_Data.asp?formid=" & formid & "&Page=" & cPageNo))
		Response.end
	end if
end if
sql="select * from "&tableName
Set ExeResult = Server.CreateObject(G_FS_RS)
ExeResult.Open Sql,Conn,1,1
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ִ�н��</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<body topmargin="2" leftmargin="2">
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu">
    <td class="xingmu"><a href="#" class="sd"><strong>������</strong></a><a href="../../help?Lable=NS_Form_Custom" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr class="hback">
    <td class="hback"><a href="FormManage.asp">������</a></td>
  </tr>
</table>
<table align="center" width="98%" border="0" cellpadding="2" cellspacing="1" class="table">
  <form name="DataList" method="post" action="?">
        <tr class="habck"> 
          <%
		Dim ArrayFieldName(5),ArrayFieldText(5)
		ArrayFieldName(0) = "form_usernum" : ArrayFieldText(0) = "�û���"
		ArrayFieldName(1) = "form_username" : ArrayFieldText(1) = "�û���"
		ArrayFieldName(2) = "form_ip" : ArrayFieldText(2) = "IP��ַ"
		ArrayFieldName(3) = "form_time" : ArrayFieldText(3) = "���ʱ��"
		ArrayFieldName(4) = "form_lock" : ArrayFieldText(4) = "����"
		ArrayFieldName(5) = "form_answer" : ArrayFieldText(5) = "�Ƿ�ظ�"
		For j = LBound(ArrayFieldText) To UBound(ArrayFieldText)
%>
          <td nowrap  height="26" class="xingmu"><div align="center"> 
              <% = ArrayFieldText(j) %>
            </div></td>
          <%
		next
%>
          <td nowrap  height="26" class="xingmu"><div align="center"> 
              <input name="FormData" onClick="selectAll(this.form);" title="ȫѡ" type="checkbox" value="">
            </div></td>
        </tr>
        <%
if Not ExeResult.Eof then
	ExeResult.PageSize = 10
	if cPageNo = "" then cPageNo = 1
	if Not IsNumeric(cPageNo) then cPageNo = 1
	cPageNo = CInt(cPageNo)
	if cPageNo > ExeResult.PageCount then cPageNo = ExeResult.PageCount
	ExeResult.AbsolutePage = cPageNo
	For i = 1 To ExeResult.PageSize
		if ExeResult.Eof then Exit For
%>
        <tr class="hback"> 
          <%
		For j = LBound(ArrayFieldName) To UBound(ArrayFieldName)
%>
          <td nowrap class="hback"> <div align="center"> 
              <%
			  if ArrayFieldName(j) = "form_lock" then
			  	if ExeResult(ArrayFieldName(j)) = 1 then Response.Write("��") else Response.Write("��")
			  else
				  if ArrayFieldName(j) = "form_answer" then
						if ExeResult(ArrayFieldName(j)) & "" = "" then Response.Write("��") else Response.Write("<font color=""red"">��</font>")
					else
						Response.Write(ExeResult(ArrayFieldName(j)) & "")
					end if
			  end if
		 %>
            </div></td>
          <% next %>
          <td nowrap class="hback"> <div align="center"><a href="javascript:void(0);" onClick="location='FormAnswer.asp?FormData=<% = ExeResult("ID") %>&formid=<% = formid %>';">�鿴���ظ�</a>&nbsp;<input name="FormData" type="checkbox" value="<% = ExeResult("ID") %>"> </div></td>
        </tr>
        <%
		ExeResult.MoveNext
	Next
%>
  <tr>
    <td colspan="<% = ExeResult.Fields.Count + 1%>"class="hback">
	  <table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td width="79%" colspan="2" align="right" class="hback"> <%
		  Response.Write "<p>"&  fPageCount(ExeResult,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
	%> </td>
        </tr>
      </table>
	 </td>
  </tr>
  <tr>
    <td height="30" colspan="<% = ExeResult.Fields.Count + 1%>"class="hback">
	<input type="submit" name="" value=" ɾ �� " onClick="return Del(this.form);">
	&nbsp;
	<input type="submit" name="" value="ɾ��ȫ��" onClick="return DelAll(this.form);">
      &nbsp;
      <input type="submit" name="Input" value=" �� �� " onClick="return Lock(this.form);">
      &nbsp;
      <input type="submit" name="Input2" value=" �� �� " onClick="return UNLock(this.form);">
      &nbsp;
      <input type="submit" style="display:none;" name="Input" value=" �� �� " onClick="return Answer(this.form);">
      &nbsp;
      <input type="submit" style="display:none;" name="Input2" value="�ظ�ȫ��" onClick="return AnswerAll(this.form);">
      <input type="hidden" name="Action">	 <input type="hidden" name="formid" value="<% = formid %>">
      <input type="hidden" name="Page" value="<% = cPageNo %>"></td>
  </tr>
<%
end if
%>
  </form>
  </table>
</body>
</html>
<%
Set Conn = Nothing
Set ExeResult = Nothing
Function RemovePlace(Str)
	Dim TempArray,Tempi,TempStr
	RemovePlace = ""
	TempArray = Split(Str,",")
	For Tempi = LBound(TempArray) To UBound(TempArray)
		TempStr = Trim(TempArray(Tempi))
		if TempStr <> "" then
			if RemovePlace = "" then
				RemovePlace = TempStr
			else
				RemovePlace = RemovePlace & "," & TempStr
			end if
		end if
	Next
End Function
%>
<script language="javascript">
function SubmitData(thsForm,Str,actionstr,formaction){
	if(confirm(Str)){
		if(formaction)thsForm.action=formaction;
		thsForm.Action.value=actionstr;
		return true;
	}
	return false;
}
function IsSelected(thsForm){
	if(thsForm.FormData.length){
		for(var i=0;i<thsForm.FormData.length;i++){
			if(thsForm.FormData[i].checked)return true;
		}
	}
	return false;
}
function Lock(thsForm){if(IsSelected(thsForm)){return SubmitData(thsForm,'ȷ��Ҫ������','Lock');}else{alert('û��ѡ�����������');return false;}}
function UNLock(thsForm){if(IsSelected(thsForm)){return SubmitData(thsForm,'ȷ��Ҫ������','UNLock');}else{alert('û��ѡ�����������');return false;}}
function Del(thsForm){if(IsSelected(thsForm)){return SubmitData(thsForm,'ȷ��Ҫɾ����','Del');}else{alert('û��ѡ�����������');return false;}}
function DelAll(thsForm){return SubmitData(thsForm,'ȷ��Ҫɾ��ȫ����','DelAll');}
function Answer(thsForm){if(IsSelected(thsForm)){return SubmitData(thsForm,'ȷ��Ҫ�ظ���','Answer','FormAnswer.asp');}else{alert('û��ѡ�����������');return false;}}
function AnswerAll(thsForm){return SubmitData(thsForm,'ȷ��Ҫ�ظ�ȫ����','AnswerAll','FormAnswer.asp');}
</script>
<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
'on error resume next
Dim Conn,User_Conn,rs_down_obj
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,i
int_RPP=20 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"				'βҳ

MF_Default_Conn
MF_User_Conn
MF_Session_TF

if not MF_Check_Pop_TF("ME_Mproducts") then Err_Show 
if not MF_Check_Pop_TF("ME032") then Err_Show 

dim str_use,str_lock,str_UserDel,deluser,strShowErr,tmp_lock
if Request.Form("Action")="Del" then
	if trim(Request.Form("did"))="" then
		strShowErr = "<li>��ѡ��һ����¼</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		User_Conn.execute("Delete From FS_ME_getThing where Id in ("& FormatIntArr(Request.Form("did")) &")")
		strShowErr = "<li>ɾ���ɹ�</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Get_Thing.asp")
		Response.end
	end if
end if
if Request.Form("Action")="Lock" then
	if trim(Request.Form("did"))="" then
		strShowErr = "<li>��ѡ��һ����¼</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		User_Conn.execute("Update FS_ME_getThing set isLock=1 where Id in ("& FormatIntArr(Request.Form("did")) &")")
		strShowErr = "<li>�����ɹ�</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Get_Thing.asp")
		Response.end
	end if
end if
if Request.Form("Action")="UnLock" then
	if trim(Request.Form("did"))="" then
		strShowErr = "<li>��ѡ��һ����¼</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		User_Conn.execute("Update FS_ME_getThing set isLock=0 where Id in ("& FormatIntArr(Request.Form("did")) &")")
		strShowErr = "<li>�����ɹ�</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Get_Thing.asp")
		Response.end
	end if
end if
Function GetFriendName(f_strNumber)
	Dim RsGetFriendName
	Set RsGetFriendName = User_Conn.Execute("Select UserName From FS_ME_Users Where UserNumber = '"& f_strNumber &"'")
	If  Not RsGetFriendName.eof  Then 
		GetFriendName = RsGetFriendName("UserName")
	Else
		GetFriendName = 0
	End If 
	set RsGetFriendName = nothing
End Function 
if Request.QueryString("Use")="1" then
	str_use = " and isUse=1"
elseif Request.QueryString("Use")="0" then
	str_use = " and isUse=0"
else
	str_use = ""
end if	
if Request.QueryString("isLock")="1" then
	str_lock = " and isLock=1"
elseif Request.QueryString("isLock")="0" then
	str_lock = " and isLock=0"
else
	str_lock = ""
end if	
if Request.QueryString("UserDel")="1" then
	str_UserDel = " and UserDel=1"
elseif Request.QueryString("UserDel")="0" then
	str_UserDel = " and UserDel=0"
else
	str_UserDel = ""
end if	
Set rs_down_obj=Server.CreateObject(G_FS_RS)
rs_down_obj.open "select * from FS_ME_GetThing where 1=1 "&str_use&str_lock&str_UserDel&" order by id desc",User_Conn,1,1

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="lib/UserJS.js" type="text/JavaScript"></script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes >
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr>
    <td class="xingmu">��Ա��Ʒ</td>
  </tr>
  <tr>
    <td class="hback"><a href="Get_Thing.asp">����</a>&nbsp;|&nbsp;<a href="Get_Thing.asp?Use=1">��ʹ��</a> | <a href="Get_Thing.asp?Use=0">δʹ��</a> | <a href="Get_Thing.asp?UserDel=1" onClick="history.back()">��Ա��ɾ��</a> | <a href="Get_Thing.asp?isLock=1" onClick="history.back()">������</a> | <a href="Get_Thing.asp?isLock=0">δ����</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="form1" method="post" action="">
    <tr class="xingmu">
      <td width="27%" align="center" class="xingmu">��Ʒ</td>
      <td width="19%" align="center" class="xingmu">�汾�ţ��ͺ�</td>
      <td width="8%" align="center" class="xingmu">��ʹ��</td>
      <td width="12%" align="center" class="xingmu">�������</td>
      <td width="12%" align="center" class="xingmu">�û���</td>
      <td width="17%" align="center" class="xingmu">�������ʱ��</td>
      <td width="5%" align="center" class="xingmu"><input name="chkall" type="checkbox" id="chkall" onClick="CheckAll(this.form);" value="checkbox" title="���ѡ�����л��߳�������ѡ��"></td>
    </tr>
    <%
	if rs_down_obj.eof then
		Response.Write"<tr  class=""hback""><td colspan=""7""  class=""hback"" height=""40"">û�м�¼��</td></tr>"
	else
			rs_down_obj.PageSize=int_RPP
			cPageNo=NoSqlHack(Request.QueryString("Page"))
			If cPageNo="" Then cPageNo = 1
			If not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo>rs_down_obj.PageCount Then cPageNo=rs_down_obj.PageCount 
			If cPageNo<=0 Then cPageNo=1
			rs_down_obj.AbsolutePage=cPageNo
			for i=0 to int_RPP
				if rs_down_obj.eof then exit for
				if rs_down_obj("isLock")=1 then
					tmp_lock = "<span class=""tx"">����</span>��"
				else
					tmp_lock = "<span>���ũ�"
				end if
				Response.Write("<tr class='hback'>"&chr(10)&chr(13)) 
				Response.Write("<td>"& tmp_lock &"<a href='#' id=item$pval[CatID]) style=""CURSOR: hand""  onmouseup=""opencat(down_"&rs_down_obj("id")&");"" onMouseOver=""this.className='bg'"" onMouseOut=""this.className='bg1'"" language=javascript>"&rs_down_obj("ProductID")&"</a></td>"&chr(10)&chr(13))
				Response.Write("<td>"&rs_down_obj("version")&"��"&rs_down_obj("Ptype")&"</td>")
				if rs_down_obj("isuse")=1 then
					Response.Write("<td><span class=""tx"">�Ѿ�����</span></td>")
				else
					Response.Write("<td>��û����</td>")
				end if
				Response.Write("<td align='center'>"&rs_down_obj("maxNUM")&"</td>")
				Response.Write("<td align='center'><a href=""../../"& G_USER_DIR &"/ShowUser.asp?UserNumber="&rs_down_obj("UserNumber")&""" target=""_blank"">"&GetFriendName(rs_down_obj("UserNumber"))&"</a></td>")
				Response.Write("<td align='center'>"&rs_down_obj("UpdateTime")&"</td>")
				Response.Write("<td align='center'><input type='checkbox' name='did' value='"&rs_down_obj("id")&"'></td>")
				Response.Write("</tr>")
				if rs_down_obj("UserDel")=1 then
					deluser = "��"
				else
					deluser = "��"
				end if
			   	Response.Write"<tr  class=""hback_1"" id=""down_"&rs_down_obj("id")&""" style=""display:none""><td colspan=""7""  class=""hback_1"" height=""50"">������ڣ�"&rs_down_obj("addtime")&"&nbsp;&nbsp;&nbsp;�������IP:"&rs_down_obj("Ip")&"&nbsp;&nbsp;&nbsp;�û��Ƿ�ɾ����"& deluser &"&nbsp;&nbsp;&nbsp;&nbsp;<a href=""Get_Thing_Edit.asp?Action=Edit&Id="&Rs_down_obj("id")&"""><span class=""tx"">�޸Ĵ˼�¼</span></a><br>���ص�ַ��"&rs_down_obj("URL_1")&"<BR>������"&rs_down_obj("Content")&"</td></tr>"
				rs_down_obj.movenext
			next
		end if%>
    <tr>
      <td align="right" colspan="7" class="hback"><%
	response.Write fPageCount(rs_down_obj,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
	rs_down_obj.close:set rs_down_obj=nothing
	%>
        <input name="AddAward" type="button" value="�� ��" onClick="location='Get_Thing_Edit.asp?Action=Add'">
        <input name="Action" type="hidden" id="Action">
        <input type="button" name="Submit" value="ɾ��"  onClick="document.form1.Action.value='Del';{if(confirm('ȷ���������ѡ��ļ�¼��')){this.document.form1.submit();return true;}return false;}">
        <input type="button" name="Submit2" value="����"  onClick="document.form1.Action.value='Lock';{if(confirm('ȷ��������')){this.document.form1.submit();return true;}return false;}">
        <input type="button" name="Submit22" value="����"  onClick="document.form1.Action.value='UnLock';{if(confirm('ȷ�Ͻ�����')){this.document.form1.submit();return true;}return false;}"></td>
    </tr>
  </form>
</table>
</body>
<script language="JavaScript" type="text/JavaScript">
function opencat(cat)
{
  if(cat.style.display=="none"){
     cat.style.display="";
  } else {
     cat.style.display="none"; 
  }
}
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = form1.elements[i];  
    if (e.name != 'chkall')  
       e.checked = form1.chkall.checked;  
    }  
	}
function AlertBeforeSubmite()
{
	var checkGroup=document.Thing_Form.DeleteAwards;
	var flag=false;
	for(var i=0;i<checkGroup.length;i++)
	{
		if(checkGroup[i].checked)
		{
			flag=true;
		}
	}
	if(flag)
	{
		if(confirm("ȷ��Ҫɾ���ü�¼?"))
		{
			document.Thing_Form.submit();
		}
	}else
	{
		alert("��ѡ��Ҫɾ���ļ�¼")
	}
}
</script>
</html>







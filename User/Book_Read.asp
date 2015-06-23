<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim P_ToUserNumber,P_strToUserNumber,GetUsermessObj
dim FS_Book,str_m_type
str_m_type = CintStr(Request.QueryString("M_type"))
if isnull(str_m_type) or not isnumeric(str_m_type) or trim(str_m_type)="" then
	strShowErr = "<li>�������</li>"
	Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
end if
Set FS_Book = new Cls_message
P_ToUserNumber = NoSqlHack(Request.QueryString("ToUserNumber"))
'����û����ͱ��
Set GetUsermessObj = server.CreateObject(G_FS_RS)
GetUsermessObj.open "select  UserID,isLock,UserName,GroupID,UserNumber From FS_ME_Users where UserNumber = '"& NoSqlHack(P_ToUserNumber) &"'",User_Conn,1,3
if GetUsermessObj.eof then
	P_strToUserNumber = ""
Else
	P_strToUserNumber = GetUsermessObj("UserName")
End if

If Request.Form("Action") = "Save" then
	Dim p_M_ReadUserName,p_M_title,p_M_Content,p_M_Type
	p_M_ReadUserName = NoSqlHack(Request.Form("M_ReadUserNumber"))
	p_M_title = NoSqlHack(Request.Form("Title"))
	p_M_Content = NoSqlHack(NoHtmlHackInput(Request.Form("Content")))
	p_M_Type= CintStr(Request.Form("M_Type"))
	If p_M_ReadUserName="" Or p_M_title="" Or p_M_Content="" Then
		strShowErr = "<li>����д����</li><li>���Ա��⡢�ռ��ˡ���Ϣ���ݲ���Ϊ��</li>"
		Call ReturnError(strShowErr,"")
	End If
	If len(p_M_Content)>1000 Then
		strShowErr = "<li>�������ݲ��ܳ���1000���ַ�</li>"
		Call ReturnError(strShowErr,"")
	End If
	If Trim(p_M_ReadUserName)=Fs_User.UserName Then
		strShowErr = "<li>�����Լ����Լ���������</li>"
		Call ReturnError(strShowErr,"")
	End if
	Dim Returvaluestr
	Returvaluestr = Fs_User.GetFriendNumber(p_M_ReadUserName)
	Dim t_RsCheckFriend
	if Returvaluestr ="0" then
		strShowErr = "<li>�Ҳ�����Ա��Ϣ�����������͵Ļ�Ա�Ѿ�ɾ��</li>"
		Call ReturnError(strShowErr,"")
	Else
		Set t_RsCheckFriend = User_Conn.Execute("select FriendType from FS_ME_Friends where UserNumber='"&NoSqlHack(Returvaluestr)&"' and F_UserNumber='"&Fs_User.UserNumber&"' and FriendType=2")
		If Not t_RsCheckFriend.EOF Then 
			strShowErr = "<li>�Է��ѽ�������������������ٸ���������Ϣ</li>"
			Call ReturnError(strShowErr,"")
		End If
	End if
	If Fs_User.UserExist(Returvaluestr)=False then
		strShowErr = "<li>û�д��û����ߴ��û��Ѿ�������</li>"
		Call ReturnError(strShowErr,"")
	End If
	Set t_RsCheckFriend = Nothing 
	If FS_Book.LenbContent(Returvaluestr)+Len(Request.Form("Content")) > 100*1024 then
		strShowErr = "<li>�Է����Կռ�������������֪ͨ�Է�ɾ����������</li>"
		Call ReturnError(strShowErr,"")
	End If
	dim save_rs
	set save_rs= Server.CreateObject(G_FS_RS)
	save_rs.open "select * From FS_ME_Book where 1=0",User_Conn,1,3
	save_rs.addnew
	save_rs("M_Title")=p_M_title
	save_rs("M_ReadUserNumber")=Returvaluestr
	save_rs("M_Content")=p_M_Content
	save_rs("M_FromUserNumber")=Fs_User.UserNumber
	save_rs("M_FromDate")=now
	save_rs("M_ReadTF")=0
	save_rs("LenContent")=len(p_M_Content)
	save_rs("M_Type")=p_M_Type
	save_rs.update
	save_rs.close:set save_rs = nothing
	Set FS_Book = Nothing 
	strShowErr = "<li>��ϲ��</li><li>���ͳɹ�</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
Else
	Dim p_BookID,RsRBookObj,str_m_title,str_M_FromUserNumber,str_M_ReadUserNumber,str_M_Content,str_M_FromDate
	p_BookID = NoSqlHack(Request.QueryString("BookID"))
	if isNumeric(p_BookID) = false then
			strShowErr = "<li>��������</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	End if
	'��������
	Set RsRBookObj = server.CreateObject(G_FS_RS)
	RsRBookObj.open "select  BookID,M_Title,M_FromUserNumber,M_ReadUserNumber,M_Content,M_FromDate,M_ReadTF,LenContent From FS_ME_Book where BookID = "& CintStr(p_BookID) ,User_Conn,1,3
	if RsRBookObj.eof then
		strShowErr = "<li>�Ҳ�����¼</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
		RsRBookObj("M_ReadTF")=1
		RsRBookObj.Update
		str_m_title = "RE:"&RsRBookObj("M_title")
		str_M_FromUserNumber = Fs_User.GetFriendName(RsRBookObj("M_FromUserNumber"))
		str_M_ReadUserNumber = RsRBookObj("M_ReadUserNumber")
		str_M_Content = vbCrLf&"---------"&str_M_FromUserNumber &"��"& RsRBookObj("M_FromDate") &"˵:--------"&vbCrLf&""&RsRBookObj("M_Content")
		str_M_FromDate = RsRBookObj("M_FromDate")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-����</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<body>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td class="hback"><strong>λ�ã�</strong><a href="../">��վ��ҳ</a> &gt;&gt; 
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; <a href="book.asp?M_Type=<%=str_m_type%>">���Թ���</a>&gt;&gt; �ظ�����&lt;&lt; 
            <%
			select case Request.QueryString("M_type")
					case "0"
						Response.Write("��Ա����")
					case "1"
						Response.Write("��������")
					case "2"
						Response.Write("��������")
					case "3"
						Response.Write("��ְ��Ƹ����")
					case "4"
						Response.Write("��������")
					case "5"
						Response.Write("��������")
			end select
			%>
          &gt;&gt; </td>
        </tr>
      </table> 
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="UserForm" method="post" action=""  onsubmit="return CheckForm();">
          <tr class="hback"> 
            <td width="16%" class="hback_1"><div align="center"><strong>�ռ���</strong></div></td>
            <td class="hback"> <div align="left"> 
                <input name="M_ReadUserNumber" type="text" id="M_ReadUserNumber" size="20" value="<% = str_M_FromUserNumber%>">
                <font color="#999999"> 
                <select name="SelectFriend" id="SelectFriend" onChange="DoTitle(this.options[this.selectedIndex].value)">
                  <option selected value="">>>ѡ�����<<</option>
                  <%=Fs_User.FriendList%> 
                </select>
                </font>����д�û���<strong>��<a href="Friend_add.asp">��Ӻ���</a></strong></div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>���Ա���</strong></div></td>
            <td class="hback"> <div align="left"> 
                <input name="Title" type="text" id="Title" value="<% = str_m_title %>" size="40" maxlength="50">
              </div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>��������</strong></div></td>
            <td class="hback"> <div align="left"> 
                <textarea name="Content"  style="width:80%" rows="15" id="Content"><% = str_M_Content %></textarea>���1000���ַ�
              </div></td>
          </tr>
          <tr class="hback"> 
            <td colspan="2" class="hback"> <div align="left">������������������
                <input name="M_Type" type="hidden" id="M_Type" value="<%=str_m_type%>">
                <input name="Action" type="hidden" id="Action" value="Save">
                <input type="submit" name="Submit" value=" ȷ������ ">
                �� 
                <input type="reset" name="Submit3" value="������д">
              </div></td>
          </tr>
          <tr class="hback"> 
            <td colspan="2" class="hback"> <div align="center"> </div></td>
          </tr>
        </form>
      </table>
      </td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="Copyright.asp" -->
        </div></td>
    </tr>
 
</table>
</body>
</html>
<script language="JavaScript" type="text/javascript">
function CheckForm()
{
	if(document.UserForm.M_ReadUserNumber.value=="")
	{
		alert("����д�ռ���!");
		document.UserForm.M_ReadUserNumber.focus();
		return false;
	}
	if(document.UserForm.Title.value=="")
	{
		alert("����д���Ա���!");
		document.UserForm.Title.focus();
		return false;
	}
	if(document.UserForm.Content.value=="")
	{
		alert("����д��������!");
		document.UserForm.Content.focus();
		return false;
	}
	}
</script>
<script language="JavaScript" type="text/JavaScript">
function DoTitle(addTitle) {  
document.UserForm.M_ReadUserNumber.value=document.UserForm.SelectFriend.value;  
document.UserForm.M_ReadUserNumber.focus(); 
 return; 
} 
</script>
<%
	End if
End if
RsRBookObj.close:set RsRBookObj = nothing
set FS_Book = nothing
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->






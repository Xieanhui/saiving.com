<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
if Request("Action") = "del" then
	Dim strFriendID,Str_type
	Str_type = Request.QueryString("Types")
	If Str_type = "" Then
		strShowErr = "<li>����Ĳ���</li>"
		Call ReturnError(strShowErr,"")
	Else
		Str_type = CintStr(Str_type)
	End If	
	strFriendID = NoSqlHack(Request.QueryString("FriendID"))
	If strFriendID = ""  or isNumeric(strFriendID)=false then
		strShowErr = "<li>����Ĳ���</li>"
		Call ReturnError(strShowErr,"")
	Else
		User_Conn.execute("Delete From FS_ME_Friends where FriendType = " & CintStr(Str_type) & " And UserNumber = '" & Fs_User.UserNumber & "' And FriendID="&CintStr(strFriendID))
		strShowErr = "<li>ɾ���ɹ���</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-�ҵ�����</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt;
			<%
			if Request("type") = 0 then
				Response.Write("�ҵĺ���")
			Elseif Request("type") = 1 then
				Response.Write("�ҵ�İ����")
			Elseif Request("type") = 2 then
				Response.Write("�ҵĺ�����")
			Else
				Response.Write("�������")
			End if
			%>
			</td>
          </tr>
        </table>
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td colspan="4" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="44%"> ��������<strong> 
                  <%
				Dim RsUserFriendObj,RsUserFriendSQL
				Dim strpage,strSQLs
				strpage=CintStr(request("page"))
				if len(strpage)=0 Or strpage<1 or trim(strpage)=""  Then strpage="1"
				Set RsUserFriendObj = Server.CreateObject(G_FS_RS)
				If cint(Request("type"))=0 or  trim(Request("type"))="" then
						  strSQLs = " and FriendType=0 " 
				Elseif cint(Request("type"))=1 then
						  strSQLs = " and FriendType=1 " 
				Elseif cint(Request("type"))=2 then
						  strSQLs = " and FriendType=2 " 
				Else
						  strSQLs = " and FriendType=0 " 
				End if
				RsUserFriendSQL = "Select FriendID,UserNumber,FriendType,F_UserNumber,AddTime,Updatetime,RealName,Content,Email,Tel,Mobile,QQ,MSN From FS_ME_Friends  where UserNumber='"&Fs_User.UserNumber&"' "& strSQLs &" Order by FriendID desc"
				RsUserFriendObj.Open RsUserFriendSQL,User_Conn,1,3
				Response.Write "<Font color=red>" & RsUserFriendObj.RecordCount&"</font>"
				%>
                  </strong> ��
				  <%
			if Request("type") = 0 then
				Response.Write("����")
			Elseif Request("type") = 1 then
				Response.Write("İ����")
			Elseif Request("type") = 2 then
				Response.Write("������")
			Else
				Response.Write("�������")
			End if
			%></td>
                <form action="Friend.asp"  method="post" name="myform" id="myform">
                  <td width="56%"><div align="center">
				  <%if request.QueryString("type") = 1 then%>
				  <a href="Friend_add.asp?type=1"><strong>���İ����</strong></a>
				  <%Elseif request.QueryString("type") = 2 then%>
				  <a href="Friend_add.asp?type=2"><strong>��Ӻ�����</strong></a>
				  <%Else%>
				  <a href="Friend_add.asp?type=2"><strong>��Ӻ���</strong></a>
				  <%End if%></div></td>
                </form>
              </tr>
            </table></td>
        </tr >
		<tr class="hback">
          <td class="xingmu">�û����</td>
          <td class="xingmu"><strong>�û���</strong></td>
          <td class="xingmu">����</td>
          <td class="xingmu"><div align="center">��ע</div></td>
        </tr>
		<%
		Dim select_count,select_pagecount,i
		if RsUserFriendObj.eof then
			   RsUserFriendObj.close
			   set RsUserFriendObj=nothing
			   Response.Write"<TR><TD colspan=""4""  class=""hback"">û�м�¼��</TD></TR>"
		else
				RsUserFriendObj.pagesize = 15
				RsUserFriendObj.absolutepage=cint(strpage)
				select_count=RsUserFriendObj.recordcount
				select_pagecount=RsUserFriendObj.pagecount
				for i=1 to RsUserFriendObj.pagesize
					if RsUserFriendObj.eof Then exit For 
						Dim Returvaluestr
						Returvaluestr = Fs_User.GetFriendName(RsUserFriendObj("F_UserNumber"))
					if RsUserFriendObj("F_UserNumber") = "0" then
						  exit For 
					Else
		%>
        
        <tr class="hback"> 
          <td width="18%" class="hback"><a href="ShowUser.asp?UserNumber=<% = RsUserFriendObj("F_UserNumber")%>"> 
            <% = RsUserFriendObj("F_UserNumber")%>
            </a></td>
          <td width="35%" class="hback"><div align="left"><a href="ShowUser.asp?UserNumber=<% = RsUserFriendObj("F_UserNumber")%>">
<% = Returvaluestr%>
              </a></div></td>
          <td width="33%" class="hback"> <div align="left"> <a href="message_write.asp?ToUserNumber=<% = RsUserFriendObj("F_UserNumber")%>">����</a>��<a href="book_write.asp?ToUserNumber=<% = RsUserFriendObj("F_UserNumber")%>&M_Type=0">����</a>��<a href="Friend_add.asp?FriendID=<% = RsUserFriendObj("FriendID")%>">�޸�</a>��<a href="Friend.asp?Action=del&FriendID=<% = RsUserFriendObj("FriendID")%>&Types=<% = request.QueryString("type") %>"  onClick="{if(confirm('ȷ��Ҫɾ����ѡ������Ŀ��?')){this.document.inbox.submit();return true;}return false;}">ɾ��</a></div></td>
          <td width="14%" class="hback"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sid<%=RsUserFriendObj("FriendID")%>);"  language=javascript><div align="center">�鿴��ע</div></td>
        </tr>
        <tr class="hback" style="display:none" id="sid<%=RsUserFriendObj("FriendID")%>"> 
          <td height="46" colspan="4" class="hback"><table width="100%" border="0" cellspacing="1" cellpadding="5" class="table">
              <tr> 
                <td width="23%" class="hback_1">������ <% = RsUserFriendObj("RealName")%></td>
                <td width="22%" class="hback_1">�绰�� <% = RsUserFriendObj("Tel")%> </td>
                <td width="23%" class="hback_1">�ֻ��� <% = RsUserFriendObj("Mobile")%> </td>
                <td width="32%" class="hback_1">Email: <a href="mailto:<% = RsUserFriendObj("Email")%>"><% = RsUserFriendObj("Email")%></a> </td>
              </tr>
              <tr> 
                <td class="hback_1">MSN�� <% = RsUserFriendObj("MSN")%> </td>
                <td class="hback_1">QQ�� <%
						if  Len(Trim(RsUserFriendObj("QQ")))>4 then
							Dim sOICQ
						    sOICQ ="<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin="& RsUserFriendObj("QQ") &"&Site=FoosunCMS&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:"& RsUserFriendObj("QQ") &":16 alt=""��������"& RsUserFriendObj("QQ") &"����Ϣ""></a>"
							Response.Write sOICQ
						Else
							Response.Write("û��")
						End if
						%> </td>
                <td colspan="2" class="hback_1">˵���� <% = RsUserFriendObj("Content")%> </td>
              </tr>
            </table></td>
        </tr>
        <%
				End if
			  RsUserFriendObj.MoveNext
		  Next
		  %>
        <tr class="hback"> 
          <td colspan="4" class="xingmu"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="80%"> <span class="top_navi"> 
                  <% 	Response.Write("ÿҳ:"& RsUserFriendObj.pagesize &"��,")
							Response.write"&nbsp;��<b>"& select_pagecount &"</b>ҳ<b>&nbsp;" & select_count &"</b>����¼����ҳ�ǵ�<b>"& strpage &"</b>ҳ��"
							if int(strpage)>1 then
								Response.Write"&nbsp;<a href=?page=1&type="&Request("type")&">��һҳ</a>&nbsp;&nbsp;"
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&type="&Request("type")&">��һҳ</a>&nbsp;&nbsp;"
							End if
							If int(strpage)<select_pagecount then
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&type="&Request("type")&">��һҳ</a>&nbsp;"
								Response.Write"&nbsp;<a href=?page="& select_pagecount &"&type="&Request("type")&">���һҳ</a>&nbsp;&nbsp;"
							End if
								Response.Write"<br>"
								RsUserFriendObj.close
								Set RsUserFriendObj=nothing
							End if
							%>
                  </SPAN></td>
              </tr>
            </table></td>
        </tr>
      </table> </td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="Copyright.asp" -->
        </div></td>
    </tr>
 
</table>
</body>
</html>
<%
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->
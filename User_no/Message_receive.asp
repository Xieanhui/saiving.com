<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-����-�ռ���</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; ���ţ��ռ���</td>
          </tr>
        </table>
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td colspan="4" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="27%"> ��������<strong> 
                  <%
				Dim RsUserFriendObj,RsUserFriendSQL
				Dim strpage,strSQLs
				strpage=request("page")
				if len(strpage)=0 Or strpage<1 or trim(strpage)=""  Then strpage="1"
				Set RsUserFriendObj = Server.CreateObject(G_FS_RS)
				If cint(Request("type"))=0 or  trim(Request("type"))="" then
						  strSQLs = " and FriendType=0 " 
				Elseif cint(Request("type"))=1 then
						  strSQLs = " and FriendType=1 " 
				Elseif cint(Request("type"))=2 then
						  strSQLs = " and FriendType=2 " 
				End if
				RsUserFriendSQL = "Select UserNumber,FriendType,F_UserNumber,AddTime,Updatetime From FS_ME_Friends  where UserNumber='"&Fs_User.UserNumber&"' "& strSQLs &" Order by FriendID desc"
				RsUserFriendObj.Open RsUserFriendSQL,User_Conn,1,3
				Response.Write "<Font color=red>" & RsUserFriendObj.RecordCount&"</font>"
				%>
                  </strong> ������</td>
                <form action="Message_receive.asp"  method="post" name="myform" id="myform">
                  <td width="7%"><div align="left">�ռ�ռ��</div></td>
                  <td width="66%">&nbsp;</td>
                </form>
              </tr>
            </table></td>
        </tr class="hback">
        <tr class="hback"> 
          <td width="34%" class="xingmu"><div align="left"><strong>�û����</strong></div></td>
          <td width="25%" class="xingmu"><div align="left"><strong>�û���</strong></div></td>
          <td width="41%" class="xingmu"><div align="left"><strong>����</strong></div></td>
        </tr>
        <%
		Dim select_count,select_pagecount,i
		if RsUserFriendObj.eof then
			   RsUserFriendObj.close
			   set RsUserFriendObj=nothing
			   Response.Write"<TR><TD colspan=""3""  class=""hback"">û�м�¼��</TD></TR>"
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
          <td class="hback"><div align="left">��<a href="ShowUser.asp?UserNumber=<% = RsUserFriendObj("F_UserNumber")%>"> 
              <% = RsUserFriendObj("F_UserNumber")%></a></div></td>
          <td class="hback"><div align="left"><a href="ShowUser.asp?UserNumber=<% = RsUserFriendObj("F_UserNumber")%>"> <% = Returvaluestr%></a></div></td>
          <td class="hback"><div align="left"> <a href="message_write.asp?ToUserNumber=<% = RsUserFriendObj("F_UserNumber")%>">����</a>��<a href="book_write.asp?ToUserNumber=<% = RsUserFriendObj("F_UserNumber")%>">����</a>��<a href="Friend.asp?Action=del&UserNumber=<% = RsUserFriendObj("F_UserNumber")%>">ɾ��</a>��<a href="Friend_Move.asp?UserNumber=<% = RsUserFriendObj("F_UserNumber")%>">ת��</a></div></td>
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
                  <% 
							Response.write"&nbsp;��<b>"& select_pagecount &"</b>ҳ<b>&nbsp;" & select_count &"</b>����¼����ҳ�ǵ�<b>"& strpage &"</b>ҳ��"
							if int(strpage)>1 then
								Response.Write"&nbsp;<a href=Callboard.asp?page=1&Keyword="&Request("Keyword")&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&">��һҳ</a>&nbsp;&nbsp;"
								Response.Write"&nbsp;<a href=Callboard.asp?page="&cstr(cint(strpage)-1)&"&Keyword="&Request("Keyword")&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&">��һҳ</a>&nbsp;&nbsp;"
							End if
							If int(strpage)<select_pagecount then
								Response.Write"&nbsp;<a href=Callboard.asp?page="&cstr(cint(strpage)+1)&"&Keyword="&Request("Keyword")&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&">��һҳ</a>&nbsp;"
								Response.Write"&nbsp;<a href=Callboard.asp?page="& select_pagecount &"&Keyword="&Request("Keyword")&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&">���һҳ</a>&nbsp;&nbsp;"
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






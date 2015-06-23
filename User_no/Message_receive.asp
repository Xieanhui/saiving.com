<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-短信-收件箱</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
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
            
          <td class="hback"><strong>位置：</strong><a href="../">网站首页</a> &gt;&gt; 
            <a href="main.asp">会员首页</a> &gt;&gt; 短信－收件箱</td>
          </tr>
        </table>
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td colspan="4" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="27%"> 共搜索到<strong> 
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
                  </strong> 个朋友</td>
                <form action="Message_receive.asp"  method="post" name="myform" id="myform">
                  <td width="7%"><div align="left">空间占用</div></td>
                  <td width="66%">&nbsp;</td>
                </form>
              </tr>
            </table></td>
        </tr class="hback">
        <tr class="hback"> 
          <td width="34%" class="xingmu"><div align="left"><strong>用户编号</strong></div></td>
          <td width="25%" class="xingmu"><div align="left"><strong>用户名</strong></div></td>
          <td width="41%" class="xingmu"><div align="left"><strong>操作</strong></div></td>
        </tr>
        <%
		Dim select_count,select_pagecount,i
		if RsUserFriendObj.eof then
			   RsUserFriendObj.close
			   set RsUserFriendObj=nothing
			   Response.Write"<TR><TD colspan=""3""  class=""hback"">没有记录。</TD></TR>"
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
          <td class="hback"><div align="left">·<a href="ShowUser.asp?UserNumber=<% = RsUserFriendObj("F_UserNumber")%>"> 
              <% = RsUserFriendObj("F_UserNumber")%></a></div></td>
          <td class="hback"><div align="left"><a href="ShowUser.asp?UserNumber=<% = RsUserFriendObj("F_UserNumber")%>"> <% = Returvaluestr%></a></div></td>
          <td class="hback"><div align="left"> <a href="message_write.asp?ToUserNumber=<% = RsUserFriendObj("F_UserNumber")%>">发信</a>｜<a href="book_write.asp?ToUserNumber=<% = RsUserFriendObj("F_UserNumber")%>">留言</a>｜<a href="Friend.asp?Action=del&UserNumber=<% = RsUserFriendObj("F_UserNumber")%>">删除</a>｜<a href="Friend_Move.asp?UserNumber=<% = RsUserFriendObj("F_UserNumber")%>">转移</a></div></td>
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
							Response.write"&nbsp;共<b>"& select_pagecount &"</b>页<b>&nbsp;" & select_count &"</b>条记录，本页是第<b>"& strpage &"</b>页。"
							if int(strpage)>1 then
								Response.Write"&nbsp;<a href=Callboard.asp?page=1&Keyword="&Request("Keyword")&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&">第一页</a>&nbsp;&nbsp;"
								Response.Write"&nbsp;<a href=Callboard.asp?page="&cstr(cint(strpage)-1)&"&Keyword="&Request("Keyword")&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&">上一页</a>&nbsp;&nbsp;"
							End if
							If int(strpage)<select_pagecount then
								Response.Write"&nbsp;<a href=Callboard.asp?page="&cstr(cint(strpage)+1)&"&Keyword="&Request("Keyword")&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&">下一页</a>&nbsp;"
								Response.Write"&nbsp;<a href=Callboard.asp?page="& select_pagecount &"&Keyword="&Request("Keyword")&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&">最后一页</a>&nbsp;&nbsp;"
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






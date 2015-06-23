<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-会员列表</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt; 会员列表统计</td>
          </tr>
        </table>
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td colspan="7" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="43%"> 共搜索到<strong> 
                  <%
				Dim RsUserListObj,RsUserSQL
				Dim strpage,strSQLs,StrOrders
				strpage=CintStr(request("page"))
				if len(strpage)=0 Or strpage<1 or trim(strpage)=""  Then strpage="1"
				Set RsUserListObj = Server.CreateObject(G_FS_RS)
				if Request("RegTime") = "0" then
					StrOrders = " order by RegTime Desc"
				Elseif Request("RegTime")= "1" then
					StrOrders = " order by RegTime asc"
				Else
					StrOrders = " order by UserID Desc"
				End If
				Dim Keyword
				Keyword=NoSqlHack(Request("Keyword"))
				If Keyword<>"" then
						if Request("searchtype") <>"" then
								if  Request("Name") = "UserName" then
									  strSQLs = " and UserName like '%" & Keyword& "%' "& StrOrders &""
								Elseif  Request("Name") = "UserNumber" then
									  strSQLs = " and UserNumber  like '%" & Keyword& "%' "& StrOrders &""
								Elseif  Request("Name") = "NickName" then
									  strSQLs = " and NickName  like '%" & Keyword& "%' "& StrOrders &""
								Elseif  Request("Name") = "RealName" then
									  strSQLs = " and RealName  like '%" & Keyword& "%' "& StrOrders &""
								Elseif  Request("Name") = "Email" then
									  strSQLs = " and Email  like '%" & Keyword& "%' "& StrOrders &""
								Elseif  Request("Name") = "QQ" then
									  strSQLs = " and QQ  like '%" & Keyword& "%' "& StrOrders &""
								Elseif  Request("Name") = "MSN" then
									  strSQLs = " and MSN  like '%" & Keyword& "%' "& StrOrders &""
								Elseif  Request("Name") = "Integral" then
									  strSQLs = " and Integral <"& Keyword &"+50 and Integral>"& Keyword &"-50 "& StrOrders &""
								Elseif  Request("Name") = "Province" then
									  strSQLs = " and Province  like '%" & Keyword& "%' "& StrOrders &""
								Elseif  Request("Name") = "city" then
									  strSQLs = " and city  like '%" & Keyword& "%' "& StrOrders &""
								End if
						Else
								if  Request("Name") = "UserName" then
									  strSQLs = " and UserName = '" & Keyword& "' "& StrOrders &""
								Elseif  Request("Name") = "UserNumber" then
									  strSQLs = " and UserNumber  = '" & Keyword& "' "& StrOrders &""
								Elseif  Request("Name") = "NickName" then
									  strSQLs = " and NickName  = '" & Keyword& "' "& StrOrders &""
								Elseif  Request("Name") = "RealName" then
									  strSQLs = " and RealName  = '" & Keyword& "' "& StrOrders &""
								Elseif  Request("Name") = "Email" then
									  strSQLs = " and Email  = '" & Keyword& "' "& StrOrders &""
								Elseif  Request("Name") = "QQ" then
									  strSQLs = " and QQ  = '" & Keyword& "' "& StrOrders &""
								Elseif  Request("Name") = "MSN" then
									  strSQLs = " and MSN  = '" & Keyword& "' "& StrOrders &""
								Elseif  Request("Name") = "Integral" then
									  strSQLs = " and Integral =" & clng(Keyword)& " "& StrOrders &""
								Elseif  Request("Name") = "Province" then
									  strSQLs = " and Province ='" & Keyword& "' "& StrOrders &""
								Elseif  Request("Name") = "city" then
									  strSQLs = " and city ='" & Keyword& "' "& StrOrders &""
								End if
						End if
				Else
						strSQLs = " "& StrOrders &""
				End if
				RsUserSQL = "Select UserID,UserName,UserNumber,RealName,Email,QQ,MSN,homepage,Integral,isLock,RegTime,Province,city From Fs_ME_Users  where isLock=0 "& strSQLs &""
				'Response.Write(RsUserSQL)
				'Response.end
				RsUserListObj.Open RsUserSQL,User_Conn,1,1
				Response.Write "<Font color=red>" & RsUserListObj.RecordCount&"</font>"
				%>
                  </strong> 个会员</td>
                <form action="UserList.asp"  method="post" name="myform" id="myform">
                  <td width="57%"><div align="left">搜索： 
                      <select name="Name" id="select">
                        <option value="UserName" <%if Request("Name")="UserName" then response.Write("selected")%>>用户名</option>
                        <option value="UserNumber" <%if Request("Name")="UserNumber" then response.Write("selected")%>>用户编号</option>
                        <option value="NickName" <%if Request("Name")="NickName" then response.Write("selected")%>>昵称</option>
                        <option value="RealName" <%if Request("Name")="RealName" then response.Write("selected")%>>姓名</option>
                        <option value="Email" <%if Request("Name")="Email" then response.Write("selected")%>>电子邮件</option>
                        <option value="QQ" <%if Request("Name")="QQ" then response.Write("selected")%>>OICQ</option>
                        <option value="MSN" <%if Request("Name")="MSN" then response.Write("selected")%>>MSN</option>
                        <option value="Integral" <%if Request("Name")="Integral" then response.Write("selected")%>>50-积分+50</option>
                        <option value="Province" <%if Request("Name")="Province" then response.Write("selected")%>>省份</option>
                        <option value="city" <%if Request("Name")="city" then response.Write("selected")%>>城市</option>
                      </select>
                      <input name="keyword" type="text" id="keyword2" value="<%=Keyword%>" size="10">
                      <input name="searchtype" type="checkbox" id="searchtype" value="1" <%if Request("searchtype")="1" then Response.Write("checked")%> >
                      模糊搜索 
                      <input type="submit" name="Submit" value="搜索">
                    </div></td>
                </form>
              </tr>
            </table></td>
        </tr class="hback">
        <tr class="hback"> 
          <td width="17%" class="xingmu"><div align="left"><strong>
		  <%If Request("RegTime") <> "" then
		  		If cint(Request("RegTime"))=1 then%>
		  <a href="UserList.asp?page=<%=strpage%>&Keyword=<%=Keyword%>&Name=<%=NoSqlHack(Request("Name"))%>&searchtype=<%=NoSqlHack(Request("searchtype"))%>&RegTime=0&CountPage=<%=NoSqlHack(Request("CountPage"))%>&Integral=" class="LinkCss">用户名</a>
		  		<%Elseif  cint(Request("RegTime"))=0 then %>
		  <a href="UserList.asp?page=<%=strpage%>&Keyword=<%=Keyword%>&Name=<%=NoSqlHack(Request("Name"))%>&searchtype=<%=NoSqlHack(Request("searchtype"))%>&RegTime=1&CountPage=<%=NoSqlHack(Request("CountPage"))%>&Integral=" class="LinkCss">用户名</a>
		  		<%End if%>
			<% Else	%>
				<a href="UserList.asp?page=<%=strpage%>&Keyword=<%=Keyword%>&Name=<%=NoSqlHack(Request("Name"))%>&searchtype=<%=NoSqlHack(Request("searchtype"))%>&RegTime=0&CountPage=<%=NoSqlHack(Request("CountPage"))%>&Integral=" class="LinkCss">用户名</a>
			<% End If %>
		  </strong></div></td>
          <td width="18%" class="xingmu"><div align="left"><strong>编号</strong></div></td>
          <td width="13%" class="xingmu"><div align="left"><strong>OICQ</strong></div></td>
          <td width="9%" class="xingmu"><div align="center"><strong>Email</strong></div></td>
          <td width="7%" class="xingmu"><div align="center"><strong>主页</strong></div></td>
          <td width="10%" class="xingmu"><div align="center"><strong>积分</strong></div></td>
          <td width="26%" class="xingmu"><div align="center"><strong>操作</strong></div></td>
        </tr>
        <%
		Dim select_count,select_pagecount,i
		if RsUserListObj.eof then
			   RsUserListObj.close
			   set RsUserListObj=nothing
			   Response.Write"<TR><TD colspan=""7""  class=""hback"">没有记录。</TD></TR>"
		else
				if Request("CountPage")="" or len(Request("CountPage"))<1 then
					RsUserListObj.pagesize = 20
				Else
					RsUserListObj.pagesize = CintStr(Request("CountPage"))
				End if
				RsUserListObj.absolutepage=CintStr(strpage)
				select_count=RsUserListObj.recordcount
				select_pagecount=RsUserListObj.pagecount
				for i=1 to RsUserListObj.pagesize
					if RsUserListObj.eof Then exit For 
		%>
        <tr  onMouseOver=overColor(this) onMouseOut=outColor(this)>
          <td class="hback"><div align="left"><a href="ShowUser.asp?UserNumber=<% = RsUserListObj("UserNumber")%>"> 
              <% = RsUserListObj("UserName")%>
              </a></div></td>
          <td class="hback"><div align="left"> 
              <% = RsUserListObj("UserNumber")%>
            </div></td>
          <td class="hback"><div align="left"> 
              <%
						if  Len(Trim(RsUserListObj("QQ")))>4 then
							Dim sOICQ
						    sOICQ ="<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin="& RsUserListObj("QQ") &"&Site=FoosunCMS&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:"& RsUserListObj("QQ") &":16 alt=""点击这里给"& RsUserListObj("QQ") &"发消息""></a>"
							Response.Write sOICQ
						Else
							Response.Write("没有")
						End if
						%>
            </div></td>
          <td class="hback"><div align="center"><a href="mailto:<% = RsUserListObj("Email")%>">发信</a></div></td>
          <td class="hback"><div align="center"><a href="<% = RsUserListObj("homepage")%>">主页</a></div></td>
          <td class="hback"><div align="center">
              <% = RsUserListObj("Integral")%>
            </div></td>
          <td class="hback"><div align="center"><a href="UserReport.asp?action=report&ToUserNumber=<%=RsUserListObj("UserNumber")%>">举报</a>&nbsp;|&nbsp;<a href="message_write.asp?ToUserNumber=<%=RsUserListObj("UserNumber")%>">发信</a>&nbsp;|&nbsp;<a href="Book_write.asp?ToUserNumber=<%=RsUserListObj("UserNumber")%>&M_type=0">留言</a>&nbsp;|&nbsp;<a href="Friend_add.asp?type=0&ToUserNumber=<%=RsUserListObj("UserNumber")%>&action=addFriend"   onClick="{if(confirm('确认添加为好友吗?')){this.document.inbox.submit();return true;}return false;}">好友</a>&nbsp;</div></td>
        </tr>
        <%
			  RsUserListObj.MoveNext
		  Next
		  %>
        <tr class="hback"> 
          <td colspan="7" class="xingmu"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="80%"> <span class="top_navi">
				<% 		Response.Write("每页:"& RsUserListObj.pagesize &"个,")
							Response.write"&nbsp;共<b>"& select_pagecount &"</b>页<b>&nbsp;" & select_count &"</b>条记录，本页是第<b>"& strpage &"</b>页。"
							if int(strpage)>1 then
								Response.Write"&nbsp;<a href=?page=1&Keyword="&Keyword&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&"&RegTime="&Request("RegTime")&"&CountPage="&Request("CountPage")&"&Integral="& Request("Integral")&">第一页</a>&nbsp;&nbsp;"
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&Keyword="&Keyword&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&"&RegTime="&Request("RegTime")&"&CountPage="&Request("CountPage")&"&Integral="& Request("Integral")&">上一页</a>&nbsp;&nbsp;"
							End if
							If int(strpage)<select_pagecount then
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&Keyword="&Keyword&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&"&RegTime="&Request("RegTime")&"&CountPage="&Request("CountPage")&"&Integral="& Request("Integral")&">下一页</a>&nbsp;"
								Response.Write"&nbsp;<a href=?page="& select_pagecount &"&Keyword="&Keyword&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&"&RegTime="&Request("RegTime")&"&CountPage="&Request("CountPage")&"&Integral="& Request("Integral")&">最后一页</a>&nbsp;&nbsp;"
							End if
								Response.Write"<br>"
								RsUserListObj.close
								Set RsUserListObj=nothing
							End if
							%> </SPAN></td>
              </tr>
            </table>
            </td>
        </tr>
      </table></td>
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
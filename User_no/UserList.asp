<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-��Ա�б�</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; ��Ա�б�ͳ��</td>
          </tr>
        </table>
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td colspan="7" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="43%"> ��������<strong> 
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
                  </strong> ����Ա</td>
                <form action="UserList.asp"  method="post" name="myform" id="myform">
                  <td width="57%"><div align="left">������ 
                      <select name="Name" id="select">
                        <option value="UserName" <%if Request("Name")="UserName" then response.Write("selected")%>>�û���</option>
                        <option value="UserNumber" <%if Request("Name")="UserNumber" then response.Write("selected")%>>�û����</option>
                        <option value="NickName" <%if Request("Name")="NickName" then response.Write("selected")%>>�ǳ�</option>
                        <option value="RealName" <%if Request("Name")="RealName" then response.Write("selected")%>>����</option>
                        <option value="Email" <%if Request("Name")="Email" then response.Write("selected")%>>�����ʼ�</option>
                        <option value="QQ" <%if Request("Name")="QQ" then response.Write("selected")%>>OICQ</option>
                        <option value="MSN" <%if Request("Name")="MSN" then response.Write("selected")%>>MSN</option>
                        <option value="Integral" <%if Request("Name")="Integral" then response.Write("selected")%>>50-����+50</option>
                        <option value="Province" <%if Request("Name")="Province" then response.Write("selected")%>>ʡ��</option>
                        <option value="city" <%if Request("Name")="city" then response.Write("selected")%>>����</option>
                      </select>
                      <input name="keyword" type="text" id="keyword2" value="<%=Keyword%>" size="10">
                      <input name="searchtype" type="checkbox" id="searchtype" value="1" <%if Request("searchtype")="1" then Response.Write("checked")%> >
                      ģ������ 
                      <input type="submit" name="Submit" value="����">
                    </div></td>
                </form>
              </tr>
            </table></td>
        </tr class="hback">
        <tr class="hback"> 
          <td width="17%" class="xingmu"><div align="left"><strong>
		  <%If Request("RegTime") <> "" then
		  		If cint(Request("RegTime"))=1 then%>
		  <a href="UserList.asp?page=<%=strpage%>&Keyword=<%=Keyword%>&Name=<%=NoSqlHack(Request("Name"))%>&searchtype=<%=NoSqlHack(Request("searchtype"))%>&RegTime=0&CountPage=<%=NoSqlHack(Request("CountPage"))%>&Integral=" class="LinkCss">�û���</a>
		  		<%Elseif  cint(Request("RegTime"))=0 then %>
		  <a href="UserList.asp?page=<%=strpage%>&Keyword=<%=Keyword%>&Name=<%=NoSqlHack(Request("Name"))%>&searchtype=<%=NoSqlHack(Request("searchtype"))%>&RegTime=1&CountPage=<%=NoSqlHack(Request("CountPage"))%>&Integral=" class="LinkCss">�û���</a>
		  		<%End if%>
			<% Else	%>
				<a href="UserList.asp?page=<%=strpage%>&Keyword=<%=Keyword%>&Name=<%=NoSqlHack(Request("Name"))%>&searchtype=<%=NoSqlHack(Request("searchtype"))%>&RegTime=0&CountPage=<%=NoSqlHack(Request("CountPage"))%>&Integral=" class="LinkCss">�û���</a>
			<% End If %>
		  </strong></div></td>
          <td width="18%" class="xingmu"><div align="left"><strong>���</strong></div></td>
          <td width="13%" class="xingmu"><div align="left"><strong>OICQ</strong></div></td>
          <td width="9%" class="xingmu"><div align="center"><strong>Email</strong></div></td>
          <td width="7%" class="xingmu"><div align="center"><strong>��ҳ</strong></div></td>
          <td width="10%" class="xingmu"><div align="center"><strong>����</strong></div></td>
          <td width="26%" class="xingmu"><div align="center"><strong>����</strong></div></td>
        </tr>
        <%
		Dim select_count,select_pagecount,i
		if RsUserListObj.eof then
			   RsUserListObj.close
			   set RsUserListObj=nothing
			   Response.Write"<TR><TD colspan=""7""  class=""hback"">û�м�¼��</TD></TR>"
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
						    sOICQ ="<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin="& RsUserListObj("QQ") &"&Site=FoosunCMS&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:"& RsUserListObj("QQ") &":16 alt=""��������"& RsUserListObj("QQ") &"����Ϣ""></a>"
							Response.Write sOICQ
						Else
							Response.Write("û��")
						End if
						%>
            </div></td>
          <td class="hback"><div align="center"><a href="mailto:<% = RsUserListObj("Email")%>">����</a></div></td>
          <td class="hback"><div align="center"><a href="<% = RsUserListObj("homepage")%>">��ҳ</a></div></td>
          <td class="hback"><div align="center">
              <% = RsUserListObj("Integral")%>
            </div></td>
          <td class="hback"><div align="center"><a href="UserReport.asp?action=report&ToUserNumber=<%=RsUserListObj("UserNumber")%>">�ٱ�</a>&nbsp;|&nbsp;<a href="message_write.asp?ToUserNumber=<%=RsUserListObj("UserNumber")%>">����</a>&nbsp;|&nbsp;<a href="Book_write.asp?ToUserNumber=<%=RsUserListObj("UserNumber")%>&M_type=0">����</a>&nbsp;|&nbsp;<a href="Friend_add.asp?type=0&ToUserNumber=<%=RsUserListObj("UserNumber")%>&action=addFriend"   onClick="{if(confirm('ȷ�����Ϊ������?')){this.document.inbox.submit();return true;}return false;}">����</a>&nbsp;</div></td>
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
				<% 		Response.Write("ÿҳ:"& RsUserListObj.pagesize &"��,")
							Response.write"&nbsp;��<b>"& select_pagecount &"</b>ҳ<b>&nbsp;" & select_count &"</b>����¼����ҳ�ǵ�<b>"& strpage &"</b>ҳ��"
							if int(strpage)>1 then
								Response.Write"&nbsp;<a href=?page=1&Keyword="&Keyword&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&"&RegTime="&Request("RegTime")&"&CountPage="&Request("CountPage")&"&Integral="& Request("Integral")&">��һҳ</a>&nbsp;&nbsp;"
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&Keyword="&Keyword&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&"&RegTime="&Request("RegTime")&"&CountPage="&Request("CountPage")&"&Integral="& Request("Integral")&">��һҳ</a>&nbsp;&nbsp;"
							End if
							If int(strpage)<select_pagecount then
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&Keyword="&Keyword&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&"&RegTime="&Request("RegTime")&"&CountPage="&Request("CountPage")&"&Integral="& Request("Integral")&">��һҳ</a>&nbsp;"
								Response.Write"&nbsp;<a href=?page="& select_pagecount &"&Keyword="&Keyword&"&Name="& Request("Name")&"&searchtype="&Request("searchtype")&"&RegTime="&Request("RegTime")&"&CountPage="&Request("CountPage")&"&Integral="& Request("Integral")&">���һҳ</a>&nbsp;&nbsp;"
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->
<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-��������</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; <a href="Callboard.asp">��Ա����</a>&gt;&gt;&gt;��������</td>
          </tr>
        </table>
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td colspan="4" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="44%"> ��������<strong> 
                  <%
				Dim RsUserNewsObj,RsUserNewsSQL
				Dim strpage,strSQLs
				strpage=CintStr(request("page"))
				if len(strpage)=0 Or strpage<1 or trim(strpage)=""  Then strpage="1"
				Set RsUserNewsObj = Server.CreateObject(G_FS_RS)
				If Request("Keyword")<>"" then
						if Request("searchtype") <>"" then
								if  Request("Name") = "title" then
									  strSQLs = " and Title like '%" & NosqlHack(Request("Keyword"))& "%' " 
								Elseif  Request("Name") = "content" then
									  strSQLs = " and Content  like '%" & NosqlHack(Request("Keyword"))& "%' "
								End if
						Else
								if  Request("Name") = "title" then
									  strSQLs = " and title = '" & NosqlHack(Request("Keyword"))& "'"
								Elseif  Request("Name") = "content" then
									  strSQLs = " and content  = '" & NosqlHack(Request("Keyword"))& "'"
								End if
						End if
				Else
						strSQLs = ""
				End if
				RsUserNewsSQL = "Select Newsid,title,content,AddTime,GroupID,NewsPoint,isLock From Fs_ME_News  where isLock=0 "& strSQLs &" Order by NewsID desc"
				RsUserNewsObj.Open RsUserNewsSQL,User_Conn,1,3
				Response.Write "<Font color=red>" & RsUserNewsObj.RecordCount&"</font>"
				%>
                  </strong> ����Ϣ</td>
                <form action="Callboard.asp"  method="post" name="myform" id="myform">
                  <td width="56%"><div align="left">������ 
                      <select name="Name" id="select">
                        <option value="title" <%if Request("Name") = "title" then response.Write("selected")%>>����</option>
                        <option value="content" <%if Request("Name") = "content" then response.Write("selected")%>>����</option>
                      </select>
                      <input name="keyword" type="text" id="keyword2" value="<%=Request("keyword")%>" size="10">
                      <input name="searchtype" type="checkbox" id="searchtype" value="1" <%if Request("searchtype")="1" then Response.Write("checked")%> >
                      ģ������ 
                      <input type="submit" name="Submit" value="����">
                    </div></td>
                </form>
              </tr>
            </table></td>
        </tr class="hback">
        <tr class="hback"> 
          <td width="34%" class="xingmu"><div align="left"><strong>���� </strong></div></td>
          <td width="20%" class="xingmu"><div align="left"><strong>����</strong></div></td>
          <td width="46%" class="xingmu"><div align="left"><strong>����</strong></div></td>
        </tr>
        <%
		Dim select_count,select_pagecount,i
		if RsUserNewsObj.eof then
			   RsUserNewsObj.close
			   set RsUserNewsObj=nothing
			   Response.Write"<TR><TD colspan=""3""  class=""hback"">û�м�¼��</TD></TR>"
		else
				RsUserNewsObj.pagesize = 20
				RsUserNewsObj.absolutepage=cint(strpage)
				select_count=RsUserNewsObj.recordcount
				select_pagecount=RsUserNewsObj.pagecount
				for i=1 to RsUserNewsObj.pagesize
					if RsUserNewsObj.eof Then exit For 
		%>
        <tr class="hback"> 
          <td class="hback"><div align="left">��<a href="ShowCallboard.asp?NewsID=<% = RsUserNewsObj("NewsID")%>"> 
              <% = RsUserNewsObj("title")%>
              </a></div></td>
          <td class="hback"><div align="left"> <% = RsUserNewsObj("addtime")%></div></td>
          <td class="hback"><div align="left">
              <% = Left(RsUserNewsObj("Content"),30)%>
              ... </div></td>
        </tr>
        <%
			  RsUserNewsObj.MoveNext
		  Next
		  %>
        <tr class="hback"> 
          <td colspan="4" class="xingmu"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="80%"> <span class="top_navi"> 
                  <% 	Response.Write("ÿҳ:"& RsUserNewsObj.pagesize &"��,")
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
								RsUserNewsObj.close
								Set RsUserNewsObj=nothing
							End if
							%>
                  </SPAN></td>
                <form name="form1" method="post" action="UserList.asp">
                </form>
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
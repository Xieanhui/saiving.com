<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-日志</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt;日志</td>
          </tr>
        </table>
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td colspan="6" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="44%"> <strong>
				<%
				  dim strTmp,strLogType,strTmp1
				  strLogType = NoSqlHack(Trim(Request.QueryString("LogTye")))
			     if Request.QueryString("LogTye")<>"" then
			  		strTmp =  " and (LogType Like '%"& strLogType &"%' Or LogType = '" & strLogType & "')"
			     Else
			  		strTmp =  " "
			    End if
				if Request("date1") <>"" and  Request("date2")<>"" then
					if isdate(Request("date1"))=false or isdate(Request("date2"))=false then
						strShowErr = "<li>您输入的日期格式不正确</li>"
						Call ReturnError(strShowErr,"")
					else
						If G_IS_SQL_User_DB =0 Then
							strTmp1 = " and Logtime>=#"&datevalue(NoSqlHack(Request("date1")))&"#  and Logtime<=#"&datevalue(NoSqlHack(Request("date2")))&"#"
						Else
							strTmp1 = " and Logtime>='"&datevalue(NoSqlHack(Request("date1")))&"'  and Logtime<='"&datevalue(NoSqlHack(Request("date2")))&"'"
						End if
					End if
				Else
						strTmp1 = ""
				End if
				Dim RsUserListObj,RsUserSQL
				Dim strpage,strSQLs,StrOrders
				strpage=CintStr(request("page"))
				if len(strpage)=0 Or strpage<1 or trim(strpage)=""  Then strpage="1"
				Set RsUserListObj = Server.CreateObject(G_FS_RS)
				RsUserSQL = "Select LogType,UserNumber,points,moneys,LogTime,LogContent,Logstyle From Fs_ME_Log  where UserNumber='"& Fs_User.UserNumber &"' "& strTmp & strTmp1 &" order by  LogID desc"
				RsUserListObj.Open RsUserSQL,User_Conn,1,3
				Response.Write "<Font color=red>" & RsUserListObj.RecordCount&"</font>"
				%>
                  </strong> 个日志　类型：<a href="history.asp?LogTye=注册">注册</a>｜<a href="history.asp?LogTye=%B5%C7%C2%BD">登陆</a>｜<a href="history.asp?LogTye=购买">购买</a>｜<a href="history.asp?LogTye=冲值">冲值</a>｜<a href="history.asp?LogTye=兑换">兑换</a>｜<a href="history.asp?LogTye=其他">其他</a></td>
                <form action="History.asp"  method="post" name="myform" id="myform">
                  <td width="56%"><div align="left">
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="63%" valign="top">从 <input name="date1" type="text" id="date1" value="<%=datevalue(date())-1%>" size="10">
                            到 <input name="date2" type="text" id="date2" value="<%=datevalue(date())%>" size="10">
                            的记录 
                            <input type="submit" name="Submit" value="搜索">
                            日期格式请用1977-6-7格式</td>
                        </tr>
                      </table>
                    </div></td>
                </form>
              </tr>
            </table></td>
        </tr class="hback">
        <tr class="hback"> 
          <td width="17%" class="xingmu"><div align="left"><strong> 类型</strong></div></td>
          <td width="15%" class="xingmu"><div align="left"><strong><%=top_moneyName%></strong></div></td>
          <td width="11%" class="xingmu"><div align="left"><strong>点数</strong></div></td>
          <td width="20%" class="xingmu"><div align="center"><strong>日期</strong></div></td>
          <td width="25%" class="xingmu"><div align="center"><strong>说明</strong></div></td>
          <td width="12%" class="xingmu"><div align="center"><strong>增加/减少</strong></div></td>
        </tr>
        <%
		Dim select_count,select_pagecount,i
		if RsUserListObj.eof then
			   RsUserListObj.close
			   set RsUserListObj=nothing
			   Response.Write"<TR><TD colspan=""6""  class=""hback"">没有记录。</TD></TR>"
		else
				if Request("CountPage")="" or len(Request("CountPage"))<1 then
					RsUserListObj.pagesize = 20
				Else
					RsUserListObj.pagesize = CintStr(Request("CountPage"))
				End if
				RsUserListObj.absolutepage=cint(strpage)
				select_count=RsUserListObj.recordcount
				select_pagecount=RsUserListObj.pagecount
				for i=1 to RsUserListObj.pagesize
					if RsUserListObj.eof Then exit For 
		%>
        <tr class="hback"> 
          <td class="hback"><div align="left"><a href=history.asp?LogTye=<% = RsUserListObj("LogType")%>><% = RsUserListObj("LogType")%></a></div></td>
          <td class="hback"><div align="left"> 
              <% = FormatNumber(RsUserListObj("points"),2,-1)%>
            </div></td>
          <td class="hback"><div align="left"> 
              <% = RsUserListObj("moneys")%>
            </div></td>
          <td class="hback"><div align="center">
              <% = RsUserListObj("LogTime")%>
            </div></td>
          <td class="hback"><div align="center">
              <% = RsUserListObj("LogContent")%>
            </div></td>
          <td class="hback"><div align="center"> 
              <%
			  if RsUserListObj("Logstyle") = 0 then
				  Response.Write("<font color=red>增加</font>")
			  Else
				  Response.Write("减少")
			  End if
			  %>
            </div></td>
        </tr>
        <%
			  RsUserListObj.MoveNext
		  Next
		  %>
        <tr class="hback"> 
          <td colspan="6" class="xingmu"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="80%"> <span class="top_navi">
                  <% 	Response.Write("每页:"& RsUserListObj.pagesize &"个,")
							Response.write"&nbsp;共<b>"& select_pagecount &"</b>页<b>&nbsp;" & select_count &"</b>条记录，本页是第<b>"& strpage &"</b>页。"
							if int(strpage)>1 then
								Response.Write"&nbsp;<a href=?page=1&LogType="&Request("LogTye")&" class=""top_navi"">第一页</a>&nbsp;&nbsp;"
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&LogType="&Request("LogTye")&" class=""top_navi"">上一页</a>&nbsp;&nbsp;"
							End if
							If int(strpage)<select_pagecount then
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&LogType="&Request("LogTye")&" class=""top_navi"">下一页</a>&nbsp;"
								Response.Write"&nbsp;<a href=?page="& select_pagecount &"&LogType="&Request("LogTye")&" class=""top_navi"">最后一页</a>&nbsp;&nbsp;"
							End if
								Response.Write"<br>"
								RsUserListObj.close
								Set RsUserListObj=nothing
							End if
							%>
                  </SPAN></td>
                <form name="form1" method="post" action="UserList.asp">
                </form>
              </tr>
            </table></td>
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






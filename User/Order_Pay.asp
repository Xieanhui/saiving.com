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
if Request.QueryString("Action") = "lock_order" then
	User_Conn.execute("Delete From FS_ME_Order  where OrderNumber='"& NoSqlHack(Request.QueryString("OrderNumber"))&"' and UserNumber='"& Fs_User.UserNumber &"'")
	User_Conn.execute("Delete From FS_ME_Order_detail  where OrderNumber='"& NoSqlHack(Request.QueryString("OrderNumber"))&"'")
	strShowErr = "<li>操作定单成功!</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-定单</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt;定单</td>
        </tr>
        <tr class="hback">
          <td class="hback"><a href="Order.asp">一般定单</a>┆<a href="Order_Pay.asp">在线支付定单</a></td>
        </tr>
      </table>
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="Order.asp">
          <tr class="hback"> 
            <td colspan="6" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="44%"> <strong> 
                    <%
				  dim strTmp,strLogType,strTmp1
				  strLogType = NoSqlHack(Request.QueryString("LogTye"))
			     if Request.QueryString("LogTye")<>"" then
			  		strTmp =  " and LogType='"& strLogType &"'"
			     Else
			  		strTmp =  " "
			    End if
				Dim RsOrderObj,RsOrderSQL
				Dim strpage,strSQLs,StrOrders
				strpage=request("page")
				if len(strpage)=0 Or strpage<1 or trim(strpage)=""  Then strpage="1"
				Set RsOrderObj = Server.CreateObject(G_FS_RS)
				RsOrderSQL = "Select * From FS_ME_Order  where UserNumber='"& Fs_User.UserNumber &"' and OrderType=3 order by  OrderID desc"
				RsOrderObj.Open RsOrderSQL,User_Conn,1,3
				Response.Write RsOrderObj.recordcount
				%>
                    </strong> 个定单</td>
                  <td width="56%"><div align="left"> </div></td>
                </tr>
              </table></td>
          </tr class="hback">
          <tr class="hback"> 
            <td width="20%" class="xingmu"><div align="left"><strong> 定单号(点定单查看详情)</strong></div></td>
            <td width="11%" class="xingmu"><div align="center">审核状态</div></td>
            <td width="21%" class="xingmu"><div align="center">成功日期</div></td>
            <td width="18%" class="xingmu"><div align="center"><strong>日期</strong></div></td>
            <td width="9%" class="xingmu"><div align="center"><strong>类型</strong></div></td>
            <td width="13%" class="xingmu"><div align="center"><strong>支付</strong></div></td>
          </tr>
          <%
		Dim select_count,select_pagecount,i
		if RsOrderObj.eof then
			   RsOrderObj.close
			   set RsOrderObj=nothing
			   set conn=nothing
			   set fs_user=nothing
			   Response.Write"<TR><TD colspan=""6""  class=""hback"">没有记录。</TD></TR>"
		else
				if Request("CountPage")="" or len(Request("CountPage"))<1 then
					RsOrderObj.pagesize = 20
				Else
					RsOrderObj.pagesize = Request("CountPage")
				End if
				RsOrderObj.absolutepage=cint(strpage)
				select_count=RsOrderObj.recordcount
				select_pagecount=RsOrderObj.pagecount
				for i=1 to RsOrderObj.pagesize
					if RsOrderObj.eof Then exit For 
		 %>
          <tr class="hback"> 
            <td class="hback"><div align="left"> 
                <% = RsOrderObj("OrderNumber")%>
              </div></td>
            <td class="hback"> <div align="center"> 
                <%
					if RsOrderObj("isLock")=1 then
						Response.Write"<span class=tx>审核中...<span>"
					Else
						Response.Write"已审核..."
					End if
					%>
              </div></td>
            <td class="hback"><div align="center"> 
                <% = RsOrderObj("M_PayDate")%>
              </div></td>
            <td class="hback"> 
              <% = RsOrderObj("AddTime")%>
            </td>
            <td class="hback"><div align="center"> 
                <%
			if RsOrderObj("OrderType")=0 then
				Response.Write("会员组")
			Elseif RsOrderObj("OrderType")=1 then
				Response.Write("商品")
			Elseif RsOrderObj("OrderType")=2 then
				Response.Write("点卡")
			Elseif RsOrderObj("OrderType")=3 then
				Response.Write("在线支付")
			Else
				Response.Write("其他")
			End if
			%>
              </div></td>
            <td class="hback"> <div align="center"> 
                <%if RsOrderObj("IsSuccess")=0 then%>失败<%Else%><span class="tx">成功</span><%End if%>
              </div></td>
          </tr>
          <%
			  RsOrderObj.MoveNext
		  Next
		  %>
          <tr class="hback"> 
            <td colspan="6" class="xingmu"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="80%"> <span class="top_navi"> 
                    <% 	Response.Write("每页:"& RsOrderObj.pagesize &"个,")
							Response.write"&nbsp;共<b>"& select_pagecount &"</b>页<b>&nbsp;" & select_count &"</b>条记录，本页是第<b>"& strpage &"</b>页。"
							if int(strpage)>1 then
								Response.Write"&nbsp;<a href=?page=1&LogType="&Request("LogTye")&">第一页</a>&nbsp;&nbsp;"
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&LogType="&Request("LogTye")&">上一页</a>&nbsp;&nbsp;"
							End if
							If int(strpage)<select_pagecount then
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&LogType="&Request("LogTye")&">下一页</a>&nbsp;"
								Response.Write"&nbsp;<a href=?page="& select_pagecount &"&LogType="&Request("LogTye")&">最后一页</a>&nbsp;&nbsp;"
							End if
								Response.Write"<br>"
								RsOrderObj.close
								Set RsOrderObj=nothing
							End if
							%>
                    </SPAN></td>
                </tr>
              </table></td>
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
<%
Set Fs_User = Nothing
set user_conn=nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->
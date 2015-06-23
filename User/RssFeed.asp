<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
dim obj_mf_sys_obj,MF_Domain,MF_Site_Name,tmp_c_path
set obj_mf_sys_obj = Conn.execute("select top 1 MF_Domain,MF_Site_Name from FS_MF_Config")
if obj_mf_sys_obj.eof then
	strShowErr = "<li>找不到主系统配置信息！</li>"
	Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
else
	MF_Domain = obj_mf_sys_obj("MF_Domain")
	MF_Site_Name = obj_mf_sys_obj("MF_Site_Name")
end if
obj_mf_sys_obj.close:set obj_mf_sys_obj = nothing
tmp_c_path =MF_Domain&"/"
%>
<script language="javascript" type="text/javascript">
function copyToClipBoardK(f_obj)
{
   var clipBoardContent=''; 
   clipBoardContent = 'http://<%=tmp_c_path%>xml/ns/'+f_obj+'.xml';
   window.clipboardData.setData("Text",clipBoardContent);
   alert("你已复制RSS订阅地址!");
}
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-RSS聚合</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt; <a href="RssFeed.asp">RSS聚合</a> &gt;&gt; <%=NoSqlHack(Request.QueryString("Sub_sys"))%></td>
        </tr>
        <tr class="hback"> 
          <td class="hback"><a href="RssFeed.asp">首页</a>┆<a href="RssFeed.asp?Sub_sys=NS">新闻</a></td>
        </tr>
      </table>
      <%
		  Sub mainHelp()
		  %>
      <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
        <tr class="hback">
          <td> <strong>·什么是RSS<font color="#3366CC"> </font></strong><br> <br>
            　RSS是站点用来和其他站点之间共享内容的一种简易方式，通常被用于新闻和其他按顺序排列的网站，例如Blog。一段项目的介绍可能包含新闻的全部介绍等。或者仅仅是额外的内容或者简短的介绍。这些项目的链接通常都能链接到全部的内容。网络用户可以在客户端借助于支持RSS的新闻聚合工具软件，在不打开网站内容页面的情况下阅读支持RSS输出的网站内容。<br> 
            <br>
            <strong>·RSS如何工作</strong><br> <br>
            　您一般需要下载和安装一个RSS新闻阅读器，然后从网站提供的聚合新闻目录列表中订阅您感兴趣的新闻栏目的内容。订阅后，您将会及时获得所订阅新闻频道的最新内容。<br>
            <br> <strong>·RSS新闻阅读器的特点</strong><br> <br>
            　a. 没有广告或者图片来影响标题或者文章概要的阅读。
<p>　b. RSS阅读器自动更新你定制的网站内容，保持新闻的及时性。</p>
            <p>　c. 用户可以加入多个定制的RSS提要，从多个来源搜集新闻整合到单个数据流中。<br>
              <br>
              <strong>·RSS阅读器下载</strong><br>
            </p>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="24%"><div align="center"> <a href="http://fox.foxmail.com.cn/" target="_blank"><img src="Images/Rss/foxmail.gif" width="134" height="51" border="0"></a> 
                    <table width="98%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><div align="center"><a href="http://fox.foxmail.com.cn/">腾讯Foxmail 
                            6</a></div></td>
                      </tr>
                    </table>
                    
                  </div></td>
                <td width="19%"><div align="center"><a href="http://www.potu.com" target="_blank"><img src="Images/Rss/zbt.gif" border="0"></a><br>
                    <table width="98%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td><div align="center"><a href="http://www.potu.com" target="_blank">周博通RSS阅读器</a></div></td>
                      </tr>
                    </table>
                    
                  </div></td>
                <td width="57%" valign="bottom"><div align="center"><a href="http://www.sharpreader.net/SharpReader0960_Setup.exe" target="_blank"><img src="images/Rss/sharp%20Reader.gif" width="91" height="18" border="0"></a>&nbsp;<a href="http://www.rssreader.com/download/rssreader.zip" target="_blank"><img src="images/Rss/RssReader.gif" width="91" height="19" border="0"></a>&nbsp;<a href="http://feeddemon.com/download/dloadhandler.asp?file=feeddemon-trial.exe" target="_blank"><img src="images/Rss/FeedDemon.gif" width="91" height="20" border="0"></a>&nbsp;<a href="http://www.newzcrawler.com/downloads.shtml" target="_blank"><img src="images/Rss/Nc.gif" width="91" height="19" border="0"></a> 
                  </div>
                  <table width="98%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td><div align="center">国外RSS聚合阅读器</div></td>
                    </tr>
                  </table>
                  
                </td>
              </tr>
            </table>
            
          </td>
        </tr>
      </table>
      <%end sub%>
      <%
	dim sub_sys
	sub_sys = NoSqlHack(Request.QueryString("sub_sys"))
	select case sub_sys
			case "NS"
				call NS()
			case else
				Call MainHelp()
	end select
	sub NS()
		Dim getParam
		set getParam = Conn.execute("select top 1 RSSTF From FS_NS_SysParam")
		if getParam(0)<>1 then
			Response.Write "<p><div align=center>新闻频道没开通RSS聚合订阅</div></p>"
			getParam.close:set getParam=nothing
		Else
			Response.Write("<table width=""98%"" border=""0"" align=""center"" cellpadding=""5"" cellspacing=""1"" class=""table"">")
			Response.Write("<tr class=""hback""><td><img src=""images/-.gif"" border=""0""></img><b>最新更新</b></td><td><img src=""../sys_images/rss.gif"" border=""0"" alt=""点击复制此RSS聚合地址"" style=""cursor:hand;"" onclick=""Javascript:copyToClipBoardK('now');""></td><td>http://"& tmp_c_path &"xml/ns/now.xml</td></tr>")
			dim str_ClassID,news_SQL,obj_news_rs,str_action,obj_news_rs_1
			str_ClassID = NoSqlHack(Request.QueryString("ClassID"))
			if str_ClassID<>"" then
				news_SQL = "Select Orderid,id,ClassID,ClassName,ClassEName,IsUrl,AddNewsType from FS_NS_NewsClass where Parentid  = '"& str_ClassID &"' and ReycleTF=0 Order by Orderid desc,ID desc"
			Else
				news_SQL = "Select Orderid,id,ClassID,ClassName,ClassEName,IsUrl,AddNewsType from FS_NS_NewsClass where Parentid  = '0'  and ReycleTF=0  Order by Orderid desc,ID desc"
			End if
			Set obj_news_rs = server.CreateObject(G_FS_RS)
			obj_news_rs.Open news_SQL,Conn,1,3
			  if Not obj_news_rs.eof then
				Do while Not obj_news_rs.eof 
					Set obj_news_rs_1 = server.CreateObject(G_FS_RS)
					obj_news_rs_1.Open "Select Count(ID) from FS_NS_NewsClass where ParentID='"& obj_news_rs("ClassID") &"'",Conn,1,1
					if obj_news_rs_1(0)>0 then
						str_action=  "<img src=""images/+.gif"" border=""0""></img><a href=""RssFeed.asp?ClassID="& obj_news_rs("ClassID") &"&Sub_sys="&NoSqlHack(Request.QueryString("Sub_sys"))&""" title=""点击查看此类下的其他子类RSS聚合"">"& obj_news_rs("ClassName") &"</a>"
					Else
						str_action=  "<img src=""images/-.gif"" border=""0""></img><a href=""#"" title=""没子类了"">"& obj_news_rs("ClassName") &"</a>"
					End if
					obj_news_rs_1.close:set obj_news_rs_1 =nothing
					Response.Write"<tr class=""hback""><td width=""220"">"
					Response.Write str_action &"</td><td width=""20""><img src=""../sys_images/rss.gif"" border=""0"" alt=""点击复制此RSS聚合地址"" style=""cursor:hand;"" onclick=""Javascript:copyToClipBoardK('"& obj_news_rs("ClassEName") &"');""></td><td>http://"& tmp_c_path &"xml/ns/"& obj_news_rs("ClassEName") &".xml</td>"
					Response.Write "</td></tr>"
					obj_news_rs.MoveNext
				loop
				response.Write("<tr class=""hback""><td colspan=""3""><span class=""tx"">提醒：点击图片复制RSS聚合地址</span></td></tr>")
				Response.Write("</table>")
			End if
			obj_news_rs.close:set obj_news_rs = nothing
		end if
%>
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr>
          <td class="hback"> <table width="76%" border="0" align="center" cellpadding="5" cellspacing="1">
              <tr>
                <td>
         <%
		 if getParam(0)=1 then
			  dim rs
			  set rs = Server.CreateObject(G_FS_RS)
			  rs.open "select top 10 id,NewsTitle,Content,Source,Author,addtime From FS_NS_News where isLock=0 and IsURL=0 and isRecyle=0 and isdraft=0 order by PopId desc,id desc",Conn,1,3
			  if rs.eof then
				response.Write("无内容")
				  rs.close:set rs=nothing
			  else
				do while not rs.eof 
					response.Write "·<a href=Public_info.asp?type=NS&Url="&rs("id")&"><b><font style=""font-size:14px"">"&rs("NewsTitle")&"</font></b></a><br>"&chr(13)
					response.Write "<br>日期："&rs("addtime")&",来源："&rs("Source")&",作者："&rs("Author")&"<br><br>"
					response.Write "<font style="" line-height:22px"">&nbsp;&nbsp;&nbsp;&nbsp;"&left(rs("Content"),200)&"...</font><p></p>"
				rs.movenext
				loop	
			  end if
			  If Not IsNull(rs) Then
				rs.close:set rs=Nothing
			  End if
		  end if
		getParam.close:set getParam=nothing
		  %>
                </td>
              </tr>
            </table></td>
        </tr>
      </table> 
      <%
	end sub
%>
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
%>

<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






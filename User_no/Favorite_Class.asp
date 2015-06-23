<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp" -->
<%
Dim rs
if request("Action")="del" then
	if Request("id")="" then
		strShowErr = "<li>错误的参数！</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		User_Conn.execute("Delete from FS_ME_FavoriteClass where UserNumber='"&Fs_User.UserNumber&"' and ClassID ="&CintStr(Request("id")))
		User_Conn.execute("update FS_ME_Favorite set FavoClassID=0 where UserNumber='"&Fs_User.UserNumber&"' and  FavoClassID="&CintStr(Request("id")))
		set User_Conn=nothing:set Fs_User=nothing
		strShowErr = "<li>删除成功！</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Favorite_Class.asp")
		Response.end
	end if
end if
if Request.Form("Action")="add_save" then
   set rs= Server.CreateObject(G_FS_RS)
   rs.open "select ClassID,ClassCName,AddTime,UserNumber From FS_ME_FavoriteClass where ClassCName='"&NoSqlHack(trim(Request.Form("ClassCName")))&"' and UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
   if not rs.eof then
   		rs.close:Set rs=nothing
		set User_Conn=nothing:set Fs_User=nothing
		strShowErr = "<li>此收藏夹已经存在！</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
   else
   		rs.addnew
   		rs("ClassCName")=CintStr(request.Form("ClassCName"))
		rs("UserNumber")=Fs_User.UserNumber
		rs("addtime")=now
		rs.update
   		rs.close:Set rs=nothing
		set User_Conn=nothing:set Fs_User=nothing
		strShowErr = "<li>添加成功！</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Favorite_Class.asp")
		Response.end
   end if
end if
if Request.Form("Action")="edit_save" then
	set rs= Server.CreateObject(G_FS_RS)
	rs.open "select ClassID,ClassCName,UserNumber From FS_ME_FavoriteClass where Classid="&CintStr(Request.Form("id"))&" and UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
	rs("ClassCName")=NoSqlHack(request.Form("ClassCName"))
	rs.update
	rs.close:Set rs=nothing
	set User_Conn=nothing:set Fs_User=nothing
	strShowErr = "<li>修改成功！</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Favorite_Class.asp")
	Response.end
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-相册管理</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt; <a href="Favorite.asp">收藏夹管理</a> 
            &gt;&gt;分类管理</td>
        </tr>
      </table> 
		  
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr>
          <td class="hback"><a href="Favorite.asp">全部</a>┆<a href="Favorite.asp?Type=0">新闻</a>┆<a href="Favorite.asp?Type=1">下载</a>┆<a href="Favorite.asp?Type=2">企业会员</a>┆<a href="Favorite.asp?Type=3">供求信息</a>┆<a href="Favorite.asp?Type=4">商品</a>┆<a href="Favorite.asp?Type=5">房产信息</a>┆<a href="Favorite.asp?Type=6">招聘</a>┆<a href="Favorite.asp?Type=7">日志</a>┆<a href="Favorite_Class.asp">收藏夹(分类)管理</a>┆<a href="Favorite_Class.asp?Action=add">增加分类</a></td>
        </tr>
        <tr> 
          <td class="hback"> 
            <%
		  response.Write("	<table width=""98%"" align=center cellpadding=""2"" cellspacing=""1""><tr>")
		  dim t_k
		  t_k=0
		  set rs = Server.CreateObject(G_FS_RS)
		  rs.open "select ClassID,ClassCName,UserNumber From FS_ME_FavoriteClass where UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
		  do while not rs.eof 
		  	Response.Write("	<td width=""24%"" valign=bottom><img src=""images/folderopened.gif""></img><a href=Favorite.asp?classid="&rs("ClassID")&">"&rs("ClassCName")&"</a><a href=Favorite_class.asp?id="&rs("Classid")&"&Action=edit><修改></a><a href=""Favorite_class.asp?Action=del&id="&rs("Classid")&""" onClick=""{if(confirm('确定删除吗？')){return true;}return false;}""><删除></a></td>")
		  rs.movenext
		  t_k = t_k+1
		  if t_k mod 4 =0 then
		  	Response.Write("	</tr>")
		  end if
		  loop
		  response.Write("	</table>")
		  rs.close:set rs=nothing
		  %>
          </td>
        </tr>
      </table> 
      <%if NoSqlHack(Request.QueryString("Action"))="add" then%>
      
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="">
          <tr> 
            <td colspan="2" class="xingmu">增加收藏夹</td>
          </tr>
          <tr> 
            <td width="24%" class="hback"><div align="right">分类名称</div></td>
            <td width="76%" class="hback"><input name="ClassCName" type="text" id="ClassCName" value="" size="40"></td>
          </tr>
          <tr> 
            <td class="hback"><div align="right"></div></td>
            <td class="hback"><input type="submit" name="Submit" value="增加收藏夹"> 
              <input name="Action" type="hidden" id="Action" value="add_save"></td>
          </tr>
        </form>
      </table>
      
      <%end if%>
       <%if NoSqlHack(Request.QueryString("Action"))="edit" then
	   if NoSqlHack(request.QueryString("id"))="" then
	   		rs.close:set rs=nothing	
			set User_Conn=nothing:set Fs_User=nothing
			response.Write("错误的参数")
			response.end
	   end if
	   set rs= Server.CreateObject(G_FS_RS)
	   rs.open "select Classid,ClassCName,UserNumber From FS_ME_FavoriteClass where Classid="&CintStr(request.QueryString("id")),User_Conn,1,3
	   %>
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="">
          <tr> 
            <td colspan="2" class="xingmu">修改收藏夹</td>
          </tr>
          <tr> 
            <td width="24%" class="hback"><div align="right">相册分类名称</div></td>
            <td width="76%" class="hback"><input name="ClassCName" type="text" id="ClassCName" value="<%=rs("ClassCName")%>" size="40"><input name="id" type="hidden" id="id" value="<%=rs("Classid")%>" size="40"></td>
          </tr>
          <tr> 
            <td class="hback"><div align="right"></div></td>
            <td class="hback"><input type="submit" name="Submit2" value="修改收藏夹"> 
              <input name="Action" type="hidden" id="Action" value="edit_save"></td>
          </tr>
        </form>
      </table>
      <%end if%>
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
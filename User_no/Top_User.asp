<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<%
dim obj_user_rs,tmp_order,tsql,order_tip,Fs_User
set Fs_User=new cls_user
tmp_order = NoSqlHack(Request.QueryString("type"))
select case tmp_order     
	case "int"
		 tsql = "select top 8 UserNumber,NickName,UserName,Integral,FS_Money,LoginNum,hits From FS_ME_Users where isLock=0 Order by Integral desc,UserID desc"
	 case "money"
		 tsql = "select  top 8 UserNumber,NickName,UserName,Integral,FS_Money,LoginNum,hits From FS_ME_Users where isLock=0 Order by FS_Money desc,UserID desc"
	 case "active"
		 tsql = "select  top 8 UserNumber,NickName,UserName,Integral,FS_Money,LoginNum,hits From FS_ME_Users where isLock=0 Order by LoginNum desc,UserID desc"
	 case "hits"
		 tsql = "select  top 8 UserNumber,NickName,UserName,Integral,FS_Money,LoginNum,hits From FS_ME_Users where isLock=0 Order by hits desc,UserID desc"
	case else
		 tsql = "select top 8  UserNumber,NickName,UserName,Integral,FS_Money,LoginNum,hits From FS_ME_Users where isLock=0 Order by Integral desc,UserID desc"
end select 
set obj_user_rs= Server.CreateObject(G_FS_RS)
obj_user_rs.open tsql,user_conn,1,3
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head></head>
<body class="hback">
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="hback">
  <tr  class="hback">
    <td>
      <%
 if Not obj_user_rs.eof then
	  Response.Write("<table width=""100%"" height=""100%""  border=""0"" cellspacing=""0"" cellpadding=""0"">")
	 Do while Not obj_user_rs.eof
		select case tmp_order
			case "int"
				 order_tip = obj_user_rs("Integral")
			 case "money"
				 order_tip = obj_user_rs("FS_Money")
			 case "active"
				 order_tip = obj_user_rs("LoginNum")
			 case "hits"
				 order_tip = obj_user_rs("hits")
			case else
				 order_tip = obj_user_rs("Integral")
		end select 
		Response.Write"<tr><td><img src=""images/dot.gif"" border=""0""><a href=ShowUser.asp?UserNumber="&obj_user_rs("UserNumber")&" title="""" target=""_blank"">"& obj_user_rs("NickName") &"("& obj_user_rs("UserName") &")</a></td><td width=""50"" align=""right"">"& order_tip &"</td></tr>"
		  obj_user_rs.movenext
	Loop
	Response.Write("</table>")	
Else
	   Response.Write("没有记录")
End if
obj_user_rs.close
set obj_user_rs = nothing
%>
    </td>
  </tr>
</table>
</body>
</html>
<%
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






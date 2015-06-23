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
	dim rs_sysobj
	set rs_sysobj = User_Conn.execute("select top 1 PointChange From FS_ME_SysPara")
	if rs_sysobj.eof then
		response.Write "配置信息错误，请与系统管理员联系"
		response.end
		rs_sysobj.close:set rs_sysobj=nothing
	else
		PointChange = rs_sysobj(0)
		rs_sysobj.close:set rs_sysobj=nothing
	end if
	Dim PointChangestr,PointChange,PointChangestr1,PointChangestr2,PointChangestr3,frmMoney
	frmMoney = Request.Form("money")
	if isnull(frmMoney) then frmMoney=0 
	frmMoney = left(frmMoney,11)
	if not isnumeric(frmMoney) then frmMoney = 0 
	if clng(frmMoney)<0 then frmMoney = 0
	
	PointChangestr = split(PointChange,",")
	if not isarray(PointChangestr) then
		response.Write"错误的参数"
		response.end
	else
		PointChangestr1=PointChangestr(0)
		PointChangestr2=PointChangestr(1)
		PointChangestr3=PointChangestr(2)
	end if
	if NoSqlHack(request.Form("Action"))="changepoint_save" then
	   if PointChangestr1<>"3" and PointChangestr1<>"2" then
			strShowErr = "<li>管理员设定不能兑换积分</li>"
			set Conn = nothing
			set User_Conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	   end if
	   '开始兑换
	   if Fs_User.NumFS_Money<clng(frmMoney) then
			strShowErr = "<li>您的金币少于您输入的金币数量</li>"
			set Conn = nothing
			set User_Conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	   end if
	   User_Conn.execute("Update FS_ME_Users Set Integral=Integral+"&clng(frmMoney)*PointChangestr2&",FS_Money=FS_Money-"&clng(frmMoney)&" where UserNumber='"&Fs_User.UserNumber&"'")
	   Call Fs_User.AddLog("兑换",Fs_User.UserNumber,0,clng(frmMoney),"减少金币",1) 
	   Call Fs_User.AddLog("兑换",Fs_User.UserNumber,0,clng(frmMoney)*PointChangestr2,"增加积分",0) 
		strShowErr = "<li>兑换金币为积分成功</li>"
		set Conn = nothing
		set User_Conn=nothing
		Response.Redirect("lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	if NoSqlHack(request.Form("Action"))="changemoney_save" then
	   if PointChangestr1<>"3" and PointChangestr1<>"1" then
			strShowErr = "<li>管理员设定不能兑换金币</li>"
			set Conn = nothing
			set User_Conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	   end if
	   '开始兑换
	   if Fs_User.NumIntegral<clng(frmMoney) then
			strShowErr = "<li>您的积分少于您输入的积分数量</li>"
			set Conn = nothing
			set User_Conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	   end if
	   dim tmp_change
	   tmp_change = 1/PointChangestr3
	   User_Conn.execute("Update FS_ME_Users Set Integral=Integral-"&clng(frmMoney)&",FS_Money=FS_Money+"&replace(formatnumber(clng(frmMoney)/tmp_change,2,-1),",","")&" where UserNumber='"&Fs_User.UserNumber&"'")
	   Call Fs_User.AddLog("兑换",Fs_User.UserNumber,0,clng(frmMoney)*PointChangestr3,"增加金币",0) 
	   Call Fs_User.AddLog("兑换",Fs_User.UserNumber,clng(frmMoney),0,"减少积分",1) 
		strShowErr = "<li>兑换金币为积分成功</li>"
		set Conn = nothing
		set User_Conn=nothing
		Response.Redirect("lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-我的帐户</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt; <a href="MyAccount.asp">我的帐户</a> &gt;&gt; 兑换</td>
          </tr>
        </table>
        <%if noSqlHack(Request.QueryString("action"))="changepoint" then%>
        <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
           <form name="form1" method="post" action="">
             <tr>
               <td height="27" colspan="2" class="hback">
			     <div align="center">
			       <%
			   dim not_tf,tmp_str
			   if PointChangestr1="3" or PointChangestr1="2" then
					if Fs_User.NumFS_Money>1 then
						tmp_str = "您可以兑换积分，您准备兑换<input name=""money"" type=""text"" size=""8"" value="""& split(FormatNumber(Fs_User.NumFS_Money,2,-1),".")(0) &""" />个金币"
						not_tf=true
					elseif Fs_User.NumFS_Money<1 then
						not_tf=false
						tmp_str = "<span class=""tx"">您的金币不够，不能兑换！</span>"
					end if
					Response.Write "系统允许您1个金币兑换"& PointChangestr2 &"个积分，您目前金币："& FormatNumber(Fs_User.NumFS_Money,2,-1) &"，"& tmp_str &""
			   else
			   		response.Write "系统不允许金币兑换积分！！"
					not_tf=false
			   end if
			   %>
	            </div></td>
             </tr>
             <tr>
           <td height="22" colspan="2" class="hback">
              <label></label>              <div align="center">
                <input type="submit" name="Submit" value="开始兑换积分"<%if not_tf=false then response.Write "disabled"%>>
                <input name="Action" type="hidden" id="Action" value="changepoint_save">
              </div></td>
          </tr></form>
      </table>
	  <%
	  elseif noSqlHack(Request.QueryString("action"))="changemoney" then 
	  %>
        <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
           <form name="form1" method="post" action="">
             <tr>
               <td height="27" colspan="2" class="hback">
			     <div align="center">
			   <%
			   if PointChangestr1="3" or PointChangestr1="1" then
					if Fs_User.NumIntegral>=1/PointChangestr3 then
						tmp_str = "您可以兑换金币，您准备兑换<input name=""money"" type=""text"" size=""8"" value="""& Fs_User.NumIntegral &""" />个积分"
						not_tf=true
					elseif Fs_User.NumIntegral<1/PointChangestr3 then
						not_tf=false
						tmp_str = "<span class=""tx"">您的积分不够，不能兑换！</span>"
					end if
					Response.Write "系统允许您"& 1/PointChangestr3 &"个积分兑换1个金币，您目前积分："& FormatNumber(Fs_User.NumIntegral,2,-1) &"，"& tmp_str &""
			   else
			   		response.Write "系统不允许积分兑换金币！！"
					not_tf=false
			   end if
			   %>
	            </div></td>
             </tr>
             <tr>
           <td height="22" colspan="2" class="hback">
              <label></label>              <div align="center">
                <input type="submit" name="Submit" value="开始兑换金币"<%if not_tf=false then response.Write "disabled"%>>
                <input name="Action" type="hidden" id="Action" value="changemoney_save">
              </div></td>
          </tr></form>
      </table>
	  <%end if%></td>
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






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
		response.Write "������Ϣ��������ϵͳ����Ա��ϵ"
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
		response.Write"����Ĳ���"
		response.end
	else
		PointChangestr1=PointChangestr(0)
		PointChangestr2=PointChangestr(1)
		PointChangestr3=PointChangestr(2)
	end if
	if NoSqlHack(request.Form("Action"))="changepoint_save" then
	   if PointChangestr1<>"3" and PointChangestr1<>"2" then
			strShowErr = "<li>����Ա�趨���ܶһ�����</li>"
			set Conn = nothing
			set User_Conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	   end if
	   '��ʼ�һ�
	   if Fs_User.NumFS_Money<clng(frmMoney) then
			strShowErr = "<li>���Ľ������������Ľ������</li>"
			set Conn = nothing
			set User_Conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	   end if
	   User_Conn.execute("Update FS_ME_Users Set Integral=Integral+"&clng(frmMoney)*PointChangestr2&",FS_Money=FS_Money-"&clng(frmMoney)&" where UserNumber='"&Fs_User.UserNumber&"'")
	   Call Fs_User.AddLog("�һ�",Fs_User.UserNumber,0,clng(frmMoney),"���ٽ��",1) 
	   Call Fs_User.AddLog("�һ�",Fs_User.UserNumber,0,clng(frmMoney)*PointChangestr2,"���ӻ���",0) 
		strShowErr = "<li>�һ����Ϊ���ֳɹ�</li>"
		set Conn = nothing
		set User_Conn=nothing
		Response.Redirect("lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	if NoSqlHack(request.Form("Action"))="changemoney_save" then
	   if PointChangestr1<>"3" and PointChangestr1<>"1" then
			strShowErr = "<li>����Ա�趨���ܶһ����</li>"
			set Conn = nothing
			set User_Conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	   end if
	   '��ʼ�һ�
	   if Fs_User.NumIntegral<clng(frmMoney) then
			strShowErr = "<li>���Ļ�������������Ļ�������</li>"
			set Conn = nothing
			set User_Conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	   end if
	   dim tmp_change
	   tmp_change = 1/PointChangestr3
	   User_Conn.execute("Update FS_ME_Users Set Integral=Integral-"&clng(frmMoney)&",FS_Money=FS_Money+"&replace(formatnumber(clng(frmMoney)/tmp_change,2,-1),",","")&" where UserNumber='"&Fs_User.UserNumber&"'")
	   Call Fs_User.AddLog("�һ�",Fs_User.UserNumber,0,clng(frmMoney)*PointChangestr3,"���ӽ��",0) 
	   Call Fs_User.AddLog("�һ�",Fs_User.UserNumber,clng(frmMoney),0,"���ٻ���",1) 
		strShowErr = "<li>�һ����Ϊ���ֳɹ�</li>"
		set Conn = nothing
		set User_Conn=nothing
		Response.Redirect("lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-�ҵ��ʻ�</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; <a href="MyAccount.asp">�ҵ��ʻ�</a> &gt;&gt; �һ�</td>
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
						tmp_str = "�����Զһ����֣���׼���һ�<input name=""money"" type=""text"" size=""8"" value="""& split(FormatNumber(Fs_User.NumFS_Money,2,-1),".")(0) &""" />�����"
						not_tf=true
					elseif Fs_User.NumFS_Money<1 then
						not_tf=false
						tmp_str = "<span class=""tx"">���Ľ�Ҳ��������ܶһ���</span>"
					end if
					Response.Write "ϵͳ������1����Ҷһ�"& PointChangestr2 &"�����֣���Ŀǰ��ң�"& FormatNumber(Fs_User.NumFS_Money,2,-1) &"��"& tmp_str &""
			   else
			   		response.Write "ϵͳ�������Ҷһ����֣���"
					not_tf=false
			   end if
			   %>
	            </div></td>
             </tr>
             <tr>
           <td height="22" colspan="2" class="hback">
              <label></label>              <div align="center">
                <input type="submit" name="Submit" value="��ʼ�һ�����"<%if not_tf=false then response.Write "disabled"%>>
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
						tmp_str = "�����Զһ���ң���׼���һ�<input name=""money"" type=""text"" size=""8"" value="""& Fs_User.NumIntegral &""" />������"
						not_tf=true
					elseif Fs_User.NumIntegral<1/PointChangestr3 then
						not_tf=false
						tmp_str = "<span class=""tx"">���Ļ��ֲ��������ܶһ���</span>"
					end if
					Response.Write "ϵͳ������"& 1/PointChangestr3 &"�����ֶһ�1����ң���Ŀǰ���֣�"& FormatNumber(Fs_User.NumIntegral,2,-1) &"��"& tmp_str &""
			   else
			   		response.Write "ϵͳ��������ֶһ���ң���"
					not_tf=false
			   end if
			   %>
	            </div></td>
             </tr>
             <tr>
           <td height="22" colspan="2" class="hback">
              <label></label>              <div align="center">
                <input type="submit" name="Submit" value="��ʼ�һ����"<%if not_tf=false then response.Write "disabled"%>>
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->






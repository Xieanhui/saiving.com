<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
dim rsOrder
on error resume next
	set rsOrder = User_Conn.execute("select OrderID From FS_ME_Order where OrderNumber='"&NoSqlHack(Request("OrderNumber"))&"' and MoneyAmount="& CintStr(Request.QueryString("Moneys"))&" and UserNumber='"&Fs_User.UserNumber&"' and OrderID="&CintStr(Request("OrderID"))&"")
	if rsOrder.eof then
		response.Write "�Ҳ�����¼��"
		Response.end
	end if
'��ʼ�ʻ�֧��
if Request.Form("Action")="save" then
	'�õ�������Ϣ
	Dim rsPay
	set rsPay = User_Conn.execute("select UserNumber,MoneyAmount,isPay From FS_ME_Order where OrderID="&CintStr(Request.Form("OrderID")))
	if not rsPay.eof then
		if RsPay("isPay")=1 then
			strShowErr = "<li>�����Ѿ�֧����������֧��!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		if rsPay("MoneyAmount")<0 then
			strShowErr = "<li>���Ķ�����Ϣ�д���,�벻Ҫ�÷Ƿ�;����ȡ��Ʒ!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		if Fs_User.NumFS_Money < rsPay("MoneyAmount") then
			strShowErr = "<li>���Ľ�Ҳ���!</li><li><a href=""pay.asp"">����˴�Ϊ�ʻ���ֵ��</a></li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		else
			'���»�Ա���
			User_Conn.execute("Update FS_ME_Users set FS_Money=FS_Money-"&rsPay("MoneyAmount")&" where UserNumber='"&Fs_User.UserNumber&"'")
			'���¶���״̬
			User_Conn.execute("Update FS_ME_Order set IsSuccess=1,IsPay=1 where UserNumber='"&Fs_User.UserNumber&"' and OrderId="&CintStr(Request.Form("OrderID")))
		end if
		strShowErr = "<li>֧���ɹ�!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Order.asp")
		Response.end
	else
		strShowErr = "<li>�������!</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Order.asp")
		Response.end
	end if
	set rsPay = nothing
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-֧������</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; ֧������</td>
          </tr>
        </table>
		<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr>
            <td colspan="2" class="xingmu">���������Ķ�����Ϣ</td>
          </tr>
          <tr class="hback">
            <td width="20%"><div align="right">�������</div></td>
            <td width="80%"><%=NoSqlHack(Request("OrderNumber"))%></td>
          </tr>
          <tr class="hback">
            <td><div align="right">�������</div></td>
            <td><%=NoSqlHack(Request("Moneys"))%>RMB</td>
          </tr>
          <tr class="hback">
            <td colspan="2"><div align="center" class="tx">�����ڰ���������֧��������ƾ�����������ǿͷ���Ա��ϵ</div></td>
          </tr>
        </table>
		<%
		select Case NoSqlHack(Request("PayStyle"))
			Case "PostOrBank"
				Call PostOrBank()
			Case "MySelfAcc"
				Call MySelfAcc()
			Case "Card"
				Call Card()
			Case Else
				Call MySelfAcc()
		end Select
		Sub PostOrBank()
		%>
        <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr>
            <td colspan="2" class="xingmu">��ѡ���ǵĸ����ǵ������ʾֻ��밴�����·�ʽ�����ǻ��</td>
          </tr>
          <tr class="hback">
            <td width="20%" height="27" class="hback_1"><div align="right">�ʾֻ������</div></td>
            <td width="80%">
			<%
			dim rsMall
			if IsExist_SubSys("MS") Then
			 set rsMall = Conn.execute("select top 1 Address,Content,PostCode,UserName From FS_MS_PayMethod")
			 if not rsMall.eof then
				 Response.Write "��ַ��"& rsMall("Address") &"<br />"
				 Response.Write "�ռ��ˣ�"& rsMall("UserName") &"<br />"
				 Response.Write "�������룺"& rsMall("PostCode") &"<br />"
			 else
				 Response.Write "�����û��ͨ�̳���ϵͳ�����ֹ��ٴ�����"
			 end if
			 rsMall.close:set rsMall = nothing
			%>
			<%else%>
			�����û��ͨ�̳���ϵͳ�����ֹ��ٴ�����
			<%end if%>			</td>
          </tr>
          <tr class="hback">
            <td height="28" class="hback_1"><div align="right">���е������</div></td>
            <td>
			<%
			if IsExist_SubSys("MS") Then
			 set rsMall = Conn.execute("select top 1 Other From FS_MS_PayMethod")
			 if not rsMall.eof then
				 Response.Write ""& rsMall("Other") &"<br />"
			 else
				 Response.Write "�����û��ͨ�̳���ϵͳ�����ֹ��ٴ�����"
			 end if
			 rsMall.close:set rsMall = nothing
			%>
			<%else%>
			�����û��ͨ�̳���ϵͳ�����ֹ��ٴ�����
			<%end if%>
			</td>
          </tr>
          <tr class="hback">
            <td height="28" class="hback_1"><div align="right">������֪</div></td>
            <td>
			<%
			if IsExist_SubSys("MS") Then
			 set rsMall = Conn.execute("select top 1 Content From FS_MS_PayMethod")
			 if not rsMall.eof then
				 Response.Write ""& rsMall("Content") &"<br />"
			 else
				 Response.Write "�����û��ͨ�̳���ϵͳ�����ֹ��ٴ�����"
			 end if
			 rsMall.close:set rsMall = nothing
			%>
			<%else%>
			�����û��ͨ�̳���ϵͳ�����ֹ��ٴ�����
			<%end if%>
			</td>
          </tr>
        </table>
		<%end sub%>
		<%sub MySelfAcc()%>
		<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr>
            <td width="100%" class="xingmu">��ѡ������ʻ�֧��</td>
          </tr>
          <tr>
            <form name="form1" method="post" action="">
              <td class="hback"><input type="button" name="Submit" value="����Ҫ֧���Ľ��Ϊ��<%=NoSqlHack(Request("Moneys"))%>����ȷ��֧����֧���ɹ��󽫿۳�������Ӧ���" onClick="{if(confirm('��ȷ��֧����?')){this.document.form1.submit();return true;}return false;}">
              <input name="Action" type="hidden" id="Action" value="save">
              <input name="OrderNumber" type="hidden" id="OrderNumber" value="<%=NoSqlHack(Request("OrderNumber"))%>">
              <input name="Moneys" type="hidden" id="Moneys" value="<%=NoSqlHack(Request("Moneys"))%>">
              <input name="OrderID" type="hidden" id="OrderID" value="<%=NoSqlHack(Request("OrderID"))%>"></td>
            </form>
          </tr>
        </table>
		<%end sub%>
		<%sub Card()%>
		<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr>
            <td width="100%" class="xingmu">�㿨֧��</td>
          </tr>
          <tr>
            <td class="hback">�㿨����֧��������ʱû��ͨ���������Ҫʹ�õ㿨֧�������ȹ���㿨��ֵ�������ʻ������ʻ�����<a href="card.asp">[�㿨��ֵ]</a> </td>
          </tr>
        </table>
		<%end sub%>
	    <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr>
            <td width="100%" class="xingmu">ѡ������֧����ʽ</td>
          </tr>
          <tr>
            <td class="hback"><a href="onlinepay.asp?OrderNumber=<%=NoSqlHack(Request("OrderNumber"))%>&Moneys=<%=NoSqlHack(Request("Moneys"))%>&OrderID=<%=NoSqlHack(Request("OrderID"))%>">����֧��</a> ��<a href="PayCenter.asp?PayStyle=MySelfAcc&OrderNumber=<%=NoSqlHack(Request("OrderNumber"))%>&Moneys=<%=NoSqlHack(Request("Moneys"))%>&OrderID=<%=NoSqlHack(Request("OrderID"))%>">�ʻ�֧��</a> </td>
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






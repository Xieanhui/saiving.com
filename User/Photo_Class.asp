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
		strShowErr = "<li>����Ĳ�����</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		User_Conn.execute("Delete from FS_ME_PhotoClass where UserNumber='"&Fs_User.UserNumber&"' and ID ="&CintStr(Request("id")))
		User_Conn.execute("update FS_ME_Photo set Classid=0 where UserNumber='"&Fs_User.UserNumber&"' and  ClassID="&CintStr(Request("id")))
		set User_Conn=nothing:set Fs_User=nothing
		strShowErr = "<li>ɾ���ɹ���</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Photo_Class.asp")
		Response.end
	end if
end if
if Request.Form("Action")="add_save" then
   set rs= Server.CreateObject(G_FS_RS)
   rs.open "select title,id,UserNumber From FS_ME_PhotoClass where title='"&NoSqlHack(Request.Form("title"))&"' and UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
   if not rs.eof then
   		rs.close:Set rs=nothing
		set User_Conn=nothing:set Fs_User=nothing
		strShowErr = "<li>������Ѿ����ڣ�</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
   else
   		rs.addnew
   		rs("title")=NoSqlHack(request.Form("title"))
		rs("UserNumber")=Fs_User.UserNumber
		rs.update
   		rs.close:Set rs=nothing
		set User_Conn=nothing:set Fs_User=nothing
		strShowErr = "<li>��ӳɹ���</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Photo_Class.asp")
		Response.end
   end if
end if
if Request.Form("Action")="edit_save" then
	set rs= Server.CreateObject(G_FS_RS)
	rs.open "select title,id,UserNumber From FS_ME_PhotoClass where id="&CintStr(NoSqlHack(Request.Form("id")))&" and UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
	rs("title")=NoSqlHack(request.Form("title"))
	rs.update
	rs.close:Set rs=nothing
	set User_Conn=nothing:set Fs_User=nothing
	strShowErr = "<li>�޸ĳɹ���</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Photo_Class.asp")
	Response.end
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-������</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; <a href="PhotoManage.asp">������</a> 
            &gt;&gt;�������</td>
        </tr>
      </table> 
		  
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr>
          <td class="hback"><a href="PhotoManage.asp">�����ҳ</a>��<a href="Photo_add.asp">����ͼƬ</a>��<a href="PhotoManage.asp?isRec=1">���Ƽ���ͼƬ</a>��<a href="Photo_filt.asp">�õ�Ƭ����</a>��<a href="Photo_Class.asp">������</a>��<a href="Photo_Class.asp?Action=add">���ӷ���</a></td>
        </tr>
        <tr> 
          <td class="hback"> 
            <%
		  response.Write("	<table width=""98%"" align=center cellpadding=""2"" cellspacing=""1""><tr>")
		  dim t_k
		  t_k=0
		  set rs = Server.CreateObject(G_FS_RS)
		  rs.open "select id,title,UserNumber From FS_ME_PhotoClass where UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
		  do while not rs.eof 
		  	Response.Write("	<td width=""24%"" valign=bottom><img src=""images/folderopened.gif""></img><a href=PhotoManage.asp?classid="&rs("id")&" title=""����鿴�˷����µ�ͼƬ"">"&rs("title")&"</a><a href=Photo_Class.asp?id="&rs("id")&"&Action=edit><�޸�></a><a href=""Photo_Class.asp?Action=del&id="&rs("id")&""" onClick=""{if(confirm('ȷ��ɾ����')){return true;}return false;}""><ɾ��></a></td>")
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
            <td colspan="2" class="xingmu">����������</td>
          </tr>
          <tr> 
            <td width="24%" class="hback"><div align="right">����������</div></td>
            <td width="76%" class="hback"><input name="title" type="text" id="title" value="" size="40"></td>
          </tr>
          <tr> 
            <td class="hback"><div align="right"></div></td>
            <td class="hback"><input type="submit" name="Submit" value="����������">
              <input name="Action" type="hidden" id="Action" value="add_save"></td>
          </tr>
        </form>
      </table>
      
      <%end if%>
       <%if NoSqlHack(Request.QueryString("Action"))="edit" then
	   if NoSqlHack(request.QueryString("id"))="" then
	   		rs.close:set rs=nothing	
			set User_Conn=nothing:set Fs_User=nothing
			response.Write("����Ĳ���")
			response.end
	   end if
	   set rs= Server.CreateObject(G_FS_RS)
	   rs.open "select id,title,UserNumber From FS_ME_PhotoClass where id="&CintStr(request.QueryString("id")),User_Conn,1,3
	   %>
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="">
          <tr> 
            <td colspan="2" class="xingmu">�޸�������</td>
          </tr>
          <tr> 
            <td width="24%" class="hback"><div align="right">����������</div></td>
            <td width="76%" class="hback"><input name="title" type="text" id="title" value="<%=rs("title")%>" size="40"><input name="id" type="hidden" id="id" value="<%=rs("id")%>" size="40"></td>
          </tr>
          <tr> 
            <td class="hback"><div align="right"></div></td>
            <td class="hback"><input type="submit" name="Submit2" value="�޸�������"> 
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->
<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<%
User_GetParm
if Request.QueryString("skin")= "1" then
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = 1
elseif Request.QueryString("skin")= "2" then
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = 2
elseif Request.QueryString("skin")= "3" then
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = 3
elseif Request.QueryString("skin")= "4" then
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = 4
elseif Request.QueryString("skin")= "5" then
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = 5
else
	if Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = "" then
		response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = 3
	else
		response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")
	end if
End if
Dim p_UserNumber,RsObj,str_lock,p_str
if Request.QueryString("UserNumber")<>"" then
	p_str = "UserNumber='"& NoSqlHack(Request.QueryString("UserNumber")) &"'"
elseif Request.QueryString("UserName")<>"" then
	p_str = "UserName='"& NoSqlHack(Request.QueryString("UserName")) &"'"
else
	response.Write("����Ĳ���")
	response.end
end if
User_Conn.execute("update FS_ME_Users set hits=hits+1 where "& p_str &"")
set RsObj = Server.CreateObject (G_FS_RS)
RsObj.Open "select isLock,UserName,RealName,GroupID,Integral,LoginNum,RegTime, LastLoginTime,LastLoginIP,UserNumber,FS_Money,ConNumber,UserID,HomePage,BothYear,Tel,MSN,QQ,Corner,Province,City,Address,PostCode,PassQuestion,SelfIntro,isOpen,Certificate,CertificateCode,Vocation,HeadPic,NickName,Mobile,CloseTime,IsCorporation,isMessage,Email,sex,safeCode,UserLoginCode,HeadPicsize,OnlyLogin,UserFavor,IsMarray,PassAnswer,hits from FS_ME_Users where "& p_str &"",User_Conn,1,1
If RsObj.eof Then 
	strShowErr = "<li>���û��ѱ�ɾ���򲻴��ڸ��û���Ϣ</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
dim Temp_Admin_Name,Temp_Admin_Pass_Word,Temp_Admin_Parent_Admin,Temp_Admin_Is_Super
'if G_ADMIN_Login_Type = 0 then
	Temp_Admin_Name = Session("Admin_Name")
	Temp_Admin_Pass_Word = Session("Admin_Pass_Word")
	Temp_Admin_Parent_Admin = Session("Admin_Parent_Admin")
	Temp_Admin_Is_Super = Session("Admin_Is_Super")
'else
'	Temp_Admin_Name = request.Cookies("FoosunAdminCookie")("Admin_Name")
'	Temp_Admin_Pass_Word = request.Cookies("FoosunAdminCookie")("Admin_Pass_Word")
'	Temp_Admin_Parent_Admin = request.Cookies("FoosunAdminCookie")("Temp_Admin_Parent_Admin")
'	Temp_Admin_Is_Super = request.Cookies("FoosunAdminCookie")("Admin_Is_Super")
'end if
if Temp_Admin_Name<>"" and Temp_Admin_Pass_Word<>"" and Temp_Admin_Parent_Admin<>"" and Temp_Admin_Is_Super<>"" then
	if RsObj("isLock")=0 then
		str_lock="��<a href=../"&G_ADMIN_DIR&"/User/LockUser.asp?UserNumber="&RsObj("UserNumber")&"&action=Lock>�������û�</a>"
	else
		str_lock="��<a href=../"&G_ADMIN_DIR&"/User/LockUser.asp?UserNumber="&RsObj("UserNumber")&"&action=UnLock>�������û�</a>"
	end if
end if
if RsObj.eof then
		strShowErr = "<li>����Ĳ���</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
Else
	if session("Admin_Name")="" and session("Admin_Pass_Word")="" and session("Admin_Is_Super")="" then
		if RsObj("isLock") = 1 then
			strShowErr = "<li>���û��ѱ�����</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-<% = RsObj("UserName")%></title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head></head>
<body>
<table width="98%" height="135" border="0" align="center" cellpadding="0" cellspacing="1" class="table">
  <tr class="back"> 
    <td width="82%" valign="top" class="hback"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="76%"  valign="top"> <table width="98%" height="256" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
              <tr> 
                <td height="29" colspan="5"  class="xingmu"> <span class="bigtitle"><strong>��<strong>
                  <%  = RsObj("UserName")%>
                  </strong></strong>�Ļ�����Ϣ������: 
                  <%  = RsObj("hits")%>
                  ��</span><strong>��</strong></td>
              </tr>
              <tr> 
                <td height="24" class="hback_1">&nbsp;</td>
                <td colspan="4" class="hback"> <div align="right"><a href="Message_write.asp?ToUserNumber=<%=RsObj("UserNumber")%>" target="_blank">����</a>��<a href="book_write.asp?ToUserNumber=<%=RsObj("UserNumber")%>&M_Type=0" target="_blank">����</a>��<a href="Friend_add.asp?type=0&UserName=<%=RsObj("UserName")%>" target="_blank">��Ϊ����</a>��<a href="UserReport.asp?action=report&ToUserNumber=<%=RsObj("UserNumber")%>">�ٱ�</a><%=str_lock%>
                    ��������������������ƫ�ã�<img src="images/changeskin.gif" width="86" height="12" border="0" usemap="#Map"></div></td>
              </tr>
              <tr> 
                <td width="18%" height="24" class="hback_1"> <div align="center">�û�����</div></td>
                <td width="16%" class="hback"><%=RsObj("RealName")%></td>
                <td width="20%" class="hback"><div align="center" style="display:none"><a href="../<%=G_USER_DIR%>?User=<%=RsObj("UserNumber")%>" ><strong>�鿴<%  = RsObj("UserName")%>�Ŀռ�</strong></a></div></td>
                <td width="46%" colspan="2" rowspan="8" class="hback"><div align="center"></div>
                  <div align="center"> </div>
                  <div align="center"> 
                    <%
				  Dim strHeadPicsizearr,strHeadPicsizearrw,strHeadPicsizearrh
				  If Not IsNull(RsObj("HeadPicsize")) then
					  strHeadPicsizearr = split(RsObj("HeadPicsize"),",")
					  strHeadPicsizearrw = strHeadPicsizearr(0)
					  strHeadPicsizearrh = strHeadPicsizearr(1)
				 End if
				  if Trim(RsObj("HeadPic")) <>"" or len(Trim(RsObj("HeadPic"))) >0 then
				  %>
                    <table width="42" border="0" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC" class="table">
                      <tr> 
                        <td height="40" bgcolor="#FFFFFF"><img src="<%=RsObj("HeadPic")%>" width="<%If Not IsNull(RsObj("HeadPicsize")) Then response.write(strHeadPicsizearr(0))%>" height="<%If Not IsNull(RsObj("HeadPicsize")) Then response.write(strHeadPicsizearr(1))%>"></td>
                      </tr>
                    </table>
                  </div>
                  <div align="center"></div></td>
                <%Else%>
                <img src="images/noHeadpic.gif" border="0"> 
                <%End if%>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">�û��ǳ�</div></td>
                <td colspan="2" class="hback"><%=RsObj("NickName")%></td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">�� �� ��</div></td>
                <td colspan="2" class="hback"><%=RsObj("UserName")%></td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">�û����</div></td>
                <td colspan="2" class="hback"><%=RsObj("UserNumber")%></td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">��½ʱ��</div></td>
                <td colspan="2" class="hback"><%=RsObj("LastLoginTime")%> </td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">ע������</div></td>
                <td height="24" colspan="2" class="hback"><%=RsObj("RegTime")%></td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">ʡ������</div></td>
                <td colspan="2" class="hback"> <%=RsObj("Province")%> </td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"><div align="center">�ǡ�����</div></td>
                <td colspan="2" class="hback"><%=RsObj("City")%></td>
              </tr>
            </table>
            <%
			if RsObj("isOpen")=1 or(session("Admin_Name")<>"" and session("Admin_Pass_Word")<>"" and session("Admin_Is_Super")<>"") then
		    %>
            <table width="98%" height="231" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
              <tr> 
                <td height="29" colspan="4"  class="xingmu"> <span class="bigtitle"><strong>����ϵ��ʽ</strong></span></td>
              </tr>
              <tr> 
                <td width="18%" height="24" class="hback_1"> <div align="center">�����ʼ�</div></td>
                <td width="36%" class="hback"><a href="mailto:<%=RsObj("Email")%>"><%=RsObj("Email")%></a></td>
                <td width="18%" class="hback_1"><div align="center">��վ��ҳ</div></td>
                <td width="28%" class="hback"><a href="<%=RsObj("Homepage")%>" target="_blank"><%=RsObj("Homepage")%></a></td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">��ϵ�绰</div></td>
                <td class="hback"><%=RsObj("Tel")%></td>
                <td class="hback_1"><div align="center">�ƶ��绰</div></td>
                <td class="hback"><%=RsObj("mobile")%></td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">��Ѷ�ѣ�</div></td>
                <td class="hback"><%=RsObj("qq")%>
                  <%
						if  Len(Trim(RsObj("QQ")))>4 then
							Dim sOICQ
						    sOICQ ="<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin="& RsObj("QQ") &"&Site=FoosunCMS&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:"& RsObj("QQ") &":16 alt=""��������"& RsObj("QQ") &"����Ϣ""></a>"
							Response.Write sOICQ
						Else
							Response.Write("û��")
						End if
				%>
                </td>
                <td class="hback_1"><div align="center">�û�MSN</div></td>
                <td class="hback"><%=RsObj("MSN")%></td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">�ء���ַ</div></td>
                <td class="hback"><%=RsObj("Province")%><%=RsObj("City")%><%=RsObj("address")%></td>
                <td class="hback_1"><div align="center">�Ƿ���</div></td>
                <td class="hback">
				<%
				if RsObj("IsMarray")=0 then 
					response.Write("����")
				Elseif RsObj("IsMarray")=1 then 
					response.Write("�ѻ�")
				Else
					response.Write("δ��")
				End if
				%></td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">��������</div></td>
                <td class="hback"><%=RsObj("bothyear")%> </td>
                <td class="hback_1"><div align="center">��½�ɣ�</div></td>
                <td class="hback"><%=RsObj("LastLoginIP")%></td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"><div align="center">��½����</div></td>
                <td class="hback"><%=RsObj("LoginNum")%></td>
                <td class="hback_1"><div align="center">�û����</div></td>
                <td class="hback"><%=RsObj("FS_Money")%>&nbsp;<%=p_MoneyName%></td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"><div align="center">�û�����</div></td>
                <td class="hback"><%=RsObj("Integral")%></td>
                <td class="hback_1"><div align="center"></div></td>
                <td class="hback">&nbsp;</td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"><div align="center">���˰���</div></td>
                <td colspan="3" class="hback"><%=RsObj("UserFavor")%></td>
              </tr>
            </table>
            <map name="MapMap">
              <area shape="rect" coords="2,0,14,18" href="ShowUser.asp?UserNumber=<% = RsObj("UserNumber")%>&skin=1">
              <area shape="rect" coords="19,1,32,17" href="ShowUser.asp?UserNumber=<% = RsObj("UserNumber")%>&skin=2">
              <area shape="rect" coords="37,0,50,13" href="ShowUser.asp?UserNumber=<% = RsObj("UserNumber")%>&skin=3">
              <area shape="rect" coords="52,-1,67,13" href="ShowUser.asp?UserNumber=<% = RsObj("UserNumber")%>&skin=4">
              <area shape="rect" coords="72,0,87,14" href="ShowUser.asp?UserNumber=<% = RsObj("UserNumber")%>&skin=5">
            </map> 
            <%Else%>
            <table width="98%" height="59" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
              <tr> 
                <td height="22"  class="xingmu"> <span class="bigtitle"><strong>����ϵ����</strong></span></td>
              </tr>
              <tr> 
                <td height="34" class="hback"> 
                  <div align="left">��ϵ��������Ϊ����</div></td>
              </tr>
            </table>
            <map name="MapMap2">
              <area shape="rect" coords="2,0,14,18" href="ShowUser.asp?UserNumber=<% = RsObj("UserNumber")%>&skin=1">
              <area shape="rect" coords="19,1,32,17" href="ShowUser.asp?UserNumber=<% = RsObj("UserNumber")%>&skin=2">
              <area shape="rect" coords="37,0,50,13" href="ShowUser.asp?UserNumber=<% = RsObj("UserNumber")%>&skin=3">
              <area shape="rect" coords="52,-1,67,13" href="ShowUser.asp?UserNumber=<% = RsObj("UserNumber")%>&skin=4">
              <area shape="rect" coords="72,0,87,14" href="ShowUser.asp?UserNumber=<% = RsObj("UserNumber")%>&skin=5">
            </map>
            <%End if%>
            <table width="98%" height="97" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
              <tr> 
                <td height="22" colspan="4" class="xingmu"><span class="bigtitle"><strong>������ר��</strong></span></td>
              </tr>
              <tr> 
                <td width="18%" height="24" class="hback_1"><div align="center"><a href="infomanage.asp">�ҵ�ר��</a></div></td>
                <td width="82%" colspan="3" class="hback"> <%
					Dim RsInfoObj
					Set RsInfoObj = server.CreateObject(G_FS_RS)
					RsInfoObj.open "select  Top 5 ClassID,ClassCName,ClassEName,ParentID,UserNumber,ClassTypes,AddTime,ClassContent From FS_ME_InfoClass where ParentID=0 and UserNumber='"& RsObj("UserNumber")&"' order by ClassID desc",User_Conn,1,1
					 if Not RsInfoObj.eof then
						 Do while Not RsInfoObj.eof
							Response.Write "<a href=""ShowInfoClass.asp?ClassID="& RsInfoObj("ClassID") & "&UserNumber=" & RsInfoObj("UserNumber") & """ Title=""" & RsInfoObj("ClassContent") & """>" & RsInfoObj("ClassCName") & "</a>&nbsp;&nbsp;&nbsp;"
							  RsInfoObj.movenext
						  Loop	
					Else
						   Response.Write("û��ר��")
					End if
					RsInfoObj.close:set RsInfoObj = nothing
				  %> </td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"><div align="center"><a href="GroupClass.asp">�ҵ���Ⱥ</a></div></td>
                <td colspan="3" class="hback"> <%
					Dim RsGroupObj
					Set RsGroupObj = server.CreateObject(G_FS_RS)
					RsGroupObj.open "select  Top 5 ClassID,Title,Content,UserNumber From FS_ME_GroupDebateManage where UserNumber='"& RsObj("UserNumber")&"' order by ClassID desc",User_Conn,1,1
					 if Not RsGroupObj.eof then
						 Do while Not RsGroupObj.eof
							Response.Write "<a href=""myGroup.asp?ClassID="& RsGroupObj("ClassID") & "&UserNumber=" & RsGroupObj("UserNumber") & """ Title=""" & RsGroupObj("Content") & """>" & RsGroupObj("Title") & "</a>&nbsp;&nbsp;&nbsp;"
							  RsGroupObj.movenext
						  Loop	
					Else
						   Response.Write("û����Ⱥ")
					End if
					RsGroupObj.close:set RsGroupObj = nothing
				  %> </td>
              </tr>
            </table>
            <%If  RsObj("IsCorporation")= 0  then%> <table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1" class="table">
              <tr> 
                <td height="22" colspan="4" class="xingmu"><span class="bigtitle"><strong>����ҵ����</strong></span></td>
              </tr>
              <tr> 
                <td height="37" colspan="4" class="hback"> <div align="left">û��ͨ��ҵ���ϣ�</div></td>
              </tr>
            </table>
            <%Else%> 
            <table width="98%"  border="0" align="center" cellpadding="5" cellspacing="1" class="table">
              <tr> 
                <td height="22" colspan="4" class="xingmu"><span class="bigtitle"><strong>����ҵ���ϡ������������� 
                  ������������������<a href="Corp_Card_add.asp?UserNumber=<%=RsObj("UserNumber")%>" class="top_navi"><strong>�ղ� 
                  <%  = RsObj("UserName")%>
                  ����Ƭ</strong></a> </strong></span></td>
              </tr>
              <%
				Dim RsCorpObj
				Set RsCorpObj = server.CreateObject(G_FS_RS)
				RsCorpObj.open "select  C_Name,C_ShortName,C_Province,C_City,C_Address,C_ConactName,C_Vocation,C_PostCode,C_Tel,C_Fax,C_BankName,C_BankUserName,isYellowPage,isYellowPageCheck From FS_ME_CorpUser where UserNumber='"& RsObj("UserNumber")&"'",User_Conn,1,1
				 if Not RsCorpObj.eof then
				%>
              <tr> 
                <td width="18%" height="24" class="hback_1"> <div align="center">��˾����</div></td>
                <td width="82%" colspan="3" class="hback"> <% = RsCorpObj("C_Name") %> </td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">��˾���</div></td>
                <td colspan="3" class="hback"> <% = RsCorpObj("C_ShortName") %> </td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"><div align="center">��˾��ϵ��</div></td>
                <td colspan="3" class="hback"> <% = RsCorpObj("C_ConactName") %> </td>
              </tr>
              <tr>
                <td height="24" class="hback_1"><div align="center">��ϵ��ְλ</div></td>
                <td colspan="3" class="hback">
                  <% = RsCorpObj("C_Vocation") %>
                </td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">��˾��ַ</div></td>
                <td colspan="3" class="hback"> <% = RsCorpObj("C_Province") &RsCorpObj("C_City") &RsCorpObj("C_Address")%>
                  ,&nbsp;&nbsp;�ʱࣺ 
                  <% = RsCorpObj("C_PostCode") %> </td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">��˾�绰</div></td>
                <td colspan="3" class="hback"> <% = RsCorpObj("C_Tel") %> </td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">��˾����</div></td>
                <td colspan="3" class="hback"> <% = RsCorpObj("C_Fax") %> </td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">�����ʻ�</div></td>
                <td colspan="3" class="hback">����: 
                  <% = RsCorpObj("C_BankName") %>
                  ���ʺ�: 
                  <% = RsCorpObj("C_BankUserName") %> </td>
              </tr>
              <tr> 
                <td height="24" class="hback_1"> <div align="center">��ҳ��ͨ</div></td>
                <td colspan="3" class="hback"> <%
				  	  if  RsCorpObj("isYellowPage") = 0 then
						  Response.Write("��û��ͨ��ҳ") 
					 Else
						   if RsCorpObj("isYellowPageCheck")=0 then
								Response.Write("�Ѿ���ͨ��ҳ��&nbsp;&nbsp;��ûͨ��")
							Else
								Response.Write("�Ѿ���ͨ��ҳ")
							End if
					 End if
				 %> </td>
              </tr>
              <%
			Else
			     Response.Write("<tr><td height=""40"" colspan=""4"" class=""hback""><b>����ҵ��Ա����û�ҵ���ҵ����</b></a></td></tr>")
			End if
			RsCorpObj.close
			Set RsCorpObj = nothing
			%>
            </table>
            <%End if%> </td>
          <td width="24%" valign="top"><table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1" class="table" style="display:none">
              <tr> 
                <td height="22" colspan="4" class="xingmu"><span class="bigtitle"><strong>�������Ƽ�</strong></span></td>
              </tr>
              <tr> 
                <td height="76" colspan="4" class="hback"> 
                  <div align="left">{$RecGQ}</div></td>
              </tr>
            </table> 
            <table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1" class="table"  style="display:none">
              <tr> 
                <td height="22" colspan="4" class="xingmu"><span class="bigtitle"><strong>�����»���</strong></span></td>
              </tr>
              <tr> 
                <td height="76" colspan="4" class="hback"> <div align="left">{$newGroup}</div></td>
              </tr>
            </table>
            <table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1" class="table">
              <tr> 
                <td height="22" colspan="4" class="xingmu"><span class="bigtitle"><strong>���û�����</strong></span></td>
              </tr>
              <tr> 
                <td height="76" colspan="4" class="hback"> <div align="left">
                    <%
					Dim RsUserTopObj,Ti
					Set RsUserTopObj = server.CreateObject(G_FS_RS)
					RsUserTopObj.open "select  UserName,UserNumber,NickName,hits,islock From FS_ME_Users where Islock=0 order by hits desc,UserID DESC",User_Conn,1,1
					 if Not RsUserTopObj.eof then
						 for Ti = 1 to 10
							If RsUserTopObj.eof Then Exit for
							Response.Write "<img src=""images/dot.gif"" border=""0""></img><a href=""ShowUser.asp?UserNumber="& RsUserTopObj("UserNumber") & """ Title=""" & RsUserTopObj("NickName") & """>" & RsUserTopObj("NickName") & "[" & RsUserTopObj("UserName") & "]</a>&nbsp;<font style=""font-size:9px"" color=""#999999"">("& RsUserTopObj("hits")&")</font><br>"
							  RsUserTopObj.movenext
						  next	
					Else
						   Response.Write("û���û�")
					End if
					RsUserTopObj.close:set RsUserTopObj = nothing
				  %>
                  </div></td>
              </tr>
            </table>
            <table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1" class="table">
              <tr> 
                <td height="22" colspan="4" class="xingmu"><span class="bigtitle"><strong>����������</strong></span></td>
              </tr>
              <tr> 
                <td height="76" colspan="4" class="hback"> <div align="left">
                    <%
					Dim RsUserPointObj,pi
					Set RsUserPointObj = server.CreateObject(G_FS_RS)
					RsUserPointObj.open "select  UserName,UserNumber,NickName,hits,Integral,islock From FS_ME_Users where Islock=0 order by Integral desc,Fs_Money desc",User_Conn,1,1
					 if Not RsUserPointObj.eof then
						 for pi = 1 to 10
							If RsUserPointObj.eof Then Exit for
							Response.Write "<img src=""images/dot.gif"" border=""0""></img><a href=""ShowUser.asp?UserNumber="& RsUserPointObj("UserNumber") & """ Title=""" & RsUserPointObj("NickName") & """>" & RsUserPointObj("NickName") & "[" & RsUserPointObj("UserName") & "]</a>&nbsp;<font style=""font-size:9px"" color=""#999999"">("& RsUserPointObj("Integral")&")</font><br>"
							  RsUserPointObj.movenext
						  Next	
					Else
						   Response.Write("û���û�")
					End if
					RsUserPointObj.close:set RsUserPointObj = nothing
				  %>
                  </div></td>
              </tr>
            </table>
            <table width="98%" height="124" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
              <tr> 
                <td width="50%" height="27"  class="xingmu"><span class="bigtitle"><strong>������</strong></span>(<a href="Top_User.asp?type=int" target="sys_Frame" class="top_navi">����</a><span class="top_navi">��</span><a href="Top_User.asp?type=money" target="sys_Frame" class="top_navi">���</a><span class="top_navi">��</span><a href="Top_User.asp?type=active" target="sys_Frame" class="top_navi">���Ծ</a><span class="top_navi">��<a href="Top_User.asp?type=hits" target="sys_Frame" class="top_navi">����</a></span>)</td>
              </tr>
              <tr> 
                <td height="94" valign="top" class="hback"><iframe src="top_user.asp?type=int" name="sys_Frame" scrolling="no" frameborder="0" align="middle" width="100%" height="175"></iframe></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr class="back"> 
    <td height="20" class="xingmu"> <div align="left"> 
        <!--#include file="Copyright.asp" -->
      </div></td>
  </tr>
</table>
<map name="Map">
  <area shape="rect" coords="2,0,14,18" href="ShowUser.asp?UserNumber=<% = RsObj("UserNumber")%>&skin=1">
  <area shape="rect" coords="19,1,32,17" href="ShowUser.asp?UserNumber=<% = RsObj("UserNumber")%>&skin=2">
  <area shape="rect" coords="37,0,50,13" href="ShowUser.asp?UserNumber=<% = RsObj("UserNumber")%>&skin=3">
  <area shape="rect" coords="52,-1,67,13" href="ShowUser.asp?UserNumber=<% = RsObj("UserNumber")%>&skin=4">
  <area shape="rect" coords="72,0,87,14" href="ShowUser.asp?UserNumber=<% = RsObj("UserNumber")%>&skin=5">
</map>
</body>
</html>
<%
End If
Set RsObj = Nothing
%>
<!--Powsered by Foosun Inc.,Products:FoosunCMS V5.0ϵ��,site:foosun.cn-->






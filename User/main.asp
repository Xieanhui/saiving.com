<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim straction
straction = Request("action")
if straction="Unmessage" then
	User_Conn.execute("update FS_ME_Users set ismessage= 0 where UserNumber='"& Fs_User.UserNumber &"'")
	strShowErr = "<li>���ı�վ����ȡ��</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../main.asp")
	Response.end
Elseif straction = "ismessage" then
	User_Conn.execute("update FS_ME_Users set ismessage= 1 where UserNumber='"& Fs_User.UserNumber &"'")
	strShowErr = "<li>���ı�վ���ϳɹ�</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../main.asp")
	Response.end
Elseif straction = "Open" then
	User_Conn.execute("update FS_ME_Users set isOpen= 1 where UserNumber='"& Fs_User.UserNumber &"'")
	strShowErr = "<li>���⿪�����Ͽ���</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../main.asp")
	Response.end
Elseif straction = "Close" then
	User_Conn.execute("update FS_ME_Users set isOpen= 0 where UserNumber='"& Fs_User.UserNumber &"'")
	strShowErr = "<li>���⿪������ȡ��</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../main.asp")
	Response.end
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>��ӭ�û�<%=Fs_User.UserName%>����<%=GetUserSystemTitle%>��Ա����</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE1 {font-size: 14px}
-->
</style>
<head></head>
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
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="Top_navi.asp" -->
    </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60%"  valign="top"> 
              <table width="98%" height="273" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
              <tr> 
                <td height="21" colspan="4"  class="hback_1"> <span class="bigtitle"><strong>��������Ϣ</strong></span><strong><a href="ShowUser.asp?UserNumber=<%=Fs_User.UserNumber%>" target="_blank">&lt;��ϸ��Ϣ&gt;</a>�����������Ƽ���</strong><a href="award/award.asp?action=myaward">���ֳ齱</a></td>
              </tr>
              <tr> 
                <td width="18%" height="24" class="hback"> <div align="center" class="hback">�û�����</div></td>
                <td width="36%" class="hback"><%=Fs_User.RealName%></td>
                <td colspan="2" rowspan="7" class="hback"> <div align="center"></div>
                  <div align="center"> 
                    <%
				  Dim strHeadPicsizearr,strHeadPicsizearrw,strHeadPicsizearrh
				  if not isNull(Fs_User.HeadPicsize) then 
					  strHeadPicsizearr = split(Fs_User.HeadPicsize,",")
					  strHeadPicsizearrw = strHeadPicsizearr(0)
					  strHeadPicsizearrh = strHeadPicsizearr(1)
				  End if
				  if Trim(Fs_User.HeadPic) <>"" or len(Trim(Fs_User.HeadPic)) >0 then
				  %>
                    <table width="42" border="0" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC" class="table">
                      <tr> 
                        <td height="40" bgcolor="#FFFFFF"><img src="<%=Fs_User.HeadPic%>" width="<% = strHeadPicsizearr(0)%>" height="<% = strHeadPicsizearr(1)%>"></td>
                      </tr>
                    </table>
                  </div>
                  <div align="center"></div></td>
                <%Else%>
                <img src="images/noHeadpic.gif" border="0"> 
                <%End if%>
              </tr>
              <tr> 
                <td height="24" class="hback"> <div align="center" class="hback">�û��ǳ�</div></td>
                <td class="hback"><%=Fs_User.NickName%></td>
              </tr>
              <tr> 
                <td height="24" class="hback"> <div align="center" class="hback">�� �� ��</div></td>
                <td class="hback"><%=Fs_User.UserName%></td>
              </tr>
              <tr> 
                <td height="24" class="hback"> <div align="center" class="hback">�û����</div></td>
                <td class="hback"><%=Fs_User.UserNumber%></td>
              </tr>
              <tr> 
                <td height="24" class="hback"> <div align="center" class="hback">�����ʼ�</div></td>
                <td class="hback"><%=Fs_User.Email%> </td>
              </tr>
              <tr> 
                <td height="24" class="hback"> <div align="center" class="hback">�� ½ IP</div></td>
                <td height="24" class="hback"><%=Fs_User.LastLoginIP%></td>
              </tr>
              <tr> 
                <td height="24" class="hback"> <div align="center" class="hback">��������</div></td>
                <td class="hback"> <%
				  		if Fs_User.CloseTime ="3000-1-1" then
				     		Response.Write"û����"
						Else
							Response.Write Fs_User.CloseTime
						End if
				  %> </td>
              </tr>
              <tr> 
                <td height="24" class="hback"> <div align="center" class="hback">��½ʱ��</div></td>
                <td class="hback"><%=Fs_User.LastLoginTime%></td>
                <td width="18%" class="hback"> <div align="center" class="hback">���Ͽ���</div></td>
                <td width="28%" class="hback"> <%
				    if Fs_User.isOpen=0 then
						Response.Write("<a href=""main.asp?action=Open"">��˽</a>")
					Else
						Response.Write("<a href=""main.asp?action=Close"">����</a>")
					End if
				  %> </td>
              </tr>
              <tr> 
                <td height="24" class="hback"><div align="center" class="hback">ע������</div></td>
                <td class="hback"><%=Fs_User.RegTime%></td>
                <td class="hback"><div align="center" class="hback">���ı�վ</div></td>
                <td class="hback"> <%
				    if Fs_User.ismessage=0 then
						Response.Write("<a href=main.asp?Action=ismessage>δ����</a>")
					Else
						Response.Write("<a href=main.asp?action=Unmessage>�Ѷ���</a>")
					End if
				  %> </td>
              </tr>
              <tr>
                <td height="24" class="hback"><div align="center" class="hback">��½����</div></td>
                <td class="hback"><%=Fs_User.NumLoginNum%></td>
                <td class="hback">&nbsp;</td>
                <td class="hback">&nbsp;</td>
              </tr>
            </table>
              
              <%If  Fs_User.isCorp = 1  then%>
              <table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1" class="table">
                <tr>
                  <td height="26" colspan="4" class="hback_1"><span class="bigtitle"><strong>����ҵ����</strong></span> <a href="corp_info.asp" class="sd">�޸�</a>��<a href="#" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(corpmainId);"  language=javascript>�鿴</a></td>
                </tr>
                <tr style="display:none" id="corpmainId">
                  <td height="26" colspan="4" class="hback">
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <%
				Dim RsCorpObj
				Set RsCorpObj = server.CreateObject(G_FS_RS)
				RsCorpObj.open "select  CorpID,C_Name,C_ShortName,C_Province,C_City,C_Address,C_PostCode,C_Tel,C_Fax,C_BankName,C_BankUserName,isYellowPage,isYellowPageCheck From FS_ME_CorpUser where UserNumber='"& Fs_User.UserNumber&"'",User_Conn,1,1
				if Not RsCorpObj.eof then
				%>
                <tr>
                  <td width="18%" height="24" class="hback"><div align="center">��˾����</div></td>
                  <td width="82%" colspan="3" class="hback"><% = RsCorpObj("C_Name") %>
                  </td>
                </tr>
                <tr>
                  <td height="24" class="hback"><div align="center">��˾���</div></td>
                  <td colspan="3" class="hback"><% = RsCorpObj("C_ShortName") %>
                  </td>
                </tr>
                <tr>
                  <td height="24" class="hback"><div align="center">��˾��ַ</div></td>
                  <td colspan="3" class="hback"><% = RsCorpObj("C_Province") &RsCorpObj("C_City") &RsCorpObj("C_Address")%>
                    ,&nbsp;&nbsp;�ʱࣺ
                    <% = RsCorpObj("C_PostCode") %>
                  </td>
                </tr>
                <tr>
                  <td height="24" class="hback"><div align="center">��˾�绰</div></td>
                  <td colspan="3" class="hback"><% = RsCorpObj("C_Tel") %>
                  </td>
                </tr>
                <tr>
                  <td height="24" class="hback"><div align="center">��˾����</div></td>
                  <td colspan="3" class="hback"><% = RsCorpObj("C_Fax") %>
                  </td>
                </tr>
                <tr>
                  <td height="24" class="hback"><div align="center">�����ʻ�</div></td>
                  <td colspan="3" class="hback">����:
                    <% = RsCorpObj("C_BankName") %>
                    ���ʺ�:
                    <% = RsCorpObj("C_BankUserName") %>
                  </td>
                </tr>
                <tr style="display:none">
                  <td height="24" class="hback"><div align="center">��ҳ��ͨ</div></td>
                  <td colspan="3" class="hback"><%
				  	  if  RsCorpObj("isYellowPage") = 0 then
						  Response.Write("��û��ͨ��ҳ��&nbsp;&nbsp;<a href=""Corp_Info.asp?CorpID="&RsCorpObj("CorpID") &"&YellowPage=Open""><b>��ͨ</b></a>") 
					 Else
						   if RsCorpObj("isYellowPageCheck")=0 then
								Response.Write("�Ѿ���ͨ��ҳ��&nbsp;&nbsp;��ûͨ��")
							Else
								Response.Write("�Ѿ���ͨ��ҳ��&nbsp;&nbsp;<a href=""Corp_Info.asp?CorpID="&RsCorpObj("CorpID") &"&YellowPage=Close""><b>�رջ�ҳ����</b>")
							End if
					 End if
				 %>
                  </td>
                </tr>
                <%
			Else
			     Response.Write("<tr><td height=""40"" colspan=""4"" class=""hback""><b>����ҵ��Ա����û�ҵ���ҵ����</b></a></td></tr>")
			End if
			RsCorpObj.close
			Set RsCorpObj = nothing
			%>
                  </table></td>
                </tr>
              </table>
			  <%else%>
			  <table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1" class="table">
                <tr>
                  <td height="26" colspan="4" class="hback_1"><span class="bigtitle"><strong>��û��ͨ��ҵ���� </strong></span>&nbsp;&nbsp; <a href="OpenCorp.asp" class="sd">��ͨ</a></td>
                </tr>
              </table>
			  
              <%End if%>
<table width="98%" height="76" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
              <tr> 
                <td height="22" colspan="4" class="hback_1"><span class="bigtitle"><strong>����Դ����</strong></span></td>
              </tr>
              <tr> 
                <td width="18%" height="24" class="hback"><div align="center"><a href="infomanage.asp">�ҵ�ר��</a></div></td>
                <td width="82%" colspan="3" class="hback">
				<%
					Dim RsInfoObj
					Set RsInfoObj = server.CreateObject(G_FS_RS)
					RsInfoObj.open "select  Top 4 ClassID,ClassCName,ClassEName,ParentID,UserNumber,ClassTypes,AddTime,ClassContent From FS_ME_InfoClass where ParentID=0 and UserNumber='"& Fs_User.UserNumber&"' order by ClassID desc",User_Conn,1,1
					 if Not RsInfoObj.eof then
						 Do while Not RsInfoObj.eof
							Response.Write "<a href=""ShowInfoClass.asp?ClassID="& RsInfoObj("ClassID") & "&UserNumber=" & RsInfoObj("UserNumber") & """ Title=""" & RsInfoObj("ClassContent") & """>" & RsInfoObj("ClassCName") & "</a>��"
							  RsInfoObj.movenext
						  Loop
						Response.Write "<a href=""infomanage.asp"">More..."	
					Else
						   Response.Write("û��ר����<a href=""infomanage.asp"" title=""����ר��""><b>����</b></a>")
					End if
					RsInfoObj.close:set RsInfoObj = nothing
				  %>
                  </a> </td>
              </tr>
              <tr> 
                <td height="24" class="hback"><div align="center"><a href="GroupClass.asp">�ҵ���Ⱥ</a></div></td>
                <td colspan="3" class="hback"> <%
					Dim RsGroupObj
					Set RsGroupObj = server.CreateObject(G_FS_RS)
					RsGroupObj.open "select  Top 4 gdID,ClassID,Title,Content,UserNumber,isLock From FS_ME_GroupDebateManage where UserNumber='"& Fs_User.UserNumber&"' order by ClassID desc",User_Conn,1,1
					 if Not RsGroupObj.eof then
						 Do while Not RsGroupObj.eof
							  Response.Write "<a href=""Group_unit.asp?GDID="& RsGroupObj("gdID")& """ Title=""" & RsGroupObj("Content") & """>" & RsGroupObj("Title") & "</a>��"&vbNewLine
							  'if cbool(RsGroupObj("isLock")) then response.Write(" �ȴ����... ")
							  'Response.Write "<a href=""GroupClass.asp?Act=Edit&gdID="& RsGroupObj("gdID")& """>���޸ġ�</a>&nbsp;"
							  RsGroupObj.movenext
						  Loop	
						Response.Write "<a href=""GroupClass.asp"">More..."	
					Else
						   Response.Write("û����Ⱥ��<a href=""GroupClass.asp?Act=Add"" title=""������Ⱥ""><b>����</b></a>")
					End if
					RsGroupObj.close:set RsGroupObj = nothing
				  %> </td>
              </tr>
            </table>
<table width="98%" height="49" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr>
    <td height="22" class="hback_1"><span class="bigtitle"><strong>���ҵ���־</strong></span></td>
    </tr>
  <tr>
    <td height="24" class="hback">
	<div align="left">
	<%
	Dim rs,sLog
	set rs = Server.CreateObject(G_FS_RS)
	rs.open "select top 8 * from FS_ME_Infoilog where UserNumber='"& Fs_User.UserNumber&"'  and isDraft=0 order by isTop desc,AddTime desc,iLogID desc",User_Conn,1,3
	do while not rs.eof
		sLog = sLog &  "<img src=""images/dot.gif"" border=""0"" /><a href=i_Blog/PublicLogEdit.asp?id="&Rs("iLogID")&" target=""_self"">"&Rs("Title")&"</a>"
		sLog = sLog &  "<span class=""hback_1"" style=""font-size:10px"">("&Rs("AddTime")&")</span><br />"
	rs.movenext
	loop
	rs.close:set rs = nothing
	Response.Write sLog
	%>
	</div>
      </a> </td>
    </tr>
</table></td>
            <td valign="top"> <table width="98%" height="176" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
                <tr> 
                  <td height="27"  class="hback_1"><span class="bigtitle"><strong>�����¹���</strong></span></td>
                </tr>
                <tr> 
                  <td valign="top" class="hback">
				  <%
					Dim RsNewsObj
					Set RsNewsObj = server.CreateObject(G_FS_RS)
					RsNewsObj.open "select  Top 6 NewsID,Title,Content,Addtime,GroupID,NewsPoint,isLock From FS_ME_News where Islock=0 order by NewsID desc",User_Conn,1,1
					 if Not RsNewsObj.eof then
						 Do while Not RsNewsObj.eof
							Response.Write "<img src=""images/dot.gif"" border=""0""></img><a href=""ShowCallboard.asp?NewsID="& RsNewsObj("NewsID") & """ Title=""" & RsNewsObj("Content") & """>" & RsNewsObj("Title") & "</a>&nbsp;<font style=""font-size:9px"">("& year(RsNewsObj("Addtime"))&"-"&month(RsNewsObj("Addtime"))&"-"&day(RsNewsObj("Addtime")) &")</font><br>"
							  RsNewsObj.movenext
						  Loop	
					Else
						   Response.Write("û�й���")
					End if
					RsNewsObj.close:set RsNewsObj = nothing
				  %>
                    <table width="98%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><div align="right"><a href="callboard.asp">&gt;&gt;more..</a></div></td>
                      </tr>
                    </table> </td>
                </tr>
              </table>
              
            <table width="98%" height="198" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
              <tr> 
                  <td width="50%" height="27"  class="hback_1"><span class="bigtitle"><strong>���ҵĺ���</strong></span></td>
                </tr>
                <tr> 
                  
                <td height="167" valign="top" class="hback"> 
                  <%
					Dim RsFriendObj
					Set RsFriendObj = server.CreateObject(G_FS_RS)
					RsFriendObj.open "select  Top 7 FriendID,UserNumber,FriendType,F_UserNumber,AddTime,Updatetime From FS_ME_Friends where UserNumber='"& Fs_User.UserNumber &"' and FriendType = 0 order by UpdateTime desc,FriendID desc",User_Conn,1,3
					 if Not RsFriendObj.eof then
					 	  Response.Write("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
						 Do while Not RsFriendObj.eof
						 	Dim Tmp_UserNumber,Tmp_UserName,strsendMail,RsTmpobj
							Set RsTmpobj = server.CreateObject(G_FS_RS)
							RsTmpobj.open "Select  top 1 NickName,UserName From FS_ME_Users where UserNumber = '"& RsFriendObj("F_UserNumber") &"'",User_Conn,1,1
							If Not RsTmpobj.eof then
								Tmp_UserNumber =RsTmpobj("NickName") 
								Tmp_UserName  =RsTmpobj("UserName") 
								Strsendmail ="<a href=message_write.asp?ToUserNumber="&RsFriendObj("F_UserNumber")&" title=""������Ϣ"">����</a>&nbsp;|&nbsp;<a href=book_write.asp?ToUserNumber="&RsFriendObj("F_UserNumber")&"&M_Type=0 title=""��������"">����</a>"
									Response.Write "<tr><td width=""70%""><img src=""images/dot.gif"" border=""0""></img><a href=""ShowUser.asp?UserNumber="& RsFriendObj("F_UserNumber") & """ target=""_blank"">" & Tmp_UserNumber & "("&Tmp_UserName&")</td><td align=""center"">"& Strsendmail &"</a></td></tr>"
									  RsFriendObj.movenext
							Else
									RsFriendObj.movenext
							End if
						Loop
						Response.Write("</table>")	
					Else
						   Response.Write("û�к���")
					End if
					set RsTmpobj = nothing
					RsFriendObj.close
					set RsFriendObj = nothing
				  %>
                  </td>
                </tr>
              </table>
				
            <table width="98%" height="124" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
              <tr> 
                <td width="50%" height="27"  class="hback_1"><span class="bigtitle"><strong>������</strong></span>(<a href="Top_User.asp?type=int" target="sys_Frame">����</a>��<a href="Top_User.asp?type=money" target="sys_Frame">���</a>��</span><a href="Top_User.asp?type=active" target="sys_Frame">���Ծ</a>��<a href="Top_User.asp?type=hits" target="sys_Frame">����</a>)</td>
              </tr>
              <tr> 
                <td height="94" valign="top" class="hback"><iframe src="top_user.asp?type=int" name="sys_Frame" scrolling="no" frameborder="0" align="middle" width="100%" height="175"></iframe></td>
              </tr>
            </table>
            
          </td>
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






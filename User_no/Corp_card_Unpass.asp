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
User_GetParm
If Request.Form("Action") = "Save" then
	Dim DelID,Str_Tmp,Str_Tmp1
	DelID = request.Form("CorpCardID")
	if DelID = "" then 
		strShowErr = "<li>�����ѡ��һ����ɾ��</li>"
		Call ReturnError(strShowErr,"")
	End if
	User_Conn.execute("Delete From  FS_ME_CorpCard where CorpCardID in ("&FormatIntArr(DelID)&")")
	strShowErr = "<li>ɾ����Ƭ�ɹ�</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
if Request.QueryString("Action") = "Pass" then
	dim RsaddsObj,RsaddPassObj
	Set RsaddsObj = server.CreateObject(G_FS_RS)
	RsaddsObj.open "select CorpCardID,UserNumber,F_UserNumber From FS_ME_CorpCard where CorpCardID="&CintStr(Request.QueryString("Id")),User_Conn,1,3
	Set RsaddPassObj = server.CreateObject(G_FS_RS)
	RsaddPassObj.open "select * From FS_ME_CorpCard where 1=0",User_Conn,1,3
	RsaddPassObj.addnew
	RsaddPassObj("UserNumber") = Fs_User.UserNumber
	RsaddPassObj("F_UserNumber") = RsaddsObj("UserNumber")
	RsaddPassObj("Addtime") = now
	RsaddPassObj("islock") = 0
	RsaddPassObj("Content") = ""
	RsaddPassObj.update
	RsaddPassObj.close
	set RsaddPassObj  = nothing
	User_Conn.execute("Update  FS_ME_CorpCard set islock=0  where CorpCardID="&CintStr(Request.QueryString("Id")))
	RsaddsObj.close
	set RsaddsObj = nothing
	strShowErr = "<li>�Ѿ��ͶԷ�������Ƭ</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-��Ƭ����</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; ��Ƭ���� </td>
          </tr>
        </table>
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="">
          <tr class="hback"> 
            <td colspan="4" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="44%"> ��������<strong> 
                    <%
				Dim RsCardObj,RsCardSQL
				Dim strpage
				strpage=NoSqlHack(request("page"))
				if len(strpage)=0 Or strpage<1 or trim(strpage)=""  Then strpage="1"
				Set RsCardObj = Server.CreateObject(G_FS_RS)
				RsCardSQL = "Select CorpCardID,UserNumber,F_UserNumber,AddTime,Content,isLock From FS_ME_CorpCard  where F_UserNumber='"&Fs_User.UserNumber&"' and islock = 1 Order by CorpCardID desc"
				RsCardObj.Open RsCardSQL,User_Conn,1,3
				Response.Write "<Font color=red>" & RsCardObj.RecordCount&"</font>"
				%>
                    </strong> �� ��Ƭ</td>
                  <td width="56%"><div align="center"><%if p_isPassCard = 1 Then Response.Write("<a href=Corp_Card_unpass.asp><b>�鿴�����ҽ�����Ƭ���û�</b></a>")%></div></td>
                </tr>
              </table></td>
          </tr >
          <tr class="hback"> 
            <td class="xingmu">�û����</td>
            <td class="xingmu">���ʱ��</td>
            <td class="xingmu"><div align="center">��ע��Ƭ���</div></td>
            <td class="xingmu"><div align="center">����</div></td>
          </tr>
          <%
		Dim select_count,select_pagecount,i
		if RsCardObj.eof then
			   RsCardObj.close
			   set RsCardObj=nothing
			   Response.Write"<TR><TD colspan=""4""  class=""hback"">û�м�¼��</TD></TR>"
		else
				RsCardObj.pagesize = 15
				RsCardObj.absolutepage=cint(strpage)
				select_count=RsCardObj.recordcount
				select_pagecount=RsCardObj.pagecount
				for i=1 to RsCardObj.pagesize
					if RsCardObj.eof Then exit For 
						Dim Returvaluestr
						Returvaluestr = Fs_User.GetFriendName(RsCardObj("UserNumber"))
					if RsCardObj("F_UserNumber") = "0" then
						  exit For 
					Else
						Dim RsGetCardUObj
						Set RsGetCardUObj = Server.CreateObject(G_FS_RS)
						RsGetCardUObj.open "select  qq,msn,UserNumber From FS_ME_Users where UserNumber='"& RsCardObj("UserNumber") &"'",User_Conn,1,1
						if RsGetCardUObj.eof then exit For 
						Dim RsGetCardObj
						Set RsGetCardObj = Server.CreateObject(G_FS_RS)
						RsGetCardObj.open "select C_logo,C_Name,C_ConactName,C_Vocation,C_Tel,C_Fax,C_WebSite,C_Province,C_City,C_address,C_PostCode,UserNumber From FS_ME_CorpUser where UserNumber='"& RsCardObj("UserNumber") &"'",User_Conn,1,1
						if RsGetCardObj.eof then exit For 
		%>
          <tr class="hback"> 
            <td width="41%" class="hback">�� 
              <% = Returvaluestr%>
              &nbsp;(
              <% = RsGetCardObj("C_Name")%>
              ) </td>
            <td width="23%" class="hback"><div align="left"><a href="ShowUser.asp?UserNumber=<% = RsCardObj("F_UserNumber")%>"> 
                <% = RsCardObj("addtime")%>
                </a></div></td>
            <td width="20%" class="hback"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(Cid<%=RsCardObj("CorpCardID")%>);"  language=javascript><div align="center">�鿴�Է���Ƭ</div></td>
            <td width="16%" class="hback"><div align="center"><a href="Corp_Card_Unpass.asp?id=<%=RsCardObj("CorpCardID")%>&Action=Pass"  onClick="{if(confirm('ȷ��ͬ��������Ƭ��?')){this.document.inbox.submit();return true;}return false;}">ͬ�⽻��</a>&nbsp;|&nbsp; 
                <input name="CorpCardID" type="checkbox" id="CorpCardID" value="<% = RsCardObj("CorpCardID")%>">
              </div></td>
          </tr>
          <tr class="hback" style="display:none" id="Cid<%=RsCardObj("CorpCardID")%>"> 
            <td height="46" colspan="4" class="hback"><table width="100%" border="0" cellspacing="1" cellpadding="5" class="table">
                <tr> 
                  <td width="45%" class="hback_1">&nbsp;</td>
                  <td width="55%" class="hback_1"> <table width="100%" border="0" cellspacing="1" cellpadding="2" class="table">
                      <tr> 
                        <td width="35%" rowspan="7" class="hback"> <div align="center"><a href="ShowUser.asp?UserNumber=<% = RsGetCardObj("UserNumber")%>" target="_blank"> 
                            <%if Trim(RsGetCardObj("C_logo")) <>"" then%>
                            <img src="<% = RsGetCardObj("C_logo")%>" alt="�鿴��ϸ����" border="0"></img> 
                            <%Else%>
                            <img src="images/nologo.gif" alt="�鿴��ϸ����" border="0"></img> 
                            <%End if%>
                            </a> </div></td>
                        <td width="65%" class="hback"><span class="bigtitle"><b> 
                          <% = RsGetCardObj("C_Name")%>
                          </b></span></td>
                      </tr>
                      <tr> 
                        <td class="hback"><strong> 
                          <% = RsGetCardObj("C_ConactName")%>
                          </strong>( <% = RsGetCardObj("C_Vocation")%>
                          )</td>
                      </tr>
                      <tr> 
                        <td class="hback"> �绰: 
                          <% = RsGetCardObj("C_Tel")%> </td>
                      </tr>
                      <tr> 
                        <td class="hback">���� : 
                          <% = RsGetCardObj("C_Fax")%> </td>
                      </tr>
                      <tr> 
                        <td class="hback">��ҳ: <a href="<% = RsGetCardObj("C_WebSite")%>"> 
                          <% = RsGetCardObj("C_WebSite")%>
                          </a> </td>
                      </tr>
                      <tr> 
                        <td class="hback">��ַ: 
                          <% = RsGetCardObj("C_Province")%> <% = RsGetCardObj("C_City")%> <% = RsGetCardObj("C_Address")%>
                          �� <% = RsGetCardObj("C_PostCode")%> </td>
                      </tr>
                      <tr> 
                        <td class="hback">QQ: 
                          <% = RsGetCardUObj("QQ")%> </td>
                      </tr>
                      <tr> 
                        <td width="35%" class="hback"> <div align="center"> 
                            <%
						if  Len(Trim(RsGetCardUObj("QQ")))>4 then
							Dim sOICQ
						    sOICQ ="<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin="& RsGetCardUObj("QQ") &"&Site=FoosunCMS&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:"& RsGetCardUObj("QQ") &":16 alt=""��������"& RsGetCardUObj("QQ") &"����Ϣ""></a>"
							Response.Write sOICQ
						Else
							Response.Write("û��QQ")
						End if
						%>
                          </div></td>
                        <td class="hback">msn: 
                          <% = RsGetCardUObj("msn")%> </td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
          <%
				End if
			  RsCardObj.MoveNext
		  Next
		  %>
          <tr class="hback"> 
            <td colspan="4" class="hback_1"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="80%"> <span class="top_navi"> 
                    <% 
							Response.write"&nbsp;��<b>"& select_pagecount &"</b>ҳ<b>&nbsp;" & select_count &"</b>����¼����ҳ�ǵ�<b>"& strpage &"</b>ҳ��"
							if int(strpage)>1 then
								Response.Write"&nbsp;<a href=?page=1>��һҳ</a>&nbsp;&nbsp;"
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&">��һҳ</a>&nbsp;&nbsp;"
							End if
							If int(strpage)<select_pagecount then
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&">��һҳ</a>&nbsp;"
								Response.Write"&nbsp;<a href=?page="& select_pagecount &">���һҳ</a>&nbsp;&nbsp;"
							End if
								Response.Write"<br>"
								RsCardObj.close
								Set RsCardObj=nothing
							End if
							%>
                    </SPAN></td>
                </tr>
              </table></td>
          </tr>
          <tr class="hback"> 
            <td colspan="4" class="hback"><div align="right"> 
                <input name="Action" type="hidden" id="Action" value="Save">
                <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
                ѡ��������Ƭ 
                <input type="submit" name="Submit" value="ɾ����ѡ�����Ŀ"  onClick="{if(confirm('ȷ���������ѡ��ļ�¼��?')){this.document.form1.submit();return true;}return false;}">
              </div></td>
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
set RsGetCardObj = nothing
set RsGetCardUObj = nothing

Set Fs_User = Nothing
%>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = form.elements[i];  
    if (e.name != 'chkall')  
       e.checked = form.chkall.checked;  
    }  
  }
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->






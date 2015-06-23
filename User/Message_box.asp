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
dim ShowChar,strAction,ShowChar_1
if NoSqlHack(Request.QueryString("type"))="rebox" then
 	ShowChar = "收件箱"
	ShowChar_1 = "发件人"
 	strAction = "rebox"
Elseif  NoSqlHack(Request.QueryString("type"))="drabox" then
 	ShowChar = "草稿箱"
	ShowChar_1 = "收件人"
 	strAction = "drabox"
Elseif  NoSqlHack(Request.QueryString("type"))="sendbox" then
 	ShowChar = "发件箱"
	ShowChar_1 = "收件人"
 	strAction = "sendbox"
End if
If Request.Form("Action") = "Del" then
	Dim DelID,Str_Tmp,Str_Tmp1
	DelID = FormatIntArr(request.Form("MessageID"))
	if DelID = "" then 
		strShowErr = "<li>你必须选择一项再删除</li>"
		Call ReturnError(strShowErr,"")
	End if
	if Trim(Request.Form("strAction")) = "drabox" then
		User_Conn.execute("Delete From FS_ME_Message where MessageId in ("&DelID&") and M_ReadUserNumber ='"& Fs_User.UserNumber&"'")
	Elseif  Trim(Request.Form("strAction")) = "sendbox"  then
			Dim RsTFSQL,RsTFObj
			Set RsTFObj = Server.CreateObject(G_FS_RS)
			RsTFSQL = "Select isDelF  From FS_ME_Message  where  MessageId in ("&DelID&") "
			RsTFObj.Open RsTFSQL,User_Conn,1,3
			if RsTFObj("isDelF") = 1 then
				User_Conn.execute("Delete From FS_ME_Message where MessageId in ("&DelID&")")
			Else
				User_Conn.execute("Update FS_ME_Message set isDelR = 1  where MessageId in ("&DelID&") and M_ReadUserNumber ='"& Fs_User.UserNumber&"'")
			End if
	Elseif   Trim(Request.Form("strAction")) = "rebox"   then
			Dim RsTFSQL1,RsTFObj1
			Set RsTFObj1 = Server.CreateObject(G_FS_RS)
			RsTFSQL1 = "Select isDelF  From FS_ME_Message  where  MessageId in ("&DelID&") "
			RsTFObj1.Open RsTFSQL1,User_Conn,1,3
			if RsTFObj1("isDelF") = 1 then
				User_Conn.execute("Delete From FS_ME_Message where MessageId in ("&DelID&") and M_ReadUserNumber ='"& Fs_User.UserNumber&"'")
			Else
				User_Conn.execute("Update FS_ME_Message set isDelR = 1  where MessageId in ("&DelID&") and M_ReadUserNumber ='"& Fs_User.UserNumber&"'")
			End if
	Else
		strShowErr = "<li>错误的参数</li>"
		Call ReturnError(strShowErr,"")
	'User_Conn.execute("Update FS_ME_Message set isDelR = 1  where MessageId in ("&DelID&") and M_ReadUserNumber ='"& Fs_User.UserNumber&"'")
	End if
	strShowErr = "<li>删除短信成功</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-短信</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt; 短信－<% = ShowChar %></td>
          </tr>
        </table>
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="">
          <tr class="hback"> 
            <td colspan="12" class="hback">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr class="hback"> 
                  <td width="27%" class="hback">  
                    <%
				Dim RsMessageObj,RsMessageSQL
				Dim strpage,strSQLs
				strpage=request("page")
				if len(strpage)=0 Or strpage<1 or trim(strpage)=""  Then strpage="1"
				Set RsMessageObj = Server.CreateObject(G_FS_RS)
				if NoSqlHack(Request.QueryString("type"))="rebox" then
					RsMessageSQL = "Select MessageID,M_Title,M_FromUserNumber,M_ReadUserNumber,M_Content,M_FromDate,M_ReadTF,isRecyle,isDelF,isDelR,LenContent,isSend,isDraft From FS_ME_Message  where M_ReadUserNumber='"&Fs_User.UserNumber&"'  and isRecyle=0 and isDelR=0 Order by MessageID desc"
				Elseif  NoSqlHack(Request.QueryString("type"))="drabox" then
					RsMessageSQL = "Select MessageID,M_Title,M_FromUserNumber,M_ReadUserNumber,M_Content,M_FromDate,M_ReadTF,isRecyle,isDelF,isDelR,LenContent,isSend,isDraft From FS_ME_Message  where M_FromUserNumber='"&Fs_User.UserNumber&"'  and isRecyle=0 and isDelR=0 and isDraft=1 Order by MessageID desc"
				Elseif  NoSqlHack(Request.QueryString("type"))="sendbox" then
					RsMessageSQL = "Select MessageID,M_Title,M_FromUserNumber,M_ReadUserNumber,M_Content,M_FromDate,M_ReadTF,isRecyle,isDelF,isDelR,LenContent,isSend,isDraft From FS_ME_Message  where M_FromUserNumber='"&Fs_User.UserNumber&"'  and isRecyle=0 and isDelF=0 and issend=1 Order by MessageID desc"
				Else
					RsMessageSQL = "Select MessageID,M_Title,M_FromUserNumber,M_ReadUserNumber,M_Content,M_FromDate,M_ReadTF,isRecyle,isDelF,isDelR,LenContent,isSend,isDraft From FS_ME_Message  where M_ReadUserNumber='"&Fs_User.UserNumber&"'  and isRecyle=0 and isDelR=0 Order by MessageID desc"
				End if
				RsMessageObj.Open RsMessageSQL,User_Conn,1,3
				%>
                    <a href="Message_box.asp?type=rebox"><img src="images/Skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>//recievebox.gif" width="40" height="40" border="0"></a><a href="Message_box.asp?type=sendbox"><img src="images/Skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>//sendbox.gif" width="40" height="40" border="0"></a><a href="Message_box.asp?type=drabox"><img src="images/Skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>//draftbox.gif" width="40" height="40" border="0"></a><a href="Message_write.asp"><img src="images/Skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>//writemessage.gif" width="40" height="40" border="0"></a></td>
                  <td width="14%" class="hback"><div align="left">空间占用:
                      <%
				     Dim UnTotle,FS_Message_1
					 Set FS_Message_1 = new Cls_Message
					UnTotle=FS_Message_1.LenContent(Fs_User.UserNumber)/(1024*100)*100
					Set FS_Message_1 = Nothing 
					If IsNull(UnTotle) then UnTotle=0
					Response.Write Formatnumber(UnTotle,2,-1)&"%"
					%>
                    </div></td>
                  <td width="59%" class="hback"> 
                    <table width="100%" height="17" border="0" cellpadding="0" cellspacing="1" class="table">
                      <tr> 
                        <td class="hback_1"><img src="images/space_pic_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.gif" width="<% = Formatnumber((UnTotle),2,-1)%>%" height="17"></td>
                      </tr>
                    </table> </td>
                </tr>
              </table></td>
          </tr class="hback">
          <%
		Dim select_count,select_pagecount,i
		if RsMessageObj.eof then
			   RsMessageObj.close
			   set RsMessageObj=nothing
			   Response.Write"<TR  class=""hback""><TD colspan=""7""  class=""hback"" height=""40"">没有记录。</TD></TR>"
		else
				RsMessageObj.pagesize = 20
				RsMessageObj.absolutepage=CintStr(strpage)
				select_count=RsMessageObj.recordcount
				select_pagecount=RsMessageObj.pagecount
		  %>
          <tr class="hback"> 
            <td width="5%" height="22" class="xingmu"><div align="left"><strong>已读</strong></div></td>
            <td width="15%" class="xingmu"><strong>
<% = ShowChar_1 %>
              </strong></td>
            <td width="36%" height="22" class="xingmu"><div align="left"><strong>标题</strong></div></td>
            <td width="20%" height="22" class="xingmu"><div align="left"><strong>日期</strong></div></td>
            <td width="11%" height="22" class="xingmu"><strong>长度</strong></td>
            <td width="7%" height="22" class="xingmu"><div align="center">查看</div></td>
            <td width="6%" height="22" class="xingmu"><div align="center"><strong>操作</strong></div></td>
          </tr>
          <%
				for i=1 to RsMessageObj.pagesize
					if RsMessageObj.eof Then exit For 
					Dim Returvaluestr_R,Returvaluestr_F,strbstat,strben,strcss,strReadTF
					if RsMessageObj("M_ReadTF") =0 then 
						strbstat = "<b>"
						strben = "</b>"
						strcss = "hback"
						strReadTF = "<font color=red><b>×</b></font>"
					Else
						strbstat = ""
						strben = ""
						strcss = "hback"
						strReadTF = "<font color=#999999><b>√</b></font>"
					End if
					if Request.QueryString("type")="rebox" then
						Returvaluestr_R = Fs_User.GetFriendName(RsMessageObj("M_FromUserNumber"))
						if Trim(RsMessageObj("M_FromUserNumber")) <> "0" then
							Returvaluestr_F = "<a href=ShowUser.asp?UserNumber="& RsMessageObj("M_FromUserNumber") &">"&Fs_User.GetFriendName(RsMessageObj("M_FromUserNumber"))&"</a>"
						Else
							Returvaluestr_F = "用户不存在"
						End if	
					Else
						Returvaluestr_R = Fs_User.GetFriendName(RsMessageObj("M_ReadUserNumber"))
						if Trim(RsMessageObj("M_ReadUserNumber")) <> "0" then
							Returvaluestr_F = "<a href=ShowUser.asp?UserNumber="& RsMessageObj("M_ReadUserNumber") &">"&Fs_User.GetFriendName(RsMessageObj("M_ReadUserNumber"))&"</a>"
						Else
							Returvaluestr_F = "用户不存在"
						End if	
					ENd if
		%>
          <tr class="hback"> 
            <td class="<% = strcss %>"><div align="center"><% = strReadTF%> </div></td>
            <td class="<% = strcss %>">
              <% =   Returvaluestr_F %>
            </td>
            <td class="<% = strcss %>"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(mid<% = RsMessageObj("MessageID")%>);"  language=javascript><% = strbstat & RsMessageObj("M_title") & strben %></td>
            <td class="<% = strcss %>"><% =  RsMessageObj("M_FromDate")  %></td>
            <td class="<% = strcss %>"><% =  RsMessageObj("LenContent")  %>
              Byte</td>
            <td class="<% = strcss %>"> 
              <div align="center">
                <%
							if NoSqlHack(Request.QueryString("type"))="rebox" then
								Response.Write "<a href=""Message_Read.asp?MessageID="& RsMessageObj("MessageID") &"&strstat="&strAction&""">回复 </a>"
							Elseif  NoSqlHack(Request.QueryString("type"))="drabox" then
								Response.Write "<a href=""Message_Read.asp?MessageID="& RsMessageObj("MessageID") &"&strstat="&strAction&""">发送 </a>"
							Elseif  NoSqlHack(Request.QueryString("type"))="sendbox" then
								Response.Write "<a href=""Message_Read.asp?MessageID="& RsMessageObj("MessageID") &"&strstat="&strAction&""">发送 </a>"
							End if
						%>
              </div></td>
            <td class="<% = strcss %>"><input name="MessageID" type="checkbox" id="MessageID" value="<% = RsMessageObj("MessageID")%>"></td>
          </tr>
          <tr class="hback" id="mid<% = RsMessageObj("MessageID")%>" style="display:none"> 
            <td colspan="12" class="hback"><table width="100%" height="62" border="0" cellpadding="5" cellspacing="1" class="table">
                <tr> 
                  <td height="60" valign="top" class="hback_1"> <a href="Message_Read.asp?MessageID = <%  = RsMessageObj("MessageID") %>"> 
                    </a> <table width="100%" border="0" cellspacing="0" cellpadding="4">
                      <tr> 
                        <td height="60" valign="top"> 
                          <% = RsMessageObj("M_Content")%>
                        </td>
                      </tr>
                      <tr> 
                        <td><div align="right"> 
                            <%
							if NoSqlHack(Request.QueryString("type"))="rebox" then
								Response.Write "<a href=""Message_Read.asp?MessageID="& RsMessageObj("MessageID") &"&strstat="&strAction&""">回复该短信 </a>"
							Elseif  NoSqlHack(Request.QueryString("type"))="drabox" then
								Response.Write "<a href=""Message_Read.asp?MessageID="& RsMessageObj("MessageID") &"&strstat="&strAction&""">发送此短信 </a>"
							Elseif  NoSqlHack(Request.QueryString("type"))="sendbox" then
								Response.Write "<a href=""Message_Read.asp?MessageID="& RsMessageObj("MessageID") &"&strstat="&strAction&""">发送此短信 </a>"
							End if
						%>
                          </div></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
          <%
			  RsMessageObj.MoveNext
		  Next
		  %>
          <tr class="hback"> 
            <td colspan="12"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr  class="hback"> 
                  <td colspan="2"> <% 	Response.Write("每页:"& RsMessageObj.pagesize &"个,")
							Response.write"&nbsp;共<b>"& select_pagecount &"</b>页<b>&nbsp;" & select_count &"</b>条记录，本页是第<b>"& strpage &"</b>页。"
							if int(strpage)>1 then
								Response.Write"&nbsp;<a href=?page=1>第一页</a>&nbsp;&nbsp;"
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&">上一页</a>&nbsp;&nbsp;"
							End if
							If int(strpage)<select_pagecount then
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&">下一页</a>&nbsp;"
								Response.Write"&nbsp;<a href=?page="& select_pagecount &">最后一页</a>&nbsp;&nbsp;"
							End if
								Response.Write"<br>"
								RsMessageObj.close
								Set RsMessageObj=nothing
							End if
							%> <div align="right"> </div></td>
                </tr>
                <tr  class="hback"> 
                  <td width="64%"><div align="right">节省每一分空间，请及时删除无用信息 
                      <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
                      选中所有短信 
                      <input name="Action" type="hidden" id="Action" value="Del">
                      <input name="strAction" type="hidden" id="strAction" value="<% = strAction%>">
                      　 </div></td>
                  <td width="18%"><input type="submit" name="Submit" value="删除选中的短信" onClick="{if(confirm('确定清除您所选择的记录吗？?')){this.document.form1.submit();return true;}return false;}"> 
                  </td>
                </tr>
              </table></td>
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->
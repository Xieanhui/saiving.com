<meta http-equiv="Content-Type" content="text/html; charset=gb2312" /><table width="100%" height="26" border="0" cellpadding="2" cellspacing="0">
  <tr> 
    <td width="66%" height="23"> <div align="left"> 
<%
Dim TotleMessage,FS_Message_top
Set FS_Message_top = new Cls_Message
FS_Message_top.UserName = Fs_User.UserNumber
TotleMessage = FS_Message_top.Number
Set FS_Message_top = Nothing 
If top_strisPromptTF = 1 then 
	if Fs_User.NumFS_Money < top_strisPromptnNum then
		Response.Write("<span class=""top_navi"">欢迎您：<b>"& Fs_User.UserName &"</b>,&nbsp;&nbsp;您的金币小于"& NoSqlHack(Top_strisPromptnNum) & NoSqlHack(top_moneyName) &",&nbsp;&nbsp;请尽快冲值,&nbsp;&nbsp;</span>&nbsp;<a href="&s_savepath&"/card.asp class=""top_navi""><b>冲值</b></a>&nbsp;&nbsp;")
	Else
		Response.Write("<span class=""top_navi"">欢迎您：<b>"& Fs_User.UserName &"</b></span>&nbsp;&nbsp;")
	End if
End if  
If TotleMessage=0 then
	Response.Write("<a href="&s_savepath&"/Message_box.asp?type=rebox class=""top_navi"">短消息(0)</a>&nbsp;&nbsp;")
Else
	Response.Write("<a href="&s_savepath&"/Message_box.asp?type=rebox class=""navitx""><b>短消息("&TotleMessage&")</font></a>&nbsp;&nbsp;")
End If
%>
      </div></td>
    <td width="34%"  class=menu> <div align="center"><A class="top_navi" onmouseover=showmenu(event,0,1,false) title=基本信息   onmouseout=delayhidemenu() href="<%=s_savepath%>/MyInfo.asp"  target=_self>基本信息</A><span class="top_navi">┊</span><A  onmouseover=showmenu(event,1,1,false)  class="top_navi"    title=账号信息 onmouseout=delayhidemenu()   href="<%=s_savepath%>/MyAccount.asp" target=_self>账号管理</A><span class="top_navi">┊<A class="top_navi" onmouseover=showmenu(event,3,1,false) title=交易管理   onmouseout=delayhidemenu() href="<%=s_savepath%>/Order.asp"  target=_self>交易管理</A>
	
	<%if Fs_User.isCorp = 0 then%>
	┊<A class="top_navi" title=稿件管理 href="<%=s_savepath%>/contr/contrManage.asp" target=_self>稿件管理</A></span><%End if%> 
        <%if Fs_User.isCorp = 1 then%>
        <span class="top_navi">┊<a class="top_navi" onMouseOver=showmenu(event,2,1,false) title=企业信息  onMouseOut=delayhidemenu() href="<%=s_savepath%>/Corp_Info.asp"  target=_self>企业信息</a></span>
<%End if%>
<%if Fs_User.isCorp = 1 then%>
	┊<A class="top_navi" title=稿件管理 href="<%=s_savepath%>/contr/contrManage.asp" target=_self>稿件管理</A>
	<%End if%> </div></td>
	   <DIV class=menuskin id=popmenu onmouseover="clearhidemenu();highlightmenu(event,'on')" style="Z-INDEX: 100"  onmouseout="highlightmenu(event,'off');dynamichide(event)" divalpha></DIV>
    
  </tr>
</table>







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
		Response.Write("<span class=""top_navi"">��ӭ����<b>"& Fs_User.UserName &"</b>,&nbsp;&nbsp;���Ľ��С��"& NoSqlHack(Top_strisPromptnNum) & NoSqlHack(top_moneyName) &",&nbsp;&nbsp;�뾡���ֵ,&nbsp;&nbsp;</span>&nbsp;<a href="&s_savepath&"/card.asp class=""top_navi""><b>��ֵ</b></a>&nbsp;&nbsp;")
	Else
		Response.Write("<span class=""top_navi"">��ӭ����<b>"& Fs_User.UserName &"</b></span>&nbsp;&nbsp;")
	End if
End if  
If TotleMessage=0 then
	Response.Write("<a href="&s_savepath&"/Message_box.asp?type=rebox class=""top_navi"">����Ϣ(0)</a>&nbsp;&nbsp;")
Else
	Response.Write("<a href="&s_savepath&"/Message_box.asp?type=rebox class=""navitx""><b>����Ϣ("&TotleMessage&")</font></a>&nbsp;&nbsp;")
End If
%>
      </div></td>
    <td width="34%"  class=menu> <div align="center"><A class="top_navi" onmouseover=showmenu(event,0,1,false) title=������Ϣ   onmouseout=delayhidemenu() href="<%=s_savepath%>/MyInfo.asp"  target=_self>������Ϣ</A><span class="top_navi">��</span><A  onmouseover=showmenu(event,1,1,false)  class="top_navi"    title=�˺���Ϣ onmouseout=delayhidemenu()   href="<%=s_savepath%>/MyAccount.asp" target=_self>�˺Ź���</A><span class="top_navi">��<A class="top_navi" onmouseover=showmenu(event,3,1,false) title=���׹���   onmouseout=delayhidemenu() href="<%=s_savepath%>/Order.asp"  target=_self>���׹���</A>
	
	<%if Fs_User.isCorp = 0 then%>
	��<A class="top_navi" title=������� href="<%=s_savepath%>/contr/contrManage.asp" target=_self>�������</A></span><%End if%> 
        <%if Fs_User.isCorp = 1 then%>
        <span class="top_navi">��<a class="top_navi" onMouseOver=showmenu(event,2,1,false) title=��ҵ��Ϣ  onMouseOut=delayhidemenu() href="<%=s_savepath%>/Corp_Info.asp"  target=_self>��ҵ��Ϣ</a></span>
<%End if%>
<%if Fs_User.isCorp = 1 then%>
	��<A class="top_navi" title=������� href="<%=s_savepath%>/contr/contrManage.asp" target=_self>�������</A>
	<%End if%> </div></td>
	   <DIV class=menuskin id=popmenu onmouseover="clearhidemenu();highlightmenu(event,'on')" style="Z-INDEX: 100"  onmouseout="highlightmenu(event,'off');dynamichide(event)" divalpha></DIV>
    
  </tr>
</table>







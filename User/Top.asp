<%
dim rs_sys_top,top_moneyName,top_strisPromptTF,top_strisPromptnNum,menu_RssFeed,top_GroupID,rs_Group
set rs_sys_top=User_Conn.execute("select top 1 MoneyName,isPrompt,RssFeed From FS_ME_SysPara")
top_moneyName=rs_sys_top(0)
top_strisPromptTF= cint(trim(split(rs_sys_top(1),",")(0)))
top_strisPromptnNum=clng(trim(split(rs_sys_top(1),",")(1)))
menu_RssFeed=rs_sys_top(2)
rs_sys_top.close:set rs_sys_top=nothing
set rs_Group = User_Conn.execute("select GroupName From FS_ME_Group where GroupID="&FS_User.NumGroupID)
if rs_Group.eof then
	top_GroupID = "--"
	rs_Group.close:set rs_Group = nothing
else
	top_GroupID = rs_Group(0)
	rs_Group.close:set rs_Group = nothing
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<table width="100%" border="0" cellpadding="5" cellspacing="1" class="hback">
  <tr> 
    <td width="15%" height="28"> 
    <div align="right"></div></td>
    <td width="10%"><div align="center"><img src="<%=s_savepath%>/images/changeskin.gif" width="86" height="12" border="0" usemap="#Map"></div></td>
    <td width="75%"><div align="right">用户编号:<%=Fs_User.UserNumber%>&nbsp;&nbsp;&nbsp;级别：<% = top_GroupID %>　　积分:<%=Fs_User.NumIntegral%>　金币:<%=Fs_User.NumFS_Money%>&nbsp;<%=top_moneyName%> 　类型: 
        <%if FS_User.isCorp=1 then %>企业<%Else%>个人<%End if%>
    　<a href="<%=s_savepath%>/main.asp"><strong>面板</strong></a>　<a href="<%=s_savepath%>/Loginout.asp"><strong>退出</strong></a>　<a href="<%=s_savepath%>/Help.asp?Dir=<% = Request.ServerVariables("URL")%>" style="cursor:help;"><strong>帮助</strong></a></div></td>
  </tr>
</table>
<map name="Map">
  <area shape="rect" coords="0,1,14,17" href="<%=s_savepath%>/changeskin.asp?Style_num=1&ReturnUrl=<%=UserUrl%>" alt="默认风格">
  <area shape="rect" coords="19,1,31,17" href="<%=s_savepath%>/changeskin.asp?Style_num=2&ReturnUrl=<%=UserUrl%>" alt="银色风格">
  <area shape="rect" coords="36,0,49,16" href="<%=s_savepath%>/changeskin.asp?Style_num=3&ReturnUrl=<%=UserUrl%>" alt="蓝色海洋">
  <area shape="rect" coords="54,1,68,19" href="<%=s_savepath%>/changeskin.asp?Style_num=4&ReturnUrl=<%=UserUrl%>" alt="浪漫咖啡">
  <area shape="rect" coords="73,1,84,15" href="<%=s_savepath%>/changeskin.asp?Style_num=5&ReturnUrl=<%=UserUrl%>" alt="青青河草">
</map>







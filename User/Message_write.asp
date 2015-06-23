<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
if NoSqlHack(request.QueryString("ToUserNumber")) = Fs_User.UserNumber then
		strShowErr = "<li>不能自己给自己发信</li>"
		Call ReturnError(strShowErr,"")
End if
Dim P_ToUserNumber,P_strToUserNumber,GetUsermessObj,FS_Message
Set FS_Message = new Cls_Message
P_ToUserNumber = NoSqlHack(Request.QueryString("ToUserNumber"))
'获得用户名和编号
Set GetUsermessObj = server.CreateObject(G_FS_RS)
GetUsermessObj.open "select  UserID,isLock,UserName,GroupID,UserNumber From FS_ME_Users where UserNumber = '"& P_ToUserNumber &"'",User_Conn,1,3
if GetUsermessObj.eof then
	P_strToUserNumber = ""
Else
	P_strToUserNumber = GetUsermessObj("UserName")
End if
If Request.Form("Action") = "Save" then
	Dim p_M_ReadUserName,p_M_title,p_M_Content
	p_M_ReadUserName = NoSqlHack(Request.Form("M_ReadUserNumber"))
	p_M_title = NoSqlHack(Request.Form("Title"))
	p_M_Content = NoHtmlHackInput(Request.Form("Content"))
	If p_M_ReadUserName="" Or p_M_title="" Or p_M_Content="" Then
		strShowErr = "<li>请填写完整</li><li>短信标题、收件人、信息内容不能为空</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If
	If len(p_M_Content)>500 Then
		strShowErr = "<li>短信内容不能超过500个字符</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If
	If Trim(p_M_ReadUserName)=Fs_User.UserName Then
		strShowErr = "<li>不能自己给自己发送短信</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	Dim Returvaluestr
	Returvaluestr = Fs_User.GetFriendNumber(p_M_ReadUserName)
	Dim t_RsCheckFriend
	if Returvaluestr ="0" then
			strShowErr = "<li>找不到会员信息，可能您发送的会员已经删除</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	Else
		Set t_RsCheckFriend = User_Conn.Execute("select FriendType from FS_ME_Friends where UserNumber='"&Returvaluestr&"' and F_UserNumber='"&Fs_User.UserNumber&"' and FriendType=2")
		If Not t_RsCheckFriend.EOF Then 
			strShowErr = "<li>对方已将你列入黑名单，不能再给他发送信息</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
	End if
	If Fs_User.UserExist(Returvaluestr)=False then
		strShowErr = "<li>没有此用户或者此用户已经被锁定</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If
	Set t_RsCheckFriend = Nothing 

	Set FS_Message = new Cls_Message
	If FS_Message.LenContent(Returvaluestr)+Len(Request.Form("Content")) > 100*1024 then
		strShowErr = "<li>对方短信空间容量已满！请通知对方删除多余短信</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If
	Dim t_fields,t_title,t_from,t_to,t_content,t_Len,t_Send,t_values,t_return,t_isDraft
	t_fields = "M_Title,M_FromUserNumber,M_ReadUserNumber,M_Content,M_FromDate,M_ReadTF,IsRecyle,isDelR,isDelF,isSend,isDraft,LenContent"
	t_title = p_M_title
	t_from = Fs_User.UserNumber
	t_to = Returvaluestr
	t_content = p_M_Content
	t_Len = Len(t_content)
	if Request.Form("isSend")<>"" then
		t_Send=1
	Else
		t_Send=0
	End if
	if Request.Form("isDraft")<>"" then
		t_isDraft=1
	Else
		t_isDraft=0
	End if
	t_values = "'"&t_title&"','"&t_from&"','"&t_to&"','"&t_content&"','"&Now()&"',0,0,0,0,"&t_Send&","&t_isDraft&","&t_Len
	t_return = FS_Message.update(t_fields,t_values,"_new_")
	Set FS_Message = Nothing 
	If t_return Then 
		strShowErr = "<li>恭喜！</li><li>发送成功</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
		strShowErr = "<li>遗憾！</li><li>发送失败</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If 	
Else
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-短消息</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt; 短消息</td>
        </tr>
      </table> 
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="">
          <tr class="hback"> 
            <td height="28" colspan="2" class="hback">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr class="hback"> 
                  <td width="13%" height="19" class="hback"> <div align="right"><a href="Message_box.asp?type=rebox"><img src="images/Skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>//recievebox.gif" width="40" height="40" border="0"></a><a href="Message_box.asp?type=sendbox"><img src="images/Skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>//sendbox.gif" width="40" height="40" border="0"></a><a href="Message_box.asp?type=drabox"><img src="images/Skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>//draftbox.gif" width="40" height="40" border="0"></a><a href="Message_write.asp"><img src="images/Skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>//writemessage.gif" width="40" height="40" border="0"></a></div></td>
                  <td width="14%" class="hback"><div align="right">空间占用 </div></td>
                  <td width="8%" class="hback"><div align="left">: 
                      <%
				     Dim UnTotle,FS_Message_1
					 Set FS_Message_1 = new Cls_Message
					UnTotle=FS_Message_1.LenContent(Fs_User.UserNumber)/(1024*100)*100
					Set FS_Message_1 = Nothing 
					If IsNull(UnTotle) then UnTotle=0
					Response.Write Formatnumber(UnTotle,2,-1)&"%"
					%>
                    </div></td>
                  <td width="65%" class="hback"><table width="100%" height="17" border="0" cellpadding="0" cellspacing="1" class="table">
                      <tr> 
                        <td class="hback_1"><img src="images/space_pic_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.gif" width="<% = Formatnumber((UnTotle),2,-1)%>%" height="17"></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
          <tr class="hback"> 
            <td width="16%" class="hback_1"><div align="center"><strong>用 户 名</strong></div></td>
            <td class="hback"> <div align="left"> 
                <input name="M_ReadUserNumber" type="text" id="M_ReadUserNumber" value="<% = P_strToUserNumber %>" size="20">
                <font color="#999999"> 
                <select name="SelectFriend" id="SelectFriend" onChange="DoTitle(this.options[this.selectedIndex].value)">
                  <option selected value="">>>选择好友<<</option>
                  <%=Fs_User.FriendList%> 
                </select>
                </font>请添写用户名<strong>｜<a href="Friend_add.asp">添加好友</a></strong></div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>消息标题</strong></div></td>
            <td class="hback"> <div align="left"> 
                <input name="Title" type="text" id="Title" size="40">
              </div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>消息内容</strong></div></td>
            <td class="hback"> <div align="left"> 
                <textarea name="Content" cols="50" rows="8" id="Content"></textarea>
              </div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>保存到草稿箱</strong></div></td>
            <td class="hback"><input name="isDraft" type="checkbox" id="isDraft" value="1">
              保存到草稿箱中 </td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>保存到发件箱</strong></div></td>
            <td class="hback"> <div align="left"> 
                <input name="isSend" type="checkbox" id="isSend" value="1" checked>
                保存到发件箱中 </div></td>
          </tr>
          <tr class="hback"> 
            <td colspan="2" class="hback"> <div align="left">　　　　　　　　　 
                <input name="Action" type="hidden" id="Action" value="Save">
                <input type="submit" name="Submit" value=" 发送消息 ">
                　 
                <input type="reset" name="Submit3" value="重新填写">
              </div></td>
          </tr>
          <tr class="hback"> 
            <td colspan="2" class="hback"> <div align="center"> </div></td>
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
<script language="JavaScript" type="text/JavaScript">
function DoTitle(addTitle) {  
document.form1.M_ReadUserNumber.value=document.form1.SelectFriend.value;  
document.form1.M_ReadUserNumber.focus(); 
 return; 
} 
</script>
<%
End if
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
dim str_m_type
str_m_type = CintStr(Request.QueryString("M_type"))
if isnull(str_m_type) or not isnumeric(str_m_type) or trim(str_m_type)="" then
	strShowErr = "<li>错误参数</li>"
	Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
end if
if NoSqlHack(request.QueryString("ToUserNumber")) = Fs_User.UserNumber then
		strShowErr = "<li>不能自己给自己留言</li>"
		Call ReturnError(strShowErr,"")
End if
Dim P_ToUserNumber,P_strToUserNumber,GetUsermessObj,FS_Book
Set FS_Book = new Cls_message
P_ToUserNumber = NoSqlHack(Request.QueryString("ToUserNumber"))
'获得用户名和编号
Set GetUsermessObj = server.CreateObject(G_FS_RS)
GetUsermessObj.open "select  UserID,isLock,UserName,GroupID,UserNumber From FS_ME_Users where UserNumber = '"& NoSqlHack(P_ToUserNumber) &"'",User_Conn,1,3
if GetUsermessObj.eof then
	P_strToUserNumber = ""
Else
	P_strToUserNumber = GetUsermessObj("UserName")
End if
If Request.Form("Action") = "Save" then
	Dim p_M_ReadUserName,p_M_title,p_M_Content,p_M_Type
	p_M_ReadUserName = NoSqlHack(Request.Form("M_ReadUserNumber"))
	p_M_title = NoSqlHack(Request.Form("Title"))
	p_M_Content = NoHtmlHackInput(Request.Form("Content"))
	p_M_Type= CintStr(Request.Form("M_Type"))
	If p_M_ReadUserName="" Or p_M_title="" Or p_M_Content="" Then
		strShowErr = "<li>请填写完整</li><li>留言标题、收件人、信息内容不能为空</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If
	If len(p_M_Content)>500 Then
		strShowErr = "<li>留言内容不能超过500个字符</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If
	If Trim(p_M_ReadUserName)=Fs_User.UserName Then
		strShowErr = "<li>不能自己给自己留言</li>"
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
		Set t_RsCheckFriend = User_Conn.Execute("select FriendType from FS_ME_Friends where UserNumber='"&NoSqlHack(Returvaluestr)&"' and F_UserNumber='"&Fs_User.UserNumber&"' and FriendType=2")
		If Not t_RsCheckFriend.EOF Then 
			strShowErr = "<li>对方已将你列入黑名单，不能再给他留言</li>"
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
	Set FS_Book = new Cls_message
	If FS_Book.LenbContent(Returvaluestr)+Len(Request.Form("Content")) > 100*1024 then
		strShowErr = "<li>对方短信空间容量已满！请通知对方删除多余短信</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If
	dim save_rs
	set save_rs= Server.CreateObject(G_FS_RS)
	save_rs.open "select * From FS_ME_Book where 1=0",User_Conn,1,3
	save_rs.addnew
	save_rs("M_Title")=NoSqlHack(p_M_title)
	save_rs("M_ReadUserNumber")=NoSqlHack(Returvaluestr)
	save_rs("M_Content")=NoSqlHack(p_M_Content)
	save_rs("M_FromUserNumber")=Fs_User.UserNumber
	save_rs("M_FromDate")=now
	save_rs("M_ReadTF")=0
	save_rs("LenContent")=NoSqlHack(len(p_M_Content))
	save_rs("M_Type")=p_M_Type
	save_rs.update
	save_rs.close:set save_rs = nothing
	Set FS_Book = Nothing 
	strShowErr = "<li>恭喜！</li><li>发送成功</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Friend.asp")
	Response.end
Else
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-留言</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt;<a href="Book.asp?M_Type=<%=str_m_type%>">留言管理</a> &gt;&gt; 撰写留言</td>
        </tr>
      </table> 
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="">
          <tr class="hback"> 
            <td height="28" colspan="2" class="hback">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr class="hback"> 
                  <td class="hback"><div align="right">空间占用 </div></td>
                  <td class="hback"><div align="left">: 
                      <%
				     Dim UnTotle,FS_Book_1
					 Set FS_Book_1 = new Cls_message
					UnTotle=FS_Book_1.LenbContent(Fs_User.UserNumber)/(1024*100)*100
					Set FS_Book_1 = Nothing 
					If IsNull(UnTotle) then UnTotle=0
					Response.Write Formatnumber(UnTotle,2,-1)&"%"
					%>
                    </div></td>
                  <td width="80%" class="hback"><table width="100%" height="17" border="0" cellpadding="0" cellspacing="1" class="table">
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
            <td class="hback_1"><div align="center"><strong>留言标题</strong></div></td>
            <td class="hback"> <div align="left"> 
                <input name="Title" type="text" id="Title" size="40">
              </div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>留言内容</strong></div></td>
            <td class="hback"> <div align="left"> 
                <textarea name="Content" cols="50" rows="8" id="Content"></textarea>
              </div></td>
          </tr>
          
          <tr class="hback"> 
            <td colspan="2" class="hback"> <div align="left">　　　　　　　　　 
                <input name="Action" type="hidden" id="Action" value="Save">
                <input name="M_Type" type="hidden" id="M_Type" value="<%=str_m_type%>">
                <input type="submit" name="Submit" value=" 确定留言 ">
                　 
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






<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,tmp_type,strShowErr
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("SS_site") then Err_Show
if not MF_Check_Pop_TF("SS001") then Err_Show

Dim RsOCObj,TempFlag
Set RsOCObj = Conn.Execute("Select top 1 * from FS_SS_SysPara")
If RsOCObj.eof then
	TempFlag = false
Else
	TempFlag = true
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>网站维护</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<body topmargin="2" leftmargin="2">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form action="" method="post" name="VOForm">
    <tr  class="xingmu"> 
      <td colspan="2" class="xingmu"><strong>参数设置</strong></td>
    </tr>
    <tr class="hback"> 
      <td width="24%">&nbsp;&nbsp;&nbsp;&nbsp; <div align="right">网站名称</div></td>
      <td width="76%"> <input name="WebName" type="text" id="WebName" style="width:90%" value="<%If TempFlag = true then Response.Write(RsOCObj("WebName")) end if%>"></td>
    </tr>
    <tr class="hback"> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp; <div align="right">网站地址</div></td>
      <td> <input name="WebUrl" type="text" id="WebUrl" style="width:90%" value="<%If TempFlag = true then Response.Write(RsOCObj("WebUrl")) end if%>"></td>
    </tr>
    <tr class="hback"> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp; <div align="right">管理员</div></td>
      <td> <input name="WebAdmin" type="text" id="WebAdmin" style="width:90%" value="<%If TempFlag = true then Response.Write(RsOCObj("WebAdmin")) end if%>"></td>
    </tr>
    <tr class="hback"> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp; <div align="right">网站信箱</div></td>
      <td> <input name="WebEmail" type="text" id="WebEmail" style="width:90%" value="<%If TempFlag = true then Response.Write(RsOCObj("WebEmail")) end if%>"></td>
    </tr>
    <tr class="hback"> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp; <div align="right">开始统计时间</div></td>
      <td> <input name="WebCountTime" type="text" readonly id="WebCountTime" style="width:71%" value="<%If TempFlag = true then Response.Write(RsOCObj("WebCountTime")) end if%>"> 
        <input type="button" name="dfgdf" value="选择日期" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',280,150,window,document.VOForm.WebCountTime);"></td>
    </tr>
    <tr class="hback">
      <td><div align="right">统计防刷新时间</div></td>
      <td><input name="ExpTime" type="text" id="ExpTime" style="width:90%" value="<%If TempFlag = true then Response.Write(RsOCObj("ExpTime")) end if%>" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
        分钟 </td>
    </tr>
    <tr class="hback"> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp; <div align="right">网站介绍</div></td>
      <td> <textarea name="WebIntro" rows="5" id="WebIntro" style="width:90%"><%If TempFlag = true then Response.Write(RsOCObj("WebIntro")) end if%></textarea></td>
    </tr>
    <tr class="hback"> 
      <td colspan="2"> <div align="center"> 
          <input type="submit" name="Submit" value=" 确 定 ">
          &nbsp;&nbsp; 
          <input name="action" type="hidden" id="action" value="trues">
          <input type="reset" name="Submit" value=" 还 原 ">
          &nbsp;&nbsp; 
          <input type="button" name="Submit" value=" 取 消 " onClick="history.back();">
        </div></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
	If Request.Form("action") = "trues" then
		Dim VOModObj,VoModSql
		Set VOModObj = Server.CreateObject(G_FS_RS)
		VoModSql = "Select top 1 * from FS_SS_SysPara"
		VOModObj.Open VoModSql,Conn,1,3
		If TempFlag = false then
			VOModObj.AddNew
		End If
		VOModObj("WebName") = NoSqlHack(Replace(Request.Form("WebName"),"""",""))
		VOModObj("WebUrl") = NoSqlHack(Request.Form("WebUrl"))
		VOModObj("WebIntro") = NoSqlHack(Request.Form("WebIntro"))
		VOModObj("WebEmail") = NoSqlHack(Request.Form("WebEmail"))
		VOModObj("WebAdmin") = NoSqlHack(Request.Form("WebAdmin"))
		VOModObj("WebCountTime") = NoSqlHack(Request.Form("WebCountTime"))
		VOModObj("ExpTime") = NoSqlHack(Request.Form("ExpTime"))
		VOModObj.Update
		VOModObj.Close
		Set VOModObj = Nothing
		strShowErr = "<li>修改成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If
Conn.Close
Set Conn = Nothing
%>






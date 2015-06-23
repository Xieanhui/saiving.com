<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp"-->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_InterFace/ns_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
 
MF_Default_Conn
Dim Conn,User_Conn
Dim Configobj,PageS,sql,MSTitle,ShowIP,s_IsUser,Style
Set Configobj=server.CreateObject (G_FS_RS)
sql="select ID,Title,IPShow,IsUser,IsAut,PageSize,Style From FS_WS_Config"
configobj.open sql,Conn,1,1
if not configobj.eof then
	PageS=configobj("PageSize")
	MSTitle=configobj("Title")
	ShowIP=configobj("IPShow")
	s_IsUser = configobj("IsUser")
	Style = configobj("Style")
	if Style<>"" then
		Style = Style
	else
		Style = "3"
	end if
end if
Response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = Style
set configobj=nothing
if s_IsUser <> "0" then
	if session("FS_UserName")="" then
		response.Write"未开放匿名发布帖！"
		response.end
	end if
end if
dim ClassID,ClassRs,ClassName
ClassID=NoSqlHack(Request.querystring("ClassID"))
ClassName=NoSqlHack(Request.querystring("ClassName"))
%>
<html>
<HEAD>
<TITLE><%=GetGuestBookTitle%></TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../<% = G_USER_DIR %>/images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="javascript">
<!--
function check(){
if (document.WriteNote.Topic.value==""){
alert("留言标题不能为空!");
document.WriteNote.Topic.focus();
return false;
}
if (document.WriteNote.Content.value==""){
alert("留言内容不能为空");
document.WriteNote.Content.focus();
return false;
}
return true;
}
-->
</script>
<body>
<form id="WriteNote" name="WriteNote" action="SaveNotes.asp?Act=Add" method="post">
  <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
    <tr class="hback"
>
      <td height="30" colspan="4" align="left" class="xingmu"><img src="images/Forum_nav.gif"><a href="index.asp" class="Top_Navi"><b><%=MSTitle%></b></a> -> <a href="DefNoteList.asp?ClassID=<%=ClassID%>" class="Top_Navi"><b><%=ClassName%></b></a></td>
    </tr>
    <tr>
      <td class="hback"
 align="right" width="18%">贴子标题:</td>
      <td width="82%" class="hback"
 >
        <input type="text" name="Topic" size="50" maxlength="80">
        &nbsp;&nbsp;
        <input type="checkbox" id="IsAdmin" name="IsAdmin" value="1">
      管理员可见</td>
    </tr>
    <tr class="hback"
>
      <td class="hback"
 align="right" width="18%">当前表情:</td>
      <td class="hback"
 align="left">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>
              <input name="FaceNum" type="radio" value="1" checked="checked">
              <img src="Images/face1.gif" width="22" height="22"> </td>
            <td>
              <input type="radio" name="FaceNum" value="2" >
              <img src="Images/face2.gif" width="22" height="22"> </td>
            <td>
              <input type="radio" name="FaceNum" value="3" >
              <img src="Images/face3.gif" width="22" height="22"> </td>
            <td>
              <input type="radio" name="FaceNum" value="4" >
              <img src="Images/face4.gif" width="22" height="22"> </td>
            <td>
              <input type="radio" name="FaceNum" value="5" >
              <img src="Images/face5.gif" width="22" height="22"> </td>
            <td>
              <input type="radio" name="FaceNum" value="6">
              <img src="Images/face6.gif" width="22" height="22"></td>
            <td>
              <input type="radio" name="FaceNum" value="7" >
              <img src="Images/face7.gif" width="22" height="22"> </td>
            <td>
              <input type="radio" name="FaceNum" value="8" >
              <img src="Images/face8.gif" width="22" height="22"> </td>
            <td>
              <input type="radio" name="FaceNum" value="9" >
              <img src="Images/face9.gif" width="22" height="22"></td>
          </tr>
          <tr>
            <td>
              <input type="radio" name="FaceNum" value="10" >
              <img src="Images/face10.gif" width="22" height="22"></td>
            <td>
              <input type="radio" name="FaceNum" value="11" >
              <img src="Images/face11.gif" width="22" height="22"> </td>
            <td>
              <input type="radio" name="FaceNum" value="12">
              <img src="Images/face12.gif" width="22" height="22"></td>
            <td>
              <input type="radio" name="FaceNum" value="13" >
              <img src="Images/face13.gif" width="22" height="22"> </td>
            <td>
              <input type="radio" name="FaceNum" value="14">
              <img src="Images/face14.gif" width="22" height="22"></td>
            <td>
              <input type="radio" name="FaceNum" value="15" >
              <img src="Images/face15.gif" width="22" height="22"></td>
            <td>
              <input type="radio" name="FaceNum" value="16" >
              <img src="Images/face16.gif" width="22" height="22"></td>
            <td>
              <input type="radio" name="FaceNum" value="17" >
              <img src="Images/face17.gif" width="22" height="22"> </td>
            <td>
              <input type="radio" name="FaceNum" value="18" >
              <img src="Images/face18.gif" width="22" height="22"> </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr class="hback"
>
      <td class="hback"
 align="right" width="18%">贴子类型:</td>
      <td class="hback"
 align="left">
        <input type="radio" name="Style" value="1">
        推荐帖子
        <input type="radio" name="Style" value="0" checked>
        普通帖子</td>
    </tr>
    <tr class="hback"
>
      <td class="hback"
 align="right">帖子内容:</td>
      <td class="hback"
 valign="top" >
        <textarea name="Content" cols="78" rows="10" id="Content"></textarea>
      </td>
    </tr>
    <tr class="hback"
>
      <td class="hback"
>&nbsp;
          <input type="hidden" name="ClassID" id="ClassID" value="<%=ClassID%>">
      </td>
      <td class="hback"
>
        <input type="submit" name="submit" value="OK发表贴子" onClick="return check()">
        &nbsp;&nbsp;
        <input type="reset" name="reset" value=" 清  空 ">
      </td>
    </tr>
  </table>
</form>
<%
dim NoteRs,Topic,Content,Face,IsUser,isTop,IsAdmin
if NoSqlHack(Request.QueryString("Act"))="Add" then
Set NoteRs=Server.CreateObject(G_FS_RS)
	ClassID=NoSqlHack(Request.form("ClassID"))
	Topic=NoHtmlHackInput(NoSqlHack(Replace(Request.form("Topic"),"'","")))
	IsTop=NoHtmlHackInput(NoSqlHack(Request.form("Style")))
	Content=NoHtmlHackInput(NoSqlHack(Request.form("Content")))
	IsAdmin=NoHtmlHackInput(NoSqlHack(Request.form("IsAdmin")))
	Face=NoHtmlHackInput(NoSqlHack(Request.Form("FaceNum")))
	if IsAdmin="" then
		IsAdmin="0"
	end if
	if ClassID="" then
		Response.write("<script>alert('参数出错!');</script>")
		response.end
	end if
	if Topic="" then
		Response.write("<script>alert('标题不能为空');</script>")
		Response.end
	end if
	if isTop="" then
		Response.write("<script>alert('帖子类型值没传过来');</script>")
		Response.end
	end if
	if Content="" then
		Response.write("<script>alert('内容不能为空');</script>")
		response.end
	end if
	NoteRs.open "Select * from FS_WS_BBS ",Conn,3,3
	NoteRs.Addnew
	NoteRs("ClassID")=ClassID
	NoteRs("User")=session("FS_UserName")
	NoteRs("Topic")=Topic
	NoteRs("Body")=Content
	NoteRs("AddDate")=now()
	NoteRs("IsTop")=IsTop
	NoteRs("Style")="普通"
	NoteRs("IsAdmin")=IsAdmin
	NoteRs("LastUpdateUser")=session("FS_UserName")
	NoteRs("Face")=Face
	NoteRs("IP")=NoSqlHack(Request.ServerVariables("REMOTE_ADDR"))
	NoteRs.update
	Set NoteRs=nothing
	Response.write "<script>location.href='DefNoteList.asp?ClassID='"&ClassID&"';</script>"
end if
Response.end
Set Conn=nothing
%>
</body>
</html>







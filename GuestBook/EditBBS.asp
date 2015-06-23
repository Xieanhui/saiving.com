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
Dim Configobj,PageS,sql,MSTitle,ShowIP,Style
Set Configobj=server.CreateObject (G_FS_RS)
sql="select ID,Title,IPShow,IsUser,IsAut,PageSize,Style From FS_WS_Config"
configobj.open sql,Conn,1,1
if not configobj.eof then
	PageS=configobj("PageSize")
	MSTitle=configobj("Title")
	ShowIP=configobj("IPShow")
	Style = configobj("Style")
	if Style<>"" then
		Style = Style
	else
		Style = "3"
	end if
end if
response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = Style
set configobj=nothing
%>
<html>
<HEAD>
<TITLE><%=GetGuestBookTitle%></TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../<% = G_USER_DIR %>/images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</HEAD>
<script language="javascript">
function goback(){
history.back();
}
</script>
<body>
<%
Dim BBsID,NoteID,ClassName,ClassID,Pag,NoteTilte,ShowRs,strShowErr
if NoSqlHack(Request.querystring("Act"))="Edit" then
NoteTilte=NoSqlHack(Request.QueryString("NoteTilte"))
BBsID=NoSqlHack(Request.QueryString("BBSID"))
NoteID=NoSqlHack(Request.QueryString("NoteID"))
ClassName=NoSqlHack(Request.QueryString("ClassName"))
ClassID=NoSqlHack(Request.QueryString("ClassID"))
Pag=NoSqlHack(Request.QueryString("Page"))
Set ShowRs=Conn.execute("Select ID,ClassID,[User],ParentID,Topic,Body,AddDate,IsTop,State,Style,IsAdmin,Answer,Hit,LastUpdateDate,LastUpdateUser,Face,IP From FS_WS_BBS Where ID="&CintStr(BBSID)&"")
%>
<form id="WriteNote" name="WriteNote" action="?Act=save" method="post">
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="hback">
    <td align="left" class="xingmu" colspan="4"><img src="images/Forum_nav.gif"><a href="index.asp" class="Top_Navi"><b><%=MSTitle%></b></a> -> <a href="DefNoteList.asp?ClassID=<%=ClassID%>" class="Top_Navi"><b><%=ClassName%></b></a> -> <a href="#" onClick="goback()" class="Top_Navi"><b><%=NoteTilte%></b></a></td>
  </tr>
  <tr>
    <td class="hback" align="right" width="20%">贴子标题:</td>
    <td class="hback" >
      <input type="text" name="Topic" size="50" maxlength="60" value="<%=NoteTilte%>" readonly>
      &nbsp;&nbsp;</td>
  </tr>
  <tr class="hback">
    <td class="hback" align="right" width="20%">当前表情:</td>
    <td class="hback" align="left">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>
            <input name="FaceNum" type="radio" value="1" <%If ShowRs("Face")=1 then response.Write("Checked")%>>
            <img src="Images/face1.gif" width="22" height="22"> </td>
          <td>
            <input type="radio" name="FaceNum" value="2" <%If ShowRs("Face")=2 then response.Write("Checked")%>>
            <img src="Images/face2.gif" width="22" height="22"> </td>
          <td>
            <input type="radio" name="FaceNum" value="3" <%If ShowRs("Face")=2 then response.Write("Checked")%>>
            <img src="Images/face3.gif" width="22" height="22"> </td>
          <td>
            <input type="radio" name="FaceNum" value="4" <%If ShowRs("Face")=4 then response.Write("Checked")%>>
            <img src="Images/face4.gif" width="22" height="22"> </td>
          <td>
            <input type="radio" name="FaceNum" value="5" <%If ShowRs("Face")=5 then response.Write("Checked")%>>
            <img src="Images/face5.gif" width="22" height="22"> </td>
          <td>
            <input type="radio" name="FaceNum" value="6" <%If ShowRs("Face")=6 then response.Write("Checked")%>>
            <img src="Images/face6.gif" width="22" height="22"></td>
          <td>
            <input type="radio" name="FaceNum" value="7" <%If ShowRs("Face")=7 then response.Write("Checked")%>>
            <img src="Images/face7.gif" width="22" height="22"> </td>
          <td>
            <input type="radio" name="FaceNum" value="8" <%If ShowRs("Face")=8 then response.Write("Checked")%>>
            <img src="Images/face8.gif" width="22" height="22"> </td>
          <td>
            <input type="radio" name="FaceNum" value="9" <%If ShowRs("Face")=9 then response.Write("Checked")%>>
            <img src="Images/face9.gif" width="22" height="22"></td>
        </tr>
        <tr>
          <td>
            <input type="radio" name="FaceNum" value="10" <%If ShowRs("Face")=10 then response.Write("Checked")%>>
            <img src="Images/face10.gif" width="22" height="22"></td>
          <td>
            <input type="radio" name="FaceNum" value="11" <%If ShowRs("Face")=11 then response.Write("Checked")%>>
            <img src="Images/face11.gif" width="22" height="22"> </td>
          <td>
            <input type="radio" name="FaceNum" value="12" <%If ShowRs("Face")=12 then response.Write("Checked")%>>
            <img src="Images/face12.gif" width="22" height="22"></td>
          <td>
            <input type="radio" name="FaceNum" value="13" <%If ShowRs("Face")=13 then response.Write("Checked")%>>
            <img src="Images/face13.gif" width="22" height="22"> </td>
          <td>
            <input type="radio" name="FaceNum" value="14" <%If ShowRs("Face")=14 then response.Write("Checked")%>>
            <img src="Images/face14.gif" width="22" height="22"></td>
          <td>
            <input type="radio" name="FaceNum" value="15" <%If ShowRs("Face")=15 then response.Write("Checked")%>>
            <img src="Images/face15.gif" width="22" height="22"></td>
          <td>
            <input type="radio" name="FaceNum" value="16" <%If ShowRs("Face")=16 then response.Write("Checked")%>>
            <img src="Images/face16.gif" width="22" height="22"></td>
          <td>
            <input type="radio" name="FaceNum" value="17" <%If ShowRs("Face")=17 then response.Write("Checked")%>>
            <img src="Images/face17.gif" width="22" height="22"> </td>
          <td>
            <input type="radio" name="FaceNum" value="18" <%If ShowRs("Face")=18 then response.Write("Checked")%>>
            <img src="Images/face18.gif" width="22" height="22"> </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr class="hback">
    <td class="hback" align="right">帖子内容:</td>
    <td class="hback" valign="top" >
      <textarea name="Content" cols="78" rows="8" id="Content" ><%=replace(ShowRs("Body"),"<br>",chr(13)&chr(10))%></textarea>
    </td>
  </tr>
  <tr class="hback">
    <td class="hback">&nbsp;
        <input type="hidden" name="ClassID" id="ClassID" value="<%=ClassID%>">
		<input type="hidden" name="BbsId" value="<%=ShowRs("ID")%>">
		<input type="hidden" name="ClassName" value="<%=ClassName%>">
		<input type="hidden" name="noteId" value="<%=NoteID%>">
		<input type="hidden" name="Page"  value="<%=Pag%>">
    </td>
    <td class="hback">
      <input type="submit" name="submit" value="OK保存贴子">
      &nbsp;&nbsp;
      <input type="reset" name="reset" value=" 重  置 ">
    </td>
  </tr>
</table>
</form>
<%
Set ShowRs=nothing
end if
dim NoteRs,Topic,Content,Face,IsUser,isTop,IsAdmin,url
if NoSqlHack(Request.QueryString("Act"))="save" then
Set NoteRs=Server.CreateObject(G_FS_RS)
	ClassName = NoSqlHack(request.form("ClassName"))
	ClassID = NoSqlHack(request.form("ClassID")
	BbsId = NoSqlHack(request.form("BbsId"))
	NoteID = NoSqlHack(request.forM("NoteID"))
	Topic=NoHtmlHackInput(NoSqlHack(Replace(Request.form("Topic"),"'","")))
	IsTop=NoHtmlHackInput(NoSqlHack(Request.form("Style")))
	Content=replace(NoHtmlHackInput(NoSqlHack(Request.form("Content"))),chr(13)&chr(10),"<br>")
	IsAdmin=NoHtmlHackInput(NoSqlHack(Request.form("IsAdmin")))
	Face=NoHtmlHackInput(NoSqlHack(Request.Form("FaceNum")))
	Pag = NoSqlHack(request.form("Page"))
	'	response.Write(ClassName)
	'Response.End()

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
	if Content="" then
		Response.write("<script>alert('内容不能为空');</script>")
		response.end
	end if
	NoteRs.open "Select ID,ClassID,[User],ParentID,Topic,Body,AddDate,IsTop,State,Style,IsAdmin,Answer,Hit,LastUpdateDate,LastUpdateUser,Face,IP From FS_WS_BBS Where ID="&CintStr(BbsId)&"",Conn,3,3
	if not Noters.eof then
	NoteRs("Body")=Content
	NoteRs("LastUpdateUser")=session("FS_UserName")	
	NoteRs("LastUpdateDate")=now()
	NoteRs("Face")=Face
	NoteRs("IP") = NoSqlHack(request.ServerVariables("REMOTE_ADDR"))
	NoteRs.update
	end if
	Set NoteRs=nothing
	url="ShowNote.asp?ClassName="&NoSqlHack(ClassName)&"&NoteID="&NoSqlHack(NoteID)&"&ClassID="&NoSqlHack(ClassID)&"&Page="&NoSqlHack(Pag)&""
	'response.End()
	Response.write("<script>location.href='"&url&"'</script>")
	Response.end
end if
Response.end
Set Conn=nothing
%>
</body>
</html>







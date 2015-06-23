<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,str_CurrPath,sRootDir
MF_Default_Conn
'session判断
MF_Session_TF 
'权限判断
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
if not MF_Check_Pop_TF("WS002") then Err_Show
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if
dim Fs_news
set Fs_news = new Cls_News
%>
<html>
<HEAD>
<TITLE>FoosunCMS留言系统</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="javascript">
function ShowNote(NoteID,ClassName,ClassID)
{
location="ShowNote.asp?NoteID="+NoteID+"&ClassName="+ClassName+"&ClassID="+ClassID;
}
</script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/PublicJS.js"></script>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Get_Domain.asp"></script>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Prototype.js"></script>
<script language="JavaScript" type="text/javascript" src="../../Editor/FS_scripts/editor.js"></script>
<script language="javascript">
function goback(){
history.go(-1);
}
</script>
<body>
<%
Dim ID,ShowRs,SaveRs,NoteRs,Topic,Content,Face,IsUser,isTop,IsAdmin,ClassName,NoteID,ClassID,strShowErr
Set ShowRs=Server.CreateObject(G_FS_RS)
Set SaveRs=Server.CreateObject(G_FS_RS)
if NoSqlHack(Request.QueryString("Act"))="Edit" then
	ID=NoSqlHack(Request.QueryString("BBSID"))
	if ID="" then
		strShowErr = "<li>参数出错!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	ShowRs.open "Select ID,ClassID,[User],Topic,Body,AddDate,IsTop,State,IsAdmin,LastUpdateDate,LastUpdateUser,Face From FS_WS_BBS Where ID="&CintStr(ID)&"",Conn,1,1
	if ShowRs.eof and ShowRs.bof then
		strShowErr = "<li>你要修改的帖子已不存在了!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	else
%>

<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="hback">
    <td align="left" colspan="2" class="xingmu">留言板&nbsp;<a href="../../help?Lable=Message" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr>
    <td colspan="2" class="hback"><a href="#" onClick="goback()">返回</a><a href="ClassMessageManager.asp"></a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
<form id="WriteNote" name="WriteNote" action="?Act=Save" method="post">
	<tr class="hback">
		<td class="hback" align="right" width="20%">贴子标题</td>
	  <td class="hback" ><input type="text" name="Topic" size="50" maxlength="80" value="<%=ShowRs("Topic")%>">
&nbsp;&nbsp;</td>
	</tr>
	<tr class="hback">
		<td class="hback" align="right" width="20%">当前表情</td>
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
		<td class="hback" align="right">帖子内容</td>
        <td class="hback" valign="top" >
				
                <!--编辑器开始-->
				<iframe id='NewsContent' src='../Editer/AdminEditer.asp?id=Content' frameborder=0 scrolling=no width='100%' height='280'></iframe>
				<input type="hidden" name="Content" value="<% = HandleEditorContent(ShowRs("Body")) %>">
                <!--编辑器结束-->	
		</td>
	</tr>
	<tr class="hback">
		<td class="hback"><input type="hidden" name="ClassName" value="<%=Request.QueryString("ClassName")%>">&nbsp;<input type="hidden" name="NoteID" ID="NoteID" value="<%=ShowRs("ID")%>"></td>
	  <td class="hback"><input type="button" name="submit1" value="OK保存贴子" onClick="checkformdata(this.form);">&nbsp;&nbsp;
	    <input type="reset" name="reset" value=" 清  空 ">
</form>
	</table>
<%
	end if
end if
set ShowRs=nothing
if Trim(Request.QueryString("Act"))="Save" then

	ID=NoSqlHack(Request.Form("NoteID"))
	Topic=NoHtmlHackInput(NoSqlHack(Request.form("Topic")))
	Content=NoHtmlHackInput(NoSqlHack(Request.form("Content")))
	IsAdmin=NoHtmlHackInput(NoSqlHack(Request.form("IsAdmin")))

	ClassName=NoSqlHack(NoHtmlHackInput(Request.Form("ClassName")))
	if IsAdmin="" then
		IsAdmin="0"
	end if
	if ID="" then
		strShowErr = "<li>参数出错!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	if Topic="" then
		strShowErr = "<li>标题不能为空!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if	
	if Content="" then
		strShowErr = "<li>内容不能为空!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	
	SaveRs.open "Select ID,ClassID,[User],Topic,Body,AddDate,IsTop,State,ParentID,IsAdmin,LastUpdateDate,LastUpdateUser,Face From FS_WS_BBS Where ID="&CintStr(ID)&"",Conn,3,3
if not SaveRs.eof then
	Face=NoSqlHack(Request.Form("FaceNum"))
	ClassID=SaveRs("ClassID")
	NoteID=SaveRs("ParentID")
	SaveRs("Topic")=Topic
	SaveRs("Body")=Content
	SaveRs("IsTop")=IsTop
	SaveRs("IsAdmin")=IsAdmin
	SaveRs("LastUpdateUser")=session("Admin_Name")
	SaveRs("LastUpdateDate")=now()
	SaveRs("Face")=Face
	SaveRs.update
	Response.write("'<script>ShowNote("&NoteID&",'"&ClassName&"','"&ClassID&"');</script>'")
else
		strShowErr = "<li>参数出错!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
end if
Set SaveRs=nothing
end if
Set Conn=nothing
%>
<script language="javascript">
function checkformdata(FormObj)
{
	if (frames["NewsContent"].g_currmode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return false;}
	FormObj.Content.value=frames["NewsContent"].GetNewsContentArray();
	FormObj.submit();
}
</script>
</body>
</html>







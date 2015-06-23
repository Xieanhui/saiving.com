<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,str_CurrPath,sRootDir
MF_Default_Conn
'session判断
MF_Session_TF 
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
End if%>
<html>
<HEAD>
<TITLE>FoosunCMS留言系统</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/PublicJS.js"></script>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Get_Domain.asp"></script>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Prototype.js"></script>
<body>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="hback">
    <td align="left" colspan="2" class="xingmu">留言板&nbsp;<a href="../../help?Lable=Message" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr>
    <td colspan="2" class="hback"><a href="ClassMessageManager.asp">管理首页</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
<form  name="WriteNote" action="?Act=Add" method="post">
	<tr class="hback">
		<td class="hback" align="right" width="20%">贴子标题</td>
		<td class="hback" ><input type="text" name="Topic" size="50" maxlength="80">
		&nbsp;&nbsp;<input type="checkbox" id="IsAdmin" name="IsAdmin" value="1">仅管理员可见</td>
	</tr>
	<tr class="hback">
		<td class="hback" align="right" width="20%">贴子类型</td>
		<td class="hback" align="left"><input type="radio" name="Style" value="1">推荐帖子<input type="radio" name="Style" value="0" checked>普通帖子</td>
	</tr>
<tr>
				<td class="hback" align="right" height="25">表情</td>
				<td class="hback" ><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td> <input name="FaceNum" type="radio" value="1" checked> 
                      <img src="Images/face1.gif" width="22" height="22"> </td>
                    <td> <input type="radio" name="FaceNum" value="2"> <img src="Images/face2.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="3"> <img src="Images/face3.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="4"> <img src="Images/face4.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="5"> <img src="Images/face5.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="6"> <img src="Images/face6.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="7"> <img src="Images/face7.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="8"> <img src="Images/face8.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="9"> <img src="Images/face9.gif" width="22" height="22"></td>
                  </tr>
                  <tr> 
                    <td> <input type="radio" name="FaceNum" value="10"> <img src="Images/face10.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="11"> <img src="Images/face11.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="12"> <img src="Images/face12.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="13"> <img src="Images/face13.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="14"> <img src="Images/face14.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="15"> <img src="Images/face15.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="16"> <img src="Images/face16.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="17"> <img src="Images/face17.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="18"> <img src="Images/face18.gif" width="22" height="22">                    </td>
                  </tr>
                </table>
				</td>
				</tr>
				<tr class="hback">
		<td class="hback" align="right">帖子内容</td>
        <td class="hback" valign="top" >
                <!--编辑器开始-->
				<iframe id='NewsContent' src='../Editer/AdminEditer.asp?id=Content' frameborder=0 scrolling=no width='100%' height='280'></iframe>
				<input type="hidden" name="Content">
                <!--编辑器结束--></td>
	</tr>
	<tr class="hback">
		<td class="hback">&nbsp;<input type="hidden" name="ClassID" ID="ClassID" value="<%=Request.querystring("ClassID")%>"></td>
	  <td class="hback"><input type="button" name="submit1" value="OK发表贴子" onClick="SubmitFun();">&nbsp;&nbsp;
	    <input type="reset" name="reset" value=" 清  空 "></td>
		</tr>
</form>
	</table>


</body>
</html>
<%
dim ClassID,NoteRs,Topic,Content,Face,IsUser,isTop,IsAdmin,strShowErr
Set NoteRs=Server.CreateObject(G_FS_RS)
if NoSqlHack(Request.QueryString("Act"))="Add" then
	ClassID=NoHtmlHackInput(Request.form("ClassID"))
	Topic=NoHtmlHackInput(trim(Replace(Request.form("Topic"),"'","")))
	IsTop=NoHtmlHackInput(trim(Request.form("Style")))
	Content=NoHtmlHackInput(Trim(Request.form("Content")))
	IsAdmin=NoHtmlHackInput(Trim(Request.form("IsAdmin")))
	if IsAdmin="" then
		IsAdmin="0"
	end if
	if ClassID="" then
		strShowErr = "<li>参数出错!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	if Topic="" then
		strShowErr = "<li>标题不能为空!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	if isTop="" then
		strShowErr = "<li>帖子类型值没传过来!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	if Content="" then
		strShowErr = "<li>内容不能为空!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	NoteRs.open "Select ClassID,[User],Topic,Body,AddDate,IsTop,State,Style,IsAdmin,Hit,LastUpdateUser,Face,IP from FS_WS_BBS where 1=2",Conn,1,3
	NoteRs.Addnew
	NoteRs("ClassID")=NoSqlHack(ClassID)
	NoteRs("User")=session("Admin_Name")
	NoteRs("Topic")=Topic
	NoteRs("Body")=Content
	NoteRs("AddDate")=now()
	NoteRs("IsTop")=IsTop
	NoteRs("Style")="普通"
	NoteRs("IsAdmin")=IsAdmin
	NoteRs("Face")=NoSqlHack(Request.Form("FaceNum"))
	NoteRs("LastUpdateUser")=session("Admin_Name")
	NoteRs("IP")=NoSqlHack(Request.ServerVariables("REMOTE_ADDR"))
	NoteRs.update
	Set NoteRs=nothing
		strShowErr = "<li>发帖成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=bbs/ClassMessageManager.asp")
		Response.end
end if
Set Conn=nothing
%>
<script language="javascript">
function SubmitFun()
{
	if (frames["NewsContent"].g_currmode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return false;}
	document.WriteNote.Content.value=frames["NewsContent"].GetNewsContentArray();
	document.WriteNote.submit();
}
</script>






<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,sRootDir,str_CurrPath
MF_Default_Conn
'session�ж�
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
End if
%>
<html>
<HEAD>
<TITLE>FoosunCMS����ϵͳ</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/PublicJS.js"></script>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Get_Domain.asp"></script>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Prototype.js"></script>
<script language="javascript">
function goback(){
history.go(-1);
}
</script>
<body>
<%
Dim ID,ShowRs,SaveRs,NoteRs,Topic,Content,Face,IsUser,isTop,IsAdmin,strShowErr
Set ShowRs=Server.CreateObject(G_FS_RS)
Set SaveRs=Server.CreateObject(G_FS_RS)
if Request.QueryString("Act")="NoteEdit" then
	ID=Request.QueryString("ID")
	if ID="" then
		strShowErr = "<li>��������!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	ShowRs.open "Select ID,ClassID,[User],Topic,Body,AddDate,IsTop,State,IsAdmin,Face,LastUpdateDate,LastUpdateUser From FS_WS_BBS Where ID="&CintStr(ID)&"",Conn,1,1
	if ShowRs.eof and ShowRs.bof then
		strShowErr = "<li>��Ҫ�޸ĵ������Ѳ�������!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	else
%>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="hback">
    <td align="left" colspan="2" class="xingmu">���԰�&nbsp;<a href="../../help?Lable=Message" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr>
    <td colspan="2" class="hback"><a href="#" onClick="goback()">����</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <form id="WriteNote" name="WriteNote" action="?Act=Save" method="post">
    <tr class="hback">
      <td class="hback" align="right" width="20%">���ӱ���</td>
      <td class="hback" ><input type="text" name="Topic" size="50" maxlength="80" value="<%=ShowRs("Topic")%>">
        &nbsp;&nbsp;
        <%
		if ShowRs("IsAdmin")="0" then
			Response.write("<input type=""checkbox"" id=""IsAdmin"" name=""IsAdmin"" value=""1"">")
		else
			Response.write("<input type=""checkbox"" id=""IsAdmin"" name=""IsAdmin"" value=""1"" checked>")
		end if
		%>
       ������Ա�ɼ�</td>
    </tr>
    <tr class="hback">
      <td class="hback" align="right" width="20%">��������</td>
      <td class="hback" align="left"><%
		if ShowRs("IsTop")=0 then 
		response.write("<input type=""radio"" name=""Style"" value=""1"">�Ƽ�����<input type=""radio"" name=""Style"" value=""0"" checked>��ͨ����")		
		else
		response.write("<input type=""radio"" name=""Style"" value=""1"" checked>�Ƽ�����<input type=""radio"" name=""Style"" value=""0"">��ͨ����")		
		end if
		%>
      </td>
    </tr>
    <tr class="hback">
      <td class="hback" align="right" width="20%">��ǰ����</td>
      <td class="hback" align="left"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><input name="FaceNum" type="radio" value="1" <%If ShowRs("Face")=1 then response.Write("Checked")%>>
              <img src="Images/face1.gif" width="22" height="22"> </td>
            <td><input type="radio" name="FaceNum" value="2" <%If ShowRs("Face")=2 then response.Write("Checked")%>>
              <img src="Images/face2.gif" width="22" height="22"> </td>
            <td><input type="radio" name="FaceNum" value="3" <%If ShowRs("Face")=2 then response.Write("Checked")%>>
              <img src="Images/face3.gif" width="22" height="22"> </td>
            <td><input type="radio" name="FaceNum" value="4" <%If ShowRs("Face")=4 then response.Write("Checked")%>>
              <img src="Images/face4.gif" width="22" height="22"> </td>
            <td><input type="radio" name="FaceNum" value="5" <%If ShowRs("Face")=5 then response.Write("Checked")%>>
              <img src="Images/face5.gif" width="22" height="22"> </td>
            <td><input type="radio" name="FaceNum" value="6" <%If ShowRs("Face")=6 then response.Write("Checked")%>>
              <img src="Images/face6.gif" width="22" height="22"></td>
            <td><input type="radio" name="FaceNum" value="7" <%If ShowRs("Face")=7 then response.Write("Checked")%>>
              <img src="Images/face7.gif" width="22" height="22"> </td>
            <td><input type="radio" name="FaceNum" value="8" <%If ShowRs("Face")=8 then response.Write("Checked")%>>
              <img src="Images/face8.gif" width="22" height="22"> </td>
            <td><input type="radio" name="FaceNum" value="9" <%If ShowRs("Face")=9 then response.Write("Checked")%>>
              <img src="Images/face9.gif" width="22" height="22"></td>
          </tr>
          <tr>
            <td><input type="radio" name="FaceNum" value="10" <%If ShowRs("Face")=10 then response.Write("Checked")%>>
              <img src="Images/face10.gif" width="22" height="22"></td>
            <td><input type="radio" name="FaceNum" value="11" <%If ShowRs("Face")=11 then response.Write("Checked")%>>
              <img src="Images/face11.gif" width="22" height="22"> </td>
            <td><input type="radio" name="FaceNum" value="12" <%If ShowRs("Face")=12 then response.Write("Checked")%>>
              <img src="Images/face12.gif" width="22" height="22"></td>
            <td><input type="radio" name="FaceNum" value="13" <%If ShowRs("Face")=13 then response.Write("Checked")%>>
              <img src="Images/face13.gif" width="22" height="22"> </td>
            <td><input type="radio" name="FaceNum" value="14" <%If ShowRs("Face")=14 then response.Write("Checked")%>>
              <img src="Images/face14.gif" width="22" height="22"></td>
            <td><input type="radio" name="FaceNum" value="15" <%If ShowRs("Face")=15 then response.Write("Checked")%>>
              <img src="Images/face15.gif" width="22" height="22"></td>
            <td><input type="radio" name="FaceNum" value="16" <%If ShowRs("Face")=16 then response.Write("Checked")%>>
              <img src="Images/face16.gif" width="22" height="22"></td>
            <td><input type="radio" name="FaceNum" value="17" <%If ShowRs("Face")=17 then response.Write("Checked")%>>
              <img src="Images/face17.gif" width="22" height="22"> </td>
            <td><input type="radio" name="FaceNum" value="18" <%If ShowRs("Face")=18 then response.Write("Checked")%>>
              <img src="Images/face18.gif" width="22" height="22"> </td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback">
      <td class="hback" align="right">��������</td>
      <td class="hback" valign="top" >
                <pre id="idTemporary" name="idTemporary" style="display:none"></pre>
                
                <!--�༭����ʼ-->
				<iframe id='NewsContent' src='../Editer/AdminEditer.asp?id=Content' frameborder=0 scrolling=no width='100%' height='280'></iframe>
				<input type="hidden" name="Content" value="<% = HandleEditorContent(ShowRs("Body")) %>">
                <!--�༭������-->
				</td>
    </tr>
    <tr class="hback">
      <td class="hback">&nbsp;
        <input type="hidden" name="NoteID" ID="NoteID" value="<%=ShowRs("ID")%>"></td>
      <td class="hback"><input type="button" name="submit1" value="OK��������"  onClick="SubmitFun();">
        &nbsp;&nbsp;
        <input type="reset" name="reset" value=" ��  �� ">
  </form>
</table>
<%
	end if
end if
set ShowRs=nothing
if NoSqlHack(Request.QueryString("Act"))="Save" then
	ID=NoSqlHack(Request.Form("NoteID"))
	Topic=NoHtmlHackInput(NoSqlHack(Replace(Request.form("Topic"),"'","")))
	IsTop=NoHtmlHackInput(NoSqlHack(Request.form("Style")))
	Content=NoHtmlHackInput(NoSqlHack(Request.form("Content")))
	IsAdmin=NoHtmlHackInput(NoSqlHack(Request.form("IsAdmin")))
	if IsAdmin="" then
		IsAdmin="0"
	end if
	if ID="" then
		strShowErr = "<li>��������!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	if Topic="" then
		strShowErr = "<li>���ⲻ��Ϊ��!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	if isTop="" then
		strShowErr = "<li>��������ֵû������!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	if Content="" then
		strShowErr = "<li>���ݲ���Ϊ��!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	SaveRs.open "Select ID,ClassID,[User],Topic,Body,AddDate,IsTop,State,Face,IsAdmin,LastUpdateDate,LastUpdateUser From FS_WS_BBS Where ID="&CintStr(ID)&"",Conn,1,3
if not SaveRs.eof then
	SaveRs("Topic")=Topic
	SaveRs("Body")=Content
	SaveRs("IsTop")=IsTop
	SaveRs("IsAdmin")=IsAdmin
	SaveRs("Face")=NoSqlHack(Request.Form("FaceNum"))
	SaveRs("LastUpdateUser")=session("Admin_Name")
	SaveRs("LastUpdateDate")=now()
	SaveRs.update
	Set SaveRs=nothing
		Response.write("<script>history.go(-2);</script>")
		Response.end
else
	Set SaveRs=nothing
		strShowErr = "<li>��������!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
end if
end if
Set Conn=nothing
%>
<script language="javascript">
function SubmitFun()
{
	if (frames["NewsContent"].g_currmode!='EDIT') {alert('����ģʽ���޷����棬���л������ģʽ');return false;}
	document.WriteNote.Content.value=frames["NewsContent"].GetNewsContentArray();
	document.WriteNote.submit();
}
</script>
</body>
</html>
<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<%
	Dim Conn,User_Conn
	Dim CharIndexStr,Fs_news
	Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
	Temp_Admin_Name = Session("Admin_Name")
	Temp_Admin_Is_Super = Session("Admin_Is_Super")
	Temp_Admin_FilesTF = Session("Admin_FilesTF")
	MF_Default_Conn
	MF_Session_TF 
	'权限判断
	'Call MF_Check_Pop_TF("NS_Class_000001") 
	'得到会员组列表 
	set Fs_news = new Cls_News
	Fs_News.GetSysParam()
	dim sRootDir,str_CurrPath
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
	if Request.Form("Action") = "save" then
		if Trim(Request.Form("pic"))<>"" then
			Conn.execute("Update FS_NS_News set NewsPicFile='"& NoSqlHack(Request.Form("pic"))&"' where NewsID='"& NoSqlHack(Request.Form("NewsiD"))&"'")
		End if
		if Trim(Request.Form("s_pic"))<>"" then
			Conn.execute("Update FS_NS_News set NewsSmallPicFile='"& NoSqlHack(Request.Form("s_pic"))&"' where NewsID='"& NoSqlHack(Request.Form("NewsiD"))&"'")
		End if
		Response.Write("<script>alert('修改成功');window.close();</script>")
		Response.end
	End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新闻管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body scrolling=yes>
<%
Dim obj_tmp_rs
set obj_tmp_rs = Conn.execute("select Newsid,NewsPicFile,NewsSmallPicFile From FS_NS_News where Newsid='"& NoSqlHack(Request.QueryString("NewsiD"))&"'")
%>
<table width="530" height="443" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="form1" method="post" action="">
    <tr> 
      <td width="211" height="13" class="hback">原图片（小） 
        <input name="Action" type="hidden" id="Action" value="save"> <input name="Newsid" type="hidden" id="Newsid" value="<%=Request.QueryString("Newsid")%>"></td>
      <td width="319" rowspan="4" class="hback">选择小图<br> <input name="s_pic" type="text" id="s_pic"> 
        <input type="button" name="PPPChoose"  value="选择图片" onClick="var returnvalue=OpenWindow('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window);document.form1.s_pic.value=returnvalue;SmallPic.src=returnvalue;"> 
        <br> <br>
        选择大图<br> <input name="pic" type="text" id="pic"> <input type="button" name="PPPChoose2"  value="选择图片" onClick="var returnvalue=OpenWindow('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window);document.form1.pic.value=returnvalue;BigPic.src=returnvalue;"> 
        <br> <br> <input type="submit" name="Submit" value="修改图片">
		 </td>
    </tr>
    <tr> 
      <td height="169" class="hback"><img id="SmallPic" src="<%=obj_tmp_rs("NewsSmallPicFile")%>" width="280"></img></td>
    </tr>
    <tr> 
      <td height="19" class="hback">原图片（大）</td>
    </tr>
    <tr> 
      <td height="188" class="hback"><img id="BigPic" src="<%=obj_tmp_rs("NewsPicFile")%>" width="280"></img></td>
    </tr>
  </form>
</table>

</body>
</html>
<%
set obj_tmp_rs = nothing
set Fs_news = nothing
%>
<script language="JavaScript" type="text/JavaScript" src="js/Public.js"></script>
<script language="JavaScript" type="text/JavaScript">
function opencat(cat)
{
  if(cat.style.display=="none"){
     cat.style.display="";
  } else {
     cat.style.display="none"; 
  }
}
function imgresize(o){
		var parentNode=o.parentNode.parentNode
		if (parentNode){
		if (o.offsetWidth>=parentNode.offsetWidth) o.style.width='98%';
		}else{
		var parentNode=o.parentNode
		if (parentNode){
			if (o.offsetWidth>=parentNode.offsetWidth) o.style.width='98%';
			}
		}
}
</script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






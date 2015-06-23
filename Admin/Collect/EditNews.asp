<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<%
Dim Conn,CollectConn
MF_Default_Conn
MF_Collect_Conn
MF_Session_TF

Dim NewsIDStr,Result,RsNewsObj,Sql
Dim Title,Links,Content,AddDate,SiteID,Author,SourceStr
Result = Request("Result")
NewsIDStr = Request("NewsIDStr")
if Result = "Submit" then
	Title = Request.Form("Title")
	Links = Request.Form("Links")
	Content = Request.Form("Content")
	AddDate = Request.Form("AddDate")
	SiteID = Request.Form("SiteID")
	Author = Request.Form("Author")
	SourceStr = Request.Form("Source")
	if NewsIDStr <> "" then
		Sql = "Select * from FS_News where ID=" & CintStr(NewsIDStr)&""
		'On Error Resume Next
		Set RsNewsObj = Server.CreateObject(G_FS_RS)
		RsNewsObj.Open Sql,CollectConn,3,3
		RsNewsObj("Title") = NoSqlHack(Title)
		RsNewsObj("Links") = NoSqlHack(Links)
		RsNewsObj("Content") = NoSqlHack(Content)
		RsNewsObj("AddDate") = NoSqlHack(AddDate)
		RsNewsObj("Author") = NoSqlHack(Author)
		RsNewsObj("Source") = NoSqlHack(SourceStr)
		RsNewsObj("SiteID") = NoSqlHack(SiteID)
		RsNewsObj.UpDate
		RsNewsObj.Close
		Set RsNewsObj = Nothing
		if Err.Number <> 0 then
%>
	<script language="JavaScript">
	alert('修改失败');
	</script>
<%
		else
			Response.Redirect("Check.asp")
		end if
	else
%>
	<script language="JavaScript">
	alert('修改的新闻不存在');
	</script>
<%
	end if
else
	if NewsIDStr <> "" then
		Sql = "Select * from FS_News where ID=" & CintStr(NewsIDStr)
		Set RsNewsObj = CollectConn.Execute(Sql)
		if Not RsNewsObj.Eof then
			Title = RsNewsObj("Title")
			Links = RsNewsObj("Links")
			Content = RsNewsObj("Content")
			AddDate = RsNewsObj("AddDate")
			SiteID = RsNewsObj("SiteID")
			Author = RsNewsObj("Author")
			SourceStr = RsNewsObj("Source")
		else
%>
	<script language="JavaScript">
	alert('新闻不存在');
	</script>
<%
		end if
	else
%>
	<script language="JavaScript">
	alert('参数错误');
	</script>
<%
	end if
end if

Dim SiteList,RsSiteObj
Set RsSiteObj = CollectConn.Execute("Select ID,SiteName from FS_Site where 1=1 order by id desc")
do while Not RsSiteObj.Eof
	if Clng(RsSiteObj("ID")) = Clng(SiteID) then
		SiteList = SiteList & "<option selected value=" & RsSiteObj("ID") & "" & ">" & RsSiteObj("SiteName") & "</option><br>"
	else
		SiteList = SiteList & "<option value=" & RsSiteObj("ID") & "" & ">" & RsSiteObj("SiteName") & "</option><br>"
	end if
	RsSiteObj.MoveNext	
loop
RsSiteObj.Close
Set RsSiteObj = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改新闻</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<style>
.LableWindow
{
	border-right: 1px solid;
	border-left: 1px solid;
	border-bottom: 1px solid;
	border-color: Black;
	cursor: default;
}
.LableDefault
{
	border-right: 1px solid;
	border-top: 1px solid;
	font-size: 12px;
	border-left: 1px solid;
	border-bottom: 1px solid;
	font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
	border-color: Black;
	cursor: default;

}
.LableSelected
{
	border-right: 1px solid;
	border-top: 1px solid;
	font-size: 12px;
	border-left: 1px solid;
	font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
	border-color: Black;
	cursor: default;

}
.ToolBarButtonLine {
	border-bottom: 1px solid;
	font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
	border-color: Black;
}
</style>
<link href="FS_css.css" rel="stylesheet">
<script language="JavaScript" type="text/JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<body topmargin="2" leftmargin="2">
<form action="" method="post" name="NewsForm">
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr class="hback"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td style="cursor:hand;" width="35" align="center" alt="保存" onClick="document.NewsForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
			<td width=2 class="Gray">|</td>
			<td style="cursor:hand;" width="35" align="center" alt="后退" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp; <input name="Result" type="hidden" id="Result" value="Submit"> 
              <input value="<% = NewsIDStr %>" name="NewsIDStr" type="hidden" id="NewsIDStr"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1">
	<tr>
		
      <td><table width="100%" border="0" cellspacing="1" cellpadding="5" class="table">
          <tr class="hback"> 
            <td width="120" height="26"> 
              <div align="center">新闻标题</div></td>
            <td> <input name="Title" type="text" id="Title2" style="width:100%;" value="<% = Title %>"></td>
          </tr>
          <tr class="hback"> 
            <td height="26"  class="hback"> 
              <div align="center">新闻联接</div></td>
            <td> 
              <input name="Links" type="text" id="Links" style="width:100%;" value="<% = Links %>"></td>
          </tr>
          <tr class="hback"> 
            <td><div align="center">采集站点</div></td>
            <td><select style="width:100%;" name="SiteID">
                <% = SiteList %>
              </select></td>
          </tr>
          <tr class="hback"> 
            <td height="26">
<div align="center">作&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;者</div></td>
            <td><input style="width:100%;" type="text" name="Author" value="<% = Author %>"></td>
          </tr>
          <tr class="hback"> 
            <td height="26">
<div align="center">来&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;源</div></td>
            <td><input style="width:100%;" type="text" name="Source" value="<% = SourceStr %>"></td>
          </tr>
          <tr class="hback"> 
            <td height="26"> 
              <div align="center">采集日期</div></td>
            <td><input name="AddDate" type="text" id="AddDate2" style="width:100%;" value="<% = AddDate %>"> 
              <div align="center"></div></td>
          </tr>
        </table></td>
	</tr>
    <tr class="hback"> 
      <td height="20" colspan="2">
<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td id="EditCodeBtn" width="100" class="LableSelected" onClick="CodeContent();" bgcolor="#EFEFEF"> <div align="center"> 
                代 码</div></td>
            <td width="5" class="ToolBarButtonLine">&nbsp;</td>
			<td id="PreviewBtn" width="100" class="LableDefault" onClick="Preview();"> <div align="center">预 
                览</div></td>
            <td class="ToolBarButtonLine">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr id="EditCodeArea"> 
      <td height="300" colspan="2"> 
        <textarea name="Content" id="NewsContent" rows="20" style="width:100%;"><% = Content %></textarea></td>
    </tr>
    <tr id="PreviewArea" style="display:none;" bgcolor="#EFEFEF"> 
      <td height="300" colspan="2"> 
        <iframe name="PreviewContent" frameborder="1" class="Composition" ID="PreviewContent" MARGINHEIGHT="1" MARGINWIDTH="1" width="98%" scrolling="yes" src="about:blank"></iframe></td>
    </tr>
</table>
</form>
</body>
</html>
<%
Set CollectConn = Nothing
Set Conn = Nothing
Set RsNewsObj = Nothing
%>
<script language="JavaScript">
function SetEditAreaHeight()
{
	var BodyHeight=document.body.clientHeight;
	var EditAreaHeight=BodyHeight-200;
	document.all.NewsContent.style.height=EditAreaHeight;
	document.all.PreviewContent.height=EditAreaHeight;
}
SetEditAreaHeight();
window.onresize=SetEditAreaHeight;
function Preview()
{
	var TempStr='';
	document.all.EditCodeArea.style.display='none';
	document.all.PreviewArea.style.display='';
	PreviewContent.document.write('<head><link href=\"../../CSS/FS_css.css\" type=\"text/css\" rel=\"stylesheet\"></head><body MONOSPACE>');
	PreviewContent.document.body.innerHTML=document.all.Content.value;
	document.all.PreviewBtn.className='LableSelected';
	document.all.PreviewBtn.style.backgroundColor='#EFEFEF';
	document.all.EditCodeBtn.className='LableDefault';
	document.all.EditCodeBtn.style.backgroundColor='#FFFFFF';
}
function CodeContent()
{
	document.all.EditCodeArea.style.display='';
	document.all.PreviewArea.style.display='none';
	document.all.EditCodeBtn.className='LableSelected';
	document.all.EditCodeBtn.style.backgroundColor='#EFEFEF';
	document.all.PreviewBtn.className='LableDefault';
	document.all.PreviewBtn.style.backgroundColor='#FFFFFF';
}
</script>






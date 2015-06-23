<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<%
Dim Conn,CollectConn
MF_Default_Conn
MF_Collect_Conn
MF_Session_TF
if not MF_Check_Pop_TF("CS001") then Err_Show
Dim RsEditObj,EditSql,SiteID,ObjUrl,WebCharset
Set RsEditObj = Server.CreateObject(G_FS_RS)
SiteID = Request("SiteID")
if SiteID <> "" then
	EditSql="Select * from FS_Site where ID=" & CintStr(SiteID)
	RsEditObj.Open EditSql,CollectConn,1,3
	if RsEditObj.Eof then
		Response.write"<script>alert(""没有修改的站点"");location.href=""javascript:history.back()"";</script>"
		Response.end
	else
		ObjUrl = RsEditObj("ObjUrl")
		WebCharset = RsEditObj("WebCharset")
	end if
else
	Response.write"<script>alert(""没有修改的站点"");location.href=""javascript:history.back()"";</script>"
	Response.end
end if

Dim ListHeadSetting,ListFootSetting,OtherPageHeadSetting,OtherPageFootSetting
Dim IndexRule,StartPageNum,EndPageNum,HandPageContent,OtherType
Dim ListSetting,OtherPageSetting
ListSetting = split(Request.Form("ListSetting"),"[列表内容]",-1,1)
ListHeadSetting = NoSqlHack(ListSetting(0))
ListFootSetting = NoSqlHack(ListSetting(1))
If Err Or ListHeadSetting="" Or ListFootSetting="" Then
	ListHeadSetting = "<body"
	ListFootSetting = "</body>"
	Err.clear
End If
If InStr(Request.Form("OtherPageSetting"),"[其他页面]")<>0 then
	OtherPageSetting = split(Request.Form("OtherPageSetting"),"[其他页面]",-1,1)
	OtherPageHeadSetting = NoSqlHack(OtherPageSetting(0))
	OtherPageFootSetting = NoSqlHack(OtherPageSetting(1))
	OtherPageHeadSetting=replace(OtherPageHeadSetting,"''","'")
	OtherPageFootSetting=replace(OtherPageFootSetting,"''","'")
End if
OtherType = NoSqlHack(Request.Form("OtherType"))
IndexRule = NoSqlHack(Request.Form("IndexRule"))
StartPageNum = NoSqlHack(Request.Form("StartPageNum"))
EndPageNum = NoSqlHack(Request.Form("EndPageNum"))
HandPageContent = Request.Form("HandPageContent")
if Request.Form("Result") = "Edit" then
    Dim RsAddObj,sql
	Set RsAddObj = Server.CreateObject(G_FS_RS)
	Sql = "select * from FS_Site where id=" & CintStr(Request.Form("SiteID"))
	RsAddObj.Open Sql,CollectConn,1,3
	RsAddObj("ListHeadSetting") = ListHeadSetting
	RsAddObj("ListFootSetting") = ListFootSetting
	Select Case OtherType
		Case "0"
			RsAddObj("OtherType") = OtherType
			RsAddObj("IndexRule") = ""
			RsAddObj("StartPageNum") = ""
			RsAddObj("EndPageNum") = ""
			RsAddObj("HandPageContent") = ""
			RsAddObj("OtherPageHeadSetting") = ""
			RsAddObj("OtherPageFootSetting") = ""
		Case "1"
			RsAddObj("OtherType") = OtherType
			RsAddObj("IndexRule") = ""
			RsAddObj("StartPageNum") = ""
			RsAddObj("EndPageNum") = ""
			RsAddObj("HandPageContent") = ""
			RsAddObj("OtherPageHeadSetting") = OtherPageHeadSetting
			RsAddObj("OtherPageFootSetting") = OtherPageFootSetting
		Case "2"
			RsAddObj("OtherType") = OtherType
			RsAddObj("IndexRule") = IndexRule
			RsAddObj("StartPageNum") = StartPageNum
			RsAddObj("EndPageNum") = EndPageNum
			RsAddObj("HandPageContent") = ""
			RsAddObj("OtherPageHeadSetting") = ""
			RsAddObj("OtherPageFootSetting") = ""
		Case "3"
			RsAddObj("OtherType") = OtherType
			RsAddObj("IndexRule") = ""
			RsAddObj("StartPageNum") = ""
			RsAddObj("EndPageNum") = ""
			RsAddObj("HandPageContent") = HandPageContent
			RsAddObj("OtherPageHeadSetting") = ""
			RsAddObj("OtherPageFootSetting") = ""
		Case Else
			RsAddObj("OtherType") = 0
			RsAddObj("IndexRule") = ""
			RsAddObj("StartPageNum") = ""
			RsAddObj("EndPageNum") = ""
			RsAddObj("HandPageContent") = ""
			RsAddObj("OtherPageHeadSetting") = ""
			RsAddObj("OtherPageFootSetting") = ""
	End Select
	RsAddObj.update
	RsAddObj.close
	Set RsAddObj = Nothing
end if
Dim ResponseAllStr,NewsListStr

ResponseAllStr = GetPageContent(ObjURL,WebCharset)
NewsListStr = GetOtherContent(ResponseAllStr,ListHeadSetting,ListFootSetting)
NewsListStr = Replace(Replace(NewsListStr,"""","%22"),"'","%27")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自动新闻采集—站点设置</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<form name="form1" method="post" action="SiteFourStep.asp" id="Form1">
<table width="98%" height="20" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback" >
			  <td style="cursor:hand" width="50" align="center" alt="第三步" onClick="window.location.href='javascript:history.go(-1)';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">上一步</td>
			  <td width=2 class="Gray">|</td>
			  <td style="cursor:hand" width="50" align="center" alt="第四步" onClick="document.form1.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">下一步</td>
			  <td width=2 class="Gray">|</td>
		      <td style="cursor:hand" width="35" align="center" alt="后退" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
			  <td>&nbsp; <input name="SiteID" type="hidden" id="SiteID2" value="<% = SiteID %>"> 
				<input name="Result" type="hidden" id="Result2" value="Edit">
              <input type="hidden" name="NewsListStr" value="<% = NewsListStr %>"></td>
        </tr>
  </table><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	  <tr class="hback" > 
      <td width="20%"> 
        <div align="center">列表URL</div></td>
		<td>	&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form1.LinkSetting.rows>2)document.Form1.LinkSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form1.LinkSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
		&nbsp;&nbsp;可用标签：<font onClick="addTag('[列表URL]')" style="CURSOR: hand"><b>[列表URL]</b></font>&nbsp;&nbsp;&nbsp;<font onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><br>
		 <textarea onFocus="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="LinkSetting" cols="50" rows="6" id="textarea2" style="width:100%;"><%=RsEditObj("LinkHeadSetting")%>[列表URL]<%=RsEditObj("LinkFootSetting")%></textarea></td>
	  </tr>
</table>
</form>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="0" class="table">
  <tr class="hback" >
    <td height="28" class="xingmu"> 
      <div align="center">代码</div></td>
  </tr>
  <tr class="hback" >
    <td height="20"><textarea name="CodeArea" rows="18" style="width:100%;"></textarea></td>
  </tr>
  <tr class="hback" > 
    <td height="28" class="xingmu"> 
      <div align="center">结果</div></td>
  </tr>
  <tr class="hback" > 
    <td><iframe frameborder="1" name="PreviewArea" src="about:blank" ID="PreviewArea" MARGINHEIGHT="1" MARGINWIDTH="1" height="300" width="100%" scrolling="yes"></iframe></td>
  </tr>
</table>
<p><p><p>
</body>
</html>
<%
Set CollectConn = Nothing
Set Conn = Nothing
Set RsEditObj = Nothing
%>
<script language="JavaScript">
function document.onreadystatechange()
{
	document.all.CodeArea.value=unescape(document.form1.NewsListStr.value);
	frames["PreviewArea"].document.write(unescape(document.form1.NewsListStr.value));
}

currObj = "uuuu";
function getActiveText(obj)
{
	currObj = obj;
}

function addTag(code)
{
	addText(code);
}

function addText(ibTag)
{
	var isClose = false;
	var obj_ta = currObj;
	if (obj_ta.isTextEdit)
	{
		obj_ta.focus();
		var sel = document.selection;
		var rng = sel.createRange();
		rng.colapse;
		if((sel.type == "Text" || sel.type == "None") && rng != null)
		{
			rng.text = ibTag;
		}
		obj_ta.focus();
		return isClose;
	}
	else return false;
}	
</script>






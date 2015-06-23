<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<%
Dim Conn,CollectConn
MF_Default_Conn
MF_Collect_Conn
MF_Session_TF

Dim RsEditObj,EditSql,SiteID
Dim NewsLinkStr,WebCharset
Set RsEditObj = Server.CreateObject(G_FS_RS)
SiteID = Request("SiteID")
if SiteID <> "" then
	EditSql="Select * from FS_Site where ID=" & SiteID
	RsEditObj.Open EditSql,CollectConn,1,3
	if RsEditObj.Eof then
		Response.write"<script>alert(""没有修改的站点"");location.href=""javascript:history.back()"";</script>"
		Response.end
	Else
		WebCharset = RsEditObj("WebCharset")
	end if
else
	Response.write"<script>alert(""没有修改的站点"");location.href=""javascript:history.back()"";</script>"
	Response.end
end if

Dim PageTitleHeadSetting,PageTitleFootSetting,PagebodyHeadSetting,PagebodyFootSetting
Dim OtherNewsPageHeadSetting,OtherNewsPageFootSetting
Dim OtherNewsType,OtherNewsPageIndexSetting
Dim OtherNewsPageIndexSettingStartPageNum,OtherNewsPageIndexSettingEndPageNum,OtherNewsPageIndexSettingHandPageContent
Dim AuthorHeadSetting,AuthorFootSetting
Dim SourceHeadSetting,SourceFootSetting
Dim AddDateHeadSetting,AddDateFootSetting
Dim HandSetAuthor,HandSetSource,HandSetAddDate
Dim TextTF,IsStyle,IsDiv,IsA,IsClass,IsFont,IsSpan,IsObjectTF,IsIFrame,IsScript
Dim PageTitleSetting,PagebodySetting,OtherNewsPageSetting,AuthorSetting,SourceSetting,AddDateSetting
If InStr(Request.Form("PageTitleSetting"),"[标题]") = 0 Then
	Response.Write "<script>alert('新闻标题没有设置或设置不正确！');history.back();</script>"
	Response.End 
End If
If InStr(Request.Form("PagebodySetting"),"[内容]") = 0 Then
	Response.Write "<script>alert('新闻内容没有设置或设置不正确！');history.back();</script>"
	Response.End 
End if
PageTitleSetting = Split(Request.Form("PageTitleSetting"),"[标题]",-1,1)
PageTitleHeadSetting = PageTitleSetting(0)
PageTitleFootSetting = PageTitleSetting(1)
PagebodySetting = Split(Request.Form("PagebodySetting"),"[内容]",-1,1)
PagebodyHeadSetting = PagebodySetting(0)
PagebodyFootSetting = PagebodySetting(1)
If InStr(Request.Form("OtherNewsPageSetting"),"[分页新闻]")<>0 then
	OtherNewsPageSetting = Split(Request.Form("OtherNewsPageSetting"),"[分页新闻]",-1,1)
	OtherNewsPageHeadSetting = OtherNewsPageSetting(0)
	OtherNewsPageFootSetting = OtherNewsPageSetting(1)
End If
'-----2007-01-17 Eidt By Ken   修正采集新闻内容页分页规则不能保存
If Request.Form("OtherNewsType") <> "" And IsNumeric(Request.Form("OtherNewsType")) Then
	OtherNewsType = Cint(Request.Form("OtherNewsType"))
Else
	OtherNewsType = 0
End If		
If Trim(Request.Form("OtherNewsPageIndexSetting")) <> "" Then
	OtherNewsPageIndexSetting = Trim(Request.Form("OtherNewsPageIndexSetting"))
Else
	OtherNewsPageIndexSetting = ""
End If
If InStr(Request.Form("AuthorSetting"),"[作者]")<>0 then
	AuthorSetting = Split(Request.Form("AuthorSetting"),"[作者]",-1,1)
	AuthorHeadSetting = AuthorSetting(0)
	AuthorFootSetting = AuthorSetting(1)
End If 
If InStr(Request.Form("SourceSetting"),"[来源]")<>0 then
	SourceSetting = Split(Request.Form("SourceSetting"),"[来源]",-1,1)
	SourceHeadSetting = SourceSetting(0)
	SourceFootSetting = SourceSetting(1)
End If
If InStr(Request.Form("AddDateSetting"),"[加入时间]")<>0 then
	AddDateSetting = Split(Request.Form("AddDateSetting"),"[加入时间]",-1,1)
	AddDateHeadSetting = AddDateSetting(0)
	AddDateFootSetting = AddDateSetting(1)
End If 
HandSetAuthor = Request.Form("HandSetAuthor")
HandSetSource = Request.Form("HandSetSource")
HandSetAddDate = Request.Form("HandSetAddDate")
if Request.Form("Result") = "Edit" then
    Dim RsAddObj,sql
	Set RsAddObj = Server.CreateObject(G_FS_RS)
	Sql = "select * from FS_Site where id=" & CintStr(Request.Form("SiteID"))
	RsAddObj.Open Sql,CollectConn,1,3
	TextTF = RsAddObj("TextTF")
	IsStyle = RsAddObj("IsStyle")
	IsDiv = RsAddObj("IsDiv")
	IsA = RsAddObj("IsA")
	IsClass = RsAddObj("IsClass")
	IsFont = RsAddObj("IsFont")
	IsSpan = RsAddObj("IsSpan")
	IsObjectTF = RsAddObj("IsObject")
	IsIFrame = RsAddObj("IsIFrame")
	IsScript = RsAddObj("IsScript")
	RsAddObj("PagebodyHeadSetting") = PagebodyHeadSetting
	RsAddObj("PagebodyFootSetting") = PagebodyFootSetting
	RsAddObj("PageTitleHeadSetting") = PageTitleHeadSetting
	RsAddObj("PageTitleFootSetting") = PageTitleFootSetting
	RsAddObj("OtherNewsPageHeadSetting") = OtherNewsPageHeadSetting
	RsAddObj("OtherNewsPageFootSetting") = OtherNewsPageFootSetting
	'-------2007-01-17
	RsAddObj("OtherNewsType") = OtherNewsType
	RsAddObj("OtherNewsPageIndexSetting") = OtherNewsPageIndexSetting
	RsAddObj("OtherNewsPageIndexSettingStartPageNum") = 0
	RsAddObj("OtherNewsPageIndexSettingEndPageNum") = 0
	RsAddObj("OtherNewsPageIndexSettingHandPageContent") = ""
	'------------------
	RsAddObj("AuthorHeadSetting") = AuthorHeadSetting
	RsAddObj("AuthorFootSetting") = AuthorFootSetting
	RsAddObj("SourceHeadSetting") = SourceHeadSetting
	RsAddObj("SourceFootSetting") =SourceFootSetting
	RsAddObj("AddDateHeadSetting") = AddDateHeadSetting
	RsAddObj("AddDateFootSetting") = AddDateFootSetting
	RsAddObj("HandSetAuthor") = HandSetAuthor
	RsAddObj("HandSetSource") = HandSetSource
	if IsDate(HandSetAddDate) then
		RsAddObj("HandSetAddDate") = HandSetAddDate
	end if
	RsAddObj.update
	RsAddObj.close
	Set RsAddObj = Nothing
end if

NewsLinkStr = NoSqlHack(Request("NewsLinkStr"))
Dim ResponseAllStr,TitleStr,NewsBodyStr,AuthorStr,SourceStr,AddDateStr
ResponseAllStr = GetPageContent(NewsLinkStr,WebCharset)
TitleStr = GetOtherContent(ResponseAllStr,PageTitleHeadSetting,PageTitleFootSetting)
NewsBodyStr = GetOtherContent(ResponseAllStr,PagebodyHeadSetting,PagebodyFootSetting)
NewsBodyStr = ReplaceContentStr(NewsBodyStr)
if HandSetAuthor <> "" then
	AuthorStr = HandSetAuthor
else
	if AuthorHeadSetting <> "" And AuthorFootSetting <> "" then 
		AuthorStr = GetOtherContent(ResponseAllStr,AuthorHeadSetting,AuthorFootSetting)
	end if
end if
if HandSetSource <> "" then
	SourceStr = HandSetSource
else
	if SourceHeadSetting <> "" And SourceFootSetting <> "" then 
		SourceStr = GetOtherContent(ResponseAllStr,SourceHeadSetting,SourceFootSetting)
	end if
end if
if HandSetAddDate <> "" then
	if Not IsDate(HandSetAddDate) then
		AddDateStr = Now
	else
		AddDateStr = HandSetAddDate
	end if
else
	if AddDateHeadSetting <> "" And AddDateFootSetting <> "" then 
		AddDateStr = GetOtherContent(ResponseAllStr,AddDateHeadSetting,AddDateFootSetting)
	end if
end if
NewsBodyStr = Replace(Replace(NewsBodyStr,"""","%22"),"'","%27")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自动新闻采集—站点设置</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr class="hback"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="50" style="cursor:hand" align="center" alt="第四步" onClick="window.location.href='javascript:history.go(-1)';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">上一步</td>
			<td width=2 class="Gray">|</td>
            <td width="35" style="cursor:hand" align="center" alt="完成" onClick="window.location.href='Site.asp';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">完成</td>
			<td width=2 class="Gray">|</td>
		    <td width="35" style="cursor:hand" align="center" alt="后退" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td><input type="hidden" name="NewsBodyStr" value="<% = NewsBodyStr %>"> &nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback">
    <td height="26">
<div align="center"><strong><font size="3"><% = TitleStr %></font></strong></div></td>
  </tr>
	<tr class="hback">
	  
    <td height="26">
<div align="center"><strong>作者</strong>： 
        <% = AuthorStr %>
        &nbsp;&nbsp;<strong>来源</strong>： 
        <% = SourceStr %>
        &nbsp;&nbsp;<strong>时间</strong>： 
        <% = AddDateStr %></div></td>
	</tr>
	<tr class="hback">
	  <td><iframe frameborder="1" name="PreviewArea" src="about:blank" ID="PreviewArea" MARGINHEIGHT="1" MARGINWIDTH="1" height="480" width="100%" scrolling="yes"></iframe></td>
	</tr>
</table>
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
	frames["PreviewArea"].document.write(unescape(document.all.NewsBodyStr.value));
}
</script>






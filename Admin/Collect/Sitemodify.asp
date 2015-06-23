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

Dim RsEditObj,EditSql,SiteID
Set RsEditObj = Server.CreateObject(G_FS_RS)
SiteID = Request("SiteID")
if SiteID <> "" then
	EditSql="Select * from FS_Site where ID=" & CintStr(SiteID)
	RsEditObj.Open EditSql,CollectConn,1,3
	if RsEditObj.Eof then
		Response.write"<script>alert(""没有修改的站点"");location.href=""javascript:history.back()"";</script>"
		Response.end
	end if
else
	Response.write"<script>alert(""没有修改的站点"");location.href=""javascript:history.back()"";</script>"
	Response.end
end if

Function SiteFolderList()
	Dim ClassListObj,SelectStr
	Set ClassListObj = CollectConn.Execute("Select * from FS_SiteFolder where 1=1 order by ID desc")
	do while Not ClassListObj.Eof
		if RsEditObj("Folder") = ClassListObj("ID") then
			SelectStr = "selected"
		else
			SelectStr = ""
		end if
		SiteFolderList = SiteFolderList & "<option " & SelectStr & " value="&ClassListObj("ID")&"" & ">&nbsp;&nbsp;|--" & ClassListObj("SiteFolder") & "</option><br>"
		ClassListObj.MoveNext	
	loop
	ClassListObj.Close
	Set ClassListObj = Nothing
End Function

if Request.Form("Result") = "Edit" then
    Dim RsAddObj,sql
    if Request.Form("SiteName")="" or Request.Form("objURL")="" then
		Response.write"<script>alert(""请填写完整！"");location.href=""javascript:history.back()"";</script>"
		Response.end
	end if
	Set RsAddObj = Server.CreateObject(G_FS_RS)
	Sql = "select * from FS_Site where id=" & CintStr(Request.Form("SiteID"))
	RsAddObj.Open Sql,CollectConn,1,3
	RsAddObj("SiteName") = NoSqlHack(Request.Form("SiteName"))
	RsAddObj("objURL") = NoSqlHack(Request.Form("objURL"))
	RsAddObj("Folder")=NoSqlHack(Request.Form("SiteFolder")) '修改分类
	'On Error Resume Next 
	Dim ListSetting,LinkSetting,PageBodySetting,PageTitleSetting,OtherNewsPageSetting,AuthorSetting,SourceSetting,AddDateSetting,OtherPageSetting,StrErr
	StrErr = ""
	If InStr(Request.Form("ListSetting"),"[列表内容]")<>0 then
		ListSetting = Split(Request.Form("ListSetting"),"[列表内容]",-1,1)
		RsAddObj("ListHeadSetting") = ListSetting(0)
		RsAddObj("ListFootSetting") = ListSetting(1)
	ElseIf ListSetting(0)="" Or ListSetting(1)="" Or err Then
		Err.clear
		RsAddObj("ListHeadSetting") = "<body"
		RsAddObj("ListFootSetting") = "</body>"
	End If
	LinkSetting = Split(Request.Form("LinkSetting"),"[列表URL]",-1,1)
	RsAddObj("LinkHeadSetting") = LinkSetting(0)
	RsAddObj("LinkFootSetting") = LinkSetting(1)
	If err Then
		StrErr = "列表URL没有设置或设置不正确！"
		Err.clear
	End if
	PageBodySetting = Split(Request.Form("PageBodySetting"),"[新闻内容]",-1,1)
	RsAddObj("PagebodyHeadSetting") = PageBodySetting(0)
	RsAddObj("PagebodyFootSetting") = PageBodySetting(1)
	If err Then
		StrErr = StrErr & "\r\n新闻内容没有设置或设置不正确！"
		Err.clear
	End if
	PageTitleSetting = Split(Request.Form("PageTitleSetting"),"[新闻标题]",-1,1) 
	RsAddObj("PageTitleHeadSetting") = PageTitleSetting(0)
	RsAddObj("PageTitleFootSetting") = PageTitleSetting(1)
	If err Then
		StrErr = StrErr & "\r\n新闻标题没有设置或设置不正确！"
		Err.clear
	End If
	If InStr(Request.Form("AuthorSetting"),"[作者]")<>0 then
		AuthorSetting = Split(Request.Form("AuthorSetting"),"[作者]",-1,1)
		RsAddObj("AuthorHeadSetting") = AuthorSetting(0)
		RsAddObj("AuthorFootSetting") = AuthorSetting(1)
	End If
	If InStr(Request.Form("SourceSetting"),"[来源]")<>0 then
		SourceSetting = Split(Request.Form("SourceSetting"),"[来源]",-1,1)
		RsAddObj("SourceHeadSetting") = SourceSetting(0)
		RsAddObj("SourceFootSetting") = SourceSetting(1)
	End If
	If InStr(Request.Form("AddDateSetting"),"[加入时间]")<>0 then
		AddDateSetting = Split(Request.Form("AddDateSetting"),"[加入时间]",-1,1)
		RsAddObj("AddDateHeadSetting") = AddDateSetting(0)
		RsAddObj("AddDateFootSetting") = AddDateSetting(1)
	End if
	
	Select Case Request.Form("OtherNewsType")
		Case "0"
			RsAddObj("OtherNewsType") = NoSqlHack(Request.Form("OtherNewsType"))
			RsAddObj("OtherNewsPageHeadSetting") = ""
			RsAddObj("OtherNewsPageFootSetting") = ""
			RsAddObj("OtherNewsPageIndexSetting") = ""
			RsAddObj("OtherNewsPageIndexSettingStartPageNum") = ""
			RsAddObj("OtherNewsPageIndexSettingEndPageNum") = ""
			RsAddObj("OtherNewsPageIndexSettingHandPageContent") = ""
		Case "1"
			RsAddObj("OtherNewsType") = NoSqlHack(Request.Form("OtherNewsType"))
			If InStr(Request.Form("OtherNewsPageSetting"),"[分页新闻]")<>0 Then
				OtherNewsPageSetting = Split(Request.Form("OtherNewsPageSetting"),"[分页新闻]",-1,1)
				RsAddObj("OtherNewsPageHeadSetting") = NoSqlHack(OtherNewsPageSetting(0))
				RsAddObj("OtherNewsPageFootSetting") = NoSqlHack(OtherNewsPageSetting(1))
			End if
			RsAddObj("OtherNewsPageIndexSetting") = ""
			RsAddObj("OtherNewsPageIndexSettingStartPageNum") = ""
			RsAddObj("OtherNewsPageIndexSettingEndPageNum") = ""
			RsAddObj("OtherNewsPageIndexSettingHandPageContent") = ""
		Case "2"
			RsAddObj("OtherNewsType") = NoSqlHack(Request.Form("OtherNewsType"))
			RsAddObj("OtherNewsPageHeadSetting") = ""
			RsAddObj("OtherNewsPageFootSetting") = ""
			RsAddObj("OtherNewsPageIndexSetting") = NoSqlHack(Request.Form("OtherNewsPageIndexSetting"))
			RsAddObj("OtherNewsPageIndexSettingStartPageNum") = NoSqlHack(Request.Form("OtherNewsPageIndexSettingStartPageNum"))
			RsAddObj("OtherNewsPageIndexSettingEndPageNum") = NoSqlHack(Request.Form("OtherNewsPageIndexSettingEndPageNum"))
			RsAddObj("OtherNewsPageIndexSettingHandPageContent") = ""
		Case "3"
			RsAddObj("OtherNewsType") = NoSqlHack(Request.Form("OtherNewsType"))
			RsAddObj("OtherNewsPageHeadSetting") = ""
			RsAddObj("OtherNewsPageFootSetting") = ""
			RsAddObj("OtherNewsPageIndexSetting") = ""
			RsAddObj("OtherNewsPageIndexSettingStartPageNum") = ""
			RsAddObj("OtherNewsPageIndexSettingEndPageNum") = ""
			RsAddObj("OtherNewsPageIndexSettingHandPageContent") = NoSqlHack(Request.Form("OtherNewsPageIndexSettingHandPageContent"))
		Case Else
			RsAddObj("OtherNewsType") = "0"
			RsAddObj("OtherNewsPageHeadSetting") = ""
			RsAddObj("OtherNewsPageFootSetting") = ""
			RsAddObj("OtherNewsPageIndexSetting") = ""
			RsAddObj("OtherNewsPageIndexSettingStartPageNum") = ""
			RsAddObj("OtherNewsPageIndexSettingEndPageNum") = ""
			RsAddObj("OtherNewsPageIndexSettingHandPageContent") = ""
	End Select
	
	If StrErr<>"" Then
		Err.clear
		Response.Write "<script>alert('"& StrErr &"');history.back();</script>"
		Response.End
	End If 
	Select Case Request.Form("OtherType")
		Case "0"
			RsAddObj("OtherType") = NoSqlHack(Request.Form("OtherType"))
			RsAddObj("IndexRule") = ""
			RsAddObj("StartPageNum") = ""
			RsAddObj("EndPageNum") = ""
			RsAddObj("HandPageContent") = ""
			RsAddObj("OtherPageHeadSetting") = ""
			RsAddObj("OtherPageFootSetting") = ""
		Case "1"
			RsAddObj("OtherType") = NoSqlHack(Request.Form("OtherType"))
			RsAddObj("IndexRule") = ""
			RsAddObj("StartPageNum") = ""
			RsAddObj("EndPageNum") = ""
			RsAddObj("HandPageContent") = ""
			OtherPageSetting = Split(Request.Form("OtherPageSetting"),"[其他页面]",-1,1)
			RsAddObj("OtherPageHeadSetting") = NoSqlHack(OtherPageSetting(0))
			RsAddObj("OtherPageFootSetting") = NoSqlHack(OtherPageSetting(1))
		Case "2"
			RsAddObj("OtherType") = NoSqlHack(Request.Form("OtherType"))
			RsAddObj("IndexRule") = NoSqlHack(Request.Form("IndexRule"))
			RsAddObj("StartPageNum") = NoSqlHack(Request.Form("StartPageNum"))
			RsAddObj("EndPageNum") = NoSqlHack(Request.Form("EndPageNum"))
			RsAddObj("HandPageContent") = ""
			RsAddObj("OtherPageHeadSetting") = ""
			RsAddObj("OtherPageFootSetting") = ""
		Case "3"
			RsAddObj("OtherType") = NoSqlHack(Request.Form("OtherType"))
			RsAddObj("IndexRule") = ""
			RsAddObj("StartPageNum") = ""
			RsAddObj("EndPageNum") = ""
			RsAddObj("HandPageContent") = NoSqlHack(Request.Form("HandPageContent"))
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
	RsAddObj("HandSetAuthor") = NoSqlHack(Request.Form("HandSetAuthor"))
	RsAddObj("HandSetSource") = NoSqlHack(Request.Form("HandSetSource"))
	if IsDate(Request.Form("HandSetAddDate")) then
		RsAddObj("HandSetAddDate") = NoSqlHack(Request.Form("HandSetAddDate"))
	end if
	if Request.Form("IsIFrame") = "1" then
		RsAddObj("IsIFrame") = True
	else
		RsAddObj("IsIFrame") = False
	end if
	if Request.Form("IsScript") = "1" then
		RsAddObj("IsScript") = True
	else
		RsAddObj("IsScript") = False
	end if
	if Request.Form("IsClass") = "1" then
		RsAddObj("IsClass") = True
	else
		RsAddObj("IsClass") = False
	end if
	if Request.Form("IsFont") = "1" then
		RsAddObj("IsFont") = True
	else
		RsAddObj("IsFont") = False
	end if
	if Request.Form("IsSpan") = "1" then
		RsAddObj("IsSpan") = True
	else
		RsAddObj("IsSpan") = False
	end if
	if Request.Form("IsObject") = "1" then
		RsAddObj("IsObject") = True
	else
		RsAddObj("IsObject") = False
	end if
	if Request.Form("IsStyle") = "1" then
		RsAddObj("IsStyle") = True
	else
		RsAddObj("IsStyle") = False
	end if
	if Request.Form("IsDiv") = "1" then
		RsAddObj("IsDiv") = True
	else
		RsAddObj("IsDiv") = False
	end if
	if Request.Form("IsA") = "1" then
		RsAddObj("IsA") = True
	else
		RsAddObj("IsA") = False
	end if
	if Request.Form("Audit") = "1" then
		RsAddObj("Audit") = True
	else
		RsAddObj("Audit") = False
	end if
	if Request.Form("IsAutoCollect") <> "" then
		RsAddObj("IsAutoCollect") = True
		RsAddObj("CollectDate") = Clng(Request.Form("CollectDate"))
	else
		RsAddObj("IsAutoCollect") = False
	end if
	if Request.Form("TextTF") = "1" then
		RsAddObj("TextTF") = True
	else
		RsAddObj("TextTF") = False
	end If
	If Request.Form("IsReverse") = "1" Then
		RsAddObj("IsReverse") = "1"
	Else
		RsAddObj("IsReverse") = "0"
	End If
	if Request.Form("SaveRemotePic") = "1" then
		RsAddObj("SaveRemotePic") = True
	else
		RsAddObj("SaveRemotePic") = False
	end if
	if Request.Form("Islock") <> "" then
		RsAddObj("Islock") = True
	else
		RsAddObj("Islock") = False
	end if
	If Request.Form("IsAutoPicNews") <> "" Then
		RsAddObj("IsAutoPicNews") = 1
	Else
		RsAddObj("IsAutoPicNews") = 0
	End If	
	RsAddObj.update
	RsAddObj.close
	Set RsAddObj = Nothing
	Response.Redirect("Site.asp")
	Response.End
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自动新闻采集—站点设置</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<form name="Form" method="post" action="" id="Form">
      <table width="98%" height="20" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback" >
            <td width="30" align="center" alt="保存" style="cursor:hand" onClick="document.Form.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
			<td width="35" align="center" alt="后退" style="cursor:hand" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp; <input name="vs" type="hidden" id="vs2" value="add"> <input name="SiteID" type="hidden" id="SiteID2" value="<% = SiteID %>"> 
              <input name="Result" type="hidden" id="Result2" value="Edit"></td>
        </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr class="hback" > 
      <td width="15%" height="26"> <div align="center">采集站点名称</div></td>
      <td> <input name="SiteName" style="width:100%;" type="text" id="SiteName" value="<%=RsEditObj("sitename")%>"> 
        <div align="right"> </div></td>
    </tr>
    <tr class="hback" > 
      <td height="26"> <div align="center">采集对象页</div></td>
      <td><input name="objURL" style="width:100%;" type="text" id="objURL" value="<%=RsEditObj("objURL")%>" size="50"></td>
    </tr>
	<tr class="hback" > 
		<td height="26"> <div align="center">采集站点分类</div></td>
      <td><select name="SiteFolder" style="width:100%;" id="SiteFolder">
		<option value="0">根栏目</option>
          <% = SiteFolderList %>
        </select></td>
    </tr>
    <tr class="hback" > 
      <td height="26"><div align="center">采集参数</div></td>
      <td>锁定 
        <input name="islock" type="checkbox" id="islock" value="1" <%if RsEditObj("islock")=true then response.Write("checked")%>>
        保存远程图片 
        <input type="checkbox" name="SaveRemotePic" value="1" <%if RsEditObj("SaveRemotePic")=true then response.Write("checked")%>>
        新闻是否已经审核 
        <input type="checkbox" name="Audit" value="1" <%if RsEditObj("Audit")=true then response.Write("checked")%>>
		是否倒序采集 
        <input name="IsReverse" type="checkbox" id="IsReverse" value="1" <%if RsEditObj("IsReverse")="1" then response.Write("checked")%>>
		<!-- 2007-02-25 Edit By Ken -->
		内容中包含图片时设置为图片新闻
		<input name="IsAutoPicNews" type="checkbox" id="IsAutoPicNews" value="1" <%if RsEditObj("IsAutoPicNews")="1" then response.Write("checked")%>>
		<!-- End -->
	</td>
    </tr>
    <tr class="hback" > 
      <td height="26"><div align="center">过滤选项</div></td>
      <td>HTML <input type="checkbox" name="TextTF" value="1" <% if RsEditObj("TextTF") = True then Response.Write("checked")%>>
        STYLE <input type="checkbox" name="IsStyle" value="1" <% if RsEditObj("IsStyle") = True then Response.Write("checked")%>>
        DIV<input type="checkbox" name="IsDiv" value="1" <% if RsEditObj("IsDiv") = True then Response.Write("checked")%>>
        A<input type="checkbox" name="IsA" value="1" <% if RsEditObj("IsA") = True then Response.Write("checked")%>>
        CLASS<input type="checkbox" name="IsClass" value="1" <% if RsEditObj("IsClass") = True then Response.Write("checked")%>>
        FONT<input type="checkbox" name="IsFont" value="1" <% if RsEditObj("IsFont") = True then Response.Write("checked")%>>
        SPAN<input type="checkbox" name="IsSpan" value="1" <% if RsEditObj("IsSpan") = True then Response.Write("checked")%>>
        OBJECT<input type="checkbox" name="IsObject" value="1" <% if RsEditObj("IsObject") = True then Response.Write("checked")%>>
        IFRAME<input type="checkbox" name="IsIFrame" value="1" <% if RsEditObj("IsIFrame") = True then Response.Write("checked")%>>
        SCRIPT<input type="checkbox" name="IsScript" value="1" <% if RsEditObj("IsScript") = True then Response.Write("checked")%>>
      </td>
    </tr>
    <tr class="hback" > 
      <td height="36" colspan="2">
<div align="center"></div>
        <div align="center">
          <input onClick="ChangeCutPara(0);" <% if RsEditObj("OtherType") = 0 then Response.Write("checked") %> name="OtherType" type="radio" value="0">
          不分页 
          <input type="radio" onClick="ChangeCutPara(1);" name="OtherType" <% if RsEditObj("OtherType") = 1 then Response.Write("checked") %> value="1">
          标记分页设置 
          <input type="radio" onClick="ChangeCutPara(2);" <% if RsEditObj("OtherType") = 2 then Response.Write("checked") %> name="OtherType" value="2">
          索引分页设置 
          <input type="radio" onClick="ChangeCutPara(3);" <% if RsEditObj("OtherType") = 3 then Response.Write("checked") %> name="OtherType" value="3">
          手工分页设置
		  <input type="radio" onClick="ChangeCutPara(4);" <% if RsEditObj("OtherType") = 4 then Response.Write("checked") %> name="OtherType" value="4">
          <b>列表内容范围设置</b></div></td>
    </tr>
    <tr class="hback"  id="TagCutPage" style="display:<% if RsEditObj("OtherType") <> 1 then Response.Write("none") %>;"> 
      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr class="hback" > 
            <td width="15%"> 
              <div align="center">其他页面</div></td>
            <td>
			&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form.OtherPageSetting.rows>2)document.Form.OtherPageSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form.OtherPageSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
			&nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.Form.OtherPageSetting);" onClick="addTag('[其他页面]')" style="CURSOR: hand"><b>[其他页面]</b></font>
			&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.OtherPageSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font>
			<br>
			<textarea  ondblclick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="OtherPageSetting" id="OtherPageSetting" rows="4" style="width:100%;"><%=RsEditObj("OtherPageHeadSetting")%>[其他页面]<%=RsEditObj("OtherPageFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback"  id="IndexCutPage" style="display:<% if RsEditObj("OtherType") <> 2 then Response.Write("none") %>;"> 
      <td colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr class="hback" > 
            <td width="15%"> 
              <div align="center">索引规则 </div></td>
            <td>
			&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form.IndexRule.rows>2)document.Form.IndexRule.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form.IndexRule.rows+=1" style='cursor:hand'><b>扩大</b></span><br>
			<textarea name="IndexRule" rows="3" id="IndexRule" style="width:100%;"><% = RsEditObj("IndexRule") %></textarea></td>
          </tr>
          <tr class="hback" > 
            <td height="26"> <div align="center">页码</div></td>
            <td>页码开始： 
              <input name="StartPageNum" type="text" id="StartPageNum" size="3" maxlength="8" value="<% = RsEditObj("StartPageNum") %>">
              页码结束 
              <input name="EndPageNum" type="text" id="EndPageNum" size="3" maxlength="8" value="<% = RsEditObj("EndPageNum") %>">&nbsp&nbsp例:在索引规则中写http://.../index_^$^.htm，其中^$^代表设定的页码</td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback"  id="HandCutPage" style="display:<% if RsEditObj("OtherType") <> 3 then Response.Write("none") %>;"> 
      <td width="10%"> <div align="center">分页内容</div></td>
      <td height="26">	  &nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form.HandPageContent.rows>2)document.Form.HandPageContent.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form.HandPageContent.rows+=1" style='cursor:hand'><b>扩大</b></span>
			<textarea  name="HandPageContent" rows="6" id="HandPageContent" style="width:100%;"><% = RsEditObj("HandPageContent") %></textarea></tr>
    <tr class="hback"   id="ListContent" style="display:none"> 
      <td colspan="2">
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr class="hback" > 
            <td width="15%"> 
	  <div align="center">列表内容</div></td>
      <td>	&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form.ListSetting.rows>2)document.Form.ListSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form.ListSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
	  &nbsp;&nbsp;可用标签:<font onmouseover="getActiveText(document.Form.ListSetting);" onClick="addTag('[列表内容]')" style="CURSOR: hand"><b>[列表内容]</b></font>
	  &nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.ListSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><br>
	   <textarea  ondblclick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="ListSetting" cols="50" rows="3" id="ListSetting" style="width:100%;"><%=RsEditObj("ListHeadSetting")%>[列表内容]<%=RsEditObj("ListFootSetting")%></textarea>
	   </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr class="hback" > 
      <td> <div align="center">列表URL<font color="#ff0000">*</font></div></td>
      <td>	&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form.LinkSetting.rows>2)document.Form.LinkSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form.LinkSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
			&nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.Form.LinkSetting);" onClick="addTag('[列表URL]')" style="CURSOR: hand"><b>[列表URL]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.LinkSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><br>
	  <textarea   ondblclick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)"  name="LinkSetting" cols="50" rows="3" id="textarea2" style="width:100%;"><%=RsEditObj("LinkHeadSetting")%>[列表URL]<%=RsEditObj("LinkFootSetting")%></textarea></td>
    </tr>
    <tr class="hback" > 
      <td> <div align="center">新闻标题<font color="#ff0000">*</font></div></td>
      <td>	&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form.PageTitleSetting.rows>2)document.Form.PageTitleSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form.PageTitleSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
			&nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.Form.PageTitleSetting);" onClick="addTag('[新闻标题]')" style="CURSOR: hand"><b>[新闻标题]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.PageTitleSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><br>
	  <textarea  ondblclick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)"  name="PageTitleSetting" cols="50" rows="3" id="textarea6" style="width:100%;"><%=RsEditObj("PageTitleHeadSetting")%>[新闻标题]<%=RsEditObj("PageTitleFootSetting")%></textarea></td>
    </tr>
    <tr class="hback" > 
      <td> <div align="center">新闻内容<font color="#ff0000">*</font></div></td>
      <td>	&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form.PagebodySetting.rows>2)document.Form.PagebodySetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form.PagebodySetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
	  &nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.Form.PagebodySetting);" onClick="addTag('[新闻内容]')" style="CURSOR: hand"><b>[新闻内容]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.PagebodySetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><br>
	   <textarea  ondblclick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)"  name="PagebodySetting" cols="50" rows="3" id="textarea8" style="width:100%;"><%=RsEditObj("PagebodyHeadSetting")%>[新闻内容]<%=RsEditObj("PagebodyFootSetting")%></textarea></td>
    </tr>
    <tr class="hback" > 
      <td height="26" colspan="4"> <div align="center">
          <input name="OtherSetType" type="radio" onClick="ChangeSetOption(0);" value="0" checked>
          设置作者 
          <input type="radio" name="OtherSetType" onClick="ChangeSetOption(1);" value="1">
          设置来源 
          <input type="radio" name="OtherSetType" onClick="ChangeSetOption(2);" value="2">
          设置时间 
		  </div></td>
    </tr>
    <tr class="hback"  id="SetAuthor" style="display:;"> 
      <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr class="hback" >
            <td height="26"> 
              <div align="center">手动设置</div></td>
            <td colspan="3"><input style="width:100%;" type="text" name="HandSetAuthor" value="<% = RsEditObj("HandSetAuthor") %>"></td>
          </tr>
          <tr class="hback" > 
            <td width="15%"> 
              <div align="center">作者</div></td>
            <td colspan="3">	&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form.AuthorSetting.rows>2)document.Form.AuthorSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form.AuthorSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
			&nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.Form.AuthorSetting);" onClick="addTag('[作者]')" style="CURSOR: hand"><b>[作者]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.AuthorSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><br>
			<textarea  ondblclick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)"  name="AuthorSetting" cols="50" rows="3" id="textarea9" style="width:100%;"><%=RsEditObj("AuthorHeadSetting")%>[作者]<%=RsEditObj("AuthorFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback"  id="SetSource" style="display:none;"> 
      <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr class="hback" >
            <td height="26">
<div align="center">手动设置</div></td>
            <td colspan="3"><input style="width:100%;" type="text" name="HandSetSource" value="<% = RsEditObj("HandSetSource") %>"></td>
          </tr>
		  <tr class="hback" > 
            <td width="15%"> 
              <div align="center">来源</div></td>
            <td colspan="3">	&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form.SourceSetting.rows>2)document.Form.SourceSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form.SourceSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
			&nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.Form.SourceSetting);" onClick="addTag('[来源]')" style="CURSOR: hand"><b>[来源]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.SourceSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><br>
			 <textarea   ondblclick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="SourceSetting" cols="50" rows="3" id="textarea9a" style="width:100%;"><%=RsEditObj("SourceHeadSetting")%>[来源]<%=RsEditObj("SourceFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback"  id="SetAddTime" style="display:none;"> 
      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr class="hback" >
            <td height="26">
<div align="center">手动设置</div></td>
            <td colspan="3"><input style="width:100%;" type="text" name="HandSetAddDate" value="<% = RsEditObj("HandSetAddDate") %>"></td>
          </tr>
		  <tr class="hback" > 
            <td width="15%"> 
              <div align="center">加入时间</div></td>
            <td>	&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form.AddDateSetting.rows>2)document.Form.AddDateSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form.AddDateSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
			&nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.Form.AddDateSetting);" onClick="addTag('[加入时间]')" style="CURSOR: hand"><b>[加入时间]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.AddDateSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><br>
			 <textarea  ondblclick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)"  name="AddDateSetting" cols="50" rows="3" id="textarea9" style="width:100%;"><%=RsEditObj("AddDateHeadSetting")%>[加入时间]<%=RsEditObj("AddDateFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback" > 
      <td height="26" colspan="4"> <div align="center">
          <input name="OtherNewsType" <% if RsEditObj("OtherNewsType") = 0 then Response.Write("checked") %> type="radio" onClick="ChangeNewsSetOption(0);" value="0" checked>
          不设置新闻分页 
          <input type="radio" <% if RsEditObj("OtherNewsType") = 1 then Response.Write("checked") %> name="OtherNewsType" onClick="ChangeNewsSetOption(1);" value="1">
          标记分页
          <input type="radio" <% if RsEditObj("OtherNewsType") = 2 then Response.Write("checked") %> name="OtherNewsType" onClick="ChangeNewsSetOption(2);" value="2">
          页码分页
          <!--input type="radio" <% if RsEditObj("OtherNewsType") = 3 then Response.Write("checked") %> name="OtherNewsType" onClick="ChangeNewsSetOption(3);" value="3">
          新闻手动分页-->
		  </div></td>
    </tr>
    <tr class="hback" id="SetCutPage" style="display:none;"> 
      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr> 
            <td width="15%"><div align="center">分页新闻<br>(下一页)</div></td>
      <td width="85%">&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form1.OtherNewsPageSetting.rows>2)document.Form1.OtherNewsPageSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form1.OtherNewsPageSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>&nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.form1.OtherNewsPageSetting);" onClick="addTag('[分页新闻]')" style="CURSOR: hand"><b>[分页新闻]</b></font>
              <table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="5"></td>
                </tr>
              </table>
            <textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="OtherNewsPageSetting" cols="50" rows="3" id="textarea5" style="width:100%;"><%=RsEditObj("OtherNewsPageHeadSetting")%>[分页新闻]<%=RsEditObj("OtherNewsPageFootSetting")%></textarea><br /><span style="color:red;">例:<% = Server.HTMLEncode("<a href=") %>"[分页新闻]"<% = Server.HTMLEncode(">") %>下一页 要求 下一页 必须为整个页面中唯一字符</span></td>
    </tr>
        </table></td>
    </tr>
    <tr class="hback"  id="SetIndexCutPage" style="display:none;"> 
      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr class="hback" > 
            <td width="15%"> 
            <div align="center">分页规则</div></td>
            <td width="85%" >
			&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form.IndexRule.rows>2)document.Form.IndexRule.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form.IndexRule.rows+=1" style='cursor:hand'><b>扩大</b></span>
			 &nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.form1.OtherNewsPageIndexSetting);" onClick="addTag('[分页新闻]')" style="CURSOR: hand"><b>[分页新闻]</b></font> &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;<font onmouseover="getActiveText(document.form1.OtherNewsPageIndexSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><br>
			<textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="OtherNewsPageIndexSetting" cols="50" rows="3" id="OtherNewsPageIndexSetting" style="width:100%;"><%=RsEditObj("OtherNewsPageIndexSetting")%></textarea><br /><span style="color:red;">例:<% = Server.HTMLEncode("&nbsp;</font><a href=") %>"[分页新闻]"<% = Server.HTMLEncode(">") %>[变量]<% = Server.HTMLEncode("</a>") %>要求 [分页新闻] 前字符串必须为整个页面中唯一代码</span></td>
          </tr>
        </table></td>
    </tr>
	<!--tr class="hback"  id="SetHandCutPage" style="display:none;"> 
      <td width="10%"> <div align="center">分页内容</div></td>
      <td height="26">&nbsp;&nbsp;输入区域：<span onClick="if(document.Form.HandPageContent.rows>2)document.Form.HandPageContent.rows-=1" style='cursor:hand'><b>缩小</b></span><span onClick="document.Form.HandPageContent.rows+=1" style='cursor:hand'><b>扩大</b></span><textarea  name="OtherNewsPageIndexSettingHandPageContent" rows="6" id="OtherNewsPageIndexSettingHandPageContent" style="width:100%;"><% = RsEditObj("OtherNewsPageIndexSettingHandPageContent") %></textarea>
	  </td>
	</tr-->	
</table>
</form>
<p><br><p><p>
</body>
</html>
<%
Set Conn = Nothing
Set CollectConn = Nothing
Set RsEditObj = Nothing
%>
<script language="JavaScript">
function ChangeCutPara(Flag)
{
	switch (Flag)
	{
		case 0 :
			document.all.TagCutPage.style.display='none';
			document.all.IndexCutPage.style.display='none';
			document.all.HandCutPage.style.display='none';
			document.all.ListContent.style.display='none';
			break;
		case 1 :
			document.all.TagCutPage.style.display='';
			document.all.IndexCutPage.style.display='none';
			document.all.HandCutPage.style.display='none';
			document.all.ListContent.style.display='none';
			break;
		case 2 :
			document.all.TagCutPage.style.display='none';
			document.all.IndexCutPage.style.display='';
			document.all.HandCutPage.style.display='none';
			document.all.ListContent.style.display='none';
			break;
		case 3 :
			document.all.TagCutPage.style.display='none';
			document.all.IndexCutPage.style.display='none';
			document.all.HandCutPage.style.display='';
			document.all.ListContent.style.display='none';
			break;
		case 4 :
			document.all.TagCutPage.style.display='none';
			document.all.IndexCutPage.style.display='none';
			document.all.HandCutPage.style.display='none';
			document.all.ListContent.style.display='';
			break;		
		default :
			document.all.TagCutPage.style.display='none';
			document.all.IndexCutPage.style.display='none';
			document.all.HandCutPage.style.display='none';
			document.all.ListContent.style.display='none';
			break;
	}
}
function ChangeSetOption(Flag)
{
	switch (Flag)
	{
		case 0 :
			document.all.SetAuthor.style.display='';
			document.all.SetSource.style.display='none';
			document.all.SetAddTime.style.display='none';
			break;
		case 1 :
			document.all.SetAuthor.style.display='none';
			document.all.SetSource.style.display='';
			document.all.SetAddTime.style.display='none';
			break;
		case 2 :
			document.all.SetAuthor.style.display='none';
			document.all.SetSource.style.display='none';
			document.all.SetAddTime.style.display='';
			break;
		case 999 :
			document.all.SetAuthor.style.display='none';
			document.all.SetSource.style.display='none';
			document.all.SetAddTime.style.display='none';
			break;
		default :
			document.all.SetAuthor.style.display='none';
			document.all.SetSource.style.display='none';
			document.all.SetAddTime.style.display='none';
			break;
	}
}

function ChangeNewsSetOption(f_Flag)
{
	switch (f_Flag)
	{
		case 0 :
			document.all.SetCutPage.style.display='none';
			document.all.SetIndexCutPage.style.display='none';
			document.all.SetHandCutPage.style.display='none';
			break;
		case 1 :
			document.all.SetCutPage.style.display='';
			document.all.SetIndexCutPage.style.display='none';
			document.all.SetHandCutPage.style.display='none';
			break;
		case 2 :
			document.all.SetCutPage.style.display='none';
			document.all.SetIndexCutPage.style.display='';
			document.all.SetHandCutPage.style.display='none';
			break;
		case 3 :
			document.all.SetCutPage.style.display='none';
			document.all.SetIndexCutPage.style.display='none';
			document.all.SetHandCutPage.style.display='';
			break;
		default :
			document.all.SetCutPage.style.display='none';
			document.all.SetIndexCutPage.style.display='none';
			document.all.SetHandCutPage.style.display='none';
			break;
	}
}		  

currObj = "uuuu";
function getActiveText(obj)
{
	obj.focus();
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
//alert("ok");
	if (obj_ta.isTextEdit)
	{
	//alert("nooooo");
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
	else
		return false;
}	
-->
</script>






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
Dim LinkHeadSetting,LinkFootSetting
Dim ObjUrl,ListHeadSetting,ListFootSetting,NewsLinkStr
Dim HandSetAuthor,HandSetSource,HandSetAddDate,WebCharset
Set RsEditObj = Server.CreateObject(G_FS_RS)
SiteID = Request("SiteID")
if SiteID <> "" then
	EditSql="Select * from FS_Site where ID=" & SiteID
	RsEditObj.Open EditSql,CollectConn,1,3
	if RsEditObj.Eof then
		Response.write("没有修改的站点")
	else
		ObjUrl = RsEditObj("ObjUrl")
		ListHeadSetting = RsEditObj("ListHeadSetting")
		ListFootSetting = RsEditObj("ListFootSetting")
		HandSetAuthor = RsEditObj("HandSetAuthor")
		HandSetSource = RsEditObj("HandSetSource")
		HandSetAddDate = RsEditObj("HandSetAddDate")
		WebCharset = RsEditObj("WebCharset")
	end if
else
	Response.write("没有修改的站点")
end if
Dim ListSetting
If InStr(Request.Form("LinkSetting"),"[列表URL]") = 0 Then
	Response.Write "<script>alert('列表URL没有设置或设置不正确！');history.back();</script>"
	Response.End 
End if
ListSetting = Split(Request.Form("LinkSetting"),"[列表URL]",-1,1)
LinkHeadSetting = ListSetting(0)
LinkFootSetting = ListSetting(1)

if Request.Form("Result") = "Edit" then
    Dim RsAddObj,sql
	Set RsAddObj = Server.CreateObject(G_FS_RS)
	Sql = "select * from FS_Site where id=" & Request.Form("SiteID")
	RsAddObj.Open Sql,CollectConn,1,3
	RsAddObj("LinkHeadSetting") = LinkHeadSetting
	RsAddObj("LinkFootSetting") = LinkFootSetting
	RsAddObj.update
	RsAddObj.close
	Set RsAddObj = Nothing
end if

Dim ResponseAllStr,NewsListStr
ResponseAllStr = GetPageContent(ObjURL,WebCharset)
NewsListStr = GetOtherContent(ResponseAllStr,ListHeadSetting,ListFootSetting)
NewsLinkStr = FormatUrl(GetOtherContent(NewsListStr,LinkHeadSetting,LinkFootSetting),ObjUrl)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自动新闻采集—站点设置</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<form name="form1" method="post" action="SiteFiveStep.asp" id="Form1">
<table width="100%" border="0" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr  class="hback">
            <td width="50" style="cursor:hand" align="center" alt="第四步" onClick="window.location.href='javascript:history.go(-1)';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">上一步</td>
			<td width=2 class="Gray">|</td>
            <td width="50" style="cursor:hand" align="center" alt="第五步" onClick="document.form1.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">下一步</td>
			<td width=2 class="Gray">|</td>
		    <td width="35" style="cursor:hand" align="center" alt="后退" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp; <input name="SiteID" type="hidden" id="SiteID2" value="<% = SiteID %>"> 
              <input name="Result" type="hidden" id="Result2" value="Edit"> <input type="hidden" name="NewsLinkStr" value="<% = NewsLinkStr %>"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr class="hback"> 
      <td width="20%"> <div align="center">标题</div></td>
      <td>	&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form1.PageTitleSetting.rows>2)document.Form1.PageTitleSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form1.PageTitleSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
	  &nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.form1.PageTitleSetting);" onClick="addTag('[标题]')" style="CURSOR: hand"><b>[标题]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.PageTitleSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font>
        <table width="95%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="5"></td>
          </tr>
        </table>
        <textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="PageTitleSetting" cols="50" rows="3" id="textarea4" style="width:100%;"><%=RsEditObj("PageTitleHeadSetting")%>[标题]<%=RsEditObj("PageTitleFootSetting")%></textarea></td>
    </tr>
    <tr class="hback"> 
      <td> <div align="center">内容</div></td>
      <td> &nbsp;&nbsp;输入区域： <span onClick="if(document.Form1.PagebodySetting.rows>2)document.Form1.PagebodySetting.rows-=1" style='cursor:hand'><b>缩小</b></span> 
        <span onClick="document.Form1.PagebodySetting.rows+=1" style='cursor:hand'><b>扩大</b></span> 
        &nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.form1.PagebodySetting);" onClick="addTag('[内容]')" style="CURSOR: hand"><b>[内容]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.PagebodySetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font>
        <textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onChange="getActiveText(this)" name="PagebodySetting" cols="50" rows="6" id="textarea" style="width:100%;"><%=RsEditObj("PagebodyHeadSetting")%>[内容]<%=RsEditObj("PagebodyFootSetting")%></textarea></td>
    </tr>
    <tr class="hback"> 
      <td height="26" colspan="4"> <div align="left"> 　　　　　　　　　　　　　　　　　
<input name="OtherSetType" type="radio" onClick="ChangeSetOption(0);" value="0" checked>
          设置作者 
          <input type="radio" name="OtherSetType" onClick="ChangeSetOption(1);" value="1">
          设置来源 
          <input type="radio" name="OtherSetType" onClick="ChangeSetOption(2);" value="2">
          设置时间 
        </div></td>
    </tr>
    <tr class="hback" id="SetAuthor" style="display:;"> 
      <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td height="26">
<div align="center">手动设置</div></td>
            <td colspan="3"><input style="width:100%;" type="text" name="HandSetAuthor" value="<% = HandSetAuthor %>"></td>
          </tr>
          <tr> 
            <td width="20%"> <div align="center">作者</div></td>
            <td colspan="3">&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form1.AuthorSetting.rows>2)document.Form1.AuthorSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form1.AuthorSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
			 &nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.form1.AuthorSetting);" onClick="addTag('[作者]')" style="CURSOR: hand"><b>[作者]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.AuthorSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font>
              <table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="5"></td>
                </tr>
              </table>
              <textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="AuthorSetting" cols="50" rows="3" id="textarea9" style="width:100%;"><%=RsEditObj("AuthorHeadSetting")%>[作者]<%=RsEditObj("AuthorFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback" id="SetSource" style="display:none;"> 
      <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td height="26">
<div align="center">手动设置</div></td>
            <td colspan="3"><input style="width:100%;" type="text" name="HandSetSource" value="<% = HandSetSource %>"></td>
          </tr>
		  <tr> 
            <td width="20%"> <div align="center">来源</div></td>
            <td colspan="3">&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form1.SourceSetting.rows>2)document.Form1.SourceSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form1.SourceSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
			&nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.form1.SourceSetting);" onClick="addTag('[来源]')" style="CURSOR: hand"><b>[来源]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.SourceSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font>
              <table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="5"></td>
                </tr>
              </table>
              <textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="SourceSetting" cols="50" rows="3" id="textarea9a" style="width:100%;"><%=RsEditObj("SourceHeadSetting")%>[来源]<%=RsEditObj("SourceFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback" id="SetAddTime" style="display:none;"> 
      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td height="26">
<div align="center">手动设置</div></td>
            <td colspan="3"><input style="width:100%;" type="text" name="HandSetAddDate" value="<% = HandSetAddDate %>"></td>
          </tr>
		  <tr> 
            <td width="20%"> <div align="center">加入时间</div></td>
            <td>&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form1.AddDateSetting.rows>2)document.Form1.AddDateSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form1.AddDateSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
			&nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.form1.AddDateSetting);" onClick="addTag('[加入时间]')" style="CURSOR: hand"><b>[加入时间]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.AddDateSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font>
              <table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="5"></td>
                </tr>
              </table>
              <textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="AddDateSetting" cols="50" rows="3" id="textarea9" style="width:100%;"><%=RsEditObj("AddDateHeadSetting")%>[加入时间]<%=RsEditObj("AddDateFootSetting")%></textarea></td>
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
            <td width="20%"> 
              <div align="center">分页新闻<br>(下一页)</div></td>
      <td>&nbsp;&nbsp;输入区域：
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
            <td width="20%"> 
            <div align="center">分页规则</div></td>
            <td width="80%" >
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
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr class="hback">
    <td colspan="2" height="28" class="xingmu"> <div align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;预览结果:<a href="<% = NewsLinkStr %>" target="_blank"> 
        <% = NewsLinkStr %>
        </a></div></td>
  </tr>
</table>
</body>
</html>
<%
Set RsEditObj = Nothing
Set CollectConn = Nothing
%>
<script language="JavaScript">
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
			//document.all.SetHandCutPage.style.display='none';
			break;
		case 1 :
			document.all.SetCutPage.style.display='';
			document.all.SetIndexCutPage.style.display='none';
			//document.all.SetHandCutPage.style.display='none';
			break;
		case 2 :
			document.all.SetCutPage.style.display='none';
			document.all.SetIndexCutPage.style.display='';
			//document.all.SetHandCutPage.style.display='none';
			break;
		//case 3 :
			//document.all.SetCutPage.style.display='none';
			//document.all.SetIndexCutPage.style.display='none';
			//document.all.SetHandCutPage.style.display='';
			//break;
		default :
			document.all.SetCutPage.style.display='none';
			document.all.SetIndexCutPage.style.display='none';
			//document.all.SetHandCutPage.style.display='none';
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






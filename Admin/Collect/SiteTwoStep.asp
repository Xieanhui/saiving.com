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
Dim RsEditObj,EditSql,SiteID
Dim ListHeadSetting,ListFootSetting,OtherPageFootSetting,OtherPageHeadSetting,OtherType,IndexRule,StartPageNum,EndPageNum,HandPageContent
Set RsEditObj = Server.CreateObject(G_FS_RS)
SiteID = Request("SiteID")
if SiteID <> "" then
	EditSql="Select * from FS_Site where ID=" & SiteID
	RsEditObj.Open EditSql,CollectConn,1,3
	if RsEditObj.Eof then
		Response.write"<script>alert(""没有修改的站点"");location.href=""javascript:history.back()"";</script>"
		Response.end
	else
		ListHeadSetting = RsEditObj("ListHeadSetting")
		ListFootSetting = RsEditObj("ListFootSetting")
		OtherPageFootSetting = RsEditObj("OtherPageFootSetting")
		OtherPageHeadSetting = RsEditObj("OtherPageHeadSetting")
		IndexRule = RsEditObj("IndexRule")
		StartPageNum = RsEditObj("StartPageNum")
		EndPageNum = RsEditObj("EndPageNum")
		HandPageContent = RsEditObj("HandPageContent")
		OtherType = RsEditObj("OtherType")
	end if
else
	Response.write"<script>alert(""没有修改的站点"");location.href=""javascript:history.back()"";</script>"
	Response.end
end if
Set RsEditObj = Nothing
if Request.Form("Result") = "Edit" then
    Dim RsAddObj,sql
    if Request.Form("SiteName")=""  or Request.Form("objURL")="" then
		Response.write"<script>alert(""请填写完整！"");location.href=""javascript:history.back()"";</script>"
		Response.end
	end if
	IF Trim(Request.Form("PicSavePath")) = "" then
		Response.write"<script>alert(""请选择图片保存路径！"");location.href=""javascript:history.back()"";</script>"
		Response.end
	End If	
	Set RsAddObj = Server.CreateObject(G_FS_RS)
	Sql = "select * from FS_Site where id=" & CintStr(Request.Form("SiteID"))
	RsAddObj.Open Sql,CollectConn,1,3
	RsAddObj("SiteName") = NoSqlHack(Request.Form("SiteName"))
	RsAddObj("objURL") = NoSqlHack(Request.Form("objURL"))
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
	if Request.Form("TextTF") = "1" then
		RsAddObj("TextTF") = True
	else
		RsAddObj("TextTF") = False
	end if
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
	'===2007-02-25 Edit By Ken =======
	RsAddObj("folder") = NoSqlHack(Request.Form("SiteFolder"))
	If Request.Form("IsAutoPicNews") <> "" then
		RsAddObj("IsAutoPicNews") = 1
	Else
		RsAddObj("IsAutoPicNews") = 0
	End If
	If Request.Form("ToClass") <> "" then
		RsAddObj("ToClassID") = NoSqlHack(Request.Form("ToClass"))
	Else
		RsAddObj("ToClassID") = "0"
	End If
	RsAddObj("NewsTemplets") = NoSqlHack(Request.Form("NewsTemp"))
	IF Request.Form("AutoCTF") = "no" Then
		RsAddObj("AutoCellectTime") = "no"
	ElseIf Request.Form("AutoCTF") = "day" Then
		RsAddObj("AutoCellectTime") = "day$$$" & NoSqlHack(Request.Form("TimeHour"))
	ElseIF Request.Form("AutoCTF") = "week" Then
		RsAddObj("AutoCellectTime") = "week$$$" & NoSqlHack(Request.Form("TimeWeek"))& "|" & NoSqlHack(Request.Form("TimeHour"))
	ElseIF Request.Form("AutoCTF") = "month" Then
		RsAddObj("AutoCellectTime") = "month$$$" & NoSqlHack(Request.Form("TimeMonth")) & "|" & NoSqlHack(Request.Form("TimeHour"))
	End IF
	if Trim(Request.Form("NewsNum")) = "" Or Not IsNumeric(Trim(Request.Form("NewsNum"))) Then
		RsAddObj("CellectNewNum") = 0
	Else
		If Cint(Trim(Request.Form("NewsNum"))) < 0 Then
			RsAddObj("CellectNewNum") = 0
		Else
			RsAddObj("CellectNewNum") = CintStr(Request.Form("NewsNum"))
		End IF	
	End If
	RsAddObj("WebCharset") = NoSqlHack(Request.Form("WebCharset"))
	RsAddObj("RulerID") = NoSqlHack(Request.Form("CS_SiteReKeyID"))
	RsAddObj("PicSavePath") = NoSqlHack(Request.Form("PicSavePath"))
	IF Request.Form("WaterPrint") = 1 Then
		RsAddObj("WaterPrintTF") = 1
	Else
		RsAddObj("WaterPrintTF") = 0
	End IF		
	'===End===========================	
	RsAddObj.update
	RsAddObj.close
	Set RsAddObj = Nothing
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
<form name="form1" method="post" action="SiteThreeStep.asp" id="Form1">
<table width="98%" height="20" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback" >
            <td width="50" style="cursor:hand" align="center" alt="第二步" onClick="window.location.href='javascript:history.go(-1)';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">上一步</td>
		    <td width="50" style="cursor:hand" align="center" alt="第三步" onClick="document.form1.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">下一步</td>
			  <td width="35" align="center" alt="后退" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut" style="cursor:hand">后退</td>
            <td>&nbsp; <input name="SiteID" type="hidden" id="SiteID2" value="<% = SiteID %>"> 
              <input name="Result" type="hidden" id="Result2" value="Edit"></td>
        </tr>
      </table>  
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
    <tr class="hback" > 
      <td width="10%"> 
        <div align="center">列表内容</div></td>
      <td>	&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form1.ListSetting.rows>2)document.Form1.ListSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form1.ListSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
	  &nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.form1.ListSetting);" onClick="addTag('[列表内容]')" style="CURSOR: hand"><b>[列表内容]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.ListSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><br>
	<textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="ListSetting" rows="10" id="ListSetting" style="width:100%;"><%=ListHeadSetting%>[列表内容]<%=ListFootSetting%></textarea></td>
    </tr>
    <tr class="hback" > 
      <td height="36" colspan="2">
<div align="center"></div>
        <div align="center">
          <input onClick="ChangeCutPara(0);" <% if OtherType = 0 then Response.Write("checked") %> name="OtherType" type="radio" value="0">
          不分页 
          <input type="radio" onClick="ChangeCutPara(1);" name="OtherType" <% if OtherType = 1 then Response.Write("checked") %> value="1">
          标记分页设置 
          <input type="radio" onClick="ChangeCutPara(2);" <% if OtherType = 2 then Response.Write("checked") %> name="OtherType" value="2">
          索引分页设置 
          <input type="radio" onClick="ChangeCutPara(3);" <% if OtherType = 3 then Response.Write("checked") %> name="OtherType" value="3">
          手工分页设置 </div></td>
    </tr>
    <tr class="hback"  id="TagCutPage" style="display:<% if OtherType <> 1 then Response.Write("none") %>;"> 
      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr class="hback" > 
            <td width="10%" bgcolor="#F5F5F5"> 
              <div align="center">其他页面</div></td>
            <td>&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form1.OtherPageSetting.rows>2)document.Form1.OtherPageSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form1.OtherPageSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
			&nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.form1.OtherPageSetting);" onClick="addTag('[其他页面]')" style="CURSOR: hand"><b>[其他页面]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.OtherPageSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font>
              <table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr class="hback" >
                  <td height="5"></td>
                </tr>
              </table>
              <textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="OtherPageSetting" rows="4" style="width:100%;"><%=OtherPageHeadSetting%>[其他页面]<%=OtherPageFootSetting%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback"  id="IndexCutPage" style="display:<% if OtherType <> 2 then Response.Write("none") %>;"> 
      <td colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr class="hback" > 
            <td width="10%" bgcolor="#F5F5F5"> 
              <div align="center">索引规则 </div></td>
            <td>&nbsp;&nbsp;输入区域： <span onClick="if(document.Form1.IndexRule.rows>2)document.Form1.IndexRule.rows-=1" style='cursor:hand'><b>缩小</b></span> 
              <span onClick="document.Form1.IndexRule.rows+=1" style='cursor:hand'><b>扩大</b></span> 
              <table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr class="hback" > 
                  <td height="5"></td>
                </tr>
              </table>
              <textarea name="IndexRule" rows="3" id="IndexRule" style="width:100%;"><% = IndexRule %></textarea></td>
          </tr>
          <tr class="hback" > 
            <td height="26" bgcolor="#F5F5F5"> 
              <div align="center">页码</div></td>
            <td>页码开始： 
              <input name="StartPageNum" type="text" id="StartPageNum" size="10" maxlength="3" value="<% = StartPageNum %>">
              页码结束 
              <input name="EndPageNum" type="text" id="EndPageNum" size="10" maxlength="3" value="<% = EndPageNum %>"></td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback"  id="HandCutPage" style="display:<% if OtherType <> 3 then Response.Write("none") %>;"> 
      <td width="10%" bgcolor="#F5F5F5"> 
        <div align="center">分页内容</div></td>
      <td height="26">&nbsp;&nbsp;输入区域： <span onClick="if(document.Form1.HandPageContent.rows>2)document.Form1.HandPageContent.rows-=1" style='cursor:hand'><b>缩小</b></span> 
        <span onClick="document.Form1.HandPageContent.rows+=1" style='cursor:hand'><b>扩大</b></span> 
        <table width="95%" border="0" cellspacing="0" cellpadding="0">
          <tr class="hback" > 
            <td height="5"></td>
          </tr>
        </table> 
        <textarea name="HandPageContent" rows="6" id="HandPageContent" style="width:100%;"><% = HandPageContent %></textarea></tr>
</table>
</form>
</body>
</html>
<%
Set CollectConn = Nothing
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
			break;
		case 1 :
			document.all.TagCutPage.style.display='';
			document.all.IndexCutPage.style.display='none';
			document.all.HandCutPage.style.display='none';
			break;
		case 2 :
			document.all.TagCutPage.style.display='none';
			document.all.IndexCutPage.style.display='';
			document.all.HandCutPage.style.display='none';
			break;
		case 3 :
			document.all.TagCutPage.style.display='none';
			document.all.IndexCutPage.style.display='none';
			document.all.HandCutPage.style.display='';
			break;
		default :
			document.all.TagCutPage.style.display='none';
			document.all.IndexCutPage.style.display='none';
			document.all.HandCutPage.style.display='none';
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






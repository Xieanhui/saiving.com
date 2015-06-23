<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<%
Dim Conn,CollectConn
MF_Default_Conn
MF_Collect_Conn
MF_Session_TF

Dim sRootDir,str_CurrPath,SaveRemotePicPath
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")

if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/" & G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if

SaveRemotePicPath = Replace(sRootDir &"/"&G_UP_FILES_DIR & "/" & G_SAVE_FILE_PATH,"//","/")

if not MF_Check_Pop_TF("CS001") then Err_Show
Dim Rs
if Request("Action") = "Del" then
	if Request("id") <> "" then
		CollectConn.Execute("delete from FS_Site where ID in (" & FormatIntArr(Replace(Request("id"),"***",",")) & ")")
	end If
	if Request("SiteFolderID") <> "" then
		CollectConn.Execute("delete from FS_SiteFolder where ID in (" & FormatIntArr(Replace(Request("SiteFolderID"),"***",",")) & ")")
	end if
	Response.Redirect("site.asp")
	Response.End
elseif Request("Action") = "Lock" then
	if Request("LockID") <> "" then
		CollectConn.Execute("Update FS_Site Set IsLock=1 where ID in (" & FormatIntArr(Replace(Request("LockID"),"***",",")) & ")")
		Response.Redirect("site.asp")
		Response.End
	end if
elseif Request("Action") = "UNLock" then
	if Request("LockID") <> "" then
		CollectConn.Execute("Update FS_Site Set IsLock=0 where ID in (" & FormatIntArr(Replace(Request("LockID"),"***",",")) & ")")
		Response.Redirect("site.asp")
		Response.End
	end if
end if
if Request.Form("vs")="add" then
    if Request.Form("SiteName")="" or Request.Form("objURL")="" then
		Response.write"<script>alert(""请填写完整！"");location.href=""javascript:history.back()"";</script>"
		Response.end
	end if
	IF Trim(Request.Form("PicSavePath")) = "" Then
		Response.write"<script>alert(""请选择图片保存路径！"");location.href=""javascript:history.back()"";</script>"
		Response.end
	End If	
    Dim Sql
	Set Rs = Server.CreateObject(G_FS_RS)
	Sql = "Select * from FS_Site where 1=0"
	Rs.Open Sql,CollectConn,1,3
	Rs.AddNew
	Rs("SiteName") = NoSqlHack(Request.Form("SiteName"))
	Rs("objURL") = NoSqlHack(Request.Form("objURL"))
	Rs("folder") = NoSqlHack(Request.Form("SiteFolder"))
	if Request.Form("IsIFrame") = "1" then
		Rs("IsIFrame") = True
	else
		Rs("IsIFrame") = False
	end if
	if Request.Form("IsReverse") = "1" then
		Rs("IsReverse") = 1
	else
		Rs("IsReverse") = 0
	end if
	if Request.Form("IsScript") = "1" then
		Rs("IsScript") = True
	else
		Rs("IsScript") = False
	end if
	if Request.Form("IsClass") = "1" then
		Rs("IsClass") = True
	else
		Rs("IsClass") = False
	end if
	if Request.Form("IsFont") = "1" then
		Rs("IsFont") = True
	else
		Rs("IsFont") = False
	end if
	if Request.Form("IsSpan") = "1" then
		Rs("IsSpan") = True
	else
		Rs("IsSpan") = False
	end if
	if Request.Form("IsObject") = "1" then
		Rs("IsObject") = True
	else
		Rs("IsObject") = False
	end if
	if Request.Form("IsStyle") = "1" then
		Rs("IsStyle") = True
	else
		Rs("IsStyle") = False
	end if
	if Request.Form("IsDiv") = "1" then
		Rs("IsDiv") = True
	else
		Rs("IsDiv") = False
	end if
	if Request.Form("IsA") = "1" then
		Rs("IsA") = True
	else
		Rs("IsA") = False
	end if
	if Request.Form("Audit") = "1" then
		Rs("Audit") = True
	else
		Rs("Audit") = False
	end if
	if Request.Form("TextTF") = "1" then
		Rs("TextTF") = True
	else
		Rs("TextTF") = False
	end if
	if Request.Form("SaveRemotePic") = "1" then
		Rs("SaveRemotePic") = True
	else
		Rs("SaveRemotePic") = False
	end if
	if Request.Form("Islock") <> "" then
		Rs("Islock") = True
	else
		Rs("Islock") = False
	end if
	'===2007-02-25 Edit By Ken =======
	If Request.Form("IsAutoPicNews") <> "" then
		Rs("IsAutoPicNews") = 1
	Else
		Rs("IsAutoPicNews") = 0
	End If
	If Request.Form("ToClass") <> "" then
		Rs("ToClassID") = NoSqlHack(Request.Form("ToClass"))
	Else
		Rs("ToClassID") = "0"
	End If
	Rs("NewsTemplets") = NoSqlHack(Request.Form("NewsTemp"))
	IF Request.Form("AutoCTF") = "no" Then
		Rs("AutoCellectTime") = "no"
	ElseIf Request.Form("AutoCTF") = "day" Then
		Rs("AutoCellectTime") = "day$$$" & NoSqlHack(Request.Form("TimeHour"))
	ElseIF Request.Form("AutoCTF") = "week" Then
		Rs("AutoCellectTime") = "week$$$" & NoSqlHack(Request.Form("TimeWeek")) & "|" & NoSqlHack(Request.Form("TimeHour"))
	ElseIF Request.Form("AutoCTF") = "month" Then
		Rs("AutoCellectTime") = "month$$$" & NoSqlHack(Request.Form("TimeMonth")) & "|" & NoSqlHack(Request.Form("TimeHour"))
	End IF
	if Trim(Request.Form("NewsNum")) = "" Or Not IsNumeric(Trim(Request.Form("NewsNum"))) Then
		Rs("CellectNewNum") = 0
	Else
		If Cint(Trim(Request.Form("NewsNum"))) < 0 Then
			Rs("CellectNewNum") = 0
		Else
			Rs("CellectNewNum") = CintStr(Request.Form("NewsNum"))
		End IF	
	End If
	Rs("WebCharset") = NoSqlHack(Request.Form("WebCharset"))
	Rs("RulerID") = NoSqlHack(Request.Form("CS_SiteReKeyID"))
	Rs("PicSavePath") = NoSqlHack(Request.Form("PicSavePath"))
	IF Request.Form("WaterPrint") = 1 Then
		Rs("WaterPrintTF") = 1
	Else
		Rs("WaterPrintTF") = 0
	End IF
	'===End===========================	
	Rs.UpDate
	Rs.Close
	Set Rs = Nothing
	Set CollectConn = Nothing
	Response.Write "<script>alert('操作成功');location.href='Site.asp';</script>"
	Response.End
elseif Request("vs")="addfolder" then
	Dim SiteFolder,SiteFolderDetail,SqlStr
	SiteFolder = Request.Form("SiteFolder")
	SiteFolderDetail = Request.Form("SiteFolderDetail")
	If SiteFolder = "" or SiteFolderDetail = "" Then
		Response.write"<script>alert(""请填写完整！"");location.href=""javascript:history.back()"";</script>"
		Response.end
	End If
	Set Rs = Server.CreateObject(G_FS_RS)
	if Request("SiteFolderID") <> "" then
		SqlStr = "Select * from FS_SiteFolder where ID=" & CintStr(Request("SiteFolderID"))
		Rs.Open SqlStr,CollectConn,1,3
	else
		SqlStr = "Select * from FS_SiteFolder where 1=0"
		Rs.Open SqlStr,CollectConn,1,3
		Rs.AddNew
	end if
	Rs("SiteFolder") = SiteFolder
	Rs("SiteFolderDetail") = SiteFolderDetail
	Rs.UpDate
	Rs.Close
	Set Rs = Nothing
	Set Conn = Nothing
	Set CollectConn = Nothing
	Response.Redirect("Site.asp")
	Response.end
elseif Request("vs")="Copy" then
	Dim SiteID,SiteFolderID,RsCopySourceObj,RsCopyObjectObj,FiledObj
	SiteID = Request("SiteID")
	SiteFolderID = Request("SiteFolderID")
	if SiteID <> "" then
		Set RsCopySourceObj = CollectConn.Execute("Select * from FS_Site where ID in (" & FormatIntArr(Replace(SiteID,"***",",")) & ")")
		do while Not RsCopySourceObj.Eof
			Set RsCopyObjectObj = Server.CreateObject(G_FS_RS)
			RsCopyObjectObj.Open "Select * from FS_Site where 1=0",CollectConn,3,3
			RsCopyObjectObj.AddNew
			For Each FiledObj In RsCopyObjectObj.Fields
				if LCase(FiledObj.name) <> "id" then
					RsCopyObjectObj(FiledObj.name) = RsCopySourceObj(FiledObj.name)
				end if
			Next
			RsCopyObjectObj.Update
			RsCopySourceObj.MoveNext
		Loop
		Set RsCopySourceObj = Nothing
		Set RsCopyObjectObj = Nothing
	end If
	if SiteFolderID <> "" then
		Set RsCopySourceObj = CollectConn.Execute("Select * from FS_SiteFolder where ID in (" & FormatIntArr(Replace(SiteFolderID,"***",",")) & ")")
		do while Not RsCopySourceObj.Eof
			Set RsCopyObjectObj = Server.CreateObject(G_FS_RS)
			RsCopyObjectObj.Open "Select * from FS_SiteFolder where 1=0",CollectConn,3,3
			RsCopyObjectObj.AddNew
			For Each FiledObj In RsCopyObjectObj.Fields
				if LCase(FiledObj.name) <> "id" then
					RsCopyObjectObj(FiledObj.name) = RsCopySourceObj(FiledObj.name)
				end if
			Next
			RsCopyObjectObj.Update
			RsCopySourceObj.MoveNext
		Loop
		Set RsCopySourceObj = Nothing
		Set RsCopyObjectObj = Nothing
	end if
	Set CollectConn = Nothing
	Response.Redirect("Site.asp")
	Response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改新闻</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<%
if Request("Action") = "Addsite" then
	Call Add()
ElseIf Request("Action") = "Addsitefolder" Then
	Call AddFolder()
ElseIf Request("Action") = "SubFolder" Then
	Call Main(Request("FolderID"))
Else
	Call Main("0")
end if
Sub Main(f_FolderID)
	if f_FolderID = "" then f_FolderID = "0"
	Session("SessionReturnValue") = ""
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
	<tr class="hback">
		<td height="26" colspan="5" valign="middle">
			<table width="100%" height="20" border="0" cellpadding="5" cellspacing="1">
				<tr>
					<td width=55 align="center" alt="添加采集栏目" style="cursor:hand" onClick="location='?Action=Addsitefolder';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新建栏目</td>
					<td width=2 class="Gray">|</td>
					<td width=55 align="center" alt="添加采集站点" style="cursor:hand" onClick="location='?Action=Addsite';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新建站点</td>
					<td width=2 class="Gray">|</td>
					<td width=35 align="center" alt="后退" style="cursor:hand" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
					<td>&nbsp;</td>
				</tr>
		  </table>
		</td>
	</tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu"> 
    <td width="50%" height="26" nowrap bgcolor="#FFFFFF" class="xingmu"> 
      <div align="center">名称</div></td>
    <td width="6%" height="26" nowrap bgcolor="#FFFFFF" class="xingmu"> 
      <div align="center">状态</div></td>
    <td width="10%" height="26" bgcolor="#FFFFFF" class="xingmu" nowrap> 
      <div align="center">采集对象页</div></td>
    <td height="26" nowrap class="xingmu">
<div align="center">操作</div></td>
  </tr>
  <%
	if f_FolderID = "0" then
		Dim RsSite,SiteSql,CheckInfo
		Dim RsSiteFolder
		Set RsSiteFolder = CollectConn.Execute("select * from FS_SiteFolder where 1=1 order by id DESC")
		Do While not RsSiteFolder.EOF
	%>
  <tr class="hback"> 
    <td height="26" nowrap><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/folder.gif" width="16" height="16"></td>
          <td nowrap><%= RsSiteFolder("SiteFolder")%></td>
        </tr>
      </table></td>
    <td nowrap><div align="center"> &nbsp; </div></td>
    <td nowrap><div align="center"> &nbsp; </div></td>
    <td nowrap><div align="center"><span style="cursor:hand;" onClick="if (confirm('确定要复制吗?')){location='Site.asp?vs=Copy&SiteFolderID=<% = RsSiteFolder("ID") %>';}">复制</span>&nbsp;&nbsp;<span style="cursor:hand;" onClick="if (confirm('确定要修改吗?')){location='?Action=Addsitefolder&SiteFolderID=<% = RsSiteFolder("ID") %>';}">修改</span>&nbsp;&nbsp;<span style="cursor:hand;" onClick="if (confirm('确定要删除吗?')){location='?action=Del&SiteFolderID=<% = RsSiteFolder("ID") %>';}">删除</span>&nbsp;&nbsp;<span style="cursor:hand;" onClick="location='site.asp?Action=SubFolder&FolderID=<% =RsSiteFolder("ID") %>'">进入</span></div></td>
  </tr>
  <%
			RsSiteFolder.MoveNext
		Loop
		Set RsSiteFolder = Nothing
	else
%>
  <tr class="hback"> 
    <td height="26" colspan="4" nowrap><img src="images/folder.gif" width="16" height="16"> 
      <a href="Site.asp">返回上一级</a> </td>
  </tr>
<%	
	end if
	
	Dim IsCollect,RsTempObj,CollectPromptInfo
	Set RsSite = Server.CreateObject(G_FS_RS)
	SiteSql="Select * from FS_Site where folder=" & NoSqlHack(f_FolderID) & " order by id desc"
		RsSite.Open SiteSql,CollectConn,1,1
	Do While not RsSite.eof
		if  RsSite("LinkHeadSetting") <> "" And  RsSite("LinkFootSetting") <> "" And RsSite("PagebodyHeadSetting") <> "" And  RsSite("PagebodyFootSetting") <> "" And  RsSite("PageTitleHeadSetting") <> "" And  RsSite("PageTitleFootSetting") <> "" then
			if RsSite("IsLock") = True then
				IsCollect = False
				CollectPromptInfo = "站点已经被锁定,不能采集"
			else
				IsCollect = True
				CollectPromptInfo = "可以采集,请检查是否设置正确，否则不能进行采集"
			end if
		else
			IsCollect = False
			CollectPromptInfo = "不能采集,请把匹配规则设置完整"
		end if
	%>
  <tr title="<% = CollectPromptInfo %>" class="hback"> 
    <td height="26" nowrap><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/SiteSet.gif" width="23" height="22"></td>
          <td nowrap><% = RsSite("SiteName") %></td>
        </tr>
      </table></td>
    <td nowrap><div align="center"> 
        <%
		if RsSite("IsLock") = True then
			Response.Write("锁定")
		ElseIf IsCollect = False Then
			Response.Write("无效")
		else
			Response.Write("有效")
		end if
		%>
      </div></td>
    <td nowrap><div align="center"><a href="<% = RsSite("objURL") %>" target="_blank"><img src="Images/objpage.gif" alt="点击访问" width="20" height="20" border="0"></a></div></td>
    <td nowrap><div align="center"><span style="cursor:hand;" onClick="if (confirm('确定要复制吗?')){location='Site.asp?vs=Copy&SiteID=<% = RsSite("ID") %>';}">复制</span>&nbsp;&nbsp;<span style="cursor:hand;" onClick="if (confirm('确定要删除吗?')){location='?action=Del&Id=<% = RsSite("ID") %>';}">删除</span>&nbsp;&nbsp;<span style="cursor:hand;" onClick="if (confirm('确定要修改吗？')){location='SitemodifyOne.asp?SiteID=<% = RsSite("ID") %>';}">向导</span><% if IsCollect = true then %>&nbsp;&nbsp;<span onClick="StartOneSiteCollect('<% = RsSite("Id") %>');" style="cursor:hand;">采集</span><% end if %></div></td>
  </tr>
  <%
		RsSite.MoveNext
	loop
%>
</table>
<%
	RsSite.close
	Set RsSite = Nothing
end Sub

Sub Add()
%>
<form name="AddSiteForm" id="AddSiteForm" method="post" action="">
		<table width="98%" height="20" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
			<tr>
				<td width=35 align="center" alt="保存"  style="cursor:hand" onClick="document.AddSiteForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
				<td width=35 align="center" alt="后退"  style="cursor:hand" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
				<td>&nbsp;
					<input name="vs" type="hidden" id="vs2" value="add">				</td>
			</tr>
  </table>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr class="hback">
			<td width="100" height="26"><div align="right">采集站点名称：</div></td>
			<td><input name="SiteName" style="width:100%;" type="text" id="SiteName2"></td>
		</tr>
		<tr class="hback">
			<td height="26"><div align="right">采集站点分类：</div></td>
			<td><select name="SiteFolder" style="width:100%;" id="SiteFolder">
					<option value="0">根栏目</option>
					<% = FolderList() %>
				</select></td>
		</tr>
		<tr class="hback">
			<td height="26"><div align="right">采集对象页：</div></td>
			<td><input style="width:100%;" name="objURL" type="text" id="objURL" value="http://"></td>
		</tr>
		<tr class="hback">
			<td height="26"><div align="right">选择编码：</div></td>
			<td><select name="WebCharset" style="width:100%;" id="WebCharset">
					<option value="GB2312" selected="selected">GB2312</option>
					<option value="UTF-8">UTF-8</option>
				</select>
			</td>
		</tr>
		<tr class="hback">
			<td height="26"><div align="right">采集参数：</div></td>
			<td>锁定
				<input name="islock" type="checkbox" id="islock" value="1">
				保存远程图片
				<input type="checkbox" name="SaveRemotePic" value="1" checked>
				新闻是否已经审核
				<input name="Audit" type="checkbox" value="1">
				是否倒序采集
				<input name="IsReverse" type="checkbox" id="IsReverse" value="1">
				<!-- 2007-02-25 Edit By Ken -->
				内容中包含图片时设置为图片新闻
				<input name="IsAutoPicNews" type="checkbox" id="IsAutoPicNews" value="1" checked>
				<!-- End -->
				</td>
		</tr>
		<tr class="hback">
			<td height="26"><div align="right">过滤选项：</div></td>
			<td>HTML
				<input type="checkbox" name="TextTF" value="1" checked>
				STYLE
				<input type="checkbox" name="IsStyle" value="1" checked>
				DIV
				<input type="checkbox" name="IsDiv" value="1">
				A
				<input type="checkbox" name="IsA" value="1" checked>
				CLASS
				<input type="checkbox" name="IsClass" value="1">
				FONT
				<input type="checkbox" name="IsFont" value="1" checked>
				SPAN
				<input type="checkbox" name="IsSpan" value="1">
				OBJECT
				<input type="checkbox" name="IsObject" value="1" checked>
				IFRAME
				<input type="checkbox" name="IsIFrame" value="1" checked>
				SCRIPT
				<input type="checkbox" name="IsScript" value="1" checked>
			</td>
		</tr>
		<tr class="hback">
			<td height="26"><div align="right">关键字过滤规则：</div></td>
			<td>
			<input style="width:50%;" name="CS_SiteReKeyName" type="text" id="CS_SiteReKeyName" value="" readonly />
			<input style="width:50%;" name="CS_SiteReKeyID" type="hidden" id="CS_SiteReKeyID" value="" />
			<% = GetAllRulerList() %>
			</td>
		</tr>
		<tr class="hback">
			<td height="26"><div align="right">入库栏目：</div></td>
			<td><select name="ToClass" style="width:100%;" id="ToClass">
				<%
					Dim obj_unite_rs,tmp_str_list
					Set obj_unite_rs = server.CreateObject(G_FS_RS)
					obj_unite_rs.Open "Select Orderid,id,ClassID,ClassName,ParentID from FS_NS_NewsClass where IsURL=0 And Parentid  = '0' and ReycleTF=0 Order by Orderid desc,ID desc",Conn,1,1
					  tmp_str_list  = ""
					  do while Not obj_unite_rs.eof 
						tmp_str_list = tmp_str_list &"<option value="""& obj_unite_rs("ClassID") &""">+"& obj_unite_rs("ClassName") &"</option>"& Chr(13) & Chr(10)
						tmp_str_list = tmp_str_list &UniteChildNewsList(obj_unite_rs("ClassID"),"")
						 obj_unite_rs.movenext
					 Loop
					 obj_unite_rs.close
					 set obj_unite_rs = nothing
					 Response.Write tmp_str_list
				%>
				</select></td>
		</tr>
		<tr class="hback">
			<td height="26"><div align="right">新闻模板：</div></td>
			<td>
			<input style="width:70%;" name="NewsTemp" type="text" id="NewsTemp" value="/Templets/NewsClass/news.htm" readonly>
			<input name="Submit5" type="button" id="selNewsTemplet" value="选择模板"  onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir%>/<% = G_TEMPLETS_DIR %>',400,300,window,document.AddSiteForm.NewsTemp);document.AddSiteForm.NewsTemp.focus();">
				</td>
		</tr>
		<tr class="hback">
			<td height="26"><div align="right">定时采集设置：</div></td>
			<td>
			<select name="AutoCTF" id="AutoCTF" style="width:30%;" onChange="Javascript:Setautocellecttime(this.options[this.selectedIndex].value)">
				<option value="no" selected="selected">不自动采集</option>
				<option value="day">每天</option>
				<option value="week">每周</option>
				<option value="month">每月</option>
			</select>
			<select name="TimeWeek" id="TimeWeek" style="width:30%; display:none;">
				<option value="7" selected="selected">周日</option>
				<option value="1">周一</option>
				<option value="2">周二</option>
				<option value="3">周三</option>
				<option value="4">周四</option>
				<option value="5">周五</option>
				<option value="6">周六</option>
			</select>
			<select name="TimeMonth" id="TimeMonth" style="width:30%; display:none;">
				<option value="1" selected="selected">1</option>
				<option value="2">2</option>
				<option value="3">3</option>
				<option value="4">4</option>
				<option value="5">5</option>
				<option value="6">6</option>
				<option value="7">7</option>
				<option value="8">8</option>
				<option value="9">9</option>
				<option value="10">10</option>
				<option value="11">11</option>
				<option value="12">12</option>
				<option value="13">13</option>
				<option value="14">14</option>
				<option value="15">15</option>
				<option value="16">16</option>
				<option value="17">17</option>
				<option value="18">18</option>
				<option value="19">19</option>
				<option value="20">20</option>
				<option value="21">21</option>
				<option value="22">22</option>
				<option value="23">23</option>
				<option value="24">24</option>
				<option value="25">25</option>
				<option value="26">26</option>
				<option value="27">27</option>
				<option value="28">28</option>
				<option value="29">29</option>
				<option value="30">30</option>
				<option value="31">31</option>
			</select>
			<select name="TimeHour" id="TimeHour" style="width:30%; display:none;">
				<option value="00:00" selected="selected">00:00</option>
				<option value="00:30">00:30</option>
				<option value="01:00">01:00</option>
				<option value="01:30">01:30</option>
				<option value="02:00">02:00</option>
				<option value="02:30">02:30</option>
				<option value="03:00">03:00</option>
				<option value="03:30">03:30</option>
				<option value="04:00">04:00</option>
				<option value="04:30">04:30</option>
				<option value="05:00">05:00</option>
				<option value="05:30">05:30</option>
				<option value="06:00">06:00</option>
				<option value="06:30">06:30</option>
				<option value="07:00">07:00</option>
				<option value="07:30">07:30</option>
				<option value="08:00">08:00</option>
				<option value="08:30">08:30</option>
				<option value="09:00">09:00</option>
				<option value="09:30">09:30</option>
				<option value="10:00">10:00</option>
				<option value="10:30">10:30</option>
				<option value="11:00">11:00</option>
				<option value="11:30">11:30</option>
				<option value="12:00">12:00</option>
				<option value="12:30">12:30</option>
				<option value="13:00">13:00</option>
				<option value="13:30">13:30</option>
				<option value="14:00">14:00</option>
				<option value="14:30">14:30</option>
				<option value="15:00">15:00</option>
				<option value="15:30">15:30</option>
				<option value="16:00">16:00</option>
				<option value="16:30">16:30</option>
				<option value="17:00">17:00</option>
				<option value="17:30">17:30</option>
				<option value="18:00">18:00</option>
				<option value="18:30">18:30</option>
				<option value="19:00">19:00</option>
				<option value="19:30">19:30</option>
				<option value="20:00">20:00</option>
				<option value="20:30">20:30</option>
				<option value="21:00">21:00</option>
				<option value="21:30">21:30</option>
				<option value="22:00">22:00</option>
				<option value="22:30">22:30</option>
				<option value="23:00">23:00</option>
				<option value="23:30">23:30</option>
			</select>
			</td>
		</tr>
		<tr class="hback">
			<td height="26"><div align="right">单次采集数量：</div></td>
			<td><input style="width:100%;" name="NewsNum" type="text" id="NewsNum" value="10" onKeyUp=if(isNaN(value)||event.keyCode==32)execCommand('undo');  onafterpaste=if(isNaN(value)||event.keyCode==32)execCommand('undo');>
			</td>
		</tr>
		<tr class="hback">
			<td height="26"><div align="right">图片保存路径：</div></td>
			<td><input name="PicSavePath" type="text" id="PicSavePath" style="width:40%" value="<% = SaveRemotePicPath %>" readonly>
			<INPUT type="button"  name="Submit4" value="选择路径" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPathFrame.asp?CurrPath=<%= str_CurrPath %>',300,250,window,document.AddSiteForm.PicSavePath);document.AddSiteForm.PicSavePath.focus();">
			保存远程图片自动水印
		<input type="checkbox" name="WaterPrint" value="1" checked>
			</td>
		</tr>
  </table>
</form>
<% End Sub

Sub AddFolder()
%>
<form name="AddSiteFolderForm" method="post" action="">
<table width="98%" height="20" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
					<tr class="hback">
						<td width=35 style="cursor:hand" align="center" alt="保存" onClick="document.AddSiteFolderForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
						<td width=35 style="cursor:hand" align="center" alt="后退" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
						<td>&nbsp;
							<input name="vs" type="hidden" id="vs2" value="addfolder">
              <input type="hidden" name="SiteFolderID" value="<% = SiteFolderID %>"> </td>
					</tr>
  </table>	
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<%
	Dim SiteFolderID,RsSiteFolder
	SiteFolderID = Request("SiteFolderID")
	If SiteFolderID<>"" Then
		Set RsSiteFolder = CollectConn.Execute("select * from FS_SiteFolder where ID=" & CintStr(SiteFolderID))
		If RsSiteFolder.EOF Then
			Response.write "<script>alert('该栏目不存在！');history.back();</script>"
			Response.End
		End If
	%>
		<tr class="hback">
			<td width="100" height="26"><div align="center">栏目名称:</div></td>
			<td><input style="width:100%" type="text" name="SiteFolder" value="<%=RsSiteFolder("SiteFolder")%>"></td>
		</tr>
		<tr class="hback">
			<td width="100" height="26"><div align="center">栏目说明:</div></td>
			<td><textarea style="width:100%" name="SiteFolderDetail" rows="10"><%=RsSiteFolder("SiteFolderDetail")%></textarea></td>
		</tr>
		<%
	Else
	%>
		<tr class="hback">
			<td width="100" height="26"><div align="center">栏目名称:</div></td>
			<td><input style="width:100%" type="text" name="SiteFolder"></td>
		</tr>
		<tr class="hback">
			<td width="100" height="26"><div align="center">栏目说明:</div></td>
			<td><textarea style="width:100%" name="SiteFolderDetail" rows="10"></textarea></td>
		</tr>
		<%
	End If
	%>
  </table>
</form>
<% End Sub %>
</body>
</html>
<%
Function FolderList()
	Dim FolderListObj,StrSelected
	Set FolderListObj = Collectconn.Execute("Select * from FS_SiteFolder where 1=1 order by ID desc")
	do while Not FolderListObj.Eof
		If CInt(Request("FolderID"))=FolderListObj("ID") Then
			StrSelected="selected"
		Else
			StrSelected=""
		End If
		FolderList = FolderList & "<option value="&FolderListObj("ID")&" " & StrSelected & ">&nbsp;&nbsp;|--" & FolderListObj("SiteFolder") & "</option><br>"
		FolderListObj.MoveNext	
	loop
	FolderListObj.Close
	Set FolderListObj = Nothing
End Function
Set CollectConn = Nothing


Function UniteChildNewsList(TypeID,f_CompatStr)  
	Dim f_ChildNewsRs,ChildTypeListStr,f_TempStr,f_isUrlStr,lng_GetCount
	Set f_ChildNewsRs = Conn.Execute("Select id,orderid,ClassName,ClassID,ParentID from FS_NS_NewsClass where IsURL=0 And ParentID='" & NoSqlHack(TypeID) & "' and ReycleTF=0 order by Orderid desc,id desc" )
	f_TempStr =f_CompatStr & "┄"
	do while Not f_ChildNewsRs.Eof
			UniteChildNewsList = UniteChildNewsList & "<option value="""& f_ChildNewsRs("ClassID") &""">"
			UniteChildNewsList = UniteChildNewsList & "├" &  f_TempStr & f_ChildNewsRs("ClassName") 
			UniteChildNewsList = UniteChildNewsList & "</option>" & Chr(13) & Chr(10)
			UniteChildNewsList = UniteChildNewsList &UniteChildNewsList(f_ChildNewsRs("ClassID"),f_TempStr)
		f_ChildNewsRs.MoveNext
	loop
	f_ChildNewsRs.Close
	Set f_ChildNewsRs = Nothing
End Function


Function GetAllRulerList()
	Dim Rs
	GetAllRulerList = "<select name=""CS_SiteReKey"" id=""CS_SiteReKey"" style=""width:48%;"" onChange=""Javascript:SelectRuler(this.options[this.selectedIndex]);"">" & VbNewLine
	GetAllRulerList = GetAllRulerList & VbTab & VbTab & " <option value="""">选择规则</option>" & VbNewLine
	GetAllRulerList = GetAllRulerList & VbTab & VbTab & " <option value=""Clear"" style=""color:#FF0000;"">清除</option>" & VbNewLine
	Set Rs = CollectConn.ExeCute("Select ID,RuleName From FS_Rule Where ID > 0 Order By ID Desc")
	Do While Not Rs.Eof
		GetAllRulerList = GetAllRulerList & VbTab & VbTab & " <option value=""" & Rs(0) & """>" & Rs(1) & "</option>" & VbNewLine
	Rs.MoveNext
	Loop
	GetAllRulerList = GetAllRulerList & VbTab & VbTab & "</select>"	
	Rs.Close : Set Rs = Nothing
End Function


Conn.Close : Set Conn = Nothing
%>
<script language="JavaScript">
function InsertScript()
{
	var ReturnValue='';
	ReturnValue=showModalDialog("NewsNum.asp",window,'dialogWidth:260pt;dialogHeight:120pt;status:no;help:no;scroll:no;');
	return ReturnValue;
}
function StartOneSiteCollect(ID)
{
	Num = InsertScript();
	if (Num!='back'&& Num!='0')
	{
		if (Num==""||Num==null) Num="allNews"
		location='Collecting.asp?SiteID='+ID+'&Num='+Num;
	}
}

function Setautocellecttime(str)
{
	if (str == "" || str == "no")
	{
		document.getElementById("TimeHour").style.display = "none";
		document.getElementById("TimeWeek").style.display = "none";
		document.getElementById("TimeMonth").style.display = "none";
	}
	if (str == "day")
	{
		document.getElementById("TimeHour").style.display = "";
		document.getElementById("TimeWeek").style.display = "none";
		document.getElementById("TimeMonth").style.display = "none";
	}
	else if (str == "week")
	{
		document.getElementById("TimeHour").style.display = "";
		document.getElementById("TimeWeek").style.display = "";
		document.getElementById("TimeMonth").style.display = "none";
	}
	else if (str == "month")
	{
		document.getElementById("TimeHour").style.display = "";
		document.getElementById("TimeWeek").style.display = "none";
		document.getElementById("TimeMonth").style.display = "";
	}
}

function SelectRuler(obj)
{
	var Ruler_V = obj.value;
	var Ruler_Str = obj.innerText;
	var Name_Str = document.AddSiteForm.CS_SiteReKeyName.value;
	var ID_Str = document.AddSiteForm.CS_SiteReKeyID.value;
	if (Ruler_V == '')
	{
		return;
	}
	if (Ruler_V == 'Clear')
	{
		document.AddSiteForm.CS_SiteReKeyName.value = '';
		document.AddSiteForm.CS_SiteReKeyID.value = '';
	}
	else
	{
		if (Name_Str == '' || ID_Str == '')
		{
			document.AddSiteForm.CS_SiteReKeyName.value = Ruler_Str;
			document.AddSiteForm.CS_SiteReKeyID.value = Ruler_V;
		}
		else
		{
			if (ID_Str.indexOf(Ruler_V) != -1)
			{
				document.AddSiteForm.CS_SiteReKeyName.value = Name_Str;
				document.AddSiteForm.CS_SiteReKeyID.value = ID_Str;
			}
			else
			{
				document.AddSiteForm.CS_SiteReKeyName.value = Name_Str + ',' + Ruler_Str;
				document.AddSiteForm.CS_SiteReKeyID.value = ID_Str + ',' + Ruler_V;
			}
		}
	}
}
</script>






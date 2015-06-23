<%@  language="VBSCRIPT" codepage="936" %>
<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<%
MF_Default_Conn
MF_Session_TF
Dim Str_Soft_Version,Str_Action,FoosunMFVersion
Dim Str_conn,Str_User_Conn,ConfigFilePath,Conn,User_Conn,Old_News_Conn,Str_Old_Conn,Collect_Conn,Str_Collect_Conn

Str_Soft_Version="5.0 build 20100507"
ConfigFilePath="FS_Inc/Const.asp"
Str_Action=Trim(Request.QueryString("Act"))
FoosunMFVersion = Request.Cookies("FoosunMFCookies")("FoosunMFVersion")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<title>FoosunCMS<%= Str_Soft_Version %>数据库升级程序</title>
	<style type="text/css">
		body, td, th { font-size: 12px; }
		#up { text-align: center; }
		a { text-decoration: none; }
		a:link { color: #000000; }
		a:visited { color: #000000; }
		a:hover { color: #FF0000; }
		a:active { color: #FF0000; }
		.nav_l2 { text-decoration: none; font-size: 14px; color: #FFFFFF; height: 0; filter: dropshadow(offX=1,offY=1,color=002E5B); }
		A.nav_l2:link { text-decoration: none; font-size: 14px; color: #FFFFFF; height: 0; }
		A.nav_l2:visited { text-decoration: none; font-size: 14px; color: #FFFFFF; height: 0; }
		A.nav_l2:active { text-decoration: underline; font-size: 14px; color: #FFFFFF; height: 0; }
		A.nav_l2:hover { text-decoration: underline; font-size: 14px; color: #fff; height: 0; }
		body { font-family: "新宋体"; font-size: 12px; text-decoration: none; line-height: 150%; background: #FFFFFF; color: #000000; padding: 0px; margin-top: 0px; margin-right: 0px; margin-bottom: 0px; margin-left: 0px; }
		Button, input, textarea, select { font-family: "Verdana, 新宋体"; font-size: 12px; color: #000000; text-decoration: none; background: #E0EEF5; }
		.xk { border: 1px solid #013767; margin-top: 5px; margin-bottom: 5px; background: #B0DAF0; }
		.table { border: 1px solid #004A6F; margin-top: 5px; margin-bottom: 5px; background: #FFFFFF; }
		.leftframetable { border: 1px solid #013767; margin-top: 3px; margin-bottom: 3px; background: #E0EEF5; }
		.titledaohang { font-size: 12px; color: #0066CC; font-weight: bolder; font-family: '宋体'; }
		A.admintx:link { font-weight: normal; font-size: 12px; color: #053EC0; text-decoration: none; font-weight: bolder; }
		A.admintx:visited { font-weight: normal; font-size: 12px; color: #053EC0; text-decoration: none; font-weight: bolder; }
		A.admintx:hover { font-weight: normal; font-size: 12px; color: #FF0000; text-decoration: none; font-weight: bolder; }
		A.admintx:active { font-weight: normal; font-size: 12px; color: #FF0000; text-decoration: none; font-weight: bolder; }
		.bt { font-size: 12px; font-weight: bold; color: #EBEBEB; text-decoration: none; }
		td { color: #000000; font-size: 12px; line-height: 18px; }
		.xingmu { font-family: "Verdana, 新宋体"; color: #FFFFFF; font-size: 12px; font-weight: bolder; background: #0F70C7; line-height: 22px; }
		.back { background: #016470; }
		.Leftback { background: #0F70C7; }
		.hback { background: #E0EEF5; }
		.hback_1 { background: #CFE4EF; }
		.bordercolor { cursor: col-resize; border-left: none; border-bottom: none; border-top: none; border-right: solid; border-right-width: 3px; border-right-color: #eeeeee; }
		.pagetile { font-weight: bolder; font-size: 20px; }
		.ischeck { color: #FF9900; }
		a.CusCss1.link { color: #FFFFFF; }
		.CusCss2 { color: #00CCCC; }
		MenuLine { color: #FFFFFF; }
		.CusCss3 { color: #FFCCCC; }
		.alps { font-size: 9pt; color: #8AADE9; font-family: '宋体'; }
		A.top:link { font-weight: normal; font-size: 12px; color: #FFFFFF; text-decoration: none; }
		A.top:visited { font-weight: normal; font-size: 12px; color: #FFFFFF; text-decoration: none; }
		A.top:hover { font-weight: normal; font-size: 12px; color: #CC3300; text-decoration: none; }
		A.top:active { font-weight: normal; font-size: 12px; color: #000000; text-decoration: none; }
		A.otherset:link { font-weight: normal; font-size: 12px; color: #000000; text-decoration: none; }
		A.otherset:visited { font-weight: normal; font-size: 12px; color: #000000; text-decoration: none; }
		A.otherset:hover { font-weight: normal; font-size: 12px; color: #FF0000; text-decoration: none; }
		A.otherset:active { font-weight: normal; font-size: 12px; color: #FF0000; text-decoration: none; }
		.tx { color: red; }
		A.sd:link { font-weight: normal; font-size: 12px; color: #FFFFFF; text-decoration: none; }
		A.sd:visited { font-weight: normal; font-size: 12px; color: #FFFFFF; text-decoration: none; }
		A.sd:hover { font-weight: normal; font-size: 12px; color: yellow; text-decoration: none; }
		A.sd:active { font-weight: normal; font-size: 12px; color: yellow; text-decoration: none; }
		.RightInput { font-size: 12px border-bottom:solid; border-bottom-color: #6FBE44; border-bottom-width: 1; border-left: solid; border-left-color: #6FBE44; border-left-width: 1; border-right: solid; border-right-color: #6FBE44; border-right-width: 1; border-top: solid; border-top-color: #6FBE44; border-top-width: 1; }
		.WarnInput { font-size: 12px border-bottom:solid; border-bottom-color: #FF0000; border-bottom-width: 1; border-left: solid; border-left-color: #FF0000; border-left-width: 1; border-right: solid; border-right-color: #FF0000; border-right-width: 1; border-top: solid; border-top-color: #FF0000; border-top-width: 1; }
	</style>
</head>
<body>
	<div id="up">
		<br />
		<br />
		<br />
		<br />
		<table width="60%" border="0" cellspacing="1" cellpadding="5" class="table">
			<tr>
				<td class="xingmu">
					FOOSUN CMS
					<%= Str_Soft_Version %>
					数据库升级程序
				</td>
			</tr>
			<tr>
				<td class="hback_1">
					适用于FOOSUN CMS 4.0 所有版本，系统检测到当前版本为<% = FoosunMFVersion %>。<br />
					<font color="#CC0000"><strong>如果不升级会导致系统不能够正常运行！</strong></font>
				</td>
			</tr>
		</table>
		<%
If Str_Action="" Then
		%>
		<table width="60%" border="0" cellspacing="1" cellpadding="5" class="table">
			<tr>
				<td class="hback">
					请确认您的以下配置是否正确
				</td>
			</tr>
			<tr>
				<td class="hback">
					1、主数据库连接;<br />
					2、会员数据库连接;<br />
					2、归档数据库连接;<br />
					3、配置文件是否可写(FS_Inc/Const.asp);<br />
					4、确认当前系统版本<select name="CurrVersion" id="CurrVersion">
						<option value="0">4.0</option>
						<option <% if InStr(LCase(FoosunMFVersion),"sp1") > 0 then Response.Write("Selected") %> value="1">4.0 SP1</option>
						<option <% if InStr(LCase(FoosunMFVersion),"sp2") > 0 then Response.Write("Selected") %> value="2">4.0 SP2</option>
						<option <% if InStr(LCase(FoosunMFVersion),"sp3") > 0 then Response.Write("Selected") %> value="3">4.0 SP3</option>
						<option <% if InStr(LCase(FoosunMFVersion),"sp4") > 0 then Response.Write("Selected") %> value="4">4.0 SP4</option>
						<option <% if InStr(LCase(FoosunMFVersion),"sp5") > 0 then Response.Write("Selected") %> value="5">4.0 SP5</option>
						<option <% if InStr(LCase(FoosunMFVersion),"sp6") > 0 then Response.Write("Selected") %> value="6">4.0 SP6</option>
						<option <% if InStr(LCase(FoosunMFVersion),"sp7") > 0 then Response.Write("Selected") %> value="7">4.0 SP7</option>
						<option <% if InStr(LCase(FoosunMFVersion),"5.0") > 0 then Response.Write("Selected") %> value="8">5.0</option>
						<option <% if InStr(LCase(FoosunMFVersion),"build 20091111") > 0 then Response.Write("Selected") %> value="9">5.0 build 20091111</option>
						<option <% if InStr(LCase(FoosunMFVersion),"build 20100129") > 0 then Response.Write("Selected") %> value="10">5.0 build 20100129</option>
					</select>;<br />
					如果数据库连接不正确，请修改FS_Inc/Const.asp文件里面的数据库连接参数。
				</td>
			</tr>
			<tr>
				<td class="hback">
					<input <% if InStr(Request.Cookies("FoosunMFCookies")("FoosunMFVersion"),Str_Soft_Version) <> 0 then Response.Write("disabled") %> type="button" name="check2" value="我确认配置好了，开始升级到<%= Str_Soft_Version %>" id="check2" onclick="location='?Act=Up&Version='+document.all.CurrVersion.value;" />
					&nbsp;
				</td>
			</tr>
		</table>
		<%
Else
	response.Write "<table width=""60%"" border=""0"" cellspacing=""1"" cellpadding=""5"" class=""table""><tr><td class=hback>"
	If G_IS_SQL_DB=1 Then
		Str_conn = "Provider=SQLOLEDB.1;"& G_DATABASE_CONN_STR &";"
	Else
		Str_conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mapPath(Add_Root_Dir(G_DATABASE_CONN_STR))
	End If

	If G_IS_SQL_User_DB=1 Then
		Str_User_Conn = "Provider=SQLOLEDB.1;"& G_User_DATABASE_CONN_STR &";"
	Else
		Str_User_Conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mapPath(Add_Root_Dir(G_User_DATABASE_CONN_STR))
	End If
	If G_IS_SQL_Old_News_DB=1 Then
		Str_Old_Conn = "Provider=SQLOLEDB.1;"& G_Old_News_DATABASE_CONN_STR & ";"
	Else
		Str_Old_Conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mapPath(Add_Root_Dir(G_Old_News_DATABASE_CONN_STR))
	End If
	If G_IS_SQL_Collect_DB = 1 Then
		Str_Collect_Conn = "Provider=SQLOLEDB.1;"& G_COLLECT_DATA_STR & ";"
	Else
		Str_Collect_Conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mapPath(Add_Root_Dir(G_COLLECT_DATA_STR))
	End If

	Set Conn = Server.CreateObject(G_FS_CONN)
	Set User_Conn = Server.CreateObject(G_FS_CONN)
	Set Old_News_Conn = Server.CreateObject(G_FS_CONN)
	Set Collect_Conn = Server.CreateObject(G_FS_CONN)
	Conn.Open Str_conn
	User_Conn.Open Str_User_Conn
	Old_News_Conn.Open Str_Old_Conn
	Collect_Conn.Open Str_Collect_Conn
	Str_Action=UpDateBase()
	Conn.Close
	Set Conn = Nothing
	User_Conn.Close
	Set User_Conn = Nothing
	Old_News_Conn.Close
	Set Old_News_Conn = Nothing
	Collect_Conn.Close
	Set Collect_Conn = Nothing
	Response.Write "<span class=tx>升级完成。</span></td></tr></table>"
End If
		%>
		<table width="60%" border="0" cellspacing="1" cellpadding="5" class="table">
			<tr>
				<td class="hback_1">
					<div align="center" class="bordercolor">
						Power By Foosun.Cn<br />
						&copy; 2002-2009 Foosun Inc.<br />
						四川风讯科技发展有限公司 版权所有</div>
				</td>
			</tr>
		</table>
		<br />
		<br />
	</div>
</body>
</html>
<%
Function OutPutInfo(f_Str)
	If Err Then
		Response.Write(f_Str & "升级失败！<br>原因：<font color=red>"&Err.Description&"</font><br />")
		Err.Clear
	Else
		Response.Write f_Str & "升级成功！<br />"
	End If
End Function
Function UpDateBase()
	On Error Resume Next
	Dim f_UserDefineVersion
	f_UserDefineVersion = NoSqlHack(Request.QueryString("Version"))
	if Not IsNumeric(f_UserDefineVersion) then f_UserDefineVersion = 7
	f_UserDefineVersion = CInt(f_UserDefineVersion)
	if f_UserDefineVersion < 0 then f_UserDefineVersion = 0
	'if f_UserDefineVersion > 8 then f_UserDefineVersion = 8
	if f_UserDefineVersion < 1 then UpDateBase = UpDateBase & SP1()
	if f_UserDefineVersion < 5 then
		UpDateBase = UpDateBase & SP2TOSP5()
		UpDateBase = UpDateBase & SP2TOSP5_UpDateLable()
	end if
	if f_UserDefineVersion < 6 then UpDateBase = UpDateBase & SP6()
	if f_UserDefineVersion < 7 then UpDateBase = UpDateBase & SP7()
	if f_UserDefineVersion < 8 then UpDateBase = UpDateBase & UP5()
	if f_UserDefineVersion < 9 then UpDateBase = UpDateBase & UP520091111()
	if f_UserDefineVersion < 10 then UpDateBase = UpDateBase & UP520100129()
	if f_UserDefineVersion < 11 then UpDateBase = UpDateBase & UP520100507()
End Function
Function SP1()
	OutPutInfo "SP1"
End Function

Function SP2TOSP5()
	Dim RenameRS,key,ExistTF,GroupID,Str_Temp,ChangeCount,Str_Fields,Str_Sql
	ChangeCount = 0
	Set RenameRS = Server.CreateObject(G_FS_RS)
	'主数据库______________________________________________________________________
	RenameRS.Open "SELECT LogContent FROM FS_MF_Oper_Log WHERE 1=0",Conn,1,3
	ExistTF=RenameRS.Fields("LogContent").Type
	RenameRS.Close
	If ExistTF<>203 Then
		Conn.execute("ALTER TABLE FS_MF_Oper_Log ALTER COLUMN LogContent ntext NULL")
		ChangeCount = ChangeCount + 1
	End If
	'______________________________________________________________________________
	RenameRS.Open "SELECT * FROM FS_MF_Labestyle WHERE 1=0",Conn,1,3
	ExistTF=False
	For key=0 To RenameRS.Fields.count-1
		If LCase(RenameRS.Fields(key).name)=LCase("LableClassID") Then
			ExistTF=True
		End If
	Next
	RenameRS.Close
	If Not ExistTF Then
		Conn.execute("ALTER TABLE FS_MF_Labestyle ADD LableClassID int NULL")
		ChangeCount = ChangeCount + 1
	End If
	Conn.ExeCute("Update FS_MF_Labestyle Set LableClassID = 0 Where ID > 0")
	'______________________________________________________________________________
	ExistTF=False
	If G_IS_SQL_DB = 1 Then
		Set Str_Temp=Conn.Execute("Select * from dbo.sysobjects where id = object_id(N'[FS_MF_StyleClass]') and OBJECTPROPERTY(id, N'IsUserTable') = 1")
		If Not Str_Temp.Eof Then
			ExistTF=True
		End IF
	Else
		Set Str_Temp=Conn.openSchema(20)
		Str_Temp.MoveFirst
		Do While not Str_Temp.EOF
			if Str_Temp("TABLE_TYPE")="TABLE" then
				if  Str_Temp("TABLE_NAME")="FS_MF_StyleClass" then
					ExistTF=True
					Exit Do
				End if
			End if
			Str_Temp.MoveNext
		Loop
	End If
	If Not ExistTF Then
		Conn.Execute("CREATE TABLE [FS_MF_StyleClass] ([ID] int NOT NULL identity primary key,[ClassName] nvarchar(50) NULL,[ClassContent] nvarchar(50) NULL,[ParentID] int NULL)")
		ChangeCount = ChangeCount + 1
	End If
	'______________________________________________________________________________
	ExistTF=False
	If G_IS_SQL_DB = 1 Then
		Set Str_Temp=Conn.Execute("Select * from dbo.sysobjects where id = object_id(N'[FS_MF_FreeLabel]') and OBJECTPROPERTY(id, N'IsUserTable') = 1")
		If Not Str_Temp.Eof Then
			ExistTF=True
		End IF
	Else
		Set Str_Temp=Conn.openSchema(20)
		Str_Temp.MoveFirst
		Do While not Str_Temp.EOF
			if Str_Temp("TABLE_TYPE")="TABLE" then
				if  Str_Temp("TABLE_NAME")="FS_MF_FreeLabel" then
					ExistTF=True
					Exit Do
				End if
			End if
			Str_Temp.MoveNext
		Loop
	End If
	If Not ExistTF Then
		Conn.Execute("CREATE TABLE [FS_MF_FreeLabel] ([ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL ,	[LabelID] nvarchar(20) NOT NULL ,[LabelName] nvarchar(50) NOT NULL ,[LabelSQl] ntext NULL ,	[NSFields] ntext NULL ,	[NCFields] ntext NULL ,	[LabelContent] ntext NULL ,	[selectNum] int NULL DEFAULT 0 ,	[DesCon] nvarchar(250) NULL ,[SysType] nvarchar(3) NULL)")
		Conn.Execute("CREATE  UNIQUE  INDEX [LabelID] ON [FS_MF_FreeLabel]([LabelID])")
		Conn.Execute("CREATE  UNIQUE  INDEX [LabelName] ON [FS_MF_FreeLabel]([LabelName])")
		Conn.Execute("CREATE  INDEX [selectNum] ON [FS_MF_FreeLabel]([selectNum])")
		ChangeCount = ChangeCount + 1
	End If
	'______________________________________________________________________________
	RenameRS.Open "SELECT * FROM FS_MF_Config WHERE 1=0",Conn,1,3
	ExistTF=False
	For key=0 To RenameRS.Fields.count-1
		If LCase(RenameRS.Fields(key).name)=LCase("Style_MaxNum") Then
			ExistTF=True
		End If
	Next
	RenameRS.Close
	If Not ExistTF Then
		Conn.execute("ALTER TABLE FS_MF_Config ADD Style_MaxNum int NULL DEFAULT 50,Define_MaxNum int NULL DEFAULT 50")
		Conn.execute("UPDATE FS_MF_Config SET Style_MaxNum=50,Define_MaxNum=50")
		ChangeCount = ChangeCount + 1
	End If
	'------------------------------------------------------------------------------
	RenameRS.Open "SELECT * FROM FS_MF_Config WHERE 1=0",Conn,1,3
	ExistTF=False
	For key=0 To RenameRS.Fields.count-1
		If LCase(RenameRS.Fields(key).name)=LCase("LabelContent_MaxNum") Then
			ExistTF=True

		End If
	Next
	RenameRS.Close
	If Not ExistTF Then
		Conn.execute("ALTER TABLE FS_MF_Config ADD LabelContent_MaxNum int NULL DEFAULT 0")
		Conn.execute("UPDATE FS_MF_Config SET LabelContent_MaxNum=0")
		ChangeCount = ChangeCount + 1
	End If
	'------------------------------------------------------------------------------
	RenameRS.Open "SELECT * FROM FS_MF_Admin WHERE 1=0",Conn,1,3
	ExistTF=False
	For key=0 To RenameRS.Fields.count-1
		If LCase(RenameRS.Fields(key).name)=LCase("Admin_FilesTF") Then
			ExistTF=True
		End If
	Next
	RenameRS.Close
	If Not ExistTF Then
		Conn.execute("ALTER TABLE FS_MF_Admin ADD Admin_FilesTF Smallint NULL DEFAULT 0")
		Conn.execute("UPDATE FS_MF_Admin SET Admin_FilesTF = 0")
		ChangeCount = ChangeCount + 1
	End If
	'______________________________________________________________________________
	If Err Then
		Response.Write("SP2-SP5主数据库升级失败！共"&ChangeCount&"处改动！<br>原因：<font color=red>"&Err.Description&"</font><br />")
		Err.Clear
	Else
		Response.Write "SP2-SP5主数据库升级成功！共"&ChangeCount&"处改动！<br />"
	End If
	'主数据库End___________________________________________________________________

	'子系统开始____________________________________________________________________

	'新闻数据库____________________________________________________________________
	ChangeCount = 0
	If IsExist_SubSys("NS") Then
		RenameRS.Open "SELECT * FROM FS_NS_NewsClass WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("IsAdPic") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_NS_NewsClass ADD IsAdPic nvarchar(1) NULL DEFAULT '0',AdPicWH nvarchar(20) NULL,AdPicLink nvarchar(250) NULL,AdPicAdress nvarchar(250) NULL")
			Conn.execute("Update FS_NS_NewsClass Set IsAdPic='0'")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_NS_TodayPic WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("TodayTitle") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_NS_TodayPic ADD TodayTitle nvarchar(150) NULL,TodayWidth int NULL DEFAULT 300")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT AdPicWH FROM FS_NS_NewsClass WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("AdPicWH").DefinedSize
		RenameRS.Close
		If ExistTF<>20 Then
			Conn.execute("ALTER TABLE FS_NS_NewsClass ALTER COLUMN AdPicWH nvarchar(20) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_NS_News WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("IsAdPic") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_NS_News ADD IsAdPic nvarchar(1) NULL DEFAULT '0',AdPicWH nvarchar(20) NULL,AdPicLink nvarchar(250) NULL,AdPicAdress nvarchar(250) NULL")
			Conn.execute("Update FS_NS_News Set IsAdPic='0'")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT AdPicWH FROM FS_NS_News WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("AdPicWH").DefinedSize
		RenameRS.Close
		If ExistTF<>20 Then
			Conn.execute("ALTER TABLE FS_NS_News ALTER COLUMN AdPicWH nvarchar(20) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_NS_SysParam WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("IsCopyFileTF") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_NS_SysParam ADD IsCopyFileTF tinyint NULL default 0")
			Conn.execute("Update FS_NS_SysParam Set IsCopyFileTF = 0 Where SysID > 0")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_NS_SysParam WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("IsEditFileTF") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_NS_SysParam ADD IsEditFileTF tinyint NULL default 0")
			Conn.execute("Update FS_NS_SysParam Set IsEditFileTF = 0 Where SysID > 0")
			ChangeCount = ChangeCount + 1
		End If
		'----------------------------------------------------------------
		'2007-07-17
		RenameRS.Open "SELECT * FROM FS_NS_Special WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("FileSaveType") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_NS_Special ADD FileSaveType tinyint NULL default 3")
			Conn.execute("Update FS_NS_Special Set FileSaveType = 3 Where SpecialID > 0")
			ChangeCount = ChangeCount + 1
		End If
		'----------------------------------------------------------------
		RenameRS.Open "SELECT * FROM FS_NS_FreeJsFile WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("NewsID") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_NS_FreeJsFile ADD NewsID nvarchar(50) NULL default 0")
			If G_IS_SQL_DB = 1 Then
				Conn.ExeCute("Update FS_NS_FreeJsFile Set NewsID = (select top 1 NewsID From FS_NS_News Where FileName = FS_NS_FreeJsFile.FileName)")
			Else
				Conn.ExeCute("Update FS_NS_FreeJsFile,FS_NS_News Set FS_NS_FreeJsFile.NewsID = FS_NS_News.NewsID Where FS_NS_News.FileName = FS_NS_FreeJsFile.FileName")
			End if	
			Conn.ExeCute("ALTER TABLE FS_NS_FreeJsFile DROP COLUMN FileName")			
			ChangeCount = ChangeCount + 1
		End If
		
		'________________________________________________________________
		If Err Then
			Response.Write("SP2-SP5新闻数据库升级失败！共"&ChangeCount&"处改动！<br>原因：<font color=red>"&Err.Description&"</font><br />")
			Err.Clear
		Else
			Response.Write "SP2-SP5新闻数据库升级成功！共"&ChangeCount&"处改动！<br />"
		End If
	End If
	'新闻数据库End_________________________________________________________________

	'房产数据库____________________________________________________________________
	ChangeCount = 0
	If IsExist_SubSys("HS") Then
		RenameRS.Open "SELECT * FROM FS_HS_Second WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("Properties") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_HS_Second ADD Properties nvarchar(20) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_HS_Second WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("Click") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_HS_Second ADD Click int NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_HS_Tenancy WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("Click") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_HS_Tenancy ADD Click int NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_HS_Tenancy WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("XiaoQuName") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_HS_Tenancy ADD XiaoQuName nvarchar(200) NULL,XingZhi smallint NULL,ZaWuJian smallint NULL DEFAULT 0,JiaoTong nvarchar(250) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_HS_Quotation WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("KaiFaShang") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_HS_Quotation ADD KaiFaShang nvarchar(100) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT equip FROM FS_HS_Second WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("equip").DefinedSize
		RenameRS.Close
		If ExistTF<>20 Then
			Conn.execute("ALTER TABLE FS_HS_Second ALTER COLUMN equip nvarchar(20) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT equip FROM FS_HS_Tenancy WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("equip").DefinedSize
		RenameRS.Close
		If ExistTF<>20 Then
			Conn.execute("ALTER TABLE FS_HS_Tenancy ALTER COLUMN equip nvarchar(20) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT Pic FROM FS_HS_Picture WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("Pic").DefinedSize
		RenameRS.Close
		If ExistTF<>100 Then
			Conn.execute("ALTER TABLE FS_HS_Picture ALTER COLUMN Pic nvarchar(100) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		If Err Then
			Response.Write("SP2-SP5房产数据库升级失败！共"&ChangeCount&"处改动！<br>原因：<font color=red>"&Err.Description&"</font><br />")
			Err.Clear
		Else
			Response.Write "SP2-SP5房产数据库升级成功！共"&ChangeCount&"处改动！<br />"
		End If
	End If
	'房产数据库end_________________________________________________________________

	'人才数据库____________________________________________________________________
	ChangeCount = 0
	If IsExist_SubSys("AP") Then
		If G_IS_SQL_DB = 1 Then
			Set Str_Temp= Conn.execute("Select object_name(cdefault) as 'default' from syscolumns where id=object_id('FS_AP_Consume') and cdefault<>0 and name='LeftCount'")
			If Not Str_Temp.Eof Then
				Conn.execute("ALTER TABLE FS_AP_Consume DROP CONSTRAINT "&Str_Temp(0))
			End If
			'________________________________________________________________
			Set Str_Temp= Conn.execute("Select object_name(cdefault) as 'default' from syscolumns where id=object_id('FS_AP_Payment') and cdefault<>0 and name='LeftCount'")
			If Not Str_Temp.Eof Then
				Conn.execute("ALTER TABLE FS_AP_Payment DROP CONSTRAINT "&Str_Temp(0))
			End If
			'________________________________________________________________
			Set Str_Temp= Conn.execute("Select object_name(cdefault) as 'default' from syscolumns where id=object_id('FS_AP_Consume') and cdefault<>0 and name='SearchUId'")
			If Not Str_Temp.Eof Then
				Conn.execute("ALTER TABLE FS_AP_Consume DROP CONSTRAINT "&Str_Temp(0))
			End If
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT LeftCount FROM FS_AP_Consume WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("LeftCount").Type
		RenameRS.Close
		If ExistTF<>3 Then
			Conn.execute("ALTER TABLE FS_AP_Consume ALTER COLUMN LeftCount INT NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT SearchUId FROM FS_AP_Consume WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("SearchUId").Type
		RenameRS.Close
		If ExistTF<>202 Then
			Conn.execute("ALTER TABLE FS_AP_Consume ALTER COLUMN SearchUId nvarchar(20) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT LeftCount FROM FS_AP_Payment WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("LeftCount").Type
		RenameRS.Close
		If ExistTF<>3 Then
			Conn.execute("ALTER TABLE FS_AP_Payment ALTER COLUMN LeftCount INT NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT SelfAppraise FROM FS_AP_Resume_Intention WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("SelfAppraise").DefinedSize
		RenameRS.Close
		If ExistTF<3000 Then
			If G_IS_SQL_DB=1 Then
				Conn.execute("ALTER TABLE FS_AP_Resume_Intention ALTER COLUMN SelfAppraise nvarchar(3000) NULL")
			Else
				Conn.execute("ALTER TABLE FS_AP_Resume_Intention ALTER COLUMN SelfAppraise ntext NULL")
			End If
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT Description FROM FS_AP_Resume_EducateExp WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("Description").DefinedSize
		RenameRS.Close
		If ExistTF<3000 Then
			If G_IS_SQL_DB=1 Then
				Conn.execute("ALTER TABLE FS_AP_Resume_EducateExp ALTER COLUMN Description nvarchar(3000) NULL")
			Else
				Conn.execute("ALTER TABLE FS_AP_Resume_EducateExp ALTER COLUMN Description ntext NULL")
			End If
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT [Language] FROM FS_AP_Resume_Language WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("Language").DefinedSize
		RenameRS.Close
		If ExistTF<30 Then
			Conn.execute("ALTER TABLE FS_AP_Resume_Language ALTER COLUMN [Language] nvarchar(30) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT Degree FROM FS_AP_Resume_Language WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("Degree").DefinedSize
		RenameRS.Close
		If ExistTF<30 Then
			Conn.execute("ALTER TABLE FS_AP_Resume_Language ALTER COLUMN Degree nvarchar(30) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT Certificate FROM FS_AP_Resume_TrainExp WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("Certificate").DefinedSize
		RenameRS.Close
		If ExistTF<50 Then
			Conn.execute("ALTER TABLE FS_AP_Resume_TrainExp ALTER COLUMN Certificate nvarchar(50) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_AP_Resume_BaseInfo WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("ShenGao") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_AP_Resume_BaseInfo ADD ShenGao INT NULL,XueLi nvarchar(30) NULL,Address nvarchar(80) NULL,HowDay nvarchar(10) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_AP_Consume WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("LeftDateNumber") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_AP_Consume ADD LeftDateNumber INT NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_AP_UserList WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("Phone") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_AP_UserList ADD Phone nvarchar(50) NULL,Email nvarchar(20) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_AP_Payment WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("LeftDateNumber") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_AP_Payment ADD LeftDateNumber INT NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_AP_Job_Public WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("NeedNum") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_AP_Job_Public ADD NeedNum int NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_AP_Job_Public WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("Jlmode") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			If G_IS_SQL_DB=1 Then
				Conn.execute("ALTER TABLE FS_AP_Job_Public ADD Jlmode nvarchar(350) NULL")
			Else
				Conn.execute("ALTER TABLE FS_AP_Job_Public ADD Jlmode ntext NULL")
			End If
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_AP_Job_Public WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("EducateExp") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_AP_Job_Public ADD EducateExp int NULL default 0,Sex int NULL default 0,WorkAge int NULL default 0,Age int NULL default 0,JobType int NULL default 0,OtherJobDes text NULL,MoneyMonth int NULL default 0,FreeMoney int NULL default 0,OtherMoneyDes text NULL,HolleType int NULL default 0")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_AP_Trade WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("TNum") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_AP_Trade ADD TNum INTEGER NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_AP_Job WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("JNum") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_AP_Job ADD JNum INTEGER NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_AP_Province WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("PNum") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_AP_Province ADD PNum INTEGER NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_AP_City WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("CNum") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_AP_City ADD CNum INTEGER NULL")
			ChangeCount = ChangeCount + 1
		End If
		'----------------------------------------------------------------
		'2007-07-31
		If G_IS_SQL_DB = 1 Then
			Set Str_Temp= Conn.execute("Select object_name(cdefault) as 'default' from syscolumns where id=object_id('FS_AP_Resume_BaseInfo') and cdefault<>0 and name='PictureExt'")
			If Not Str_Temp.Eof Then
				Conn.execute("ALTER TABLE FS_AP_Resume_BaseInfo DROP CONSTRAINT "&Str_Temp(0))
			End If
		End If
		RenameRS.Open "SELECT * FROM FS_AP_Resume_BaseInfo WHERE 1=0",Conn,1,3
		ExistTF = RenameRS.Fields("PictureExt").Type
		RenameRS.Close
		If ExistTF <> 202 Then
			Conn.execute("ALTER TABLE FS_AP_Resume_BaseInfo ALTER COLUMN PictureExt nvarchar(250) null")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		If Err Then
			Response.Write("SP2-SP5人才数据库升级失败！共"&ChangeCount&"处改动！<br>原因：<font color=red>"&Err.Description&"</font><br />")
			Err.Clear
		Else
			Response.Write "SP2-SP5人才数据库升级成功！共"&ChangeCount&"处改动！<br />"
		End If
	End If

	'人才数据库end_________________________________________________________________

	'下载数据库____________________________________________________________________
	ChangeCount = 0
	If IsExist_SubSys("DS") Then
		'----------------------------------------------------------------
		RenameRS.Open "SELECT TOP 1 IndexPage FROM FS_DS_SysPara",Conn,1,3
		If Not RenameRS.eof Then
			ExistTF=Replace(RenameRS("IndexPage"),",",".")
		End If
		RenameRS.Close
		Conn.execute("Update FS_DS_SysPara set IndexPage = '"&ExistTF&"'")
		'________________________________________________________________
		If G_IS_SQL_DB = 1 Then
			Set Str_Temp= Conn.execute("Select object_name(cdefault) as 'default' from syscolumns where id=object_id('FS_DS_List') and cdefault<>0 and name='FileExtName'")
			If Not Str_Temp.Eof Then
				Conn.execute("ALTER TABLE FS_DS_List DROP CONSTRAINT "&Str_Temp(0))
			End If
		End If
		RenameRS.Open "SELECT FileExtName FROM FS_DS_List WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("FileExtName").Type
		RenameRS.Close
		If ExistTF<>202 Then
			Conn.execute("ALTER TABLE FS_DS_List ALTER COLUMN FileExtName nvarchar(6) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_DS_List WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("SpeicalID") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_DS_List ADD SpeicalID int NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		ExistTF=False
		If G_IS_SQL_DB = 1 Then
			Set Str_Temp=Conn.Execute("Select * from dbo.sysobjects where id = object_id(N'[FS_DS_SpeciaList]') and OBJECTPROPERTY(id, N'IsUserTable') = 1")
			If Not Str_Temp.Eof Then
				ExistTF=True
			End IF
		Else
			Set Str_Temp=Conn.openSchema(20)
			Str_Temp.MoveFirst
			Do While not Str_Temp.EOF
				if Str_Temp("TABLE_TYPE")="TABLE" then
					if  Str_Temp("TABLE_NAME")="FS_DS_SpeciaList" then
						ExistTF=True
						Exit Do
					End if
				End if
				Str_Temp.MoveNext
			Loop
		End If

		If ExistTF Then
			Conn.Execute("DROP TABLE [FS_DS_SpeciaList]")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		ExistTF=False
		If G_IS_SQL_DB = 1 Then
			Set Str_Temp=Conn.Execute("Select * from dbo.sysobjects where id = object_id(N'[FS_DS_Special]') and OBJECTPROPERTY(id, N'IsUserTable') = 1")
			If Not Str_Temp.Eof Then
				ExistTF=True
			End IF
		Else
			Set Str_Temp=Conn.openSchema(20)
			Str_Temp.MoveFirst
			Do While not Str_Temp.EOF
				if Str_Temp("TABLE_TYPE")="TABLE" then
					if  Str_Temp("TABLE_NAME")="FS_DS_Special" then
						ExistTF=True
						Exit Do
					End if
				End if
				Str_Temp.MoveNext
			Loop
		End If
		If Not ExistTF Then
			Conn.Execute("CREATE TABLE [FS_DS_Special] ([SpecialID] int IDENTITY PRIMARY KEY NOT NULL ,[ParentID] int NULL DEFAULT 0,[SpecialCName] nvarchar(50) NULL ,[SpecialEName] nvarchar(50) NULL ,[SpecialTemplet] nvarchar(250) NULL ,[IsUrl] smallint NULL DEFAULT 0,[Addtime] datetime NULL ,[Domain] nvarchar(20) NULL ,[IsLimited] smallint NULL ,[naviPic] nvarchar(250) NULL ,[isLock] smallint NULL DEFAULT 0,[naviText] ntext NULL ,[FileExtName] nvarchar(6) NULL ,[Savepath] nvarchar(250) NULL)")
			ChangeCount = ChangeCount + 1
		Else
			'________________________________________________________________
			RenameRS.Open "SELECT * FROM FS_DS_Special WHERE 1=0",Conn,1,3
			ExistTF=False
			For key=0 To RenameRS.Fields.count-1
				If LCase(RenameRS.Fields(key).name)=LCase("SpecialTemplet") Then
					ExistTF=True
				End If
			Next
			RenameRS.Close
			If Not ExistTF Then
				Conn.execute("DROP TABLE [FS_DS_Special]")
				Conn.Execute("CREATE TABLE [FS_DS_Special] ([SpecialID] int IDENTITY PRIMARY KEY NOT NULL ,[ParentID] int NULL DEFAULT 0,[SpecialCName] nvarchar(50) NULL ,[SpecialEName] nvarchar(50) NULL ,[SpecialTemplet] nvarchar(250) NULL ,[IsUrl] smallint NULL DEFAULT 0,[Addtime] datetime NULL ,[Domain] nvarchar(20) NULL ,[IsLimited] smallint NULL ,[naviPic] nvarchar(250) NULL ,[isLock] smallint NULL DEFAULT 0,[naviText] ntext NULL ,[FileExtName] nvarchar(6) NULL ,[Savepath] nvarchar(250) NULL)")
				ChangeCount = ChangeCount + 1
			End If
		End If
		'________________________________________________________________
		If G_IS_SQL_DB = 1 Then
			Set Str_Temp= Conn.execute("Select object_name(cdefault) as 'default' from syscolumns where id=object_id('FS_DS_List') and cdefault<>0 and name='ClickNum'")
			If Not Str_Temp.Eof Then
				Conn.execute("ALTER TABLE FS_DS_List DROP CONSTRAINT "&Str_Temp(0))
			End If
			Conn.execute("IF EXISTS (SELECT * FROM sysindexes WHERE name = 'Idx_ClickNum') DROP INDEX FS_DS_List.Idx_ClickNum")
		End If
		RenameRS.Open "SELECT ClickNum FROM FS_DS_List WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("ClickNum").Type
		RenameRS.Close
		If ExistTF<>20 Then
			Conn.execute("ALTER TABLE FS_DS_List ALTER COLUMN ClickNum INTEGER NULL")
			ChangeCount = ChangeCount + 1
		End If
		If G_IS_SQL_DB = 1 Then
			Conn.execute("CREATE NONCLUSTERED INDEX Idx_ClickNum ON FS_DS_List(ClickNum) ON [PRIMARY]")
			Conn.execute("ALTER TABLE dbo.FS_DS_List ADD CONSTRAINT DF_FS_DS_List_ClickNum DEFAULT 0 FOR ClickNum")
		End If
		'________________________________________________________________
		If Err Then
			Response.Write("SP2-SP5下载数据库升级失败！共"&ChangeCount&"处改动！<br />原因：<font color=red>"&Err.Description&"</font><br />")
			Err.Clear
		Else
			Response.Write "SP2-SP5下载数据库升级成功！共"&ChangeCount&"处改动！<br />"
		End If
	End If
	'下载数据库End_________________________________________________________________

	'供求数据库____________________________________________________________________
	ChangeCount = 0
	If IsExist_SubSys("SD") Then
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_SD_Config WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("IndexTemplet") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_SD_Config ADD IndexTemplet nvarchar(200) NULL")
			Conn.execute("UPDATE FS_SD_Config SET IndexTemplet='/"&G_TEMPLETS_DIR&"/Supply/index.htm'")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_SD_Address WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("Templets") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_SD_Address ADD Templets nvarchar(200) NULL")
			Conn.execute("UPDATE FS_SD_Address SET Templets='/"&G_TEMPLETS_DIR&"/Supply/Area.htm'")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_SD_Address WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("ClassOrder") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_SD_Address ADD ClassOrder INT NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		RenameRS.Open "SELECT * FROM FS_SD_Config WHERE 1=0",Conn,1,3
		ExistTF=False
		For key=0 To RenameRS.Fields.count-1
			If LCase(RenameRS.Fields(key).name)=LCase("PublistTemplet") Then
				ExistTF=True
			End If
		Next
		RenameRS.Close
		If Not ExistTF Then
			Conn.execute("ALTER TABLE FS_SD_Config ADD PublistTemplet nvarchar(200) NULL")
			ChangeCount = ChangeCount + 1
		End If
		'________________________________________________________________
		If G_IS_SQL_DB = 1 Then
			Set Str_Temp= Conn.execute("Select object_name(cdefault) as 'default' from syscolumns where id=object_id('FS_SD_News') and cdefault<>0 and name='PubNumber'")
			If Not Str_Temp.Eof Then
				Conn.execute("ALTER TABLE FS_SD_News DROP CONSTRAINT "&Str_Temp(0))
			End If
		End If
		RenameRS.Open "SELECT PubNumber FROM FS_SD_News WHERE 1=0",Conn,1,3
		ExistTF=RenameRS.Fields("PubNumber").Type
		RenameRS.Close
		If ExistTF=2 Then
			Conn.execute("ALTER TABLE FS_SD_News ALTER COLUMN PubNumber INT NULL")
			ChangeCount = ChangeCount + 1
		End If
		Conn.execute("Update FS_SD_News Set PubNumber = 0 Where PubNumber Is Null")
		'----------------------------------------------------------------
		If Err Then
			Response.Write("SP2-SP5供求数据库升级失败！共"&ChangeCount&"处改动！<br>原因：<font color=red>"&Err.Description&"</font><br />")
			Err.Clear
		Else
			Response.Write "SP2-SP5供求数据库升级成功！共"&ChangeCount&"处改动！<br />"
		End If
	End If
	'供求数据库End_________________________________________________________________

	'子系统结束____________________________________________________________________
	'主数据库end___________________________________________________________________
	'归档数据库____________________________________________________________________
	'__________________________________________________________
	ChangeCount = 0
	RenameRS.Open "SELECT * FROM FS_Old_News WHERE 1=0",Old_News_Conn,1,3
	ExistTF=False
	For key=0 To RenameRS.Fields.count-1
		If LCase(RenameRS.Fields(key).name)=LCase("IsAdPic") Then
			ExistTF=True
		End If
	Next
	RenameRS.Close
	If Not ExistTF Then
		Old_News_Conn.execute("ALTER TABLE FS_Old_News ADD IsAdPic nvarchar(1) NULL DEFAULT '0',AdPicWH nvarchar(20) NULL,AdPicLink nvarchar(250) NULL,AdPicAdress nvarchar(250) NULL")
		Old_News_Conn.execute("Update FS_Old_News Set IsAdPic='0'")
		ChangeCount = ChangeCount + 1
	End If

	If Err Then
		Response.Write("SP2-SP5归档数据库升级失败！共"&ChangeCount&"处改动！<br>原因：<font color=red>"&Err.Description&"</font><br />")
		Err.Clear
	Else
		Response.Write "SP2-SP5归档数据库升级成功！共"&ChangeCount&"处改动！<br />"
	End If
	'归档数据库End_________________________________________________________________

	'会员数据库____________________________________________________________________
	'__________________________________________________________
	ChangeCount = 0
	RenameRS.Open "SELECT * FROM FS_ME_SysPara WHERE 1=0",User_Conn,1,3
	ExistTF=False
	For key=0 To RenameRS.Fields.count-1
		If LCase(RenameRS.Fields(key).name)=LCase("DefaultGroupID") Then
			ExistTF=True
		End If
	Next
	RenameRS.Close
	If Not ExistTF Then
		User_Conn.execute("ALTER TABLE FS_ME_SysPara ADD DefaultGroupID INT NULL")
		ChangeCount = ChangeCount + 1
	End If
	'__________________________________________________________
	RenameRS.Open "SELECT * FROM FS_ME_SysPara WHERE 1=0",User_Conn,1,3
	ExistTF=False
	For key=0 To RenameRS.Fields.count-1
		If LCase(RenameRS.Fields(key).name)=LCase("UserSystemName") Then
			ExistTF=True
		End If
	Next
	RenameRS.Close
	If Not ExistTF Then
		User_Conn.execute("ALTER TABLE FS_ME_SysPara ADD UserSystemName varchar(60) NULL")
		ChangeCount = ChangeCount + 1
	End If
	'__________________________________________________________
	RenameRS.Open "SELECT * FROM FS_ME_CorpUser WHERE 1=0",User_Conn,1,3
	ExistTF=False
	For key=0 To RenameRS.Fields.count-1
		If LCase(RenameRS.Fields(key).name)=LCase("DeFen") Then
			ExistTF=True
		End If
	Next
	RenameRS.Close
	If Not ExistTF Then
		User_Conn.execute("ALTER TABLE FS_ME_CorpUser ADD DeFen INT NULL")
		ChangeCount = ChangeCount + 1
	End If
	'__________________________________________________________
	RenameRS.Open "select GroupName,UpfileNum,UpfileSize,GroupDate,GroupPoint,GroupMoney,GroupType,CorpTemplet,LimitInfoNum,GroupDebateNum,JuniorDomain,KeywordsNumber,ProductDiscount,isHtml,BcardNumber,Templetwatermark From FS_ME_Group WHERE GroupName='默认会员组'",User_Conn,1,3
	If RenameRS.Eof Then
		RenameRS.addnew
		RenameRS("GroupName")="默认会员组"
		RenameRS("UpfileNum")=0
		RenameRS("UpfileSize")="2048"
		RenameRS("GroupDate")=0
		RenameRS("GroupPoint")=0
		RenameRS("GroupMoney")=0
		RenameRS("GroupType")=1
		RenameRS("ProductDiscount")=1
		RenameRS("LimitInfoNum")=0
		RenameRS("GroupDebateNum")="0,0"
		RenameRS("JuniorDomain")=1
		RenameRS.update
		RenameRS.Close
		GroupID=User_Conn.execute("SELECT GroupID FROM FS_ME_Group WHERE GroupName='默认会员组'")(0)
		User_Conn.execute("UPDATE FS_ME_SysPara SET DefaultGroupID="&GroupID)
		ChangeCount = ChangeCount + 1
	Else
		RenameRS.Close
	End If
	'__________________________________________________________
	ExistTF=False
	If G_IS_SQL_User_DB = 1 Then
		Set Str_Temp=User_Conn.Execute("Select * from dbo.sysobjects where id = object_id(N'[FS_ME_POP]') and OBJECTPROPERTY(id, N'IsUserTable') = 1")
		If Not Str_Temp.Eof Then
			ExistTF=True
		End IF
	Else
		Set Str_Temp=User_Conn.openSchema(20)
		Str_Temp.MoveFirst
		Do While not Str_Temp.EOF
			if Str_Temp("TABLE_TYPE")="TABLE" then
				if  Str_Temp("TABLE_NAME")="FS_ME_POP" then
					ExistTF=True
					Exit Do
				End if
			End if
			Str_Temp.MoveNext
		Loop
	End If
	If Not ExistTF Then
		User_Conn.Execute("CREATE TABLE [FS_ME_POP] ([ID] int NOT NULL identity primary key,[UserNumber] nvarchar (50) NULL,[InfoId] nvarchar(50) NULL,[SubType] nvarchar(50) NULL,[AddTime] datetime NULL,[isClass] int NULL)")
		ChangeCount = ChangeCount + 1
	End If
	'__________________________________________________________
	RenameRS.Open "SELECT * FROM FS_ME_Order WHERE 1=0",User_Conn,1,3
	ExistTF=False
	For key=0 To RenameRS.Fields.count-1
		If LCase(RenameRS.Fields(key).name)=LCase("IsPay") Then
			ExistTF=True
		End If
	Next
	RenameRS.Close
	If Not ExistTF Then
		User_Conn.execute("ALTER TABLE FS_ME_Order ADD IsPay tinyint NULL default 0")
		ChangeCount = ChangeCount + 1
	End If
	'__________________________________________________________
	RenameRS.Open "SELECT * FROM FS_ME_Order WHERE 1=0",User_Conn,1,3
	ExistTF=False
	For key=0 To RenameRS.Fields.count-1
		If LCase(RenameRS.Fields(key).name)=LCase("IsOnPay") Then
			ExistTF=True
		End If
	Next
	RenameRS.Close
	If Not ExistTF Then
		User_Conn.execute("ALTER TABLE FS_ME_Order ADD IsOnPay tinyint NULL default 0")
		ChangeCount = ChangeCount + 1
	End If
	'__________________________________________________________
	RenameRS.Open "SELECT Content FROM FS_ME_Friends WHERE 1=0",User_Conn,1,3
	ExistTF=RenameRS.Fields("Content").Type
	RenameRS.Close
	If ExistTF<>203 Then
		User_Conn.execute("ALTER TABLE FS_ME_Friends ALTER COLUMN Content ntext NULL")
		ChangeCount = ChangeCount + 1
	End If
	'__________________________________________________________
	RenameRS.Open "SELECT ProductDiscount FROM FS_ME_Group WHERE 1=0",User_Conn,1,3
	ExistTF=RenameRS.Fields("ProductDiscount").Type
	RenameRS.Close
	If ExistTF<>202 Then
		If G_IS_SQL_User_DB = 1 Then
			Set Str_Temp= User_Conn.execute("Select object_name(cdefault) as 'default' from syscolumns where id=object_id('FS_ME_Group') and cdefault<>0 and name='ProductDiscount'")
			If Not Str_Temp.Eof Then
				User_Conn.execute("ALTER TABLE FS_ME_Group DROP CONSTRAINT "&Str_Temp(0))
			End If
			User_Conn.execute("ALTER TABLE FS_ME_Group ALTER COLUMN ProductDiscount nvarchar(50) NULL")
			User_Conn.execute("ALTER TABLE FS_ME_Group ADD CONSTRAINT DF__FS_ME_Gro__Produ__1920BF5C DEFAULT 0 FOR ProductDiscount")
		Else
			User_Conn.execute("ALTER TABLE FS_ME_Group ALTER COLUMN ProductDiscount nvarchar(50) NULL DEFAULT '0'")
		End If
		ChangeCount = ChangeCount + 1
	End If
	'__________________________________________________________
	RenameRS.Open "SELECT AwardPic FROM FS_ME_Award WHERE 1=0",User_Conn,1,3
	ExistTF=RenameRS.Fields("AwardPic").DefinedSize
	RenameRS.Close
	If ExistTF < 200 Then
		User_Conn.execute("ALTER TABLE FS_ME_Award ALTER COLUMN AwardPic nvarchar(200) NULL")
		ChangeCount = ChangeCount + 1
	End If
	'__________________________________________________________
	If Err Then
		Response.Write("SP2-SP5会员数据库升级失败！共"&ChangeCount&"处改动！<br>原因：<font color=red>"&Err.Description&"</font><br />")
		Err.Clear
	Else
		Response.Write "SP2-SP5会员数据库升级成功！共"&ChangeCount&"处改动！<br />"
	End If
	'会员数据库End____________________________________________________________________

	'升级采集数据库开始_______________________________________________________________
	ChangeCount = 0
	RenameRS.Open "SELECT * FROM FS_Site WHERE 1=0",Collect_Conn,1,3
	ExistTF = False
	Str_Fields = ""
	Str_Sql = ""
	For key = 0 To RenameRS.Fields.count - 1
		If LCase(RenameRS.Fields(key).name)=LCase("IsAutoPicNews") Then
			ExistTF=True
		End If
		Str_Fields = Str_Fields & "," & LCase(RenameRS.Fields(key).name)
	Next
	RenameRS.Close
	Set RenameRS=Nothing
	Str_Fields = Str_Fields& ","
	If Not InStr(Str_Fields,LCase(",IsAutoPicNews,"))>0 Then
		If Str_Sql = "" Then
			Str_Sql = "ADD IsAutoPicNews smallint NULL default 0" 
		Else
			Str_Sql = Str_Sql&",IsAutoPicNews smallint NULL default 0" 
		End If
	End If
	If Not InStr(Str_Fields,LCase(",ToClassID,"))>0 Then
		If Str_Sql = "" Then
			Str_Sql = "ADD ToClassID nvarchar(20) NULL" 
		Else
			Str_Sql = Str_Sql&",ToClassID nvarchar(20) NULL" 
		End If
	End If
	If Not InStr(Str_Fields,LCase(",NewsTemplets,"))>0 Then
		If Str_Sql = "" Then
			Str_Sql = "ADD NewsTemplets nvarchar(200) NULL" 
		Else
			Str_Sql = Str_Sql&",NewsTemplets nvarchar(200) NULL" 
		End If
	End If
	If Not InStr(Str_Fields,LCase(",AutoCellectTime,"))>0 Then
		If Str_Sql = "" Then
			Str_Sql = "ADD AutoCellectTime nvarchar(100) NULL" 
		Else
			Str_Sql = Str_Sql&",AutoCellectTime nvarchar(100) NULL" 
		End If
	End If
	If Not InStr(Str_Fields,LCase(",CellectNewNum,"))>0 Then
		If Str_Sql = "" Then
			Str_Sql = "ADD CellectNewNum int NULL" 
		Else
			Str_Sql = Str_Sql&",CellectNewNum int NULL" 
		End If
	End If
	If Not InStr(Str_Fields,LCase(",WebCharset,"))>0 Then
		If Str_Sql = "" Then
			Str_Sql = "ADD WebCharset nvarchar(10) NULL" 
		Else
			Str_Sql = Str_Sql&",WebCharset nvarchar(10) NULL" 
		End If
	End If
	If Not InStr(Str_Fields,LCase(",RulerID,"))>0 Then
		If Str_Sql = "" Then
			Str_Sql = "ADD RulerID nvarchar(50) NULL" 
		Else
			Str_Sql = Str_Sql&",RulerID nvarchar(50) NULL" 
		End If
	End If
	If Not InStr(Str_Fields,LCase(",PicSavePath,"))>0 Then
		If Str_Sql = "" Then
			Str_Sql = "ADD PicSavePath nvarchar(200) NULL" 
		Else
			Str_Sql = Str_Sql&",PicSavePath nvarchar(200) NULL" 
		End If
	End If
	If Not InStr(Str_Fields,LCase(",WaterPrintTF,"))>0 Then
		If Str_Sql = "" Then
			Str_Sql = "ADD WaterPrintTF int NULL default 0" 
		Else
			Str_Sql = Str_Sql&",WaterPrintTF int NULL default 0" 
		End If
	End If
	If Not Str_Sql="" Then
		Collect_Conn.execute("ALTER TABLE FS_Site "&Str_Sql)
		ChangeCount = ChangeCount + 1
	End If
	If Err Then
		Response.Write("SP2-SP5采集数据库升级失败！共"&ChangeCount&"处改动！<br>原因：<font color=red>"&Err.Description&"</font><br />")
		Err.Clear
	Else
		Collect_Conn.execute("Update FS_Site Set IsAutoPicNews = 0")
		Collect_Conn.execute("Update FS_Site Set WaterPrintTF = 0")
		Collect_Conn.execute("Update FS_Site Set PicSavePath = '"&G_UP_FILES_DIR&"/File'")
		Collect_Conn.execute("Update FS_Site Set AutoCellectTime = 'no'")
		Response.Write "SP2-SP5采集数据库升级成功！共"&ChangeCount&"处改动！<br />"
	End If
	'采集数据库升级结束_______________________________________________________________
End Function

Function SP2TOSP5_UpDateLable()
	Dim RsLable,RegExp_Obj,RegExp_result,RegExp_match,LableStr,LableContent,LablePar,ForNum,LableLeft,LableRight,UpCount
	UpCount = 0
	Set RegExp_Obj = New RegExp
	RegExp_Obj.Pattern = "{FS:.*?}"
	RegExp_Obj.IgnoreCase = True'不区分大小写。
	RegExp_Obj.Global = True'全部匹配
	RegExp_Obj.Multiline = True'多行匹配
	Set Conn = Server.CreateObject(G_FS_CONN)
	Conn.Open Str_conn
	Set RsLable = Server.CreateObject(G_FS_RS)
	RsLable.Open "Select ID,LableContent From FS_MF_Lable",Conn,1,3
	While Not RsLable.Eof
		Set RegExp_result = RegExp_Obj.Execute(RsLable("LableContent"))
		For Each RegExp_match In RegExp_result
			LableStr = RegExp_match.value
			If InStr(LableStr,"{FS:NS=ClassList")>0 Or InStr(LableStr,"{FS:NS=SpecialList")>0 Then
				LableContent = Mid(LableStr,8,Len(LableStr) - 8)
				LablePar = Split(LableContent,"┆")
				if Ubound(LablePar)=18 Then
					LableStr = "{FS:NS=ClassList"
					LableLeft = "":LableRight = ""
					For ForNum=Lbound(LablePar)+1 To Ubound(LablePar)
						If ForNum < 14 Then
							LableLeft = LableLeft & "┆" & LablePar(ForNum)
						Else
							LableRight = LableRight & "┆" & LablePar(ForNum)
						End If
					Next
					LableStr = LableStr & LableLeft & "┆分隔数量$5┆分隔样式$" & LableRight & "}"
				End If
				if Ubound(LablePar)=24 Then
					LableStr = "{FS:NS=ClassList"
					LableLeft = "":LableRight = ""
					For ForNum=Lbound(LablePar)+1 To Ubound(LablePar)
						If ForNum < 20 Then
							LableLeft = LableLeft & "┆" & LablePar(ForNum)
						Else
							LableRight = LableRight & "┆" & LablePar(ForNum)
						End If
					Next
					LableStr = LableStr & LableLeft & "┆分隔数量$5┆分隔样式$" & LableRight & "}"
				End If
				If Ubound(LablePar) = 18 or Ubound(LablePar) = 24 Then
					RsLable("LableContent") = Replace(RsLable("LableContent"),RegExp_match.value,LableStr)
					RsLable.Update()
					UpCount = UpCount + 1
				End If
			End If
			If InStr(LableStr,"{FS:NS=FlashFilt")>0 Then
				LableContent = Mid(LableStr,8,Len(LableStr) - 8)
				LablePar = Split(LableContent,"┆")
				If Ubound(LablePar) = 6 Then
					LableStr = Left(LableStr,Len(LableStr) - 1)
					LableStr = LableStr & "┆背景颜色$#FFFFFF}"
					RsLable("LableContent") = Replace(RsLable("LableContent"),RegExp_match.value,LableStr)
					RsLable.Update()
					UpCount = UpCount + 1
				End If
			End If
		Next
		RsLable.Movenext
	Wend
	REsponse.Write("SP2-SP5标签库更新成功。共更新了"&UpCount&"条标签！<br />")
End Function

Function SP6()
	OutPutInfo "SP6"
End Function
Function SP7()
	OutPutInfo "SP7"
End Function
Function UP5()
	If G_IS_SQL_DB=1 Then
		Dim f_FSO,f_File,f_SQL,f_Stream
		Set f_FSO = Server.CreateObject(G_FS_FSO)
		Set f_File = f_FSO.GetFile(Server.MapPath("SQL.txt"))
		Set f_Stream = f_File.OpenAsTextStream(1)
		If Not f_Stream.AtEndOfStream Then 
			f_SQL = f_Stream.ReadAll
		Else 
			f_SQL = ""
		End If
		Set f_Stream = Nothing
		Set f_File = Nothing
		Set f_FSO = Nothing
		if f_SQL <> "" then Conn.Execute(f_SQL)
	Else
		f_SQL = "CREATE TABLE [FS_MF_CustomForm] ([id] autoincrement(1,1) primary key,[formName] Text(50) NOT NULL,"
		f_SQL = f_SQL & "[tableName] Text(50) NOT NULL,[upfileSaveUrl] Text(200) NULL,[upfileSize] Long NULL,[state] Long NULL,"
		f_SQL = f_SQL & "[TimeLimited] Long NULL,[StartTime] Date NULL,[EndTime] Date NULL,[SubmitType] Long NULL,"
		f_SQL = f_SQL & "[GoldFactor] Long NULL,[PointFactor] Long NULL,[UserGroup] Memo NULL,[UserOnce] Long NULL,"
		f_SQL = f_SQL & "[Validate] Long NULL,[remark] Memo NULL,[VerifyLogin] Long NULL,[DataInitStatus] Long NULL)"
		Conn.Execute(f_SQL)
		f_SQL = "CREATE TABLE [FS_MF_CustomForm_Item]([FormItemID] autoincrement(1,1) primary key,[FormID] Long NOT NULL,"
		f_SQL = f_SQL & "[ItemName] Text(50) NOT NULL,[FieldName] Text(50) NULL,[orderby] Long NULL,[State] Long NULL,"
		f_SQL = f_SQL & "[IsNull] Long NULL,[ItemType] Text(50) NULL,[MaxSize] Long NULL,[DefaultValue] Text(50) NULL,"
		f_SQL = f_SQL & "[SelectItem] Memo NULL,[Remark] Memo NULL)"
		Conn.Execute(f_SQL)
	End If
	f_SQL = "UPDATE FS_MF_Config SET MF_Soft_Version='5.0'"
	Conn.execute(f_SQL)
	OutPutInfo "5.0"
End Function

Function UP520091111()
	Dim RsLable,RegExp_Obj,RegExp_result,RegExp_match,LableStr,LableContent,LablePar,ForNum,LableLeft,LableRight,UpCount
	UpCount = 0
	'update FS_MF_Lable set LableContent = replace(LableContent,'输出方式','输出格式')
	Set Conn = Server.CreateObject(G_FS_CONN)
	Conn.Open Str_conn
	Set RsLable = Server.CreateObject(G_FS_RS)
	RsLable.Open "Select ID,LableContent From FS_MF_Lable",Conn,1,3
	While Not RsLable.Eof
		if Instr(RsLable("LableContent"),"输出方式")>0 then
			RsLable("LableContent") = Replace(RsLable("LableContent"),"输出方式","输出格式")
			RsLable.Update()
			UpCount=UpCount+1
		end if
		RsLable.Movenext
	Wend
	Response.Write("5.0 build 20091111标签库更新成功。共更新了"&UpCount&"条标签！<br />")
	Conn.execute("UPDATE FS_MF_Config SET MF_Soft_Version='5.0 build 20091111'")
	OutPutInfo "5.0 build 20091111"
End Function

Function UP520100129()
	if G_IS_SQL_User_DB=0 then
		Dim catalog,tbl
		Set catalog = Server.CreateObject("ADOX.Catalog")
		Set catalog.ActiveConnection = User_Conn
		Set tbl = catalog.Tables("FS_ME_Users")
		tbl.Columns("PassQuestion").Properties("Jet OLEDB:Allow Zero Length") = True
		tbl.Columns("PassQuestion").Properties("Jet OLEDB:Allow Zero Length") = True
		tbl.Columns("PassAnswer").Properties("Jet OLEDB:Allow Zero Length") = True
		tbl.Columns("safeCode").Properties("Jet OLEDB:Allow Zero Length") = True
		tbl.Columns("RealName").Properties("Jet OLEDB:Allow Zero Length") = True
		set tbl=nothing
		Set catalog=nothing
	end if
	Conn.execute("UPDATE FS_MF_Config SET MF_Soft_Version='5.0 build 20100129'")
	OutPutInfo "5.0 build 20100129"
	Response.Cookies("FoosunMFCookies")("FoosunMFVersion") = "5.0 build 20100129"
End Function

Function UP520100507()
	if G_IS_SQL_User_DB=0 then
		Dim catalog,tbl
		Set catalog = Server.CreateObject("ADOX.Catalog")
		Set catalog.ActiveConnection = User_Conn
		Set tbl = catalog.Tables("FS_ME_Users")
		tbl.Columns("NickName").Properties("Jet OLEDB:Allow Zero Length") = True
		set tbl=nothing
		Set catalog=nothing
	end if
	Conn.execute("UPDATE FS_MF_Config SET MF_Soft_Version='5.0 build 20100507'")
	OutPutInfo "5.0 build 20100507"
	Response.Cookies("FoosunMFCookies")("FoosunMFVersion") = "5.0 build 20100507"
End Function
%>
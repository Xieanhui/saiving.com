<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Public.asp" -->
<!--#include file="../../FS_InterFace/CLS_Foosun.asp" -->
<!--#include file="lib/Cls_RefreshJs.asp"-->
<!--#include file="lib/cls_js.asp"-->
<% 
Dim Conn,FS_NS_JS_Obj,FS_NS_JS_Sql,sErrStr
Dim Temp_Admin_Is_Super,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
MF_Default_Conn
'session�ж�
MF_Session_TF 
if not MF_Check_Pop_TF("NS041") then Err_Show
Dim FileID,RsSysModObj,FunClassID,Types,ScrLink,ScrNewsType  ,sRootDir,str_CurrPath,str_CurrPathPic,db_NewsDir,str_SavePath,FunClassName,FileTypeStr


Set RsSysModObj = Conn.Execute("select top 1 NewsDir from FS_NS_SysParam")
if not RsSysModObj.eof then 
	db_NewsDir = RsSysModObj(0)
end if
RsSysModObj.close
if isnull(db_NewsDir) or db_NewsDir="" then 
	response.Redirect("lib/error.asp?ErrCodes=<li>��Ǹ����ϵͳ������ ����ϵͳǰ̨Ŀ¼ ����Ϊ�ա��޷�������</li>&ErrorUrl=../Js_Sys_Add.asp")
	response.End()
end if
str_SavePath = sRootDir &"/"& db_NewsDir 

if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""

if Temp_Admin_Is_Super = 1 then
	str_CurrPathPic = sRootDir &"/"&G_UP_FILES_DIR 
Else
	str_CurrPathPic = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&Temp_Admin_Name,"//","/")
End if

str_CurrPath = replace(sRootDir,"//","/") &"/"& db_NewsDir
str_CurrPath = Replace(str_CurrPath,"'","\'")



If Request.QueryString("FileID")="" or isnull(Request.QueryString("FileID")) then
	response.Redirect("lib/error.asp?ErrCodes=<li>�������ݴ���</li>")
	Response.End
Else
	FileID = NoSqlHack(Request.QueryString("FileID"))
	Types = "System" ''ϵͳJS
	Set RsSysModObj = Conn.Execute("Select * from FS_NS_Sysjs where ID="&FileID&"")
	If RsSysModObj.eof then
		response.Redirect("lib/error.asp?ErrCodes=<li>δ��ѯ����ؼ�¼</li>")
		Response.End
	End IF
	FunClassID = NoSqlHack(RsSysModObj("ClassID"))
	dim tmprs
	set tmprs = Conn.execute("select ClassName from FS_NS_NewsClass where ClassID='"&FunClassID&"'")
	if not tmprs.eof then FunClassName = tmprs(0)
	tmprs.close
	ScrLink = RsSysModObj("MoreContent")
	ScrNewsType = RsSysModObj("NewsType")
End IF
Function ClassList()
	Dim Rs,SelectStr
	Set Rs = Conn.Execute("select ClassID,ClassCName from FS_NS_NewsClass where ParentID = '0' and DelFlag=0 order by AddTime desc")
	do while Not Rs.Eof
		If Cstr(FunClassID) = Cstr(Rs("ClassID")) then
			SelectStr = " selected"
		Else
			SelectStr = ""
		End If
		ClassList = ClassList & "<option value="""&Rs("ClassID")&""""& SelectStr & ">" & Rs("ClassCName") & chr(10) & chr(13)
		ClassList = ClassList & ChildClassList(Rs("ClassID"),"")
		Rs.MoveNext	
	loop
	Rs.Close
	Set Rs = Nothing
End Function
Function ChildClassList(ClassID,Temp)
	Dim TempRs,TempStr,SelectStrs
	Set TempRs = Conn.Execute("Select ClassID,ClassCName,ChildNum from FS_NS_NewsClass where ParentID = '" & NoSqlHack(ClassID) & "' and DelFlag=0 order by AddTime desc ")
	TempStr = Temp & " - "
	do while Not TempRs.Eof
		If Cstr(FunClassID) = Cstr(TempRs("ClassID")) then
			SelectStrs = " selected"
		Else
			SelectStrs = ""
		End If
		if TempRs("ChildNum") = 0 then
			ChildClassList = ChildClassList & "<option value="""&TempRs("ClassID")&"""" & SelectStrs & ">" & TempStr & TempRs("ClassCName") & "</option>"& chr(10) & chr(13)
		else
			ChildClassList = ChildClassList & "<option value="""&TempRs("ClassID")&"""" & SelectStrs & ">" & TempStr & TempRs("ClassCName") & "</option>"& chr(10) & chr(13)
		end if
		ChildClassList = ChildClassList & ChildClassList(TempRs("ClassID"),TempStr)
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title>ϵͳJS����___Powered by foosun Inc.</title>
	<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
	<script type="text/JavaScript" src="js/Public.js"></script>
	<script type="text/JavaScript" src="../../FS_Inc/Prototype.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" scroll="yes" oncontextmenu="return true;">
	<form action="?action=add&FileID=<%=FileID%>" method="post" name="ClassJSForm">
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr class="hback">
			<td class="xingmu">
				ϵͳJS����
			</td>
		</tr>
		<tr class="hback">
			<td class="hback">
				<a href="Js_Sys_manage.asp?Act=View">������ҳ</a>&nbsp;|&nbsp; <a href="Js_Sys_Add.asp">����</a> &nbsp;|&nbsp; <a href="Js_Sys_manage.asp?Act=Search">��ѯ</a>
			</td>
		</tr>
	</table>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr class="hback">
			<td width="15%" height="26">
				&nbsp;&nbsp;&nbsp;&nbsp;��������
			</td>
			<td width="35%">
				<input name="FileCName" type="text" id="FileCName" style="width: 90%" value="<%=RsSysModObj("FileCName")%>">
			</td>
			<td width="15%">
				&nbsp;&nbsp;&nbsp;&nbsp;�ļ�����
			</td>
			<td width="35%">
				<input name="FileName" type="text" id="FileName" readonly="" style="width: 90%" value="<%=RsSysModObj("FileName")%>">
			</td>
		</tr>
		<tr class="hback">
			<td width="15%" height="26">
				&nbsp;&nbsp;&nbsp;&nbsp;ѡ����Ŀ
			</td>
			<td width="35%">
				<input name="strClassName" type="text" id="strClassName" style="width: 50%" value="<%=FunClassName%>" readonly="">
				<input name="strClassID" type="hidden" id="strClassID" value="<%=FunClassID%>">
				<input type="button" name="Submit" value="ѡ����Ŀ" onclick="SelectClass();">
			</td>
			<td colspan="2">
				��ѡ�����������
			</td>
		</tr>
		<tr class="hback">
			<td height="26">
				&nbsp;&nbsp;&nbsp;&nbsp;��������
			</td>
			<td>
				<select name="NewsType" style="width: 90%" onchange="ChooseNewsType(this.options[this.selectedIndex].value);">
					<option value="RecNews" <%if RsSysModObj("NewsType") = "RecNews" then Response.Write("selected")%>>�Ƽ�����</option>
					<option value="MarqueeNews" <%if RsSysModObj("NewsType") = "MarqueeNews" then Response.Write("selected")%>>��������</option>
					<option value="SBSNews" <%if RsSysModObj("NewsType") = "SBSNews" then Response.Write("selected")%>>��������</option>
					<option value="PicNews" <%if RsSysModObj("NewsType") = "PicNews" then Response.Write("selected")%>>ͼƬ����</option>
					<option value="NewNews" <%if RsSysModObj("NewsType") = "NewNews" then Response.Write("selected")%>>��������</option>
					<option value="HotNews" <%if RsSysModObj("NewsType") = "HotNews" then Response.Write("selected")%>>�ȵ�����</option>
					<option value="WordNews" <%if RsSysModObj("NewsType") = "WordNews" then Response.Write("selected")%>>��������</option>
					<option value="TitleNews" <%if RsSysModObj("NewsType") = "TitleNews" then Response.Write("selected")%>>��������</option>
					<option value="ProclaimNews" <%if RsSysModObj("NewsType") = "ProclaimNews" then Response.Write("selected")%>>��������</option>
				</select>
			</td>
			<td>
				&nbsp;&nbsp;���������
			</td>
			<td>
				<input name="setHitsValue" type="text" onchange="if(/\D/.test(this.value)){alert('ֻ����������');this.value='';}" value="<%=RsSysModObj("MarSpeed")%>" size="5" maxlength="4">
				�����ȵ�����,����������ڶ���
			</td>
		</tr>
		<tr class="hback">
			<td height="26">
				&nbsp;&nbsp;&nbsp;��������
			</td>
			<td>
				<select name="MoreContent" id="MoreContent" style="width: 90%" <%If Types = "System" then Response.Write("disabled")%> onchange="ChooseLink(this.options[this.selectedIndex].value);">
					<option value="1" <%If RsSysModObj("MoreContent")=1 then Response.Write("selected")%>>��</option>
					<option value="0" <%If RsSysModObj("MoreContent")=0 then Response.Write("selected")%>>��</option>
				</select>
			</td>
			<td>
				&nbsp;&nbsp;&nbsp;&nbsp;��������
			</td>
			<td>
				<input name="LinkWord" type="text" id="LinkWord" style="width: 90%" value="<%=RsSysModObj("LinkWord")%>" <%If Types = "System" then Response.Write("disabled")%>>
			</td>
		</tr>
		<tr class="hback">
			<td height="26">
				&nbsp;&nbsp;&nbsp;&nbsp;��������
			</td>
			<td>
				<input name="NewsNum" type="text" id="NewsNum" style="width: 90%" value="<%=RsSysModObj("NewsNum")%>">
			</td>
			<td>
				&nbsp;&nbsp;&nbsp;&nbsp;ÿ������
			</td>
			<td>
				<input name="RowNum" type="text" id="RowNum" style="width: 90%" value="<%=RsSysModObj("RowNum")%>">
			</td>
		</tr>
		<tr class="hback">
			<td height="26">
				&nbsp;&nbsp;&nbsp;&nbsp;������ʽ
			</td>
			<td>
				<input name="LinkCSS" type="text" id="LinkCSS" style="width: 90%" value="<%=RsSysModObj("LinkCSS")%>" <%If Types = "System" then Response.Write("disabled")%>>
			</td>
			<td>
				&nbsp;&nbsp;&nbsp;&nbsp;��������
			</td>
			<td>
				<input name="TitleNum" type="text" id="TitleNum" style="width: 90%" value="<%=RsSysModObj("TitleNum")%>">
			</td>
		</tr>
		<tr class="hback">
			<td height="26">
				&nbsp;&nbsp;&nbsp;&nbsp;ͼƬ���
			</td>
			<td>
				<input name="PicWidth" type="text" id="PicWidth" style="width: 90%" value="<%=RsSysModObj("PicWidth")%>">
			</td>
			<td>
				&nbsp;&nbsp;&nbsp;&nbsp;������ʽ
			</td>
			<td>
				<input name="TitleCSS" type="text" id="TitleCSS" style="width: 90%" value="<%=RsSysModObj("TitleCSS")%>">
			</td>
		</tr>
		<tr class="hback">
			<td height="26">
				&nbsp;&nbsp;&nbsp;&nbsp;ͼƬ�߶�
			</td>
			<td>
				<input name="PicHeight" type="text" id="PicHeight" style="width: 90%" value="<%=RsSysModObj("PicHeight")%>">
			</td>
			<td>
				&nbsp;&nbsp;&nbsp;&nbsp;�����о�
			</td>
			<td>
				<input name="RowSpace" type="text" id="RowSpace" style="width: 90%" value="<%=RsSysModObj("RowSpace")%>">
			</td>
		</tr>
		<tr class="hback">
			<td height="26">
				&nbsp;&nbsp;&nbsp;&nbsp;�����ٶ�
			</td>
			<td>
				<input name="MarSpeed" type="text" id="MarSpeed" style="width: 90%" value="<%=RsSysModObj("MarSpeed")%>">
			</td>
			<td>
				&nbsp;&nbsp;&nbsp;&nbsp;��������
			</td>
			<td>
				<select name="MarDirection" id="MarDirection" style="width: 90%">
					<option value="up" <%If RsSysModObj("MarDirection")="up" then Response.Write("selected")%>>����</option>
					<option value="down" <%If RsSysModObj("MarDirection")="down" then Response.Write("selected")%>>����</option>
					<option value="left" <%If RsSysModObj("MarDirection")="left" then Response.Write("selected")%>>����</option>
					<option value="right" <%If RsSysModObj("MarDirection")="right" then Response.Write("selected")%>>����</option>
				</select>
			</td>
		</tr>
		<tr class="hback">
			<td height="26">
				&nbsp;&nbsp;&nbsp;&nbsp;������
			</td>
			<td>
				<input name="MarWidth" type="text" id="MarWidth" style="width: 90%" value="<%=RsSysModObj("MarWidth")%>">
			</td>
			<td>
				&nbsp;&nbsp;&nbsp;&nbsp;����߶�
			</td>
			<td>
				<input name="MarHeight" type="text" id="MarHeight" style="width: 90%" value="<%=RsSysModObj("MarHeight")%>">
			</td>
		</tr>
		<tr class="hback">
			<td height="26">
				&nbsp;&nbsp;&nbsp;&nbsp;��ʾ����
			</td>
			<td>
				<select name="ShowTitle" id="ShowTitle" style="width: 90%">
					<option value="1" <%If RsSysModObj("ShowTitle")=1 then Response.Write("selected")%>>��</option>
					<option value="0" <%If RsSysModObj("ShowTitle")=0 then Response.Write("selected")%>>��</option>
				</select>
			</td>
			<td>
				&nbsp;&nbsp;&nbsp;&nbsp;�¿�����
			</td>
			<td>
				<select name="OpenMode" id="OpenMode" style="width: 90%">
					<option value="1" <%If RsSysModObj("OpenMode")=1 then Response.Write("selected")%>>��</option>
					<option value="0" <%If RsSysModObj("OpenMode")=0 then Response.Write("selected")%>>��</option>
				</select>
			</td>
		</tr>
		<tr class="hback">
			<td height="26">
				&nbsp;&nbsp;&nbsp;&nbsp;����ͼƬ
			</td>
			<td>
				<input name="NaviPic" type="text" id="NaviPic" style="width: 60%" value="<%=RsSysModObj("NaviPic")%>">
				<input type="button" name="bnt_ChoosePic_naviPic" value="ѡ��ͼƬ" onclick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPathPic %>',500,300,window,document.ClassJSForm.NaviPic);">
			</td>
			<td>
				&nbsp;&nbsp;&nbsp;&nbsp;�м�ͼƬ
			</td>
			<td>
				<input name="RowBetween" type="text" id="RowBetween" style="width: 52%" value="<%=RsSysModObj("RowBetween")%>">
				<input type="button" name="bnt_ChoosePic_rowBettween" value="ѡ��ͼƬ" onclick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPathPic %>',500,300,window,document.ClassJSForm.RowBetween);">
			</td>
		</tr>
		<tr class="hback">
			<td height="26">
				&nbsp;&nbsp;&nbsp;&nbsp;��������
			</td>
			<td>
				<select name="DateType" id="DateType" style="width: 90%">
					<option value="0">������������</option>
					<option value="1" <%if RsSysModObj("DateType") = "1" then Response.Write("selected") end if%>>
						<%=Year(Now)&"-"&Month(Now)&"-"&Day(Now)%></option>
					<option value="2" <%if RsSysModObj("DateType") = "2" then Response.Write("selected") end if%>>
						<%=Year(Now)&"."&Month(Now)&"."&Day(Now)%></option>
					<option value="3" <%if RsSysModObj("DateType") = "3" then Response.Write("selected") end if%>>
						<%=Year(Now)&"/"&Month(Now)&"/"&Day(Now)%></option>
					<option value="4" <%if RsSysModObj("DateType") = "4" then Response.Write("selected") end if%>>
						<%=Month(Now)&"/"&Day(Now)&"/"&Year(Now)%></option>
					<option value="5" <%if RsSysModObj("DateType") = "5" then Response.Write("selected") end if%>>
						<%=Day(Now)&"/"&Month(Now)&"/"&Year(Now)%></option>
					<option value="6" <%if RsSysModObj("DateType") = "6" then Response.Write("selected") end if%>>
						<%=Month(Now)&"-"&Day(Now)&"-"&Year(Now)%></option>
					<option value="7" <%if RsSysModObj("DateType") = "7" then Response.Write("selected") end if%>>
						<%=Month(Now)&"."&Day(Now)&"."&Year(Now)%></option>
					<option value="8" <%if RsSysModObj("DateType") = "8" then Response.Write("selected") end if%>>
						<%=Month(Now)&"-"&Day(Now)%></option>
					<option value="9" <%if RsSysModObj("DateType") = "9" then Response.Write("selected") end if%>>
						<%=Month(Now)&"/"&Day(Now)%></option>
					<option value="10" <%if RsSysModObj("DateType") = "10" then Response.Write("selected") end if%>>
						<%=Month(Now)&"."&Day(Now)%></option>
					<option value="11" <%if RsSysModObj("DateType") = "11" then Response.Write("selected") end if%>>
						<%=Month(Now)&"��"&Day(Now)&"��"%></option>
					<option value="12" <%if RsSysModObj("DateType") = "12" then Response.Write("selected") end if%>>
						<%=day(Now)&"��"&Hour(Now)&"ʱ"%></option>
					<option value="13" <%if RsSysModObj("DateType") = "13" then Response.Write("selected") end if%>>
						<%=day(Now)&"��"&Hour(Now)&"��"%></option>
					<option value="14" <%if RsSysModObj("DateType") = "14" then Response.Write("selected") end if%>>
						<%=Hour(Now)&"ʱ"&Minute(Now)&"��"%></option>
					<option value="15" <%if RsSysModObj("DateType") = "15" then Response.Write("selected") end if%>>
						<%=Hour(Now)&":"&Minute(Now)%></option>
					<option value="16" <%if RsSysModObj("DateType") = "16" then Response.Write("selected") end if%>>
						<%=Year(Now)&"��"&Month(Now)&"��"&Day(Now)&"��"%></option>
				</select>
			</td>
			<td>
				&nbsp;&nbsp;&nbsp;&nbsp;����·��
			</td>
			<td>
				<input name="SaveFilePath" type="text" id="SaveFilePath" style="width: 52%" value="<%=RsSysModObj("FileSavePath")%>">
				<input type="button" name="Submit4" value="ѡ��·��" onclick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPathFrame.asp?CurrPath=<%= str_SavePath %>',300,250,window,document.ClassJSForm.SaveFilePath);document.ClassJSForm.SaveFilePath.focus();">
			</td>
		</tr>
		<tr class="hback">
			<td height="26">
				&nbsp;&nbsp;&nbsp;&nbsp;������ʽ
			</td>
			<td>
				<input name="DateCSS" type="text" id="DateCSS" style="width: 90%" value="<%=RsSysModObj("DateCSS")%>">
			</td>
			<td>
				&nbsp;&nbsp;&nbsp;&nbsp;��ʾ��Ŀ
			</td>
			<td>
				<select name="ClassName" id="ClassName" style="width: 90%">
					<option value="1" <%If RsSysModObj("ClassName")=1 then Response.Write("selected")%>>��ʾ</option>
					<option value="0" <%If RsSysModObj("ClassName")=0 then Response.Write("selected")%>>����ʾ</option>
				</select>
			</td>
		</tr>
		<tr class="hback">
			<td height="26">
				&nbsp;&nbsp;&nbsp;&nbsp;��������
			</td>
			<td>
				<select name="SonClass" id="SonClass" style="width: 90%">
					<option value="1" <%If RsSysModObj("SonClass")=1 then Response.Write("selected")%>>��</option>
					<option value="0" <%If RsSysModObj("SonClass")=0 then Response.Write("selected")%>>��</option>
				</select>
			</td>
			<td>
				&nbsp;&nbsp;&nbsp;&nbsp;�����Ҷ���
			</td>
			<td>
				<select name="RightDate" id="RightDate" style="width: 90%">
					<option value="1" <%If RsSysModObj("RightDate")=1 then Response.Write("selected")%>>��</option>
					<option value="0" <%If RsSysModObj("RightDate")=0 then Response.Write("selected")%>>��</option>
				</select>
			</td>
		</tr>
		<tr class="hback">
			<td colspan="10" align="center">
				<input type="submit" name="submit" value=" ���� ">&nbsp;&nbsp;
				<input type="reset" name="reset" value=" ���� ">
			</td>
		</tr>
	</table>
	</form>
	<p>
</body>
</html>
<%
RsSysModObj.close
set RsSysModObj=nothing
If Request.QueryString("action") = "add" then

	Dim ResultStr,TempFsoObj,FileNameStr
	FileNameStr = NoSqlHack(Trim(request.Form("FileName")))
	If Request.Form("FileCName")="" then
		sErrStr = sErrStr & "<li>�ļ��������Ʋ���Ϊ��</li>" '�ļ����Ʋ���Ϊ�ջ����зǷ��ַ�
	End If
	If FileNameStr="" then
		sErrStr = sErrStr & "<li>�ļ����Ʋ���Ϊ��</li>" '�ļ����Ʋ���Ϊ�ջ����зǷ��ַ�
	End If
	If Request.Form("SaveFilePath")="" then
		sErrStr = sErrStr & "<li>δָ���ļ�����·��</li>" '�ļ����Ʋ���Ϊ�ջ����зǷ��ַ�
	End If
	If isnumeric(Request.Form("NewsNum"))=false then
		sErrStr = sErrStr & "<li>����������������Ϊ������</li>" '����������������Ϊ������
	End If
	If isnumeric(Request.Form("TitleNum"))=false then
		sErrStr = sErrStr & "<li>������������Ϊ������</li>"
	End If
	If isnumeric(Request.Form("RowNum"))=false then
		sErrStr = sErrStr & "<li>����ÿ��������������Ϊ������</li>"
	End If
	If isnumeric(Request.Form("RowSpace"))=false then
		sErrStr = sErrStr & "<li>�����о����Ϊ������</li>"
	End IF
	If Types="Class" and Request.Form("ClassID")="" then
		sErrStr = sErrStr & "<li>��ĿID�������ݴ���</li>"  '��ĿID�������ݴ���
	End If
	If Request.Form("NewsType")="PicNews" or Request.Form("NewsType")="FilterNews" then
		If isnumeric(Request.Form("PicWidth"))=false or isnumeric(Request.Form("PicHeight"))=false then
			sErrStr = sErrStr & "<li>ͼƬ������Ϊ������</li>"
		End If
	End If
	If Request.Form("MoreContent")=1 then
		If Request.Form("LinkWord")="" then
			sErrStr = sErrStr & "<li>��������������</li>"
		End If
	End If
	If Request.Form("NewsType")="MarqueeNews" or Request.Form("NewsType")="ProclaimNews" then
		If isnumeric(Request.Form("MarSpeed"))=false then
			sErrStr = sErrStr & "<li>���Ź����ٶȱ���Ϊ������</li>"
		End If
	End If
	if sErrStr<>"" then 
		response.Redirect("lib/error.asp?ErrCodes="&server.URLEncode(sErrStr))
		response.End()
	end if
	Dim ClassJsAddObj,RsClassSql
	Set ClassJsAddObj = Server.CreateObject(G_FS_RS)
	RsClassSql = "Select * from FS_NS_Sysjs where ID="&CintStr(FileID)&""
	ClassJsAddObj.Open RsClassSql,Conn,1,3
	ClassJsAddObj("FileCName") = NoSqlHack(Request.Form("FileCName"))
	ClassJsAddObj("ClassID") = Replacestr(NoSqlHack(Request.Form("strClassID")),":0,else:"&Request.Form("strClassID"))
	ClassJsAddObj("NewsType") = NoSqlHack(Request.Form("NewsType"))
	ClassJsAddObj("NewsNum") = CintStr(Request.Form("NewsNum"))
	ClassJsAddObj("TitleNum") = CintStr(Request.Form("TitleNum"))
	ClassJsAddObj("TitleCSS") = NoSqlHack(Cstr(Request.Form("TitleCSS")))
	ClassJsAddObj("RowNum") = CintStr(Request.Form("RowNum"))
	ClassJsAddObj("NaviPic") = NoSqlHack(Request.Form("NaviPic"))
	ClassJsAddObj("RowBetween") = NoSqlHack(Request.Form("RowBetween"))
	ClassJsAddObj("FileSavePath") = Cstr(Request.Form("SaveFilePath"))
	ClassJsAddObj("RowSpace") = CintStr(Request.Form("RowSpace"))
	ClassJsAddObj("DateType") = CintStr(Request.Form("DateType"))
	ClassJsAddObj("DateCSS") = Cstr(Request.Form("DateCSS"))
	If Request.Form("ClassName")<>0 then
		ClassJsAddObj("ClassName") = 1
	Else
		ClassJsAddObj("ClassName") = 0
	End If
	If Request.Form("SonClass")<>0 then
		ClassJsAddObj("SonClass") = 1
	Else
		ClassJsAddObj("SonClass") = 0
	End If
	If Request.Form("RightDate")<>0 then
		ClassJsAddObj("RightDate") = 1
	Else
		ClassJsAddObj("RightDate") = 0
	End If
	If Request.Form("MoreContent")<>"" and isnull(Request.Form("MoreContent"))=false then
		ClassJsAddObj("MoreContent") = NoSqlHack(Request.Form("MoreContent"))
	End if
	If Request.Form("MoreContent")<>0 then
		ClassJsAddObj("LinkWord") = NoSqlHack(Request.Form("LinkWord"))
		ClassJsAddObj("LinkCSS") = NoSqlHack(Request.Form("LinkCSS"))
	End If
	If Request.Form("PicWidth")<>"" and isnull(Request.Form("PicWidth"))=false then
		ClassJsAddObj("PicWidth") = CintStr(Request.Form("PicWidth"))
	End If
	If Request.Form("PicHeight")<>"" and isnull(Request.Form("PicHeight"))=false then
		ClassJsAddObj("PicHeight") = CintStr(Request.Form("PicHeight"))
	End If
	If Request.Form("MarSpeed")<>"" and isnull(Request.Form("MarSpeed"))=false then
		ClassJsAddObj("MarSpeed") = CintStr(Request.Form("MarSpeed"))
	End If
	If Request.Form("MarDirection")<>"" and isnull(Request.Form("MarDirection"))=false then
		ClassJsAddObj("MarDirection") = Cstr(Request.Form("MarDirection"))
	End If
	If Request.Form("ShowTitle")<>"" and isnull(Request.Form("ShowTitle"))=false then
		ClassJsAddObj("ShowTitle") = NoSqlHack(Request.Form("ShowTitle"))
	End If
	If Request.Form("OpenMode")<>1 then
		ClassJsAddObj("OpenMode") = 0
	Else
		ClassJsAddObj("OpenMode") = 1
	End If
	If Request.Form("MarWidth")<>"" and isnull(Request.Form("MarWidth"))=false then
		ClassJsAddObj("MarWidth") = NoSqlHack(Request.Form("MarWidth"))
	End If
	If Request.Form("MarHeight")<>"" and isnull(Request.Form("MarHeight"))=false then
		ClassJsAddObj("MarHeight") = NoSqlHack(Request.Form("MarHeight"))
	End If
	if request.Form("NewsType")="HotNews" then 
		if isnumeric(request.Form("setHitsValue")) then 
			ClassJsAddObj("MarSpeed") = NoSqlHack(request.Form("setHitsValue"))
		else
			ClassJsAddObj("MarSpeed") = 1
		end if	
	end if
	'------2007-01-18    
	If Trim(Request.Form("strClassName")) = "" Or Trim(Request.Form("strClassID")) = "" Then
		FileTypeStr = 2
	Else
		FileTypeStr = 1
	End If		
	ClassJsAddObj("FileType") = FileTypeStr
	'--------------------
	ClassJsAddObj.Update
	ClassJsAddObj.Close
	Set ClassJsAddObj = Nothing
	ResultStr = CreateSysJS(FileNameStr)
	if ResultStr = true then
	response.Redirect("lib/Success.asp?ErrorUrl="&server.URLEncode( "../Js_Sys_manage.asp?Act=View" )&"&ErrCodes=<li>��ϲ���޸ĳɹ���</li>")
	else
		response.Redirect("lib/error.asp?ErrCodes=<li>"&ResultStr&"</li>&ErrorUrl=../Js_Sys_manage.asp")
	end if
	Response.End
End If
Conn.Close
Set Conn = Nothing
%>
<script>
	var TempLink = '<% = ScrLink %>';
	var TempNewsType = '<% = ScrNewsType %>';
	function ChooseLink(Link) {
		if (Link == 1) {
			document.ClassJSForm.LinkWord.disabled = false;
			document.ClassJSForm.LinkCSS.disabled = false;
		}
		else {
			document.ClassJSForm.LinkWord.disabled = true;
			document.ClassJSForm.LinkCSS.disabled = true;
		}
	}

	function ChooseNewsType(NewsType) {
		if ((NewsType == 'MarqueeNews') || (NewsType == 'ProclaimNews')) {
			document.ClassJSForm.MarSpeed.disabled = false;
			document.ClassJSForm.MarDirection.disabled = false;
			document.ClassJSForm.MarWidth.disabled = false;
			document.ClassJSForm.MarHeight.disabled = false;
		}
		else {
			document.ClassJSForm.MarSpeed.disabled = true;
			document.ClassJSForm.MarDirection.disabled = true;
			document.ClassJSForm.MarWidth.disabled = true;
			document.ClassJSForm.MarHeight.disabled = true;
		}
		if ((NewsType == 'PicNews') || (NewsType == 'FilterNews')) {
			document.ClassJSForm.PicWidth.disabled = false;
			document.ClassJSForm.PicHeight.disabled = false;
			document.ClassJSForm.ShowTitle.disabled = false;
		}
		else {
			document.ClassJSForm.PicWidth.disabled = true;
			document.ClassJSForm.PicHeight.disabled = true;
			document.ClassJSForm.ShowTitle.disabled = true;
		}
	}
	ChooseLink(TempLink);
	ChooseNewsType(TempNewsType);

	function SelectClass() {
		var ReturnValue = '', TempArray = new Array();
		ReturnValue = OpenWindow('lib/SelectClassFrame.asp', 400, 300, window);
		try {
			$("strClassID").value = ReturnValue[0][0];
			$("strClassName").value = ReturnValue[1][0];
		}
		catch (ex) { }
	}

 
</script>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->

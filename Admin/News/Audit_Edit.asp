<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="lib/cls_main.asp"-->
<%
Dim Conn,s_News,User_Conn,newsInfo_Rs,newsClass_Rs,sql_cmd,sql_class_cmd,contID,NewsTitle,CurtTitle,Content,Keywords,Author,classID,ClassName
DIm Fs_News,str_CurrPath,sRootDir
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if
if Not MF_Session_TF() then Top_Go_To_Login_Page
'MF_Check_Pop_TF�����������ļ�������+����
'if Not MF_Check_Pop_TF("NS_News_01") then Top_Go_To_Error_Page 

MF_User_Conn
MF_Default_Conn
MF_Session_TF 
Set Fs_News=New Cls_News
contID=CintStr(Request.QueryString("contid"))

Set newsInfo_Rs=Server.CreateObject(G_FS_RS)
Set newsClass_Rs=Server.CreateObject(G_FS_RS)
'��ø���Ļ�������
sql_cmd="Select ContTitle,SubTitle,ContContent,KeyWords,UserNumber,MainID,AuditTF from FS_ME_InfoContribution where contid="&contID
newsInfo_Rs.open sql_cmd,User_Conn,1,1
if not  newsInfo_Rs.eof then
	NewsTitle=newsInfo_Rs("ContTitle")
	CurtTitle=newsInfo_Rs("SubTitle")
	Content=newsInfo_Rs("ContContent")
	Keywords=newsInfo_Rs("KeyWords")
	Author=newsInfo_Rs("UserNumber")
	classID=newsInfo_Rs("MainID")
ENd if
'��÷����������� 
sql_class_cmd="Select ClassName,Classid from  FS_NS_NewsClass where id="&classID
newsClass_Rs.open sql_class_cmd,Conn,1,1 
dim newsClassID
if not newsClass_Rs.eof then
	ClassName=newsClass_Rs("ClassName")
	newsClassID=newsClass_Rs("Classid")
End if
if err.number="" then
	Response.Redirect("lib/error.asp?ErrCodes=<li>�����쳣���뷵��</li>")
	Response.End()
End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>CMS5.0</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" src="../../FS_Inc/Prototype.js"></script>
<script language="JavaScript" src="js/Public.js"></script>
<script language="JavaScript" src="js/CheckJs.js"></script>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Get_Domain.asp"></script>
<script language="JavaScript" type="text/javascript" src="../../Editor/FS_scripts/editor.js"></script>
<body class="hback">
<form name="auditEdit_Form" id="auditEdit_Form" action="Constr_Action.asp?act=editaudit" method="post">
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td colspan="2" class="xingmu">Ͷ�����ǰ�༭</td>
		</tr>
		<tr class="hback">
			<td width="20%" align="right">���ű��⣺</td>
			<td width="80%" aling="left">
				<input name="txt_newsTitle" type="text" size="50" id="txt_newsTitle" value="<%=NewsTitle%>">
				<input name="hid_contID" type="hidden" id="hid_contID" value="<%=ContID%>">
				<span id="span_newsTitle"></span> </td>
		</tr>
		<tr class="hback">
			<td align="right">���Ÿ����⣺</td>
			<td aling="left">
				<input name="txt_curtTitle" type="text" size="50" id="txt_curtTitle" value="<%=CurtTitle%>">
			</td>
		</tr>
		<tr class="hback">
			<td align="right">���ģ�</td>
			<td aling="left">
				<!--�༭����ʼ-->
				<iframe id='NewsContent' src='../Editer/AdminEditer.asp?id=txt_content' frameborder=0 scrolling=no width='100%' height='280'></iframe>
				<input type="hidden" name="txt_content" value="<% = HandleEditorContent(Content)%>">
                <!--�༭������-->
				<span id="span_content"></span></td>
		</tr>
		<tr class="hback">
			<td align="right">�ؼ��ʣ�</td>
			<td aling="left">
				<input name="txt_keywords" type="text" size="50" id="txt_keywords" value="<%=Keywords%>" onKeyUp="ReplaceDot('txt_keywords')">
			</td>
		</tr>
		<tr class="hback">
			<td align="right">����������Ŀ��</td>
			<td aling="left">
				<input name="txt_classCName" type="text" size="50" id="txt_classCName" value="<%=ClassName%>">
				<input name="hid_ClassID" type="hidden" id="hid_ClassID" value="<% = newsClassID %>">
				<input type="button" name="Submit" value="ѡ����Ŀ"   onClick="SelectClass();">
				<span id="span_ClassID"></span></td>
		</tr>
		<tr class="hback">
			<td align="right">ѡ��ģ�壺</td>
			<td aling="left">
				<input name="NewsTemplet" type="text" id="NewsTemplet" size="50" value="<%=Replace("/"& G_TEMPLETS_DIR &"/NewsClass/news.htm","//","/")%>">
				<input name="Submit532" type="button" id="Submit53" value="ѡ��ģ��"  onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=Replace("/"&G_VIRTUAL_ROOT_DIR&"/"& G_TEMPLETS_DIR,"//","/") %>',400,300,window,document.auditEdit_Form.NewsTemplet);document.auditEdit_Form.NewsTemplet.focus();">
				<span id="span_NewsTemplet"></span></td>
		</tr>
		<tr class="hback">
			<td align="right">���ߣ�</td>
			<td aling="left"><a href="../../user/ShowUser.asp?UserNumber=<%=Author%>" title="����鿴���û�����" target="_blank"><%=Fs_News.GetUserName(Author)%></a>
				<input name="hid_Author" type="hidden" size="50" readonly="true" id="hid_Author" value="<%=Author%>">
			</td>
		</tr>
		<tr class="hback">
			<td align="right">&nbsp;</td>
			<td aling="left">
				<%if newsInfo_Rs("AuditTF")<>1 then%>
				<input name="sub_button" type="button" id="sub_button" value="���" onClick="checkRight(this.form)">
				&nbsp;&nbsp;
				<%Else%>
				<input name="sub_button" type="button" id="sub_button" value="�����"  disabled="disabled">
				<%End if%>
				<input type="reset" name="Submit" value="����">
				<input type="button" name="bbb" onClick="window.history.back()" value="����">
			</td>
		</tr>
	</table>
</form>
</body>
<script language="JavaScript">
function SelectClass()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('lib/SelectClassFrame.asp',400,300,window);
	try {
		$("hid_ClassID").value= ReturnValue[0][0];
		$("txt_classCName").value= ReturnValue[1][0];
	}
	catch (ex) { }
}
function checkRight(FormObj)
{
	var flag1=isEmpty("txt_newsTitle","span_newsTitle");
	var flag2=isEmpty("txt_content","span_content");
	var flag3=isEmpty("hid_ClassID","span_ClassID");
	var flag4=isEmpty("NewsTemplet","span_NewsTemplet");
	if(flag1&&flag2&&flag3&&flag4)
	{
		if (frames["NewsContent"].g_currmode!='EDIT') {alert('����ģʽ���޷����棬���л������ģʽ');return false;}
		FormObj.txt_content.value=frames["NewsContent"].GetNewsContentArray();
		FormObj.submit();
	}
}
</script>
</html>







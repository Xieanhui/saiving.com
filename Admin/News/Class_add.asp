<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%
Dim Conn,User_Conn
MF_Default_Conn   
MF_User_Conn
MF_Session_TF 
Dim Fs_news,NS_ClassNameValure,sRootDir,strShowErr,str_newsDir
set Fs_news = new Cls_News
MF_GetUserGroupID
Fs_News.GetSysParam()
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
if Fs_news.newsDir<>"" then str_newsDir = "/"+Fs_news.newsDir else str_newsDir=""
Dim obj_Class_Rs,ClassID,str_ClassKeywords,str_Classdescription,str_currpath
Dim lng_OrderID,str_ClassName,str_ClassEName_add,str_ParentID,str_Templet,str_NewsTemplet,str_Domain,lng_AdminID,int_RefreshNumber
Dim  lng_GroupID,lng_PointNumber,flt_Money,str_FileExtName,dtm_Addtime,int_isConstr,int_IsURL,str_UrlAddress,lng_Oldtime,int_isShow
Dim str_ClassNaviContent,str_ClassNaviPic,lng_DefineID,int_NewsCheck,int_AddNewsType,tmp_fileExtName,str_SavePath,str_FileSaveType,int_isConstrDel,str_GetParentID,IsAdPic,AdPicWH,AdPicLink,AdPicAdress
ClassID = NoSqlHack(Trim(Request.QueryString("ClassID")))
Select Case fs_news.fileExtName
		Case 0
			tmp_fileExtName ="html"
		Case 1
			tmp_fileExtName ="htm"
		Case 2
			tmp_fileExtName ="shtml"
		Case 3
			tmp_fileExtName ="shtm"
		Case 4
			tmp_fileExtName ="asp"
End Select
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
if Request.QueryString("Action")="add" then
	if not Get_SubPop_TF(ClassID,"NS016","NS","class") then Err_Show
	str_Templet = Replace("//"&G_TEMPLETS_DIR&"/NewsClass/class.htm","//","/")
	str_NewsTemplet = Replace("/"&G_TEMPLETS_DIR&"/NewsClass/news.htm","//","/")
	dtm_Addtime = now
	lng_AdminID = session("Admin_Name")
	lng_OrderID = 10
	lng_PointNumber = ""
	flt_Money = ""
	str_SavePath = Replace(str_newsDir,"//","/")
	str_UrlAddress = "http://"
	str_FileExtName = tmp_fileExtName
	int_isShow = 1
	int_RefreshNumber = 0
	int_AddNewsType=Fs_news.addNewsType
	str_FileSaveType = Fs_news.ClassSaveType
	if NoSqlHack(ClassID)<>"" then
		str_GetParentID = ClassID
	Else
		str_GetParentID = "0"
	End if
	lng_Oldtime = 180
	if ClassID<>"" then
		Dim obj_IsUrlTF_Rs
		Set obj_IsUrlTF_Rs = server.CreateObject(G_FS_RS)
		obj_IsUrlTF_Rs.Open "Select IsUrl from FS_NS_NewsClass where ClassID='"& NoSqlHack(ClassID) &"' order by id desc",Conn,1,1
		if not obj_IsUrlTF_Rs.eof then
			if obj_IsUrlTF_Rs(0) = 1 then
				strShowErr = "<li>�ⲿ��Ŀ�����������</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			elseif obj_IsUrlTF_Rs(0) = 2 then
				strShowErr = "<li>��ҳ��Ŀ�����������</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
		end if
	End if
Elseif Request.QueryString("Action")="edit" then
	if not Get_SubPop_TF(ClassID,"NS017","NS","class") then Err_Show
	Set obj_Class_Rs = server.CreateObject(G_FS_RS)
	obj_Class_Rs.open "select ClassID,OrderID,ClassName,ClassEName,ParentID,Templet,NewsTemplet,[Domain],RefreshNumber,ClassAdmin,isPop,FileExtName,Addtime,isConstr,IsURL,UrlAddress,Oldtime,isShow,ClassNaviContent,ClassNaviPic,DefineID,NewsCheck,AddNewsType,SavePath,FileSaveType,isConstrDel,ClassKeywords,Classdescription,IsAdPic,AdPicWH,AdPicLink,AdPicAdress From FS_NS_NewsClass where ClassID = '"& NoSqlHack(ClassID) &"'",Conn,1,3
	if  not obj_Class_Rs.eof then
		if obj_Class_Rs("isPop")=1 then
			Dim obj_tmppop_rs
			set obj_tmppop_rs = Conn.execute("select GroupName,PointNumber,FS_Money,InfoID,PopType,isClass From FS_MF_POP where InfoID='"& NoSqlHack(obj_Class_Rs("ClassID")) &"' and isClass=1 and PopType='NS'")
			if obj_tmppop_rs.eof then
					lng_GroupID = ""
					lng_PointNumber=""
					flt_Money = ""
					obj_tmppop_rs.close:set obj_tmppop_rs = nothing
			Else
					lng_GroupID = obj_tmppop_rs("GroupName")
					if obj_tmppop_rs("PointNumber") = 0 or isnull(trim(obj_tmppop_rs("PointNumber"))) then:lng_PointNumber="" else:lng_PointNumber=obj_tmppop_rs("PointNumber"):end if
					if obj_tmppop_rs("FS_Money") = 0 or isnull(trim(obj_tmppop_rs("FS_Money"))) then:flt_Money="" else:flt_Money=obj_tmppop_rs("FS_Money"):end if
					obj_tmppop_rs.close:set obj_tmppop_rs = nothing
			End if
		Else
					lng_GroupID = ""
					lng_PointNumber=""
					flt_Money = ""
		End if
		lng_OrderID = obj_Class_Rs("OrderID")
		str_ClassName = obj_Class_Rs("ClassName")
		str_ClassEName_add = obj_Class_Rs("ClassEName")
		str_ParentID = obj_Class_Rs("ParentID")
		str_GetParentID = obj_Class_Rs("ParentID")
		str_Templet = obj_Class_Rs("Templet")
		str_NewsTemplet = obj_Class_Rs("NewsTemplet")
		str_Domain = obj_Class_Rs("Domain")
		lng_AdminID = obj_Class_Rs("ClassAdmin")
		int_RefreshNumber = obj_Class_Rs("RefreshNumber")
		str_FileExtName = obj_Class_Rs("FileExtName")
		dtm_Addtime = obj_Class_Rs("Addtime")
		int_isConstr = obj_Class_Rs("isConstr")
		int_IsURL = obj_Class_Rs("IsURL")
		str_UrlAddress = obj_Class_Rs("UrlAddress")
		lng_Oldtime = obj_Class_Rs("Oldtime")
		int_isShow = obj_Class_Rs("isShow")
		str_ClassNaviContent = obj_Class_Rs("ClassNaviContent")
		str_ClassNaviPic = obj_Class_Rs("ClassNaviPic")
		lng_DefineID = obj_Class_Rs("DefineID")
		int_NewsCheck = obj_Class_Rs("NewsCheck")
		int_AddNewsType = obj_Class_Rs("AddNewsType")
		str_SavePath = obj_Class_Rs("SavePath")
		str_FileSaveType = obj_Class_Rs("FileSaveType")
		int_isConstrDel = obj_Class_Rs("isConstrDel")
		str_ClassKeywords  = obj_Class_Rs("ClassKeywords")
		str_Classdescription  = obj_Class_Rs("Classdescription")
		IsAdPic = obj_Class_Rs("IsAdPic")
		AdPicWH = obj_Class_Rs("AdPicWH")
		AdPicLink = obj_Class_Rs("AdPicLink")
		AdPicAdress = obj_Class_Rs("AdPicAdress")
		obj_Class_Rs.close
		set  obj_Class_Rs = nothing
	Else
		strShowErr = "<li>����Ĳ���</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
Else
		strShowErr = "<li>����Ĳ���</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
End if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ŀ����___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function insertType(f_classType) { 
	if (f_classType=="1")
	{
		OutUrl.style.display='';
		InUrl.style.display='none'
		PageClass.style.display='none'
	}
	else if(f_classType=="0")
	{
		OutUrl.style.display='none';
		InUrl.style.display=''
		PageClass.style.display=''
	}
	else if(f_classType=="2")
	{
		OutUrl.style.display='none';
		InUrl.style.display='none'
		PageClass.style.display=''
	}
}
//-->
</script>
<script language="JavaScript" src="../../FS_Inc/GetLettersByChinese.js"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
</head>
<body>
<form name="ClassForm" method="post" action="Class_Save.asp">
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr class="hback">
      <td class="xingmu">��Ŀ����<a href="../../help?Lable=NS_Class_add" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
    </tr>
    <tr>
      <td height="18" class="hback"><div align="left"><a href="Class_Manage.asp">������ҳ</a>��<a href="Class_add.asp?ClassID=&Action=add">��Ӹ���Ŀ</a>��<a href="Class_Action.asp?Action=one">һ����Ŀ����</a>��<a href="Class_Action.asp?Action=n">N����Ŀ����</a>��<a href="Class_Action.asp?Action=reset"   onClick="{if(confirm('ȷ�ϸ�λ������Ŀ��\n\n���ѡ��ȷ�������е���Ŀ������Ϊһ������!!')){return true;}return false;}">��λ������Ŀ</a>��<a href="Class_Action.asp?Action=unite">��Ŀ�ϲ�</a>��<a href="Class_Action.asp?Action=allmove">��Ŀת��</a>��<a href="Class_Action.asp?Action=clearClass"  onClick="{if(confirm('ȷ�����������Ŀ���������\n\n���ѡ��ȷ��,���е���Ŀ�����Ž����ŵ�����վ��!!')){return true;}return false;}">ɾ��������Ŀ</a> <a href="../../help?Lable=NS_Class_add_1" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></div></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr class="hback">
      <td colspan="3" class="xingmu">�����Ŀ</td>
    </tr>
    <tr>
      <td width="23%" height="29" class="hback"><div align="right">��Ŀ���ͣ� </div></td>
      <td width="14%" class="hback"><input name="isUrl" type="radio" id="isUrl"  onClick="insertType(0)" value="0" <%if int_IsURL = 0 then response.Write("Checked")%>>
        ��ͨ��Ŀ<br />
		<input name="isUrl" type="radio" id="isUrl"  onClick="insertType(1)" value="1" <%if int_IsURL = 1 then response.Write("Checked")%>>
        �ⲿ��Ŀ<br />
		<input name="isUrl" type="radio" id="isUrl"  onClick="insertType(2)" value="2" <%if int_IsURL = 2 then response.Write("Checked")%>>
        ��ҳ��Ŀ
        <input name="ClassID" type="hidden" id="ClassID" value="<% = ClassID %>">
        <input name="ParentID" type="hidden" id="ParentID" value="<% = str_GetParentID %>" readonly>
      </td>
      <td width="63%" class="hback"><span class="tx">��ͨ��Ŀ������ϸ�Ĳ������á������������Ŀ������<br>��ҳ��ĿΪ��ҳ�治���������Ŀ�����ţ��繫˾���ܣ���ϵ���ǵ�<br>
        �ⲿ��Ŀָ���ӵ���ϵͳ����ĵ�ַ�С�������Ŀ׼�����ӵ���վ�е�����ϵͳʱ����ʹ�����ַ�ʽ���������ⲿ��Ŀ��������ţ�Ҳ�����������Ŀ��</span></td>
    </tr>
    <tr>
      <td width="23%" class="hback"><div align="right">��Ŀ�������ƣ�</div></td>
      <td colspan="2" class="hback"><input name="ClassName" type="text" id="ClassName" size="40" maxlength="100" value="<% = str_ClassName%>" onBlur="value=value.replace(/[\s]/g,'');<% if Request.QueryString("Action")="add" then %>SetClassEName(value,document.ClassForm.ClassEName);<% end if %>" onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[\s]/g,''));">
        <span class="tx"> *3-100���ַ�</span></td>
    </tr>
    <tr>
      <td height="22" class="hback"><div align="right">������ĿID��</div></td>
      <td height="22" colspan="2" class="hback"><%
	  Dim str_Parentvalue
	  if Request.QueryString("Action") = "add" then
	  		if Not isnull(Trim(ClassID)) then
				str_Parentvalue = Fs_news.GetClassName(ClassID)
			Else
				str_Parentvalue = "����Ŀ" 
			End if
	 Elseif Request.QueryString("Action") = "edit" then
	 		if str_ParentID = "0" then
				str_Parentvalue = "����Ŀ"
			Else		
				str_Parentvalue = Fs_news.GetClassName(str_ParentID)
			End if
	 End if
	  %>
        <input name="ParentIDs" type="text" id="ParentIDs" value="<% = str_Parentvalue %>" size="40" readonly>
        <span class="tx"> *0Ϊ����Ŀ</span></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table" style="<%if int_IsURL = 1 then%>display:<%else%>display:none<%end if%>" id="OutUrl">
    <tr >
      <td width="23%" height="19" class="hback"><div align="right">�ⲿ��ַ��</div></td>
      <td width="78%" height="19" class="hback"><input name="UrlAddress" type="text" id="UrlAddress" size="40" maxlength="250" value="<% = str_UrlAddress%>">
        <span class="tx"> *</span> ���250���ַ�</td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table"  id="PageClass" style="<%if int_IsURL = 2 OR int_IsURL = 0 then%>display:<%else%>display:none<%end if%>">
    <tr style="dispay:">
      <td width="23%" class="hback"><div align="right">��ĿӢ�����ƣ�</div></td>
      <td width="77%" class="hback"><input name="ClassEName" type="text" id="ClassEName" size="40" maxlength="50" value="<% =str_ClassEName_add%>" <%if Request.QueryString("Action")="edit" then response.Write("Readonly")%>>
        <span class="tx"> *<br>
        3-50���ַ�,��������ĸ�����֣��л��ߣ��»���,@,.��һ��ȷ��,�������޸�</span></td>
    </tr>
    <tr style="dispay:">
      <td class="hback"><div align="right">��Ŀģ���ַ��</div></td>
      <td class="hback"><input name="Templet" type="text" id="Templet" value="<% = str_Templet %>" size="50" maxlength="250" readonly>
        <input type="button" name="Submit" value="ѡ��ģ��" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir %>/<% = G_TEMPLETS_DIR %>',400,300,window,document.ClassForm.Templet);document.ClassForm.Templet.focus();">
        <span class="tx"> *250���ַ�</span></td>
    </tr>
    <tr style="dispay:">
      <td class="hback"><div align="right">����������</div></td>
      <td class="hback"><input name="Domain" type="text" id="Domain" size="40" maxlength="150" value="<% = str_Domain %>">
        <span class="tx">150���ַ�,����д��ȷ������</span></td>
    </tr>
    <tr style="dispay:">
      <td class="hback"><div align="right">����Ա��</div></td>
      <td class="hback"><SELECT name="ClassAdmin" id="ClassAdmin">
          <%
			Dim obj_AdminList_Rs
			set obj_AdminList_Rs = Conn.Execute("Select Admin_Name,Admin_Real_Name from FS_MF_Admin Where Admin_Parent_Admin='"&Temp_Admin_Name&"' or Admin_Name='"&Temp_Admin_Name&"' order by ID asc")
			If not obj_AdminList_Rs.eof Then
				if lng_AdminID = obj_AdminList_Rs("Admin_Name") then
					Response.Write "<OPTION value=""" & obj_AdminList_Rs("Admin_Name") & """ selected>����Ա�ʺţ�" & obj_AdminList_Rs("Admin_Name") & "������Ա������" & obj_AdminList_Rs("Admin_Real_Name") & "</OPTION>"
				Else
					Response.Write "<OPTION value=""" & obj_AdminList_Rs("Admin_Name") & """>����Ա�ʺţ�" & obj_AdminList_Rs("Admin_Name") & "������Ա������" & obj_AdminList_Rs("Admin_Real_Name") & "</OPTION>"
				End if
				obj_AdminList_Rs.MoveNext
			End If
			Do while not obj_AdminList_Rs.eof
				if lng_AdminID = obj_AdminList_Rs("Admin_Name") then
					Response.Write "<OPTION value=""" & obj_AdminList_Rs("Admin_Name") & """ selected>����Ա�ʺţ�" & obj_AdminList_Rs("Admin_Name") & "������Ա������" & obj_AdminList_Rs("Admin_Real_Name") & "</OPTION>"
				Else
					Response.Write "<OPTION value=""" & obj_AdminList_Rs("Admin_Name") & """>����Ա�ʺţ�" & obj_AdminList_Rs("Admin_Name") & "������Ա������" & obj_AdminList_Rs("Admin_Real_Name") & "</OPTION>"
				End if
				obj_AdminList_Rs.Movenext
			Loop
			obj_AdminList_Rs.Close
			Set obj_AdminList_Rs = Nothing
			%>
        </SELECT>
        <span class="tx">����Ա����ѡ��<a href="../../help?Lable=NS_Class_Admin" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></span>�� </td>
    </tr>
    <tr style="dispay:">
      <td class="hback"><div align="right">��̬�ļ���չ����</div></td>
      <td class="hback"><select name="FileExtName" id="FileExtName">
          <option value="html" <% if  Trim(str_FileExtName) = "html"  then response.Write("selected")%>>.html</option>
          <option value="htm" <% if  Trim(str_FileExtName) = "htm"  then response.Write("selected")%>>.htm</option>
          <option value="shtml" <% if  Trim(str_FileExtName) = "shtml"  then response.Write("selected")%>>.shtml</option>
          <option value="shtm" <% if  Trim(str_FileExtName)= "shtm"  then response.Write("selected")%>>.shtm</option>
          <option value="asp" <% if  Trim(str_FileExtName) = "asp"  then response.Write("selected")%>>.asp</option>
        </select>
        <span class="tx"> *�����Ҫ�Ķ�Ȩ�ޣ���������Ϊ.asp</span></td>
    </tr>
    <tr style="dispay:">
      <td class="hback"><div align="right">��Ŀ��ҳ����ģʽ��</div></td>
      <td class="hback"><select name="FileSaveType" id="FileSaveType">
          <option value="0" <%if str_FileSaveType = 0 Then response.Write("selected")%>>��ĿӢ��/index.html</option>
          <option value="1" <%if str_FileSaveType = 1 Then response.Write("selected")%>>��ĿӢ��/��ĿӢ��.html</option>
          <option value="2" <%if str_FileSaveType = 2 Then response.Write("selected")%>>��ĿӢ��.html</option>
        </select>
        <span class="tx"> *</span></td>
    </tr>
    <tr style="dispay:">
      <td class="hback"><div align="right">��Ŀ����·����</div></td>
      <td class="hback"><input name="SavePath" type="text" id="SavePath" value="<%=str_SavePath%>" size="40" maxlength="255" readonly>
        <%if Request.QueryString("Action")="add" then%>
        <INPUT type="button"  name="Submit4" value="ѡ��·��" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPathFrame.asp?CurrPath=<%= Replace(sRootDir & str_newsDir,"//","/")%>',320,280,window,document.ClassForm.SavePath);document.ClassForm.SavePath.focus();">
        <%End if%>
        <span class="tx"> *<br>
        һ����д�������޸�;���ѡ���˶�������������������Ŀ¼�ĸ�Ŀ¼</span></td>
    </tr>
    <tr style="dispay:">
      <td class="hback"><div align="right">��ĿMETA�ؼ��֣�</div></td>
      <td class="hback"><textarea name="ClassKeywords" style="width:80%" rows="5" id="ClassKeywords"><% = str_ClassKeywords %>
</textarea>
        <span class="tx"><br>
        ���200���ַ�,�û������������������������Ŀ����������������¼�Ļ���</span></td>
    </tr>
    <tr style="dispay:">
      <td class="hback"><div align="right">��ĿMETA������</div></td>
      <td class="hback"><textarea name="Classdescription"  style="width:80%" rows="5" id="Classdescription"><% = str_Classdescription %>
</textarea>
        <span class="tx"><br>
        ���200���ַ�,�û������������������������Ŀ����������������¼�Ļ���</span></td>
    </tr>
  </table>
  
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table"  id="InUrl" style="<%if int_IsURL = 0 then%>display:<%else%>display:none<%end if%>">
    <tr id="InUrl3" style="dispay:">
      <td class="hback"><div align="right">����ģ���ַ��</div></td>
      <td class="hback"><input name="NewsTemplet" type="text" id="NewsTemplet" value="<% = str_NewsTemplet %>" size="50" maxlength="250" readonly>
        <input type="button" name="Submit2"  value="ѡ��ģ��" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir %>/<% = G_TEMPLETS_DIR %>',400,300,window,document.ClassForm.NewsTemplet);document.ClassForm.NewsTemplet.focus();">
        <span class="tx"> *250���ַ�</span></td>
    </tr>
    <tr id="InUrl6" style="dispay:">
      <td class="hback"><div align="right">�����Ա�飺</div></td>
      <td class="hback"><input name="BrowPop"  id="BrowPop" type="text" value="<% = lng_GroupID %>" onMouseOver="this.title=this.value;" readonly>
        <select name="selectPop" id="selectPop" style="overflow:hidden;" onChange="ChooseExeName();">
          <option value="" selected>ѡ���Ա��</option>
          <option value="del" style="color:red;">���</option>
          <% = MF_GetUserGroupID %>
        </select>
        ��Ҫ����
        <input name="PointNumber" type="text" id="PointNumber" size="8" maxlength="5" value="<% = lng_PointNumber %>"  onChange="ChooseExeName();"  onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
        ��Ҫ���
        <input name="Money" type="text" id="Money" size="8" maxlength="5" value="<% = flt_Money %>"  onChange="ChooseExeName();"  onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')"></td>
    </tr>
    <tr id="InUrl8" style="dispay:">
      <td class="hback"><div align="right">�Ƿ�����Ͷ�壺</div></td>
      <td class="hback"><input name="isConstr" type="checkbox" id="isConstr" value="1" <%if int_isConstr = 1 Then response.Write("checked")%>>
        ��ԱͶ���Ƿ�����ɾ��
        <input name="isConstrDel" type="checkbox" id="isConstrDel" value="1"  <%if int_isConstrDel = 1 Then response.Write("checked")%>>
        ��</td>
    </tr>
    <tr id="InUrl9" style="dispay:">
      <td class="hback"><div align="right">�������ģʽ��</div></td>
      <td class="hback"><input name="AddNewsType" type="checkbox" id="AddNewsType" value="0"  <%if int_AddNewsType = 0 Then response.Write("checked")%>>
        ���ģʽ</td>
    </tr>
    <tr style="dispay:">
      <td class="hback"><div align="right">�Զ���ѡ��</div></td>
      <td class="hback"><select name="DefineID" id="DefineID">
          <option value="0" selected>ѡ���Զ������</option>
          <% = Fs_News.GetDefineClassId%>
        </select>
      </td>
    </tr>
    <tr style="dispay:">
      <td class="hback"><div align="right">������Ҫ��ˣ�</div></td>
      <td class="hback"><input name="NewsCheck" type="checkbox" id="NewsCheck" value="1" <%if int_NewsCheck = 1 then response.Write("checked")%>>
        ��Ҫ���</td>
    </tr>
    <tr style="dispay:">
      <td class="hback"><div align="right"> �������鵵��</div></td>
      <td class="hback"><input  name="Oldtime" type="text" id="Oldtime" value="<% = lng_Oldtime %>" size="40"  onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')"></td>
    </tr>
    <tr style="dispay:">
      <td class="hback"><div align="right">�������¶�������Ϣ</div></td>
      <td class="hback"><input  name="RefreshNumber" type="text" id="RefreshNumber" value="<% = int_RefreshNumber %>" size="40"  onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
        <span class="tx">���Ϊ0������</span></td>
    </tr>
	    <tr>
      <td height="2" class="hback"><div align="right">�Ƿ�����������ʾ���л���</div></td>
      <td height="2" class="hback"><input name="IsAdPic" type="checkbox" id="IsAdPic" value="1" onClick="javascript:ShowAdpicInfo();" <% If Cint(IsAdPic)=1 or Cint(IsAdPic)=2  Then Response.Write("checked")%>></td>
    </tr>
		<tr id="selectAp" style="display:none" class="hback">
		<td class="hback"></td>
		    <td  colspan="2" class="hback" align="left"> ͼƬ���л�
		
                <input id="Checkbox1" name="Checkbox1" type="checkbox" onClick="javascript:ShowAdpicInfo1();" <% If Cint(IsAdPic)=1 Then Response.Write("checked") %>> &nbsp;&nbsp;&nbsp;���ֻ��л�
		     
                <input id="Checkbox2" name="Checkbox2" type="checkbox"  onClick="javascript:ShowAdpicInfo2();" <% If  Cint(IsAdPic)=2 Then Response.Write("checked") %> value="1">
          </td>
             		<td class="hback"></td>

		</tr>
    <tr id="Adpic" style="display:none" class="hback"><td colspan="2">
      <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr>
          <td width="23%" height="2" class="hback"><div align="right">���л��������ã�</div></td>
          <td width="78%" height="2" class="hback"><input name="AdPicWH" type="text" id="AdPicWH" size="20" maxlength="20" value="<%If AdPicWH="" Or IsNull(AdPicWH) Then:Response.Write("100,100,1,400"):Else:Response.Write(AdPicWH):End If%>">
(���,�߶�,��(1)��(0),����λ������������ǰ������(������)����100,100,1,400)</td>
        </tr>
        <tr>
          <td height="5" class="hback"><div align="right">ͼƬ��ַ��</div></td>
          <td height="5" class="hback"><input name="AdPicAdress" type="text" id="AdPicAdress"  size="20" maxlength="250" readonly value="<%=AdPicAdress%>">
            <input name="SelectAdPic" type="button" id="SelectAdPic" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath%>',500,300,window,document.ClassForm.AdPicAdress);"  value="ѡ��ͼƬ��FLASH">
            ���ӵ�ַ
            <input name="AdPicLink" type="text" id="AdPicLink"  size="36" maxlength="250" value="<%=AdPicLink%>"></td>
        </tr>
      </table></td></tr>
	  
	 <tr id="wzPic" style="display:none" class="hback">
		         <td colspan="4">
		        <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		                  <tr>
                <td width="23%" height="2" class="hback"><div align="right">���л�����</div></td>
                <td width="78%" height="2" class="hback"><input name="AdPicWHw" type="text" id="Text2" size="20" maxlength="20" value="<%if Cint(IsAdPic)=2 then response.write(AdPicWH) end if%>">
                (����λ������������ǰ������(������)����100</td>
              </tr>
              <tr>
	         <td class="hback" align="right">���л�����
		     </td>
		     <td class="hback" colspan="3"  align="left">
                <textarea id="IsApicArea" name="IsApicArea" cols="80" rows="10"><%
				if Cint(IsAdPic)=2 then response.write(AdPicLink) end if
				%></textarea>
		      </td>
		     </tr>
		     </table>
		    </td>
	</tr> 
  </table>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td width="23%" height="21" class="hback"><div align="right">������ڣ�</div></td>
      <td width="78%" height="21" class="hback"><input  name="Addtime" type="text" id="Addtime" value="<% = dtm_Addtime %>" size="40" readonly>
        <input name="SelectDate" type="button" id="SelectDate" value="ѡ��ʱ��" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.Addtime);" ></td>
    </tr>
    <tr>
      <td height="10" class="hback"><div align="right">
          <DIV align="right">�Ƿ��ڵ�������ʾ��</DIV>
        </div></td>
      <td height="10" class="hback"><input name="isShow" type="checkbox" id="isShow" value="1" <% if int_isShow = 1 then response.Write("checked") %>></td>
    </tr>
    <tr>
      <td height="21" class="hback"><div align="right">���ݷ�ҳ��ǩ[FS:PAGE]<br>
				  <a href="javascript:void(0);" onClick="InsertHTML('[FS:PAGE]','NewsContent')"><span class="tx">�����ҳ��ǩ</span></a><br>
      </div></td>
      <td height="21" class="hback">
	  			<!--�༭����ʼ-->
				<iframe id='NewsContent' src='../Editer/AdminEditer.asp?id=ClassNaviContent' frameborder=0 scrolling=no width='100%' height='380'></iframe>
				<input type="hidden" name="ClassNaviContent" value="<% = HandleEditorContent(str_ClassNaviContent) %>">
                <!--�༭������-->
            </td>
    </tr>
    <tr>
      <td height="21" class="hback"><div align="right">��Ŀ����ͼƬ��</div></td>
      <td height="21" class="hback"><input name="ClassNaviPic" type="text" id="ClassNaviPic" value="<% = str_ClassNaviPic%>" size="40">
        <input type="button" name="PPPChoose"  value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath%>',500,300,window,document.ClassForm.ClassNaviPic);"></td>
    </tr>
    <tr>
      <td height="10" class="hback"><div align="right">����Ȩ�أ�</div></td>
      <td height="10" class="hback"><input name="OrderID" type="text" id="OrderID" value="<% = lng_OrderID%>" size="40"  onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')"></td>
    </tr>
    <tr>
      <td height="21" class="hback"><div align="right"></div></td>
      <td height="21" class="hback"><input type="button" name="Submit4222" value="������Ŀ" onClick="{if(confirm('ȷ�ϱ���������Ŀ��Ϣ��?')){SubmitFun();return true;}return false;}">
        <input type="reset" name="Submit5222" value="����">
        <input name="str_add" type="hidden" id="str_add" value="<% = Request.QueryString("Action")%>"></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
'����䱻SamJun��2009��5��26�գ������Ŀ���ֻ��л�ʱע���
'If Cint(IsAdPic)=1 Then Response.Write("<script language=""javascript"">document.all.Adpic.style.display="""";< /script>")
set Fs_news = nothing
%>
<SCRIPT language="JavaScript">
function SetClassEName(Str,Obj)
{
	Obj.value=ConvertToLetters(Str,1);
}

function SubmitFun()
{
	if (frames["NewsContent"].g_currmode!='EDIT') {alert('����ģʽ���޷����棬���л������ģʽ');return false;}
	document.ClassForm.ClassNaviContent.value=frames["NewsContent"].GetNewsContentArray();
	document.ClassForm.submit();
}

var DocumentReadyTF=false;
function document.onreadystatechange()
{
	ChooseExeName();
}
function ChooseExeName()
{
  var ObjValue = document.ClassForm.selectPop.options[document.ClassForm.selectPop.selectedIndex].value;
  if (ObjValue!='')
  {
	if (document.ClassForm.BrowPop.value=='')
		document.ClassForm.BrowPop.value = ObjValue;
	else if(document.ClassForm.BrowPop.value.indexOf(ObjValue)==-1)
		document.ClassForm.BrowPop.value = document.ClassForm.BrowPop.value+","+ObjValue;
	if (ObjValue=='del')
  		document.ClassForm.BrowPop.value ='';
  }
   CheckNumber(document.ClassForm.PointNumber,"����۵�ֵ");
  if (document.ClassForm.PointNumber.value>32767||document.ClassForm.PointNumber.value<-32768||document.ClassForm.PointNumber.value=='0')
	{
		alert('����۵�ֵ��������Χ��\n���32767���Ҳ���Ϊ0');
		document.ClassForm.PointNumber.value='';
		document.ClassForm.PointNumber.focus();
	}
   CheckNumber(document.ClassForm.Money,"������ֵ");
  if (document.ClassForm.Money.value>32767||document.ClassForm.Money.value<-32768||document.ClassForm.Money.value=='0')
	{
		alert('������ֵ��������Χ��\n���32767���Ҳ���Ϊ0');
		document.ClassForm.Money.value='';
		document.ClassForm.Money.focus();
	}
  if (document.ClassForm.BrowPop.value!=''||document.ClassForm.PointNumber.value!=''||document.ClassForm.Money.value!=''){document.ClassForm.FileExtName.options[4].selected=true;document.ClassForm.FileExtName.readonly=true;}
  else {document.ClassForm.FileExtName.readonly=false;}
}

function CheckFileExtName(Obj)
{
	if (Obj.value!='')
	{
		for (var i=0;i<document.all.FileExtName.length;i++)
		{
			if (document.all.FileExtName.options(i).value=='asp') document.all.FileExtName.options(i).selected=true;
		}
		document.all.FileExtName.readonly=true;
	}
	else
	{
		document.all.FileExtName.readonly=false;
	}
}

function ShowAdpicInfo()
{
	if (document.all.IsAdPic.checked==true)
    {
        document.all.selectAp.style.display="";
        document.all.Checkbox1.value="0";
        document.all.Checkbox2.value="0";
    }
    else
    {
        document.all.selectAp.style.display="none";
        document.all.wzPic.style.display="none";
        document.all.Checkbox2.checked=false;
        document.all.Checkbox1.checked=false;
        document.all.Checkbox1.value="0";
        document.all.Checkbox2.value="0";
        document.all.Adpic.style.display="none";
    }
}
function ShowAdpicInfo1()
{
	if (document.all.Checkbox1.checked==true)
    {   
        document.all.Checkbox1.value="1";
         document.all.Checkbox2.value="0";
        document.all.Adpic.style.display="";
        document.all.Checkbox2.checked=false;
        document.all.wzPic.style.display="none";
        document.all.IsAdPic.checked=true;
    }
    else
    {
        document.all.Checkbox1.value="0";
        document.all.Adpic.style.display="none";
    }
}
function ShowAdpicInfo2()
{
	if (document.all.Checkbox2.checked==true)
    {
        document.all.Checkbox2.value="1";        
        document.all.wzPic.style.display="";
        document.all.Checkbox1.checked=false;
         document.all.Checkbox1.value="0";
        document.all.Adpic.style.display="none";
        document.all.IsAdPic.checked=true
    }
    else
    {
        document.all.Checkbox2.value="0";
        document.all.wzPic.style.display="none";
    }
}
//����ʱ���Զ���ʾ���л�
//2009-5-26 by SamJun
    if (document.all.IsAdPic.checked==true)
    {
        document.all.selectAp.style.display="";
       
    }
    else
    {
        document.all.selectAp.style.display="none";
        document.all.wzPic.style.display="none";
        document.all.Checkbox1.value="0";
        document.all.Checkbox2.value="0";
        document.all.Adpic.style.display="none";
    }
    if (document.all.Checkbox1.checked==true)
    {   
        document.all.Checkbox1.value="1";
         document.all.Checkbox2.value="0";
        document.all.Adpic.style.display="";
        document.all.Checkbox2.checked=false;
        document.all.wzPic.style.display="none";
        document.all.IsAdPic.checked=true;
    }
    else
    {
        document.all.Checkbox1.value="0";
        document.all.Adpic.style.display="none";
    }
	if (document.all.Checkbox2.checked==true)
    {
        document.all.Checkbox2.value="1";        
        document.all.wzPic.style.display="";
        document.all.Checkbox1.checked=false;
         document.all.Checkbox1.value="0";
        document.all.Adpic.style.display="none";
        document.all.IsAdPic.checked=true
    }
    else
    {
        document.all.Checkbox2.value="0";
        document.all.wzPic.style.display="none";
    }

</SCRIPT>







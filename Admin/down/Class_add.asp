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
if not MF_Check_Pop_TF("DS_Class") then Err_Show
Dim Fs_Down,NS_ClassNameValure,sRootDir,strShowErr,str_DownDir
set Fs_Down = new Cls_News
MF_GetUserGroupID
Fs_Down.GetSysParam()
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
if Fs_Down.DownDir<>"" then str_DownDir = "/"+Fs_Down.DownDir else str_DownDir=""
Dim obj_Class_Rs,ClassID,str_ClassKeywords,str_Classdescription,str_currpath
Dim lng_OrderID,str_ClassName,str_ClassEName_add,str_ParentID,str_Templet,str_NewsTemplet,str_Domain,lng_AdminID,int_RefreshNumber
Dim  lng_GroupID,lng_PointNumber,flt_Money,str_FileExtName,dtm_Addtime,int_isConstr,int_IsURL,str_UrlAddress,lng_Oldtime,int_isShow
Dim str_ClassNaviContent,str_ClassNaviPic,lng_DefineID,int_NewsCheck,tmp_fileExtName,str_SavePath,str_FileSaveType,int_isConstrDel,str_GetParentID
ClassID = NoSqlHack(Trim(Request.QueryString("ClassID")))
Select Case Fs_Down.fileExtName
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
	if not MF_Check_Pop_TF("DS010") then Err_Show
	str_Templet = Replace("//"&G_TEMPLETS_DIR&"/Down/class.htm","//","/")
	str_NewsTemplet = Replace("/"&G_TEMPLETS_DIR&"/Down/Down.htm","//","/")
	dtm_Addtime = now
	lng_AdminID = session("Admin_Name")
	lng_OrderID = 10
	lng_PointNumber = ""
	flt_Money = ""
	str_SavePath = str_DownDir
	str_UrlAddress = "http://"
	str_FileExtName = tmp_fileExtName
	int_isShow = 1
	int_RefreshNumber = 0
	str_FileSaveType = Fs_Down.ClassSaveType
	if NoSqlHack(ClassID)<>"" then
		str_GetParentID = ClassID
	Else
		str_GetParentID = "0"
	End if
	lng_Oldtime = 180
	if ClassID<>"" then
		Dim obj_IsUrlTF_Rs
		Set obj_IsUrlTF_Rs = server.CreateObject(G_FS_RS)
		obj_IsUrlTF_Rs.Open "Select IsUrl from FS_DS_Class where ClassID='"& NoSqlHack(ClassID) &"' order by id desc",Conn,1,1
		if obj_IsUrlTF_Rs(0) = 1 then
			strShowErr = "<li>外部栏目不能添加子类</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
	End if
Elseif Request.QueryString("Action")="edit" then
	if not MF_Check_Pop_TF("DS011") then Err_Show
	Set obj_Class_Rs = server.CreateObject(G_FS_RS)
	obj_Class_Rs.open "select ClassID,OrderID,ClassName,ClassEName,ParentID,Templet,NewsTemplet,[Domain],RefreshNumber,ClassAdmin,isPop,FileExtName,Addtime,isConstr,IsURL,UrlAddress,Oldtime,isShow,ClassNaviContent,ClassNaviPic,DefineID,NewsCheck,AddNewsType,SavePath,FileSaveType,isConstrDel,ClassKeywords,Classdescription From FS_DS_Class where ClassID = '"& NoSqlHack(ClassID) &"'",Conn,1,3
	if  not obj_Class_Rs.eof then
		if obj_Class_Rs("isPop")=1 then
			Dim obj_tmppop_rs
			set obj_tmppop_rs = Conn.execute("select GroupName,PointNumber,FS_Money,InfoID,PopType,isClass From FS_MF_POP where InfoID='"& obj_Class_Rs("ClassID") &"' and isClass=1 and PopType='DS'")
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
		str_SavePath = obj_Class_Rs("SavePath")
		str_FileSaveType = obj_Class_Rs("FileSaveType")
		int_isConstrDel = obj_Class_Rs("isConstrDel")
		str_ClassKeywords  = obj_Class_Rs("ClassKeywords")
		str_Classdescription  = obj_Class_Rs("Classdescription")
		obj_Class_Rs.close
		set  obj_Class_Rs = nothing
	Else
		strShowErr = "<li>错误的参数</li>"
		Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
Else
		strShowErr = "<li>错误的参数</li>"
		Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
End if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>栏目管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function insertType() { 
	if (document.ClassForm.isUrl.checked==true)
		OutUrl.style.display=''
	else
		OutUrl.style.display='none'
	if (document.ClassForm.isUrl.checked==true)
		InUrl.style.display='none';
	else
		InUrl.style.display='';
}
//-->
</script>
<script language="JavaScript" src="../../FS_Inc/GetLettersByChinese.js"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
</head>
  <body>
<form name="ClassForm" method="post" action="Class_Save.asp">
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr class="hback"> 
      <td class="xingmu">栏目管理<a href="../../help?Lable=NS_Class_add" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
    </tr>
    <tr> 
      <td height="18" class="hback"><div align="left"><a href="Class_Manage.asp">管理首页</a>┆<a href="Class_add.asp?ClassID=&Action=add">添加根栏目</a>┆<a href="Class_Action.asp?Action=one">一级栏目排序</a>┆<a href="Class_Action.asp?Action=n">N级栏目排序</a>┆<a href="Class_Action.asp?Action=reset"   onClick="{if(confirm('确认复位所有栏目？\n\n如果选择确定，所有的栏目将设置为一级分类!!')){return true;}return false;}">复位所有栏目</a>┆<a href="Class_Action.asp?Action=unite">栏目合并</a>┆<a href="Class_Action.asp?Action=allmove">栏目转移</a>┆<a href="Class_Action.asp?Action=clearClass"  onClick="{if(confirm('确认清空所有栏目里的数据吗？\n\n如果选择确定,所有的栏目的下载将被放到回收站中!!')){return true;}return false;}">删除所有栏目</a> ┆ <a  href="#" onClick="javascirp:history.back()">后退</a>  
          <a href="../../help?Lable=NS_Class_add_1" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></div></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr class="hback"> 
      <td colspan="3" class="xingmu">添加栏目</td>
    </tr>
    <tr> 
      <td width="23%" height="29" class="hback"> <div align="right">栏目类型： </div></td>
      <td width="14%" class="hback"><input name="isUrl" type="checkbox" id="isUrl"  onClick="insertType()" value="1" <%if int_IsURL = 1 then response.Write("Checked")%>>
        外部栏目 
        <input name="ClassID" type="hidden" id="ClassID" value="<% = ClassID %>"> 
        <input name="ParentID" type="hidden" id="ParentID" value="<% = str_GetParentID %>" readonly> 
      </td>
      <td width="63%" class="hback"><span class="tx">内部栏目具有详细的参数设置。可以添加子栏目和下载<br>
        外部栏目指链接到本系统以外的地址中。当此栏目准备链接到网站中的其他系统时，请使用这种方式。不能在外部栏目中添加下载，也不能添加子栏目。</span></td>
    </tr>
    <tr> 
      <td width="23%" class="hback"><div align="right">栏目中文名称：</div></td>
      <td colspan="2" class="hback"><input onBlur="<% if Request.QueryString("Action")="add" then %>SetClassEName(this.value,document.ClassForm.ClassEName);<% end if %>" name="ClassName" type="text" id="ClassName" size="40" maxlength="100" value="<% = str_ClassName%>">
        <span class="tx"> *3-100个字符</span></td>
    </tr>
    <tr> 
      <td height="22" class="hback"><div align="right">父级栏目ID：</div></td>
      <td height="22" colspan="2" class="hback"> <%
	  Dim str_Parentvalue
	  if Request.QueryString("Action") = "add" then
	  		if Not isnull(Trim(ClassID)) then
				str_Parentvalue = Fs_Down.GetClassName(NoSqlHack(ClassID))
			Else
				str_Parentvalue = "根栏目" 
			End if
	 Elseif Request.QueryString("Action") = "edit" then
	 		if str_ParentID = "0" then
				str_Parentvalue = "根栏目"
			Else		
				str_Parentvalue = Fs_Down.GetClassName(NoSqlHack(str_ParentID))
			End if
	 End if
	  %> 
        <input name="ParentIDs" type="text" id="ParentIDs" value="<% = str_Parentvalue %>" size="40" readonly> 
        <span class="tx"> *0为根栏目</span></td>
    </tr>
  </table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table" style="<%if int_IsURL = 1 then%>display:<%else%>display:none<%end if%>" id="OutUrl">
    <tr > 
      <td width="23%" height="19" class="hback"> 
        <div align="right">外部地址：</div></td>
      <td width="78%" height="19" class="hback"><input name="UrlAddress" type="text" id="UrlAddress" size="40" maxlength="250" value="<% = str_UrlAddress%>">
       <span class="tx"> *</span> 最大250个字符</td>
    </tr>
</table>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table"  id="InUrl" style="<%if int_IsURL = 1 then%>display:none<%else%>display:<%end if%>">
    <tr id="InUrl1" style="dispay:"> 
      <td width="23%" class="hback"><div align="right">栏目英文名称：</div></td>
      <td width="77%" class="hback"><input name="ClassEName" type="text" id="ClassEName" size="40" maxlength="50" value="<% =str_ClassEName_add%>" <%if Request.QueryString("Action")="edit" then response.Write("Readonly")%>> 
        <span class="tx"> *<br>
        3-50个字符,必须是字母，数字，中划线，下划线,@,.，一旦确认,将不能修改</span></td>
    </tr>
    <tr id="InUrl2" style="dispay:"> 
      <td class="hback"><div align="right">栏目模板地址：</div></td>
      <td class="hback"><input name="Templet" type="text" id="Templet" value="<% = str_Templet %>" size="50" maxlength="250" readonly> 
        <input type="button" name="Submit" value="选择模板" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir %>/<% = G_TEMPLETS_DIR %>',400,300,window,document.ClassForm.Templet);document.ClassForm.Templet.focus();"> 
        <span class="tx"> *250个字符</span></td>
    </tr>
    <tr id="InUrl3" style="dispay:"> 
      <td class="hback"><div align="right">下载模板地址：</div></td>
      <td class="hback"><input name="NewsTemplet" type="text" id="NewsTemplet" value="<% = str_NewsTemplet %>" size="50" maxlength="250" readonly> 
        <input type="button" name="Submit2"  value="选择模板" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir %>/<% = G_TEMPLETS_DIR %>',400,300,window,document.ClassForm.NewsTemplet);document.ClassForm.NewsTemplet.focus();"> 
        <span class="tx"> *250个字符</span></td>
    </tr>
    <tr id="InUrl4" style="dispay:"> 
      <td class="hback"><div align="right">捆绑域名：</div></td>
      <td class="hback"><input name="Domain" type="text" id="Domain" size="40" maxlength="150" value="<% = str_Domain %>"  <%if len(Trim(str_Domain))>=6 then Response.Write("readonly")%>>
        <span class="tx">150个字符,请填写正确的域名</span></td>
    </tr>
    <tr id="InUrl5" style="dispay:"> 
      <td class="hback"><div align="right">管理员：</div></td>
      <td class="hback"> <SELECT name="ClassAdmin" id="ClassAdmin">
          <%
			Dim obj_AdminList_Rs
			set obj_AdminList_Rs = Conn.Execute("Select Admin_Name,Admin_Real_Name from FS_MF_Admin Where Admin_Parent_Admin='"&Temp_Admin_Name&"' or Admin_Name='"&Temp_Admin_Name&"' order by ID asc")
			If not obj_AdminList_Rs.eof Then
				if lng_AdminID = obj_AdminList_Rs("Admin_Name") then
					Response.Write "<OPTION value=""" & obj_AdminList_Rs("Admin_Name") & """ selected>管理员帐号：" & obj_AdminList_Rs("Admin_Name") & "　管理员姓名：" & obj_AdminList_Rs("Admin_Real_Name") & "</OPTION>"
				Else
					Response.Write "<OPTION value=""" & obj_AdminList_Rs("Admin_Name") & """>管理员帐号：" & obj_AdminList_Rs("Admin_Name") & "　管理员姓名：" & obj_AdminList_Rs("Admin_Real_Name") & "</OPTION>"
				End if
				obj_AdminList_Rs.MoveNext
			End If
			Do while not obj_AdminList_Rs.eof
				if lng_AdminID = obj_AdminList_Rs("Admin_Name") then
					Response.Write "<OPTION value=""" & obj_AdminList_Rs("Admin_Name") & """ selected>管理员帐号：" & obj_AdminList_Rs("Admin_Name") & "　管理员姓名：" & obj_AdminList_Rs("Admin_Real_Name") & "</OPTION>"
				Else
					Response.Write "<OPTION value=""" & obj_AdminList_Rs("Admin_Name") & """>管理员帐号：" & obj_AdminList_Rs("Admin_Name") & "　管理员姓名：" & obj_AdminList_Rs("Admin_Real_Name") & "</OPTION>"
				End if
				obj_AdminList_Rs.Movenext
			Loop
			obj_AdminList_Rs.Close
			Set obj_AdminList_Rs = Nothing
			%>
        </SELECT> <span class="tx">管理员必须选择<a href="../../help?Lable=NS_Class_Admin" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></span>　 
      </td>
    </tr>
    <tr id="InUrl6" style="dispay:"> 
      <td class="hback"><div align="right">浏览会员组：</div></td>
      <td class="hback"> <input name="BrowPop"  id="BrowPop" type="text" value="<% = lng_GroupID %>" onMouseOver="this.title=this.value;" readonly> 
        <select name="selectPop" id="selectPop" style="overflow:hidden;" onChange="ChooseExeName();">
          <option value="" selected>选择会员组</option>
          <option value="del" style="color:red;">清空</option>
          <% = MF_GetUserGroupID %>
        </select>
        需要点数 
        <input name="PointNumber" type="text" id="PointNumber" size="8" maxlength="5" value="<% = lng_PointNumber %>"  onChange="ChooseExeName();">
        需要金币 
        <input name="Money" type="text" id="Money" size="8" maxlength="5" value="<% = flt_Money %>"  onChange="ChooseExeName();"></td>
    </tr>
    <tr id="InUrl7" style="dispay:"> 
      <td class="hback"><div align="right">静态文件扩展名：</div></td>
      <td class="hback"><select name="FileExtName" id="FileExtName">
          <option value="html" <% if  Trim(str_FileExtName) = "html"  then response.Write("selected")%>>.html</option>
          <option value="htm" <% if  Trim(str_FileExtName) = "htm"  then response.Write("selected")%>>.htm</option>
          <option value="shtml" <% if  Trim(str_FileExtName) = "shtml"  then response.Write("selected")%>>.shtml</option>
          <option value="shtm" <% if  Trim(str_FileExtName)= "shtm"  then response.Write("selected")%>>.shtm</option>
          <option value="asp" <% if  Trim(str_FileExtName) = "asp"  then response.Write("selected")%>>.asp</option>
        </select> <span class="tx"> *如果需要阅读权限，必须设置为.asp</span></td>
    </tr>
    <tr id="InUrl8" style="dispay:"> 
      <td class="hback"><div align="right">是否允许投稿：</div></td>
      <td class="hback"><input name="isConstr" type="checkbox" id="isConstr" value="1" <%if int_isConstr = 1 Then response.Write("checked")%>>
        　　　会员投稿是否允许删除 
        <input name="isConstrDel" type="checkbox" id="isConstrDel" value="1"  <%if int_isConstrDel = 1 Then response.Write("checked")%>>
        是</td>
    </tr>
    <tr id="InUrl10" style="dispay:"> 
      <td class="hback"><div align="right">栏目首页保存模式：</div></td>
      <td class="hback"><select name="FileSaveType" id="FileSaveType">
          <option value="0" <%if str_FileSaveType = 0 Then response.Write("selected")%>>栏目英文/index.html</option>
          <option value="1" <%if str_FileSaveType = 1 Then response.Write("selected")%>>栏目英文/栏目英文.html</option>
          <option value="2" <%if str_FileSaveType = 2 Then response.Write("selected")%>>栏目英文.html</option>
        </select> <span class="tx"> *</span></td>
    </tr>
    <tr id="InUrl11" style="dispay:"> 
      <td class="hback"><div align="right">栏目保存路径：</div></td>
      <td class="hback"><input name="SavePath" type="text" id="SavePath" value="<%=str_SavePath%>" size="40" maxlength="255" readonly> 
        <%if Request.QueryString("Action")="add" then%> <INPUT type="button"  name="Submit4" value="选择路径" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPathFrame.asp?CurrPath=<%=sRootDir & str_DownDir%>',320,280,window,document.ClassForm.SavePath);document.ClassForm.SavePath.focus();"> 
        <%End if%>
        <span class="tx"> *<br>
        一旦填写将不能修改;如果选择了二级域名，请存放在下载目录的根目录</span></td>
    </tr>
    <tr id="InUrl11" style="dispay:"> 
      <td class="hback"><div align="right">自定义选择：</div></td>
      <td class="hback"> <select name="DefineID" id="DefineID">
          <option value="0" selected>选择自定义分类</option>
          <% = Fs_Down.GetDefineClassId%>
        </select> </td>
    </tr>
    <tr id="InUrl11" style="dispay:"> 
      <td class="hback"><div align="right">下载需要审核：</div></td>
      <td class="hback"><input name="NewsCheck" type="checkbox" id="NewsCheck" value="1" <%if int_NewsCheck = 1 then response.Write("checked")%>>
        需要审核</td>
    </tr>
    <tr id="InUrl11" style="dispay:"> 
      <td class="hback"><div align="right">栏目META关键字：</div></td>
      <td class="hback"><textarea name="ClassKeywords" style="width:80%" rows="5" id="ClassKeywords"><% = str_ClassKeywords %></textarea> 
        <span class="tx"><br>
        最多200个字符,用户搜索引擎搜索，可以提高栏目被搜索引擎搜索收录的机会</span></td>
    </tr>
    <tr id="InUrl11" style="dispay:"> 
      <td class="hback"><div align="right">栏目META描述：</div></td>
      <td class="hback"><textarea name="Classdescription"  style="width:80%" rows="5" id="Classdescription"><% = str_Classdescription %></textarea> 
        <span class="tx"><br>
        最多200个字符,用户搜索引擎搜索，可以提高栏目被搜索引擎搜索收录的机会</span></td>
    </tr>
    <tr id="InUrl11" style="dispay:"> 
      <td class="hback"><div align="right"> 多少天后归档：</div></td>
      <td class="hback"><input  name="Oldtime" type="text" id="Oldtime" value="<% = lng_Oldtime %>" size="40"></td>
    </tr>
    <tr id="InUrl11" style="dispay:">
      <td class="hback"><div align="right">发布最新多少条信息</div></td>
      <td class="hback"><input  name="RefreshNumber" type="text" id="RefreshNumber" value="<% = int_RefreshNumber %>" size="40">
        <span class="tx">如果为0则不限制</span></td>
    </tr>
  </table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr> 
      <td width="23%" height="21" class="hback"><div align="right">添加日期：</div></td>
      <td width="78%" height="21" class="hback"><input  name="Addtime" type="text" id="Addtime" value="<% = dtm_Addtime %>" size="40"></td>
    </tr>
    <tr> 
      <td height="22" class="hback"><div align="right">是否在导航中显示：</div></td>
      <td height="22" class="hback"><input name="isShow" type="checkbox" id="isShow" value="1" <% if int_isShow = 1 then response.Write("checked") %>></td>
    </tr>
    <tr> 
      <td height="21" class="hback"><div align="right">栏目导航说明：</div></td>
      <td height="21" class="hback"><textarea name="ClassNaviContent"  style="width:80%" rows="6" id="ClassNaviContent"><% = str_ClassNaviContent%></textarea></td>
    </tr>
    <tr> 
      <td height="21" class="hback"><div align="right">栏目导航图片：</div></td>
      <td height="21" class="hback"><input name="ClassNaviPic" type="text" id="ClassNaviPic" value="<% = str_ClassNaviPic%>" size="40">
        <input type="button" name="PPPChoose"  value="选择图片" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath%>',500,300,window,document.ClassForm.ClassNaviPic);"></td>
    </tr>
    <tr> 
      <td height="21" class="hback"><div align="right">排列权重：</div></td>
      <td height="21" class="hback"><input name="OrderID" type="text" id="OrderID" value="<% = lng_OrderID%>" size="40"></td>
    </tr>
    <tr> 
      <td height="21" class="hback"><div align="right"></div></td>
      <td height="21" class="hback"><input type="button" name="Submit4222" value="保存栏目" onClick="{if(confirm('确认保存您的栏目信息吗?')){this.document.ClassForm.submit();return true;}return false;}"> 
        <input type="reset" name="Submit5222" value="重置">
        <input name="str_add" type="hidden" id="str_add" value="<% = Request.QueryString("Action")%>"></td>
    </tr>
</table>
</form>
</body>
</html>
<%
set Fs_Down = nothing
%>
<SCRIPT language="JavaScript">
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
   CheckNumber(document.ClassForm.PointNumber,"浏览扣点值");
  if (document.ClassForm.PointNumber.value>32767||document.ClassForm.PointNumber.value<-32768||document.ClassForm.PointNumber.value=='0')
	{
		alert('浏览扣点值超过允许范围！\n最大32767，且不能为0');
		document.ClassForm.PointNumber.value='';
		document.ClassForm.PointNumber.focus();
	}
   CheckNumber(document.ClassForm.Money,"浏览金币值");
  if (document.ClassForm.Money.value>32767||document.ClassForm.Money.value<-32768||document.ClassForm.Money.value=='0')
	{
		alert('浏览金币值超过允许范围！\n最大32767，且不能为0');
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
function SetClassEName(Str,Obj)
{
	Obj.value=ConvertToLetters(Str,1);
}
</SCRIPT>
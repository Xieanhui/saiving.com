<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_InterFace/HS_Function.asp" -->
<!--#include file="../../FS_InterFace/AP_Function.asp" -->
<!--#include file="../../FS_Inc/Cls_SysConfig.asp"-->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,obj_Label_Rs,SQL,strShowErr,str_CurrPath,sRootDir
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
MF_Default_Conn
'session�ж�
MF_Session_TF 
if not MF_Check_Pop_TF("MF025") then Err_Show
Dim LableName,txt_Content,LableClassID,Labelclass_SQL,obj_Labelclass_rs,obj_Count_rs,isDel,tmps_LableName
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if
Dim Label_ConMaxNum,Sys_Obj
Set Sys_Obj = New Cls_SysConfig
Sys_Obj.getSysParam()
Label_ConMaxNum = Sys_Obj.Label_MaxNum
If Label_ConMaxNum = "" Or IsNull(Label_ConMaxNum) Then Label_ConMaxNum = 0
Set Sys_Obj = NOthing
Rem End

LableClassID = NoSqlHack(Request.QueryString("LableClassID"))
LableName = Trim(Request.Form("LableName"))
txt_Content = Trim(Request.Form("TxtFileds"))
isDel = Trim(Request.Form("isDel"))

if Request.Form("Action") = "add_save" then
	if LableName ="" or txt_Content =""  then
		strShowErr = "<li>����д����</li>"
		Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	If Clng(Label_ConMaxNum) > 0 Then
		if len(txt_Content) > Clng(Label_ConMaxNum)  then
			strShowErr = "<li>��ǩ���ݲ��������" & Clng(Label_ConMaxNum) & "���ַ�</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
	End If	
	Labelclass_SQL = "Select LableName,LableContent,AddDate,LableClassID,isDel from FS_MF_Lable where LableName ='"& "{FS400_"&LableName&"}"&"'"
	Set obj_Labelclass_rs = server.CreateObject(G_FS_RS)
	obj_Labelclass_rs.Open Labelclass_SQL,Conn,1,3
	if obj_Labelclass_rs.eof then
		obj_Labelclass_rs.addnew
		obj_Labelclass_rs("LableName") = "{FS400_"& LableName &"}"
		obj_Labelclass_rs("LableContent") = txt_Content
		obj_Labelclass_rs("AddDate") =now
		if isDel<>"" then
			obj_Labelclass_rs("isDel") =1
		else
			obj_Labelclass_rs("isDel") =0
		end if
		obj_Labelclass_rs("LableClassID") =Request.Form("LableClassID")
		obj_Labelclass_rs.update
	else
		strShowErr = "<li>�����ظ�,����������</li>"
		Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	obj_Labelclass_rs.close:set obj_Labelclass_rs =nothing
	strShowErr = "<li>��ӳɹ�</li><li><a href=Label/Label_Creat_txt.asp>�������</a></li><li><a href=Label/All_Label_Stock.asp?classid="&Request.Form("LableClassID")&">���ر�ǩ����</a></li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Label/Label_Creat_txt.asp")
	Response.end
elseif Request.Form("Action") = "edit_save" then
	dim rstf
	tmps_LableName="{FS400_"&LableName&"}"
	Set rstf = Conn.execute("Select LableName,LableContent,AddDate,LableClassID,isDel from FS_MF_Lable where LableName ='"& NoSqlHack(tmps_LableName) &"' and id <>"& CintStr(Request.Form("ID")))
	if not rstf.eof then
		strShowErr = "<li>�����ظ�,����������</li>"
		Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	If Clng(Label_ConMaxNum) > 0 Then
		if len(txt_Content) > Clng(Label_ConMaxNum)  then
			strShowErr = "<li>��ǩ���ݲ��������" & Clng(Label_ConMaxNum) & "���ַ�</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
	End If	
	Labelclass_SQL = "Select id,isDel,LableName,LableContent,AddDate,LableClassID from FS_MF_Lable where id ="& NosqlHack(Request.Form("ID")) 
	Set obj_Labelclass_rs = server.CreateObject(G_FS_RS)
	obj_Labelclass_rs.Open Labelclass_SQL,Conn,1,3
	if not obj_Labelclass_rs.eof then
		obj_Labelclass_rs("LableName") = "{FS400_"& LableName &"}"
		obj_Labelclass_rs("LableContent") = txt_Content
		obj_Labelclass_rs("AddDate") =now
		if isDel<>"" then
			obj_Labelclass_rs("isDel") =1
		else
			obj_Labelclass_rs("isDel") =0
		end if
		obj_Labelclass_rs("LableClassID") =Request.Form("LableClassID")
		obj_Labelclass_rs.update
		obj_Labelclass_rs.close:set obj_Labelclass_rs =nothing
	else
		obj_Labelclass_rs.close:set obj_Labelclass_rs =nothing
		strShowErr = "<li>����Ĳ���</li>"
		Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	strShowErr = "<li>�޸ĳɹ�</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Label/all_Label_Stock.asp")
	Response.end
end if
if Request.QueryString("DelTF")="1" then
	Conn.execute("Delete From FS_MF_Labestyle where StyleType='"& NoSqlHack(Request.QueryString("Label_Sub"))&"' and id="&CintStr(Request.QueryString("id")))
	strShowErr = "<li>ɾ���ɹ�</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Label/All_Label_style.asp?Label_Sub="& Request.Form("Label_Sub")&"")
	Response.end
end if
dim tmp_LableName,tmp_LableClassID,tmp_LableContent,tmp_isDel,tmp_id,tmp_action
if Request.QueryString("type")="edit" then
	dim rs
	if not isnumeric(Request.QueryString("id")) then
		strShowErr = "<li>����Ĳ���</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	set rs = Conn.execute("select id,LableName,LableClassID,LableContent,isDel From FS_MF_Lable where id="&CintStr(Request.QueryString("id")))
	if rs.eof then
		rs.close:set rs=nothing
		strShowErr = "<li>����Ĳ���</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		tmp_LableName=Replace(Replace(rs("LableName"),"{FS400_",""),"}","")
		tmp_LableClassID=rs("LableClassID")
		tmp_LableContent=rs("LableContent")
		tmp_isDel=rs("isDel")
		tmp_id = rs("id")
		tmp_action = "edit_save"
	end if
else
	tmp_action = "add_save"
end if
%>
<html>
<head>
<title>��ǩ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Get_Domain.asp"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/Prototype.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/CheckJs.js" type="text/JavaScript"></script>
<body>
<table width="98%" height="56" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
	<tr class="hback" >
		<td width="100%" height="20"  align="Left" class="xingmu">��ǩ��</td>
	</tr>
	<tr class="hback" >
		<td height="27" align="center" class="hback"><div align="left"><a href="All_Label_Stock.asp">���б�ǩ</a>��<a href="../FreeLabel/FreeLabelList.asp"><font color="#FF0000">���ɱ�ǩ</font></a>��<a href="All_Label_Stock.asp?isDel=1">���ݿ�</a>��<a href="label_creat.asp">������ǩ</a>��<a href="label_creat_txt.asp">�ı�������ǩ</a>��<a href="Label_Class.asp" target="_self">��ǩ����</a>&nbsp;��<a href="All_label_style.asp">��ʽ����</a>&nbsp;<a href="../../help?Label=MF_Label_Stock" target="_blank" style="cursor:help;"><img src="../Images/_help.gif" width="50" height="17" border="0"></a></div></td>
	</tr>
</table>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
	<tr class="xingmu">
		<td colspan="2" class="xingmu">������ǩ</td>
	</tr>
	<form name="NewsForm" method="post" action="" onSubmit="return CheckForm(this);" target="_self">
		<tr class="hback">
			<td width="8%">
				<div align="right">��ǩ����</div>
			</td>
			<td width="92%"><span class="tx">{FS400_
				<input name="LableName" type="text" value="<%=tmp_LableName%>" id="LableName" size="18"  style="border-top-width: 0px;border-right-width: 0px;border-bottom-width: 0px;border-left-width:0px;border-bottom-color: #000000;"  onFocus="Do.these('LableName',function(){return CheckContentLen('LableName','span_LableName','2-30')})" onKeyUp="Do.these('LableName',function(){return CheckContentLen('LableName','span_LableName','2-30')})">
				}</span><span id="span_LableName"></span>
				<select name="LableClassID" id="LableClassID">
					<option value="0">ѡ��������Ŀ</option>
					<%
				  dim class_rs_obj
				  set class_rs_obj=Conn.execute("select id,ParentID,ClassName From FS_MF_LableClass where ParentID=0 order by id desc")
				  do while not class_rs_obj.eof
						If CStr(tmp_LableClassID)=CStr(class_rs_obj("id")) Then 
							response.Write "<option value="""&class_rs_obj("id")&""" selected >"&class_rs_obj("ClassName")&"</option>"
						Else
							response.Write "<option value="""&class_rs_obj("id")&""">"&class_rs_obj("ClassName")&"</option>"
						End If 
						response.Write get_childList(class_rs_obj("id"),"")
					class_rs_obj.movenext
					
				  loop
				  class_rs_obj.close:set class_rs_obj=nothing
				  %>
				</select>
				<label>
				<input name="isDel" type="checkbox" id="isDel" value="1" <%if tmp_isDel=1 then response.Write"checked"%>>
				���뱸�ݿ�</label>
			</td>
		</tr>
		<tr class="hback" <%if request.QueryString("Label_Sub")<>"DS" then response.Write("style=""display:'none';""") else response.Write("style=""display:'';"" ") end if%>> </tr>
		<tr class="hback">
			<td>&nbsp;</td>
			<td>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td height="16" valign="top">
							<!--<OBJECT ID="Lable"
			CLASSID="CLSID:389961B3-2025-4E6D-92E8-AE75352096E2"
			CODEBASE="Foosun.CAB#version=1,0,0,0">
			</OBJECT>
			-->
						<td height="16" valign="top"> <span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/News_label.asp',500,480,'obj');" style="cursor:hand;"> <a href="#" title="�����б����ǩ">�����б�</a> | </span><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/News_C_label.asp',500,350,'obj');" style="cursor:hand;"> <a href="#" title="���ų������ǩ">���ų���</a> | </span><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/News_Un_label.asp',500,350,'obj');" style="cursor:hand;"> <a href="#" title="���Ų��������ű�ǩ">����������</a> | </span><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBMS")=1 then%><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/Mall_label.asp',500,450,'obj');" style="cursor:hand;"> <a href="#" title="�̳��б����ǩ">�̳��б�</a> | </span><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/Mall_C_label.asp',450,350,'obj');" style="cursor:hand;"> <a href="#" title="�̳ǳ������ǩ">�̳ǳ���</a> | </span><%end if%><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/down_label.asp',500,450,'obj');" style="cursor:hand;"><a href="#" title="�����б����ǩ">�����б�</a> | </span><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/down_C_label.asp',400,320,'obj');" style="cursor:hand;"><a href="#" title="���س������ǩ">���س���</a> | </span><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/All_label.asp',500,360,'obj');" style="cursor:hand;display:;"><a href="#" title="ͨ�����ǩ-������ϵͳ����ʹ��">ͨ�ñ�ǩ</a> | </span><a href="All_label_style.asp" target="_self" title="��ʽ����"> ��ʽ����</a> | <span  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(id_templet);" onMouseOver="this.className='bg'" onMouseOut="this.className='bg1'" language=javascript><a href="#" title="�����ǩ">�����ǩ</a></span></td>
					</tr>
				</table>
				<table width="100%" border="0" cellspacing="0" cellpadding="0" id="id_templet" style="display:none;">
					<tr>
						<td valign="top"> <%if Request.Cookies("FoosunSUBCookie")("FoosunSUBAP")=1 then%><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/job_label.asp',500,380,'obj');" style="cursor:hand;"><a href="#" title="�˲����ǩ">�˲ű�ǩ</a> | </span><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/job_Search_label.asp',500,260,'obj');" style="cursor:hand;"><a href="#" title="�˲����ǩ">�˲�����</a> | </span><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/job_C_label.asp',300,250,'obj');" style="cursor:hand;display:none"><a href="#" title="�˲����ǩ">�˲����ǩ</a> | </span><%End if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBSD")=1 then%><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/supply_C_label.asp',480,460,'obj');" style="cursor:hand;"><a href="#" title="�������ǩ">�����ǩ</a> | </span><%end if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBHS")=1 then%><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/house_label.asp',500,400,'obj');" style="cursor:hand;"><a href="#" title="�������ǩ">������ǩ</a> | </span><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/house_C_label.asp',350,350,'obj');" style="cursor:hand;"><a href="#" title="�����ೣ���ǩ">��������</a> | </span><%end if%><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/FL_C_label.asp',300,350,'obj');" style="cursor:hand;"><a href="#" title="�����������ǩ">��������</a> | </span><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/st_C_label.asp',300,250,'obj');" style="cursor:hand;"><a href="#" title="����ͳ�����ǩ">����ͳ��</a> | </span><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/vote_C_label.asp',380,200,'obj');" style="cursor:hand;"><a href="#" title="ͶƱ���ǩ">ͶƱ</a> | </span><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/ads_C_label.asp',300,250,'obj');" style="cursor:hand;"><a href="#" title="��泣�����ǩ">���</a> | </span><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/book_C_label.asp',400,420,'obj');" style="cursor:hand;"><a href="#" title="�������ǩ">����</a> | </span> </span><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/log_C_label.asp',400,400,'obj');" style="cursor:hand;"><a href="#" title="��־���ǩ">��־</a> | </span><span onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/photo_C_label.asp',320,350,'obj');" style="cursor:hand;"><a href="#" title="������ǩ">���</a> </span></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr class="hback">
			<td>
				<div align="right">��ǩ����</div>
			</td>
		  <td>
			    <textarea style="width:100%;" name="TxtFileds" rows="16"><%=Server.HTMLEncode(tmp_LableContent)%></textarea></td>
		</tr>
		<tr class="hback">
			<td>&nbsp;</td>
			<td>
				<input type="submit" name="Submit" value="ȷ�ϱ����ǩ">
				<input name="Action" type="hidden" id="Action" value="<%=tmp_action%>">
				<input name="id" type="hidden" id="Action" value="<%=tmp_id%>">
				<input type="reset" name="Submit2" value="����">
			</td>
		</tr>
	</form>
</table>
<%
Function get_childList(TypeID,f_CompatStr)  
	Dim f_ChildNewsRs,ChildTypeListStr,f_TempStr,f_isUrlStr,lng_GetCount
	Set f_ChildNewsRs = Conn.Execute("Select id,ParentID,ClassName from FS_MF_LableClass where ParentID=" & CintStr(TypeID) & " order by id desc" )
	f_TempStr =f_CompatStr & "��"
	do while Not f_ChildNewsRs.Eof
			get_childList = get_childList & "<option value="""& f_ChildNewsRs("id")&""""
			If CStr(tmp_LableClassID)=CStr(f_ChildNewsRs("id")) then
				get_childList = get_childList & " selected" & Chr(13) & Chr(10)	
			End If
			get_childList = get_childList & ">��" &  f_TempStr & f_ChildNewsRs("ClassName") 
			get_childList = get_childList & "</option>" & Chr(13) & Chr(10)
			get_childList = get_childList &get_childList(f_ChildNewsRs("id"),f_TempStr)
		f_ChildNewsRs.MoveNext
	loop
	f_ChildNewsRs.Close
	Set f_ChildNewsRs = Nothing
End Function
Set Conn=nothing
%>
</html>
<script language="JavaScript" type="text/JavaScript">
function Insertlabel_News(URL,widthe,heighte,obj)
{

  var obj=window.OpenWindowAndSetValue("../../Fs_Inc/convert.htm?"+URL,widthe,heighte,'window',obj)
  if (obj==undefined)return false;
  if (obj!='')InsertEditor(obj);
}
function InsertEditor(InsertValue)
{
	document.NewsForm.TxtFileds.focus(); 
	document.selection.createRange().text=InsertValue; 
}
function opencat(cat)
{
  if(cat.style.display=="none"){
     cat.style.display="";
  } else {
     cat.style.display="none"; 
  }
}
function CheckForm(FormObj)
{
	return true;
}
</script>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->

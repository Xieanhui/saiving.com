<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_InterFace/HS_Function.asp" -->
<!--#include file="../../FS_InterFace/AP_Function.asp" -->
<!--#include file="Func_page.asp" -->
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
MF_Session_TF 
Dim GetSysConfigObj,GetStyleMaxNum
Set GetSysConfigObj = New Cls_SysConfig
GetSysConfigObj.getSysParam()
GetStyleMaxNum = Clng(GetSysConfigObj.Style_MaxNum)
Set GetSysConfigObj = Nothing
Dim str_StyleName,txt_Content,Labelclass_SQL,obj_Labelclass_rs,obj_Count_rs
Dim Label_Sub

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
Label_Sub = NoSqlHack(Request.QueryString("Label_Sub"))
str_StyleName = Trim(Request.Form("StyleName"))
txt_Content = Trim(Request.Form("TxtFileds"))
if Request.Form("Action") = "Add_save" then
	if str_StyleName ="" or txt_Content ="" then
		strShowErr = "<li>����д����</li>"
		Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	set obj_Count_rs = server.CreateObject(G_FS_RS)
	obj_Count_rs.Open "Select StyleName,Content,AddDate,LableClassID from FS_MF_Labestyle Order by id desc",Conn,1,3
	if Not obj_Count_rs.eof then
		if obj_Count_rs.recordcount>GetStyleMaxNum then
			strShowErr = "<li>����������ʽ�Ѿ�����" & GetStyleMaxNum & "�����㽫����������\n�����Ҫ���ӣ���ɾ��������ʽ</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Label/All_Label_style.asp?Action=Add&Label_Sub="&Request.Form("Label_Sub")&"")
			Response.end
		end if
	end if
	Labelclass_SQL = "Select StyleName,Content,AddDate,StyleType,LableClassID from FS_MF_Labestyle where StyleName ='"& str_StyleName &"'"
	Set obj_Labelclass_rs = server.CreateObject(G_FS_RS)
	obj_Labelclass_rs.Open Labelclass_SQL,Conn,1,3
	if obj_Labelclass_rs.eof then
		obj_Labelclass_rs.addnew
		obj_Labelclass_rs("StyleName") = str_StyleName
		obj_Labelclass_rs("content") = txt_Content
		obj_Labelclass_rs("AddDate") =now
		obj_Labelclass_rs("StyleType") =Request.Form("Label_Sub")
		obj_Labelclass_rs("LableClassID") =Request.Form("LableClassID")'---д�����ݿ�-------2/1 by chen---------
		'tmp_LableClassID=obj_Labelclass_rs("LableClassID")
		obj_Labelclass_rs.update
	else
			strShowErr = "<li>����ʽ�����ظ�,����������</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	end if
	obj_Labelclass_rs.close:set obj_Labelclass_rs =nothing
	strShowErr = "<li>��ӳɹ�</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Label/All_Label_style.asp?Label_Sub="& Request.Form("Label_Sub")&"")
	Response.end
elseif Request.Form("Action") = "Add_edit" then
	Labelclass_SQL = "Select StyleName,Content,AddDate,LableClassID from FS_MF_Labestyle where id ="& NosqlHack(Request.Form("ID")) 
	Set obj_Labelclass_rs = server.CreateObject(G_FS_RS)
	obj_Labelclass_rs.Open Labelclass_SQL,Conn,1,3
	if not obj_Labelclass_rs.eof then
		obj_Labelclass_rs("StyleName") = str_StyleName
		obj_Labelclass_rs("content") = txt_Content
		'obj_Labelclass_rs("AddDate") =now
	obj_Labelclass_rs("LableClassID")=Request.Form("LableClassID")'--------д�����ݿ�----2/1 by chen--------------------
	'tmp_LableClassID=obj_Labelclass_rs("LableClassID")
		obj_Labelclass_rs.update
	End if
	obj_Labelclass_rs.close:set obj_Labelclass_rs =nothing
	strShowErr = "<li>�޸ĳɹ�</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Label/All_Label_style.asp?Label_Sub="& Request.Form("Label_Sub")&"")
	Response.end
end if
if Request.QueryString("DelTF")="1" then
	Conn.execute("Delete From FS_MF_Labestyle where StyleType='"& Request.QueryString("Label_Sub")&"' and id="&CintStr(Request.QueryString("id")))
	strShowErr = "<li>ɾ���ɹ�</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Label/All_Label_style.asp?Label_Sub="& Request.Form("Label_Sub")&"")
	Response.end
end if
%>
<html>
<head>
<title>��ǩ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/Prototype.js"></script>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Get_Domain.asp"></script>
<body>
<table width="98%" height="81" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="hback" >
    <td width="100%" height="20"  align="Left" class="xingmu">������ʽ����</td>
  </tr>
  <tr class="hback" >
    <td class="hback" align="center"><div align="left"><a href="../Templets_List.asp">ģ�����</a>
        <%
	if Request.QueryString("TF") = "NS" then
		Response.Write("��<a href=News_Label.asp target=_self>���ر�ǩ����</a>") 
	elseif Request.QueryString("TF") = "DS" then
		Response.Write("��<a href=Down_Label.asp target=_self>���ر�ǩ����</a>") 
	elseif Request.QueryString("TF") = "SD" then
		Response.Write("��<a href=supply_Label.asp target=_self>���ر�ǩ����</a>") 
	elseif Request.QueryString("TF") = "HS" then
		Response.Write("��<a href=House_Label.asp target=_self>���ر�ǩ����</a>") 
	elseif Request.QueryString("TF") = "AP" then
		Response.Write("��<a href=job_Label.asp target=_self>���ر�ǩ����</a>") 
	elseif Request.QueryString("TF") = "MS" then
		Response.Write("��<a href=Mall_Label.asp target=_self>����</a>")
	else
		Response.Write("") 
	end if
	%><!--�����ı��༭ by sicend -->
        ��<a href="All_Label_Style.asp" target="_self">������ʽ</a>��<a href="Label_Style_Class.asp" target="_self">������ʽ����</a> <a href="../../help?Label=MF_Label_Creat" target="_blank" style="cursor:help;"><img src="../Images/_help.gif" width="50" height="17" border="0"></a>��<a href="javascript:history.back();">����</a></div></td>
  </tr>
  <tr class="hback" >
    <td valign="top" class="hback"><strong>����:</strong> <a href="All_Label_style.asp?Label_Sub=NS" title="��� ����ϵͳ ��ǩ������ʽ" target="_self">����ϵͳ</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=NS&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ��������ϵͳ��ǩ������ʽ��" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=NS&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ��������ϵͳ��ǩ������ʽ��" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a>��<a href="All_Label_style.asp?Label_Sub=DS" title="��� ����ϵͳ ��ǩ������ʽ" target="_self">����ϵͳ</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=DS&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ��������ϵͳ��ǩ������ʽ��" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=DS&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ��������ϵͳ��ǩ������ʽ��" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a>
      ��
      <%if Request.Cookies("FoosunSUBCookie")("FoosunSUBSD")=1 then%>
      <a href="All_Label_style.asp?Label_Sub=SD" title="��� ����ϵͳ ��ǩ������ʽ" target="_self">����ϵͳ</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=SD&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ��������ϵͳ��ǩ������ʽ��" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=SD&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ��������ϵͳ��ǩ������ʽ��" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a>
      ��
      <%end if%>
      <%if Request.Cookies("FoosunSUBCookie")("FoosunSUBHS")=1 then%>
      <a href="All_Label_style.asp?Label_Sub=HS" title="��� ����¥�� ��ǩ������ʽ" target="_self">����¥��</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=HS&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ��������ϵͳ��ǩ������ʽ��" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=HS&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ��������ϵͳ��ǩ������ʽ��" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a>
    <%end if%>��
    <a href="All_Label_style.asp?Label_Sub=CForm" title="��� �Զ���� ��ǩ������ʽ" target="_self">�Զ����</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=CForm&ClassID=<%= Request.QueryString("ClassId") %>" title="���ñ༭��ģʽ������Ա��½��ǩ������ʽ��" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=CForm&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ��������ϵͳ��ǩ������ʽ��" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a></td>
  </tr>
  <tr class="hback" >
    <td valign="top" class="hback"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong>
      <%if Request.Cookies("FoosunSUBCookie")("FoosunSUBAP")=1 then%>
      <a href="All_Label_style.asp?Label_Sub=AP" title="��� ��Ƹ��ְ ��ǩ������ʽ" target="_self">��Ƹ��ְ</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=AP&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ�����˲�ϵͳ��ǩ������ʽ��" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=AP&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ��������ϵͳ��ǩ������ʽ��" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a>��
      <%End if%>
      <%if Request.Cookies("FoosunSUBCookie")("FoosunSUBMS")=1 then%>
      <a href="All_Label_style.asp?Label_Sub=MS" title="��� �̳�B2C ��ǩ������ʽ" target="_self">�̳�B2C</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=MS&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ�����̳�ϵͳ��ǩ������ʽ��" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=MS&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ��������ϵͳ��ǩ������ʽ��" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a>��
      <%end if%>
      <a href="All_Label_style.asp?Label_Sub=ME" title="��� ��Աϵͳ ��ǩ������ʽ" target="_self">��Աϵͳ</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=ME&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ������Աϵͳ��ǩ������ʽ��" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=ME&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ��������ϵͳ��ǩ������ʽ��" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a>��
	  <a href="All_Label_style.asp?Label_Sub=Login" title="��� ��Ա��½ ��ǩ������ʽ" target="_self">��Ա��½</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=Login&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ������Ա��½��ǩ������ʽ��" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=Login&ClassID=<%= Request.QueryString("ClassId") %>" title="�����ı�ģʽ��������ϵͳ��ǩ������ʽ��" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a></td>
  </tr>
</table>
<!------2/1 by chen ��������������ʽ�ķ����ļ��� ����ȡ��������ʽ------------------------------------>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
 <tr>
    <td width="8%" class="xingmu"><div align="center">��ʽID��</div></td>
    <td width="32%" class="xingmu"><div align="center">��ʽ����</div></td>
    <td width="25%" class="xingmu"><div align="center">���ò鿴</div></td>
	<td width="14%" class="xingmu"><div align="center">����ϵͳ</div></td>
	<td width="21%" class="xingmu"><div align="center">����/����</div></td>
</tr>
  <%
  dim rs_class,str_ParentID
  if trim(Request.QueryString("ClassID"))<>"" then
		str_ParentID = " and ParentID="&Request.QueryString("ClassID")&""
  elseif not isnumeric(trim(Request.QueryString("ClassID"))) then
		str_ParentID = " and ParentID=0"
  else
		str_ParentID = " and ParentID=0"
  end if
  set rs_class=Conn.execute("select id,ClassName,ClassContent,ParentID From FS_MF_StyleClass where 1=1"&str_ParentID&" order by id desc")
  do while not rs_class.eof 
  %>
  <tr class="hback">
    <td valign="top"><div align="center"><img src="../Images/Folder/folder.gif" alt="�ļ���" width="20" height="16"></div></td>
    <td><a href="All_Label_Style.asp?ClassId=<% = rs_class("id")%>&ParentID=<%=rs_class("id")%>">
      <% = rs_class("ClassName")%>
      </a></td>
    <td><% = rs_class("ClassContent")%></td>
	<td></td>
	<td></td>
  </tr>
  <%
  rs_class.movenext
  loop
  rs_class.close:set rs_class = nothing
  %>
  <tr class="hback_1">
    <td colspan="7" height="2"></td>
  </tr>
  <%
	dim rs_stock,ClassId,LableClassID1
	if Request.QueryString("ClassId")<>"" then
		LableClassID1 = NoSqlHack(Request.QueryString("ClassID"))
	Else
		LableClassID1=0
	End if
	ClassId = " and LableClassID="&LableClassID1&""
	set rs_stock= Server.CreateObject(G_FS_RS)
	rs_stock.open "select ID,StyleName,Content,LableClassID From FS_MF_Labestyle Where ID > 0" & ClassId &" order by ID desc",Conn,1,1
	if rs_stock.eof then
	   rs_stock.close
	   set rs_stock=nothing
	   Response.Write"<TR  class=""hback""><TD colspan=""7""  class=""hback"" height=""40"">û�м�¼��</TD></TR>"
	 end if
	%>
<!---------2/2 by chen------------------------------------------------------------------------------------------>
<%
	  Select Case Request.QueryString("Action")
	  			Case "Add"
					Call Add()
				Case "Add_Save"
					Call Add_Save()
				Case else
					Call Main()
	End Select
	Sub Main()
	%>
<%
			dim tmp_Label_Sub,LableClassID
	
			Set obj_Label_Rs = server.CreateObject(G_FS_RS)
			if trim(Label_Sub) <>"" then:tmp_Label_Sub = "and StyleType='"& Label_Sub &"'":else:tmp_Label_Sub = "":end if
			if Request.QueryString("ClassId")<>"" then
				LableClassID = " and LableClassID = " & NoSqlHack(Request.QueryString("ClassID"))
			Else
				LableClassID= " and (LableClassID < 1 Or LableClassID is Null)"
			End if
			SQL = "Select  ID,StyleName,LoopContent,Content,AddDate,StyleType,LableClassID from FS_MF_Labestyle where id>0 "& tmp_Label_Sub & LableClassID & " Order by id desc"
			obj_Label_Rs.Open SQL,Conn,1,3
			If not obj_Label_Rs.Eof Then
				obj_Label_Rs.PageSize=int_RPP
				cPageNo=NoSqlHack(Request.QueryString("page"))
				If cPageNo="" Then cPageNo = 1
				If not isnumeric(cPageNo) Then cPageNo = 1
				cPageNo = Clng(cPageNo)
				If cPageNo>obj_Label_Rs.PageCount Then cPageNo=obj_Label_Rs.PageCount 
				If cPageNo<=0 Then cPageNo=1		
				obj_Label_Rs.AbsolutePage=cPageNo
			
				For int_Start=1 TO int_RPP  
			%>
  <tr class="hback">
<td width="8%" class="hback" align="center"><% = obj_Label_Rs("ID") %></td>
    <td width="32%" class="hback">�� <a href="All_Label_style.asp?Action=Add&type=edit&id=<%= obj_Label_Rs("id")%>&Label_Sub=<%= obj_Label_Rs("StyleType")%>&ClassID=<% If Request.QueryString("ClassId")<>"" And IsNumeric(Request.QueryString("ClassId")) Then : Response.Write Request.QueryString("ClassId") : Else : Response.Write "0" : End IF %>"  target="_self">
      <% = obj_Label_Rs("StyleName") %>
    </a></td>
    <td width="25%" class="hback"   id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(Label<%=obj_Label_Rs("ID")%>);"  language=javascript>������ʽ�鿴</td>
    <td width="14%"><% = obj_Label_Rs("StyleType")%>
    </td>
    <td width="21%"><a href="All_Label_style.asp?id=<%=obj_Label_Rs("ID")%>&Label_Sub=<%=obj_Label_Rs("StyleType")%>&DelTF=1" onClick="{if(confirm('ȷ��ɾ����������ʽ��')){return true;}return false;}">ɾ��</a> <a href="All_Label_style_txt.asp?Action=Add&type=edit&id=<%= obj_Label_Rs("id")%>&Label_Sub=<%= obj_Label_Rs("StyleType")%>&ClassID=<% If Request.QueryString("ClassId")<>"" And IsNumeric(Request.QueryString("ClassId")) Then : Response.Write Request.QueryString("ClassId") : Else : Response.Write "0" : End IF %>"  target="_self">TXT</a></td>
  </tr>
  <tr id="Label<%=obj_Label_Rs("ID")%>" style="display:none"  class="hback">
    <td height="42" colspan="7"  class="hback"><table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
        <tr>
          <td height="48" class="hback"><%
			Dim regEx,result
			Set regEx = New RegExp '
			regEx.Pattern = "<img(.+?){(.+?)}(.+?)>" '  
			regEx.IgnoreCase = true ' 
			regEx.Global = True '  
			result = regEx.replace(obj_label_Rs("Content"),"<img src='../images/default.png'/>") 
			Response.Write(result)
		  %>
          </td>
        </tr>
      </table></td>
  </tr>
  <%
				obj_Label_Rs.MoveNext
				If obj_Label_Rs.Eof or obj_Label_Rs.Bof Then Exit For
			Next
			'Response.Write "<tr><td class=""hback"" colspan=""7"" align=""left"">"&fPageCount(obj_Label_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ȫѡ<input type=""checkbox"" name=""Checkallbox"" onclick=""javascript:CheckAll('Checkallbox');"" value=""0""></td></tr>"
		End If
			%>
  <tr  class="hback">
    <td height="21" colspan="7"  class="hback"><span class="tx">�����ʽ�����޸���ʽ</span> </td>
  </tr>
</table>
<div align="center"></div>
<p>
  <%
	End Sub
	%>
  <%
	  Sub Add()
	  	  dim str_id,tmp_id,tmp_StyleName,tmp_Content,tmp_Action,labelclassidd
		  str_id = NoSqlHack(Request.QueryString("id"))
		  if Request.QueryString("type")="edit" then
		  	if NoSqlHack(Request("IsPostBack")) <> "1" then
				Set obj_Label_Rs = server.CreateObject(G_FS_RS)
				obj_Label_Rs.Open "Select  ID,StyleName,LoopContent,Content,AddDate,StyleType,LableClassID from FS_MF_Labestyle where id="& str_id &"",Conn,1,3
				tmp_id = obj_Label_Rs("id")
				tmp_StyleName = obj_Label_Rs("StyleName")
				tmp_Content = obj_Label_Rs("Content")
				labelclassidd=obj_Label_Rs("LableClassID")
			end if
			tmp_Action = "Add_edit"
		  Else
			tmp_id = ""
			tmp_StyleName = ""
			tmp_Content = ""
			labelclassidd=""
			tmp_Action = "Add_save"
		  End if
		  if Request.Form("IsPostBack") = "1" then
			tmp_id = NoSqlHack(Request("id"))
			tmp_StyleName = NoSqlHack(Request("StyleName"))
			tmp_Content = NoSqlHack(Request("TxtFileds"))
			labelclassidd = NoSqlHack(Request("LableClassID"))
		  end if
		  
	%>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="xingmu">
    <td colspan="2" class="xingmu">������ǩ������ʽ(���������<span class="tx"><% = GetStyleMaxNum %></span>����ʽ)</td>
  </tr>
  <form name="Label_Form" method="post" action="" target="_self">
    <tr class="hback">
      <td width="13%"><div align="right"> ��ʽ����</div></td>
      <td width="87%"><input name="StyleName" type="text" id="StyleName" size="40" value="<% = tmp_StyleName %>">
        <input name="id" type="hidden" id="id" value="<% = tmp_id %>">
        <!----2/1/ by chen  ѡ����ʽ����  ��ʼ---------------->
        <select name="LableClassID" id="LableClassID">
          <option value="0">ѡ��������Ŀ</option>
          <%
				  dim class_rs_obj,str
				  set class_rs_obj=Conn.execute("select id,ParentID,ClassName From FS_MF_StyleClass where ParentID=0 order by id desc")
				  do while not class_rs_obj.eof
						If CStr(labelclassidd)=CStr(class_rs_obj("id")) Then 
							response.Write "<option value="""&class_rs_obj("id")&""" selected >"&class_rs_obj("ClassName")&"</option>"
						Else
							response.Write "<option value="""&class_rs_obj("id")&""">"&class_rs_obj("ClassName")&"</option>"
						End If 
						if labelclassidd = class_rs_obj("id") then
							response.Write "selected"
						end if
						response.Write get_childList(class_rs_obj("id"),"")
					class_rs_obj.movenext
				  loop
				  class_rs_obj.close:set class_rs_obj=nothing
				  %>
        </select>
        <!----2/1/ by chen  ѡ����ʽ���� ����---------------->
      </td>
    </tr>
    <tr class="hback">
      <td><div align="right">�����ֶ�</div></td>
      <td><%
			Dim Label_Sub
			Label_Sub = NoSqlHack(Request.QueryString("Label_Sub"))
			select case Label_Sub
					case "NS"
						Call NS_select()
					case "DS"
						Call DS_select()
					case "SD"
						Call SD_select()
					case "HS"
						Call HS_select()
					case "AP"
						Call AP_select()
					case "MS"
						Call MS_select()
					case "ME"
						Call ME_select()
					Case "Login"
						Call ME_Login()
					Case "CForm"
						Call MF_CustomForm()
					case else
						Call NS_select()
			end select
			%>
    </tr>
    <tr class="hback" <%if request.QueryString("Label_Sub")<>"DS" then response.Write("style=""display:'none';""") else response.Write("style=""display:'';"" ") end if%>> </tr>
    <tr class="hback">
      <td><div align="right">��ʽ����</div></td>
      <td><!--�༭����ʼ,�ı���ʽ by sicend-->
      <textarea rows="20" style="width:100%;" name="TxtFileds" ONSELECT="this.pos=document.selection.createRange();" onCLICK="this.pos=document.selection.createRange();" onKEYUP="this.pos=document.selection.createRange();"><%=Server.HTMLEncode(tmp_Content)%></textarea>
        <!--�༭������-->
      </td>
    </tr>
    <tr class="hback">
      <td>&nbsp;</td>
      <td><span class="tx">�ر����ѡ����ڸ߼�������Ա����html�൱��Ϥ����Ա����������ڱ�ǩ������ʽ������������дhtml��ʽ���ﵽ����ǰ̨ҳ���Ч�������ڶ�html��̫��Ϥ����Ա����鿴�����ĵ����߲���򵥵ı�ǩ��ʽ�Ϳ��ԣ�</span></td>
    </tr>
    <tr class="hback">
      <td>&nbsp;</td>
      <td><%
			 if tmp_Action = "Add_save" then
				dim obj_Count_rs_1,tmp_str,tmp_display
				Set obj_Count_rs_1 = server.CreateObject(G_FS_RS)
				obj_Count_rs_1.Open "Select StyleName,Content,AddDate from FS_MF_Labestyle where StyleType='NS' Order by id desc",Conn,1,3
				if Not obj_Count_rs_1.eof then
					if obj_Count_rs_1.recordcount>GetStyleMaxNum then
						tmp_str = "--��ǩ��ʽ�Ѿ�����" & GetStyleMaxNum & "��,���ܴ�����"
						tmp_display = "disabled"
					Else
						tmp_str = ""
						tmp_display = ""
					End if
				Else
					tmp_str = ""
					tmp_display = ""
				End if
			 Else
					tmp_str = ""
					tmp_display = ""
			 End if
			  %>
        <input type="submit" name="Submit" value="��HTML������ʽ<% = tmp_str %>"<% = tmp_display %> onClick="return Label_Form_sumit(this.form,1);">
        <input name="Action" type="hidden" id="Action" value="<% = tmp_Action %>" >
        <input name="Label_Sub" type="hidden" value="<%=Request.QueryString("Label_Sub")%>">
        <input type="hidden" name="IsPostBack" value="1">
        <input type="submit" name="Submit3" value="��XHTML������ʽ<% = tmp_str %>"<% = tmp_display %> onClick="return Label_Form_sumit(this.form,0);">
        <input type="reset" name="Submit2" value="����">
      </td>
    </tr>
  </form>
</table>
<script language="JavaScript" type="text/JavaScript">
function Label_Form_sumit(FormObj,IsHTML)
{
	if(FormObj.StyleName.value == "")
	{
		alert("����д��ǩ����")
		FormObj.StyleName.focus();
		return false;
	}

	if(FormObj.TxtFileds.value == "")
	{
		alert("�������ǩ��ʽ����")
		return false;
	}
	return true;
}

function Insertlabel_Sel(Lable_obj)
{
	if(Lable_obj.options[Lable_obj.selectedIndex].value==''){
	return false;
	}else{
	InsertEditor(Lable_obj.options[Lable_obj.selectedIndex].value);
	}
}
function InsertEditor(InsertValue)
{
	//Label_Form.TxtFileds.value = Label_Form.TxtFileds.value+InsertValue;//�޸��ı���ӱ�ǩ by sicend
	try
	    {
	        Label_Form.TxtFileds.pos.text=InsertValue;
	    }
	    catch(e)
	    {}
}
</script>
<%
End Sub
sub NS_select()
			'�õ��Զ����ֶ�
			dim ns_D_rs,ns_list
			ns_list = ""
			set ns_D_rs = Server.CreateObject(G_FS_RS)
			ns_D_rs.open "select D_Coul,D_Name From FS_MF_DefineTable where D_SubType='NS'",Conn,1,3
			if ns_D_rs.eof then
				ns_list =ns_list& "<option value="""">û���Զ����ֶ�</option>"
				ns_D_rs.close:set ns_D_rs=nothing
			else
				do while not ns_D_rs.eof 
					ns_list = ns_list & "<option value=""{NS=Define|"&ns_D_rs("D_Coul")&"}"">"& ns_D_rs("D_Name")&"</option>"
					ns_D_rs.movenext
				loop
				ns_D_rs.close:set ns_D_rs=nothing
			end if
			%>
<select name="NewsFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">���������ֶΩ���</option>
  <option value="{NS:FS_ID}">�Զ����</option>
  <option value="{NS:FS_NewsID}">NewsID</option>
  <option value="<a href={NS:FS_NewsURL}>{NS:FS_NewsTitle}</a>">���ű���(�ض�)������ʹ����title��</option>
  <option value="{NS:FS_NewsTitleAll}">������������(����New��־)</option>
  <option value="{NS:FS_NewsURL}"> ���ŷ���·��</option>
  <option value="{NS:FS_CurtTitle}"> ���Ÿ�����</option>
  <option value="{NS:FS_NewsNaviContent}"> ���ŵ���</option>
  <option value="{NS:FS_Content}"> ��������</option>
  <option value="{NS:FS_AddTime}"> �����������</option>
  <option value="{NS:FS_Author}"> ��������</option>
  <option value="{NS:FS_Editer}"> �������α༭</option>
  <option value="" style="background:#88AEFF;color:000000">����Ԥ�����ֶΩ���</option>
  <option value="{NS:FS_hits}">�����</option>
  <option value="{NS:FS_KeyWords}">�ؼ���</option>
  <option value="{NS:FS_TxtSource}"> ������Դ</option>
  <option value="{NS:FS_SmallPicPath}">ͼƬ���ŵ�ͼƬ��ַ(Сͼ)</option>
  <option value="{NS:FS_PicPath}">ͼƬ���ŵ�ͼƬ��ַ(��ͼ)</option>
  <option value="{NS:FS_FormReview}">���۱�</option>
  <option value="{NS:FS_ReviewURL}">��������(����ַ)</option>
  <option value="{NS:FS_ShowComment}">��ʾ�����б�</option>
  <option value="{NS:FS_AddFavorite}">�����ղ�</option>
  <option value="{NS:FS_SendFriend}">���͸�����</option>
  <option value="{NS:FS_SpecialList}">����ר���б�</option>
  <option value="{NS:FS_PrevPage}"> ��һƪ����</option>
  <option value="{NS:FS_NextPage}"> ��һƪ����</option>
  <option value="" style="background:#88AEFF;color:000000">����ר��ɶ����ֶΩ���</option>
  <option value="{NS:FS_SpecialName}">ר����������</option>
  <option value="" style="background:#88AEFF;color:000000">���������Զ����ֶΩ���</option>
  <%=ns_list%>
</select>
<select name="SingleClassFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">������Ŀ�ɶ����ֶΩ���</option>
  <option value="{NS:FS_ClassName}">��Ŀ��������</option>
  <option value="{NS:FS_ClassURL}">��Ŀ����·��</option>
  <option value="{NS:FS_ClassNaviPicURL}">��Ŀ����ͼƬ��ַ</option>
  <option value="{NS:FS_ClassNaviDescript}">��Ŀ����˵��</option>
  <option value="" style="background:#88AEFF;color:000000">������ҳ��Ŀ�ɶ����ֶΩ���</option>
  <option value="{NS:FS_PageContent}">��Ŀ����</option>
  <option value="{NS:FS_Keywords}">��ĿMETA�ؼ���</option>
  <option value="{NS:FS_description}">��ĿMETA����</option>
</select>
</td>
<%end sub%>
<%sub DS_select()
			'�õ��Զ����ֶ�
			dim ds_D_rs,ds_list
			ds_list = ""
			set ds_D_rs = Server.CreateObject(G_FS_RS)
			ds_D_rs.open "select D_Coul,D_Name From FS_MF_DefineTable where D_SubType='DS'",Conn,1,3
			if ds_D_rs.eof then
				ds_list =ds_list& "<option value="">û���Զ����ֶ�</option>"
				ds_D_rs.close:set ds_D_rs=nothing
			else
				do while not ds_D_rs.eof 
					ds_list = ds_list & "<option value=""{DS=Define|"&ds_D_rs("D_Coul")&"}"">"& ds_D_rs("D_Name")&"</option>"
					ds_D_rs.movenext
				loop
				ds_D_rs.close:set ds_D_rs=nothing
			end if%>
<select name="NewsFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">���������ֶΩ���</option>
  <option value="{DS:FS_ID}">�Զ����</option>
  <option value="{DS:FS_DownLoadID}">DownLoadID</option>
  <option value="<a href={DS:FS_DownURL}>{DS:FS_Name}</a>">���ر���(�ض�)</option>
  <option value="<a href={DS:FS_DownURL}>{DS:FS_NameAll}</a>">���ر���(����)</option>
  <option value="{DS:FS_Description}">���ؼ��</option>
  <option value="{DS:FS_AddTime}">���ʱ��</option>
  <option value="{DS:FS_EditTime}">�޸�ʱ��</option>
  <option value="{DS:FS_SystemType}">ϵͳƽ̨</option>
  <option value="{DS:FS_Accredit}">������Ȩ</option>
  <option value="{DS:FS_Version}">�汾</option>
  <option value="{DS:FS_Appraise}">�Ǽ�����</option>
  <option value="{DS:FS_FileSize}">�ļ���С</option>
  <option value="{DS:FS_Language}">����</option>
  <option value="{DS:FS_PassWord}">��ѹ����</option>
  <!--<option value="{DS:FS_Property}">��������</option>-->
  <option value="{DS:FS_Provider}">������</option>
  <option value="{DS:FS_ProviderUrl}">�ṩ��Url��ַ</option>
  <option value="{DS:FS_EMail}">��ϵ��EMAIL</option>
  <option value="{DS:FS_Types}">��������</option>
  <option value="{DS:FS_OverDue}">��������</option>
  <option value="{DS:FS_ConsumeNum}">���ѵ���</option>
  <option value="{DS:FS_Address$&amp;lt;br /&amp;gt;}">���ص�ַ</option>
  <option value="" style="background:#88AEFF;color:000000">����Ԥ�����ֶΩ���</option>
  <option value="{DS:FS_Hits}">�����</option>
  <option value="{DS:FS_ClickNum}">���ش���</option>
  <option value="{DS:FS_Pic}">��ʾͼƬ��ַ</option>
  <option value="{DS:FS_ReviewURL}">��������(����ַ)</option>
  <option value="{DS:FS_FormReview}">���۱�</option>
  <option value="{DS:FS_ShowComment}">��ʾ�����б�</option>
  <option value="{DS:FS_SpecialList}">����ר���б�</option>
  <option value="{DS:FS_AddFavorite}">�����ղ�</option>
  <option value="{DS:FS_SendFriend}">���͸�����</option>
  <option value="{DS:FS_DownURL}">���ط���·��</option>
  <option value="" style="background:#88AEFF;color:000000">������Ŀ�ɶ����ֶΩ���</option>
  <option value="{DS:FS_ClassName}">��Ŀ��������</option>
  <option value="{DS:FS_ClassURL}">��Ŀ����·��</option>
  <option value="{DS:FS_ClassNaviPicURL}">��Ŀ����ͼƬ��ַ</option>
  <option value="{DS:FS_ClassNaviDescript}">��Ŀ����˵��</option>
  <option value="{DS:FS_ClassKeywords}">��Ŀ�ؼ���</option>
  <option value="{DS:FS_Classdescription}">��Ŀ����</option>
  <option value="" style="background:#88AEFF;color:000000">����ר���ɶ����ֶΩ���</option>
  <option value="{DS:FS_SpecialName}">ר����������</option>
  <option value="" style="background:#88AEFF;color:000000">���������Զ����ֶΩ���</option>
  <%=ds_list%>
</select>
<%end sub%>
<%sub SD_select()
'�õ��Զ����ֶ�
			dim sd_D_rs,sd_list
			sd_list = ""
			set sd_D_rs = Server.CreateObject(G_FS_RS)
			sd_D_rs.open "select D_Coul,D_Name From FS_MF_DefineTable where D_SubType='SD'",Conn,1,3
			if sd_D_rs.eof then
				sd_list =sd_list& "<option value="">û���Զ����ֶ�</option>"
				sd_D_rs.close:set sd_D_rs=nothing
			else
				do while not sd_D_rs.eof 
					sd_list = sd_list & "<option value=""{SD=Define|"&sd_D_rs("D_Coul")&"}"">"& sd_D_rs("D_Name")&"</option>"
					sd_D_rs.movenext
				loop
				sd_D_rs.close:set sd_D_rs=nothing
			end if%>
<select name="NewsFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">���������ֶΩ���</option>
  <option value="&lt;a href=&quot;{SD:FS_URL}&quot; target=_blank&gt;{SD:FS_title}">����</option>
  <option value="{SD:FS_Alltitle}">��������(���ض�)</option>
  <option value="{SD:FS_URL}">��������·��</option>
  <option value="{SD:FS_PubType}">����</option>
  <option value="{SD:FS_PubTypeLink}">�����ӵ�����</option>
  <option value="{SD:FS_PubContent}">����</option>
  <option value="{SD:FS_AreaID}">��������</option>
  <option value="{SD:FS_ClassID}">��������</option>
  <option value="{SD:FS_Keyword}">�ؼ���</option>
  <option value="{SD:FS_CompType}">��Ӫ��ʽ</option>
  <option value="{SD:FS_PubNumber}">��Ʒ����</option>
  <option value="{SD:FS_PubPrice}">��Ʒ�۸�</option>
  <option value="{SD:FS_PubPack}">��װ˵��</option>
  <option value="{SD:FS_Pubgui}">��Ʒ���</option>
  <option value="{SD:FS_PubPic_1}">ͼƬһ��ַ</option>
  <option value="{SD:FS_PubPic_2}">ͼƬ��</option>
  <option value="{SD:FS_PubPic_3}">ͼƬ��</option>
  <option value="{SD:FS_Addtime}">����ʱ��</option>
  <option value="{SD:FS_EditTime}">������ʱ��</option>
  <option value="{SD:FS_ValidTime}">��Чʱ��</option>
  <option value="{SD:FS_PubAddress}">����</option>
  <option value="{SD:FS_Fax}">��ϵ����</option>
  <option value="{SD:FS_User}">������Ա�û���</option>
  <option value="{SD:FS_tel}">��ϵ�绰</option>
  <option value="{SD:FS_Mobile}">�ƶ��绰</option>
  <option value="{SD:FS_otherLink}">������ϵ��ʽ</option>
  <option value="{SD:FS_hits}">���</option>
  <!--option value="{SD:FS_ReviewURL}">��������(����ַ)</option-->
  <option value="{SD:FS_review}">��������</option>
  <option value="{SD:FS_reviewcontent}">��������</option>
  <option value="" style="background:#88AEFF;color:000000">���������Զ����ֶΩ���</option>
  <%=sd_list%>
</select>
</span>
<%end sub%>
<%sub HS_select()
				dim hs_D_rs,hs_list
				hs_list = ""
				set hs_D_rs = Server.CreateObject(G_FS_RS)
				hs_D_rs.open "select D_Coul,D_Name From FS_MF_DefineTable where D_SubType='HS'",Conn,1,3
				if hs_D_rs.eof then
					hs_list =hs_list& "<option value="""">û���Զ����ֶ�</option>"
					hs_D_rs.close:set hs_D_rs=nothing
				else
					do while not hs_D_rs.eof 
						hs_list = hs_list & "<option value=""{HS=Define|"&hs_D_rs("D_Coul")&"}"">"& hs_D_rs("D_Name")&"</option>"
						hs_D_rs.movenext
					loop
					hs_D_rs.close:set hs_D_rs=nothing
				end if
			  %>
<select name="HouseFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">���������ֶΩ���</option>
  <option value="{all:HS_FS_ID}">�Զ����</option>
  <option value="{all:HS_FS_Price}">����</option>
  <option value="{all:HS_FS_PubDate}">����ʱ��</option>
  <option value="{all:HS_FS_UserNumber}">�����߱��</option>
  <option value="" style="background:#88AEFF;color:000000">����Ԥ�����ֶΩ���</option>
  <option value="{HS_FS_FormReview}">���۱�</option>
  <option value="{HS_FS_ReviewURL}">����������ַ</option>
  <option value="{HS_FS_ShowComment}">��ʾ�����б�</option>
  <option value="{HS_FS_AddFavorite}">�����ղ�</option>
  <option value="{HS_FS_SendFriend}">���͸�����</option>
  <option value="{HS_FS_HouseURL}"> ��Ϣ����·��</option>
  <option value="" style="background:#88AEFF;color:000000">���������Զ����ֶΩ���</option>
  <%=hs_list%>
</select>
<select name="LouFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">����¥����Ϣ����</option>
  <option value="{Lou:HS_FS_HouseName}">¥������</option>
  <option value="{Lou:HS_FS_KaiFaShang}">������</option>
  <option value="{Lou:HS_FS_Position}">¥��λ��</option>
  <option value="{Lou:HS_FS_Direction}">¥�̷�λ</option>
  <option value="{Lou:HS_FS_Class}">��Ŀ���</option>
  <option value="{Lou:HS_FS_OpenDate}">��������</option>
  <option value="{Lou:HS_FS_PreSaleRange}">Ԥ�۷�Χ</option>
  <option value="{Lou:HS_FS_Status}">����״��</option>
  <option value="{Lou:HS_FS_introduction}">���ݽ���</option>
  <option value="{Lou:HS_FS_Contact}"> ��ϵ��ʽ</option>
  <option value="{Lou:HS_FS_hits}">�����</option>
  <option value="{Lou:HS_FS_Pic}">ͼƬ</option>
</select>
<select name="SecondFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">����������Ϣ����</option>
  <option value="{Second:HS_FS_UseFor}">��;</option>
  <option value="{Second:HS_FS_Label}">���ݱ��</option>
  <option value="{Second:HS_FS_FloorType}">סլ���</option>
  <option value="{Second:HS_FS_BelongType}">��Ȩ����</option>
  <option value="{Second:HS_FS_HouseStyle}">����</option>
  <option value="{Second:HS_FS_Structure}">�����ṹ</option>
  <option value="{Second:HS_FS_Area}">�������</option>
  <option value="{Second:HS_FS_BuildDate}">�������</option>
  <option value="{Second:HS_FS_CityArea}">����</option>
  <option value="{Second:HS_FS_Address}">��ַ</option>
  <option value="{Second:HS_FS_Floor}">¥��</option>
  <option value="{Second:HS_FS_Decoration}">װ�����</option>
  <option value="{Second:HS_FS_equip}">������ʩ</option>
  <option value="{Second:HS_FS_Remark}">��ע</option>
  <option value="{Second:HS_FS_LinkMan}">��ϵ��</option>
  <option value="{Second:HS_FS_Contact}">��ϵ��ʽ</option>
  <option value="{Second:HS_FS_hits}">�����</option>
  <option value="{Second:HS_FS_Pic}">ͼƬ</option>
</select>
<select name="TenancyFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">����������Ϣ����</option>
  <option value="{Tenancy:HS_FS_UseFor}">ʹ������</option>
  <option value="{Tenancy:HS_FS_XingZhi}">��������</option>
  <option value="{Tenancy:HS_FS_Class}">����</option>
  <option value="{Tenancy:HS_FS_ZaWuJian}">�����</option>
  <option value="{Tenancy:HS_FS_CityArea}">����</option>
  <option value="{Tenancy:HS_FS_HouseStyle}">����</option>
  <option value="{Tenancy:HS_FS_Area}">�������</option>
  <option value="{Tenancy:HS_FS_Period}">��Ч��</option>
  <option value="{Tenancy:HS_FS_XiaoQuName}">С������</option>
  <option value="{Tenancy:HS_FS_Position}">��Դ��ַ</option>
  <option value="{Tenancy:HS_FS_JiaoTong}">��ͨ״��</option>
  <option value="{Tenancy:HS_FS_Floor}">¥��</option>
  <option value="{Tenancy:HS_FS_BuildDate}">�������</option>
  <option value="{Tenancy:HS_FS_Decoration}">װ�����</option>
  <option value="{Tenancy:HS_FS_equip}">������ʩ</option>
  <option value="{Tenancy:HS_FS_Remark}">��ע</option>
  <option value="{Tenancy:HS_FS_LinkMan}">��ϵ��</option>
  <option value="{Tenancy:HS_FS_Contact}">��ϵ��ʽ</option>
  <option value="{Tenancy:HS_FS_hits}">�����</option>
  <option value="{Tenancy:HS_FS_Pic}">ͼƬ</option>
</select>
<%end sub%>
<%sub AP_select()
				dim ap_D_rs,ap_list
				ap_list = ""
				set ap_D_rs = Server.CreateObject(G_FS_RS)
				ap_D_rs.open "select D_Coul,D_Name From FS_MF_DefineTable where D_SubType='AP'",Conn,1,3
				if ap_D_rs.eof then
					ap_list =ap_list& "<option value="""">û���Զ����ֶ�</option>"
					ap_D_rs.close:set ap_D_rs=nothing
				else
					do while not ap_D_rs.eof 
						ap_list = ap_list & "<option value=""{AP=Define|"&ap_D_rs("D_Coul")&"}"">"& ap_D_rs("D_Name")&"</option>"
						ap_D_rs.movenext
					loop
					ap_D_rs.close:set ap_D_rs=nothing
				end if			  
			  %>
<select name="APFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">����Ԥ�����ֶΩ���</option>
  <option value="{AP:FS_AddFavorite}">�����ղ�</option>
  <option value="{AP:FS_SendFriend}">���͸�����</option>
  <option value="" style="background:#88AEFF;color:000000">�����˲��Զ����ֶΩ���</option>
  <%=ap_list%>
</select>
<SELECT name="APFields_1" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">����Ƹ��Ϣ�����ֶΩ�</option>
  <option value="{INV:AP:FS_ID}">�Զ����</option>
  <option value="{INV:AP:FS_UserNumber}">��˾����</option>
  <!-----------------2/2 by chen------------------------------------------->
  <option value="{INV:AP:FS_Mobile}">��˾�绰</option>
  <option value="{INV:AP:FS_Fax}">��˾����</option>
  <option value="{INV:AP:FS_Address}">��˾��ַ</option>
  <option value="{INV:AP:FS_ConnectPer}">��˾��ϵ��</option>
  <option value="{INV:AP:FS_WebSit}">��˾��վ</option>
  <!------------------2/2  by chen----------------------------------------->
  <option value="{INV:AP:FS_JobName}">ְλ����</option>
  <option value="{INV:AP:FS_JobDescription}">ְλ����</option>
  <option value="{INV:AP:FS_ResumeLang}">������������</option>
  <option value="{INV:AP:FS_WorkCity}">�����ص�</option>
  <option value="{INV:AP:FS_PublicDate}">��������</option>
  <option value="{INV:AP:FS_EndDate}">��Ч����</option>
  <option value="{INV:AP:FS_NeedNum}">��Ƹ����</option>
  <option value="{INV:AP:FS_APURL}"> ��Ϣ����·��</option>
</SELECT>
<SELECT name="APFields_2" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">���˲���Ϣ�����ֶΩ�</option>
  <option value="{TO:AP:FS_ID}">�Զ����</option>
  <option value="{TO:AP:FS_UserNumber}">�û����</option>
  <option value="{TO:AP:FS_UserName}">�û���</option>
  <option value="{TO:AP:FS_PersonURL}">�û���ְ����·��</option>
  <option value="{TO:AP:FS_JobReadURL}">�鿴�û�����·��</option>
  <option value="{TO:AP:FS_Sex}">�Ա�</option>
  <option value="{TO:AP:FS_Pic}">ͼƬ</option>
  <option value="{TO:AP:FS_Birthday}">����</option>
  <option value="{TO:AP:FS_CertificateClass}">֤������</option>
  <option value="{TO:AP:FS_CertificateNo}">֤������</option>
  <option value="{TO:AP:FS_CurrentWage}">Ŀǰ��н</option>
  <option value="{TO:AP:FS_CurrencyType}">����</option>
  <option value="{TO:AP:FS_WorkAge}">��������</option>
  <option value="{TO:AP:FS_Province}">����ʡ</option>
  <option value="{TO:AP:FS_City}">������</option>
  <option value="{TO:AP:FS_HomeTel}">��ͥ�绰</option>
  <option value="{TO:AP:FS_CompanyTel}">��˾�绰</option>
  <option value="{TO:AP:FS_Mobile}">�ֻ�</option>
  <option value="{TO:AP:FS_Email}">�����ʼ�</option>
  <option value="{TO:AP:FS_QQ}">QQ</option>
  <option value="{TO:AP:FS_click}">�����</option>
  <option value="{TO:AP:FS_lastTime}">����޸�ʱ��</option>
  <option value="{TO:AP:FS_ShenGao}">���</option>
  <option value="{TO:AP:FS_XueLi}">ѧ��</option>
  <option value="{TO:AP:FS_HowDay}">��ÿ��Ե���</option>
  <option value="">--������λ--</option>
  <option value="{TO:AP_1:FS_Job}">������λ</option>
</SELECT>
<%end sub%>
<%sub MS_select()
			'�õ��Զ����ֶ�
			dim ms_D_rs,ms_list
			ms_list = ""
			set ms_D_rs = Server.CreateObject(G_FS_RS)
				ms_D_rs.open "select D_Coul,D_Name From FS_MF_DefineTable where D_SubType='MS'",Conn,1,1
			if ms_D_rs.eof then
				ms_list ="<option value="""">û���Զ����ֶ�</option>"
				ms_D_rs.close:set ms_D_rs=nothing
			else
				do while not ms_D_rs.eof 
					ms_list = ms_list & "<option value=""{MS=Define|"&ms_D_rs("D_Coul")&"}"">"& ms_D_rs("D_Name")&"</option>"
					ms_D_rs.movenext
				loop
				ms_D_rs.close:set ms_D_rs=nothing
			end if
			%>
<select name="ProductsFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">���������ֶΩ���</option>
  <option value="{MS:FS_ID}">�Զ����</option>
  <option value="{MS:FS_ProductTitle}"> ��Ʒ����(�ض�)</option>
  <option value="{MS:FS_ProductTitleAll}"> ��Ʒ���ƣ�������</option>
  <option value="{MS:FS_Barcode}"> ��Ʒ������</option>
  <option value="{MS:FS_Serialnumber}">��Ʒ���к�</option>
  <option value="{MS:FS_ProductURL}">��Ʒ�������·��</option>
  <option value="{MS:FS_Stockpile}"> ��Ʒ���</option>
  <option value="{MS:FS_OldPrice}"> �г��۸�</option>
  <option value="{MS:FS_NewPrice}"> ʵ�ʼ۸�</option>
  <!--�������-->
  <option value="{MS:FS_Mail_money}">�������</option>
  <option value="{MS:FS_NowMoney}">��������ú�ļ۸�</option>
  <!--�������-->
  <option value="{MS:FS_ProductContent}"> ��Ʒ����</option>
  <option value="{MS:FS_RepairContent}"> ��������</option>
  <option value="{MS:FS_AddTime}"> ��Ʒ�������</option>
  <option value="{MS:FS_AddMember}"> ��Ʒ�����</option>
  <option value="{MS:FS_ProductAddress}"> ��Ʒ����</option>
  <option value="{MS:FS_MakeFactory}"> ��������</option>
  <option value="{MS:FS_MakeTime}"> ��������</option>
  <option value="{MS:FS_saleNumber}"> �۳�����</option>
  <option value="{MS:FS_SaleStyle}"> ������ʽ</option>
  <option value="{MS:FS_DiscountStartDate}"> ���ۿ�ʼʱ��</option>
  <option value="{MS:FS_DiscountEndDate}"> ���۽���ʱ��</option>
  <option value="{MS:FS_Discount}"> �ۿ���</option>
  <option value="" style="background:#88AEFF;color:000000">�����ɶ����ֶΩ���</option>
  <option value="{MS:FS_hits}">�����</option>
  <option value="{MS:FS_KeyWords}">�ؼ���(�����ӵ�����)</option>
  <option value="{MS:FS_TitleKeyWords}">�ؼ���(����������������)</option>
  <!--option value="{MS:FS_TxtSource}"> ��Ʒ��Դ</option-->
  <option value="{MS:FS_SmallPicPath}">��ƷͼƬ(Сͼ)</option>
  <option value="{MS:FS_PicPath}">��ƷͼƬ(��ͼ)</option>
  <option value="{MS:FS_ShopBagURL}">���ﳵ��ַ</option>
  <option value="{MS:FS_FormReview}">���۱�</option>
  <option value="{MS:FS_ReviewTF}">�������ʾ��������</option>
  <option value="{MS:FS_ShowComment}">��ʾ�����б�</option>
  <option value="{MS:FS_AddFavorite}">�����ղ�</option>
  <option value="{MS:FS_SendFriend}">���͸�����</option>
  <option value="{MS:FS_SpecialList}">����ר���б�</option>
  <!--<option value="{MS:FS_ProductURL}"> ���ŷ���·��</option>-->
  <!--<option value="" style="background:#88AEFF;color:000000">������Ŀ�ɶ����ֶΩ���</option>-->
  <option value="{MS:FS_ClassName}">��Ŀ��������</option>
  <option value="{MS:FS_ClassURL}">��Ŀ�������·��</option>
  <option value="{MS:FS_ClassNaviPicURL}">��Ŀ����ͼƬ</option>
  <option value="{MS:FS_ClassNaviContent}">��Ŀ����˵��</option>
  <option value="{MS:FS_ClassKeywords}">��Ŀ�ؼ���(�����̳���Ŀ���ռ��б��Ż���������)</option>
  <option value="{MS:FS_Classdescription}">��Ŀ����(�����̳���Ŀ���ռ��б��Ż���������)</option>
  <!--<option value="{MS:FS_Classdescription}">��Ŀ����</option>-->
  <!--<option value="" style="background:#88AEFF;color:000000">����ר��ɶ����ֶΩ���</option>-->
  <option value="{MS:FS_SpecialName}">ר����������</option>
  <!-- <option value="{MS:FS_SpecialURL}">ר���������·��</option>-->
  <option value="{MS:FS_SpecialNaviPicURL}">ר�⵼��ͼƬ</option>
  <option value="{MS:FS_SpecialNaviDescript}">ר�⵼��˵��</option>
  <option value="" style="background:#88AEFF;color:000000">�����Զ����ֶΩ���</option>
  <%=ms_list%>
</select>
<%end sub
sub ME_select()
%>
<SELECT name="APFields_1" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">�����˻�Ա�����ֶΩ�</option>
  <option value="{ME:FS_UserNumber}">�û����</option>
  <option value="{ME:FS_UserName}">�û���</option>
  <option value="{ME:FS_NickName}">�û��ǳ�</option>
  <option value="{ME:FS_RealName}">��ʵ����</option>
  <option value="{ME:FS_Sex}">�Ա�</option>
  <option value="{ME:FS_HeadPic}">ͷ��</option>
  <option value="{ME:FS_tel}">�绰</option>
  <option value="{ME:FS_Email}">Email</option>
  <option value="{ME:FS_HomePage}">������ҳ</option>
  <option value="{ME:FS_QQ}">QQ</option>
  <option value="{ME:FS_MSN}">MSN</option>
  <option value="{ME:FS_Province}">ʡ��</option>
  <option value="{ME:FS_City}">����</option>
  <option value="{ME:FS_Address}">��ַ</option>
  <option value="{ME:FS_PostCode}">��������</option>
  <option value="{ME:FS_Vocation}">ְҵ</option>
  <option value="{ME:FS_BothYear}">��������</option>
  <option value="{ME:FS_Age}">����</option>
  <option value="{ME:FS_Integral}">����</option>
  <option value="{ME:FS_FS_Money}">���</option>
  <option value="{ME:FS_IsMarray}">���</option>
  <option value="{ME:FS_RegTime}">ע������</option>
</SELECT>
<SELECT name="APFields_2" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">����ҵ��Ա�����ֶΩ�</option>
  <option value="{ME:FS_C_Name}">��ҵ����</option>
  <option value="{ME:FS_C_ShortName}">��ҵ���</option>
  <option value="{ME:FS_C_logo}">��ҵLogo</option>
  <option value="{ME:FS_C_Tel}">�绰</option>
  <option value="{ME:FS_C_Fax}">����</option>
  <option value="{ME:FS_C_VocationClassID}">������ҵ</option>
  <option value="{ME:FS_C_WebSite}">��˾��վ</option>
  <option value="{ME:FS_C_Operation}">ҵ��Χ</option>
  <option value="{ME:FS_C_Products}">��˾��Ʒ</option>
  <option value="{ME:FS_C_Content}">��˾���</option>
  <option value="{ME:FS_C_Province}">ʡ��</option>
  <option value="{ME:FS_C_City}">����</option>
  <option value="{ME:FS_C_Address}">��ַ</option>
  <option value="{ME:FS_C_PostCode}">��������</option>
  <option value="{ME:FS_C_Vocation}">��ϵ��ְ��</option>
  <option value="{ME:FS_C_BankName}">��������</option>
  <option value="{ME:FS_C_BankUserName}">�����ʺ�</option>
  <option value="{ME:FS_C_property}">��˾����</option>
</SELECT>
<SELECT name="APFields_3" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">�������ֶΩ�</option>
  <option value="{ME:FS_UserURL}/ShowUser.asp?UserNumber={ME:FS_UserNumber}">��ϸ��ϢURL</option>
  <option value="{ME:FS_UserURL}/Message_write.asp?ToUserNumber={ME:FS_UserNumber}">����URL</option>
  <option value="{ME:FS_UserURL}/book_write.asp?ToUserNumber={ME:FS_UserNumber}&M_Type=0">����URL</option>
  <option value="{ME:FS_UserURL}/Friend_add.asp?type=0&UserName={ME:FS_UserName}">��Ϊ����URL</option>
  <option value="{ME:FS_UserURL}/UserReport.asp?action=report&ToUserNumber={ME:FS_UserNumber}">�ٱ�URL</option>
  <option value="{ME:FS_UserURL}/Corp_card_add.asp?UserNumber={ME:FS_UserNumber}">�ղ���ƬURL</option>
  <option value="{ME:FS_UserURL}/?User={ME:FS_UserNumber}">��Ա������ҳURL</option>
</SELECT>
<%end sub%>
<% Sub ME_Login() %>
<SELECT name="Select_Login" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000" selected="selected">ѡ���Ա��¼��ǩ����Ҫ���ֶ�</option>
  <option value="{Login_Name}" style="color:#FF0000;">�û��������(��ѡ)</option>
  <option value="{Login_Password}" style="color:#FF0000;">���������(��ѡ)</option>
  <option value="{Login_Simbut}" style="color:#FF0000;">��¼�ύ��ť(��ѡ)</option>
  <option value="{Login_Type}">��¼��ʽѡ���</option>
  <option value="{Login_Reset}">��¼ȡ����ť</option>
  <option value="{Reg_LinkUrl}">ע�����û�����</option>
  <option value="{Get_PassLink}">ȡ����������</option>
</SELECT>
<SELECT name="Login_Display" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000" selected="selected">ѡ���¼����ʾ�����ֶ�</option>
  <option value="{User_Name}">��Ա����</option>
  <option value="{User_JiFen}">��Ա����</option>
  <option value="{User_JinBi}">��Ա���</option>
  <option value="{User_LoginTimes}">��¼����</option>
  <option value="{User_TouGao}">Ͷ����</option>
  <option value="{User_ConCenter}">�����������</option>
  <option value="{User_LogOut}">�˳�����</option>
</SELECT>
<br />
<span style="margin-left:10px; text-align:left; font-size:12px; color:#FF0000;">���ڴ˴�����html����������ʾ��ʽ����������ڱ�ǩ����;��½��ʽ����ʾ��ʽ��<font color="#0033FF">"$*$"</font>�ָ�����½��ʽ $*$ ��ʾ��ʽ�������������ʾ����</span>
<% End Sub %>
<% Sub MF_CustomForm() %>
<SELECT name="CustomFormID" onChange="this.form.Action.value='';this.form.submit();">
  <option value="" style="background:#88AEFF;color:000000" selected="selected">ѡ���Զ����</option>
  <%
  Dim CustomFormRS,CustomFormItemRS,CustomFormID,CustomFormItemArray,i
  CustomFormID = NoSqlHack(Request("CustomFormID"))
  Set CustomFormRS = Conn.Execute("Select * from FS_MF_CustomForm")
  Do while Not CustomFormRS.Eof
  %>
  <option <% if CustomFormRS("ID") & "" = CustomFormID then Response.Write("selected") %> value="<% = CustomFormRS("ID") %>" style="color:#FF0000;"><% = CustomFormRS("formname") %></option>
  <%
  	CustomFormRS.MoveNext
  Loop
  CustomFormRS.Close
  Set CustomFormRS = Nothing
  %>
</SELECT>
<%
  if CustomFormID <> "" then
  	SQL = "Select ItemName,FieldName from FS_MF_CustomForm_Item Where FormID=" & CustomFormID
	Set CustomFormItemRS = Server.CreateObject(G_FS_RS)
	CustomFormItemRS.Open SQL,Conn,1,1
	CustomFormItemArray = CustomFormItemRS.GetRows
	CustomFormItemRS.Close
	Set CustomFormItemRS = Nothing
%>
<SELECT name="Select_CustomFormField" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000" selected="selected">ѡ����ֶ�</option>
  <option value="{CustomFormHeader}">��ͷ</option>
  <option value="{CustomFormTailor}">��β</option>
  <option value="{CustomFormValidate}">��֤��</option>
  <option value='<input name="" type="submit" value="�ύ"/>'>�ύ��ť</option>
  <option value='<input name="" type="reset" value="����"/>'>���ť</option>
  <option value='<input name="" type="button" value="��ͨ"/>'>��ͨ��ť</option>
  <%
  if IsArray(CustomFormItemArray) then
  	For i = LBound(CustomFormItemArray,2) to UBound(CustomFormItemArray,2)
  %>
  <option value="{CustomForm_<% = CustomFormItemArray(1,i) %>}"><% = CustomFormItemArray(0,i) %></option>
  <%
  	Next
  end if
  %>
</SELECT>
<SELECT name="Select_CustomFormField" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000" selected="selected">ѡ��������ʾ�ֶ�</option>
  <option value="{CustomFormData_form_usernum}">�û�ID</option>
  <option value="{CustomFormData_form_username}">�û���</option>
  <option value="{CustomFormData_form_ip}">��ԴIP��ַ</option>
  <option value="{CustomFormData_form_time}">���ʱ��</option>
  <option value="{CustomFormData_form_answer}">�ظ�����</option>
  <%
  if IsArray(CustomFormItemArray) then
  	For i = LBound(CustomFormItemArray,2) to UBound(CustomFormItemArray,2)
  %>
  <option value="{CustomFormData_<% = CustomFormItemArray(1,i) %>}"><% = CustomFormItemArray(0,i) %></option>
  <%
  	Next
  end if
  %>
</SELECT>
<%
  end if
%>
<% End Sub %>
</body>
<% 
Sub Add_Save()

End Sub
Set Conn=nothing
%>
</html><script language="JavaScript" type="text/JavaScript">
function opencat(cat)
{
  if(cat.style.display=="none"){
     cat.style.display="";
  } else {
     cat.style.display="none"; 
  }
}
</script>
<!------------2/1 by chen ��ȡ��ʽ��������� ���뿪ʼ-------------------------------->
<%
Function get_childList(TypeID,f_CompatStr)  
	Dim f_ChildNewsRs,ChildTypeListStr,f_TempStr,f_isUrlStr,lng_GetCount
	Set f_ChildNewsRs = Conn.Execute("Select id,ParentID,ClassName from FS_MF_StyleClass where ParentID=" & TypeID & " order by id desc" )
	f_TempStr =f_CompatStr & "��"
	do while Not f_ChildNewsRs.Eof
			get_childList = get_childList & "<option value="""& f_ChildNewsRs("id")&""""
			If CStr(Request.QueryString("ClassID"))=CStr(f_ChildNewsRs("id")) then
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
<!------------2/1 by chen ��ȡ��ʽ��������� �������-------------------------------->
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->

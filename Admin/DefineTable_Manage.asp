<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp"-->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_InterFace/NS_Function.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%'Copyright (c) 2006 Foosun Inc.  
Dim Conn,VClass_Rs,VClass_Sql
MF_Default_Conn
'session�ж�
MF_Session_TF
if not MF_Check_Pop_TF("MF_Define") then Err_Show
Dim CheckStr,str_Url_Add
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=15 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=10 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings>9</font>"   			'��ҳ 
toP10_=" <font face=webdings>7</font>"			'��ʮ
toP1_=" <font face=webdings>3</font>"			'��һ   
toN1_=" <font face=webdings>4</font>"			'��һ
toN10_=" <font face=webdings>8</font>"			'��ʮ
toL_="<font face=webdings>:</font>"				'βҳ
'******************************************************************
Sub DelClass()       
	Dim DelID,Str_Tmp,Str_Tmp1   
	DelID = request.Form("DelID")
	if DelID = "" then 
		response.Redirect("error.asp?ErrorUrl=&ErrCodes=<li>�����ѡ��һ����ɾ����</li>")
		response.End()
	end if
	DelID = replace(DelID," ","")
	Str_Tmp1 = DelID
	''**********************
	do while Str_Tmp1 <> "" 
		Str_Tmp1 = Get_DefineID_DefineID_TO_Del(Str_Tmp1)
		Str_Tmp = Str_Tmp & Str_Tmp1 
	loop		
	if right(Str_Tmp,1) = "," then 
		Str_Tmp	= Str_Tmp & DelID
	elseif Str_Tmp<>"" then 
		Str_Tmp	= Str_Tmp &","& DelID	
	else
		Str_Tmp = DelID	
	end if
	'response.Write(Str_Tmp)
	''**********************
	Conn.execute("delete from FS_MF_DefineTableClass where DefineID in ("&FormatIntArr(Str_Tmp)&")")
	''+++++++++++++++++
	''ɾ����ص�����������.
	Conn.execute("delete from FS_MF_DefineTable where ClassID in ("&FormatIntArr(Str_Tmp)&")")
	''+++++++++++++++++
	response.Redirect("Success.asp?ErrorUrl=DefineTable_Manage.asp&ErrCodes=<li>��ϲ��ɾ���ɹ���</li>")
End Sub

Sub SaveClass()
	call Say_Limit_Class("SaveMode") 
	Dim Edit_PartID,ii,Str_Req_Child,Str_Req_Tmp,Str_Biangeng_ParentID_ParentID,Err_info
	VClass_Sql = "select DefineID,DefineName,ParentID from FS_MF_DefineTableClass"
	if NoSqlHack(request.Form("ParID"))<>"" then 
	''����
		if CheckCardCF(NoSqlHack(Trim(request.Form("DefineName"))),request.Form("ParID"))<>"" then 
			response.Redirect("error.asp?ErrCodes=<li>ͬ��������: "&request.Form("DefineName")&" �Ѿ����ڡ�</li>")
			response.End()
		end if	
		VClass_Sql = VClass_Sql & " where DefineID=0"
	elseif 	NoSqlHack(request.Form("DefineID"))<>"" then 
	''�޸�
		VClass_Sql = VClass_Sql & " where DefineID=" & NoSqlHack(request.Form("DefineID"))
		''**********************
		''���Զ����ֶ����Ҫ���ʱ�жϵ�ǰ����������Ӽ��ټ���Ҫ��������һ���Ƿ���4�����ڡ��������ⷵ�ء�
		for ii = 1 to 1 step -1
			if request.Form("vclass"&ii)<>"" then Edit_PartID = request.Form("vclass"&ii) : exit for
		next
		''''''''''''''''''''''''''
		if Edit_PartID = "[ChangeToTop]" then Edit_PartID = 0
		if Edit_PartID > 0 then
		''''''''''''''''
		''Str_Biangeng_ParentID_ParentID׼��������ĸ��������䱾��ĸ������
		Str_Req_Tmp = Edit_PartID		
		do while Str_Req_Tmp<>""
			Str_Req_Tmp = Get_ParID_ParID_TO_Save(Str_Req_Tmp)
			Str_Biangeng_ParentID_ParentID = Str_Biangeng_ParentID_ParentID & Str_Req_Tmp 
		loop		
		if right(Str_Biangeng_ParentID_ParentID,1) = "," then 
			Str_Biangeng_ParentID_ParentID	= Str_Biangeng_ParentID_ParentID & Edit_PartID
		elseif Str_Biangeng_ParentID_ParentID<>"" then 
			Str_Biangeng_ParentID_ParentID	= Str_Biangeng_ParentID_ParentID &","& Edit_PartID	
		else
			Str_Biangeng_ParentID_ParentID = Edit_PartID	
		end if
		Err_info = Err_info & "<li>��������ĸ����������и���:"&Str_Biangeng_ParentID_ParentID&"</li>"
		response.Write(Err_info)
		'response.End()
		''''''''''''''''''''''''''
		''Str_Req_Child��ǰ������������		
		Str_Req_Tmp = request.Form("DefineID")		
		do while Str_Req_Tmp <> "" 
			Str_Req_Tmp = Get_DefineID_DefineID_TO_Del(Str_Req_Tmp)
			Str_Req_Child = Str_Req_Child & Str_Req_Tmp 
		loop		
		if right(Str_Req_Child,1) = "," then 
			Str_Req_Child	= Str_Req_Child & request.Form("DefineID")
		elseif Str_Req_Child<>"" then 
			Str_Req_Child	= Str_Req_Child &","& request.Form("DefineID")	
		else
			Str_Req_Child = request.Form("DefineID")	
		end if
		Err_info = Err_info & "<li>��ǰ��������������:"&Str_Req_Child&"</li>"
		Err_info = Err_info & "<li>Ԥ�ϵ��ܼ���:"&ubound(split(Str_Biangeng_ParentID_ParentID & "," & Str_Req_Child,",")) + 1&"</li>"
		response.Write(Err_info)
		'response.End()
		if ubound(split(Str_Biangeng_ParentID_ParentID & "," & Str_Req_Child,",")) + 1 > 4 then
			Err_info = Err_info & "<li>��Ǹ,��Ҫ���Ѿ������ļ�!���в��ܸ��ĵ�������.</li>"
			response.Redirect("error.asp?ErrCodes="&Err_info&"")
			response.End()
		end if
		'''''''''''''''
		end if
		''**********************
	else
		response.Redirect("error.asp?ErrCodes=<li>��Ҫ���Զ����ֶ�IDû���ṩ��</li>")	
		response.End()
	end if
	'response.Write(VClass_Sql)
	'response.End()
	Set VClass_Rs	= CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,Conn,3,3
	if NoSqlHack(request.Form("ParID"))<>"" then 
		VClass_Rs.AddNew
		VClass_Rs("ParentID") = NoSqlHack(request.Form("ParID"))
		VClass_Rs("DefineName") = NoSqlHack(request.Form("DefineName"))
		VClass_Rs.update
		VClass_Rs.close
		response.Redirect("Success.asp?ErrorUrl="&server.URLEncode( "DefineTable_Manage.asp?Act=Add&DefineID="&NoSqlHack(request.Form("ParID"))&"&VCText="&NoSqlHack(request.Form("VCText")) )&"&ErrCodes=<li>����ɹ���</li>")
	end if
	'''�޸�
	if NoSqlHack(request.Form("DefineID"))<>"" then 
		'response.Write("<br>VClass_Sql:"&VClass_Sql&"<br>Edit_PartID:"&Edit_PartID)
		'response.End()
		VClass_Rs("DefineName") = NoSqlHack(request.Form("DefineName"))
		if Edit_PartID<>"" then VClass_Rs("ParentID") = Edit_PartID
		VClass_Rs.update

		if Edit_PartID<>"" then 
			Dim PartID_PartID_Rs
			''ȡ�ñ����ĸ���ID�ĸ���ID�����������ص�DefineID���Ա���ʾ����IDͬ������������
			set PartID_PartID_Rs = Conn.execute("select ParentID from FS_MF_DefineTableClass where DefineID="&Edit_PartID)
			if not PartID_PartID_Rs.eof then  Edit_PartID = PartID_PartID_Rs(0)
			PartID_PartID_Rs.close
			set PartID_PartID_Rs = nothing
		else
			Edit_PartID = VClass_Rs("ParentID")
		end if		
		VClass_Rs.close
		response.Redirect("Success.asp?ErrorUrl="&server.URLEncode( "DefineTable_Manage.asp?Act=View&DefineID="&Edit_PartID )&"&ErrCodes=<li>����ɹ���</li>")
	end if
End Sub

Function CheckCardCF(DefineName,ParentID)
''���¼����Ƿ��ظ�,�ظ��򷵻�DefineName,���ظ��򷵻�""
	Dim CheckCardCF_Rs
	set CheckCardCF_Rs = Conn.execute( "select Count(*) from FS_MF_DefineTableClass where DefineName='"&NoSqlHack(DefineName)&"' and ParentID="&CintStr(ParentID) )
	if  CheckCardCF_Rs(0)>0 then 
		CheckCardCF = DefineName
	else 
		CheckCardCF = ""
	end if
	CheckCardCF_Rs.close	
End Function

Sub Say_Limit_Class(SownMode)
Dim Arr_Tmp,str_Session
str_Session = session("TopMenu_DaoHang_DefineID_List")
select case SownMode
case "AddMode"
	''============================
	''�Զ����ֶ��������ļ����ж�.
	if str_Session <> "" then 
		if left(str_Session,1)="," then str_Session = mid(str_Session,2,len(str_Session))
		if right(str_Session,1)="," then str_Session = mid(str_Session,1,len(str_Session) - 1)
		Arr_Tmp = split(str_Session,",")
		if ubound(Arr_Tmp)>0 then 
		''��Ť
			response.Write(" disabled title=""�Զ����ֶη������һ��!��ǰ�ȼ�:��"&cstr(ubound(Arr_Tmp) + 1)&"��."" ")	
		end if
	end if
	''============================
case "SaveMode"
	''============================
	''�Զ����ֶ��������ļ����ж�.
	if str_Session <> "" then 
		if left(str_Session,1)="," then str_Session = mid(str_Session,2,len(str_Session))
		if right(str_Session,1)="," then str_Session = mid(str_Session,1,len(str_Session) - 1)
		Arr_Tmp = split(str_Session,",")
		if ubound(Arr_Tmp)>0 then 
			response.Redirect("error.asp?ErrCodes=<li>�Զ����ֶ�����ܳ���һ����</li>")	
			response.End()
		else
			response.Write("�Զ����ֶη������һ��!��ǰ�ȼ�:��"&cstr(ubound(Arr_Tmp) + 1)&"��.")	
		end if
	end if
	''============================
case else

end select 	
End Sub

Function Get_VClass(DefineID)
''�ݹ������ʾ����
	Dim Get_Html
	VClass_Sql = "select DefineID,DefineName,ParentID from FS_MF_DefineTableClass"
	if DefineID<>"" and DefineID>0 then 
		VClass_Sql = VClass_Sql &" where ParentID = "&DefineID&" order by DefineID desc"
	else
		VClass_Sql = VClass_Sql &" where ParentID = 0 order by DefineID desc"	
	end if
	Set VClass_Rs	= CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,Conn,1,1	
	IF not VClass_Rs.eof THEN

	VClass_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>VClass_Rs.PageCount Then cPageNo=VClass_Rs.PageCount 
	VClass_Rs.AbsolutePage=cPageNo
	
	  FOR int_Start=1 TO int_RPP 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& VClass_Rs("DefineName") & "</td>" & vbcrlf		
		Get_Html = Get_Html & "<td align=""center""><a href=""DefineTable_Info_Manage.asp?Act=View&Add_Sql="& server.URLEncode(Encrypt( "ClassID="&VClass_Rs("DefineID") )) &"""  class=""otherset"">�鿴�����ֶ�</FONT></a></td>" & vbcrlf
''		Get_Html = Get_Html & "<td align=""center""><a href=DefineTable_Manage.asp?Act=View&DefineID="&VClass_Rs("DefineID")&">"& VClass_Rs("DefineName") & "</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""DefineTable_Info_Manage.asp?Act=Add&ClassID="&VClass_Rs("DefineID")&"""  class=""otherset"">���������ֶ�</FONT></a></td>" & vbcrlf
''		Get_Html = Get_Html & "<td align=""center""><a href=""DefineTable_Manage.asp?Act=Add&DefineID="&VClass_Rs("DefineID")&"&VCText="&VClass_Rs("DefineName")&"""  class=""otherset"">����</FONT></a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""DefineTable_Manage.asp?Act=Edit&DefineID="&VClass_Rs("DefineID")&"&VCText="&VClass_Rs("DefineName")&"""  class=""otherset"">����</FONT></a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" "&CheckStr&" name=""DelID"" id=""DelID"" value="""&VClass_Rs("DefineID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		CheckStr = ""	
		VClass_Rs.MoveNext
 		if VClass_Rs.eof or VClass_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(VClass_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""button"" name=""button1112"" value="" ����ͬ����Ŀ "" onClick=""location='DefineTable_Manage.asp?Act=Add"& str_Url_Add &"'"">" & vbcrlf
	Get_Html = Get_Html & "<input type=""submit"" name=""submit"" value="" ɾ�� "" onclick=""javascript:return confirm('ȷ��Ҫɾ����ѡ��Ŀ��?');""></td>"& vbcrlf
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	VClass_Rs.close
	Get_VClass = Get_Html
End Function

Function Get_DefineID_DefineID_TO_Del(Req_DefineID)
''ѭ�����ã�ͨ�������DefineID�õ��Ӽ���DefineID���������������һ�𴫸�DelClass������ɾ������������
Dim Str_Tmp,This_Fun_Sql
if Req_DefineID="" or isnull(Req_DefineID) or Req_DefineID="," then Get_DefineID_DefineID_TO_Del="" : exit Function
This_Fun_Sql="select DefineID from FS_MF_DefineTableClass where ParentID in ("&Req_DefineID&")"
set VClass_Rs=Conn.execute(This_Fun_Sql)
do while not VClass_Rs.eof 
	Str_Tmp = Str_Tmp & VClass_Rs(0) & ","	
	VClass_Rs.movenext
loop	
VClass_Rs.close
Get_DefineID_DefineID_TO_Del = Str_Tmp
End Function

Function Get_ParID_ParID_TO_Save(Req_ParID)
''ѭ�����ã�ͨ�������ParentID�õ����ϵ�����ParentID���������������һ�𴫸�SaveClass�������ж��Ƿ񳬹�4��.
Dim Str_Tmp,This_Fun_Sql
if Req_ParID="" or isnull(Req_ParID) or Req_ParID="," then Get_ParID_ParID_TO_Save="" : exit Function
This_Fun_Sql="select ParentID from FS_MF_DefineTableClass where DefineID in ("&Req_ParID&")"
set VClass_Rs=Conn.execute(This_Fun_Sql)
do while not VClass_Rs.eof 
	if VClass_Rs(0)=0 then exit do : Get_ParID_ParID_TO_Save="" : exit Function 
	Str_Tmp = Str_Tmp & VClass_Rs(0) & ","	
	VClass_Rs.movenext
loop	
VClass_Rs.close
Get_ParID_ParID_TO_Save = Str_Tmp
End Function

Function Get_PatID_TO_View(View_ID)
''�������ã��õ�����ID�Ա㷵�ظ�����鿴
	Dim This_Fun_Sql
	if View_ID="" then Get_PatID_TO_View=0 : exit Function
	This_Fun_Sql = "select ParentID from FS_MF_DefineTableClass where DefineID="&CintStr(View_ID)
	Set VClass_Rs = Conn.execute(This_Fun_Sql)
	if not VClass_Rs.eof then 
		Get_PatID_TO_View = VClass_Rs(0)
	else
		Get_PatID_TO_View = 0	
	end if
	VClass_Rs.close
End Function

Function Get_PatTxt_TO_View(View_ID)
''�������ã��õ�������Ŀ�����Ա㷵�ظ�����鿴���������Ӧ
Dim VClass_Rs1,This_Fun_Sql
	if View_ID="" then Get_PatTxt_TO_View="��" : exit Function
	This_Fun_Sql = "select ParentID from FS_MF_DefineTableClass where DefineID='"&NoSqlHack(View_ID)&"'"
	Set VClass_Rs = Conn.execute(This_Fun_Sql)
	if not VClass_Rs.eof then 
		set VClass_Rs1 = Conn.execute( "select DefineName from FS_MF_DefineTableClass where DefineID="&VClass_Rs(0) )
		if not VClass_Rs1.eof then 
			Get_PatTxt_TO_View = VClass_Rs1(0)
		else
			Get_PatTxt_TO_View = "��"	
		end if
		VClass_Rs1.close
		set VClass_Rs1=nothing		
	else
		Get_PatTxt_TO_View = "��"	
	end if
	VClass_Rs.close
End Function

Function TopMenu_DaoHang()
''��㵼������ session("TopMenu_DaoHang_DefineID_List") = ',1,2,3,4,'
Dim str_Req_DefineID,str_Session,Sql_Session,Per_ID_Session,Str_Tmp,Str_Remove,Arr_Tmp,This_Fun_Sql
	if request.QueryString("DefineID")="" or request.QueryString("DefineID")="0" then 
		session("TopMenu_DaoHang_DefineID_List")="" : TopMenu_DaoHang = "" : exit Function
	end if
	''**************
	str_Req_DefineID = NoSqlHack(request.QueryString("DefineID"))
	str_Req_DefineID = ","&str_Req_DefineID&","
	'if session("TopMenu_DaoHang_DefineID_List") = "" then session("TopMenu_DaoHang_DefineID_List") = str_Req_DefineID
	str_Session = session("TopMenu_DaoHang_DefineID_List")
	if str_Session = "" or isEmpty(str_Session) then str_Session = str_Req_DefineID 
	if left(str_Session,1)<>"," then str_Session = ","&str_Session
	if right(str_Session,1)<>"," then str_Session = str_Session&","
	session("TopMenu_DaoHang_DefineID_List") = str_Session
	'response.Write("ԭ"&session("TopMenu_DaoHang_DefineID_List"))
	',1,2,3,4,' �� ',3,'
	if right(str_Session,len(str_Req_DefineID)) <> str_Req_DefineID then 
	if instr(str_Session,str_Req_DefineID)>0 then
	''��ȡ
		str_Session = mid(str_Session,1,instr(str_Session,str_Req_DefineID) + len(str_Req_DefineID) - 1)
		session("TopMenu_DaoHang_DefineID_List") = str_Session
		'response.Write("��"&session("TopMenu_DaoHang_DefineID_List"))
	else
	''���
		str_Session = str_Session & NoSqlHack(request.QueryString("DefineID")) &","
		session("TopMenu_DaoHang_DefineID_List") = str_Session
		'response.Write("��"&session("TopMenu_DaoHang_DefineID_List"))
	end if
	end if
	''**************
	Sql_Session = str_Session
	if left(Sql_Session,1)="," then Sql_Session = mid(Sql_Session,2,len(Sql_Session))
	if right(Sql_Session,1)="," then Sql_Session = mid(Sql_Session,1,len(Sql_Session) - 1)
	'response.Write(session("TopMenu_DaoHang_DefineID_List"))
	''**************
	''˳����ʵ�����
	str_Session = ""
	Arr_Tmp = split(Sql_Session,",")	
	for each Per_ID_Session in Arr_Tmp
		This_Fun_Sql = "select DefineID,DefineName from FS_MF_DefineTableClass where DefineID = '"&Per_ID_Session&"'"
		Set VClass_Rs = Conn.execute(This_Fun_Sql)
		if not VClass_Rs.eof then 
			Str_Tmp = Str_Tmp & "<a href=""DefineTable_Manage.asp?Act=View&DefineID=" &VClass_Rs(0)&""" class=""sd""><b>"&VClass_Rs(1)&"</b><a> >> "
			str_Session = str_Session &VClass_Rs(0)& "," 
		end if
		VClass_Rs.close
	next
	session("TopMenu_DaoHang_DefineID_List") = "," & str_Session	
	if right(Str_Tmp,len(" >> "))=" >> " then Str_Tmp = mid(Str_Tmp,1,len(Str_Tmp) - len(" >> "))
	TopMenu_DaoHang = Str_Tmp
	'response.Write(session("TopMenu_DaoHang_DefineID_List"))
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<%if request.QueryString("Act")="Edit" then%><script language="javascript" src="../FS_Inc/class_liandong.js" type="text/javascript"></script><%end if%>
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return true;">
<%
'******************************************************************
select case request.QueryString("Act")
	case "View",""
		ViewClass
	case "Add"
		AddClass
	case "Del"
		DelClass
	case "Edit"
		EditClass
	case "Save"	
		SaveClass
end select
'******************************************************************

Sub ViewClass()
Dim View_DefineID,IsOk,VClass_Rs1
IsOk = false
View_DefineID = NoSqlHack(request.QueryString("DefineID"))
if View_DefineID<>"" then if isnumeric(View_DefineID) then if View_DefineID>0 then IsOk = true
if IsOk=false then 
	str_Url_Add = "&DefineID=0&VCText=��"	
else	
	set VClass_Rs1 = Conn.execute( "select DefineName from FS_MF_DefineTableClass where DefineID="&View_DefineID )
	if not VClass_Rs1.eof then 
		str_Url_Add = "&DefineID="&View_DefineID&"&VCText="	& VClass_Rs1(0)
	else
		str_Url_Add = "&DefineID=0&VCText=��"	
	end if
	VClass_Rs1.close
	set VClass_Rs1=nothing		
end if		
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu">
    <td class="xingmu"><a href="#" onMouseOver="this.T_BGCOLOR='#404040';this.T_FONTCOLOR='#FFFFFF';return escape('<div align=\'center\'>FoosunCMS5.0<br>  <BR>Copyright (c) 2006 Foosun Inc</div>')" class="sd"><strong>�Զ������</strong></a></td>
  </tr>
  <tr class="hback">
    <td class="hback"><!--<a href="DefineTable_Info_Manage.asp?Act=View">������ҳ</a> -->
      <!-- | <a href="DefineTable_Manage.asp?Act=View&DefineID=<%=Get_PatID_TO_View(request.QueryString("DefineID"))%>" class="sd"><b>���ظ���</b></a>
	   | <a href="#" onClick="javascript:history.back();" class="sd"><b>����</b></a>-->
     <!-- | --><a href="DefineTable_Info_Manage.asp">�ֶ����ݹ���</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="VCForm" id="VCForm" method="post" action="?Act=Del">
    <tr  class="hback"> 
      <td align="left" class="xingmu" colspan="5"> 
        <!-- | ���ർ����<a href="DefineTable_Manage.asp" class="sd"><b>��ҳ</b></a>  >> <%=TopMenu_DaoHang()%>-->
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="center"  class="xingmu">��������</td>
      <td width="25%" align="center" class="xingmu">�鿴�Զ����ֶ�</td>
      <td width="25%" align="center" class="xingmu">�����Զ����ֶ�</td>
      <!--<td width="10%" align="center" class="xingmu">��������Ŀ</td>-->
      <td width="10%" align="center" class="xingmu">����</td>
      <td width="5%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_VClass(request.QueryString("DefineID")) )
	%>
  </form>
</table>
<%End Sub
Sub AddClass()
dim View_Par_ID
if request.QueryString("DefineID")<>"" then 
	View_Par_ID = request.QueryString("DefineID")
else
	View_Par_ID = request.QueryString("ParID")
end if		
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu"> 
    <td class="xingmu">�Զ������</td>
  </tr>
  <tr class="hback"> 
    <td class="hback"><a href="DefineTable_Manage.asp?Act=View">������ҳ</a> 
      <!-- | <a href="DefineTable_Manage.asp?Act=View&DefineID=<%=Get_PatID_TO_View(request.QueryString("DefineID"))%>" class="sd"><b>���ظ���</b></a>
	   | <a href="#" onClick="javascript:history.back();" class="sd"><b>����</b></a>-->
      | <a href="DefineTable_Info_Manage.asp">�ֶ����ݹ���</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="VCForm" id="VCForm" method="post" action="?Act=Save">
    <tr  class="hback"> 
      <td colspan="2" align="left" class="xingmu" >�����Զ����ֶη�����Ϣ �Զ����ֶη���Ŀǰһ��</td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��һ���Զ����ֶ�����</td>
      <td width="596"> <input name="VCText" id="VCText" readonly type="text" size="50" value="<%=request.QueryString("VCText")%>"/> 
        <input type="hidden" name="ParID" id="ParID" value="<%=View_Par_ID%>"/> 
      </td>
    </tr>
    <tr class="hback"> 
      <td align="right">�µ��Զ����ֶ����ƣ�</td>
      <td width="596"><input name="DefineName" type="text" id="DefineName" size="50" onBlur="if(this.value==''){Chk_vCName.innerText='���������д';this.focus();} else if(Chk_vCName.innerText!='')Chk_vCName.innerText=''" />
        <span id="Chk_vCName" class="tx"></span> </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4"> <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit" <%call Say_Limit_Class("AddMode")%> name="SaveClass" id="SaveClass" value=" ���� " onClick="if(DefineName.value==''){alert('��Ҫ�Ĳ���������д');DefineName.focus();return false}" /> 
              &nbsp; <input type="reset" name="ReSet" id="ReSet" value=" ���� " /> 
            </td>
            <td width="50" align="right"> <a href="DefineTable_Manage.asp?Act=View&DefineID=<%=Get_PatID_TO_View(View_Par_ID)%>">���ظ���</a> 
            </td>
            <td width="50" align="right"> <a href="#" onClick="javascript:history.back();">����</a> 
            </td>
          </tr>
        </table></td>
    </tr>
  </form>
</table>
<%End Sub
Sub EditClass()
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu"> 
    <td class="xingmu">�Զ������</td>
  </tr>
  <tr class="hback"> 
    <td class="hback"><a href="DefineTable_Manage.asp?Act=View">������ҳ</a> 
      <!-- | <a href="DefineTable_Manage.asp?Act=View&DefineID=<%=Get_PatID_TO_View(request.QueryString("DefineID"))%>" class="sd"><b>���ظ���</b></a>
	   | <a href="#" onClick="javascript:history.back();" class="sd"><b>����</b></a>-->
      | <a href="DefineTable_Info_Manage.asp">�ֶ����ݹ���</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="VCForm" id="VCForm" method="post" action="?Act=Save">
    <tr  class="hback"> 
      <td colspan="2" align="left" class="xingmu" >�޸��Զ����ֶη�����Ϣ �Զ����ֶη���Ŀǰһ��</td>
    </tr>
    <tr  class="hback" style="display:none"> 
      <td align="right">����ϼ����<br>
        <span class="tx">�������������</span></td>
      <td width="596"> <SELECT NAME="vclass1" ID="vclass1" style="width:100px">
          <OPTION></OPTION>
        </SELECT> 
        <!--		  
<!---�����˵���ʼ--- >	
	<SELECT NAME="vclass1" ID="vclass1" onBlur="javascript:RemoveChildopt(this,'vclass2,vclass3,vclass4');"  style="width:100px">
         <OPTION></OPTION>
    </SELECT>
	<SELECT NAME="vclass2" ID="vclass2" onBlur="javascript:RemoveChildopt(this,'vclass3,vclass4');" style="width:100px">
         <OPTION></OPTION>
    </SELECT>
	<SELECT NAME="vclass3" ID="vclass3" onBlur="javascript:RemoveChildopt(this,'vclass4');" style="width:100px">
        <OPTION></OPTION>
    </SELECT>
	<SELECT NAME="vclass4" ID="vclass4" style="width:100px">
    	<OPTION></OPTION>
    </SELECT>
<!---�����˵�����--- > -->
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�Զ����ֶ�����</td>
      <td width="596"> <input name="DefineName" id="DefineName" type="text" size="50" value="<%=request.QueryString("VCText")%>" onBlur="if(this.value==''){Chk_vCName.innerText='���������д';this.focus()} else if(Chk_vCName.innerText!='')Chk_vCName.innerText=''" />
        <span id="Chk_vCName" class="tx"></span> <input type="hidden" name="DefineID" id="DefineID" value="<%=request.QueryString("DefineID")%>"> 
        <!--<input type="text" name="ParentID" id="ParentID" value="">-->
      </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4"> <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit" <%call Say_Limit_Class("AddMode")%> name="SaveClass2" id="SaveClass2" value=" ���� " onClick="if(DefineName.value==''){DefineName.focus();return false}" /> 
              &nbsp; <input type="reset" name="ReSet2" id="ReSet2" value=" ���� " /> 
            </td>
            <td width="50" align="right"> <a href="DefineTable_Manage.asp?Act=View&DefineID=<%=Get_PatID_TO_View(request.QueryString("DefineID"))%>">���ظ���</a> 
            </td>
            <td width="50" align="right"> <a href="#" onClick="javascript:history.back();">����</a> 
            </td>
          </tr>
        </table></td>
    </tr>
  </form>
</table>
<%End Sub%>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
</body>
<%if request.QueryString("Act")="Edit" then%>
<script language="javascript">
<!-- 
//awen created
//�����˵�---�Զ����ֶ����   ���4��  Ŀǰ1�� --start 
//���ݸ�ʽ ID������ID������
var array=new Array();
<%dim NS_JS_Sql,NS_JS_RS,NS_JS_i
  NS_JS_Sql="select DefineID,ParentID,DefineName from FS_MF_DefineTableClass where DefineID<>"&request.QueryString("DefineID")
  set NS_JS_RS=Conn.execute(NS_JS_Sql)
  NS_JS_i=0
  do while not NS_JS_RS.eof
%>
array[<%=NS_JS_i%>]=new Array("<%=NS_JS_RS("DefineID")%>","<%=NS_JS_RS("ParentID")%>","<%=NS_JS_RS("DefineName")%>"); 
<%
	NS_JS_RS.movenext
	NS_JS_i=NS_JS_i+1
loop
NS_JS_RS.close
%>

var liandong=new CLASS_LIANDONG_YAO(array)
liandong.firstSelectChange("0","vclass1");
/*
liandong.subSelectChange("vclass1","vclass2");
liandong.subSelectChange("vclass2","vclass3");
liandong.subSelectChange("vclass3","vclass4");

//---------------------------������������������
function RemoveChildopt(obj,StrList)
{
	var TmpArr = StrList.split(',');
	if(obj.selectedIndex<2)
	{		
		for (var i=TmpArr.length-1 ; i>=0; i--)
		{
			//alert(TmpArr[i]);
			if (TmpArr[i]!='') 
				//�����������
				for (var j=document.getElementById(TmpArr[i]).options.length-1 ; j>=0 ; j--)
				document.getElementById(TmpArr[i]).options.remove(j);				
		}	
	}
}    */
//end 
-->
</script>
<%end if%>
<%
Set VClass_Rs=nothing
Conn.close
Set Conn=nothing
%>

</html>







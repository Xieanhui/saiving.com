<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<% 
Dim Conn,User_Conn,VClass_Rs,VClass_Sql
Dim str_Url_Add,CheckStr
MF_Default_Conn
MF_User_Conn
MF_Session_TF

if not MF_Check_Pop_TF("ME_HY") then Err_Show 

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
	if not MF_Check_Pop_TF("ME026") then Err_Show 
	Dim DelID,Str_Tmp,Str_Tmp1
	DelID = request.Form("DelID")
	if DelID = "" then 
		response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>�����ѡ��һ����ɾ����</li>")
		response.End()
	end if
	DelID = replace(DelID," ","")
	Str_Tmp1 = DelID
	''**********************
	do while Str_Tmp1 <> "" 
		Str_Tmp1 = Get_VCID_VCID_TO_Del(Str_Tmp1)
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
	User_Conn.execute("delete from FS_ME_VocationClass where VCID in ("&FormatIntArr(Str_Tmp)&")")
	response.Redirect("../Success.asp?ErrorUrl=User/VocationClass.asp&ErrCodes=<li>��ϲ��ɾ���ɹ���</li>")
End Sub

Sub SaveClass()
	if not MF_Check_Pop_TF("ME026") then Err_Show 
	call Say_Limit_Class("SaveMode") 
	Dim Edit_PartID,ii,Str_Req_Child,Str_Req_Tmp,Str_Biangeng_ParentID_ParentID,Err_info
	VClass_Sql = "select VCID,vClassName,vClassName_En,ParentID from FS_ME_VocationClass"
	if NoSqlHack(request.Form("ParID"))<>"" then 
	''����
		if CheckCardCF(NoSqlHack(request.Form("vClassName")),request.Form("ParID"))<>"" then 
			response.Redirect("../Error.asp?ErrCodes=<li>ͬ��������: "&request.Form("vClassName")&" �Ѿ����ڡ�</li>")
			response.End()
		end if	
		VClass_Sql = VClass_Sql & " where VCID=0"
	elseif 	NoSqlHack(request.Form("VCID"))<>"" then 
	''�޸�
		VClass_Sql = VClass_Sql & " where VCID=" & NoSqlHack(request.Form("VCID"))
		''**********************
		''����ҵ���Ҫ���ʱ�жϵ�ǰ����������Ӽ��ټ���Ҫ��������һ���Ƿ���4�����ڡ��������򷵻ء�
		for ii = 4 to 1 step -1
			if request.Form("vclass"&ii)<>"" then Edit_PartID = NoSqlHack(request.Form("vclass"&ii)) : exit for
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
		Str_Req_Tmp = request.Form("VCID")		
		do while Str_Req_Tmp <> "" 
			Str_Req_Tmp = Get_VCID_VCID_TO_Del(Str_Req_Tmp)
			Str_Req_Child = Str_Req_Child & Str_Req_Tmp 
		loop		
		if right(Str_Req_Child,1) = "," then 
			Str_Req_Child	= Str_Req_Child & NoSqlHack(request.Form("VCID"))
		elseif Str_Req_Child<>"" then 
			Str_Req_Child	= Str_Req_Child &","& NoSqlHack(request.Form("VCID"))	
		else
			Str_Req_Child = NoSqlHack(request.Form("VCID"))	
		end if
		Err_info = Err_info & "<li>��ǰ��������������:"&Str_Req_Child&"</li>"
		Err_info = Err_info & "<li>Ԥ�ϵ��ܼ���:"&ubound(split(Str_Biangeng_ParentID_ParentID & "," & Str_Req_Child,",")) + 1&"</li>"
		response.Write(Err_info)
		'response.End()
		if ubound(split(Str_Biangeng_ParentID_ParentID & "," & Str_Req_Child,",")) + 1 > 4 then
			Err_info = Err_info & "<li>��Ǹ,��Ҫ���Ѿ������ļ�!���в��ܸ��ĵ�������.</li>"
			response.Redirect("../error.asp?ErrCodes="&Err_info&"")
			response.End()
		end if
		'''''''''''''''
		end if
		''**********************
	else
		response.Redirect("../error.asp?ErrCodes=<li>��Ҫ����ҵIDû���ṩ��</li>")	
		response.End()
	end if
	'response.Write(VClass_Sql)
	'response.End()
	Set VClass_Rs	= CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,3,3
	if NoSqlHack(request.Form("ParID"))<>"" then 
		VClass_Rs.AddNew
		VClass_Rs("ParentID") = NoSqlHack(request.Form("ParID"))
		VClass_Rs("vClassName") = NoSqlHack(request.Form("vClassName"))
		VClass_Rs("vClassName_En") = NoSqlHack(request.Form("vClassName_En"))
		VClass_Rs.update
		VClass_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/VocationClass.asp?Act=Add&VCID="&NoSqlHack(request.Form("ParID"))&"&VCText="&NoSqlHack(request.Form("VCText")) )&"&ErrCodes=<li>����ɹ���</li>")
	end if
	'''�޸�
	if NoSqlHack(request.Form("VCID"))<>"" then 
		'response.Write("<br>VClass_Sql:"&VClass_Sql&"<br>Edit_PartID:"&Edit_PartID)
		'response.End()
		VClass_Rs("vClassName") = NoSqlHack(request.Form("vClassName"))
		VClass_Rs("vClassName_En") = NoSqlHack(request.Form("vClassName_En"))
		if Edit_PartID<>"" then VClass_Rs("ParentID") = Edit_PartID
		VClass_Rs.update

		if Edit_PartID<>"" then 
			Dim PartID_PartID_Rs
			''ȡ�ñ����ĸ���ID�ĸ���ID�����������ص�VCID���Ա���ʾ����IDͬ������������
			set PartID_PartID_Rs = User_Conn.execute("select ParentID from FS_ME_VocationClass where VCID="&NoSqlHack(Edit_PartID))
			if not PartID_PartID_Rs.eof then  Edit_PartID = PartID_PartID_Rs(0)
			PartID_PartID_Rs.close
			set PartID_PartID_Rs = nothing
		else
			Edit_PartID = VClass_Rs("ParentID")
		end if		
		VClass_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/VocationClass.asp?Act=View&VCID="&Edit_PartID )&"&ErrCodes=<li>����ɹ���</li>")
	end if
End Sub

Function CheckCardCF(vClassName,ParentID)
''���¼����Ƿ��ظ�,�ظ��򷵻�vClassName,���ظ��򷵻�""
	Dim CheckCardCF_Rs
	set CheckCardCF_Rs = User_Conn.execute( "select Count(*) from FS_ME_VocationClass where vClassName='"&NoSqlHack(vClassName)&"' and ParentID="&NoSqlHack(ParentID))
	if  CheckCardCF_Rs(0)>0 then 
		CheckCardCF = vClassName
	else 
		CheckCardCF = ""
	end if
	CheckCardCF_Rs.close	
End Function

Sub Say_Limit_Class(SownMode)
Dim Arr_Tmp,str_Session
str_Session = session("TopMenu_DaoHang_VCID_List")
select case SownMode
case "AddMode"
	''============================
	''��ҵ�������ļ����ж�.
	if str_Session <> "" then 
		if left(str_Session,1)="," then str_Session = mid(str_Session,2,len(str_Session))
		if right(str_Session,1)="," then str_Session = mid(str_Session,1,len(str_Session) - 1)
		Arr_Tmp = split(str_Session,",")
		if ubound(Arr_Tmp)>2 then 
		''��Ť
			response.Write(" disabled title=""��ҵ��������ļ�!��ǰ�ȼ�:��"&cstr(ubound(Arr_Tmp) + 1)&"��."" ")	
		end if
	end if
	''============================
case "SaveMode"
	''============================
	''��ҵ�������ļ����ж�.
	if str_Session <> "" then 
		if left(str_Session,1)="," then str_Session = mid(str_Session,2,len(str_Session))
		if right(str_Session,1)="," then str_Session = mid(str_Session,1,len(str_Session) - 1)
		Arr_Tmp = split(str_Session,",")
		if ubound(Arr_Tmp)>2 then 
			response.Redirect("../error.asp?ErrCodes=<li>��ҵ����ܳ����ļ���</li>")	
			response.End()
		else
			response.Write("��ҵ��������ļ�!��ǰ�ȼ�:��"&cstr(ubound(Arr_Tmp) + 1)&"��.")
		end if
	end if
	''============================
case else

end select 	
End Sub

Function Get_VClass(vcid)
''�ݹ������ʾ����
	Dim Get_Html
	VClass_Sql = "select VCID,vClassName,vClassName_En,ParentID from FS_ME_VocationClass"
	if vcid<>"" and vcid>0 then 
		VClass_Sql = VClass_Sql &" where ParentID = "&vcid
	else
		VClass_Sql = VClass_Sql &" where ParentID = 0"	
	end if
	Set VClass_Rs	= CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,1,1	
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
		Get_Html = Get_Html & "<td align=""center""><a href=VocationClass.asp?Act=View&VCID="&VClass_Rs("VCID")&">"& VClass_Rs("vClassName") & "</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& VClass_Rs("vClassName_En") & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""VocationClass.asp?Act=Add&VCID="&VClass_Rs("VCID")&"&VCText="&VClass_Rs("vClassName")&"""  class=""otherset"">����</FONT></a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""VocationClass.asp?Act=Edit&VCID="&VClass_Rs("VCID")&"&VCText="&VClass_Rs("vClassName")&"""  class=""otherset"">����</FONT></a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" "&CheckStr&" name=""DelID"" id=""DelID"" value="""&VClass_Rs("VCID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		CheckStr = ""	
		VClass_Rs.MoveNext
 		if VClass_Rs.eof or VClass_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(VClass_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""button"" name=""button1112"" value="" ����ͬ����Ŀ "" onClick=""location='VocationClass.asp?Act=Add"& str_Url_Add &"'"">" & vbcrlf
	Get_Html = Get_Html & "<input type=""submit"" name=""submit"" value="" ɾ�� "" onclick=""javascript:return confirm('ȷ��Ҫɾ����ѡ��Ŀ��?');""></td>"& vbcrlf
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	VClass_Rs.close
	Get_VClass = Get_Html
End Function

Function Get_VCID_VCID_TO_Del(Req_VCID)
''ѭ�����ã�ͨ�������VCID�õ��Ӽ���VCID���������������һ�𴫸�DelClass������ɾ������������
Dim Str_Tmp,This_Fun_Sql
if Req_VCID="" or isnull(Req_VCID) or Req_VCID="," then Get_VCID_VCID_TO_Del="" : exit Function
This_Fun_Sql="select VCID from FS_ME_VocationClass where ParentID in ("&FormatIntArr(Req_VCID)&")"
on error resume next
set VClass_Rs=User_Conn.execute(This_Fun_Sql)
do while not VClass_Rs.eof 
	Str_Tmp = Str_Tmp & VClass_Rs(0) & ","	
	VClass_Rs.movenext
loop	
VClass_Rs.close
Get_VCID_VCID_TO_Del = Str_Tmp
	'User_Conn.execute("delete from FS_ME_VocationClass where VCID in ("&Req_VCID&")")
	'response.Redirect("../Success.asp?ErrorUrl=User/VocationClass.asp&ErrCodes=<li>��ϲ��ɾ���ɹ���</li>")
End Function

Function Get_ParID_ParID_TO_Save(Req_ParID)
''ѭ�����ã�ͨ�������ParentID�õ����ϵ�����ParentID���������������һ�𴫸�SaveClass�������ж��Ƿ񳬹�4��.
Dim Str_Tmp,This_Fun_Sql
if Req_ParID="" or isnull(Req_ParID) or Req_ParID="," then Get_ParID_ParID_TO_Save="" : exit Function
This_Fun_Sql="select ParentID from FS_ME_VocationClass where VCID in ("& FormatIntArr(Req_ParID)&")"
on error resume next
set VClass_Rs=User_Conn.execute(This_Fun_Sql)
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
	This_Fun_Sql = "select ParentID from FS_ME_VocationClass where VCID="&NoSqlHack(View_ID)
	Set VClass_Rs = User_Conn.execute(This_Fun_Sql)
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
	This_Fun_Sql = "select ParentID from FS_ME_VocationClass where VCID="&NoSqlHack(View_ID)
	Set VClass_Rs = User_Conn.execute(This_Fun_Sql)
	if not VClass_Rs.eof then 
		set VClass_Rs1 = User_Conn.execute( "select vClassName from FS_ME_VocationClass where VCID="&VClass_Rs(0) )
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
''��㵼������ session("TopMenu_DaoHang_VCID_List") = ',1,2,3,4,'
Dim str_Req_VCID,str_Session,Sql_Session,Per_ID_Session,Str_Tmp,Str_Remove,Arr_Tmp,This_Fun_Sql
	if request.QueryString("VCID")="" or request.QueryString("VCID")="0" then 
		session("TopMenu_DaoHang_VCID_List")="" : TopMenu_DaoHang = "" : exit Function
	end if
	''**************
	str_Req_VCID = NoSqlHack(request.QueryString("VCID"))
	str_Req_VCID = ","&str_Req_VCID&","
	'if session("TopMenu_DaoHang_VCID_List") = "" then session("TopMenu_DaoHang_VCID_List") = str_Req_VCID
	str_Session = session("TopMenu_DaoHang_VCID_List")
	if str_Session = "" or isEmpty(str_Session) then str_Session = str_Req_VCID 
	if left(str_Session,1)<>"," then str_Session = ","&str_Session
	if right(str_Session,1)<>"," then str_Session = str_Session&","
	session("TopMenu_DaoHang_VCID_List") = str_Session
	'response.Write("ԭ"&session("TopMenu_DaoHang_VCID_List"))
	',1,2,3,4,' �� ',3,'
	if right(str_Session,len(str_Req_VCID)) <> str_Req_VCID then 
	if instr(str_Session,str_Req_VCID)>0 then
	''��ȡ
		str_Session = mid(str_Session,1,instr(str_Session,str_Req_VCID) + len(str_Req_VCID) - 1)
		session("TopMenu_DaoHang_VCID_List") = str_Session
		'response.Write("��"&session("TopMenu_DaoHang_VCID_List"))
	else
	''���
		str_Session = str_Session & NoSqlHack(request.QueryString("VCID")) &","
		session("TopMenu_DaoHang_VCID_List") = str_Session
		'response.Write("��"&session("TopMenu_DaoHang_VCID_List"))
	end if
	end if
	''**************
	Sql_Session = str_Session
	if left(Sql_Session,1)="," then Sql_Session = mid(Sql_Session,2,len(Sql_Session))
	if right(Sql_Session,1)="," then Sql_Session = mid(Sql_Session,1,len(Sql_Session) - 1)
	'response.Write(session("TopMenu_DaoHang_VCID_List"))
	''**************
	''˳����ʵ�����
	str_Session = ""
	Arr_Tmp = split(Sql_Session,",")	
	for each Per_ID_Session in Arr_Tmp
		This_Fun_Sql = "select VCID,vClassName from FS_ME_VocationClass where VCID = "&NoSqlHack(Per_ID_Session)
		Set VClass_Rs = User_Conn.execute(This_Fun_Sql)
		if not VClass_Rs.eof then 
			Str_Tmp = Str_Tmp & "<a href=""VocationClass.asp?Act=View&VCID=" &VClass_Rs(0)&""">"&VClass_Rs(1)&"<a> >> "
			str_Session = str_Session &VClass_Rs(0)& "," 
		end if
		VClass_Rs.close
	next
	session("TopMenu_DaoHang_VCID_List") = "," & str_Session	
	if right(Str_Tmp,len(" >> "))=" >> " then Str_Tmp = mid(Str_Tmp,1,len(Str_Tmp) - len(" >> "))
	TopMenu_DaoHang = Str_Tmp
	'response.Write(session("TopMenu_DaoHang_VCID_List"))
End Function

Function Get_OneClassEname(ClassID,ClassName)
'�õ���Ŀ��Ӣ����
If ClassID = "" Or ClassName = "" Then Exit Function
Dim Rs
Set Rs = User_Conn.ExeCute("Select vClassName_En From FS_ME_VocationClass Where VCID = " & NoSqlHack(ClassID) & " And vClassName = '" & NoSqlHack(ClassName) & "'")
If Rs.Eof Then
	Get_OneClassEname = ""
Else
	Get_OneClassEname = Rs(0)
End If
Rs.Close : Set Rs = NOthing		
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
-->
</script>
</HEAD>
<script language="JavaScript" src="../../FS_Inc/GetLettersByChinese.js"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<%if request.QueryString("Act")="Edit" then%><script language="javascript" src="../../FS_Inc/class_liandong.js" type="text/javascript"></script><%end if%>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
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
Dim View_VCID,IsOk,VClass_Rs1
IsOk = false
View_VCID = NoSqlHack(request.QueryString("VCID"))
if View_VCID<>"" then if isnumeric(View_VCID) then if View_VCID>0 then IsOk = true
if IsOk=false then 
	str_Url_Add = "&VCID=0&VCText=��"	
else	
	set VClass_Rs1 = User_Conn.execute( "select vClassName from FS_ME_VocationClass where VCID="&NoSqlHack(View_VCID) )
	if not VClass_Rs1.eof then 
		str_Url_Add = "&VCID="&View_VCID&"&VCText="	& VClass_Rs1(0)
	else
		str_Url_Add = "&VCID=0&VCText=��"	
	end if
	VClass_Rs1.close
	set VClass_Rs1=nothing		
end if		
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="VCForm" id="VCForm" method="post" action="?Act=Del">
  <tr  class="hback"> 
    <td class="xingmu" colspan="5">��ҵ�������</td>
  </tr>
    <tr  class="hback"> 
      <td align="left" colspan="5">���ർ����<a href="VocationClass.asp" >������ҳ</a> 
        >> <%=TopMenu_DaoHang()%></td>
    </tr>
    <tr  class="hback"> 
      <td align="left" class="xingmu">��������</td>
      <td width="20%" align="center" class="xingmu">Ӣ������</td>
	  <td width="10%" align="center" class="xingmu">��������Ŀ</td>
      <td width="10%" align="center" class="xingmu">����</td>
      <td width="5%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_VClass(request.QueryString("VCID")) )
	%>
  </form>
</table>
<%End Sub
Sub AddClass()
	if not MF_Check_Pop_TF("ME026") then Err_Show 
dim View_Par_ID
if request.QueryString("VCID")<>"" then 
	View_Par_ID = NoSqlHack(request.QueryString("VCID"))
else
	View_Par_ID = NoSqlHack(request.QueryString("ParID"))
end if		
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="VCForm" id="VCForm" method="post" action="?Act=Save">
    <tr  class="hback"> 
      <td align="left" class="xingmu" colspan="5">���ർ����<a href="VocationClass.asp" >������ҳ</a> 
        >> <%=TopMenu_DaoHang()%></td>
    </tr>
    <tr  class="hback"> 
      <td colspan="2" align="left" class="xingmu" >������ҵ��Ϣ ��ҵ��������ļ�</td>
    </tr>
    <tr  class="hback"> 
      <td width="171" align="right">��һ����ҵ����</td>
      <td width="803">
	  <input name="VCText" id="VCText" readonly type="text" size="50" value="<%=request.QueryString("VCText")%>"/>
	  <input type="hidden" name="ParID" id="ParID" value="<%=View_Par_ID%>"/>
	  </td>
    </tr>
    <tr class="hback"> 
      <td align="right">�µ���ҵ���ƣ�</td>
      <td width="803"><input onBlur="SetClassEName(this.value,document.VCForm.vClassName_En);" name="vClassName" type="text" id="vClassName" size="50" onBlur="if(this.value==''){Chk_vCName.innerText='���������д';this.focus();} else if(Chk_vCName.innerText!='')Chk_vCName.innerText=''"><span id="Chk_vCName" class="tx"></span> 
      </td>
    </tr>
    <tr class="hback"> 
      <td align="right">��ӦӢ������</td>
      <td width="803"><input name="vClassName_En" type="text" id="vClassName_En" size="50"  onKeyUp="value=value.replace(/[^a-zA-Z0-9_-]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^a-zA-Z0-9_-]/g,''))"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
  		<tr>
		  <td align="center">
			<input type="submit" <%call Say_Limit_Class("AddMode")%> name="SaveClass" id="SaveClass" value=" ���� " onClick="if(vClassName.value==''){alert('��Ҫ�Ĳ���������д');vClassName.focus();return false}" /> &nbsp;
			<input type="reset" name="ReSet" id="ReSet" value=" ���� " />
		  </td>
		  <td width="50" align="right">
			<a href="VocationClass.asp?Act=View&VCID=<%=Get_PatID_TO_View(View_Par_ID)%>">���ظ���</a>
		  </td>
		  <td width="50" align="right">
			<a href="#" onClick="javascript:history.back();">����</a>
		  </td>
		</tr>  
      </table>
      </td>
    </tr>	
  </form>
</table>
<%End Sub
Sub EditClass()
	if not MF_Check_Pop_TF("ME026") then Err_Show 

%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="VCForm" id="VCForm" method="post" action="?Act=Save">
    <tr  class="hback"> 
      <td align="left" class="xingmu" colspan="5">���ർ����<a href="VocationClass.asp" >������ҳ</a> >> <%=TopMenu_DaoHang()%></td>
    </tr>
    <tr  class="hback"> 
      <td colspan="2" align="left" class="xingmu" >�޸���ҵ��Ϣ ��ҵ��������ļ�</td>
    </tr>
	<tr  class="hback"> 
      <td width="168" align="right">����ϼ����<br><span class="tx">�������������</span></td>
      <td width="806">
<!---�����˵���ʼ--->	
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
<!---�����˵�����--->		
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��ҵ����</td>
      <td width="806">
	  <input name="vClassName" id="vClassName" type="text" size="50" value="<%=request.QueryString("VCText")%>" onBlur="if(this.value==''){Chk_vCName.innerText='���������д';this.focus()} else if(Chk_vCName.innerText!='')Chk_vCName.innerText=''"><span id="Chk_vCName" class="tx"></span>
      <input type="hidden" name="VCID" id="VCID" value="<%=request.QueryString("VCID")%>">
  	  <!--<input type="text" name="ParentID" id="ParentID" value="">-->
	  </td>
    </tr>
    <tr class="hback"> 
      <td align="right">��ӦӢ������</td>
      <td width="806"><input name="vClassName_En" type="text" id="vClassName_En" size="50"  onKeyUp="value=value.replace(/[^a-zA-Z0-9_-]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^a-zA-Z0-9_-]/g,''))" value="<% = Get_OneClassEname(request.QueryString("VCID"),request.QueryString("VCText")) %>"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
  		<tr>
		  <td align="center">
			<input type="submit" <%call Say_Limit_Class("AddMode")%> name="SaveClass" id="SaveClass" value=" ���� " onClick="if(vClassName.value==''){vClassName.focus();return false}" /> &nbsp;
			<input type="reset" name="ReSet" id="ReSet" value=" ���� " />
		  </td>      
		  <td width="50" align="right">
			<a href="VocationClass.asp?Act=View&VCID=<%=Get_PatID_TO_View(request.QueryString("VCID"))%>">���ظ���</a>
		  </td>
		  <td width="50" align="right">
			<a href="#" onClick="javascript:history.back();">����</a>
		  </td>
		</tr>  
      </table>
      </td>
    </tr>	
  </form>
</table>
<%End Sub%>
</body>
<script language="javascript">
<!-- 
<%if request.QueryString("Act")="Edit" then%>
//awen created
//�����˵�---��ҵ���   ���4��   --start 
//���ݸ�ʽ ID������ID������
var array=new Array();
<%dim sql,rs,i
  sql="select VCID,ParentID,vClassName from FS_ME_VocationClass where VCID<>"&NoSqlHack(request.QueryString("VCID"))
  set rs=User_Conn.execute(sql)
  i=0
  do while not rs.eof
%>
array[<%=i%>]=new Array("<%=rs("VCID")%>","<%=rs("ParentID")%>","<%=rs("vClassName")%>"); 
<%
	rs.movenext
	i=i+1
loop
rs.close
%>

var liandong=new CLASS_LIANDONG_YAO(array)
liandong.firstSelectChange("0","vclass1");
liandong.subSelectChange("vclass1","vclass2");
liandong.subSelectChange("vclass2","vclass3");
liandong.subSelectChange("vclass3","vclass4");

//---------------------------������������������
function RemoveChildopt(obj,StrList)
{
/*
	if (StrList=='') 
	if (document.getElementById('ParentID')!=null) 
	{
		if (obj.value!='')
		{document.getElementById('ParentID').value=obj.value;return;}
		else
		{document.getElementById('ParentID').value=document.getElementById('vclass3').value;return;}
	}
*/	
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
/*	else
	//ѡ��ʵ����Ŀ��ֵ�����������Ա��޸ĵ�ʱ����̡�
	{
		var Tmpstr = '';
		for (var i=TmpArr.length-1 ; i>=0; i--)
		{			
			if (TmpArr[i]!='')
			{	if (document.getElementById(TmpArr[i]).selectedIndex>1)
				Tmpstr += document.getElementById(TmpArr[i]).options[document.getElementById(TmpArr[i]).selectedIndex].value;
			}
		}	
		//alert(document.VCForm.ParentID.value);	
		if (Tmpstr=='')
		{
			//����������				
			if (document.getElementById('ParentID')!=null) document.all.ParentID.value=obj.value;
		}
	}
*/		
} 
<%end if%>
function SetClassEName(Str,Obj)
{
	Obj.value=ConvertToLetters(Str,1);
}
-->
</script>
<%
Set VClass_Rs=nothing
User_Conn.close
Set User_Conn=nothing
%>
</html>
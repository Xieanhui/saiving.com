<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%
	Dim Conn,User_Conn
	MF_Default_Conn
	MF_User_Conn
	'session�ж�
	MF_Session_TF 
	if not MF_Check_Pop_TF("DS_Class") then Err_Show
	Dim Fs_news
	set Fs_news = new cls_news
	If Not Fs_news.IsSelfRefer Then response.write "�Ƿ��ύ����":Response.end
	'��ȡ����
	Dim obj_Save_Rs1,ClassID1,str_ClassKeywords1,str_Classdescription1,str_action1,strShowErr
	Dim str_ClassID1,lng_OrderID1,str_ClassName1,str_ClassEName1,str_ParentID1,str_Templet1,str_NewsTemplet1,str_Domain1,lng_AdminID1,lng_RefreshNumber
	Dim  lng_GroupID1,lng_PointNumber1,flt_Money1,str_FileExtName1,dtm_Addtime1,int_isConstr1,int_IsURL1,str_UrlAddress1,lng_Oldtime1,int_isShow1
	Dim str_ClassNaviContent1,str_ClassNaviPic1,lng_DefineID1,int_NewsCheck1,int_AddNewsType1,str_SavePath1,str_FileSaveType1,int_isConstrDel1,str_GetParentID
	str_action1 = NoSqlHack(Request.Form("str_add"))
	str_ClassID1 = NoSqlHack(Request.Form("ClassID"))
	lng_OrderID1 = NoSqlHack(Request.Form("OrderID"))
	str_ClassName1 = NoSqlHack(Request.Form("ClassName"))
	str_ClassEName1 = NoSqlHack(Request.Form("ClassEName"))
	str_ParentID1 = NoSqlHack(Request.Form("ParentID"))
	str_Templet1 = NoSqlHack(Request.Form("Templet"))
	str_NewsTemplet1 = NoSqlHack(Request.Form("NewsTemplet"))
	str_Domain1 = NoSqlHack(Request.Form("Domain"))
	lng_AdminID1 = NoSqlHack(Request.Form("ClassAdmin"))
	lng_RefreshNumber = NoSqlHack(Request.Form("RefreshNumber"))
	lng_GroupID1 = NoSqlHack(Request.Form("BrowPop"))
	lng_PointNumber1 = NoSqlHack(Request.Form("PointNumber")) 
	flt_Money1 = NoSqlHack(Request.Form("Money"))
	str_FileExtName1 = NoSqlHack(Request.Form("FileExtName"))
	dtm_Addtime1 = NoSqlHack(Request.Form("Addtime"))
	int_isConstr1 = NoSqlHack(Request.Form("isConstr"))
	int_IsURL1 = NoSqlHack(Request.Form("IsURL"))
	str_UrlAddress1 = NoSqlHack(Request.Form("UrlAddress"))
	lng_Oldtime1 = NoSqlHack(Request.Form("Oldtime"))
	int_isShow1 = NoSqlHack(Request.Form("isShow"))
	
	str_ClassNaviContent1 = NoSqlHack(Request.Form("ClassNaviContent"))
	str_ClassNaviPic1 = NoSqlHack(Request.Form("ClassNaviPic"))
	lng_DefineID1 = NoSqlHack(Request.Form("DefineID"))
	int_NewsCheck1 = NoSqlHack(Request.Form("NewsCheck"))
	int_AddNewsType1 = NoSqlHack(Request.Form("AddNewsType"))
	if  Trim(Request.Form("SavePath")) = "" then
		str_SavePath1 = "/"
	Else
		str_SavePath1 = NoSqlHack(Request.Form("SavePath"))
	End if
	str_FileSaveType1 = NoSqlHack(Request.Form("FileSaveType"))
	int_isConstrDel1 = NoSqlHack(Request.Form("isConstrDel"))
	str_ClassKeywords1  = NoSqlHack(Request.Form("ClassKeywords"))
	str_Classdescription1  = NoSqlHack(Request.Form("Classdescription"))
	'�ж������Ƿ���ȷ
	if str_Domain1 <>"" then
		if len(Trim(str_Domain1))<6  then
			strShowErr = "<li>����ȷ��д���Ķ�������</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
	End if
	if trim(int_IsURL1) <>"" then 
		if isnull(Trim(str_UrlAddress1))  then
			strShowErr = "<li>����д�ⲿ��ַ</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		if Trim(str_UrlAddress1)="http://" then
			strShowErr = "<li>����д�ⲿ��ַ</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		if trim(str_ClassName1) = ""  then
			strShowErr = "<li>����д�ⲿ��Ŀ����</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
	Else
			if trim(str_SavePath1) = ""  then
				strShowErr = "<li>����д��Ŀ����·��</li>"
				Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			if trim(str_ClassName1) = "" or trim(str_ClassEName1) = ""  or trim(str_Templet1) = ""  or trim(str_NewsTemplet1) = ""  or trim(str_SavePath1) = ""  then
				strShowErr = "<li>��*���Ǳ�����д��</li>"
				Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			if isdate(dtm_Addtime1) =false then
				strShowErr = "<li>����д��ȷ�����ڸ�ʽ</li>"
				Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			if isnumeric(lng_Oldtime1) =false or isnumeric(lng_RefreshNumber) = false then
				strShowErr = "<li>�鵵���ڲ�����ȷ������</li>"
				Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			if isnumeric(lng_OrderID1) = false then
				strShowErr = "<li>����Ȩ��(���)������ȷ������</li>"
				Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			if fs_news.chkinputchar(str_ClassEName1) = false then
				strShowErr = "<li>Ӣ������ֻ��ΪӢ�ġ����ּ��»���</li>"
				Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			if trim(lng_GroupID1)<>"" or lng_PointNumber1 <>"" or flt_Money1<>"" then 
				if trim(str_FileExtName1)<>"asp" then
						strShowErr = "<li>�����������Ȩ�ޣ���չ������Ϊ.asp</li>"
						Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
						Response.end
				End if
			End if
			if Len(str_ClassKeywords1) >200 Then
				strShowErr = "<li>��ĿMETA�ؼ��ֲ��ܳ���200</li>"
				Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			if Len(str_Classdescription1) >200 Then
				strShowErr = "<li>��ĿMETA�������ܳ���200</li>"
				Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
	End if
	Dim GetClassReturnValue1,obj_SaveTF_Rs1,obj_TF_Rs1,Newsadd_SQL1,NewsaddTF_SQL1
	GetClassReturnValue1 = Fs_news.GetRamCode(15)
	Set obj_Save_Rs1 = server.CreateObject(G_FS_RS)
	If str_action1 = "add" then
		if not MF_Check_Pop_TF("DS010") then Err_Show
		Set obj_SaveTF_Rs1 = server.CreateObject(G_FS_RS)
		obj_SaveTF_Rs1.Open "Select ID from FS_DS_Class where ClassID='"& NoSqlHack(GetClassReturnValue1) &"' order by id desc",Conn,1,3
		if  Not obj_SaveTF_Rs1.eof then
				strShowErr = "<li>��ĿClassID��������ظ�������������</li>"
				Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
		End if
		set obj_SaveTF_Rs1 = nothing
		Set obj_TF_Rs1 = server.CreateObject(G_FS_RS)
		If str_ParentID1="0" Then 
			NewsaddTF_SQL1 ="Select ID from FS_DS_Class where ClassEName='"& NoSqlHack(str_ClassEName1) &"' And ParentID='0'"
		Else
			NewsaddTF_SQL1 ="Select ID from FS_DS_Class where ParentID='"&NoSqlHack(str_ParentID1)&"' and ClassEName='"& NoSqlHack(str_ClassEName1) &"'"
		End If
		obj_TF_Rs1.Open NewsaddTF_SQL1,Conn,1,3
		if Not (obj_TF_Rs1.eof and obj_TF_Rs1.bof)  then
					strShowErr = "<li>��ĿӢ�������ظ�������������</li>"
					Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
		End if
		set obj_TF_Rs1 = nothing
		Newsadd_SQL1 = "Select * from FS_DS_Class where 1=2"
		obj_Save_Rs1.Open Newsadd_SQL1,Conn,1,3
		obj_Save_Rs1.AddNew
		obj_Save_Rs1("ClassEName") = NoSqlHack(str_ClassEName1)
		obj_Save_Rs1("ClassID") = NoSqlHack(GetClassReturnValue1)
		obj_Save_Rs1("ParentID") = NoSqlHack(str_ParentID1)
	ElseIf str_action1 = "edit" then
		if not MF_Check_Pop_TF("DS011") then Err_Show
		Newsadd_SQL1 = "Select * from FS_DS_Class where ClassID='"& NoSqlHack(str_ClassID1) &"'"
		obj_Save_Rs1.Open Newsadd_SQL1,Conn,1,3
	End if
		if int_isShow1 <>"" then:obj_Save_Rs1("isShow") = 1:Else:obj_Save_Rs1("isShow") = 0:End if
		obj_Save_Rs1("Addtime") = dtm_Addtime1
		obj_Save_Rs1("RefreshNumber") = NoSqlHack(lng_RefreshNumber)
		obj_Save_Rs1("ClassNaviContent") = NoSqlHack(str_ClassNaviContent1)
		obj_Save_Rs1("ClassNaviPic") = NoSqlHack(str_ClassNaviPic1)
		if lng_OrderID1<>"" then:obj_Save_Rs1("OrderID") = clng(lng_OrderID1):Else:obj_Save_Rs1("OrderID") = 10:End if
		if int_IsURL1 <>"" then
			obj_Save_Rs1("IsURL") = 1
			obj_Save_Rs1("UrlAddress") = NoSqlHack(str_UrlAddress1)
			obj_Save_Rs1("ClassName") = NoSqlHack(str_ClassName1)
		Else
			obj_Save_Rs1("ClassName") = NoSqlHack(str_ClassName1)
			obj_Save_Rs1("Templet") = NoSqlHack(str_Templet1)
			obj_Save_Rs1("NewsTemplet") = NoSqlHack(str_NewsTemplet1)
			obj_Save_Rs1("Domain") = NoSqlHack(str_Domain1)
			obj_Save_Rs1("ClassAdmin") = NoSqlHack(lng_AdminID1)
			obj_Save_Rs1("FileExtName") = NoSqlHack(str_FileExtName1)
			if int_isConstr1 <>"" then:obj_Save_Rs1("isConstr") = 1:Else:obj_Save_Rs1("isConstr") = 0:End if
			obj_Save_Rs1("IsURL") = 0
			obj_Save_Rs1("UrlAddress") = ""
			obj_Save_Rs1("Oldtime") = clng(lng_Oldtime1)
			obj_Save_Rs1("DefineID") = NoSqlHack(lng_DefineID1)
			if int_NewsCheck1 <> "" then:obj_Save_Rs1("NewsCheck") = 1:Else:obj_Save_Rs1("NewsCheck") = 0:End if
			if int_AddNewsType1 <>"" then:obj_Save_Rs1("AddNewsType") = 0:Else:obj_Save_Rs1("AddNewsType") = 1:End if
			obj_Save_Rs1("SavePath") = NoSqlHack(str_SavePath1)
			obj_Save_Rs1("FileSaveType") = NoSqlHack(str_FileSaveType1)
			if int_isConstrDel1 <>"" then:obj_Save_Rs1("isConstrDel") = 1:Else:obj_Save_Rs1("isConstrDel") = 0:End if
			if Trim(lng_GroupID1) <>"" or lng_PointNumber1 <> "" or flt_Money1<>"" then:obj_Save_Rs1("isPop") = 1:Else:obj_Save_Rs1("isPop") = 0:End if
			obj_Save_Rs1("ClassKeywords") = NoSqlHack(str_ClassKeywords1)
			obj_Save_Rs1("Classdescription") = NoSqlHack(str_Classdescription1)
		End if
		'����Ȩ�����ݱ�
	'	lng_GroupID1,lng_PointNumber1,flt_Money1
		if Trim(lng_GroupID1) <>"" or lng_PointNumber1 <> "" or flt_Money1<>"" then 
			Dim obj_insert_rs
			set obj_insert_rs = Server.CreateObject(G_FS_RS)
			If str_action1 = "add" then
				obj_insert_rs.Open "select  GroupName,PointNumber,FS_Money,InfoID,PopType,isClass From FS_MF_POP",Conn,1,3
				obj_insert_rs.addnew
				obj_insert_rs("InfoID")=GetClassReturnValue1
			elseIf str_action1 = "edit" then
				obj_insert_rs.Open "select  GroupName,PointNumber,FS_Money,InfoID,PopType,isClass From FS_MF_POP  where InfoID='"& NoSqlHack(str_ClassID1) &"' and PopType='DS' and isClass=1",Conn,1,3
				obj_insert_rs("InfoID")=NoSqlHack(str_ClassID1)
			End if
			obj_insert_rs("GroupName")=NoSqlHack(lng_GroupID1)
			if lng_PointNumber1 <>""  then:obj_insert_rs("PointNumber")=NoSqlHack(lng_PointNumber1):Else:obj_insert_rs("PointNumber")=0:End if
			if flt_Money1 <>"" then:obj_insert_rs("FS_Money")=NoSqlHack(flt_Money1):Else:obj_insert_rs("FS_Money")=0:End if
			obj_insert_rs("PopType")="DS"
			obj_insert_rs("isClass")=1
			obj_insert_rs.update
			obj_insert_rs.close:set obj_insert_rs = nothing
		End if
		'������ڲ����ӣ������ɾ�̬Ŀ¼
	'���ɾ�̬Ŀ¼
	'**************
	' ����xml
	 
	obj_Save_Rs1.update
	obj_Save_Rs1.close
	set obj_Save_Rs1 = nothing
	Call Makexml(NoSqlHack(str_ParentID1))
	strShowErr = "<li>��ϲ����Ŀ����ɹ�</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Down/Class_Manage.asp")
	Response.end
	set Fs_news = nothing 
%>







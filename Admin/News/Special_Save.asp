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
	'Ȩ���ж�
	'Call MF_Check_Pop_TF("NS_Special_000001")
	'�õ���Ա���б�
	Dim Fs_news,SaveReturnvalue
	set Fs_news = new cls_news
	If Not Fs_news.IsSelfRefer Then response.write "�Ƿ��ύ����":Response.end
	'��ȡ����
	Dim obj_Save_Rs,strShowErr 
	Dim lng_SpecialID,str_Templet ,str_SpecialCName,str_SpecialEName,str_SpecialSize,str_SpecialContent,naviPic,adminName
	Dim str_SavePath,str_ExtName,bit_isLock,dtm_Addtime,int_sPoint,lng_GroupID,Arr_Tmp,Int_SaveType
	lng_SpecialID = NoSqlHack(request.Form("SpecialID"))
	str_SpecialCName = NoSqlHack(request.Form("SpecialCName"))
	str_SpecialEName = NoSqlHack(request.Form("SpecialEName"))
	str_SpecialSize = NoSqlHack(request.Form("SpecialSize"))
	str_SpecialContent = NoSqlHack(request.Form("SpecialContent"))
	str_SavePath = NoSqlHack(request.Form("SavePath"))
	str_Templet = NoSqlHack(request.Form("Templet"))
	str_ExtName = NoSqlHack(request.Form("ExtName"))
	bit_isLock = NoSqlHack(request.Form("isLock"))
	dtm_Addtime = NoSqlHack(request.Form("Addtime"))
	naviPic = NoSqlHack(request.Form("naviPic"))
	adminName= NoSqlHack(request.Form("adminName"))
	Int_SaveType = NoSqlHack(request.Form("SaveType"))
	'�ж������Ƿ���ȷ
	if trim(str_SpecialCName) = "" or trim(str_SpecialEName) = "" or trim(str_ExtName) = "" or trim(str_Templet) = ""  or trim(str_SpecialSize) = ""  or trim(str_SavePath) = ""  then
		strShowErr = "<li>��*���Ǳ�����д��</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	if instr(str_SpecialSize,",")=0 then 
		strShowErr = "<li>��ʽ���󣬸�ʽ���߶�,��ȣ�150,120��û�ж��š�</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		Arr_Tmp = split(str_SpecialSize,",")
		if ubound(Arr_Tmp)<>1 then 
			strShowErr = "<li>��ʽ���󣬸�ʽ���߶�,��ȣ�150,120������Ķ��š�</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end		
		else
			if (not (isnumeric(Arr_Tmp(0)) and Arr_Tmp(0)>=0 )) or (not (isnumeric(Arr_Tmp(1)) and Arr_Tmp(1)>=0) ) then 
				strShowErr = "<li>��ʽ���󣬸�ʽ���߶�,��ȣ�150,120������Ϊ�Ǹ����֡�</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end		
			end if
		end if	
	end if
	if isdate(dtm_Addtime) =false then
		strShowErr = "<li>����д��ȷ�����ڸ�ʽ</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	'if not isnull(int_sPoint) then 
	'if isnumeric(int_sPoint) =false then
	'	strShowErr = "<li>��Ҫ����������ȷ������</li>"
	'	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	'	Response.end
	'End if
	'end if
	if fs_news.chkinputchar(str_SpecialEName) = false then
	strShowErr = "<li>Ӣ������ֻ��ΪӢ�ġ����ּ��»���</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	'if not isnull(int_sPoint) then 
	'	if int_sPoint <= 0 then int_sPoint = null
	'end if	
	'if int_sPoint="" then int_sPoint = null
	'if not isnull(int_sPoint) or lng_GroupID<>"" then str_ExtName = "asp"
		
	Dim obj_SaveTF_Rs,obj_SaveENameTF_Rs
	Set obj_Save_Rs = server.CreateObject(G_FS_RS)
	If Request.Form("Action") = "add" then
		if not Get_SubPop_TF("","NS026","NS","specail") then Err_Show
		Set obj_SaveENameTF_Rs = server.CreateObject(G_FS_RS)
		obj_SaveENameTF_Rs.Open "Select SpecialEName from FS_NS_Special where SpecialEName='"& NoSqlHack(str_SpecialEName) &"' order by SpecialID desc",Conn,1,1
		if Not obj_SaveENameTF_Rs.eof then
					strShowErr = "<li>Ӣ�������ظ�������������</li>"
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
		End if
		set obj_SaveENameTF_Rs = nothing
		obj_Save_Rs.Open "Select * from FS_NS_Special where 1=2",Conn,1,3
		obj_Save_Rs.AddNew
	ElseIf Request.Form("Action") = "edit" then
		obj_Save_Rs.Open "Select * From FS_NS_Special where SpecialID="& NoSqlHack(lng_SpecialID) ,Conn,1,3
	End if
		obj_Save_Rs("SpecialCName") = str_SpecialCName
		obj_Save_Rs("SpecialEName") = str_SpecialEName
		obj_Save_Rs("SpecialSize") = str_SpecialSize
		obj_Save_Rs("SpecialContent") = str_SpecialContent
		obj_Save_Rs("SavePath") = str_SavePath
		obj_Save_Rs("Templet") = str_Templet
		obj_Save_Rs("ExtName") = str_ExtName
		if bit_isLock<>"" then
			obj_Save_Rs("isLock") = 1
		else
			obj_Save_Rs("isLock") = 0
		end if
		obj_Save_Rs("Addtime") = dtm_Addtime
		obj_Save_Rs("naviPic") = naviPic
		obj_Save_Rs("adminName") = adminName
		obj_Save_Rs("FileSaveType") = Int_SaveType
		'obj_Save_Rs("GroupID") = lng_GroupID
		'obj_Save_Rs("sPoint") = int_sPoint
	'������ڲ����ӣ������ɾ�̬Ŀ¼
	'���ɾ�̬Ŀ¼
	'**************
	obj_Save_Rs.update
	obj_Save_Rs.close
	set obj_Save_Rs = nothing
	strShowErr = "<li>��ϲ��ר�Ᵽ��ɹ�</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Special_Manage.asp")
	set Fs_news = nothing 
	Response.end
%>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. --> 







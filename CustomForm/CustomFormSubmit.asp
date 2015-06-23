<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp"-->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<!--#include file="../user/lib/Cls_user.asp" -->
<!--#include file="../Admin/CommPages/SelectManageDir/News_Upfile.asp"-->
<%
dim Conn,User_Conn,formid,obj_form_rs,i,VerifyLogin,DataInitStatus,VerifyCode
dim formName,tableName,upfileSaveUrl,stateSet,TimeLimited,StartTime,EndTime,sql,RemoteIP,strShowErr
dim SubmitType,GoldFactor,PointFactor,UserGroup,UserOnce,Validate,remark,ArrUserGroup,upfileSize
MF_Default_Conn
MF_User_Conn

Set UpFileObj = New UpFileClass
UpFileObj.GetData
formid=NoSqlHack(UpFileObj.Form("formid"))
VerifyCode=NoSqlHack(UpFileObj.Form("VerifyCode"))
if VerifyCode <> Session("CustomFormGetCode") then AlertInfo "��֤�벻��ȷ��",True
Session("CustomFormGetCode") = ""
if formid="" then AlertInfo "���������ݲ���ȷ��",True
sql="select formName,tableName,upfileSaveUrl,upfileSize,state,TimeLimited,StartTime,EndTime,SubmitType,GoldFactor,PointFactor,UserGroup,UserOnce,Validate,remark,VerifyLogin,DataInitStatus from FS_MF_CustomForm where id="&formid&" and state=0"
set obj_form_rs=conn.execute(sql)
if obj_form_rs.eof then AlertInfo "���ѹرջ��ѱ�ɾ����",True
VerifyLogin=obj_form_rs("VerifyLogin")
Set Fs_User = New Cls_User
if VerifyLogin = 1 then
	Dim Fs_User,ReturnValue
	if Fs_User.checkStat(Session("FS_UserName"),Session("FS_UserPassword")) = False then
		Set Fs_User = Nothing
		obj_form_rs.Close
		Set obj_form_rs = Nothing
		AlertInfo "�û�û�е�¼",True
	end if
end if

formName=obj_form_rs("formName")
tableName=obj_form_rs("tableName")
upfileSaveUrl=obj_form_rs("upfileSaveUrl")
upfileSize=obj_form_rs("upfileSize")
stateSet=obj_form_rs("state")
TimeLimited=obj_form_rs("TimeLimited")
StartTime=obj_form_rs("StartTime")
EndTime=obj_form_rs("EndTime")
SubmitType=obj_form_rs("SubmitType")
GoldFactor=obj_form_rs("GoldFactor")
PointFactor=obj_form_rs("PointFactor")
UserGroup=obj_form_rs("UserGroup")
UserOnce=obj_form_rs("UserOnce")
Validate=obj_form_rs("Validate")
DataInitStatus=obj_form_rs("DataInitStatus")
remark=obj_form_rs("remark")
'ʱ������
if TimeLimited=0 then
	if now()<cdate(startTime) or now()>cdate(EndTime) then AlertInfo "�Ѿ��������������ύ���ޣ�",True
end if
'�û�������
if UserGroup<>"" then
	dim userRestriction
	userRestriction=false
	ArrUserGroup=split(UserGroup,",")
	for i=0 to ubound(ArrUserGroup)
		if cstr(Fs_User.NumGroupID)=trim(cstr(ArrUserGroup(i))) then
			userRestriction=true
		end if
	next
	if userRestriction=false then AlertInfo "�����ڵ��û��鲻�ܹ��ύ�ñ���",True
end if
'�ύ����
'0Ϊ�����ã�1Ϊ�۽�ң�2Ϊ�ۻ��֣�3Ϊ�۽�Һͻ��֣�4Ϊ�ﵽ��ң�5Ϊ�ﵽ���֣�6Ϊ�ﵽ��Һͻ���)
select case SubmitType
	case 1
		if Fs_User.NumFS_Money<GoldFactor then
			strShowErr = "�ύ�������۳���"&GoldFactor&"��ң����ĵ�ǰ���ý��Ϊ��"&Fs_User.NumFS_Money&"�����Ľ�Ҳ�����֧����д������"
		    AlertInfo strShowErr,True
		end if
	case 2
		if Fs_User.NumIntegral<PointFactor then
			strShowErr = "�ύ�������۳���"&PointFactor&"���֣����ĵ�ǰ���û���Ϊ��"&Fs_User.NumIntegral&"�����Ļ��ֲ�����֧����д������"
		    AlertInfo strShowErr,True
		end if
	case 3
		if Fs_User.NumFS_Money<GoldFactor then
			strShowErr = "�ύ�������۳���"&PointFactor&"���ּ�"&GoldFactor&"</font> ��ң����ĵ�ǰ���ý��Ϊ��"&Fs_User.NumFS_Money&"�����Ľ�Ҳ�����֧����д������"
		    AlertInfo strShowErr,True
		end if
		if Fs_User.NumIntegral<PointFactor then
			strShowErr = "�ύ�������۳���"&PointFactor&"���ּ�"&GoldFactor&"��ң����ĵ�ǰ���û���Ϊ��"&Fs_User.NumIntegral&"�����Ļ��ֲ�����֧����д������"
		    AlertInfo strShowErr,True
		end if
	case 4
		if Fs_User.NumFS_Money<GoldFactor then
			strShowErr = "�ύ����Ҫ��ﵽ"&GoldFactor&"��ң����ĵ�ǰ���ý��Ϊ��"&Fs_User.NumFS_Money&"��"
		    AlertInfo strShowErr,True
		end if
	case 5
		if Fs_User.NumIntegral<PointFactor then
			strShowErr = "�ύ����Ҫ��ﵽ"&PointFactor&"���֣����ĵ�ǰ���û���Ϊ��"&Fs_User.NumFS_Money&"��"
		    AlertInfo strShowErr,True
		end if
	case 6
		if Fs_User.NumFS_Money<GoldFactor then
			strShowErr = "�ύ����Ҫ��ﵽ"&GoldFactor&"��Һʹﵽ"&PointFactor&"���֣����ĵ�ǰ���ý��Ϊ��"&Fs_User.NumFS_Money&"��"
		    AlertInfo strShowErr,True
		end if
		if Fs_User.NumIntegral<PointFactor then
			strShowErr = "�ύ����Ҫ��ﵽ <font color=red>"&GoldFactor&"��Һʹﵽ"&PointFactor&"���֣����ĵ�ǰ���û���Ϊ��"&Fs_User.NumFS_Money&"��"
		    AlertInfo strShowErr,True
		end if
end select
'�ظ��ύ����
RemoteIP = GetIP
if VerifyLogin = 1 then
	if UserOnce=0 then
		sql="select id from "&tableName&" where form_usernum="&Fs_User.UserID&" and form_username='"&Fs_User.UserName&"'"
		Set obj_form_rs=conn.execute(sql)
		if not obj_form_rs.eof then
			strShowErr = "���Ѿ���ñ��ύ�������ˣ��������ظ��ύ��"
			AlertInfo strShowErr,True
		end if
	end if
else
	if UserOnce=0 then
		sql="select id from "&tableName&" where form_ip='"&RemoteIP&"'"
		Set obj_form_rs=conn.execute(sql)
		if not obj_form_rs.eof then
			strShowErr = "���Ѿ���ñ��ύ�������ˣ��������ظ��ύ��"
			AlertInfo strShowErr,True
		end if
	end if
end if
i=0
'ʹ���ϴ�
dim AllowFileExtStr,MaxFileSize,UpFileObj,upFileName
if upfileSize<>"" then MaxFileSize=clng(upfileSize) Else MaxFileSize =1024 End if
AllowFileExtStr = "jpg,gif,jpeg,png,bmp,txt,doc"
dim TempData,arrTempData
dim FieldStr,ValueStr
FieldStr=""
if G_IS_SQL_DB = 1 then
	ValueStr="getdate()"
else
	ValueStr="now()"
end if
sql="select formitemid,ItemName,FieldName,IsNull,ItemType,MaxSize,DefaultValue,SelectItem,Remark from  FS_MF_CustomForm_Item where formid="&formid&" and State=0 order by orderby"
set obj_form_rs=conn.execute(sql)
do while not obj_form_rs.eof
	if obj_form_rs("ItemType")="UploadFile" then
		set TempData=UpFileObj.File(""&obj_form_rs("FieldName")&"")
		'�������
		if obj_form_rs("IsNull")=0 then
			'û���ϴ�����
			if TempData.FileSize<=0 then
				strShowErr = "����û��ѡ��"&obj_form_rs("ItemName")&"�ļ���"
				AlertInfo strShowErr,True
			end if
		end if
		if TempData.FileSize>0 then
			'��С����
			if TempData.FileSize > CLng(MaxFileSize)*1024 then
				strShowErr = obj_form_rs("ItemName")&"���������ƣ����ֻ���ϴ�" & MaxFileSize & "K���ļ�"
				AlertInfo strShowErr,True
			end if
			'�ļ���������
			if CheckFileType(AllowFileExtStr,TempData.FileExt) = False then
				ErrStr = ErrStr & FileName & "�ļ�:�������ϴ�,�ϴ��ļ�������" + AllowExtStr + "\n"
				strShowErr = obj_form_rs("ItemName")&"�ļ�:�������ϴ�,�ϴ��ļ�������" + AllowExtStr
				AlertInfo strShowErr,True
			end if
			upFileName=Year(Now())&Right("0"&Month(Now()),2)&Right("0"&Day(Now()),2)&Right("0"&Hour(Now()),2)&Right("0"&Minute(Now()),2)&Right("0"&Second(Now()),2)&GetRand(4)
			upFileName=upfileSaveUrl&"/"&upFileName&"."&TempData.FileExt
			TempData.SaveToFile server.MapPath(upFileName)
			set TempData=nothing
			TempData=upFileName
		end if
	else
		TempData=NoSqlHack(UpFileObj.Form(""&obj_form_rs("FieldName")&""))
		if obj_form_rs("IsNull")=0 then
			if TempData="" then
				strShowErr = obj_form_rs("FieldName")&"��������Ϊ�գ�"
				AlertInfo strShowErr,True
			end if
		end if
		if len(TempData)>obj_form_rs("MaxSize") and obj_form_rs("MaxSize")<>0 then
			strShowErr = obj_form_rs("FieldName")&"�ĳ��ȳ������������"&obj_form_rs("MaxSize")
		    AlertInfo strShowErr,True
		end if
	end if
	i=i+1
	FieldStr=FieldStr&","&obj_form_rs("FieldName")
	ValueStr=ValueStr&",'"&TempData&"'"
	obj_form_rs.movenext
loop

if VerifyLogin = 1 then
	sql="insert into "&tableName&"(form_usernum,form_username,form_ip,form_lock,form_time"&FieldStr&") values("&Fs_User.UserID&",'"&Fs_User.UserName&"','"&RemoteIP&"'," & DataInitStatus & ","&ValueStr&")"
else
	sql="insert into "&tableName&"(form_usernum,form_username,form_ip,form_lock,form_time"&FieldStr&") values(0,'Anonymous','"&RemoteIP&"'," & DataInitStatus & ","&ValueStr&")"
end if
conn.execute sql
Set Fs_User = Nothing
Set Conn = Nothing
AlertInfo "�����ύ�����ɹ���",True

Sub AlertInfo(Str,IsBack)
	Dim AlertStr,f_RedirectPath
	f_RedirectPath = G_VIRTUAL_ROOT_DIR & "/User/main.asp"
	if IsBack then
		AlertStr = "<script language=""javascript"">alert('" & Str & "');history.back();</script>"
	else
		AlertStr = "<script language=""javascript"">alert('" & Str & "');location='" & f_RedirectPath & "';</script>"
	end if
	Response.Write(AlertStr)
	Response.end
End Sub

Function CheckFileType(AllowExtStr,FileExtName)
	Dim i,AllowArray
	AllowArray = Split(AllowExtStr,",")
	FileExtName = LCase(FileExtName)
	CheckFileType = False
	For i = LBound(AllowArray) to UBound(AllowArray)
		if LCase(AllowArray(i)) = LCase(FileExtName) then
			CheckFileType = True
		end if
	Next
	if FileExtName="asp" or FileExtName="asa" or FileExtName="aspx" or  FileExtName="php" or  FileExtName="php3" or  FileExtName="php4"  or  FileExtName="php5"then
		CheckFileType = False
	end if
End Function
%>
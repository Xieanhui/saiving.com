<% Option Explicit %>
<!--#include file="FS_Inc/Const.asp" -->
<!--#include file="FS_Inc/Function.asp" -->
<!--#include file="FS_InterFace/MF_Function.asp" -->
<!--#include file="FS_InterFace/Dynamic_Function.asp" -->
<%
''��̬����ҳ����վͨ�ã����������˲ţ����󣬷����Լ�һЩ��Ҫ�Զ����µ�ϵͳ��
''ʹ�÷���,IIS�н�index.asp�����ȼ������õø�һЩ,�������Զ����¡�
''���߷���/index.asp?isNews=1 ����������ֶ����¡�
''�����ֹ�ɾ��ǰ̨�ľ�̬�ļ���index.html�����ݺ�̨���õ���
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
'''���µ�λ�� RefreshTime = 12*60*60 Ϊ1��12Сʱ
Dim RefreshTime,isNew
RefreshTime = 180
''�Ƿ����ɾ�̬ 0 Ϊ���� 1 Ϊ������
isNew = request.QueryString("isNew")
if not isnumeric(isNew) then isNew = 0  else isNew=cint(isNew) end if

Dim conn,Str_SD_Templet,User_Conn,tmpRS_SD,MF_Index_File_Name,Dynamic_HTML

MF_Default_Conn

set tmpRS_SD = Conn.execute("select top 1 MF_Index_Templet,MF_Index_File_Name,MF_Index_Refresh from FS_MF_Config")
if not tmpRS_SD.eof then
	Str_SD_Templet = tmpRS_SD("MF_Index_Templet")
	If G_VIRTUAL_ROOT_DIR<>"" Then
		Str_SD_Templet="/"&G_VIRTUAL_ROOT_DIR&Str_SD_Templet
	End If
	MF_Index_File_Name = tmpRS_SD("MF_Index_File_Name")
	RefreshTime = CInt(tmpRS_SD("MF_Index_Refresh"))*60
end if
tmpRS_SD.close : set tmpRS_SD=nothing

if MF_Index_File_Name="" or isnull(MF_Index_File_Name) then
	MF_Index_File_Name = "index.html"   ''��ҳ��̬���� Ϊ�������ɾ�̬
end if

if isNew=0 then
	''===================================================
	 ''�ж�MF_Index_File_Name�Ƿ�Ϸ�������Ϸ�������MF_Index_File_Name
	Dim IsExistsHtmPage_
	IsExistsHtmPage_ = IsExistsHtmPage(RefreshTime) ''12��Сʱ����һ��
	if IsExistsHtmPage_ Then
		Conn.close
		response.Redirect(MF_Index_File_Name)
		response.End()
	end if
end If
	''===================================================

if Str_SD_Templet="" then
	Str_SD_Templet="/Job/index.htm"
	If G_VIRTUAL_ROOT_DIR<>"" Then
		If G_TEMPLETS_DIR<>"" Then
			Str_SD_Templet="/"&G_VIRTUAL_ROOT_DIR&"/"&G_TEMPLETS_DIR&Str_SD_Templet
		Else
			Str_SD_Templet="/"&G_VIRTUAL_ROOT_DIR&Str_SD_Templet
		End If
	Else
		If G_TEMPLETS_DIR<>"" Then
			Str_SD_Templet="/"&G_TEMPLETS_DIR&Str_SD_Templet
		Else
			Str_SD_Templet=Str_SD_Templet
		End If
	End If
end if

MF_User_Conn

''��ȡ��̬��Ϣ

Dynamic_HTML = Get_Dynamic_Refresh_Content(Str_SD_Templet,"","MF",0,"")

if isNew = 0 then
	''���ɾ�̬
	Call WriteLocalFile(MF_Index_File_Name,Dynamic_HTML)
	''��ת����̬
	response.Redirect(MF_Index_File_Name)
else
	response.Write(	Dynamic_HTML )
end if

''-----------------
'���MF_Index_File_Name�Ƿ���� ���������ݣ��������������ڣ�Ϊtrue����ת��MF_Index_File_Name�������򷵻�false.
Function IsExistsHtmPage(HowSecond)
	'HowSecond = 12*60*60 Ϊ1��12Сʱ
	if MF_Index_File_Name = "" then IsExistsHtmPage=False : exit function
	if HowSecond="" then HowSecond=180
	Dim Fso,MyFile,PhFileName,isTrue
	PhFileName = MF_Index_File_Name
	isTrue = False
	Set Fso = CreateObject(G_FS_FSO)
	If Fso.FileExists(server.MapPath(PhFileName)) Then
		set MyFile = Fso.GetFile(server.MapPath(PhFileName))
		if (MyFile.Size > 10 and datediff("s",MyFile.DateLastModified,now()) < HowSecond) Or HowSecond<0 then
			if request.QueryString("isNew")="1" then
				MyFile.Delete(True)
				isTrue = False
			else
				isTrue = True
			end if
		else
			MyFile.Delete(True)
			isTrue = False
		end if
		set MyFile = nothing
	End If
	Set Fso = Nothing
	IsExistsHtmPage = isTrue
End Function

'������д�뱾���ļ�
Sub WriteLocalFile(PhFileName,strFile)
	Dim Fso, MyFile
	Set Fso = CreateObject(G_FS_FSO)
	Set MyFile = fso.CreateTextFile(server.MapPath(PhFileName), True)
	MyFile.Write strFile
	Set MyFile = Nothing
	Set Fso = Nothing
End Sub

%>






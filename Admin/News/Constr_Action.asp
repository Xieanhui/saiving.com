<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_InterFace/NS_Public.asp" -->
<!--#include file="../../FS_InterFace/MS_Public.asp" -->
<!--#include file="../../FS_InterFace/DS_Public.asp" -->
<!--#include file="../../FS_InterFace/ME_Public.asp" -->
<!--#include file="../../FS_InterFace/MF_Public.asp" -->
<!--#include file="../../FS_InterFace/SD_Public.asp" -->
<!--#include file="../../FS_InterFace/HS_Public.asp" -->
<!--#include file="../../FS_InterFace/AP_Public.asp" -->
<!--#include file="../../FS_InterFace/Other_Public.asp" -->
<!--#include file="../../FS_InterFace/Refresh_Function.asp" -->
<%
Dim Conn,User_Conn,FS_News,values,action,info_Rs,news_Rs,sql_cmd,sql_news_cmd,str_FileName
Dim contrPoint,contrMoney,contrAuditPoint,contrAuditMoney,Rs,MainID
Dim HaveNewsIDTF,ChedkNewsIDObj,Temp_NewsID_Str,TempNewsID
'Admin_Login_State'判断是否登陆
MF_User_Conn
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("NS_Constr") then Err_Show
Set info_Rs=Server.CreateObject(G_FS_RS)
Set news_Rs=Server.CreateObject(G_FS_RS)
Set FS_News=New Cls_News
FS_News.GetSysParam()
values=NoSqlHack(request.QueryString("values"))
values=DelHeadAndEndDot(values)'去处首尾逗号
action=NoSqlHack(request.QueryString("act"))
Set Rs=User_Conn.execute("Select top 1 contrPoint,contrMoney,contrAuditPoint,contrAuditMoney from FS_ME_SysPara")
if Rs.eof then
	Response.Redirect("../error.asp?ErrCodes=<li>请先设置会员参数</li>&ErrorUrl=")
ELse
	contrPoint=Rs("contrPoint")
	if not isnumeric(contrPoint) then contrPoint=0
	contrMoney=Rs("contrMoney")
	if not isnumeric(contrMoney) then contrMoney=0
	contrAuditPoint=Rs("contrAuditPoint")
	if not isnumeric(contrAuditPoint) then contrAuditPoint=0
	contrAuditMoney=Rs("contrAuditMoney")
	if not isnumeric(contrAuditMoney) then contrAuditMoney=0
End if
Rs.close
Set Rs=nothing
'审核
if action="audit" then
	if not MF_Check_Pop_TF("NS030") then Err_Show
	sql_news_cmd="select id,NewsID,NewsTitle,CurtTitle,Content,Keywords,Editor,isPicNews,Author,classID,Templet,NewsPicFile,SaveNewsPath,FileName,FileExtName,addtime,NewsProperty from FS_NS_News"
	if instr(values,",")=0 then
		sql_cmd="select AuditTF,ContTitle,SubTitle,ContContent,KeyWords,UserNumber,MainID,PicFile,NewsID from FS_ME_InfoContribution where contid="&CintStr(values)
		info_Rs.open sql_cmd,User_Conn,1,3
		news_Rs.open sql_news_cmd ,Conn,1,3
		news_Rs.addNew
		if not info_Rs.eof then
			news_Rs("NewsTitle")=info_Rs("ContTitle")
			news_Rs("CurtTitle")=info_Rs("SubTitle")
			news_Rs("Keywords")=info_Rs("KeyWords")
			If Fs_News.CopyFileTF = 1 Then
				news_Rs("Content")=GetFilePicPath(info_Rs("ContContent"),"Con")
				if trim(info_Rs("PicFile"))<>"" then
					news_Rs("NewsPicFile")=GetFilePicPath(info_Rs("PicFile"),"only")
					news_Rs("isPicNews")=1
				ENd if
			Else
				news_Rs("Content")=info_Rs("ContContent")
				if trim(info_Rs("PicFile"))<>"" then
					news_Rs("NewsPicFile")=info_Rs("PicFile")
					news_Rs("isPicNews")=1
				ENd if
			End If	
			news_Rs("Author")=FS_News.GetUserName(info_Rs("userNumber"))
			news_Rs("Editor")=session("Admin_Name")
			if not isnull(info_Rs("MainID")) and not info_Rs("MainID")="" then
				Set Rs= Conn.execute("select Classid from FS_NS_NewsClass where id="&info_Rs("MainID"))
				if not Rs.eof then
					MainID=Rs("Classid")
				Else
					MainID="0"
				End if
			Else
				MainID="0"
			End If
			If MainID="0" Then
				Response.Redirect("lib/error.asp?ErrCodes=<li>请选择栏目n</li>")
				Response.End()
			End if
			news_Rs("classID")=MainID
			'============================================================
			HaveNewsIDTF = False
			Do While Not HaveNewsIDTF
				Temp_NewsID_Str = FS_News.getRamCode(15)
				Set ChedkNewsIDObj = Conn.ExeCute("Select NewsID From FS_NS_News Where NewsID = '" & Temp_NewsID_Str & "'")
				If ChedkNewsIDObj.Eof Then
					TempNewsID = Temp_NewsID_Str
					HaveNewsIDTF = True
					Exit Do
				End IF
				ChedkNewsIDObj.Close : Set ChedkNewsIDObj = NOthing	
			Loop
			'===========================================================
			news_Rs("NewsID")=TempNewsID
			news_Rs("SaveNewsPath") = Fs_news.SaveNewsPath(Fs_news.fileDirRule)
			news_Rs("FileName")=Fs_news.strFileNameRule(Fs_news.fileNameRule,0,0)
			news_Rs("NewsProperty")="0,1,1,1,0,0,0,0,1,0,0"
			if trim(request.Form("NewsTemplet"))<>"" then
				news_Rs("Templet")=NoSqlHack(request.Form("NewsTemplet"))
			Else
				news_Rs("Templet")=Replace("/"& G_TEMPLETS_DIR &"/NewsClass/news.htm","//","/")
			ENd if
			Dim fileExtName
			Select Case Fs_news.fileExtName
				Case 1 fileExtName="htm"
				Case 2 fileExtName="shtml"
				Case 3 fileExtName="shtml"
				Case 4 fileExtName="shtm"
				Case Else fileExtName="html"
			End Select
			news_Rs("FileExtName")=fileExtName
			news_Rs("addtime")=Now()
			news_Rs.update
			If G_IS_SQL_DB = 0 Then
				newsID = news_Rs("ID")
			Else
				newsID = Conn.execute("select ident_current('FS_NS_News')")(0)
			End if
			str_FileName=news_Rs("FileName")
			Dim TempRsObj
			If Instr(str_FileName,"自动编号ID") Then
				str_FileName = Replace(str_FileName,"自动编号ID",newsID)
				Set TempRsObj=server.CreateObject(G_FS_RS)
				TempRsObj.open "select FileName From [Fs_NS_News] where NewsID='"&news_Rs("NewsID")&"' and ID="&Clng(newsID)&"",Conn,1,3
				if not TempRsObj.eof Then
					TempRsObj("FileName") = Replace(TempRsObj("FileName"),"自动编号ID",newsID)
					TempRsObj.update
				End If
				TempRsObj.Close
			End IF
			Dim TempRsObj_1
			If Instr(str_FileName,"唯一NewsID") Then
				str_FileName = Replace(str_FileName,"唯一NewsID",news_Rs("NewsID"))
				Set TempRsObj_1=server.CreateObject(G_FS_RS)
				TempRsObj_1.open "select FileName From [Fs_NS_News] where NewsID='"&news_Rs("NewsID")&"'",Conn,1,3
				if not TempRsObj_1.eof Then
					TempRsObj_1("FileName") = str_FileName
					TempRsObj_1.update
				End If
				TempRsObj_1.Close
			End IF
			Call Refresh("NS_news",newsID)

			User_Conn.execute("Update FS_ME_InfoContribution set NewsID="&CintStr(newsID)&" where ContID="&CintStr(values))
			User_Conn.execute("Update FS_ME_Users set Integral=Integral+"&(NoSqlHack(contrAuditPoint))&" , FS_Money=FS_Money+"&(NoSqlHack(contrAuditMoney))&",ConNumberNews=ConNumberNews+1 where usernumber='"&info_Rs("UserNumber")&"'")
		Else
			Response.Redirect("lib/error.asp?ErrCodes=<li>投稿数据异常</li>")
			Response.End()
		End if
		news_Rs.close
		info_Rs.close
	else
		if not MF_Check_Pop_TF("NS030") then Err_Show
		Dim i,contID_Array'稿件id数组,
		contID_Array=split(values,",")
		for i =0 to Ubound(contID_Array)
			sql_cmd="select AuditTF,ContTitle,SubTitle,ContContent,KeyWords,UserNumber,MainID,PicFile from FS_ME_InfoContribution where contid="&CintStr(contID_Array(i))
			info_Rs.open sql_cmd,User_Conn,1,3
			if not info_Rs.eof then
				news_Rs.open sql_news_cmd,Conn,1,3
				news_Rs.addNew
				news_Rs("NewsTitle")=info_Rs("ContTitle")
				news_Rs("CurtTitle")=info_Rs("SubTitle")
				news_Rs("Keywords")=info_Rs("KeyWords")
				If Fs_News.CopyFileTF = 1 Then
					news_Rs("Content")=GetFilePicPath(info_Rs("ContContent"),"Con")
					if trim(info_Rs("PicFile"))<>"" then
						news_Rs("NewsPicFile")=GetFilePicPath(info_Rs("PicFile"),"only")
						news_Rs("isPicNews")=1
					ENd if
				Else
					news_Rs("Content")=info_Rs("ContContent")
					if trim(info_Rs("PicFile"))<>"" then
						news_Rs("NewsPicFile")=info_Rs("PicFile")
						news_Rs("isPicNews")=1
					ENd if
				End If
				news_Rs("Author")=FS_News.GetUserName(info_Rs("UserNumber"))
				news_Rs("Editor")=session("Admin_Name")
				if not isnull(info_Rs("MainID")) and not info_Rs("MainID")="" then
				Set Rs= Conn.execute("select Classid from FS_NS_NewsClass where id="&info_Rs("MainID"))
					if not Rs.eof then
						MainID=Rs("Classid")
					Else
						MainID="0"
					End if
				Else
					MainID="0"
				End If
				If MainID="0" Then
					Response.Redirect("lib/error.asp?ErrCodes=<li>请选择栏目a</li>")
					Response.End()
				End if
				news_Rs("classID")=MainID
				'============================================================
				HaveNewsIDTF = False
				Do While Not HaveNewsIDTF
					Temp_NewsID_Str = FS_News.getRamCode(15)
					Set ChedkNewsIDObj = Conn.ExeCute("Select NewsID From FS_NS_News Where NewsID = '" & CintStr(Temp_NewsID_Str) & "'")
					If ChedkNewsIDObj.Eof Then
						TempNewsID = Temp_NewsID_Str
						HaveNewsIDTF = True
						Exit Do
					End IF
					ChedkNewsIDObj.Close : Set ChedkNewsIDObj = NOthing	
				Loop
				'===========================================================
				news_Rs("NewsID")=TempNewsID
				news_Rs("SaveNewsPath") = Fs_news.SaveNewsPath(Fs_news.fileDirRule)
				news_Rs("FileName")=Fs_news.strFileNameRule(Fs_news.fileNameRule,0,0)
				news_Rs("NewsProperty")="0,1,1,1,0,0,0,0,1,0,0"
				if trim(request.Form("NewsTemplet"))<>"" then
					news_Rs("Templet")=NoSqlHack(request.Form("NewsTemplet"))
				Else
					news_Rs("Templet")=Replace("/"& G_TEMPLETS_DIR &"/NewsClass/news.htm","//","/")
				ENd if
				select case Fs_news.fileExtName
					case 1 fileExtName="htm"
					case 2 fileExtName="shtml"
					case 3 fileExtName="shtml"
					case 4 fileExtName="shtm"
					case else fileExtName="html"
				End select
				news_Rs("FileExtName")=fileExtName
				news_Rs("addtime")=Now()
				news_Rs.update
				if G_IS_SQL_DB = 0 then'是否是sqlserver
					newsID = news_Rs("ID")
				Else
					newsID = Conn.execute("select ident_current('FS_NS_News')")(0)
				End if
				User_Conn.execute("Update FS_ME_InfoContribution set NewsID="&CintStr(newsID)&" where ContID="&CintStr(contID_Array(i)))
				str_FileName=news_Rs("FileName")
				If Instr(str_FileName,"自动编号ID") Then
					str_FileName = Replace(str_FileName,"自动编号ID",newsID)
					Set TempRsObj=server.CreateObject(G_FS_RS)
					TempRsObj.open "select FileName From [Fs_NS_News] where NewsID='"&news_Rs("NewsID")&"' and ID="&CintStr(newsID)&"",Conn,1,3
					if not TempRsObj.eof Then
						TempRsObj("FileName") =str_FileName
						TempRsObj.update
					End If
					TempRsObj.Close
				End IF
				If Instr(str_FileName,"唯一NewsID") Then
					str_FileName = Replace(str_FileName,"唯一NewsID",news_Rs("NewsID"))
					Set TempRsObj_1=server.CreateObject(G_FS_RS)
					TempRsObj_1.open "select FileName From [Fs_NS_News] where NewsID='"&news_Rs("NewsID")&"'",Conn,1,3
					if not TempRsObj_1.eof Then
						TempRsObj_1("FileName") =str_FileName
						TempRsObj_1.update
					End If
					TempRsObj_1.Close
				End if
				Call Refresh("NS_news",newsID)
				User_Conn.execute("Update FS_ME_InfoContribution set NewsID="&CintStr(newsID)&" WHERE ContID = "&CintStr(contID_Array(i)))
				User_Conn.execute("Update FS_ME_Users set Integral=Integral+"&CintStr(contrAuditPoint)&" , FS_Money=FS_Money+"&CintStr(contrAuditMoney)&",ConNumberNews=ConNumberNews+1 where usernumber='"&info_Rs("UserNumber")&"'")
			Else
				Response.Redirect("lib/error.asp?ErrCodes=<li>投稿数据异常</li>")
				Response.End()
			ENd if
			info_Rs.close
			news_Rs.close
		next
	End if
	User_Conn.execute("Update FS_ME_InfoContribution set audittf=1 where ContID in("&FormatIntArr(values)&")")
'撤消审核
elseif action="recall" then
	if not MF_Check_Pop_TF("NS030") then Err_Show
	User_Conn.execute(sql_cmd)
'管理员锁定
elseif action="lock" then
	if not MF_Check_Pop_TF("NS031") then Err_Show
	sql_cmd="update  FS_ME_InfoContribution set AdminLock=1 where ContID in("&FormatIntArr(values)&")"
	User_Conn.execute(sql_cmd)
'解除锁定
elseif action="unlock" then
	if not MF_Check_Pop_TF("NS030") then Err_Show
	sql_cmd="update  FS_ME_InfoContribution set AdminLock=0 where ContID in("&FormatIntArr(values)&")"
	User_Conn.execute(sql_cmd)
'设为推荐
elseif action="tf" then
	sql_cmd="update FS_ME_InfoContribution set isTF=1 where ContID in("&FormatIntArr(values)&")"
	User_Conn.execute(sql_cmd)
'取消推荐
elseif action="untf" then
	if not MF_Check_Pop_TF("NS_Constr") then Err_Show
	sql_cmd="update FS_ME_InfoContribution set isTF=0 where ContID in("&FormatIntArr(values)&")"
	User_Conn.execute(sql_cmd)
'删除
ElseIf action="delete" Then
	if not MF_Check_Pop_TF("NS032") then Err_Show
	sql_cmd="delete From FS_ME_InfoContribution where ContID in ("&FormatIntArr(values)&")"
	User_Conn.execute(sql_cmd)
ElseIf action="deleteAll" Then
	if not MF_Check_Pop_TF("NS032") then Err_Show
	sql_cmd="delete From FS_ME_InfoContribution where isLock=0"
	User_Conn.execute(sql_cmd)
'退稿
elseif action="untread" then
	if not MF_Check_Pop_TF("NS033") then Err_Show
	sql_cmd="update FS_ME_InfoContribution Set untread=1  where ContID in("&FormatIntArr(values)&")"
	User_Conn.execute("Update FS_ME_Users set Integral=Integral-"&CintStr(contrAuditPoint)&" , FS_Money=FS_Money-"&CintStr(contrAuditMoney)&" where usernumber in (select usernumber from FS_ME_InfoContribution where ContID in ("&FormatIntArr(values)&"))")
	User_Conn.execute(sql_cmd)
'编辑审核
elseif action="editaudit" then
if not MF_Check_Pop_TF("NS030") then Err_Show
Dim ContID,NewsTitle,CurtTitle,Content,Keywords,Author,ClassID
ContID=CintStr(Request.Form("hid_contID"))
NewsTitle=NoSqlHack(Request.Form("txt_newsTitle"))
CurtTitle=NoSqlHack(Request.Form("txt_curtTitle"))
Content=NoSqlHack(Request.Form("txt_content"))
Keywords=NoSqlHack(Request.Form("txt_keywords"))
Author=NoSqlHack(Request.Form("hid_Author"))
ClassID=NoSqlHack(Request.Form("hid_ClassID"))
sql_news_cmd="select  ID,NewsID,NewsTitle,CurtTitle,Content,Keywords,Author,classID,isPicNews,NewsPicFile,SaveNewsPath,FileName,FileExtName,addtime,Templet,NewsProperty,Editor from FS_NS_News"
sql_cmd="select audittf,PicFile,UserNumber,MainID from FS_ME_InfoContribution where ContID="&ContID
news_Rs.open sql_news_cmd,Conn,1,3
info_Rs.open sql_cmd,User_Conn,1,3
news_Rs.addNew
'============================================================
HaveNewsIDTF = False
Do While Not HaveNewsIDTF
	Temp_NewsID_Str = FS_News.getRamCode(15)
	Set ChedkNewsIDObj = Conn.ExeCute("Select NewsID From FS_NS_News Where NewsID = '" & Temp_NewsID_Str & "'")
	If ChedkNewsIDObj.Eof Then
		TempNewsID = Temp_NewsID_Str
		HaveNewsIDTF = True
		Exit Do
	End IF
	ChedkNewsIDObj.Close : Set ChedkNewsIDObj = NOthing	
Loop
'===========================================================
news_Rs("NewsID")=TempNewsID
news_Rs("NewsTitle")=NewsTitle
news_Rs("CurtTitle")=CurtTitle
news_Rs("Keywords")=Keywords
news_Rs("Author")=Author
news_Rs("Editor")=session("Admin_Name")
news_Rs("NewsProperty")="0,1,1,1,0,0,0,0,1,0,0"
news_Rs("Templet")=NoSqlHack(request.Form("NewsTemplet"))
news_Rs("classID")=classid
If Fs_News.CopyFileTF = 1 Then
	news_Rs("Content")=GetFilePicPath(Content,"Con")
	if trim(info_Rs("PicFile"))<>"" then
		news_Rs("NewsPicFile")=GetFilePicPath(info_Rs("PicFile"),"only")
		news_Rs("isPicNews")=1
	ENd if
Else
	news_Rs("Content")=Content
	if trim(info_Rs("PicFile"))<>"" then
		news_Rs("NewsPicFile")=info_Rs("PicFile")
		news_Rs("isPicNews")=1
	ENd if
End If
news_Rs("Author")=FS_News.GetUserName(info_Rs("UserNumber"))
if trim(ClassID)="" then
	Response.Redirect("lib/error.asp?ErrCodes=<li>请选择栏目</li>")
	Response.End()
End if
news_Rs("classID")=ClassID
news_Rs("SaveNewsPath") = Fs_news.SaveNewsPath(Fs_news.fileDirRule)
news_Rs("FileName")=Fs_news.strFileNameRule(Fs_news.fileNameRule,0,0)
select case Fs_news.fileExtName
	case 1 fileExtName="htm"
	case 2 fileExtName="shtml"
	case 3 fileExtName="shtml"
	case 4 fileExtName="shtm"
	case else fileExtName="html"
End select
news_Rs("FileExtName")=fileExtName
news_Rs("AddTime")=Now()
news_Rs.update
if G_IS_SQL_DB = 0 then'是否是sqlserver
	newsID = news_Rs("ID")
Else
'select ident_current('FS_NS_News') 有待修正
	newsID = Conn.execute("select ident_current('FS_NS_News')")(0)
End if
str_FileName=news_Rs("FileName")
If Instr(str_FileName,"自动编号ID") Then
	str_FileName = Replace(str_FileName,"自动编号ID",newsID)
	Set TempRsObj=server.CreateObject(G_FS_RS)
	TempRsObj.open "select FileName From [Fs_NS_News] where NewsID='"&news_Rs("NewsID")&"' and ID="&CintStr(newsID)&"",Conn,1,3
	if not TempRsObj.eof Then
		TempRsObj("FileName") = str_FileName
		TempRsObj.update
	End If
	TempRsObj.Close
End IF
If Instr(str_FileName,"唯一NewsID") Then
	str_FileName = Replace(str_FileName,"唯一NewsID",news_Rs("NewsID"))
	Set TempRsObj_1=server.CreateObject(G_FS_RS)
	TempRsObj_1.open "select FileName From [Fs_NS_News] where NewsID='"&news_Rs("NewsID")&"'",Conn,1,3
	if not TempRsObj_1.eof Then
		TempRsObj_1("FileName") = str_FileName
		TempRsObj_1.update
	End If
	TempRsObj_1.Close
End IF
'User_Conn.execute("Update FS_ME_InfoContribution set NewsID='"&news_Rs("NewsID")&"' where ContID="&ContID)
User_Conn.execute("Update FS_ME_Users set Integral=Integral+"&(contrAuditPoint)&" , FS_Money=FS_Money+"&(contrAuditMoney)&",ConNumberNews=ConNumberNews+1 where usernumber='"&info_Rs("UserNumber")&"'")
news_Rs.close
info_Rs("audittf")=1
info_Rs.update
info_Rs.close

'删除投稿
elseif action="delete" then
	if not MF_Check_Pop_TF("NS032") then Err_Show
	Dim Temp_Rs,NewsID
	Set Temp_Rs=User_Conn.execute("Select NewsID From FS_ME_InfoContribution where ContID in("&FormatIntArr(values)&")")
	while not Temp_Rs.eof
		NewsID=Temp_Rs("NewsID")&","&NewsID
		Temp_Rs.movenext
	wend
	NewsID=DelHeadAndEndDot(NewsID)
	User_Conn.Execute("Delete *  From FS_ME_InfoContribution where ContID in("&FormatIntArr(values)&")")
	Conn.Execute("Delete *  From FS_NS_News where ID in("&FormatIntArr(NewsID)&")")
	Set Temp_Rs=nothing
End If

if err.number=0 then
	Response.Redirect("lib/success.asp?ErrCodes=<li>操作成功</li>&ErrorURL=../Constr_List.asp")
	Response.End()
Else
	Response.Redirect("lib/error.asp?ErrCodes=<li>请检查输入是否合法</li>")
	Response.End()
end if

'==========================================
'审核投稿时，复制内容中的图片到系统主目录
'==========================================
Function GetFilePicPath(ConStr,ConType)
	Dim Str_Con,Re,InnerPicAll,PicUrlStr,PicUrls,AllPicUrl
	Dim Num,Arr,NewPicStr,Str_PicPath,Ruslt_con
	Str_Con = ConStr & ""
	If Str_Con = "" Then Exit Function
	If ConType = "only" Then
		GetFilePicPath = CopyPicFiles(Str_Con)
	Else
		Set Re = New RegExp
		Re.IgnoreCase = True
		Re.Global = True
		Re.Pattern = "((src|href)\S+\.{1}(gif|jpg|png|swf|doc|rar)("">|""|\'|>| |\'>\s))"
		InnerPicAll = ""
		Set InnerPicAll = Re.Execute(Str_Con)
		Set Re = Nothing
		PicUrls = ""
		For Each PicUrlStr in InnerPicAll
			PicUrls = Replace(Replace(Replace(Replace(Replace(PicUrlStr,"src=",""),"'",""),"""",""),"href=",""),">","")
			AllPicUrl = AllPicUrl & "|||" & PicUrls
		Next
		If AllPicUrl <> "" Then
			If Left(AllPicUrl,3) = "|||" Then
				AllPicUrl = Right(AllPicUrl,Len(AllPicUrl) - 3)
			End If
		Else
			GetFilePicPath = Str_Con
			Exit Function
		End If
		If Instr(AllPicUrl,"|||") > 0 Then
			Arr = Split(AllPicUrl,"|||")
			For Num = LBound(Arr) To UBound(Arr)
				Str_PicPath = Arr(Num)
				NewPicStr = CopyPicFiles((Arr(Num)))
				Str_Con = Replace(Str_Con,Str_PicPath,NewPicStr)
			Next
			GetFilePicPath = Str_Con
		Else
			Str_PicPath = AllPicUrl
			NewPicStr = CopyPicFiles(AllPicUrl)
			GetFilePicPath = Replace(Str_Con,Str_PicPath,NewPicStr)
		End If			
	End If
End Function

Function CopyPicFiles(Pic)
	Dim FileTypeArr,i,FileName,FileTF,PicNewPath,FileAllName,PicPath
	Dim MFDemo,Vroot,PicUrlStart,PicTruePath,PicVPath
	Dim FsoObj,PicUrl
	If G_VIRTUAL_ROOT_DIR = "" Then
		Vroot = ""
	Else
		Vroot = "/" & G_VIRTUAL_ROOT_DIR
	End If
	Vroot = Replace(Vroot,"//","/")	
	MFDemo = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
	If Left(Lcase(MFDemo),7) = "http://" Then
		MFDemo = Right(MFDemo,Len(MFDemo) - 7)
	End If
	If Right(MFDemo,1) = "/" Then
		MFDemo = Left(MFDemo,Len(MFDemo) - 1)
	End If		 
	FileTypeArr = Array("jpg","gif","png","doc","swf","rar")
	PicUrl = LCase(Pic)
	If PicUrl = "" Or Instr(PicUrl,".") = 0 Or Instr(PicUrl,"/") = 0 Then
		Exit Function
	End If
	FileName = Lcase(Split(PicUrl,".")(UBound(Split(PicUrl,"."))))
	FileTF = False
	For i = LBound(FileTypeArr) To UBound(FileTypeArr)
		If FileName = FileTypeArr(i) Then
			FileTF = True
			Exit For
		End If
	Next
	If FileTF = True Then
		ON Error Resume NExt
		PicVPath = Replace(Replace(Replace(PicUrl,"http://",""),MFDemo,""),Vroot,"")
		PicVPath = Replace(PicVPath,"//","/")
		If Lcase(Left(PicVPath,Len("/" & G_USERFILES_DIR))) = LCase("/" & G_USERFILES_DIR) Then
			PicVPath = Right(PicVPath,Len(PicVPath) - Len("/" & G_USERFILES_DIR))
		End If
		If Left(PicVPath,1) = "/" Then
			PicVPath = Right(PicVPath,Len(PicVPath) - 1)
		End If
		FileAllName = Split(PicVPath,"/")(UBound(Split(PicVPath,"/")))
		PicUrlStart = "/" & Replace(PicVPath,FileAllName,"")
		PicNewPath = Server.MapPath(Replace("/" & G_UP_FILES_DIR & PicUrlStart,"//","/"))
		Set FsoObj = Server.CreateObject(G_FS_FSO)
		If FsoObj.FolderExists(PicNewPath) = False Then
			FsoObj.CreateFolder(PicNewPath)
		End If
		PicPath = Server.MapPath(Replace("/" & G_UP_FILES_DIR & PicUrlStart & FileAllName,"//","/"))
		PicTruePath = Server.MapPath(PicUrl)		
		If FsoObj.FileExists(PicPath) = False Then
			FsoObj.CopyFile PicTruePath,PicPath,True
		End If
		Set FsoObj = Nothing	
		If Err Then
			Err.Clear
			CopyPicFiles = PicUrl
		Else
			CopyPicFiles = Replace(Vroot & "/" & G_UP_FILES_DIR & PicUrlStart & FileAllName,"//","/")
		End If	 
	End If
End Function
%>
<%
Conn.close
User_Conn.close
Set Conn=nothing
Set User_Conn=nothing
Set info_Rs=nothing
Set Fs_News=nothing
%>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






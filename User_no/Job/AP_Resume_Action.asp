<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<% Option Explicit %>
<%Session.CodePage=65001%>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
Dim action,part,resumeRs,sqlstatement,tmpRs,errorMsg,id
'baseInfo
Dim base_BID,UserNumber,Uname,Sex,PictureExt,Birthday,CertificateClass,CertificateNo,CurrentWage,CurrencyType,WorkAge,Province,City,HomeTel,CompanyTel,Mobile,Email,QQ,isPublic,click,lastTime,address,ShenGao,XueLi,HowDay
'Intention
Dim WorkType,Salary,SelfAppraise
'workcity
Dim workprovince,workcity
'WorkExp
Dim BeginDate,EndDate,CompanyName,CompanyKind,Trade,Job,Department,workDescription,Certifier,CertifierTel
'EducateExp
Dim  edu_BeginDate,edu_EndDate,edu_SchoolName,edu_Specialty,edu_Diploma,edu_Description 
'TrainExp
Dim train_BeginDate,train_EndDate,train_TrainOrgan,train_TrainAdress,train_TrainContent,train_Certificate
'language
Dim Language,Degree
'Certificate
Dim FetchDate,Certificate,Score
'ProjectExp
Dim Pro_BeginDate,Pro_EndDate,Project,SoftSettings,HardSettings,Tools,ProjectDescript,Duty
'other
Dim title,content
'mail
Dim mailTitle,mailContent
'¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö
response.Charset="GB2312"
part = NoSqlHack(request("part"))
action = NoSqlHack(request("action"))
id = CintStr(request("id"))
Set resumeRs=server.CreateObject(G_FS_RS)
if part="baseinfo" then
	if id<>"" then 
		sqlstatement="select bid,UserNumber,Uname,Sex,PictureExt,Birthday,CertificateClass,CertificateNo,CurrentWage,CurrencyType,WorkAge,Province,City,HomeTel,CompanyTel,Mobile,Email,QQ,isPublic,click,lastTime,address,ShenGao,XueLi,HowDay from FS_AP_Resume_BaseInfo where bid="&CintStr(id)
	Else
		sqlstatement="select bid,UserNumber,Uname,Sex,PictureExt,Birthday,CertificateClass,CertificateNo,CurrentWage,CurrencyType,WorkAge,Province,City,HomeTel,CompanyTel,Mobile,Email,QQ,isPublic,click,lastTime,address,ShenGao,XueLi,HowDay from FS_AP_Resume_BaseInfo where UserNumber='"&session("FS_UserNumber")&"'"
	End if
	resumeRs.open sqlstatement,Conn,1,3
	if  resumeRs.eof then
		resumeRs.addnew
	End if
	UserNumber=session("FS_UserNumber")
	Uname=NoSqlHack(request.Form("txt_Uname"))
	Sex=NoSqlHack(request.Form("sel_sex"))
	PictureExt=NoSqlHack(request.Form("sel_PictureExt"))
	Birthday=NoSqlHack(request.Form("txt_Birthday"))
	CertificateClass=NoSqlHack(request.Form("sel_CertificateClass"))
	CertificateNo=NoSqlHack(request.Form("txt_CertificateNo"))
	CurrentWage=NoSqlHack(request.Form("sel_CurrentWage"))
	CurrencyType=NoSqlHack(request.Form("sel_CurrencyType"))
	WorkAge=NoSqlHack(request.Form("sel_WorkAge"))
	Province = NoSqlHack(request.Form("txt_Province"))
	City = NoSqlHack(request.Form("txt_City"))
	HomeTel = NoSqlHack(request.Form("txt_HomeTel"))
	CompanyTel = NoSqlHack(request.Form("txt_CompanyTel"))
	Mobile = NoSqlHack(request.Form("txt_Mobile"))
	Email = NoSqlHack(request.Form("txt_Email"))
	QQ = NoSqlHack(request.Form("txt_QQ"))
	isPublic = NoSqlHack(request.Form("sel_isPublic"))
	lastTime=Now
	
	address = NoSqlHack(request.Form("txt_address"))
	ShenGao = NoSqlHack(request.Form("txt_ShenGao"))
	XueLi = NoSqlHack(request.Form("txt_XueLi"))
	HowDay = NoSqlHack(request.Form("txt_HowDay"))
	'--------------------------------------
	resumeRs("UserNumber")=UserNumber
	if Trim(Uname)="" then 
		errorMsg="1*" 
	else 
		resumeRs("Uname")=Uname
	End if
	resumeRs("Sex")=Sex
	resumeRs("PictureExt")=PictureExt
	resumeRs("Birthday")=Birthday
	resumeRs("CertificateClass")=CertificateClass
	resumeRs("CertificateNo")=CertificateNo
	resumeRs("CurrentWage")=CurrentWage
	resumeRs("CurrencyType")=CurrencyType
	if Trim(WorkAge)="" then 
		errorMsg=errorMsg+"3*" 
	else 
		if not isNumeric(WorkAge) then 
			errorMsg=errorMsg+"3*" 
		else 
			resumeRs("WorkAge")=WorkAge
		End if
	End if
	resumeRs("WorkAge")=WorkAge
	resumeRs("Province")=Province
	resumeRs("City")=City
	resumeRs("HomeTel")=HomeTel
	resumeRs("CompanyTel")=CompanyTel
	resumeRs("Mobile")=Mobile
	resumeRs("Email")=Email
	resumeRs("QQ")=QQ
	resumeRs("isPublic")=isPublic
	resumeRs("lastTime")=lastTime

	resumeRs("address")=address
	if isnumeric(ShenGao) then resumeRs("ShenGao")=ShenGao
	resumeRs("XueLi")=XueLi
	resumeRs("HowDay")=HowDay
	if Trim(errorMsg)="" then 
		resumeRs.update 
		resumeRs.close
	Else
		Response.Write(errorMsg)
		Conn.close
		Set Conn=nothing
		response.End()
	End if
ElseIf part="intention" then '¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö
	if id<>"" then
		sqlstatement="select UserNumber,WorkType,Salary,SelfAppraise from FS_AP_Resume_Intention where bid="&CintStr(id)
	Else
		sqlstatement="select UserNumber,WorkType,Salary,SelfAppraise from FS_AP_Resume_Intention where UserNumber='"&session("FS_UserNumber")&"'"
	End if
	WorkType = NoSqlHack(request.Form("sel_WorkType"))
	Salary = NoSqlHack(request.Form("sel_Salary"))
	SelfAppraise = NoSqlHack(request.Form("txt_SelfAppraise"))
	resumeRs.open sqlstatement,Conn,1,3
	if 	resumeRs.eof then
		resumeRs.addNew
	End if
	resumeRs("UserNumber")=session("FS_UserNumber")
	resumeRs("WorkType")=WorkType
	resumeRs("Salary")=Salary
	resumeRs("SelfAppraise")=SelfAppraise
	resumeRs.update
	resumeRs.close
ElseIf part="workexp" then '¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö¡ö
	if id<>"" then
		sqlstatement="select bid,UserNumber,BeginDate,EndDate,CompanyName,CompanyKind,Trade,Job,Department,Description,Certifier,CertifierTel from FS_AP_Resume_WorkExp where bid="&CintStr(id)
	Else
		sqlstatement="select bid,UserNumber,BeginDate,EndDate,CompanyName,CompanyKind,Trade,Job,Department,Description,Certifier,CertifierTel from FS_AP_Resume_WorkExp where UserNumber='"&session("FS_UserNumber")&"'"
	End if
	UserNumber=session("FS_UserNumber")
	BeginDate = NoSqlHack(request.Form("txt_BeginDate"))
	EndDate = NoSqlHack(request.Form("txt_EndDate"))
	CompanyName = NoSqlHack(request.Form("txt_CompanyName"))
	CompanyKind = NoSqlHack(request.Form("sel_CompanyKind"))
	Trade = NoSqlHack(request.Form("txt_Trade"))
	Job = NoSqlHack(request.Form("txt_Job"))
	Department = NoSqlHack(request.Form("txt_Department"))
	workDescription = NoSqlHack(request.Form("txt_Description"))
	Certifier = NoSqlHack(request.Form("txt_Certifier"))
	CertifierTel = NoSqlHack(request.Form("txt_CertifierTel"))
	'----------------------------------------
	resumeRs.open sqlstatement,Conn,1,3
	if action<>"edit" or resumeRs.eof then
		resumeRs.addNew
	End if
	resumeRs("UserNumber")=UserNumber
	if trim(BeginDate)="" then 
		errorMsg=errorMsg&"4*"
	Else
		resumeRs("BeginDate")=BeginDate
	End If
	if trim(EndDate)="" then 
		errorMsg=errorMsg&"7*"
	Else
		resumeRs("EndDate")=EndDate
	End if
	if trim(CompanyName)="" then 
		errorMsg=errorMsg&"5*"
	Else
		resumeRs("CompanyName")=CompanyName
	End if
	resumeRs("CompanyKind")=CompanyKind
	resumeRs("Trade")=Trade
	if trim(Job)="" then 
		errorMsg=errorMsg&"6*"
	Else
		resumeRs("Job")=Job
	End if
	resumeRs("Department")=Department
	resumeRs("Description")=workDescription
	resumeRs("Certifier")=Certifier
	resumeRs("CertifierTel")=CertifierTel
	if Trim(errorMsg)="" then
		resumeRs.update
		resumeRs.close
	Else
		Response.Write(errorMsg)
		Conn.close
		Set Conn=nothing
		response.End()
	End if
Elseif part="educateexp" then
	if id<>"" then
		sqlstatement="select bid,UserNumber,BeginDate,EndDate,SchoolName,Specialty,Diploma,Description from FS_AP_Resume_EducateExp where bid="&CintStr(id)
	Else
		sqlstatement="select bid,UserNumber,BeginDate,EndDate,SchoolName,Specialty,Diploma,Description from FS_AP_Resume_EducateExp where UserNumber='"&session("FS_UserNumber")&"'"
	End if
	UserNumber=session("FS_UserNumber")
	edu_BeginDate = NoSqlHack(request.Form("txt_BeginDate"))
	edu_EndDate = NoSqlHack(request.Form("txt_EndDate"))
	edu_SchoolName = NoSqlHack(request.Form("txt_SchoolName"))
	edu_Specialty = NoSqlHack(request.Form("txt_Specialty"))
	edu_Diploma = NoSqlHack(request.Form("txt_Diploma"))
	edu_Description = NoSqlHack(request.Form("txt_Description"))
	if trim(edu_BeginDate)="" then errorMsg=errorMsg&"4*"
	if trim(edu_SchoolName)="" then errorMsg=errorMsg&"8*"
	if errorMsg<>"" then 
		Response.Write(errorMsg)
		Conn.close
		Set Conn=nothing
		response.End()
	End if
	resumeRs.open sqlstatement,conn,1,3
	if action<>"edit" or resumeRs.eof then
		resumeRs.addNew
	End if
	resumeRs("UserNumber")=UserNumber
	resumeRs("BeginDate")=edu_BeginDate
	resumeRs("EndDate")=edu_EndDate
	resumeRs("SchoolName")=edu_SchoolName
	resumeRs("Specialty")=edu_Specialty
	resumeRs("Diploma")=edu_Diploma	
	resumeRs("Description")=edu_Description
	resumeRs.update
	resumeRs.close
Elseif part="trainexp" then
	if id<>"" then
		sqlstatement="select bid,UserNumber,BeginDate,EndDate,TrainOrgan,TrainAdress,TrainContent,Certificate from FS_AP_Resume_TrainExp where bid="&CintStr(id)
	Else
		sqlstatement="select bid,UserNumber,BeginDate,EndDate,TrainOrgan,TrainAdress,TrainContent,Certificate from FS_AP_Resume_TrainExp where UserNumber='"&session("FS_UserNumber")&"'"
	End if
	UserNumber=session("FS_UserNumber")
	train_BeginDate = NoSqlHack(request.Form("txt_BeginDate"))
	train_EndDate = NoSqlHack(request.Form("txt_EndDate"))
	train_TrainOrgan = NoSqlHack(request.Form("txt_TrainOrgan"))
	train_TrainAdress = NoSqlHack(request.Form("txt_TrainAdress"))
	train_TrainContent = NoSqlHack(request.Form("txt_TrainContent"))
	train_Certificate = NoSqlHack(request.Form("txt_Certificate"))
	if trim(train_BeginDate)="" then errorMsg=errorMsg&"4*"
	if trim(train_TrainOrgan)="" then errorMsg=errorMsg&"9*"
	if errorMsg<>"" then 
		Response.Write(errorMsg)
		Conn.close
		Set Conn=nothing
		response.End()
	End if
	resumeRs.open sqlstatement,conn,1,3
	if action<>"edit" or resumeRs.eof then
		resumeRs.addNew
	End if
	resumeRs("UserNumber")=UserNumber
	resumeRs("BeginDate")=train_BeginDate
	resumeRs("EndDate")=train_EndDate
	resumeRs("TrainOrgan")=train_TrainOrgan
	resumeRs("TrainAdress")=train_TrainAdress
	resumeRs("TrainContent")=train_TrainContent	
	resumeRs("Certificate")=train_Certificate
	resumeRs.update
	resumeRs.close
Elseif part="language" then
	if id<>"" then
		sqlstatement="select bid,UserNumber,Language,Degree from FS_AP_Resume_Language where bid="&CintStr(id)
	Else
		sqlstatement="select bid,UserNumber,Language,Degree from FS_AP_Resume_Language where UserNumber='"&session("FS_UserNumber")&"'"
	End if
	UserNumber=session("FS_UserNumber")
	Language = NoSqlHack(request.Form("txt_Language"))
	Degree = NoSqlHack(request.Form("txt_Degree"))
	if trim(Language)="" then errorMsg=errorMsg&"10*"
	if trim(Degree)="" then errorMsg=errorMsg&"11*"
		if errorMsg<>"" then 
		Response.Write(errorMsg)
		Conn.close
		Set Conn=nothing
		response.End()
	End if
	resumeRs.open sqlstatement,conn,1,3
	if action<>"edit" or resumeRs.eof then
		resumeRs.addNew
	End if
	resumeRs("UserNumber")=UserNumber
	resumeRs("Language")=Language
	resumeRs("Degree")=Degree
	resumeRs.update
	resumeRs.close
Elseif part="certificate" then
	if id<>"" then
		sqlstatement="select UserNumber,FetchDate,Certificate,Score from FS_AP_Resume_Certificate where bid="&CintStr(id)
	Else
		sqlstatement="select UserNumber,FetchDate,Certificate,Score from FS_AP_Resume_Certificate where UserNumber='"&session("FS_UserNumber")&"'"
	End if
	UserNumber=session("FS_UserNumber")
	FetchDate = NoSqlHack(request.Form("txt_FetchDate"))
	Certificate = NoSqlHack(request.Form("txt_Certificate"))
	Score = NoSqlHack(request.Form("txt_Score"))

	if trim(FetchDate)="" then errorMsg=errorMsg&"12*"
	if trim(Certificate)="" then errorMsg=errorMsg&"13*"
	if trim(Score)="" then errorMsg=errorMsg&"11*"
	if errorMsg<>"" then 
		Response.Write(errorMsg)
		Conn.close
		Set Conn=nothing
		response.End()
	End if
	resumeRs.open sqlstatement,conn,1,3
	if action<>"edit" or resumeRs.eof then
		resumeRs.addNew
	End if
	resumeRs("UserNumber")=UserNumber
	resumeRs("FetchDate")=FetchDate
	resumeRs("Certificate")=Certificate
	resumeRs("Score")=Score
	resumeRs.update
	resumeRs.close
Elseif part="projectexp" then
	if id<>"" then
		sqlstatement="select UserNumber,BeginDate,EndDate,Project,SoftSettings,HardSettings,Tools,ProjectDescript,Duty from FS_AP_Resume_ProjectExp where bid="&CintStr(id)
	Else
		sqlstatement="select UserNumber,BeginDate,EndDate,Project,SoftSettings,HardSettings,Tools,ProjectDescript,Duty from FS_AP_Resume_ProjectExp where UserNumber='"&session("FS_UserNumber")&"'"
	End if
	UserNumber=session("FS_UserNumber")
	Pro_BeginDate = NoSqlHack(request.Form("txt_BeginDate"))
	Pro_EndDate = NoSqlHack(request.Form("txt_EndDate"))
	Project = NoSqlHack(request.Form("txt_Project"))
	SoftSettings = NoSqlHack(request.Form("txt_SoftSettings"))
	HardSettings = NoSqlHack(request.Form("txt_HardSettings"))
	Tools = NoSqlHack(request.Form("txt_Tools"))
	ProjectDescript = NoSqlHack(request.Form("txt_ProjectDescript"))
	Duty = NoSqlHack(request.Form("txt_Duty"))
	if trim(Pro_BeginDate)="" then errorMsg=errorMsg&"4*"
	if trim(Project)="" then errorMsg=errorMsg&"14*"
	if trim(ProjectDescript)="" then errorMsg=errorMsg&"15*"
	if trim(Duty)="" then errorMsg=errorMsg&"16*"
	if errorMsg<>"" then 
		Response.Write(errorMsg)
		Conn.close
		Set Conn=nothing
		response.End()
	End if
	resumeRs.open sqlstatement,conn,1,3
	if action<>"edit" or resumeRs.eof then
		resumeRs.addNew
	End if
	resumeRs("UserNumber")=UserNumber
	resumeRs("BeginDate")=Pro_BeginDate
	resumeRs("EndDate")=Pro_EndDate
	resumeRs("Project")=Project
	resumeRs("Project")=Project
	resumeRs("SoftSettings")=SoftSettings
	resumeRs("HardSettings")=HardSettings
	resumeRs("Tools")=Tools
	resumeRs("ProjectDescript")=ProjectDescript
	resumeRs("Duty")=Duty
	resumeRs.update
	resumeRs.close
Elseif part="other" then
	if id<>"" then
		sqlstatement="select UserNumber,Title,Content from FS_AP_Resume_Other where bid="&CintStr(id)
	Else
		sqlstatement="select UserNumber,Title,Content from FS_AP_Resume_Other where UserNumber='"&session("FS_UserNumber")&"'"
	End if
	title = NoSqlHack(request.Form("txt_Title"))
	content = NoSqlHack(request.Form("txt_Content"))
	if trim(title)="" then errorMsg=errorMsg&"17*"
	if trim(content)="" then errorMsg=errorMsg&"18*"
	if errorMsg<>"" then 
		Response.Write(errorMsg)
		Conn.close
		Set Conn=nothing
		response.End()
	End if
	resumeRs.open sqlstatement,conn,1,3
	if action<>"edit" or resumeRs.eof then
		resumeRs.addNew
	End if
	resumeRs("UserNumber")=session("FS_UserNumber")
	resumeRs("title")=title
	resumeRs("content")=content
	resumeRs.update
	resumeRs.close
Elseif part="mail" then
	if id<>"" then
		sqlstatement="select UserNumber,MailName,Content from FS_AP_Resume_Mail where bid="&id
	Else
		sqlstatement="select UserNumber,MailName,Content from FS_AP_Resume_Mail where UserNumber='"&session("FS_UserNumber")&"'"
	End if
	mailTitle = NoSqlHack(request.Form("txt_MailName"))
	mailContent = NoSqlHack(request.Form("txt_Content"))
	mailContent = replace(mailContent,vbcrlf,"<br />")
	if trim(mailTitle)="" then errorMsg=errorMsg&"17*"
	if trim(mailContent)="" then errorMsg=errorMsg&"18*"
	if errorMsg<>"" then 
		Response.Write(errorMsg)
		Conn.close
		Set Conn=nothing
		response.End()
	End if
	resumeRs.open sqlstatement,conn,1,3
	if action<>"edit" or resumeRs.eof then
		resumeRs.addNew
	End if
	resumeRs("UserNumber")=session("FS_UserNumber")
	resumeRs("MailName")=mailTitle
	resumeRs("Content")=mailContent
	resumeRs.update
	resumeRs.close
Elseif part="position" then
	sqlstatement="select UserNumber,Trade,Job from FS_AP_Resume_Position where 1=2"
	trade=NoSqlHack(request.Form("hid_trade"))
	job=NoSqlHack(request.Form("hid_job"))
	if trade="" or job="" then 
		errorMsg="19"
		response.Write(errorMsg)
		response.End()
	End if
	Dim tmp_rs
	Set tmp_rs=Conn.execute("Select BID From FS_AP_Resume_Position where trade='"&NoSqlHack(trade)&"' and job='"&NoSqlHack(job)&"' and usernumber='"&session("FS_UserNumber")&"'")
	if not tmp_rs.eof then
		errorMsg="20"
		response.Write(errorMsg)
		tmp_rs.close()
		Set tmp_rs=nothing
		response.End()
	end if
	resumeRs.open sqlstatement,conn,1,3
	resumeRs.addNew
	resumeRs("UserNumber")=session("FS_UserNumber")
	resumeRs("trade")=trade
	resumeRs("job")=job
	resumers.update
	resumeRs.close
Elseif part="workcity" then
	sqlstatement="select UserNumber,province,city from FS_AP_Resume_WorkCity where 1=2"
	workprovince=NoSqlHack(request.Form("hid_province"))
	workcity=NoSqlHack(request.Form("hid_city"))
	if workprovince="" or workcity="" then 
		errorMsg="21"
		response.Write(errorMsg)
		response.End()
	End if
	Set tmp_rs=Conn.execute("Select BID From FS_AP_Resume_WorkCity where province='"&NoSqlHack(workprovince)&"' and city='"&NoSqlHack(workcity)&"' and usernumber='"&session("FS_UserNumber")&"'")
	if not tmp_rs.eof then
		errorMsg="22"
		response.Write(errorMsg)
		tmp_rs.close()
		Set tmp_rs=nothing
		response.End()
	end if
	resumeRs.open sqlstatement,conn,1,3
	resumeRs.addNew
	resumeRs("UserNumber")=session("FS_UserNumber")
	resumeRs("province")=workprovince
	resumeRs("city")=workcity
	resumers.update
	resumeRs.close
Elseif action="del" then
	Dim deltarget
	deltarget = NoSqlHack(request("delpart"))
	if id="" then response.End()
	select case deltarget
		case "baseinfo" conn.execute("delete from FS_AP_Resume_BaseInfo where bid="&CintStr(id))
		case "intention" conn.execute("delete from FS_AP_Resume_Intention where bid="&CintStr(id))
		case "workexp" conn.execute("delete from FS_AP_Resume_WorkExp where bid="&CintStr(id))
		case "educateexp" conn.execute("delete from FS_AP_Resume_EducateExp where bid="&CintStr(id))
		case "trainexp" conn.execute("delete from FS_AP_Resume_TrainExp where bid="&CintStr(id))
		case "language" conn.execute("delete from FS_AP_Resume_Language where bid="&CintStr(id))
		case "certificate" conn.execute("delete from FS_AP_Resume_Certificate where bid="&CintStr(id))
		case "projectexp" conn.execute("delete from FS_AP_Resume_ProjectExp where bid="&CintStr(id))
		case "other" conn.execute("delete from FS_AP_Resume_Other where bid="&CintStr(id))
		case "mail" conn.execute("delete from FS_AP_Resume_Mail where bid="&CintStr(id))
		case "position" conn.execute("delete from FS_AP_Resume_position where bid="&CintStr(id))
		case "workcity" conn.execute("delete from FS_AP_Resume_Workcity where bid="&CintStr(id))
	End select
End if
Response.write("ok")
Conn.close
Set Conn=nothing
Set tmpRs=nothing
%>






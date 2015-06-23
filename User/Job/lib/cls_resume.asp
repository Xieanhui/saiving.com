<%'
Class cls_resume
	'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	'baseInfo
	Private base_BID,base_UserNumber,base_Uname,base_Sex,base_PictureExt,base_Birthday,base_CertificateClass,base_CertificateNo,base_CurrentWage,base_CurrencyType,base_WorkAge,base_Province,base_City,base_HomeTel,base_CompanyTel,base_Mobile,base_Email,base_QQ,base_isPublic,base_click,base_lastTime,base_address,base_ShenGao,base_XueLi,base_HowDay
	'Intention
	Private Itn_WorkType,Itn_Salary,Itn_SelfAppraise
	'Position
	Private pos_trade,pos_job
	'Position
	Private w_province,w_city
	'WorkExp
	Private wep_BeginDate,wep_EndDate,wep_CompanyName,wep_CompanyKind,wep_Trade,wep_Job,wep_Department,wep_workDescription,wep_Certifier,wep_CertifierTel
	'EducateExp
	Private  edu_BeginDate,edu_EndDate,edu_SchoolName,edu_Specialty,edu_Diploma,edu_Description 
	'TrainExp
	Private train_BeginDate,train_EndDate,train_TrainOrgan,train_TrainAdress,train_TrainContent,train_Certificate
	'language
	Private lng_Language,lng_Degree
	'Certificate
	Private cer_FetchDate,cer_Certificate,cer_Score
	'ProjectExp
	Private Pro_BeginDate,Pro_EndDate,Pro_Project,Pro_SoftSettings,Pro_HardSettings,Pro_Tools,Pro_ProjectDescript,Pro_Duty
	'other
	Private o_title,o_content
	'mail
	Private mailTitle,mailContent
	'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	Public function getResumeInfo(part,id)
		Dim sqlstatement,resumeRs
		Set resumeRs=Server.CreateObject(G_FS_RS)
		select case NoSqlHack(part)
			case "baseinfo"    sqlstatement="select BID,UserNumber,Uname,Sex,PictureExt,Birthday,CertificateClass,CertificateNo,CurrentWage,CurrencyType,WorkAge,Province,City,HomeTel,CompanyTel,Mobile,Email,QQ,isPublic,click,lastTime,address,ShenGao,XueLi,HowDay from FS_AP_Resume_BaseInfo where bid="&CintStr(id)
			case "intention"   sqlstatement="select BID,UserNumber,WorkType,Salary,SelfAppraise from FS_AP_Resume_Intention where bid="&id
			case "position"    sqlstatement="select BID,UserNumber,trade,job from FS_AP_Resume_Position where bid="&id
			case "workcity"    sqlstatement="select BID,UserNumber,Province,City from FS_AP_Resume_WorkCity where bid="&id
			case "workexp"     sqlstatement="select BID,UserNumber,BeginDate,EndDate,CompanyName,CompanyKind,Trade,Job,Department,Description,Certifier,CertifierTel from FS_AP_Resume_WorkExp where bid="&CintStr(id)
			case "educateexp"  sqlstatement="select BID,UserNumber,BeginDate,EndDate,SchoolName,Specialty,Diploma,Description from FS_AP_Resume_EducateExp where bid="&CintStr(id)
			case "trainexp"    sqlstatement="select BID,UserNumber,BeginDate,EndDate,TrainOrgan,TrainAdress,TrainContent,Certificate from FS_AP_Resume_TrainExp where bid="&CintStr(id)
			case "language"    sqlstatement="select BID,UserNumber,Language,Degree from FS_AP_Resume_Language where bid="&CintStr(id)
			case "certificate" sqlstatement="select BID,UserNumber,FetchDate,Certificate,Score from FS_AP_Resume_Certificate where bid="&CintStr(id)
			case "projectexp"  sqlstatement="select BID,UserNumber,BeginDate,EndDate,Project,SoftSettings,HardSettings,Tools,ProjectDescript,Duty from FS_AP_Resume_ProjectExp where bid="&CintStr(id)
			case "other"       sqlstatement="select BID,UserNumber,Title,Content from FS_AP_Resume_Other where bid="&CintStr(id)
			case "mail"        sqlstatement="select BID,UserNumber,MailName,Content from FS_AP_Resume_Mail where bid="&CintStr(id)
		End select
		resumeRs.open sqlstatement,Conn,1,3
		if resumeRs.eof then exit function
		if part="baseinfo" then
			 base_BID=resumeRs("BID")
			 base_UserNumber=resumeRs("UserNumber")
			 base_Uname=resumeRs("Uname")
			 base_Sex=resumeRs("Sex")
			 base_PictureExt=resumeRs("PictureExt")
			 base_Birthday=resumeRs("Birthday")
			 base_CertificateClass=resumeRs("CertificateClass")
			 base_CertificateNo=resumeRs("CertificateNo")
			 base_CurrentWage=resumeRs("CurrentWage")
			 base_CurrencyType=resumeRs("CurrencyType")
			 base_WorkAge=resumeRs("WorkAge")
			 base_Province=resumeRs("Province")
			 base_City=resumeRs("City")
			 base_HomeTel=resumeRs("HomeTel")
			 base_CompanyTel=resumeRs("CompanyTel")
			 base_Mobile=resumeRs("CompanyTel")
			 base_Email=resumeRs("Email")
			 base_QQ=resumeRs("QQ")
			 base_isPublic=resumeRs("isPublic")
			 base_click=resumeRs("click")
			 base_lastTime=resumeRs("lastTime")
			 
			 base_address=resumeRs("address")
			 base_ShenGao=resumeRs("ShenGao")
			 base_XueLi=resumeRs("XueLi")
			 base_HowDay=resumeRs("HowDay")
		Elseif part="intention" then
			Itn_WorkType=resumeRs("WorkType")
			Itn_Salary=resumeRs("Salary")
			Itn_SelfAppraise=resumeRs("SelfAppraise")
		Elseif part="position" then
			pos_trade=resumeRs("trade")
			pos_job=resumeRs("job")
		Elseif part="workexp" then
			wep_BeginDate=resumeRs("BeginDate")
			wep_EndDate=resumeRs("EndDate")
			wep_CompanyName=resumeRs("CompanyName")
			wep_CompanyKind=resumeRs("CompanyKind")
			wep_Trade=resumeRs("Trade")
			wep_Job=resumeRs("job")
			wep_Department=resumeRs("Department")
			wep_workDescription=resumeRs("Description")
			wep_Certifier=resumeRs("Certifier")
			wep_CertifierTel=resumeRs("CertifierTel")
		Elseif part="educateexp" then
			edu_BeginDate=resumeRs("BeginDate")
			edu_EndDate=resumeRs("EndDate")
			edu_SchoolName=resumeRs("SchoolName")
			edu_Specialty=resumeRs("Specialty")
			edu_Diploma=resumeRs("Diploma")
			edu_Description=resumeRs("Description")
		Elseif part="trainexp" then
			train_BeginDate=resumeRs("BeginDate")
			train_EndDate=resumeRs("EndDate")
			train_TrainOrgan=resumeRs("TrainOrgan")
			train_TrainAdress=resumeRs("TrainAdress")
			train_TrainContent=resumeRs("TrainContent")
			train_Certificate=resumeRs("Certificate")
		Elseif part="language" then
			lng_Language=resumeRs("Language")
			lng_Degree=resumeRs("Degree")
		Elseif part="certificate" then
			cer_FetchDate=resumeRs("FetchDate")
			cer_Certificate=resumeRs("Certificate")
			cer_Score=resumeRs("Score")
		Elseif part="projectexp" Then
			Pro_BeginDate=resumeRs("BeginDate")
			Pro_EndDate=resumeRs("EndDate")
			Pro_Project=resumeRs("Project")
			Pro_SoftSettings=resumeRs("SoftSettings")
			Pro_HardSettings=resumeRs("HardSettings")
			Pro_Tools=resumeRs("Tools")
			Pro_ProjectDescript=resumeRs("ProjectDescript")
			Pro_Duty=resumeRs("Duty")
		Elseif part="other" Then
			o_title=resumeRs("title")
			o_content=resumeRs("content")
		Elseif part="mail" Then
			mailTitle=resumeRs("MailName")
			mailContent=resumeRs("Content")
		End if
	End function
	'BaseInfo■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	public property get bs_bid'基本信息ID
		bs_bid=base_BID
	End property
	
	public property get bs_usernumber'用户编号
		bs_usernumber=base_UserNumber
	End property

	public property get bs_Uname'用户名
		bs_Uname=base_Uname
	End property
	
	public property get bs_sex
		bs_sex=base_Sex
	End property
	
	public property get bs_PictureExt'图片后缀
		bs_PictureExt=base_PictureExt
	End property

	public property get bs_Birthday'生日
		bs_Birthday=base_Birthday
	End property

	public property get bs_CertificateClass'证件类型
		bs_CertificateClass=base_CertificateClass
	End property

	public property get bs_CertificateNo'证件号码
		bs_CertificateNo=base_CertificateNo
	End property
	
	public property get bs_CurrentWage'目前年薪
		bs_CurrentWage=base_CurrentWage
	End property
	
	public property get bs_CurrencyType'货币种类
		bs_CurrencyType=base_CurrencyType
	End property
	
	public property get bs_WorkAge'工作年限
		bs_WorkAge=base_WorkAge
	End property
	
	public property get bs_Province'所在省
		bs_Province=base_Province
	End property
	
	public property get bs_City'所在城市
		bs_City=base_City
	End property
	
	public property get bs_HomeTel'家庭电话
		bs_HomeTel=base_HomeTel
	End property

	public property get bs_CompanyTel'公司电话
		bs_CompanyTel=base_CompanyTel
	End property
	
	public property get bs_Mobile'移动电话
		bs_Mobile=base_Mobile
	End property

	public property get bs_Email'电子邮件
		bs_Email=base_Email
	End property

	public property get bs_QQ'QQ号码
		bs_QQ=base_QQ
	End property

	public property get bs_address'基本信息ID
		bs_address=base_address
	End property
	public property get bs_ShenGao'基本信息ID
		bs_ShenGao=base_ShenGao
	End property
	public property get bs_XueLi'基本信息ID
		bs_XueLi=base_XueLi
	End property
	public property get bs_HowDay'基本信息ID
		bs_HowDay=base_HowDay
	End property

	
	public property get bs_isPublic'是否公开
		bs_isPublic=base_isPublic
	End property

	public property get bs_click'被浏览数量
		bs_click=base_click
	End property

	public property get bs_lastTime'最后修改时间
		bs_lastTime=base_lastTime
	End property
	'intention■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	'Itn_WorkType,Itn_Salary,Itn_SelfAppraise
	public property get WorkTypee'工作类型
		WorkTypee=Itn_WorkType
	End property
	
	public property get Salary'期望工资
		Salary=Itn_Salary
	End property

	public property get SelfAppraise'自我介绍
		SelfAppraise=Itn_SelfAppraise
	End property
	'position■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	public property get p_trade'行业
		p_trade=pos_trade
	End property
	
	public property get p_job'岗位
		p_job=pos_job
	End property
	'workcity■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	public property get workprovince'工作地点（省）
		workprovince=w_province
	End property
	
	public property get workcity'工作地点（市）
		workcity=w_city
	End property
	
	'WorkExp■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	'wep_BeginDate,wep_EndDate,wep_CompanyName,wep_CompanyKind,wep_Trade,wep_Job,wep_Department,wep_workDescription,wep_Certifier,wep_CertifierTel
	public property get wBeginDate'开始时间
		wBeginDate=wep_BeginDate
	End property
	
	public property get wEndDate'结束时间
		wEndDate=wep_EndDate
	End property


	public property get CompanyName'公司名称
		CompanyName=wep_CompanyName
	End property
	
	public property get CompanyKind'公司性质
		CompanyKind=wep_CompanyKind
	End property

	public property get Trade'行业
		Trade=wep_Trade
	End property

	public property get Job'职业
		Job=wep_Job
	End property

	public property get Department'部门
		Department=wep_Department
	End property

	public property get workDescription'工作描述
		workDescription=wep_workDescription
	End property

	public property get Certifier'证明人
		Certifier=wep_Certifier
	End property

	public property get CertifierTel'证明人联系方式
		CertifierTel=wep_CertifierTel
	End property
	'EducateExp■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	'edu_BeginDate,edu_EndDate,edu_SchoolName,edu_Specialty,edu_Diploma,edu_Description 
	public property get eBeginDate'开始时间
		eBeginDate=edu_BeginDate
	End property

	public property get eEndDate'结束时间
		eEndDate=edu_EndDate
	End property

	public property get SchoolName'学校名称
		SchoolName=edu_SchoolName
	End property

	public property get Specialty'专业
		Specialty=edu_Specialty
	End property

	public property get Diploma'学历
		Diploma=edu_Diploma
	End property

	public property get eDescription'专业描述
		eDescription=edu_Description
	End property
	'TrainExp■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	'train_BeginDate,train_EndDate,train_TrainOrgan,train_TrainAdress,train_TrainContent,train_Certificate
	public property get tBeginDate'
		tBeginDate=train_BeginDate
	End property

	public property get tEndDate'
		tEndDate=train_EndDate
	End property

	public property get TrainOrgan'培训机构
		TrainOrgan=train_TrainOrgan
	End property

	public property get TrainAdress'机构地址
		TrainAdress=train_TrainAdress
	End property

	public property get TrainContent'培训内容
		TrainContent=train_TrainContent
	End property

	public property get tCertificate'证书
		tCertificate=train_Certificate
	End property	
	
	'language■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	' Language,Degree
	public property get Language'语言
		Language=lng_Language
	End property

	public property get Degree'等级
		Degree=lng_Degree
	End property
	'Certificat■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	'cer_FetchDate,cer_Certificate,cer_Score

	public property get FetchDate'获证时间
		FetchDate=cer_FetchDate
	End property

	public property get Certificate'证书名
		Certificate=cer_Certificate
	End property

	public property get Score'分数
		Score=cer_Score
	End property
	'Project■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	'Pro_BeginDate,Pro_EndDate,Pro_Project,Pro_SoftSettings,Pro_HardSettings,Pro_Tools,Pro_ProjectDescript,Pro_Duty
	public property get pBeginDate'开始时间
		pBeginDate=Pro_BeginDate
	End property

	public property get pEndDate'结束时间
		pEndDate=Pro_EndDate
	End property

	public property get Project'项目名称
		Project=Pro_Project
	End property
	
	public property get SoftSettings'软件环境
		SoftSettings=Pro_SoftSettings
	End property

	public property get HardSettings'硬件环境
		HardSettings=Pro_HardSettings
	End property
	
	public property get Tools'开发工具
		Tools=Pro_Tools
	End property
	
	public property get ProjectDescript'项目描述
		ProjectDescript=Pro_ProjectDescript
	End property
	
	public property get Duty'职责
		Duty=Pro_Duty
	End property
	'Other■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	'title,content
	public property get title'主题
		title=o_title
	End property
	
	public property get content'内容
		content=o_content
	End property
	'email■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	'mailTitle,mailContent
	public property get mtitle'主题
		mtitle=mailTitle
	End property
	
	public property get mcontent'内容
		mcontent=mailContent
	End property
	'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
End Class
%>






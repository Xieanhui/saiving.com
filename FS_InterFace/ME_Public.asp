<%
Class cls_ME
	'风讯新闻刷新代码获得类，用于获得代码
	'方法：Fs_ns.Get_LableChar("标签代码","")
	'Copyright (c)2002-2008 Foosun.cn
	'非经风讯公司正版许可，请勿用于商业用途
	'官方站：Foosun.cn
	'技术支持论坛：bbs.foosun.net
	Public Function get_LableChar(f_Lable,f_UserId_t,f_LableType)
	dim f_UserId,f_UserId_1
	if instr(1,f_UserId_t,"***",1)>0 then
		f_UserId = split(f_UserId_t,"***")(0)
		f_UserId_1 = split(f_UserId_t,"***")(1)
	else
		f_UserId = f_UserId_t
		f_UserId_1 = ""
	end if
	if Request.Cookies("FoosunMECookies")("FoosunMELogDir") = "" then
		dim get_sysparam,blogs_dir
		set get_sysparam = User_Conn.execute("select top 1 Dir From FS_ME_iLogSysParam")
		if get_sysparam.eof then
			Response.Cookies("FoosunMECookies")("FoosunMELogDir") = "blog"
			get_sysparam.close:set get_sysparam = nothing
		else
			Response.Cookies("FoosunMECookies")("FoosunMELogDir") = get_sysparam(0)
			get_sysparam.close:set get_sysparam = nothing
		end if
		Response.Cookies("FoosunMECookies").Expires=Date()+1
	end if
	select case LCase(f_Lable.LableFun)
		case "logpage"
			get_LableChar=Log_Page(f_Lable,f_UserId)
		case "loguserpage"
			get_LableChar=log_usertile(f_Lable,f_UserId)
		case "lastlog"
			get_LableChar=LogList(f_Lable,"lastlog",f_UserId)
		case "toplog"
			get_LableChar=LogList(f_Lable,"toplog",f_UserId)
		case "hotlog"
			get_LableChar=LogList(f_Lable,"hotlog",f_UserId)
		case "classlist"
			get_LableChar=ClassList(f_Lable)
		case "topsubject"
			get_LableChar=TopSubject(f_Lable,f_UserId)
		case "infoclass"
			get_LableChar=InfoClass(f_Lable)
		case "photo_list"
			get_LableChar=Photo_List(f_Lable,"photo_list")
		case "showphoto"
			get_LableChar=Photo_show(f_Lable,f_UserId)
		case "infolist"
			get_LableChar=InfoList(f_Lable,"infolist",f_UserId,f_UserId_1)
		case "log_lastreview"
			get_LableChar=Log_LastReview(f_Lable,"log_lastreview",f_UserId)
		case "log_lastform"
			get_LableChar=Log_LastForm(f_Lable,"log_lastform",f_UserId)
		case "log_publiclog"
			get_LableChar=Log_PublicLog(f_Lable)
		case "log_search"
			get_LableChar=Log_Search(f_Lable,f_UserId)
		case "log_navi"
			get_LableChar=Log_Navi(f_Lable)
		case "log_pagetitle"
			get_LableChar=Log_PageTitle(f_Lable,f_UserId)
		case "log_title"
			get_LableChar=Log_title(f_Lable,f_UserId)
		case "userlist"
			get_LableChar=UserList(f_Lable)
	end select
	End Function
	Public Function List_Addtional_HTML(f_type,f_tf,f_divid,f_divclass,f_ulid,f_ulclass,f_liid,f_liclass)
		if f_tf = 1 And f_divid = "" And f_divclass = "" And f_ulid = "" And f_ulclass = "" And f_liid = "" And f_liclass = "" then
			List_Addtional_HTML = ""
			Exit Function
		end if
		Select Case f_type
			Case "div","ul"
				List_Addtional_HTML = table_str_list_head(f_tf,f_divid,f_divclass,f_ulid,f_ulclass)
			Case "li"
				List_Addtional_HTML = table_str_list_middle_1(f_tf,f_liid,f_liclass)
			Case "div1","ul1"
				List_Addtional_HTML = table_str_list_bottom(f_tf)
			Case "li1"
				List_Addtional_HTML = table_str_list_middle_2(f_tf)
			Case Else
				List_Addtional_HTML = ""
		End Select
	End Function
	'DIV输出_____________________________________________________________________
	'得到div,table_______________________________________________________________
	Public Function table_str_list_head(f_tf,f_divid,f_divclass,f_ulid,f_ulclass)   
		Dim table_,tr_
		Dim f_divid_1,f_divclass_1,f_ulid_1,f_ulclass_1
		if f_tf=1 then
			if f_divid<>"" then:f_divid_1 = " id="""& f_divid &"""":else:f_divid_1 = "":end if
			if f_divclass<>"" then:f_divclass_1 = " class="""& f_divclass &"""":else:f_divclass_1 = "":end if
			if f_ulid<>"" then:f_ulid_1 = " id="""& f_ulid &"""":else:f_ulid_1 = "":end if
			if f_ulclass<>"" then:f_ulclass_1 = " class="""& f_ulclass &"""":else:f_ulclass_1 = "":end if
			table_="<div"& f_divid_1 & f_divclass_1 &">"
			tr_="<ul"& f_ulid_1 & f_ulclass_1 &">"
			table_str_list_head =  table_&chr(10)
			table_str_list_head = table_str_list_head &" "& tr_&chr(10)
		else
			table_="<table border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"">"
			table_str_list_head =  table_&chr(10)
		end if
	End Function
	
	'得到div,table_________________________________________________________________________________
	public Function table_str_list_middle_1(f_tf,f_liid,f_liclass)
		Dim f_liid_1,f_liclass_1,td_
		if f_tf=1 then
			if f_liid<>"" then:f_liid_1 = " id="""& f_liid &"""":else:f_liid_1 = "":end if
			if f_liclass<>"" then:f_liclass_1 = " class="""& f_liclass &"""":else:f_liclass_1 = "":end if
			td_="<li"&f_liid_1&f_liclass_1&">"
			table_str_list_middle_1 ="  "&td_
		end if
	End Function
	
	'得到div,table_________________________________________________________________________________
	public Function table_str_list_middle_2(f_tf)
		Dim td__,tr__
		if f_tf=1 then
			td__="</li>"
		else
			td__="</td>"
		end if
			table_str_list_middle_2 =  td__&chr(10)
	End Function
	
	'得到div,table_________________________________________________________________________________
	Public Function table_str_list_middle_3(f_tf)
		if f_tf=1 then
			table_str_list_middle_3 = ""
		else
			table_str_list_middle_3 = "</tr>"
		end if
	End Function
	
	
	'得到div,table_________________________________________________________________________________
	Public Function table_str_list_bottom(f_tf)
		Dim table__,tr__
		if f_tf=1 then
			table__="</div>"
			tr__="</ul>"
			table_str_list_bottom = " "&tr__&chr(10)
		else
			table__="</table>"
			table_str_list_bottom = ""
		end if
		table_str_list_bottom = table_str_list_bottom &table__&chr(10)
	End Function
	'DIV输出结束_________________________________________________________________
	'最新日志____________________________________________________________________
	Public Function LogList(f_Lable,f_type,f_UserId)
		dim TitleNumber,TitleNumberChar,leftTitleNumber,datetf,dateType,TitleCss,div_tf,TitleCssstr,classNews_head
		dim sql_char,f_sql,rs,inSQL_Search,LastLogstr,DateChar,classNews_bottom,classNews_middle1,classNews_middle2
		TitleNumber = f_Lable.LablePara("调用数量,标题字数")
		TitleNumberChar = split(TitleNumber,",")(0)
		leftTitleNumber = split(TitleNumber,",")(1)
		if trim(leftTitleNumber)<>"" and isnumeric(leftTitleNumber) then
			leftTitleNumber = leftTitleNumber
		else
			leftTitleNumber = 40
		end if
		if trim(TitleNumberChar)<>"" and isnumeric(TitleNumberChar) then
			TitleNumberChar = TitleNumberChar
		else
			TitleNumberChar = 10
		end if
		
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
		f_DivID = f_Lable.LablePara("DivID")
		f_DivClass = f_Lable.LablePara("Divclass")
		f_UlID = f_Lable.LablePara("ulid")
		f_ULClass = f_Lable.LablePara("ulclass")
		f_LiID = f_Lable.LablePara("liid")
		f_LiClass = f_Lable.LablePara("liclass")
		datetf= f_Lable.LablePara("显示日期")
		dateType = f_Lable.LablePara("日期样式")
		if f_Lable.LablePara("输出格式") = "out_DIV" then
			div_tf = 1
			TitleCssstr = ""
			classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		else
			div_tf=0
			TitleCss = f_Lable.LablePara("CSS")
			if trim(TitleCss)<>"" then
				TitleCssstr = " class="""& TitleCss &""""
			else
				TitleCssstr = ""
			end if
		end if
		if len(f_UserId)>6 then
			inSQL_Search = " and UserNumber='"& f_UserId &"'"
		else
			inSQL_Search = ""
		end if
		classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		if f_type="lastlog" then
			sql_char = " order by iLogID desc"
		elseif f_type="toplog" then
			sql_char = " order by hits desc,iLogID Desc"
		elseif f_type="hotlog" then
			sql_char = " order by istop desc,hits desc,iLogID Desc"
		end if
		f_sql="select top "& TitleNumberChar &" iLogID,iLogStyle,Title,UserNumber,addtime from FS_ME_Infoilog where islock=0 and isDraft=0 and adminLock=0 "& inSQL_Search &" "&sql_char&""
		set rs = User_Conn.execute(f_sql)
		if rs.eof then
			LastLogstr = ""
			rs.close:Set rs =nothing
		else
			do while not rs.eof
				if datetf="1" then
					DateChar = get_DateChar(dateType,rs("AddTime"))
				else
					DateChar = ""
				end if
				if div_tf=0 then
					dim s_paths
					s_paths = replace("/"&G_VIRTUAL_ROOT_DIR&"sys_images/log/dot1.jpg","//","/")
					LastLogstr = LastLogstr & " <img src="""& s_paths &""" border=""0""><a href="""& Get_LogLink(rs("iLogID"))&""""&TitleCssstr&">"&GotTopic(rs("Title"),leftTitleNumber)&"</a>"&DateChar&"<br />"&chr(10)
				else
					LastLogstr = LastLogstr & classNews_middle1& "<a href="""& Get_LogLink(rs("iLogID"))&""">"&GotTopic(rs("Title"),leftTitleNumber)&"</a>"&DateChar&""&classNews_middle2
				end if
				rs.movenext
			loop
			rs.close:Set rs =nothing
		end if
		if div_tf=1 then
			LastLogstr = classNews_head & LastLogstr & classNews_bottom
		end if
		f_Lable.DictLableContent.Add "1",LastLogstr
	End Function
	'得到分类日志列表
	Public Function ClassList(f_Lable)
		dim ClassId,TitleNumber,lefttile,CSS,div_tf,TitleCssstr,datetf,dateType,classNews_head
		dim f_sql,rs,DateChar,classNews_middle1,classNews_bottom,classNews_middle2
		ClassID = f_Lable.LablePara("ClassId")
		if trim(ClassID)<>"" then
			TitleNumber = f_Lable.LablePara("调用数量")
			if trim(TitleNumber)<>"" and isnumeric(TitleNumber) then
				TitleNumber = TitleNumber
			else
			   TitleNumber = 10
			end if
			lefttile = f_Lable.LablePara("标题字数")
			if trim(lefttile)<>"" and isnumeric(lefttile) then
				lefttile = lefttile
			else
				lefttile = 40
			end if
			CSS = f_Lable.LablePara("标题CSS")
			Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
			f_DivID = f_Lable.LablePara("DivID")
			f_DivClass = f_Lable.LablePara("Divclass")
			f_UlID = f_Lable.LablePara("ulid")
			f_ULClass = f_Lable.LablePara("ulclass")
			f_LiID = f_Lable.LablePara("liid")
			f_LiClass = f_Lable.LablePara("liclass")
			datetf= f_Lable.LablePara("显示日期")
			dateType = f_Lable.LablePara("日期格式")
			if f_Lable.LablePara("输出格式") = "out_DIV" then
				div_tf = 1
				TitleCssstr = ""
				classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
				classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			else
				div_tf=0
				if trim(CSS)<>"" then
					TitleCssstr = " class="""& CSS &""""
				else
					TitleCssstr = ""
				end if
			end if
			classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			f_sql="select top "& TitleNumber &" iLogID,MainID,iLogStyle,Title,UserNumber,addtime from FS_ME_Infoilog where islock=0 and isDraft=0 and adminLock=0 and MainID="&ClassId&" order by iLogID desc"
			set rs = User_Conn.execute(f_Sql)
			ClassList = ""
			if rs.eof then
				ClassList = ""
				rs.close:set rs = nothing
			else
				do while Not rs.eof
					if datetf="1" then
						DateChar = get_DateChar(dateType,rs("AddTime"))
					else
						DateChar = ""
					end if
					if div_tf=0 then
						dim s_paths
						s_paths = replace("/"&G_VIRTUAL_ROOT_DIR&"sys_images/log/dot1.jpg","//","/")
						ClassList = ClassList & " <img src="""& s_paths &""" border=""0""><a href="""& Get_LogLink(rs("iLogID"))&""""&TitleCssstr&">"&GotTopic(rs("Title"),lefttile)&"</a><a href="""&get_AuthorLink(rs("UserNumber"))&""" target=""_blank"">("&get_UserNameLink(rs("UserNumber"))&")</a>"&DateChar&"<br />"&chr(10)
					else
						ClassList = ClassList & classNews_middle1& "<a href=""http://"& Request.Cookies("FoosunMFCookies")("FoosunMFDomain") &Get_LogLink(rs("iLogID"))&""">"&GotTopic(rs("Title"),lefttile)&"</a><a href="""&get_AuthorLink(rs("UserNumber"))&""" target=""_blank"">("&get_UserNameLink(rs("UserNumber"))&")</a>"&DateChar&""&classNews_middle2
					end if
					rs.movenext
				loop
				rs.close:set rs = nothing
			end if
			if div_tf=1 then
				ClassList = classNews_head & ClassList & classNews_bottom
			end if
			ClassList = ClassList
		else
			ClassList = "错误的标签"
		end if
		f_Lable.DictLableContent.Add "1",ClassList
	end Function
	''=======================================会员标签
	Public Function UserList(f_Lable)
		Dim UserType,OrderBy,UserSex,TitleNumber,ColsNumber,leftTitle,DateType,out_char,ContentNumber,NewsStyle,IsCorporation
		Dim div_tf,f_sql,rs,cols_i,cols_j,PubType_JiTR,PubType_OuTR,VC_Class,DivID,Divclass,ulid,ulclass,liid,liclass
		UserType = f_Lable.LablePara("会员类型")
		OrderBy = f_Lable.LablePara("列表类型")
		UserSex = f_Lable.LablePara("会员性别")
		TitleNumber = f_Lable.LablePara("数量")
		if TitleNumber="" or not isnumeric(TitleNumber) then
			TitleNumber = 10
		else
			TitleNumber = TitleNumber
		end if
		ColsNumber = f_Lable.LablePara("列数")
		if ColsNumber = "" or not isnumeric(ColsNumber) then
			ColsNumber = 1
		else
			ColsNumber = cint(ColsNumber)
		end if
		leftTitle = f_Lable.LablePara("字数")
		DateType = f_Lable.LablePara("日期格式")
		ContentNumber = f_Lable.LablePara("内容字数")
		PubType_JiTR = f_Lable.LablePara("奇数行样式")
		PubType_OuTR = f_Lable.LablePara("偶数行样式")
		VC_Class = f_Lable.LablePara("行业分类")

		if PubType_JiTR<>"" then 
			if left(PubType_JiTR,1)="#" then 
				PubType_JiTR = " style=""background:"&PubType_JiTR&""""
			else
				PubType_JiTR = " class="""&PubType_JiTR&""""
			end if	
		end if	
		if PubType_OuTR<>"" then 
			if left(PubType_OuTR,1)="#" then 
				PubType_OuTR = " style=""background:"&PubType_OuTR&""""
			else
				PubType_OuTR = " class="""&PubType_OuTR&""""
			end if	
		end if	
		
		if f_Lable.LablePara("输入格式") = "out_Table" then
			div_tf = 0
		else
			div_tf = 1
			 DivId = f_Lable.LablePara("DivID")
			 if DivId <>"" then:DivId = " id="""&DivId&"""":else:DivId = "":end if
			 Divclass = f_Lable.LablePara("Divclass")
			 if Divclass <>"" then:Divclass = " class="""&Divclass&"""":else:Divclass = "":end if
			 ulid = f_Lable.LablePara("ulid")
			 if ulid <>"" then:ulid = " id="""&ulid&"""":else:ulid = "":end if
			 ulclass = f_Lable.LablePara("ulclass")
			 if ulclass <>"" then:ulclass = " class="""&ulclass&"""":else:ulclass = "":end if
			 liid = f_Lable.LablePara("liid")
			 if liid <>"" then:liid = " id="""&liid&"""":else:liid = "":end if
			 liclass = f_Lable.LablePara("liclass")
			 if liclass <>"" then:liclass = " class="""&liclass&"""":else:liclass = "":end if
		end if
		UserSex = Replacestr(UserSex,"All:,else: and Sex="&UserSex)
		OrderBy = " order by "&OrderBy&" desc,UserID Desc"

		Select Case UserType
		Case "1" '企业会员
			if VC_Class="0" then
			f_sql = "select top "&TitleNumber&" A.UserID,A.UserNumber,UserName,HeadPic,tel,Email,HomePage,QQ,RegTime,MSN,Province,City,Address,PostCode,NickName,RealName,Vocation,Sex,BothYear,Integral,FS_Money,IsMarray,hits,C_Name,C_ShortName,C_logo,C_Province,C_City,C_Address,C_ConactName,C_Sex,C_Vocation,C_Tel,C_Fax,C_VocationClassID,C_WebSite,C_size,C_Operation,C_Products,C_Content,C_BankName,C_BankUserName,C_property,C_Byear,C_PostCode from FS_ME_Users A,FS_ME_CorpUser B where A.UserNumber=B.UserNumber and A.isLock=0 and B.isLockCorp=0 "& UserSex & OrderBy
			else
			f_sql = "select top "&TitleNumber&" A.UserID,A.UserNumber,UserName,HeadPic,tel,Email,HomePage,QQ,RegTime,MSN,Province,City,Address,PostCode,NickName,RealName,Vocation,Sex,BothYear,Integral,FS_Money,IsMarray,hits,C_Name,C_ShortName,C_logo,C_Province,C_City,C_Address,C_ConactName,C_Sex,C_Vocation,C_Tel,C_Fax,C_VocationClassID,C_WebSite,C_size,C_Operation,C_Products,C_Content,C_BankName,C_BankUserName,C_property,C_Byear,C_PostCode from FS_ME_Users A,FS_ME_CorpUser B where B.C_VocationClassID="&cint(VC_Class)&" and A.UserNumber=B.UserNumber and A.isLock=0 and B.isLockCorp=0 "& UserSex & OrderBy '有行业选择时的分类情况 where B.C_VocationClassID="&cint(VC_Class)&" 为所加条件进行判断
			end if
			IsCorporation = 1	
		Case "0" '个人会员
			f_sql = "select top "&TitleNumber&" UserID,UserNumber,UserName,HeadPic,tel,Email,HomePage,QQ,RegTime,MSN,Province,City,Address,PostCode,NickName,RealName,Vocation,Sex,BothYear,Integral,FS_Money,IsMarray,hits,IsCorporation from FS_ME_Users where IsCorporation=0 and isLock=0 "& UserSex & OrderBy
			IsCorporation = 0
		Case "All" '所有会员
			f_sql = "select top "&TitleNumber&" UserID,UserNumber,UserName,HeadPic,tel,Email,HomePage,QQ,RegTime,MSN,Province,City,Address,PostCode,NickName,RealName,Vocation,Sex,BothYear,Integral,FS_Money,IsMarray,hits from FS_ME_Users where isLock=0 "& UserSex & OrderBy
			IsCorporation = 9
		End Select
		set rs = User_Conn.execute(f_sql)
		UserList = ""
		if rs.eof then
			UserList = ""
			rs.close:set rs = nothing
		else
			cols_i = 0 : cols_j = -1
			do while not rs.eof
				if div_tf = 1 then
					UserList = UserList & "  <li"&liid&liclass&">"&getlist_sd(rs,f_Lable.FoosunStyle.StyleContent,leftTitle,ContentNumber,DateType,"UserList",IsCorporation)&"</li>"
				else
					UserList = UserList & "<td>"&getlist_sd(rs,f_Lable.FoosunStyle.StyleContent,leftTitle,ContentNumber,DateType,"UserList",IsCorporation)&"</td>"

				end if
				rs.movenext 
				if rs.eof then exit do
				cols_i = cols_i + 1 : cols_j = cols_j + 1
				if div_tf = 0 then
					if cols_i mod ColsNumber = 0 then
						if cols_j mod 2 = 0 then
							UserList = UserList & "</tr>"&vbNewLine&"<tr"&PubType_OuTR&">"
						else
							UserList = UserList & "</tr>"&vbNewLine&"<tr"&PubType_JiTR&">"
						end if	
					end if
				end if
			loop
			rs.close:set rs = nothing
		end if
		if div_tf = 1 then
			UserList = "<div"&DivId&Divclass&">"&chr(10)&" <ul"&ulid&ulclass&">"&UserList&"  </ul>"&chr(10)&"</div>"&vbNewLine
		else
			UserList = "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr"&PubType_JiTR&">"&chr(10)&UserList&"</tr>"&chr(10)&"</table>"&vbNewLine
		end if	
		f_Lable.DictLableContent.Add "1",UserList
	end Function

	Public Function TopSubject(f_Lable,f_UserId)
		dim TitleNumber,leftTitle,TitleCssstr,CSS,datetf,dateType,classNews_head,classNews_middle1
		dim f_sql,rs,div_tf,f_TitlePara,classNews_bottom,classNews_middle2
		f_TitlePara = f_Lable.LablePara("调用数量,标题字数")
		TitleNumber = split(f_TitlePara,",")(0)
		leftTitle = split(f_TitlePara,",")(1)
		if leftTitle <> "" then
			leftTitle = leftTitle
		else
			leftTitle = 30 
		end if
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
		f_DivID = f_Lable.LablePara("DivID")
		f_DivClass = f_Lable.LablePara("Divclass")
		f_UlID = f_Lable.LablePara("ulid")
		f_ULClass = f_Lable.LablePara("ulclass")
		f_LiID = f_Lable.LablePara("liid")
		f_LiClass = f_Lable.LablePara("liclass")
		datetf= f_Lable.LablePara("显示日期")
		dateType = f_Lable.LablePara("日期样式")
		if f_Lable.LablePara("输出格式") = "out_DIV" then
			div_tf = 1
			TitleCssstr = ""
			classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		else
			div_tf=0
			CSS=f_Lable.LablePara("CSS")
			if trim(CSS)<>"" then
				TitleCssstr = " class="""& CSS &""""
			else
				TitleCssstr = ""
			end if
		end if
		classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		if f_UserId <>"" then
			f_sql = "select top "& TitleNumber &" ClassCName,ClassEName,ClassID,UserNumber from FS_ME_InfoClass where ClassTypes=7 and UserNumber='"& f_UserId &"'"
		else
			f_sql = "select top "& TitleNumber &" ClassCName,ClassEName,ClassID,UserNumber from FS_ME_InfoClass where ClassTypes=7"
		end if
		set rs = User_Conn.execute(f_sql)
		if rs.eof then
			TopSubject = ""
			rs.close:set rs=nothing
		else
			if f_UserId<>"" then
				TopSubject = "·<a href=""Blog.asp?id="&rs("ID")&""">首页</a><br />"
			end if
			do while not rs.eof
				if div_tf = 0 then
					TopSubject = TopSubject & "·<a href=""list.asp?id="&rs("ID")&"&ClassId="&rs("ClassID")&""">"& GotTopic(rs("ClassCName"),leftTitle)&"</a><br />"
				else
					TopSubject = TopSubject & classNews_middle1 & "<a href="""& get_subjectLink(rs("UserNumber"),rs("ClassEName"))&""">"& GotTopic(rs("ClassCName"),leftTitle)&"</a>"&classNews_middle2
				end if
				rs.movenext
			loop
			rs.close:set rs=nothing
		end if
		if div_tf=1 then TopSubject = classNews_head & TopSubject & classNews_bottom
		f_Lable.DictLableContent.Add "1",TopSubject
	End Function
	'总分类
	Public Function InfoClass(f_Lable)
		Dim div_tf,css,classNews_head,classNews_middle1,classNews_bottom,classNews_middle2
		dim f_sql,rs,sys_rs_obj,f_dir1,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
		f_DivID = f_Lable.LablePara("DivID")
		f_DivClass = f_Lable.LablePara("Divclass")
		f_UlID = f_Lable.LablePara("ulid")
		f_ULClass = f_Lable.LablePara("ulclass")
		f_LiID = f_Lable.LablePara("liid")
		f_LiClass = f_Lable.LablePara("liclass")
		css= f_Lable.LablePara("CSS")
		if f_Lable.LablePara("输出格式") = "out_DIV" then
			div_tf = 1
			classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		else
			div_tf = 0
			if css<>"" then
				css = " class="""&css&""""
			else
				css = ""
			end if
		end if
		classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		f_sql = "select id,ClassName from FS_ME_iLogClass order by id asc"
		set rs = User_Conn.execute(f_sql)
		InfoClass = ""
		if rs.eof then
			InfoClass = ""
			rs.close
			set rs = nothing
		else
			do while not rs.eof
				if div_tf = 0 then
					InfoClass = InfoClass & "<a href=""List.asp?id="& rs("id") &""""&css&">"&rs("ClassName")&"</a>&nbsp;"
				else
					InfoClass = InfoClass & "<a href=""list.asp?id="& rs("id") &""">"&rs("ClassName")&"</a>&nbsp;"
				end if
			rs.movenext
			loop
			if div_tf = 1 then
				InfoClass =  classNews_head & classNews_middle1 & InfoClass & classNews_middle2 & classNews_bottom
			end if
			rs.close
			set rs = nothing
		end if
		f_Lable.DictLableContent.Add "1",InfoClass
	End Function
	'日志终极分类
	Public Function InfoList(f_Lable,f_type,f_UserId,f_UserId_str)
		Dim CodeTitle,CodeTitlestr,TitleNumber,div_tf,titleCSS,DateTF,DateStyle
		Dim f_sql,rs,i,ReviewCountstr,str_tmps,DivId,DivClass,Ulid,ulclass,liid,liclass
		Dim f_TableName,f_SelectFieldNames,f_PageIndex,f_Where,f_PaginationStr,f_NewsContent
		Dim f_IDSRS,f_IDSArray,f_IDS,f_BeginIndex
		CodeTitle = f_Lable.LablePara("调用数量,标题字数")
		if split(CodeTitle,",")(0) = "" and isnumeric(split(CodeTitle,",")(0))=false then
			CodeTitlestr = 10
		else
			CodeTitlestr = split(CodeTitle,",")(0)
		end if
		if split(CodeTitle,",")(1) = "" and isnumeric(split(CodeTitle,",")(1))=false then
			TitleNumber = 40
		else
			TitleNumber = split(CodeTitle,",")(1)
		end if
		titleCSS = f_Lable.LablePara("CSS")
		DateTF = f_Lable.LablePara("显示日期")
		DateStyle = f_Lable.LablePara("日期样式")
		if f_Lable.LablePara("输出格式") = "out_Table" then
			div_tf = 0
		else
			div_tf = 1
			DivId = f_Lable.LablePara("DivID")
			DivClass = f_Lable.LablePara("DivClass")
			Ulid = f_Lable.LablePara("ulid")
			ulclass = f_Lable.LablePara("ulclass")
			liid = f_Lable.LablePara("liid")
			liclass = f_Lable.LablePara("liclass")
			if DivId<>"" then
				DivId = " id="""&DivId&""""
			else
				DivId = ""
			end if
			if DivClass<>"" then
				DivClass = " class="""&DivClass&""""
			else
				DivClass = ""
			end if
		end if
		if titleCSS<>"" then
			titleCSS = " class="""&titleCSS&""""
		else
			titleCSS = ""
		end if
		if f_UserId_str<>"" then
			str_tmps = " and UserNumber = '"&f_UserId&"' and ClassID="& f_UserId_str&""
		else
			str_tmps = " and UserNumber = '"&f_UserId&"'"
		end if
		f_PaginationStr = "3,CC0066,"
		f_Lable.DictLableContent.Item("0") = f_PaginationStr
		f_Where = "isDraft=0 and adminLock=0 "&str_tmps&" order by isTop desc,isTF desc,iLogID desc"
		f_TableName = "FS_ME_Infoilog"
		f_SelectFieldNames = "iLogID,iLogStyle,Title,KeyWords,Content,UserNumber,ClassID,Hits,savePath,FileName,FileExtName,AddTime"
		Set f_IDSRS = Server.CreateObject(G_FS_RS)
		f_IDSRS.Open "Select iLogID from " & f_TableName & " Where " & f_Where,Conn,0,1
		if Not f_IDSRS.Eof then
			f_PageIndex = 1
			f_IDSArray = f_IDSRS.GetRows()
			Set  rs = Server.CreateObject(G_FS_RS)
			Do While True
				f_IDS = ""
				f_BeginIndex = (f_PageIndex - 1) * PageNumber
				for i = f_BeginIndex To PageNumber * f_PageIndex - 1
					if i > UBound(f_IDSArray,2) then Exit For
					if f_IDS = "" then
						f_IDS = f_IDSArray(0,i)
					else
						f_IDS = f_IDS & "," & f_IDSArray(0,i)
					end if
				Next
				if f_IDS = "" then Exit Do
				f_Where = " iLogID In (" & f_IDS & ")"
				f_sql = "Select " & f_SelectFieldNames & " from " & f_TableName & " where " & f_Where
				rs.open f_sql,Conn,0,1
				i = 0
				InfoList = ""
				Do While Not rs.Eof
					if div_tf = 1 then InfoList = InfoList & "<a name=""top"" id=""top""></a>"&chr(10)&"<div"&DivId&DivClass&">"
					dim rs_review
					set rs_review = User_Conn.execute("select Count("&rs("iLogId")&") from FS_ME_Review where InfoID = "& rs("iLogId")&" and ReviewTypes=5 and AdminLock=0 and isLock=0")
					if rs_review.eof then
						ReviewCountstr = 0
						rs_review.close:set rs_review= nothing
					else
						ReviewCountstr = rs_review(0)
						rs_review.close:set rs_review= nothing
					end if
					dim date_tmp
					if DateTF="1" then
						date_tmp = get_DateChar(DateStyle,rs("AddTime")) &"┆"
					else
						date_tmp = ""
					end if
					InfoList = InfoList & "<h4><a href="""& Get_LogLink(rs("iLogID")) &""""&titleCSS&"><font style=""font-size:14px""><strong>"&GotTopic(""&rs("Title"),TitleNumber)&"</a></strong></font></h4>"&chr(10)
					InfoList = InfoList & ""&GetCStrLen(""&rs("Content"),800)&"...<br /><br />"&chr(10)
					if trim(replace(rs("KeyWords"),",",""))<>"" then
						dim o_i
						InfoList = InfoList & "<font style=""font-size:14px"">Tags: "
						for o_i = LBound(split(rs("KeyWords"),",")) to UBound(split(rs("KeyWords"),","))
							InfoList = InfoList & "<a href="""&replace("/"& G_VIRTUAL_ROOT_DIR&"/","//","/")&"Tags.html?Tags="&split(rs("KeyWords"),",")(o_i)&""">"&split(rs("KeyWords"),",")(o_i)&"</a>&nbsp;"
						next
						InfoList = InfoList & "</font><br /><br />"
					end if
					InfoList = InfoList & "日期：<img src="""& replace("/"& G_VIRTUAL_ROOT_DIR&"/","//","/")&"sys_images/post_yellow.gif"" border=""0"" />&nbsp;"& date_tmp &"<a href="""& Get_LogLink(rs("iLogID")) &""">阅读全文</a>┆发布：<a href="""& replace("/"& G_VIRTUAL_ROOT_DIR&"/","//","/")& G_USER_DIR &"/ShowUser.asp?UserNumber="&rs("UserNumber")&""" target=""_blank"">用户信息</a>,<a href=""list.asp?id="&rs("ID")&""" target=""_blank"">"&get_UserNameLink(rs("UserNumber"))&"的日志</a>┆分类:<a href=""list.asp?id="&rs("ID")&"&ClassId="& rs("ClassID") &""">"&get_UserClassName(rs("ClassID"))&"</a>┆浏览:"&rs("hits")&"┆评论:"&ReviewCountstr&"┆<a href=""#"">↑TOP</a><br />"&chr(10)
					if div_tf = 1 then
						InfoList = InfoList & "</div>"
					end if
					rs.movenext
					i = i + 1
				Loop
				f_Lable.DictLableContent.Add f_PageIndex & "",InfoList
				rs.Close
				f_PageIndex = f_PageIndex + 1
			Loop 
			rs.close
			set rs = nothing
		else
			f_Lable.DictLableContent.Add "1",""
		end if
		f_IDSRS.Close
		Set f_IDSRS = Nothing
	End Function
	'最新评论
	Public Function Log_LastReview(f_Lable,f_type,f_UserId)
		Dim CodeNumber,TitleNumber,lefttitle,div_tf,CSS,DateTF,DateType,rs,DivId,DivClass,Ulid,ulclass,liid,liclass,f_SQL
		CodeNumber = f_Lable.LablePara("调用数量,标题字数")
		TitleNumber = split(CodeNumber,",")(0)
		if TitleNumber="" or isnumeric(TitleNumber)=false then
			TitleNumber = 10
		else
			TitleNumber = TitleNumber
		end if
		lefttitle = split(CodeNumber,",")(1)
		if lefttitle="" or isnumeric(lefttitle)=false then
			lefttitle = 40
		else
			lefttitle = lefttitle
		end if
		CSS = f_Lable.LablePara("CSS")
		DateTF = f_Lable.LablePara("显示日期")
		DateType = f_Lable.LablePara("日期样式")
		if f_Lable.LablePara("输出格式") = "out_Table" then
			div_tf = 0
		else
			div_tf = 1
			DivId = f_Lable.LablePara("DivID")
			if DivId = "" then
				DivId = ""
			else
				DivId = " id="""&DivId&""""
			end if
			DivClass = f_Lable.LablePara("DivClass")
			if DivClass = "" then
				DivClass = ""
			else
				DivClass = " class="""&DivClass&""""
			end if
			Ulid = f_Lable.LablePara("ulid")
			if Ulid = "" then
				Ulid = ""
			else
				Ulid = " id="""&Ulid&""""
			end if
			ulclass = f_Lable.LablePara("ulclass")
			if ulclass = "" then
				ulclass = ""
			else
				ulclass = " class="""&ulclass&""""
			end if
			liid = f_Lable.LablePara("liid")
			if liid = "" then
				liid = ""
			else
				liid = " id="""&liid&""""
			end if
			liclass = f_Lable.LablePara("liclass")
			if liclass = "" then
				liclass = ""
			else
				liclass = " class="""&liclass&""""
			end if
		end if
		if CSS="" then
			CSS = ""
		else
			CSS = " class="""& CSS &""""
		end if
		if f_UserId <> "" then
			f_SQL = "select top "& TitleNumber &" ReviewID,InfoID,UserNumber,ReviewTypes,Title,Content,AddTime From FS_ME_Review where infoId="&f_UserId&" and ReviewTypes=5 and AdminLock=0 and isLock=0 order by ReviewID desc"
		else
			f_SQL = "select top "& TitleNumber &" ReviewID,InfoID,UserNumber,ReviewTypes,Title,Content,AddTime From FS_ME_Review where ReviewTypes=5 and AdminLock=0 and isLock=0 order by ReviewID desc"
		end if
		set rs  = User_Conn.execute(f_SQL)
		Log_LastReview = ""
		if rs.eof then
			Log_LastReview = "没有评论"
			rs.close
			set rs = nothing
		else	
			if div_tf = 1 then
				Log_LastReview = Log_LastReview & "<div"&DivId&DivClass&"><ul"&Ulid&ulclass&">"&chr(10)
			end if	
			do while not rs.eof 
				dim date_tmp
				if DateTF="1" then
					date_tmp = get_DateChar(DateType,rs("AddTime"))
				else
					date_tmp = ""
				end if
				if div_tf = 1 then
					Log_LastReview = Log_LastReview & " <li"&liid&liclass&">"&GotTopic(""&rs("Title"),lefttitle)&" <a href=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"&G_USER_DIR&"/ShowUser.asp?UserNumber="& rs("UserNumber")&""" target=""_blank"">"&get_UserNameLink(rs("UserNumber"))&"</a><font style=""font-size:11px""> "&date_tmp&"</font></li>"&chr(10)
				else
					Log_LastReview = Log_LastReview & "<span"&CSS&">·"&GotTopic(""&rs("Title"),lefttitle)&" <a href=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"&G_USER_DIR&"/ShowUser.asp?UserNumber="& rs("UserNumber")&""" target=""_blank"">"&get_UserNameLink(rs("UserNumber"))&"</a><font style=""font-size:11px""> "&date_tmp&"</font></span><br />"
				end if 
			rs.movenext
			loop
			rs.close
			set rs = nothing
		end if
		f_Lable.DictLableContent.Add "1",Log_LastReview
	End Function
	'评论表单
	Public Function Log_LastForm(f_Lable,f_type,f_UserId)
		Dim FormReview
		FormReview="	<form action=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/ReviewUrl.asp"" name=""reviewform"" method=""post"" onsubmit=""//return checkreview();""/>"&chr(10)
		FormReview=FormReview&"	  用户名<input name=""UserNumber"" type=""text"" id=""UserNumber"" size=""15""/>"&chr(10)
		FormReview=FormReview&"	  匿名<input name=""noname"" type=""checkbox"" id=""noname"" value=""1""/>"&chr(10)
		FormReview=FormReview&"	  密码<input name=""password"" type=""password"" id=""password"" size=""12""/><input type=""hidden"" name=""newsid"" value="""&f_UserId&"""/><input type=""hidden"" name=""Action"" value=""add_save""/><br />"&chr(10)
		FormReview=FormReview&"	  标　题<input name=""title"" type=""text"" id=""title"" size=""30""/><input type=""hidden"" name=""type"" value=""Log""/><br />"&chr(10)
		FormReview=FormReview&"	  <textarea name=""content"" cols=""50"" rows=""5"" id=""content""/></textarea><br />"&chr(10)
		FormReview=FormReview&"	  <input type=""submit"" name=""Submit"" value=""发表评论""/>&nbsp;&nbsp;<input type=""reset"" name=""Submit2"" value=""重新填写""/>"&chr(10)
		FormReview=FormReview&"	</form>"&chr(10)
		f_Lable.DictLableContent.Add "1",FormReview
	End Function
	'发表日志接口
	Public Function Log_PublicLog(f_Lable)
		Log_PublicLog = "<a href=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"&G_USER_DIR&"/i_Blog/PublicLog.asp"" target=""_blank""><img src=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/sys_images/Pub_log.gif"" border=""0"" alt=""发布日志""></a>"
		f_Lable.DictLableContent.Add "1",Log_PublicLog
	End Function
	'搜索
	Public Function Log_Search(f_Lable,f_id)
		Log_Search = "<form action=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Log_Search.asp"" name=""Log_Form"" method=""post"" id=""Log_Form""><input name=""keywords"" class=""f-text"" value="""" type=""text"" size=""12""><input name=""id"" value="""&f_id&""" type=""hidden"" size=""12"">&nbsp;<input name=""sumbit"" class=""f-button"" type=""submit"" value=""搜索""></form>"
		f_Lable.DictLableContent.Add "1",Log_Search
	End Function
	'总站日志导航
	Public Function Log_Navi(f_Lable)
		dim div_tf,CSS,DivId,DivClass,Ulid,ulclass,liid,liclass,rs
		CSS = f_Lable.LablePara("CSS")
		if f_Lable.LablePara("输出格式") = "out_Table" then
			div_tf = 0 
		else
			div_tf = 1
			DivId = f_Lable.LablePara("DivID")
			if DivId = "" then:DivId = "":else:DivId = " id="""&DivId&"""":end if
			DivClass = f_Lable.LablePara("DivClass")
			if DivClass = "" then:DivClass = "":else:DivClass = " class="""&DivClass&"""":end if
			Ulid = f_Lable.LablePara("ulid")
			if Ulid = "" then:Ulid = "":else:Ulid = " id="""&Ulid&"""":end if
			ulclass = f_Lable.LablePara("ulclass")
			if ulclass = "" then:ulclass = "":else:ulclass = " class="""&ulclass&"""":end if
			liid = f_Lable.LablePara("liid")
			if liid = "" then:liid = "":else:liid = " id="""&liid&"""":end if
			liclass = f_Lable.LablePara("liclass")
			if liclass = "" then:liclass = "":else:liclass = " class="""&liclass&"""":end if
		end if
		if CSS="" then:CSS = "":else:CSS = " class="""& CSS &"""":end if
		set rs = User_Conn.execute("select id,ClassName From FS_ME_iLogClass order by id desc")
		Log_Navi = ""
		if rs.eof then
			Log_Navi = "没系统分类"
			rs.close:set rs =nothing
		else
			if div_tf = 1 then
				Log_Navi = Log_Navi & "<div"&DivId&DivClass&">"&chr(10)&" <ul"&Ulid&ulclass&">"&chr(10)
			end if 
			do while not rs.eof
				if div_tf = 1 then
					Log_Navi = Log_Navi & "   <li"&liid&liclass&"><a href=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"&Request.Cookies("FoosunMECookies")("FoosunMELogDir")&"/syslist.asp?id="&rs("id")&""">"&rs("ClassName")&"</a></li>"&chr(10)
				else
					Log_Navi = Log_Navi & "   <span"&CSS&"><a href=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"&Request.Cookies("FoosunMECookies")("FoosunMELogDir")&"/syslist.asp?id="&rs("id")&""">"&rs("ClassName")&"</a></span>&nbsp;"
				end if
				rs.movenext
			loop
			rs.close:set rs =nothing
		end if
		if div_tf = 1 then
			Log_Navi = Log_Navi &chr(10)&" <ul>"&chr(10)&"</div>"
		end if
		f_Lable.DictLableContent.Add "1",Log_Navi
	End Function
	'页面标题
	Public Function Log_PageTitle(f_Lable,f_id)
		dim rs
		set rs = User_Conn.execute("select top 1 siteName From FS_ME_InfoiLogParam")
		if rs.eof then
			Log_PageTitle = rs(0)
		else
			Log_PageTitle = ""
		end if
		rs.close
		set rs = nothing
		f_Lable.DictLableContent.Add "1",Log_PageTitle
	End Function
	'日志标题
	Public Function Log_Page(f_Lable,f_id)
		dim rs,f_LableName
		f_LableName = f_Lable.Lablename
		If f_id="" Or IsNull(f_id) Then f_id=0
		set rs =User_Conn.execute("select * From FS_ME_Infoilog where iLogID="&f_id)
		if rs.eof then
			Log_Page = ""
			rs.close
			set rs = nothing
		else
			if instr(f_LableName,"Log_title")>0 then
				Log_Page =  replace(f_LableName,"Log_title",""&rs("title"))
			end if
			if instr(f_LableName,"Log_Content")>0 then
				Log_Page =  replace(f_LableName,"Log_Content",""&rs("Content"))
			end if
			if instr(f_LableName,"Log_Author")>0 then
				Log_Page =  replace(f_LableName,"Log_Author","<a href=""list.asp?id="&rs("iLogID")&""">"&get_UserNameLink(rs("UserNumber"))&"</a>")
			end if
			if instr(f_LableName,"Log_hits")>0 then
				Dim hits_str
				hits_str = "<script language=""JavaScript"" src=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Click.asp?type=ajax&Subsys=Log&spanid=Log_id_click_"&rs("iLogID")&"""></script>"&chr(10)
				Log_Page =  replace(f_LableName,"Log_hits",hits_str&"<span id=""Log_id_click_"&rs("iLogID")&""">loading...</span>")
			end if
			if instr(f_LableName,"Log_keywords")>0 then
				dim i 
				if instr(rs("keywords"),",")>0 then
					for i = 0 to Ubound(split(rs("keywords"),","))
						if trim(split(rs("keywords"),",")(i))<>"" then
							Log_Page =  Log_Page & replace(f_LableName,"Log_keywords","<a href=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/tags.html?tags="& split(rs("keywords"),",")(i)&""" target=""_blank"">"&split(rs("keywords"),",")(i)&"</a>&nbsp;")
						else
							Log_Page = ""
						end if
					next
				end if
			end if
			if instr(f_LableName,"Log_AddTime")>0 then
				Log_Page =  replace(f_LableName,"Log_AddTime",""&rs("AddTime"))
			end if
			if instr(f_LableName,"Log_LogType")>0 then
				Log_Page =  replace(f_LableName,"Log_LogType","主分类:<a href=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"& Request.Cookies("FoosunMECookies")("FoosunMELogDir") &"/list.asp?ClassId="& rs("MainID") &""">"&get_UsersysClassName(rs("MainID"))&"</a>,个人分类：<a href=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"& Request.Cookies("FoosunMECookies")("FoosunMELogDir") &"/list.asp?ClassId="& rs("ClassID") &"&id="&rs("iLogID")&""">"&get_UserClassName(rs("ClassID"))&"</a>")
			end if
			if instr(f_LableName,"Log_LogType")>0 then
				Log_Page =  replace(f_LableName,"Log_LogType","主分类:<a href=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"& Request.Cookies("FoosunMECookies")("FoosunMELogDir") &"/list.asp?ClassId="& rs("MainID") &""">"&get_UsersysClassName(rs("MainID"))&"</a>,个人分类：<a href=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"& Request.Cookies("FoosunMECookies")("FoosunMELogDir") &"/list.asp?ClassId="& rs("ClassID") &"&id="&rs("iLogID")&""">"&get_UserClassName(rs("ClassID"))&"</a>")
			end if
			if instr(f_LableName,"Log_ReviewList")>0 then
				Log_Page = replace(f_LableName,"Log_ReviewList","<script language=""JavaScript"" type=""text/javascript"" src=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/ShowReview.asp?Id="&RS("iLogID")&"&Type=LOG&SpanId=Log_show_review_"& rs("iLogID") &"""></script><label id=""Log_show_review_"& rs("iLogID") &""">评论加载中...</label>")
			end if
			if instr(f_LableName,"Log_ReviewForm")>0 then
				Log_Page = replace(f_LableName,"Log_ReviewForm","<span id=""Review_TF_"& rs("iLogID") &""">loading...</span><script language=""JavaScript"" type=""text/javascript"" src=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/ReviewTF.asp?Id="& rs("iLogID") &"&Type=LOG""></script>")
			end if
			rs.close
			set rs = nothing
		end if
		f_Lable.DictLableContent.Add "1",Log_Page
	End Function
	'得到日志连接
	Public Function Get_LogLink(f_id)
		dim rs,rs_sys
		set rs = User_Conn.execute("select iLogID,FileName,FileExtName,UserNumber,savePath From FS_ME_Infoilog where iLogID="&f_id)
		Get_LogLink = "http://"&Get_MF_Domain&"/"& replace(Request.Cookies("FoosunMECookies")("FoosunMELogDir")&"/Blog.asp?id="&rs("iLogID"),"//","/")
		rs.close:set rs=nothing
		Get_LogLink = Get_LogLink
	End Function
	'相册调用
	Public Function Photo_List(f_Lable,f_type)
		Dim BookPop,TitleNumber,ColsNumber,leftTitle,PizSize,div_tf,PicCSS,TitleCSS
		Dim classNews_head,classNews_middle1,classNews_bottom,classNews_middle2,pizh,pizw
		dim f_sql,rs,DivId,DivClass,Url_Domain_path,a_href_link,orderby,rec_tf,tmp_I
		BookPop = f_Lable.LablePara("类型")
		TitleNumber = f_Lable.LablePara("数量")
		if TitleNumber="" or isnumeric(TitleNumber)=false then
			TitleNumber = 10
		else
			TitleNumber = TitleNumber
		end if
		ColsNumber = f_Lable.LablePara("列数")
		if ColsNumber="" or isnumeric(ColsNumber)=false then
			ColsNumber = 5
		else
			ColsNumber = ColsNumber
		end if
		leftTitle = f_Lable.LablePara("字数")
		PizSize = f_Lable.LablePara("图片大小")
		if PizSize<>"" and instr(PizSize,",")>0 then
			if split(PizSize,",")(0)= "0" then
				pizh = ""
				pizw = " width="""&split(PizSize,",")(1)&""""
			else
				if split(PizSize,",")(1)= "0" then
					pizw = ""
					pizh = " width="""&split(PizSize,",")(0)&""""
				else
					pizh = " height="""& split(PizSize,",")(0) &""""
					pizw = " width="""&split(PizSize,",")(1)&""""
				end if
			end if
		else
			pizh = " height=""100"""
			pizw = " width=""80"""
		end if
		PicCSS = f_Lable.LablePara("图片CSS")
		TitleCSS = f_Lable.LablePara("标题CSS")
		if PicCSS<>"" then
			PicCSS = " class="""&PicCSS&""""
		else
			PicCSS = ""
		end if
		if TitleCSS<>"" then
			TitleCSS = " class="""&TitleCSS&""""
		else
			TitleCSS = ""
		end if
		
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
		f_DivID = f_Lable.LablePara("DivID")
		f_DivClass = f_Lable.LablePara("Divclass")
		f_UlID = f_Lable.LablePara("ulid")
		f_ULClass = f_Lable.LablePara("ulclass")
		f_LiID = f_Lable.LablePara("liid")
		f_LiClass = f_Lable.LablePara("liclass")
		
		if f_Lable.LablePara("输入格式") = "out_DIV" then
			div_tf=1
			classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		else
			div_tf=0
		end if
		classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		'---------------chen-----------
		Select case BookPop
		case "0"'正常情况下显示除推荐外的所有像册。
			orderby = " order by id desc"
			rec_tf = " and isRec=0"
		case "1" '显示推荐的像册 。
			orderby = " order by id desc"
			rec_tf = " and isRec=1"
		case "2" '根据点饥率来显示像册 。
			orderby = " order by Hits desc,id desc"
			rec_tf = ""
		End Select
		'-------------------chen----------------
		f_sql = "select top "& TitleNumber &" Id,title,PicSavePath,Content,Addtime,UserNumber,PicSize,Hits,isRec From FS_ME_Photo where 1=1 "&rec_tf&" "&orderby&""
		set rs = User_Conn.execute(f_sql)
		Photo_List = ""
		Url_Domain_path = "http://" & request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		if rs.eof then
			Photo_List = ""
			rs.close:set rs = nothing
		else
			tmp_I = 0 
			do while not rs.eof
				a_href_link=Url_Domain_path & "/"&Request.Cookies("FoosunMECookies")("FoosunMELogDir")&"/ShowPhoto.asp?Id="&rs("Id")&""
				if div_tf=1 then
					if f_DivId<>"" Then:f_DivId = " id="""&f_DivId&"""":else:f_DivId = "":end if
					if f_DivClass<>"" Then:f_DivClass = " class="""&f_DivClass&"""":else:f_DivClass = "":end if
					Photo_List = Photo_List & "  <div"& f_DivId & f_DivClass &" align=""center""><a href="""&a_href_link&"""><img src="""& Url_Domain_path & rs("PicSavePath") &""""&PicCSS&""& pizh & pizw &" border=""0"" /></a><br /><a href="""&a_href_link&""""&TitleCSS&">"&GotTopic(""&rs("title"),leftTitle)&"</a></div>"&chr(10)
				else
					Photo_List = Photo_List & "  <td width="""&cint(100/ColsNumber)&"%"" align=""center""><a href="""&a_href_link&"""><img src="""& Url_Domain_path & rs("PicSavePath") &""""&PicCSS&""& pizh & pizw &" border=""0"" /></a><br /><a href="""&a_href_link&""""&TitleCSS&">"&GotTopic(""&rs("title"),leftTitle)&"</a></td>"&chr(10)
				end if
				rs.movenext
				if div_tf = 0 then
					tmp_I = tmp_I + 1
					if tmp_I mod ColsNumber =0 then
						Photo_List = Photo_List & "</tr><tr>"
					end if
				end if
			loop
			rs.close:set rs = nothing
		end if
		if div_tf=1 then
			Photo_List =  Photo_List 
		else
			Photo_List = "<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""5""><tr>" & Photo_List & "</tr></table>"
		end if
		Photo_List = Photo_List
		f_Lable.DictLableContent.Add "1",Photo_List
	End Function
	'相册浏览
	Public Function Photo_show(f_Lable,f_id)
		Dim lefttitle,Picsize,piccss,titlecss,picheight,picwidth,rs,f_sql
		lefttitle = f_Lable.LablePara("字数")
		Picsize = f_Lable.LablePara("图片大小")
		piccss = f_Lable.LablePara("图片CSS")
		titlecss = f_Lable.LablePara("标题CSS")
		if not isnumeric(lefttitle) then:lefttitle =50:else:lefttitle = cint(lefttitle):end if
		if trim(piccss)<>"" then
			piccss =" class="""& piccss &""""
		else
			piccss = ""
		end if
		if trim(titlecss)<>"" then:titlecss =" class="""& titlecss &"""":else:titlecss = "":end if
		picheight = split(Picsize,",")(0)
		if picheight <> "0" then
			picheight = " height ="""&picheight&""""
		else
			picheight = ""
		end if
		picwidth = split(Picsize,",")(1)
		if picwidth <> "0" then
			picwidth = " width ="""&picwidth&""""
		else
			picwidth = ""
		end if
		Photo_show = ""
		if not isnumeric(f_id) then
			Photo_show = "错误参数"
		else
			f_sql = "select top 1 id,title,Content,PicSavePath,Addtime,UserNumber,Hits From FS_ME_Photo where id="& clng(f_id)
			set rs = User_Conn.execute(f_sql)
			if rs.eof then
				Photo_show = "错误参数"
				rs.close:set rs = nothing
			else
				Photo_show = Photo_show & "<a href="""& rs("PicSavePath") &""" target=""_blank""><img src = """& rs("PicSavePath") &""""& piccss & picheight & picwidth &" /></a><br /><br />"& rs("title")& "<br /><div align=""left"">说明:"& rs("Content")& "</div><br /><br />"
				Photo_show = Photo_show & "作者：<a href=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"&G_USER_DIR&"/ShowUser.asp?UserNumber="&rs("UserNumber")&""" target=""_blank"">"&get_UserNameLink(rs("UserNumber"))&"</a>&nbsp;&nbsp;日期："&rs("Addtime")&"&nbsp;&nbsp;阅读次数:<script language=""JavaScript"" src=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Click.asp?type=ajax&SubSys=PHOTO&spanid=PH_id_click_"&rs("id")&"""></script><div id=PH_id_click_"&rs("id")&"></div>"
				rs.close:set rs = nothing
			end if
		end if
		f_Lable.DictLableContent.Add "1",Photo_show
	End function
	'专题地址
	Public Function get_subjectLink(f_Usernumber,f_dir)
		get_subjectLink = "http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"&Request.Cookies("FoosunMECookies")("FoosunMELogDir")&"/"&f_Usernumber&"/"&f_dir&".html"
	End Function
	'得到作者地址
	Public Function get_AuthorLink(f_user)
			get_AuthorLink = "http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"&G_USER_DIR&"/ShowUser.asp?UserNumber="&f_user&""
	End Function
	'得到作者用户名
	Public Function get_UserNameLink(f_usernumber)
		dim rs
		set rs = User_Conn.execute("select UserName From FS_ME_Users where UserNumber='"&f_usernumber&"'")
		if rs.eof then
			get_UserNameLink = ""
			rs.close:set rs=nothing
		else
			get_UserNameLink = rs("UserName")
			rs.close:set rs=nothing
		end if
	end Function
	'得到分类名称
	Public Function get_UserClassName(f_id)
		dim rs
		set rs = User_Conn.execute("select ClassCName From FS_ME_InfoClass where ClassID="&f_id&"")
		if rs.eof then
			get_UserClassName = "未分类"
			rs.close:set rs=nothing
		else
			get_UserClassName = rs("ClassCName")
			rs.close:set rs=nothing
		end if
	end Function
	'得到系统分类名称
	Public Function get_UsersysClassName(f_id)
		dim rs
		set rs = User_Conn.execute("select ClassName From FS_ME_iLogClass where id="&f_id&"")
		if rs.eof then
			get_UsersysClassName = "未分类"
			rs.close:set rs=nothing
		else
			get_UsersysClassName = rs("ClassName")
			rs.close:set rs=nothing
		end if
	end Function
	'得到日期格式
	Public Function get_DateChar(f_datestyle,f_addtime)
		dim tmp_f_datestyle
		tmp_f_datestyle = f_datestyle
		if instr(f_datestyle,"YY02")>0 then
			tmp_f_datestyle= replace(tmp_f_datestyle,"YY02",right(year(f_addtime),2))
		end if
		if instr(f_datestyle,"YY04")>0 then
			tmp_f_datestyle= replace(tmp_f_datestyle,"YY04",year(f_addtime))
		end if
		if instr(f_datestyle,"MM")>0 then
			if month(f_addtime)<10 then
				tmp_f_datestyle= replace(tmp_f_datestyle,"MM","0"&month(f_addtime))
			else
				tmp_f_datestyle= replace(tmp_f_datestyle,"MM",month(f_addtime))
			end if
		end if
		if instr(f_datestyle,"DD")>0 then
			if day(f_addtime)<10 then
				tmp_f_datestyle= replace(tmp_f_datestyle,"DD","0"&day(f_addtime))
			else
			
				tmp_f_datestyle= replace(tmp_f_datestyle,"DD",day(f_addtime))
			end if
		end if
		if instr(f_datestyle,"HH")>0 then
			if hour(f_addtime)<10 then
				tmp_f_datestyle= replace(tmp_f_datestyle,"HH","0"&hour(f_addtime))
			else
				tmp_f_datestyle= replace(tmp_f_datestyle,"HH",hour(f_addtime))
			end if
		end if
		if instr(f_datestyle,"MI")>0 then
			if minute(f_addtime)<10 then
				tmp_f_datestyle= replace(tmp_f_datestyle,"MI","0"&minute(f_addtime))
			else
				tmp_f_datestyle= replace(tmp_f_datestyle,"MI",minute(f_addtime))
			end if
		end if
		if instr(f_datestyle,"SS")>0 then
			if second(f_addtime)<10 then
				tmp_f_datestyle= replace(tmp_f_datestyle,"SS","0"&second(f_addtime))
			else
				tmp_f_datestyle= replace(tmp_f_datestyle,"SS",second(f_addtime))
			end if
		end if
		get_DateChar = tmp_f_datestyle
	end Function
	
	Public Function log_usertile(f_Lable,f_id)
		dim rs,f_LableName
		f_LableName = f_Lable.LableName
		if InStr(f_LableName,"Log_Usertitle") > 0 then
			set rs = User_Conn.execute("select siteName from FS_ME_InfoiLogParam where UserNumber='"& f_id &"'")
			if not rs.eof then
				log_usertile = rs("siteName")
				rs.close:set rs =nothing
			else
				log_usertile = "风讯blog"
				rs.close:set rs =nothing
			end if
		elseif InStr(f_LableName,"Log_UserContent") > 0 then
			set rs = User_Conn.execute("select Content from FS_ME_InfoiLogParam where UserNumber='"& f_id &"'")
			if not rs.eof then
				log_usertile = rs("Content")
				rs.close:set rs =nothing
			else
				log_usertile = "风讯blog"
				rs.close:set rs =nothing
			end if
		elseif InStr(f_LableName,"Log_UserName") > 0 then
			set rs = User_Conn.execute("select UserName,RealName,userNumber from FS_ME_Users where UserNumber='"& f_id &"'")
			if not rs.eof then
				log_usertile = rs("UserName")
				rs.close:set rs =nothing
			else
				log_usertile = ""
				rs.close:set rs =nothing
			end if
		elseif InStr(f_LableName,"Log_NickName") > 0 then
			set rs = User_Conn.execute("select NickName from FS_ME_Users where UserNumber='"& f_id &"'")
			if not rs.eof then
				log_usertile = rs("NickName")
				rs.close:set rs =nothing
			else
				log_usertile = ""
				rs.close:set rs =nothing
			end if
		end if
		f_Lable.DictLableContent.Add "1",log_usertile
	End Function
	'会员列表
	Public Function getlist_sd(f_obj,s_Content,f_titlenumber,f_contentnumber,f_datestyle,f_subsys_ListType,IsCorporation)
		dim f_obj1
		if instr(s_Content,"{ME:FS_UserNumber}")>0 then
			s_Content = replace(s_Content,"{ME:FS_UserNumber}",""&f_obj("UserNumber"))
		end if
		if instr(s_Content,"{ME:FS_UserName}")>0 then
			s_Content = replace(s_Content,"{ME:FS_UserName}",""&GotTopic(""&f_obj("UserName"),f_titlenumber))
		end if
		if instr(s_Content,"{ME:FS_NickName}")>0 then
			s_Content = replace(s_Content,"{ME:FS_NickName}",""&GotTopic(""&f_obj("NickName"),f_titlenumber))
		end if
		if instr(s_Content,"{ME:FS_RealName}")>0 then
			s_Content = replace(s_Content,"{ME:FS_RealName}",""&GotTopic(""&f_obj("RealName"),f_titlenumber))
		end if
		if instr(s_Content,"{ME:FS_Sex}")>0 then
			s_Content = replace(s_Content,"{ME:FS_Sex}",""&Replacestr(f_obj("Sex"),"0:男,1:女"))
		end if
		if instr(s_Content,"{ME:FS_QQ}")>0 then
			s_Content = replace(s_Content,"{ME:FS_QQ}",""&f_obj("QQ"))
		end if
		if instr(s_Content,"{ME:FS_tel}")>0 then
			s_Content = replace(s_Content,"{ME:FS_tel}",""&f_obj("tel"))
		end if
		if instr(s_Content,"{ME:FS_Email}")>0 then
			s_Content = replace(s_Content,"{ME:FS_Email}","<a href=""mailto:"&f_obj("Email")&""">"&f_obj("Email")&"</a>")
		end if
		if instr(s_Content,"{ME:FS_HomePage}")>0 then
			s_Content = replace(s_Content,"{ME:FS_HomePage}",""&f_obj("HomePage"))
		end if
		if instr(s_Content,"{ME:FS_HeadPic}")>0 then
			if isnull(trim(f_obj("HeadPic"))) or len(trim(f_obj("HeadPic")))<5 then
				s_Content = replace(s_Content,"{ME:FS_HeadPic}","http://"& Request.Cookies("FoosunMFCookies")("FoosunMFDomain") &"/sys_images/NoPic.jpg")
			else
				s_Content = replace(s_Content,"{ME:FS_HeadPic}",""&f_obj("HeadPic"))
			end if
		end if
		if instr(s_Content,"{ME:FS_QQ}")>0 then
			s_Content = replace(s_Content,"{ME:FS_QQ}","<a http://wpa.qq.com/msgrd?V=1&amp;Uin="&f_obj("QQ")&"&amp;Site=&amp;Menu=yes"" target=""_blank"" title=""点击和"&f_obj("QQ")&"交谈"">"&f_obj("QQ")&"</a>")
		end if
		if instr(s_Content,"{ME:FS_MSN}")>0 then
			s_Content = replace(s_Content,"{ME:FS_MSN}",""&f_obj("MSN"))
		end if
		if instr(s_Content,"{ME:FS_Province}")>0 then
			s_Content = replace(s_Content,"{ME:FS_Province}",""&f_obj("Province"))
		end if
		if instr(s_Content,"{ME:FS_City}")>0 then
			s_Content = replace(s_Content,"{ME:FS_City}",""&f_obj("City"))
		end if
		if instr(s_Content,"{ME:FS_Address}")>0 then
			s_Content = replace(s_Content,"{ME:FS_Address}",""&f_obj("Address"))
		end if
		if instr(s_Content,"{ME:FS_PostCode}")>0 then
			s_Content = replace(s_Content,"{ME:FS_PostCode}",""&f_obj("PostCode"))
		end if
		if instr(s_Content,"{ME:FS_Vocation}")>0 then
			s_Content = replace(s_Content,"{ME:FS_Vocation}",""&f_obj("Vocation"))
		end if
		if instr(s_Content,"{ME:FS_BothYear}")>0 then
			s_Content = replace(s_Content,"{ME:FS_BothYear}",""&f_obj("BothYear"))
		end if
		if instr(s_Content,"{ME:FS_Age}")>0 then
			s_Content = replace(s_Content,"{ME:FS_Age}",""&datediff("yyyy",f_obj("BothYear"),date())&"岁")
		end if
		if instr(s_Content,"{ME:FS_Integral}")>0 then
			s_Content = replace(s_Content,"{ME:FS_Integral}",""&f_obj("Integral"))
		end if
		if instr(s_Content,"{ME:FS_FS_Money}")>0 then
			s_Content = replace(s_Content,"{ME:FS_FS_Money}",""&f_obj("FS_Money"))
		end if
		if instr(s_Content,"{ME:FS_IsMarray}")>0 then
			s_Content = replace(s_Content,"{ME:FS_IsMarray}",""&Replacestr(f_obj("IsMarray"),"1:已婚,2:未婚"))
		end if
		
		if instr(s_Content,"{ME:FS_UserURL}")>0 then
			s_Content = replace(s_Content,"{ME:FS_UserURL}","http://"& Request.Cookies("FoosunMFCookies")("FoosunMFDomain") &"/"&G_USER_DIR)
		end if
		
		if instr(s_Content,"{ME:FS_RegTime}")>0 then
			dim tmp_f_datestyle
			tmp_f_datestyle = f_datestyle
			if instr(f_datestyle,"YY02")>0 then
				tmp_f_datestyle= replace(tmp_f_datestyle,"YY02",right(year(f_obj("RegTime")),2))
			end if
			if instr(f_datestyle,"YY04")>0 then
				tmp_f_datestyle= replace(tmp_f_datestyle,"YY04",year(f_obj("RegTime")))
			end if
			if instr(f_datestyle,"MM")>0 then
				if month(f_obj("RegTime"))<10 then
					tmp_f_datestyle= replace(tmp_f_datestyle,"MM","0"&month(f_obj("RegTime")))
				else
					tmp_f_datestyle= replace(tmp_f_datestyle,"MM",month(f_obj("RegTime")))
				end if
			end if
			if instr(f_datestyle,"DD")>0 then
				if day(f_obj("RegTime"))<10 then
					tmp_f_datestyle= replace(tmp_f_datestyle,"DD","0"&day(f_obj("RegTime")))
				else

					tmp_f_datestyle= replace(tmp_f_datestyle,"DD",day(f_obj("RegTime")))
				end if
			end if
			if instr(f_datestyle,"HH")>0 then
				if hour(f_obj("RegTime"))<10 then
					tmp_f_datestyle= replace(tmp_f_datestyle,"HH","0"&hour(f_obj("RegTime")))
				else
					tmp_f_datestyle= replace(tmp_f_datestyle,"HH",hour(f_obj("RegTime")))
				end if
			end if
			if instr(f_datestyle,"MI")>0 then
				if minute(f_obj("RegTime"))<10 then
					tmp_f_datestyle= replace(tmp_f_datestyle,"MI","0"&minute(f_obj("RegTime")))
				else
					tmp_f_datestyle= replace(tmp_f_datestyle,"MI",minute(f_obj("RegTime")))
				end if
			end if
			if instr(f_datestyle,"SS")>0 then
				if second(f_obj("RegTime"))<10 then
					tmp_f_datestyle= replace(tmp_f_datestyle,"SS","0"&second(f_obj("RegTime")))
				else
					tmp_f_datestyle= replace(tmp_f_datestyle,"SS",second(f_obj("RegTime")))
				end if
			end if
			s_Content = replace(s_Content,"{ME:FS_RegTime}",""&tmp_f_datestyle&"")
		end if
		if instr(s_Content,"{ME:FS_hits}")>0 then
			s_Content = replace(s_Content,"{ME:FS_hits}",""&f_obj("hits"))
		end if
		if IsCorporation=9 then
			set f_obj1 = User_Conn.execute("select * from FS_ME_CorpUser where UserNumber='"&f_obj("UserNumber")&"'")
			s_Content = getlist_sd1(f_obj1,s_Content,f_titlenumber,f_contentnumber,f_datestyle)
			f_obj1.close:set f_obj1 = nothing
		elseif IsCorporation=1 then
			s_Content = getlist_sd1(f_obj,s_Content,f_titlenumber,f_contentnumber,f_datestyle)	
		end if	
		getlist_sd = s_Content
	End Function
	''企业用户的替换
	Function getlist_sd1(f_obj,s_Content,f_titlenumber,f_contentnumber,f_datestyle)
		if instr(s_Content,"{ME:FS_C_Name}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_Name}",""&GotTopic(f_obj("C_Name"),f_titlenumber))
			else
				s_Content = replace(s_Content,"{ME:FS_C_Name}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_ShortName}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_ShortName}",""&GotTopic(f_obj("C_ShortName"),f_titlenumber))
			else
				s_Content = replace(s_Content,"{ME:FS_C_ShortName}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_logo}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_logo}",""&f_obj("C_logo"))
			else
				s_Content = replace(s_Content,"{ME:FS_C_logo}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_Tel}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_Tel}",""&f_obj("C_Tel"))
			else
				s_Content = replace(s_Content,"{ME:FS_C_Tel}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_Fax}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_Fax}",""&f_obj("C_Fax"))
			else
				s_Content = replace(s_Content,"{ME:FS_C_Fax}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_VocationClassID}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_VocationClassID}",""&f_obj("C_VocationClassID"))
			else
				s_Content = replace(s_Content,"{ME:FS_C_VocationClassID}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_WebSite}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_WebSite}",""&Replacestr(f_obj("C_WebSite"),"<a href="""&f_obj("C_WebSite")&""" target=_blank>"&f_obj("C_WebSite")&"</a>"))
			else
				s_Content = replace(s_Content,"{ME:FS_C_WebSite}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_Operation}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_Operation}",""&f_obj("C_Operation"))
			else
				s_Content = replace(s_Content,"{ME:FS_C_Operation}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_Products}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_Products}",""&f_obj("C_Products"))
			else
				s_Content = replace(s_Content,"{ME:FS_C_Products}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_Content}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_Content}",""&GotTopic(f_obj("C_Content"),f_contentnumber))
			else
				s_Content = replace(s_Content,"{ME:FS_C_Content}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_Province}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_Province}",""&f_obj("C_Province"))
			else
				s_Content = replace(s_Content,"{ME:FS_C_Province}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_City}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_City}",""&f_obj("C_City"))
			else
				s_Content = replace(s_Content,"{ME:FS_C_City}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_Address}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_Address}",""&f_obj("C_Address"))
			else
				s_Content = replace(s_Content,"{ME:FS_C_Address}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_PostCode}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_PostCode}",""&f_obj("C_PostCode"))
			else
				s_Content = replace(s_Content,"{ME:FS_C_PostCode}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_Vocation}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_Vocation}",""&f_obj("C_Vocation"))
			else
				s_Content = replace(s_Content,"{ME:FS_C_Vocation}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_BankName}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_BankName}",""&f_obj("C_BankName"))
			else
				s_Content = replace(s_Content,"{ME:FS_C_BankName}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_BankUserName}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_BankUserName}",""&f_obj("C_BankUserName"))
			else
				s_Content = replace(s_Content,"{ME:FS_C_BankUserName}","")
			end if		
		end if
		if instr(s_Content,"{ME:FS_C_property}")>0 then
			if not f_obj.eof then
				s_Content = replace(s_Content,"{ME:FS_C_property}",""&f_obj("C_property"))
			else
				s_Content = replace(s_Content,"{ME:FS_C_property}","")
			end if		
		end if
		getlist_sd1 = s_Content
	End Function
End Class
%>
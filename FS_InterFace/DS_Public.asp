<%
Class cls_DS
	Private m_Rs,m_FSO,m_Dict
	Private m_PathDir,m_Path_UserDir,m_Path_User,m_Path_adminDir,m_Path_UserPageDir,m_Path_Templet
	Private m_Err_Info,m_Err_NO

	Public Property Get Err_Info()
		Err_Info = m_Err_Info
	End Property

	Public Property Get Err_NO()
		Err_NO = m_Err_NO
	End Property

	Private Sub Class_initialize()
		Set m_Rs = Server.CreateObject(G_FS_RS)
		Set m_FSO = Server.CreateObject(G_FS_FSO)
		Set m_Dict = Server.CreateObject(G_FS_DICT)
		m_PathDir = replace("/"&G_VIRTUAL_ROOT_DIR&"/","//","/")
		m_Path_UserDir = replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_USER_DIR&"/","//","/")
		m_Path_UserPageDir = replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_USERFILES_DIR&"/","//","/")
	End Sub

	Private Sub Class_Terminate()
		Set m_Rs = Nothing
		Set m_FSO = Nothing
		Set m_Dict = Nothing
	End Sub

	Public Function get_LableChar(f_Lable,f_Id,f_RefreshPageType)
		select case LCase(f_Lable.LableFun)
			case "classnews","lastnews","recnews","hotnews","downhotnews","downpicnews"
				get_LableChar=ClassNews(f_Lable,LCase(f_Lable.LableFun),f_Id)
			case "classlist"
				get_LableChar=classlist(f_Lable,"classlist",f_Id)
			case "specialdown"v
				get_LableChar=classNewsDown(f_Lable,"specialdown",f_Id)
			case "speciallistdown"
				get_LableChar=speciallistdown(f_Lable,"speciallist",f_Id)
			case "readnews"
				get_LableChar=ReadNews(f_Lable,"readnews",f_Id)
			case "sitemap"
				get_LableChar=SiteMap(f_Lable,"sitemap",f_Id)
			case "search"
				get_LableChar=Search(f_Lable,"search")
			case "infostat"
				get_LableChar=infoStat(f_Lable,"infostat")
			case "classnavi"
				get_LableChar=ClassNavi(f_Lable,"classnavi",f_Id)
			case "subclasslist"
				get_LableChar=subClassList(f_Lable,"subclasslist",f_Id)
			case "specialnavi"
				get_LableChar=SpecialNavi(f_Lable,"specialnavi",f_Id)
			case "specialcode"
				get_LableChar=SpecialCode(f_Lable,"specialcode",f_Id)
			Case "down_relative"
				get_LableChar=downrelative(f_Lable,"down_relative",f_Id)	
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
	'得到div,table_________________________________________________________________________________
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
			table_="<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
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
	'DIV格式输出结束_____________________________________________________________________

	'开始读取标签____综合类标签__________________________________________________
	Public Function ClassNews(f_Lable,f_LableType,f_Id)
		Dim classNews_head,classNews_bottom,classNews_middle1,classNews_middle2,div_tf,style_Content,Content_List,f_rs_s_obj,Content_more,c_i_k
		Dim newnumber,classid,orderby,orderdesc,colnumber,contentnumber,navinumber,datenumber,titlenumber,picshowtf,datenumber_tmp,morechar,datestyle,openstyle,open_target,containSubClass,childClass,datenumber_tmp_
		Dim f_sql,f_rs_obj,f_rs_configobj,f_configSql,CharIndexStr
		Dim MF_Domain,marqueedirec,marqueespeed,marqueestyle,search_str,order_in_str_awen,RefreshNumber
		Dim ClassName,ClassEName,c_Domain,ClassNaviContent,ClassNaviPic,c_SavePath,c_FileSaveType,search_inSQL
		search_str = f_Lable.LablePara("栏目")
		newnumber =f_Lable.LablePara("Loop")
		datenumber= f_Lable.LablePara("多少天")
		titlenumber= f_Lable.LablePara("标题数")
		picshowtf=f_Lable.LablePara("图文标志")
		openstyle=f_Lable.LablePara("打开窗口")
		containSubClass=f_Lable.LablePara("包含子类")
		orderby = f_Lable.LablePara("排列字段")
		orderdesc = f_Lable.LablePara("排列方式")
		morechar = f_Lable.LablePara("更多连接")
		datestyle = f_Lable.LablePara("日期格式")
		colnumber= f_Lable.LablePara("排列数")
		contentnumber= f_Lable.LablePara("内容字数")
		if f_Lable.LablePara("输出格式") = "out_DIV" then
			div_tf=1
		else
			div_tf=0
		end if
		if not isnumeric(titlenumber) then
			titlenumber = 30
		else
			titlenumber = titlenumber
		end if
		if isnumeric(colnumber)=false then
			colnumber = 30
		else
			colnumber = colnumber
		end if
		if isnumeric(contentnumber)=false then
			contentnumber = 100
		else
			contentnumber = contentnumber
		end if
		if right(lcase(morechar),4)=".jpg" or right(lcase(morechar),4)=".gif" or right(lcase(morechar),4)=".png" or right(lcase(morechar),4)=".ico" or right(lcase(morechar),4)=".bmp" or right(lcase(morechar),5)=".jpeg" then
			morechar = "<img src = "&morechar&" border=""0"" />"
		end if
		
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
		f_DivID = f_Lable.LablePara("DivID")
		f_DivClass = f_Lable.LablePara("Divclass")
		f_UlID = f_Lable.LablePara("ulid")
		f_ULClass = f_Lable.LablePara("ulclass")
		f_LiID = f_Lable.LablePara("liid")
		f_LiClass = f_Lable.LablePara("liclass")
		classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		''永不过期和尚未到期的
		 If G_IS_SQL_DB=1 Then
		datenumber_tmp_="datediff(day,AddTime,'"&date()&"')"
		else
		datenumber_tmp_="datediff('d',AddTime,'"&date()&"')"
		end if  
		
		datenumber_tmp=" and  (OverDue=0 or (OverDue>0 and "&datenumber_tmp_&"<= OverDue))"
		If datenumber ="0" Then
			datenumber_tmp = ""
		Else
			If G_IS_SQL_DB=1 Then
				datenumber_tmp = " and dateadd(d,"&datenumber&",AddTime)>=getdate()"
			Else
				datenumber_tmp = " and dateadd('d',"&datenumber&",AddTime)>=now"
			End If
		End If
		''已经在外部定义好了。
		order_in_str_awen = ""&orderby&" "&orderdesc&""
		if ucase(orderby)<>"ID" then order_in_str_awen = order_in_str_awen & ",ID Desc"
		childClass = DelHeadAndEndDot(""&getNewsSubClass(search_str))
		If containSubClass=1 And childClass<>"" Then
			childClass=" or classid in ('"& FormatStrArr(childClass) &"')"
		Else
			childClass=""
		End if
		select case f_LableType
			case "classnews"
				If childClass<>"" Then
					search_inSQL = " and (ClassId='"& search_str &"'"&childClass&")"
				Else
					search_inSQL = " and ClassId='"& search_str &"'"
				End if
			case "specialproducts"
				if search_str<>"" then
					search_inSQL = " and SpecialID="&search_str
				ELse
					search_inSQL = " and SpecialID<>0"
				End if
				'----------------------------------------
					if f_type="speciallist" then
				dim specaial_rs,f_spen,Pro_SpecialID
				If IsNumeric(f_Id) Then
					Pro_SpecialID=f_Id
				Else
					Pro_SpecialID=0
				End If
				set specaial_rs = Conn.execute("select SpecialEName From FS_DS_Special where SpecialID="&Pro_SpecialID&"")
				if Not specaial_rs.eof then
					f_spen = specaial_rs(0)
				end if
				specaial_rs.close:set specaial_rs = nothing
				if G_IS_SQL_DB=0 then
					search_inSQL = " and instr(SpecialEName,'"&f_spen&"')>0"
				else
					search_inSQL = " and charindex('"&f_spen&"',SpecialEName)>0"
				end if
				RefreshNumber = ""
			else
				set rs_c=conn.execute("select RefreshNumber from FS_DS_Class where  ClassId='"& f_Id &"'")
				if rs_c.eof then
						RefreshNumber = ""
					 rs_c.close:set rs_c = nothing
				else
					if rs_c(0)=0 then
						RefreshNumber = ""
					else
						RefreshNumber = "top "&rs_c(0)&""
					end if
					rs_c.close:Set rs_c=nothing
				end if
				search_inSQL=" and ClassId='"& f_Id &"'"
			end if
			'------------------------------------------------
				if G_IS_SQL_DB=0 then
					if datenumber ="0" then:datenumber_tmp = "":else:datenumber_tmp = " and datevalue(addtime)+"&datenumber&">=datevalue(now)":end if
				Elseif G_IS_SQL_DB=1 then
					if datenumber ="0" then:datenumber_tmp = "":else:datenumber_tmp = " and dateadd(d,"&datenumber&",addtime)>='"&datevalue(now())&"'":end if
				End if
			case "lastnews","hotnews","downhotnews"
				if trim(search_str)<>"" Then
					If childClass<>"" Then
						search_inSQL = " and (ClassId='"& search_str &"'"&childClass&")"
					Else
						search_inSQL = " and ClassId='"& search_str &"'"
					End if
				else
					if trim(f_id)<>"" Then
						childClass = DelHeadAndEndDot(""&getNewsSubClass(f_id))
						If containSubClass = 1 And childClass <> "" Then
							childClass = " Or ClassID in ("&childClass&")"
						Else
							childClass = ""
						End If
						If childClass<>"" then
							search_inSQL = " and (ClassId='"& f_id &"'"&childClass&")"
						Else
							search_inSQL = " and ClassId='"& f_id &"'"
						End if
					else
						search_inSQL = ""
					end if
				end if
				f_LableType = "classnews" ''处理后还原
			case "recnews"
				if trim(search_str)<>"" Then
					If childClass<>"" then
						search_inSQL = " and RecTF=1 and (ClassId='"& search_str &"'"&childClass&")"
					Else
						search_inSQL = " and RecTF=1 and ClassId='"& search_str &"'"
					End if
				else
					if trim(f_id)<>"" Then
						childClass = DelHeadAndEndDot(""&getNewsSubClass(f_id))
						If containSubClass = 1 And childClass <> "" Then
							childClass = " Or ClassID in ("&childClass&")"
						Else
							childClass = ""
						End If	
						If childClass<>"" then
							search_inSQL = " and RecTF=1 and (ClassId='"& f_id &"'"&childClass&")"
						Else
							search_inSQL = " and RecTF=1 and ClassId='"& f_id &"'"
						End if
					else
						search_inSQL = " and RecTF=1"
					end if
				end if
				f_LableType = "classnews" ''处理后还原
			case "downpicnews"
				if trim(search_str)<>"" Then
					If childClass<>"" then
						search_inSQL = " and (ClassId='"& search_str &"'"&childClass&") and pic<>''"
					Else
						search_inSQL = " and ClassId='"& search_str &"' and pic<>'' " 
					End if
				else
					if trim(f_id)<>"" Then
						If childClass<>"" then
							search_inSQL = " and  pic<>'' and (ClassId='"& f_id &"'"&childClass&")"
						Else
							search_inSQL = " and  pic<>'' and ClassId='"& f_id &"'"
						End if
					else
						search_inSQL = " and  pic<>''"
					end if
				end if
				f_LableType = "classnews"
		end select
		f_sql="select top "& newnumber &" ID,DownLoadID,ClassID,Description,Accredit,AddTime,Appraise,AuditTF,BrowPop," _
			&"ClickNum,EditTime,EMail,SavePath,FileExtName,FileName,FileSize,[Language],Name,NewsTemplet,PassWord,Pic," _
			&"Property,Provider,ProviderUrl,RecTF,ReviewTF,ShowReviewTF,SystemType,Types,Version,OverDue,ConsumeNum,Hits"
		f_sql = f_sql &" From FS_DS_List where AuditTF=1 "&search_inSQL & datenumber_tmp &" order by "&order_in_str_awen
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		Content_List=""
		set f_rs_obj = Conn.execute(f_sql)
		if f_rs_obj.eof then
			Content_List=""
			f_rs_obj.close:set f_rs_obj=nothing
		else
			if f_LableType="marnews" then
					Content_List = Content_List & "<marquee onmouseover=""this.stop();"" scrollamount="""& marqueespeed &""" direction="""& marqueedirec &""" onmouseout=""this.start();"">"
					do while not f_rs_obj.eof
						Content_List= Content_List &getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,f_LableType)&"&nbsp;&nbsp;&nbsp;"
						f_rs_obj.movenext
					loop
					Content_List = Content_List & "</marquee>"
			else
					if div_tf = 0 then
						c_i_k = 0
						if cint(colnumber)<>1 then
							Content_List = Content_List &  "  <tr>"
						end if
					end if
					do while not f_rs_obj.eof
							if div_tf=1 then
								Content_List= Content_List & classNews_middle1 & getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,f_LableType) & classNews_middle2
							else
								if cint(colnumber) =1 then
									Content_List= Content_List & chr(10)&"   <tr><td>" & getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,f_LableType) & "</td></tr>"
								else
									Content_List= Content_List & "<td width="""& cint(100/cint(colnumber))&"%"">" & getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,f_LableType) & "</td>"
								end if
							end if
							f_rs_obj.movenext
							if div_tf = 0 then
								if cint(colnumber)<>1 then
								c_i_k = c_i_k+1
									if c_i_k mod cint(colnumber) = 0 then
										Content_List = Content_List & "</tr>"&chr(10)&"  <tr>"
									end if
								end if
							end if
					   loop
					   if div_tf=0 then
							if cint(colnumber)<>1 then
								Content_List = Content_List & "</tr>"&chr(10)
							end if
					   end if
			end if
			'得到栏目路径
			if f_LableType="classnews" then
					dim Query_rs,newsclass_SavePath,FileSaveType,UrlDomain,all_savepath
					set Query_rs=Conn.execute("select ClassEName,SavePath,FileExtName,[Domain],FileSaveType,IsURL,UrlAddress From FS_DS_Class where ClassId='"& search_str &"'")
					if Query_rs.eof then
							Query_rs.close:set Query_rs=nothing
					else
						  all_savepath = get_ClassLink(search_str)
						  Query_rs.close:set Query_rs=nothing
					end if
					if openstyle=0 then
						open_target=" "
					else
						open_target=" target=""_blank"""
					end if
					if morechar<>"" then
						if div_tf=1 then
							Content_more = "  <li><a href="""&  all_savepath &""" "& open_target &" title=""更多..."">"&morechar&"</a></li>"&chr(10)
						else
							Content_more = "  <tr><td colspan="""& cint(colnumber) &"""><div align=""right""><a href="""&  all_savepath &""" "& open_target &" title=""更多..."">"&morechar&"</a></div></td></tr>"&chr(10)
						end if
					end if
					'-------------------
					elseif f_LableType="specialproducts" then
					dim special_rs,special_Path
					set special_rs =Conn.execute("select SpecialID,SpecialCName,SpecialEName,naviText,SavePath,FileExtName,isLock,naviPic from FS_DS_Special where isLock=0 and SpecialEName='"&trim(search_str)&"'")
					if Not special_rs.eof then
							if LinkType=1 then
								'Mf_Domain,IsDomain
									if trim(IsDomain)<>"" then
										m_PathDir = IsDomain & m_PathDir
									else
										m_PathDir = Mf_Domain & m_PathDir
									end if
									special_Path = "http://"&replace(m_PathDir & special_rs("SavePath")&"/special_"&special_rs("SpecialEName")&"."&special_rs("FileExtName"),"//","/")
							else
									m_PathDir = m_PathDir
									special_Path = replace(m_PathDir & special_rs("SavePath")&"/special_"&special_rs("SpecialEName")&"."&special_rs("FileExtName"),"//","/")
							end if
							special_rs.close:set special_rs=nothing
					end if
					if morechar<>"" then
						if div_tf=1 then
								Content_more = "  <li><a href="""&  special_Path &""" "& open_target &" title=""更多..."">"&morechar&"</a></li>"&chr(10)
						else
								Content_more = "  <tr><td colspan="""& cint(colnumber) &"""><div align=""right""><a href="""&  special_Path &""" "& open_target &" title=""更多..."">"&morechar&"</a></div></td></tr>"&chr(10)
						end if
					end if
					'--------------------------
			 end if
			 f_rs_obj.close:set f_rs_obj=nothing
			if f_LableType="marnews" then
				Content_List= Content_List
			else
				Content_List=classNews_head & Content_List & Content_more & classNews_bottom
			end if
		end if
		f_Lable.DictLableContent.Add "1",Content_List
	End Function
	'专区列表
	Public function classNewsDown(f_Lable,f_type,f_Id)
		dim div_tf,specialId,TitleNumber,datenumber_tmp,MF_Domain
		dim daynum,lefttitle,pictf,openstyle,OrderBy,OrderDesc,MoreStr,DateStyle,ColsNumber,ContentNumber
		specialId = f_Lable.LablePara("专题")
		TitleNumber = f_Lable.LablePara("Loop")
		daynum = f_Lable.LablePara("多少天")
		lefttitle = f_Lable.LablePara("标题数")
		pictf = f_Lable.LablePara("图文标志")
		openstyle = f_Lable.LablePara("打开窗口")
		OrderBy = f_Lable.LablePara("排列字段")
		OrderDesc = f_Lable.LablePara("排列方式")
		MoreStr = f_Lable.LablePara("更多连接")
		DateStyle = f_Lable.LablePara("日期格式")
		ColsNumber = f_Lable.LablePara("排列数")
		ContentNumber = f_Lable.LablePara("内容字数")
		if f_Lable.LablePara("输出格式") = "out_DIV" then
			div_tf=1
		else
			div_tf=0
		end if
		if not isnumeric(TitleNumber) then:TitleNumber = 10:else:TitleNumber = cint(TitleNumber):end if
		if isnumeric(lefttitle)=false then:lefttitle = 30:else:lefttitle = cint(lefttitle):end if
		if isnumeric(ContentNumber)=false then:ContentNumber = 100:else:ContentNumber = cint(ContentNumber):end if
		if right(lcase(MoreStr),4)=".jpg" or right(lcase(MoreStr),4)=".gif" or right(lcase(MoreStr),4)=".png" or right(lcase(MoreStr),4)=".ico" or right(lcase(MoreStr),4)=".bmp" or right(lcase(MoreStr),5)=".jpeg" then MoreStr = "<img src = "&MoreStr&" border=""0"" />"
		dim classNews_head,classNews_middle1,classNews_bottom,classNews_middle2
		
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
		f_DivID = f_Lable.LablePara("DivID")
		f_DivClass = f_Lable.LablePara("Divclass")
		f_UlID = f_Lable.LablePara("ulid")
		f_ULClass = f_Lable.LablePara("ulclass")
		f_LiID = f_Lable.LablePara("liid")
		f_LiClass = f_Lable.LablePara("liclass")
		classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		
		if G_IS_SQL_DB=0 then
			if daynum ="0" then:datenumber_tmp = "":else:datenumber_tmp = " and datevalue(AddTime)+"&clng(daynum)&">=datevalue(now)":end if
		Elseif G_IS_SQL_DB=1 then
			if daynum ="0" then:datenumber_tmp = "":else:datenumber_tmp = " and dateadd(d,"&clng(daynum)&",AddTime)>='"&datevalue(now())&"'":end if
		End if
		if pictf = "1" then:pictf = 1:else:pictf = 0:end if
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		dim rs,f_sql,i_d,down_more
		f_sql="select top "& cint(TitleNumber) &" ID,DownLoadID,ClassID,Description,Accredit,AddTime,Appraise,AuditTF,BrowPop," _
			&"ClickNum,EditTime,EMail,SavePath,FileExtName,FileName,FileSize,[Language],Name,NewsTemplet,PassWord,Pic," _
			&"Property,Provider,ProviderUrl,RecTF,ReviewTF,ShowReviewTF,SystemType,Types,Version,OverDue,ConsumeNum,Hits,speicalId"
		f_sql = f_sql &" From FS_DS_List where AuditTF=1 and speicalId="& specialId &"" & datenumber_tmp &" order by AddTime desc,"&OrderBy &" "& OrderDesc &""
		i_d = 0 
		set Rs = Conn.execute(f_sql)
		if Rs.eof then
			Rs.close:set Rs = nothing
			classNewsDown = ""
		else 
			if div_tf = 0 then
				classNewsDown = classNewsDown & "<tr>"
			end if
			do while not Rs.eof
				if div_tf = 1 then
					classNewsDown= classNewsDown & classNews_middle1 & getlist_news(rs,f_Lable.FoosunStyle.StyleContent,lefttitle,ContentNumber,200,pictf,DateStyle,openstyle,MF_Domain,"classNewsDown") & classNews_middle2
				else
					classNewsDown = classNewsDown & "<td width="""& cint(100/cint(ColsNumber)) &"%"">"& getlist_news(rs,f_Lable.FoosunStyle.StyleContent,lefttitle,ContentNumber,200,pictf,DateStyle,openstyle,MF_Domain,"classNewsDown") &"</td>"
				end if
				Rs.movenext
				i_d = i_d + 1
				if div_tf = 0 then
					if i_d mod cint(ColsNumber) =0 then
						classNewsDown = classNewsDown & "</tr>"&chr(10)&"<tr>"
					end if
				end if
			loop
			Dim sRs,sSpecialEName
			set sRs = Conn.execute("select SpecialEName from FS_DS_Special where SpecialID = "& specialId )
			if Not sRs.eof then
				sSpecialEName = sRs(0)
				sRs.close:set sRs=nothing
			else
				sSpecialEName = "0"
				sRs.close:set sRs=nothing
			end if
			if openstyle ="1" then:openstyle = " target=""_blank""":else:openstyle = " target=""_self""":end if
			if MoreStr<>"" then
				down_more = "<div align=""right""><a href="""& get_specialLink(sSpecialEName) &""""& openstyle &" title=""更多..."">"& MoreStr &"</a></div>"
			end if
			if div_tf = 0 then
				classNewsDown = classNews_head & classNewsDown & classNews_bottom & down_more
			else
				classNewsDown = classNewsDown & down_more
			end if
			Rs.close:set Rs = nothing
		end if
		f_Lable.DictLableContent.Add "1",classNewsDown
	End Function
	''正则查找
	Public Function Test_KeyWord(f_Str,Patt)
		Dim f_regEx
		Set f_regEx = New RegExp
		f_regEx.Pattern = Patt
		f_regEx.IgnoreCase = True
		f_regEx.Global = True
		Test_KeyWord = f_regEx.test(Cstr(f_Str))
	End Function

	Public Function replace_KeyWord(f_obj,f_MF_Domain,f_Str,Patstr)
		Dim regEx,Match, Matches,f_oldStr
		f_oldStr = f_Str
		Set regEx = New RegExp
		regEx.Pattern = Patstr
		regEx.IgnoreCase = True
		regEx.Global = True
		Set Matches = regEx.Execute(Cstr(f_Str))      ' 执行搜索。
		'Show_Html = Matches(Matches.Count-1).value   'Matches.Count-1 才是最后一个数值
		For Each Match in Matches         ' 遍历 Matches 集合。
			if Match.Value<>"" then replace_KeyWord = replace_KeyWord & doregExValue(f_obj,f_MF_Domain,f_oldStr,Match.Value)
		Next
	End Function

	Public Function doregExValue(f_obj,f_MF_Domain,f_oldStr,ExValue)
		dim s_m_Rs1,k_tmp_Chararray,k_tmp_Chararray1,k_tmp_uchar,tmp_list,k_tmp_uchar1,k_tmp_uchar2,k_tmp_i
		k_tmp_Chararray = split(ExValue,"$")
		k_tmp_uchar="":tmp_list="":k_tmp_uchar1="":k_tmp_uchar2=""
		if ubound(k_tmp_Chararray)=1 then
			k_tmp_uchar =  replace(k_tmp_Chararray(1),"}","")
			k_tmp_uchar=Replace(k_tmp_uchar,"&amp;","&")
			k_tmp_uchar=Replace(k_tmp_uchar,"&lt;","<")
			k_tmp_uchar=Replace(k_tmp_uchar,"&gt;",">")
			select case k_tmp_uchar
				case "1"
				k_tmp_uchar = "<br />"
				case "2"
				k_tmp_uchar = "&nbsp;"
				case "3"
				k_tmp_uchar = "</tr><tr>"
				case else
				'' $<br />:10 表示 <br />循环10次
				if instr(k_tmp_uchar,":") then
					k_tmp_Chararray1 = split(k_tmp_uchar,":")
					if ubound(k_tmp_Chararray1) = 1 then
						if k_tmp_Chararray1(1)<>"" then
							if isnumeric(k_tmp_Chararray1(1)) then
								k_tmp_uchar = ""
								for k_tmp_i=1 to cint(k_tmp_Chararray1(1))
									select case k_tmp_Chararray1(0)
										case "1"
										k_tmp_uchar2 = "<br />"
										case "2"
										k_tmp_uchar2 = "&nbsp;"
										case "3"
										k_tmp_uchar2 = "</tr><tr>"
										case else
										k_tmp_uchar2 = k_tmp_Chararray1(0)
									end select
									k_tmp_uchar = k_tmp_uchar & k_tmp_uchar2
								next
							end if
						end if
					end if
				end if
			end select
		end if
		set s_m_Rs1 = Conn.execute("select ID,AddressName from FS_DS_Address where DownLoadID='"&f_obj("DownLoadID")&"' order by Number asc")
		do while not s_m_Rs1.eof
			tmp_list = tmp_list & "<a href=""http://"&f_MF_Domain&"/Down.asp?DownLoadID="&f_obj("DownLoadID")&"&ID="&s_m_Rs1("ID")&""" target=""_blank"">"& s_m_Rs1("AddressName") &"</a>"
			tmp_list = tmp_list & k_tmp_uchar
			s_m_Rs1.movenext
		loop
		s_m_Rs1.close
		set s_m_Rs1=nothing
		doregExValue = replace(f_oldStr,ExValue,tmp_list)
	End Function
	
	'专区终极列表
	Public function speciallistdown(f_Lable,f_type,f_Id)
		if not isnumeric(f_Id) then
			f_Lable.DictLableContent.Add "1","标签错误，请不要非专区模板的标签插入下载专区标签"
		else
			Dim div_tf,TitleNumber,divid,divclass,ulid,ulclass,liid,liclass,daynum,lefttitle,PicTF,OpenStyle,OrderBy,OrderDesc,PageTF,PageStyle
			Dim PageNum,PageCSS,DateStyle,ColsNumber,ContentNumber,Page_flag_TF,datenumber_tmp,MF_Domain,f_rs_obj,f_sql,i_d,cl_i
			Dim f_TableName,f_SelectFieldNames,f_PageIndex,f_Where,f_PaginationStr,f_NewsContent,f_IDSRS,f_IDSArray,f_IDS,i,f_BeginIndex,pagenumber
			TitleNumber = f_Lable.LablePara("Loop")
			daynum = f_Lable.LablePara("多少天")
			lefttitle = f_Lable.LablePara("标题数")
			PicTF = f_Lable.LablePara("图文标志")
			OpenStyle =f_Lable.LablePara("打开窗口")
			OrderBy =f_Lable.LablePara("排列字段")
			OrderDesc =f_Lable.LablePara("排列方式")
			PageTF=f_Lable.LablePara("分页")
			PageStyle=f_Lable.LablePara("分页样式")
			PageNum=f_Lable.LablePara("每页数量")
			PageCSS=f_Lable.LablePara("PageCSS")
			DateStyle=f_Lable.LablePara("日期格式")
			ColsNumber=f_Lable.LablePara("排列数")
			ContentNumber=f_Lable.LablePara("内容字数")
			pagenumber = f_Lable.LablePara("每页数量")
			if pagenumber = "" OR Not IsNumeric(pagenumber) then pagenumber = 10
			if f_Lable.LablePara("输出格式") = "out_DIV" then
				div_tf=1
			else
				div_tf=0
			end if
			if not isnumeric(TitleNumber) then:TitleNumber = 10:else:TitleNumber = cint(TitleNumber):end if
			if isnumeric(lefttitle)=false then:lefttitle = 30:else:lefttitle = cint(lefttitle):end if
			if isnumeric(ContentNumber)=false then:ContentNumber = 100:else:ContentNumber = cint(ContentNumber):end if
			dim classNews_head,classNews_middle1,classNews_bottom,classNews_middle2
			
			Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
			f_DivID = f_Lable.LablePara("DivID")
			f_DivClass = f_Lable.LablePara("Divclass")
			f_UlID = f_Lable.LablePara("ulid")
			f_ULClass = f_Lable.LablePara("ulclass")
			f_LiID = f_Lable.LablePara("liid")
			f_LiClass = f_Lable.LablePara("liclass")
			classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			
			if G_IS_SQL_DB=0 then
				if daynum ="0" then:datenumber_tmp = "":else:datenumber_tmp = " and datevalue(AddTime)+"&clng(daynum)&">=datevalue(now)":end if
			Elseif G_IS_SQL_DB=1 then
				if daynum ="0" then:datenumber_tmp = "":else:datenumber_tmp = " and dateadd(d,"&clng(daynum)&",AddTime)>='"&datevalue(now())&"'":end if
			End if
			if pictf = "1" then:pictf = 1:else:pictf = 0:end if
			MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
			f_PaginationStr = PageStyle & "," & PageCSS
			f_Lable.DictLableContent.Item("0") = f_PaginationStr
			f_Where = "AuditTF=1 and speicalId="& f_id &"" & datenumber_tmp &" order by AddTime desc," & OrderBy & " " & OrderDesc
			f_TableName = "FS_DS_List"
			f_SelectFieldNames = "ID,DownLoadID,ClassID,Description,Accredit,AddTime,Appraise,AuditTF,BrowPop,ClickNum,EditTime,EMail,SavePath,FileExtName,FileName,FileSize,[Language],Name,NewsTemplet,PassWord,Pic,Property,Provider,ProviderUrl,RecTF,ReviewTF,ShowReviewTF,SystemType,Types,Version,OverDue,ConsumeNum,Hits,speicalId"
			
			Set f_IDSRS = Server.CreateObject(G_FS_RS)
			f_IDSRS.Open "Select ID from " & f_TableName & " Where " & f_Where,Conn,0,1
			if Not f_IDSRS.Eof then
				f_PageIndex = 1
				f_IDSArray = f_IDSRS.GetRows()
				Set  f_rs_obj = Server.CreateObject(G_FS_RS)
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
					f_Where = " ID In (" & f_IDS & ")"
					f_sql = "Select " & f_SelectFieldNames & " from " & f_TableName & " where " & f_Where
					f_rs_obj.open f_sql,Conn,0,1
					cl_i = 1
					speciallistdown = classNews_head
					Do While Not f_rs_obj.Eof
						f_NewsContent = getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,lefttitle,contentnumber,200,pictf,datestyle,openstyle,MF_Domain,"ClassList")
						if div_tf = 1 then
							speciallistdown= speciallistdown & classNews_middle1 & f_NewsContent & classNews_middle2
						else
							if cint(ColsNumber) = 1 then
								speciallistdown= speciallistdown &"   <tr><td>" & f_NewsContent & "</td></tr>"
							else
								if cl_i mod ColsNumber = 1 then speciallistdown= speciallistdown & "<tr>"
								speciallistdown= speciallistdown & "<td width="""& cint(100/cint(ColsNumber))&"%"">" & f_NewsContent & "</td>"
								if cl_i mod ColsNumber = 0 then speciallistdown= speciallistdown & "</tr>"
							end if
						end if
						f_rs_obj.movenext
						cl_i = cl_i + 1
					Loop
					speciallistdown = speciallistdown & classNews_bottom
					f_Lable.DictLableContent.Add f_PageIndex & "",speciallistdown
					f_rs_obj.Close
					f_PageIndex = f_PageIndex + 1
				Loop 
				Set f_rs_obj = Nothing
			else
				f_Lable.DictLableContent.Add "1",""
			end if
			f_IDSRS.Close
			Set f_IDSRS = Nothing
		end if
	End Function

	'替换样式列表________________________________________________________________
	Public Function getlist_news(f_obj,s_Content,f_titlenumber,f_contentnumber,f_navinumber,f_picshowtf,f_datestyle,f_openstyle,f_MF_Domain,f_subsys_ListType)
		Dim f_target,get_SpecialEName,ListSql,Rs_ListObj,s_NewsPathUrl,Rs_Authorobj,k_i,k_tmp_Char,k_tmp_uchar,k_tmp_Chararray,FormReview
		Dim s_m_Rs,s_m_Rs1,s_array,s_t_i,tmp_list,s_f_classSql,m_Rs_class,class_path,str_newstitle,get_SpecialID
		select case f_subsys_ListType
			case "classlist"
				get_SpecialID = f_obj("SpecialID")
			case "specialproducts"
				get_SpecialID = ""
			case else
		end select
		if f_openstyle=0 then
			f_target=" "
		else
			f_target=" target=""_blank"""
		end if
		if instr(s_Content,"{DS:FS_ID}")>0 then
			s_Content = replace(s_Content,"{DS:FS_ID}",f_obj("Id"))
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_DownLoadID}")>0 then
			s_Content = replace(s_Content,"{DS:FS_DownLoadID}",f_obj("DownLoadID"))
		end if
		'---------------------------------2/12---by chen ---下载完整标题--------------------
		if instr(s_Content,"{DS:FS_NameAll}")>0 then
			if f_subsys_ListType="readnews" then
				str_newstitle = f_obj("Name")
			else
				str_newstitle = Replace(Replace(Replace(Replace(Lose_Html(f_obj("Name"))," ","&nbsp;"),Chr(34),"&quot;"),">","&gt;"),"<","&lt;")
				if f_picshowtf=1 then
					if f_obj("Pic")<>"" then
						str_newstitle = ""&str_newstitle&"<img src="""&m_PathDir&"sys_images/img.gif"" alt=""有图片"" border=""0"">"
					end if
				end if
			end if
			s_Content = replace(s_Content,"{DS:FS_NameAll}",str_newstitle)
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_Name}")>0 then
			if f_subsys_ListType="readnews" then
				str_newstitle = f_obj("Name")
			else
				str_newstitle =GotTopic(f_obj("Name"),f_titlenumber)
				if f_picshowtf=1 then
					if f_obj("Pic")<>"" then
						str_newstitle = ""&str_newstitle&"<img src="""&m_PathDir&"sys_images/img.gif"" alt=""有图片"" border=""0"">"
					end if
				end if
			end if
			s_Content = replace(s_Content,"{DS:FS_Name}",str_newstitle)
		end if
		dim news_SavePath,s_Query_rs,news_Domain,news_UrlDomain,news_ClassEname,s_all_savepath
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_Description}")>0 then
			if f_subsys_ListType="readnews" then
				if instr(f_obj("Description"),"[FS:PAGE]")>0 then
					s_Content = replace(s_Content,"{DS:FS_Description}","[FS:CONTENT_START]"&f_obj("Description")&"[FS:CONTENT_END]")
				else
					s_Content = replace(s_Content,"{DS:FS_Description}",""&f_obj("Description"))
				end if
			else
				s_all_savepath = get_DownLink(f_obj("DownLoadID"))
				s_NewsPathUrl = s_all_savepath
				s_Content = replace(s_Content,"{DS:FS_Description}",replace(replace(GetCStrLen(""&replace(""&f_obj("Description"),"[FS:PAGE]","")&"",f_contentnumber),"&nbsp;",""),vbCrLf,"")&"...<a href="""& s_NewsPathUrl &""">详细内容</a>")
			end if
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_DownURL}")>0 or  instr(s_Content,"{DS:FS_Description}")>0 then
				s_all_savepath = get_DownLink(f_obj("DownLoadID"))
				s_NewsPathUrl = s_all_savepath
			s_Content = replace(s_Content,"{DS:FS_DownURL}",s_NewsPathUrl)
		end if
		'''下载地址列表
		'_________________________________________________________________________________________________
		if Test_KeyWord(s_Content,"\{DS\:FS_Address(\$[^{]+)?\}") then
			s_Content = replace_KeyWord(f_obj,f_MF_Domain,s_Content,"\{DS\:FS_Address(\$[^{]+)?\}")
		end if
		'_________________________________________________________________________________________________
		dim tmp_f_datestyle
		tmp_f_datestyle = f_datestyle
		if instr(s_Content,"{DS:FS_AddTime}")>0 then
			if instr(f_datestyle,"YY02")>0 then
				tmp_f_datestyle= replace(tmp_f_datestyle,"YY02",right(year(f_obj("AddTime")),2))
			end if
			if instr(f_datestyle,"YY04")>0 then
				tmp_f_datestyle= replace(tmp_f_datestyle,"YY04",year(f_obj("AddTime")))
			end if
			if instr(f_datestyle,"MM")>0 then
				if month(f_obj("AddTime"))<10 then
					tmp_f_datestyle= replace(tmp_f_datestyle,"MM","0"&month(f_obj("AddTime")))
				else
					tmp_f_datestyle= replace(tmp_f_datestyle,"MM",month(f_obj("AddTime")))
				end if
			end if
			if instr(f_datestyle,"DD")>0 then
				if day(f_obj("AddTime"))<10 then
					tmp_f_datestyle= replace(tmp_f_datestyle,"DD","0"&day(f_obj("AddTime")))
				else

					tmp_f_datestyle= replace(tmp_f_datestyle,"DD",day(f_obj("AddTime")))
				end if
			end if
			if instr(f_datestyle,"HH")>0 then
				if hour(f_obj("AddTime"))<10 then
					tmp_f_datestyle= replace(tmp_f_datestyle,"HH","0"&hour(f_obj("AddTime")))
				else
					tmp_f_datestyle= replace(tmp_f_datestyle,"HH",hour(f_obj("AddTime")))
				end if
			end if
			if instr(f_datestyle,"MI")>0 then
				if minute(f_obj("AddTime"))<10 then
					tmp_f_datestyle= replace(tmp_f_datestyle,"MI","0"&minute(f_obj("AddTime")))
				else
					tmp_f_datestyle= replace(tmp_f_datestyle,"MI",minute(f_obj("AddTime")))
				end if
			end if
			if instr(f_datestyle,"SS")>0 then
				if second(f_obj("AddTime"))<10 then
					tmp_f_datestyle= replace(tmp_f_datestyle,"SS","0"&second(f_obj("AddTime")))
				else
					tmp_f_datestyle= replace(tmp_f_datestyle,"SS",second(f_obj("AddTime")))
				end if
			end if
			s_Content = replace(s_Content,"{DS:FS_AddTime}",""&tmp_f_datestyle&"")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_EditTime}")>0 then
			if instr(f_datestyle,"YY02")>0 then
				tmp_f_datestyle= replace(tmp_f_datestyle,"YY02",right(year(f_obj("EditTime")),2))
			end if
			if instr(f_datestyle,"YY04")>0 then
				tmp_f_datestyle= replace(tmp_f_datestyle,"YY04",year(f_obj("EditTime")))
			end if
			if instr(f_datestyle,"MM")>0 then
				if month(f_obj("EditTime"))<10 then
					tmp_f_datestyle= replace(tmp_f_datestyle,"MM","0"&month(f_obj("EditTime")))
				else
					tmp_f_datestyle= replace(tmp_f_datestyle,"MM",month(f_obj("EditTime")))
				end if
			end if
			if instr(f_datestyle,"DD")>0 then
				if day(f_obj("EditTime"))<10 then
					tmp_f_datestyle= replace(tmp_f_datestyle,"DD","0"&day(f_obj("EditTime")))
				else

					tmp_f_datestyle= replace(tmp_f_datestyle,"DD",day(f_obj("EditTime")))
				end if
			end if
			if instr(f_datestyle,"HH")>0 then
				if hour(f_obj("EditTime"))<10 then
					tmp_f_datestyle= replace(tmp_f_datestyle,"HH","0"&hour(f_obj("EditTime")))
				else
					tmp_f_datestyle= replace(tmp_f_datestyle,"HH",hour(f_obj("EditTime")))
				end if
			end if
			if instr(f_datestyle,"MI")>0 then
				if minute(f_obj("EditTime"))<10 then
					tmp_f_datestyle= replace(tmp_f_datestyle,"MI","0"&minute(f_obj("EditTime")))
				else
					tmp_f_datestyle= replace(tmp_f_datestyle,"MI",minute(f_obj("EditTime")))
				end if
			end if
			if instr(f_datestyle,"SS")>0 then
				if second(f_obj("EditTime"))<10 then
					tmp_f_datestyle= replace(tmp_f_datestyle,"SS","0"&second(f_obj("EditTime")))
				else
					tmp_f_datestyle= replace(tmp_f_datestyle,"SS",second(f_obj("EditTime")))
				end if
			end if
			s_Content = replace(s_Content,"{DS:FS_EditTime}",""&tmp_f_datestyle&"")
		end if
		dim ajax_str,hits_str
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_Hits}")>0 then
				hits_str = "<span id=""DS_id_click_"&f_obj("ID")&"""></span><script language=""JavaScript"" src=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Click_ajax.asp?type=js&SubSys=DS&spanid=DS_id_click_"&f_obj("ID")&"""></script>"&chr(10)
				s_Content = replace(s_Content,"{DS:FS_Hits}",hits_str)
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_ClickNum}")>0 then
				hits_str = "<span id=""DS_id_click_"&f_obj("ID")&"_Down""></span><script language=""JavaScript"" src=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Click_ajax.asp?type=js&SubSys=DS&Get=ClickNum&spanid=DS_id_click_"&f_obj("ID")&"_Down""></script>"&chr(10)
				s_Content = replace(s_Content,"{DS:FS_ClickNum}",hits_str)
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_SystemType}")>0 then
			s_Content = replace(s_Content,"{DS:FS_SystemType}",""&f_obj("SystemType")&"")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_Accredit}")>0 then
			s_Content = replace(s_Content,"{DS:FS_Accredit}",""&Replacestr(f_obj("Accredit"),"1:免费,2:共享,3:试用,4:演示,5:注册,6:破解,7:零售,8:其它")&"")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_Version}")>0 then
			s_Content = replace(s_Content,"{DS:FS_Version}",""&f_obj("Version")&"")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_Appraise}")>0 then
			dim db_ii,Str_Tmp,ii_,vpath
			if G_VIRTUAL_ROOT_DIR<>"" then
				vpath="/"&G_VIRTUAL_ROOT_DIR
			else
				vpath=G_VIRTUAL_ROOT_DIR
			end if
			db_ii = f_obj("Appraise")
			if db_ii = "" or isnull(db_ii) then  db_ii = 0
			if db_ii>6 then db_ii=6
			Str_Tmp = ""
			for ii_ = 1 to db_ii
				Str_Tmp = Str_Tmp & "<img border=0 src="""&vpath&"/sys_images/icon_star_2.gif"" title="""&f_obj("Appraise")&"星"">"
			next 
			for ii_ = 1 to 6 - db_ii
				Str_Tmp = Str_Tmp & "<img border=0 src="""&vpath&"/sys_images/icon_star_1.gif"" title="""&f_obj("Appraise")&"星"">"
			next 
			s_Content = replace(s_Content,"{DS:FS_Appraise}",""&Str_Tmp&"")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_FileSize}")>0 then
			s_Content = replace(s_Content,"{DS:FS_FileSize}",""&f_obj("FileSize")&"")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_Language}")>0 then
			s_Content = replace(s_Content,"{DS:FS_Language}",""&f_obj("Language")&"")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_PassWord}")>0 then
			s_Content = replace(s_Content,"{DS:FS_PassWord}",""&f_obj("PassWord")&"")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_Provider}")>0 then
			s_Content = replace(s_Content,"{DS:FS_Provider}",""&f_obj("Provider")&"")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_ProviderUrl}")>0 then
			s_Content = replace(s_Content,"{DS:FS_ProviderUrl}",""&f_obj("ProviderUrl")&"")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_EMail}")>0 then
			s_Content = replace(s_Content,"{DS:FS_EMail}","<a href=""mailto:"&f_obj("EMail")&""">"&f_obj("EMail")&"</a>")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_Types}")>0 then
			s_Content = replace(s_Content,"{DS:FS_Types}",""&Replacestr(f_obj("Types"),"1:图片,2:文件,3:程序,4:Flash,5:音乐,6:影视,7:其它"))
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_OverDue}")>0 then
			if f_obj("OverDue")=0 then
				s_Content = replace(s_Content,"{DS:FS_OverDue}","永不过期")
			else
				s_Content = replace(s_Content,"{DS:FS_OverDue}",""&f_obj("OverDue")&"天")
			end if
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_ConsumeNum}")>0 then
			s_Content = replace(s_Content,"{DS:FS_ConsumeNum}",""&f_obj("ConsumeNum")&"")
		end if
		'_________________________________________________________________________________________________
		'___专区部分______________________________________________________________________________________________
		If instr(s_Content,"{DS:FS_SpecialList}")>0 Then
			if trim(get_SpecialEName)<>"" then
				tmp_list = ""
				s_array = split(get_SpecialEName,",")
				for s_t_i = 0 to ubound(s_array)
							set s_m_Rs=Conn.execute("select  SpecialID,SpecialCName,SpecialEName From FS_DS_Special where SpecialEName='"&trim(s_array(s_t_i))&"' order by SpecialID desc")
							if not s_m_Rs.eof then
								tmp_list = tmp_list &"<a href="""& get_specialLink(s_m_Rs("SpecialEName")) &""">" & s_m_Rs("SpecialCName") &"</a>&nbsp;"
							else
								tmp_list = tmp_list
							end if
							s_m_Rs.close:set s_m_Rs=nothing
				next
				tmp_list = tmp_list
			end if
			s_Content = replace(s_Content,"{DS:FS_SpecialList}",tmp_list)
		end if
		'_________________________________________________________________________________________________
		Dim m_Rs_special,m_sp_sql,array_special,i_special,s_SpecialName,m_save_special,special_UrlDomain
		if instr(s_Content,"{DS:FS_SpecialName}")>0 then
			if trim(get_SpecialEName)<>"" then
				s_SpecialName = ""
				array_special = split(get_SpecialEName,",")
				for i_special = 0 to Ubound(array_special)
					m_sp_sql = "select SpecialID,SpecialCName,naviText,SpecialEName,SavePath,FileExtName,isLock,naviPic From FS_NS_Special where isLock=0 and SpecialEName='"&trim(array_special(i_special))&"'"
					set m_Rs_special=Conn.execute(m_sp_sql)
					if not m_Rs_special.eof then
						m_save_special = get_specialLink(m_Rs_special("SpecialEName"))
						if i_special=ubound(array_special) then
							s_SpecialName = s_SpecialName & "<a href="""&m_save_special&""">" &m_Rs_special("SpecialCName")&"</a>"
						else
							s_SpecialName = s_SpecialName & "<a href="""&m_save_special&""">" &m_Rs_special("SpecialCName")&"</a>&nbsp;"
						end if
					else
						s_SpecialName = ""
					end if
					m_Rs_special.close:set m_Rs_special = nothing
				next
				s_Content = replace(s_Content,"{DS:FS_SpecialName}",s_SpecialName)
			else
				s_Content = replace(s_Content,"{DS:FS_SpecialName}","")
			end if
		end if
		'-----------------------------------------------------------------------------------------------------
		if instr(s_Content,"{DS:FS_Pic}")>0 then
			if trim(f_obj("Pic"))<>"" then
				s_Content = replace(s_Content,"{DS:FS_Pic}",f_obj("Pic"))
			else
				s_Content = replace(s_Content,"{DS:FS_Pic}","/sys_images/NoPic.jpg")
			end if
		end if
		'_________________________________________________________________________________________________
		'以下是自定义字段替换
		if instr(s_Content,"{DS=Define|")>0 then
			dim define_rs_sql,define_rs
			define_rs_sql="select ID,TableEName,ColumnName,ColumnValue,InfoID,InfoType From FS_MF_DefineData where InfoType='DS' and InfoID='"&f_obj("DownLoadID")&"' order by ID desc"
			set define_rs=Conn.execute(define_rs_sql)
			if not define_rs.eof  then
				do while not define_rs.eof
					s_Content = replace(s_Content,"{DS=Define|"&define_rs("TableEName")&"}",""&define_rs("ColumnValue"))
					define_rs.movenext
				loop
				define_rs.close:set define_rs=nothing
			else
				dim define_class_sql,define_class_rs
				define_class_sql="select D_Coul From FS_MF_DefineTable where D_SubType='DS' order  by DefineID desc"
				set define_class_rs=Conn.execute(define_class_sql)
				if not define_class_rs.eof then
					do while not define_class_rs.eof
						s_Content = replace(s_Content,"{DS=Define|"&define_class_rs("D_Coul")&"}","")
						define_class_rs.movenext
					loop
				end if
				define_class_rs.close:set define_class_rs=nothing
				define_rs.close:set define_rs=nothing
			end if
		end if
'--------------------------------------------------------------------------------------------------------------

		if instr(s_Content,"{DS:FS_FormReview}")>0 Then
			Dim ReviewStr
			if f_obj("ReviewTF")=1 Then
				ReviewStr = "<span id=""Review_TF_"& f_obj("ID") &""">loading...</span><script language=""JavaScript"" type=""text/javascript"" src=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/ReviewTF.asp?Id="& f_obj("ID") &"&Type=DS""></script>"
			Else
				ReviewStr = ""
			End If
			s_Content = replace(s_Content,"{DS:FS_FormReview}",ReviewStr)
		End If
		'_________________________________________________________________________________________________
		If instr(s_Content,"{DS:FS_ShowComment}")>0 Then
			if f_obj("ReviewTF")=1 then
				s_Content = replace(s_Content,"{DS:FS_ShowComment}","<span id=""DS_show_review_"& f_obj("ID") &""">评论加载中...</span><script language=""JavaScript"" type=""text/javascript"" src=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/ShowReview.asp?Id="&f_obj("ID")&"&Type=DS&SpanId=DS_show_review_"& f_obj("ID") &"""></script>")
			Else
				s_Content = replace(s_Content,"{DS:FS_ShowComment}","")
			End If	
		End If
		'_________________________________________________________________________________________________
		If InStr(s_Content,"{DS:FS_ReviewURL}")>0 Then
			if f_obj("ReviewTF")=1 then
				s_Content = replace(s_Content,"{DS:FS_ReviewURL}","<a href=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/ShowReviewList.asp?type=DS&Id="&f_obj("ID")&""">评</a>")
			Else
				s_Content = replace(s_Content,"{DS:FS_ReviewURL}","")
			end if
		End if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_AddFavorite}")>0 then
			s_Content = replace(s_Content,"{DS:FS_AddFavorite}","<a href=""http://"& Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/User/AddFavor.asp?Id="&f_obj("id")&"&Type=ds""><img src="""& m_PathDir &"sys_images/Favorite.gif"" border=""0"" alt=""加入收藏夹""></a>")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{DS:FS_SendFriend}")>0 then
			s_Content = replace(s_Content,"{DS:FS_SendFriend}","<a href=""http://"& Request.Cookies("FoosunMFCookies")("FoosunMFDomain") &"/"&G_USER_DIR&"/Sendmail.asp?Id="&f_obj("id")&"&Type=ds""><img src="""& m_PathDir &"sys_images/sendmail.gif"" border=""0"" alt=""发送给好友""></a>")
		end if

		'获得栏目地址_______________________________________________________________________________________________
		s_f_classSql = "select ClassID,ClassName,ClassEName,[Domain],ClassNaviContent,ClassNaviPic,SavePath,FileSaveType,ClassKeywords,Classdescription,FileExtName from FS_DS_Class where ClassId='"&f_obj("ClassId")&"' and ReycleTF=0 order by OrderID desc,id desc"
		if instr(s_Content,"{DS:FS_ClassURL}")>0 then
			class_path=get_ClassLink(f_obj("ClassId"))
			s_Content = replace(s_Content,"{DS:FS_ClassURL}",class_path)
		end if
		if instr(s_Content,"{DS:FS_ClassName}")>0 then
			set m_Rs_class = Conn.execute(s_f_classSql)
			s_Content = replace(s_Content,"{DS:FS_ClassName}",""&m_Rs_class("ClassName")&"")
			m_Rs_class.close:set m_Rs_class=nothing
		end if
		if instr(s_Content,"{DS:FS_ClassNaviPicURL}")>0 then
			set m_Rs_class = Conn.execute(s_f_classSql)
			if trim(m_Rs_class("ClassNaviPic"))<>"" then
				s_Content = replace(s_Content,"{DS:FS_ClassNaviPicURL}",""& m_Rs_class("ClassNaviPic") &"")
			else
				s_Content = replace(s_Content,"{DS:FS_ClassNaviPicURL}","")
			end if
			m_Rs_class.close:set m_Rs_class=nothing
		end if
		if instr(s_Content,"{DS:FS_ClassNaviDescript}")>0 then
			set m_Rs_class = Conn.execute(s_f_classSql)
			if trim(m_Rs_class("ClassNaviContent"))<>"" then
				s_Content = replace(s_Content,"{DS:FS_ClassNaviDescript}",""& m_Rs_class("ClassNaviContent") &"")
			else
				s_Content = replace(s_Content,"{DS:FS_ClassNaviDescript}","")
			end if
			m_Rs_class.close:set m_Rs_class=nothing
		end if
		if instr(s_Content,"{DS:FS_ClassNaviContent}")>0 then
			set m_Rs_class = Conn.execute(s_f_classSql)
			s_Content = replace(s_Content,"{DS:FS_ClassNaviContent}",""&m_Rs_class("ClassNaviContent"))
			m_Rs_class.close:set m_Rs_class=nothing
		end if
		if instr(s_Content,"{DS:FS_ClassKeywords}")>0 then
			set m_Rs_class = Conn.execute(s_f_classSql)
			s_Content = replace(s_Content,"{DS:FS_ClassKeywords}",""&m_Rs_class("ClassKeywords"))
			m_Rs_class.close:set m_Rs_class=nothing
		end if
		if instr(s_Content,"{DS:FS_Classdescription}")>0 then
			set m_Rs_class = Conn.execute(s_f_classSql)
			s_Content = replace(s_Content,"{DS:FS_Classdescription}",""&m_Rs_class("Classdescription"))
			m_Rs_class.close:set m_Rs_class=nothing
		end if
		getlist_news = s_Content
	End Function

	'开始读取标签____下载终极类标签_____________________________________________________________________
	Public Function ClassList(f_Lable,f_type,f_Id)
		if f_Id<>"" then
			dim div_tf,newnumber,datenumber,titlenumber,picshowtf,openstyle,orderby,orderdesc,pageTF,pagestyle,pagenumber,pagecss,datestyle,colnumber,contentnumber,navinumber,datenumber_tmp_
			dim classNews_head,classNews_bottom,classNews_middle1,classNews_middle2,search_inSQL,order_in_str_awen
			dim datenumber_tmp,f_sql,f_configsql,f_rs_obj,f_rs_configobj,MF_Domain
			dim TPageNum,perPageNum,PageNum,sPageCount,cl_i,c_i_k,rs_c,RefreshNumber,Page_flag_TF
			Dim f_Content,f_TableName,f_SelectFieldNames,f_PageIndex,f_Where,f_PaginationStr,f_IDSRS,f_IDSArray,f_IDS,i,f_BeginIndex,f_IndexID
			m_Err_Info = ""
			datenumber= f_Lable.LablePara("多少天")
			titlenumber= f_Lable.LablePara("标题数")
			picshowtf=f_Lable.LablePara("图文标志")
			openstyle=f_Lable.LablePara("打开窗口")
			orderby = f_Lable.LablePara("排列字段")
			orderdesc = f_Lable.LablePara("排列方式")
			pageTF = f_Lable.LablePara("分页")
			pagestyle = f_Lable.LablePara("分页样式")
			pagenumber = f_Lable.LablePara("每页数量")
			if pagenumber = "" OR Not IsNumeric(pagenumber) then pagenumber = 10
			pagecss = f_Lable.LablePara("PageCSS")
			datestyle = f_Lable.LablePara("日期格式")
			colnumber= f_Lable.LablePara("排列数")
			contentnumber= f_Lable.LablePara("内容字数")
			if f_Lable.LablePara("输出格式") = "out_DIV" then
				div_tf=1
			else
				div_tf=0
			end if
			Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
			f_DivID = f_Lable.LablePara("DivID")
			f_DivClass = f_Lable.LablePara("Divclass")
			f_UlID = f_Lable.LablePara("ulid")
			f_ULClass = f_Lable.LablePara("ulclass")
			f_LiID = f_Lable.LablePara("liid")
			f_LiClass = f_Lable.LablePara("liclass")
			classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			
			f_PaginationStr = pagestyle & "," & pagecss
			f_Lable.DictLableContent.Item("0") = f_PaginationStr
			MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
			If G_IS_SQL_DB=1 Then
				datenumber_tmp_="datediff(day,AddTime,'"&date()&"')"
			else
				datenumber_tmp_="datediff('d',AddTime,'"&date()&"')"
			end if  
			datenumber_tmp=" and  (OverDue=0 or (OverDue>0 and "&datenumber_tmp_&"<= OverDue))"
			If datenumber ="0" Then
				datenumber_tmp = ""
			Else
				If G_IS_SQL_DB=1 Then
					datenumber_tmp = " and dateadd(d,"&datenumber&",AddTime)>=getdate()"
				Else
					datenumber_tmp = " and dateadd('d',"&datenumber&",AddTime)>=now"
				End If
			End If
			order_in_str_awen = ""&orderby&" "&orderdesc&""
			if ucase(orderby)<>"ID" then order_in_str_awen = order_in_str_awen & ",ID Desc"
			set rs_c=conn.execute("select RefreshNumber from FS_DS_Class where  ClassId='"& f_Id &"'")
			if rs_c.eof then
					RefreshNumber = ""
				 rs_c.close:set rs_c = nothing
			else
				if rs_c(0)=0 then
					RefreshNumber = ""
				else
					RefreshNumber = "top "&rs_c(0)&""
				end if
				rs_c.close:Set rs_c=nothing
			end if
			search_inSQL=" and ClassId='"& f_Id &"'"
			f_Where = "AuditTF=1 "&search_inSQL & datenumber_tmp &" order by "&order_in_str_awen
			f_TableName = "FS_DS_List"
			f_IndexID = "ID"
			f_SelectFieldNames = "ID,DownLoadID,ClassID,Description,Accredit,AddTime,Appraise,AuditTF,BrowPop,ClickNum,EditTime,EMail,SavePath,FileExtName,FileName,FileSize,[Language],Name,NewsTemplet,PassWord,Pic,Property,Provider,ProviderUrl,RecTF,ReviewTF,ShowReviewTF,SystemType,Types,Version,OverDue,ConsumeNum,Hits"
			Set f_IDSRS = Server.CreateObject(G_FS_RS)
			f_IDSRS.Open "Select " & f_IndexID & " from " & f_TableName & " Where " & f_Where,Conn,0,1
			if Not f_IDSRS.Eof then
				f_PageIndex = 1
				f_IDSArray = f_IDSRS.GetRows()
				Set  f_rs_obj = Server.CreateObject(G_FS_RS)
				Do While True
					f_IDS = ""
					f_BeginIndex = (f_PageIndex - 1) * pagenumber
					for i = f_BeginIndex To pagenumber * f_PageIndex - 1
						if i > UBound(f_IDSArray,2) then Exit For
						if f_IDS = "" then
							f_IDS = f_IDSArray(0,i)
						else
							f_IDS = f_IDS & "," & f_IDSArray(0,i)
						end if
					Next
					if f_IDS = "" then Exit Do
			
					f_Where = " ID In (" & f_IDS & ") order by "&order_in_str_awen
					f_sql = "Select " & f_SelectFieldNames & " from " & f_TableName & " where " & f_Where
					f_rs_obj.open f_sql,Conn,0,1
					cl_i = 1
					ClassList = classNews_head
					Do While Not f_rs_obj.Eof
						f_Content = getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,"ClassList")
						if div_tf=1 then
							ClassList= ClassList & classNews_middle1 & f_Content & classNews_middle2
						else
							if cint(colnumber) =1 then
								ClassList= ClassList &"   <tr><td>" & f_Content & "</td></tr>"
							else
								if cl_i mod colnumber = 1 then ClassList= ClassList & "<tr>"
								ClassList= ClassList & "<td width="""& cint(100/cint(colnumber))&"%"">" & f_Content & "</td>"
								if cl_i mod colnumber = 0 then ClassList= ClassList & "</tr>"
							end if
						end if
						f_rs_obj.movenext
						cl_i = cl_i + 1
					Loop
					ClassList = ClassList & classNews_bottom
					f_Lable.DictLableContent.Add f_PageIndex & "",ClassList
					f_rs_obj.Close
					f_PageIndex = f_PageIndex + 1
				Loop 
				Set f_rs_obj = Nothing
			else
				f_Lable.DictLableContent.Add "1",""
			end if
			f_IDSRS.Close
			Set f_IDSRS = Nothing
		else
			f_Lable.DictLableContent.Add "1","不能非栏目页面插入标签"
		end if
	End Function
	'得到子类新闻列表
	Public Function subClassList(f_Lable,f_type,f_Id)
		dim rs,f_sql,rs_n,rs_f_sql
		dim div_tf,c_s_i,bg_ground,c_cols,c_cols_1,datenumber,datenumber_tmp,orderby,orderdesc,Inc_SubClass,SQL_Inc_SubClass
		dim titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,loopnumber
		dim classNews_head,classNews_middle1,classNews_bottom,classNews_middle2,c_s_i_1
		c_cols = f_Lable.LablePara("栏目排列数")
		c_cols_1 = f_Lable.LablePara("排列数")
		bg_ground = f_Lable.LablePara("背景底纹")
		datenumber = f_Lable.LablePara("多少天")
		orderby = f_Lable.LablePara("排列字段")
		orderdesc = f_Lable.LablePara("排列方式")
		loopnumber = f_Lable.LablePara("Loop")
		titlenumber = f_Lable.LablePara("标题数")
		contentnumber = f_Lable.LablePara("内容字数")
		picshowtf = f_Lable.LablePara("图文标志")
		datestyle =  f_Lable.LablePara("日期格式")
		openstyle =  f_Lable.LablePara("打开窗口")
		Inc_SubClass = f_Lable.LablePara("包含子类")
		if f_Lable.LablePara("输出格式") = "out_DIV" then
			div_tf=1
		else
			div_tf=0
		end if
		if not isnumeric(titlenumber) then:titlenumber = 30:else:titlenumber = cint(titlenumber):end if
		if not isnumeric(loopnumber) then:loopnumber = 10:else:loopnumber = cint(loopnumber):end if
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
		f_DivID = f_Lable.LablePara("DivID")
		f_DivClass = f_Lable.LablePara("Divclass")
		f_UlID = f_Lable.LablePara("ulid")
		f_ULClass = f_Lable.LablePara("ulclass")
		f_LiID = f_Lable.LablePara("liid")
		f_LiClass = f_Lable.LablePara("liclass")
		classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		
		if len(bg_ground)>5 then
			 bg_ground = " background="""& bg_ground &""""
		else
			 bg_ground = ""
		end If
		If datenumber <>"0" Then
			If G_IS_SQL_DB=1 Then
				datenumber_tmp = " and dateadd(d,"&datenumber&",AddTime)>=getdate()"
			Else
				datenumber_tmp = " and dateadd('d',"&datenumber&",AddTime)>=now"
			End If
		Else
			datenumber_tmp = ""
		End If 
		if f_Id="" then f_id="0"
		f_sql = "select ClassId,ClassName,ClassEName,ParentID From FS_DS_Class Where ParentID='"& f_Id &"' and IsURL=0 and ReycleTF=0 order by OrderID desc,id desc"
		set rs = Conn.execute(f_sql)
		if not rs.eof then
			subClassList = "<table border=""0"" cellspacing=""3"" cellpadding=""0"" width=""100%"">"&vbNewLine&"  <tr>"
			c_s_i = 0
			do while not rs.eof
				If (c_s_i mod c_cols =0) And c_s_i>0 Then
					subClassList = subClassList & "</tr>"&vbNewLine&"<tr>"&vbNewLine
				End If
				subClassList = subClassList & "<td width="""&cint(100/c_cols)&"%"" valign=""top""><table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"""& bg_ground &"><tr><td height=""26"">&nbsp;&nbsp;<a href="""& get_ClassLink(rs("ClassId"))&""">"&rs("ClassName")&"</a></td><td align=""center"" width=""20%""><a href="""& get_ClassLink(rs("ClassId"))&"""><img src=""http://"& Request.Cookies("FoosunMFCookies")("FoosunMFDomain") &"/sys_images/news_more.gif"" border=""0"" alt="""& rs("ClassName") &"..更多""></a></td></tr></table>"
				If Inc_SubClass="1" Then
					If getNewsSubClass(rs("ClassId"))<>"" Then
						SQL_Inc_SubClass=" and (ClassID='"&rs("ClassId")&"' OR ClassID IN ("&getNewsSubClass(rs("ClassId"))&"))"
					Else
						SQL_Inc_SubClass=" AND ClassID='"&rs("ClassId")&"'"
					End If
				Else
					SQL_Inc_SubClass=" AND ClassID='"&rs("ClassId")&"'"
				End If
				rs_f_sql = "select top "& loopnumber &" * From FS_DS_List where AuditTF=1 "& datenumber_tmp & SQL_Inc_SubClass &" order by "&orderby&" "&orderdesc
				if ucase(orderby)<>"ID" then  rs_f_sql = rs_f_sql&",id "&orderdesc&""
				set rs_n = Conn.execute(rs_f_sql)
				If Not rs_n.eof then
					if div_tf=1 then
						subClassList = subClassList & vbNewLine & classNews_head
					end if
					do while not rs_n.eof
						If div_tf = 1 Then
							subClassList = subClassList & classNews_middle1 &getlist_news(rs_n,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,"subClassList")&classNews_middle2
						Else
							subClassList = subClassList & "<div>" & getlist_news(rs_n,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,"subClassList")&"<div>"&vbNewLine
						End If
						rs_n.movenext
					loop
					if div_tf=1 then
						subClassList = subClassList & classNews_bottom
					end if
				end If
				subClassList = subClassList & "</td>"&vbNewLine
				rs.movenext
				c_s_i = c_s_i + 1
			Loop
			If c_s_i Mod 2 <>0 Then
				subClassList = subClassList & "<td width="""&cint(100/c_cols)&"%"" valign=""top""></td>"&vbNewLine
			End If
			rs.close:set rs = Nothing
			rs_n.close:set rs_n = nothing
			subClassList = subClassList & "</tr>"&vbNewLine&"</table>"&vbNewLine
		else
			subClassList = ""
		end If
		f_Lable.DictLableContent.Add "1",subClassList
	End Function

	'得到下载浏览________________________________________________________________
	Public Function ReadNews(f_Lable,f_type,f_Id)
		Dim ReadSql,RsReadObj,tmpsql_
		Dim MF_Domain,datestyle
		datestyle = f_Lable.LablePara("日期格式")
		ReadNews = ""
		if trim(f_Id)="" then
			ReadNews = ""
		else
			if G_IS_SQL_DB=1 then 
				tmpsql_ = "datediff(day,AddTime,'"&date()&"')<= OverDue"
			else
				tmpsql_ = "datediff('d',AddTime,'"&date()&"')<= OverDue"
			end if	
			ReadSql="select ID,DownLoadID,ClassID,Description,Accredit,AddTime,Appraise,AuditTF,BrowPop," _
				&"ClickNum,EditTime,EMail,SavePath,FileExtName,FileName,FileSize,[Language],Name,NewsTemplet,PassWord,Pic," _
				&"Property,Provider,ProviderUrl,RecTF,ReviewTF,ShowReviewTF,SystemType,Types,Version,OverDue,ConsumeNum,Hits"
			ReadSql = ReadSql &" From FS_DS_List where AuditTF=1 and DownLoadID='"& f_Id &"'  and  OverDue=0 or OverDue>0 and  "&tmpsql_
			Set  RsReadObj = Server.CreateObject(G_FS_RS)
			RsReadObj.open ReadSql,Conn,0,1
			MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
			if not RsReadObj.eof then
				ReadNews = ReadNews&getlist_news(RsReadObj,f_Lable.FoosunStyle.StyleContent,0,0,0,0,datestyle,0,MF_Domain,"readnews")
			else
				ReadNews =""
			end if
		end if
		f_Lable.DictLableContent.Add "1",ReadNews
	End Function
	'专区导航
	Public function SpecialNavi(f_Lable,f_type,f_Id)
		dim cols,titlecss,div_tf,cssstyle,titleNavi,rs,SpecialNavistr,classproducts_head,classproducts_bottom,classproducts_middle1,classproducts_middle2,cols_str
		cols = f_Lable.LablePara("方向")
		titlecss = f_Lable.LablePara("CSS")
		titleNavi = f_Lable.LablePara("导航")
		if f_Lable.LablePara("输出格式") = "out_DIV" then
			div_tf = 1
			cssstyle = ""
		else
			div_tf=0
			if titlecss<>"" then
				cssstyle = " Class="""& titlecss &""""
			else
				cssstyle = ""
			end if
		end if
		
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
		f_DivID = f_Lable.LablePara("DivID")
		f_DivClass = f_Lable.LablePara("Divclass")
		f_UlID = f_Lable.LablePara("ulid")
		f_ULClass = f_Lable.LablePara("ulclass")
		f_LiID = f_Lable.LablePara("liid")
		f_LiClass = f_Lable.LablePara("liclass")
		classproducts_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classproducts_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classproducts_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classproducts_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		
		set rs = Conn.execute("select SpecialID,SpecialCName,SpecialEName,isLock from fS_DS_Special where isLock=0  Order by SpecialID desc")
		SpecialNavistr = ""
		if rs.eof then
			SpecialNavistr = ""
			rs.close:set rs=nothing
		else
			do while not rs.eof 
				if div_tf=1 then
					SpecialNavistr = SpecialNavistr &classproducts_middle1 & "<a href="""&get_specialLink(rs("SpecialEName"))&""">"&rs("SpecialCName")&"</a>" & classproducts_middle2
				else
					if cols="0" then
						cols_str = "&nbsp;"
					else
						cols_str = "<br />"
					end if
					SpecialNavistr = SpecialNavistr & titleNavi & "<a href="""&get_specialLink(rs("SpecialEName"))&""""& cssstyle&">"&rs("SpecialCName")&"</a>"& cols_str &""
				end if
				rs.movenext
			loop
			rs.close:set rs=nothing
		end if
		if div_tf=1 then
			SpecialNavistr=classproducts_head & SpecialNavistr & classproducts_bottom
		end if
		f_Lable.DictLableContent.Add "1",SpecialNavistr
	End function
	'RSS聚合
	Public function Rssfeed(f_Lable,f_type,f_Id)
		dim ClassID,Rssfeedstr
		ClassID = f_Lable.LablePara("栏目")
		if trim(ClassID)<>"" then
				Rssfeedstr = "<a href=""http://"& Request.Cookies("foosunMfCookies")("foosunMfDomain") &"/xml/MS/"&ClassID&".xml""><img src=""http://"& Request.Cookies("foosunMfCookies")("foosunMfDomain") &"/sys_images/rss.gif"" border=""0""></a>"
		elseif trim(f_id)<>"" then
			Rssfeedstr = "<a href=""http://"& Request.Cookies("foosunMfCookies")("foosunMfDomain") &"/xml/MS/"&f_id&".xml""><img src=""http://"& Request.Cookies("foosunMfCookies")("foosunMfDomain") &"/sys_images/rss.gif"" border=""0""></a>"
		else
			Rssfeedstr = "<a href=""http://"& Request.Cookies("foosunMfCookies")("foosunMfDomain") &"/xml/MS/index.html""><img src=""http://"& Request.Cookies("foosunMfCookies")("foosunMfDomain") &"/sys_images/rss.gif"" border=""0""></a>"
		end if
		f_Lable.DictLableContent.Add "1",Rssfeedstr
	End function
	'专区调用
	Public function SpecialCode(f_Lable,f_type,f_Id)
		dim ClassID,titleNavi,SpecialCodestr,cols,pictf,picsize,piccss,piccssstr,ContentTf,ContentNumber,div_tf,titlecss,cssstyle,ContentCSS,ContentCSSstr,classproducts_head,classproducts_bottom,classproducts_middle1,classproducts_middle2
		ClassID = f_Lable.LablePara("专题")
		titleNavi = f_Lable.LablePara("导航")
		titlecss = f_Lable.LablePara("专题名称CSS")
		ContentCSS = f_Lable.LablePara("导航内容CSS")
		cols=f_Lable.LablePara("排列方式")
		pictf=f_Lable.LablePara("图片显示")
		picsize=f_Lable.LablePara("图片尺寸")
		piccss=f_Lable.LablePara("图片css")
		ContentTf = f_Lable.LablePara("导航内容")
		ContentNumber= f_Lable.LablePara("导航内容字数")
		if piccss<>"" then
			piccssstr = " class="""& piccss &""""
		else
			piccssstr = ""
		end if
		if trim(ClassID)="" then
			f_Lable.DictLableContent.Add "1","错误的标签,by foosun.cn"
			Exit Function
		end if
		if f_Lable.LablePara("输出格式") = "out_DIV" then
			div_tf = 1
			cssstyle = ""
		else
			div_tf=0
			if titlecss<>"" then
				cssstyle = " class="""& titlecss &""""
			else
				cssstyle = ""
			end if
			if ContentCSS<>"" then
				ContentCSSstr = " class="""& ContentCSS &""""
			else
				ContentCSSstr = ""
			end if
		end if
		
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
		f_DivID = f_Lable.LablePara("DivID")
		f_DivClass = f_Lable.LablePara("Divclass")
		f_UlID = f_Lable.LablePara("ulid")
		f_ULClass = f_Lable.LablePara("ulclass")
		f_LiID = f_Lable.LablePara("liid")
		f_LiClass = f_Lable.LablePara("liclass")
		classproducts_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classproducts_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classproducts_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classproducts_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		
		SpecialCodestr = ""
		dim rs
		set rs=Conn.execute("select SpecialCName,SpecialEName,naviText,naviPic from fS_DS_Special Where SpecialEName='"& ClassID &"'")
		if rs.eof then
			SpecialCodestr = ""
			rs.close:set rs=nothing
		else'
			if div_tf=1 then
				if pictf="1" then
					if trim(picsize)="" then
						SpecialCodestr = SpecialCodestr & ""
					else
						SpecialCodestr = SpecialCodestr & "  <a href="""&get_specialLink(rs("SpecialEName"))&"""><img "&piccssstr&" src="""&rs("naviPic")&""" width="""&split(picsize,",")(1)&""" height="""&split(picsize,",")(0)&""" border=""0"" /></a>" &chr(10) 
					end if
				end if
				SpecialCodestr = SpecialCodestr & classproducts_middle1 & "<a href="""&get_specialLink(rs("SpecialEName"))&""">"& rs("SpecialCName") &"</a>" & classproducts_middle2  
				if ContentTf="1" then
					SpecialCodestr = SpecialCodestr & classproducts_middle1 &"&nbsp;"& GotTopic(""&rs("naviText"),ContentNumber) & "&nbsp;<a href="""& get_specialLink(rs("SpecialEName")) &""">详细</a>" & classproducts_middle2  
				end if
				SpecialCodestr = classproducts_head & SpecialCodestr & classproducts_bottom
			else
				SpecialCodestr = SpecialCodestr& "<table width=""99%"" border=""0"" cellspacing=""0"" cellpadding=""5"">"&chr(10)&" <tr>"
				if pictf="1" then
					if trim(picsize)="" then
						SpecialCodestr = SpecialCodestr & ""
					else
						if cols="0" then
							SpecialCodestr = SpecialCodestr & "<td align=""center""><a href="""&get_specialLink(rs("SpecialEName"))&"""><img "&piccssstr&" src="""&rs("naviPic")&""" width="""&split(picsize,",")(1)&""" height="""&split(picsize,",")(0)&""" border=""0"" /></a></td>" &chr(10) 
						else
							SpecialCodestr = SpecialCodestr & "<td><a href="""&get_specialLink(rs("SpecialEName"))&"""><img "&piccssstr&" src="""&rs("naviPic")&""" width="""&split(picsize,",")(1)&""" height="""&split(picsize,",")(0)&""" border=""0"" /></a></td></tr>" &chr(10) 
						end if
					end if
				end if
				if cols="0" then
					SpecialCodestr = SpecialCodestr & "<td>"& titleNavi &"<a href="""&get_specialLink(rs("SpecialEName"))&""""&cssstyle&">"& rs("SpecialCName") &"</a>"
					if ContentTf="1" then
						SpecialCodestr = SpecialCodestr & "<br /><div align=""left"" "&ContentCSSstr&">&nbsp;"& GotTopic(""&rs("naviText"),ContentNumber)&"&nbsp;<a href="""& get_specialLink(rs("SpecialEName")) &""">详细</a></div></td>"&chr(10)
					else
						SpecialCodestr = SpecialCodestr & "</td></tr>"&chr(10)
					end if
				else
					SpecialCodestr = SpecialCodestr & " <tr><td>"& titleNavi &"<a href="""&get_specialLink(rs("SpecialEName"))&""""&cssstyle&">"& rs("SpecialCName") &"</a>"&chr(10)
					if ContentTf="1" then
						SpecialCodestr = SpecialCodestr & "<br /><div align=""left"" "&ContentCSSstr&">&nbsp;"& GotTopic(""&rs("naviText"),ContentNumber)&"&nbsp;<a href="""& get_specialLink(rs("SpecialEName")) &""">详细</a></div></td></tr>"&chr(10)
					else
						SpecialCodestr = SpecialCodestr & "</td></tr>"&chr(10)
					end if
				end if
				SpecialCodestr = SpecialCodestr & "</table>"
			end if
			rs.close:set rs=nothing
		end if
		f_Lable.DictLableContent.Add "1",SpecialCodestr
	End function

	'站点地图____________________________________________________________________
	Public Function SiteMap(f_Lable,f_type,f_Id)
		dim classId,cssstyle,SiteMapstr,RsClassObj,i,br_str
		classId = f_Lable.LablePara("栏目")
		cssstyle = f_Lable.LablePara("标题CSS")
		If classId = "" Or IsNull(classId) Then classId = 0
		SiteMapstr = ""
		set RsClassObj = Conn.execute("select ClassId,ClassName,ClassEName,IsURL,ParentID From FS_DS_Class where ReycleTF=0 and ParentID='" & classId & "' order by OrderID desc,id desc")
		if Not RsClassObj.eof then
			i=0
			do while Not RsClassObj.eof
				if RsClassObj("ParentID")<>"0" then
					if i=0 then
						br_str=""
					else
						br_str="<br />"
					end if
				else
					if i=0 then
						br_str=""
					else
						br_str="<br />"
					end if
				end if
				if cssstyle<>"" then
					SiteMapstr = SiteMapstr & br_str & "<img src="""&m_PathDir&"sys_images/+.gif"" border=""0"" /><a href="""&get_ClassLink(RsClassObj("ClassId"))&""" class="""& cssstyle &""">"&RsClassObj("ClassName")&"</a>"&chr(10)
				else
					SiteMapstr = SiteMapstr & br_str & "<img src="""&m_PathDir&"sys_images/+.gif"" border=""0"" /><a href="""&get_ClassLink(RsClassObj("ClassId"))&""">"&RsClassObj("ClassName")&"</a>"&chr(10)
				end if
				SiteMapstr = SiteMapstr & get_ClassList(RsClassObj("ClassId"),"&nbsp;",cssstyle)
				RsClassObj.movenext
				i=i+1
			loop
			RsClassObj.close:set RsClassObj=nothing
		else
			SiteMapstr = ""
			RsClassObj.close:set RsClassObj=nothing
		end if
		f_Lable.DictLableContent.Add "1",SiteMapstr
	End Function

	'得到搜索表单________________________________________________________________
	Public Function Search(f_Lable,f_type)
		dim Searchstr,showdate,showClass,rs,select_str,datestr,classstr,MF_Domain
		showdate = f_Lable.LablePara("显示日期")
		showClass = f_Lable.LablePara("显示栏目")
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		Searchstr = ""
		if showdate = "1" then
			datestr = "开始日期:<input name=""s_date""  class=""f-text"" type=""text"" value="""&date()-1&""" size=""10"" /> 结束日期：<input name=""e_date"" class=""f-text"" type=""text"" value="""&date()&""" size=""10"" />"
		else
			datestr = ""
		end if
		if showClass = "1" then
			classstr = " <select class=""f-select"" name=""ClassId"" id=""ClassId"">"
			classstr = classstr&"<option value="""">--所有--</option>"&chr(10)
			set rs = Conn.execute("select ClassId,id,classname,classEname,ParentId From FS_DS_Class where ReycleTF=0 and isURL=0 and ParentId='0' order by OrderId desc,id desc")
			do while not rs.eof
				classstr = classstr & "<option value="""&rs("ClassId")&""">┝"&rs("classname")&"</option>"&chr(10)
				classstr = classstr & get_optionNewsList(rs("ClassId"),"┝")&chr(10)
				rs.movenext
			loop
			rs.close:set rs=nothing
			classstr = classstr&"</select>"
		else
			classstr = ""
		end if
		select_str =" <select class=""f-select"" name=""s_type"" id=""s_type"">"&chr(10)

		select_str=select_str&"<option value=""title"" selected=""selected"">标题</option>"&chr(10)
		select_str=select_str&"<option value=""content"">全文</option>"&chr(10)
		select_str=select_str&"</select>"&chr(10)
		Searchstr = Searchstr & "<table width=""100%""><form action=""http://"& MF_Domain & "/Search.html"" id=""SearchForm"" name=""SearchForm"" method=""get"">" _
			&"<tr><td><input name=""SubSys"" type=""hidden"" id=""SubSys"" value=""DS"" /><input class=""f-text"" name=""Keyword"" type=""text"" id=""Keyword"" size=""15"" /> " _
			&select_str&datestr&classstr&" <input type=""submit""class=""f-button"" value=""搜索"" /></td></tr></form></table>"
		f_Lable.DictLableContent.Add "1",Searchstr
	 End Function

	 '得到统计信息_______________________________________________________________
	Public Function infoStat(f_Lable,f_type)
		dim cols,br_str,infoStatstr
		cols = f_Lable.LablePara("显示方向")
		if cols="0" then
			br_str = "&nbsp;"
		else
			br_str = "<br />"
		end if
		dim rs,c_rs1,c_rs2,c_rs3,c_rs4,c_rs5
		If G_IS_SQL_DB=1 Then
			set rs = Conn.execute("select count(id) From FS_DS_List where AuditTF=1  and (OverDue=0 or (OverDue>0 and datediff(d,AddTime,getdate())<= OverDue))")
		Else
			set rs = Conn.execute("select count(id) From FS_DS_List where AuditTF=1  and (OverDue=0 or (OverDue>0 and datediff('d',AddTime,now)<= OverDue))")
		End If 
		c_rs1=rs(0)
		rs.close:set rs=nothing
		set rs = Conn.execute("select count(id) From FS_DS_Class where ReycleTF=0")
		c_rs2=rs(0)
		rs.close:set rs=nothing
		set rs = User_Conn.execute("select count(Userid) From FS_ME_Users")
		c_rs4=rs(0)
		rs.close:set rs=nothing
		if G_IS_SQL_User_DB=0 then
			set rs = User_Conn.execute("select count(Userid) From FS_ME_Users where RegTime=now")
		else
			set rs = User_Conn.execute("select count(Userid) From FS_ME_Users where RegTime=getdate()")
		end if
		c_rs5=rs(0)
		rs.close:set rs=nothing
		infoStatstr = "总下载:"&"<strong>"&c_rs1&"</strong>"&br_str
		infoStatstr = infoStatstr &"总栏目:"&"<strong>"&c_rs2&"</strong>"&br_str
		infoStatstr = infoStatstr &"会员数:"&"<strong>"&c_rs4&"</strong>"&br_str
		infoStatstr = infoStatstr &"今日注册:"&"<strong>"&c_rs5&"</strong>"&br_str
		f_Lable.DictLableContent.Add "1",infoStatstr
	End Function

	'栏目导航____________________________________________________________________
	Public Function ClassNavi(f_Lable,f_type,f_Id)
		dim ClassId,cols,titlecss,rs,ClassNavistr,ParentIDstr,cols_str,div_tf
		dim classNews_head,classNews_bottom,classNews_middle1,classNews_middle2,cssstyle,titleNavi
		ClassId = f_Lable.LablePara("栏目")
		cols = f_Lable.LablePara("方向")
		titlecss = f_Lable.LablePara("标题CSS")
		titleNavi = f_Lable.LablePara("标题导航")
		if f_Lable.LablePara("输出格式") = "out_DIV" then
			div_tf = 1
			cssstyle = ""
		else
			div_tf=0
			if titlecss<>"" then
				cssstyle = " Class="""& titlecss &""""
			else
				cssstyle = ""
			end if
		end if
		
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
		f_DivID = f_Lable.LablePara("DivID")
		f_DivClass = f_Lable.LablePara("Divclass")
		f_UlID = f_Lable.LablePara("ulid")
		f_ULClass = f_Lable.LablePara("ulclass")
		f_LiID = f_Lable.LablePara("liid")
		f_LiClass = f_Lable.LablePara("liclass")
		if div_tf=1 then
			classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		end if
		classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		if trim(ClassId)<>"" then
			ParentIDstr = " and ParentId = '"& ClassId&"'"
		else
			if f_id<>"" then
				ParentIDstr = " and ParentId = '"& f_Id&"'"
			else
				ParentIDstr = " and ParentId = '0'"
			end if
		end if
		set rs = Conn.execute("select ClassID,OrderID,ClassName,isShow,ParentID,ReycleTF From FS_DS_Class where isShow=1 and ReycleTF=0 "& ParentIDstr &" Order by OrderId desc,id desc")
		ClassNavistr = ""
		if rs.eof then
			ClassNavistr = ""
			rs.close:set rs=nothing
		else
			do while not rs.eof
				if div_tf=1 then
					ClassNavistr = ClassNavistr &classNews_middle1 & "<a href="""&get_ClassLink(rs("ClassId"))&""">"&rs("ClassName")&"</a>" & classNews_middle2
				else
					if cols="0" then
						cols_str = "&nbsp;"
					else
						cols_str = "<br />"
					end if
					ClassNavistr = ClassNavistr & titleNavi & "<a href="""&get_ClassLink(rs("ClassId"))&""""& cssstyle&">"&rs("ClassName")&"</a>"& cols_str &""
				end if
				rs.movenext
			loop
			rs.close:set rs=nothing
			if div_tf=1 then
				ClassNavistr=classNews_head & ClassNavistr & classNews_bottom
			end if
		end if
		f_Lable.DictLableContent.Add "1",ClassNavistr
	End Function

	'得到下载单个地址____________________________________________________________
	Public Function get_DownLink(f_id)
		get_DownLink = ""
		dim rs,config_rs,config_mf_rs,class_rs
		dim SaveNewsPath,FileName,FileExtName,ClassId,LinkType,MF_Domain,Url_Domain,ClassEName,c_Domain,c_SavePath,IsDomain
		set rs = Conn.execute("select ID,ClassId,DownLoadID,SavePath,FileName,FileExtName From FS_DS_List where DownLoadID='"&f_id&"'")
		SaveNewsPath = rs("SavePath")
		if right(SaveNewsPath,1)="/" then SaveNewsPath = left(SaveNewsPath,len(SaveNewsPath)-1)
		FileName = rs("FileName")
		FileExtName = rs("FileExtName")
		ClassId = rs("ClassId")
		LinkType = Request.Cookies("FoosunDSCookies")("FoosunDSLinkType")
		IsDomain = Request.Cookies("FoosunDSCookies")("FoosunDSDomain")
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		set class_rs = Conn.execute("select ClassEName,IsURL,URLAddress,[Domain],SavePath From FS_DS_Class where ClassId='"&ClassId&"'")
		if not class_rs.eof then
			ClassEName = class_rs("ClassEName")
			c_Domain = class_rs("Domain")
			c_SavePath = class_rs("SavePath")
			class_rs.close:set class_rs=nothing
		else
			ClassEName = ""
			class_rs.close:set class_rs=nothing
		end if
		if not rs.eof then
			if LinkType = 1 then
				if trim(c_Domain)<>"" then
					Url_Domain = "http://"&c_Domain
				else
					if trim(IsDomain)<>"" then
						Url_Domain = "http://"&IsDomain
						c_SavePath = ""
					else
						Url_Domain = "http://"&MF_Domain
						c_SavePath = c_SavePath 
					end if
				end if
			else
				if trim(c_Domain)<>"" then
					Url_Domain = "http://"&c_Domain
				else
					if trim(IsDomain)<>"" then
						Url_Domain = "http://"&IsDomain
						c_SavePath = ""
					else
						if G_VIRTUAL_ROOT_DIR<>"" then
							Url_Domain = "/"& G_VIRTUAL_ROOT_DIR
						else
							Url_Domain = ""
						end if
						c_SavePath = c_SavePath 
					end if
				end if
			end if
			if trim(c_Domain)<>"" then
				get_DownLink = Url_Domain & Replace(SaveNewsPath &"/"&FileName&"."&FileExtName,"//","/")
			else
				get_DownLink = Url_Domain & Replace(c_SavePath& "/" & ClassEName &SaveNewsPath &"/"& FileName&"."&FileExtName,"//","/")
			end if
		rs.close:set rs=nothing
	  else
			get_DownLink = ""
			rs.close:set rs=nothing
	  end if
	  get_DownLink = get_DownLink
	End Function

	'得到栏目地址________________________________________________________________
	Public function get_ClassLink(f_id)
		dim IsDomain,LinkType,MF_Domain,c_rs,ClassEName,c_Domain,Url_Domain,ClassSaveType,class_savepath,FileExtName,c_SavePath
		LinkType = Request.Cookies("FoosunDSCookies")("FoosunDSLinkType")
		IsDomain = Request.Cookies("FoosunDSCookies")("FoosunDSDomain")
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		set c_rs=conn.execute("select id,IsURL,UrlAddress,ClassId,ClassEName,[Domain],FileSaveType,FileExtName,SavePath From FS_DS_Class where ClassId='"&f_id&"'")
		if not c_rs.eof then
			if c_rs("IsURL")=0 then
				ClassEName = c_rs("ClassEName")
				c_Domain = c_rs("Domain")
				FileExtName = c_rs("FileExtName")
				ClassSaveType = c_rs("FileSaveType")
				c_SavePath= c_rs("SavePath")
				if G_VIRTUAL_ROOT_DIR<>"" then
					c_SavePath = "/"& G_VIRTUAL_ROOT_DIR&c_SavePath
				end if
				if ClassSaveType=0 then
					class_savepath = ClassEName&"/Index."&FileExtName
				elseif ClassSaveType=1 then
					class_savepath = ClassEName&"/"& ClassEName &"."&FileExtName
				else
					class_savepath = ClassEName &"."&FileExtName
				end if
				if LinkType = 1 then
					if trim(c_Domain)<>"" then
						Url_Domain = "http://"&c_Domain
					else
						if trim(IsDomain)<>"" then
							Url_Domain = "http://"&IsDomain
						else
							Url_Domain = "http://"&MF_Domain
						end if
					end if
				else
					if trim(c_Domain)<>"" then
						Url_Domain = "http://"&c_Domain
					else
						if trim(IsDomain)<>"" then
							Url_Domain = "http://"&IsDomain
						else
							Url_Domain = ""
						end if
					end if
				end if
				'判断域名是否为空为空就不处理，不为空就检测程序前台目录并过滤掉根目录
				If Url_Domain = "" Or Isnull(Url_Domain) Then
					get_ClassLink = Url_Domain&Replace(c_SavePath&"/"&class_savepath,"//","/")
				Else
					Dim Savepathlen,SubPathlen
					SubPathlen = Len("/"&Conn.execute("select Sub_Sys_Path from FS_MF_Sub_Sys Where Sub_Sys_ID='DS'")(0))
					Savepathlen = Len(c_SavePath)
					c_SavePath = Right(c_SavePath,Savepathlen - SubPathlen)
					get_ClassLink = Url_Domain&Replace(c_SavePath&"/"&class_savepath,"//","/")
				End if
				'结束
			else
				get_ClassLink = c_rs("UrlAddress")
			end if
			c_rs.close:set c_rs=nothing
		else
			c_rs.close:set c_rs=nothing
			get_ClassLink = ""
		end if
		get_ClassLink = get_ClassLink
	End function
	
	'得到专题地址________________________________________________________________
	Public function get_specialLink(f_id)
		dim IsDomain,LinkType,MF_Domain,c_rs,SpecialEName,ExtName,c_SavePath,Url_Domain
		LinkType = Request.Cookies("FoosunNSCookies")("FoosunNSLinkType")
		IsDomain = Request.Cookies("FoosunNSCookies")("FoosunNSDomain")
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		set c_rs=conn.execute("select SpecialID,SpecialCName,SpecialEName,SavePath,FileExtName,isLock From FS_DS_Special where SpecialEName='"&f_id&"'")
		if not c_rs.eof then
			SpecialEName = c_rs("SpecialEName")
			ExtName = c_rs("FileExtName")
			c_SavePath= c_rs("SavePath")
			if LinkType = 1 then
				if trim(IsDomain)<>"" then
					Url_Domain = "http://"&IsDomain
				else
					Url_Domain = "http://"&MF_Domain
				end if
			else
				if trim(IsDomain)<>"" then
					Url_Domain = "http://"&IsDomain
				else
					if G_VIRTUAL_ROOT_DIR<>"" then
						Url_Domain = "/"& G_VIRTUAL_ROOT_DIR
					else
						Url_Domain = ""
					end if
				end if
			end if
			get_specialLink = Url_Domain&Replace(c_SavePath&"/special_"&SpecialEName&"."&ExtName,"//","/")
			c_rs.close:set c_rs=nothing
		else
			c_rs.close:set c_rs=nothing
			get_specialLink = "找不到参数，错误的地址"
		end if
		get_specialLink = get_specialLink
	End function
	
	'得到子类____________________________________________________________________
	Public Function get_ClassList(TypeID,CompatStr,f_css)
		Dim ChildTypeListRs,ChildTypeListStr,TempStr,i,Sql
		Set ChildTypeListRs = Server.CreateObject(G_FS_RS)
		Sql = "Select ParentID,ClassID,ClassName from FS_DS_Class where ParentID='" & TypeID & "' and ReycleTF=0 order by OrderID desc,id desc"
		ChildTypeListRs.open sql,Conn,0,1
		TempStr = CompatStr
		do while Not ChildTypeListRs.Eof
			get_ClassList = get_ClassList & TempStr
			get_ClassList = get_ClassList & "<img src="""&m_PathDir&"sys_images/-.gif"" border=""0""><a href="""&get_ClassLink(ChildTypeListRs("ClassId"))&""" class="""& f_css &""">"&ChildTypeListRs("ClassName")&"</a>"&chr(10)
			get_ClassList = get_ClassList & get_ClassList(ChildTypeListRs("ClassID"),TempStr,f_css)
			ChildTypeListRs.MoveNext
		loop
		ChildTypeListRs.Close:Set ChildTypeListRs = Nothing
	End Function

	'得到option子类______________________________________________________________
	Public Function get_optionNewsList(TypeID,CompatStr)
		Dim ChildTypeListRs,ChildTypeListStr,TempStr
		Set ChildTypeListRs = Conn.execute("Select ParentID,ClassID,ClassName from FS_DS_Class where ParentID='" & TypeID & "' and ReycleTF=0 order by OrderID desc,id desc")
		TempStr = CompatStr & "┉"
		do while Not ChildTypeListRs.Eof
			get_optionNewsList = get_optionNewsList &"<option value="""&ChildTypeListRs("ClassId")&""">"& TempStr
			get_optionNewsList = get_optionNewsList & "┉"&ChildTypeListRs("ClassName")&"</option>"&chr(10)
			get_optionNewsList = get_optionNewsList & get_optionNewsList(ChildTypeListRs("ClassID"),TempStr)
			ChildTypeListRs.MoveNext
		loop
		ChildTypeListRs.Close:Set ChildTypeListRs = Nothing
	End Function
	Public Function getNewsSubClass(typeID)
		Dim childClassRs,Str_SubClassID
		Set childClassRs=Conn.execute("Select ParentID,ClassID from FS_DS_Class where ParentID='" & TypeID & "' and ReycleTF=0 order by OrderID desc,id desc")
		While Not childClassRs.eof
			Str_SubClassID=getNewsSubClass(childClassRs("classID"))
			if Str_SubClassID <> "" then
				getNewsSubClass = getNewsSubClass & "," &childClassRs("classID")&"" & "," & Str_SubClassID
			else
				getNewsSubClass = getNewsSubClass & "," &childClassRs("classID")&""
			end if
			childClassRs.movenext
		Wend
		childClassRs.close:Set childClassRs=nothing
	End Function
	'====================相关下载标签
	Public Function downrelative(f_Lable,f_Type,f_Id)
		IF f_Id = "" Then
			f_Lable.DictLableContent.Add "1","不能够在非下载信息页面插入此标签"
			Exit Function
		End If
		Dim C_Ruler,C_Num,TitleNum,DateStr,ConNum,PicTF,DateNum,OrderStr,DescStr
		Dim OrderDescStr,ThisDTitle,ThisDAuth,GetThisObj,TrueDateStr,RulerSqlStr,OverDueStr
		Dim CSql,CRs,MF_Domain,IDStr
		Dim TitleStr,AuthStr
		C_Ruler = f_Lable.LablePara("相关条件")
		C_Num = f_Lable.LablePara("显示数量")
		TitleNum = f_Lable.LablePara("标题字数")
		ConNum = f_Lable.LablePara("简介字数")
		PicTF = f_Lable.LablePara("图文标记")
		DateNum = f_Lable.LablePara("日期范围")
		OrderStr = f_Lable.LablePara("排序字段")
		DescStr = f_Lable.LablePara("排序方式")
		DateStr = f_Lable.LablePara("日期格式")
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		'---判断数字型参数是否合法
		If TitleNum = "" Or Not IsNumeric(TitleNum) Then : TitleNum = 40 : Else : If Cint(TitleNum) < 1 Then : TitleNum = 40 : Else : TitleNum = Cint(TitleNum) : End If : End If
		If C_Num = "" Or Not IsNumeric(C_Num) Then : C_Num = 10 : Else : If Cint(C_Num) < 1 Then : 	C_Num = 10 : Else : C_Num = Cint(C_Num) : End If : End If
		If ConNum = "" Or Not IsNumeric(ConNum) Then : ConNum = 100 : Else : If Cint(ConNum) < 1 Then :	ConNum = 100 : Else : ConNum = Cint(ConNum) : End If : End If
		If DateNum = "" Or Not IsNumeric(DateNum) Then : DateNum = 0 : Else : If Cint(DateNum) <= 0 Then : 	DateNum = 0 : Else : DateNum = Cint(DateNum) : End If : End If
		'---排序方式
		IF LCase(OrderStr) <> "id" Then
			OrderDescStr = " Order By " & OrderStr & " " & DescStr & ",ID " & DescStr & ""
		Else
			OrderDescStr = " Order By " & OrderStr & " " & DescStr & ""
		End IF
		'---累加查询条件
		Set GetThisObj = Conn.ExeCute("Select [ID],[Name],[Provider] From FS_DS_List Where DownLoadID = '" & f_Id & "'")
		IF GetThisObj.Eof Then
			f_Lable.DictLableContent.Add "1","暂无相关内容"
			Exit Function
		Else
			IDStr = GetThisObj(0)
			ThisDTitle = GetThisObj(1)
			ThisDAuth = GetThisObj(2) 
		End If
		GetThisObj.Close : Set GetThisObj = Nothing
		If ThisDTitle <> "" Then
			TitleStr = " And ([Name] Like '%" & ThisDTitle & "%'"
			If G_IS_SQL_DB = 1 Then
				TitleStr = TitleStr & " Or charindex([Name],'" & ThisDTitle & "') > 0)"
			Else
				TitleStr = TitleStr & " Or Instr('" & ThisDTitle & "',[Name]) > 0)"
			End If
		End If
		If ThisDAuth <> "" Then
			AuthStr = " And ([Provider] Like '%" & ThisDAuth & "%'"
			If G_IS_SQL_DB = 1 Then
				AuthStr = AuthStr & " Or charindex([Provider],'" & ThisDAuth & "') > 0)"
			Else
				AuthStr = AuthStr & " Or Instr('" & ThisDAuth & "',[Provider]) > 0)"
			End If
		End If
		IF IDStr <> "" And Not IsNull(IDStr) Then
			IDStr = " And ID <> " & Clng(IDStr)
		Else
			IDStr = ""
		End If	
		If C_Ruler = 0 Then
			RulerSqlStr = TitleStr & IDStr
		Else
			RulerSqlStr = AuthStr & IDStr
		End If
			
		IF Cint(DateNum) = 0 Then
			TrueDateStr	= ""
		Else
			If G_IS_SQL_DB = 1 Then
				TrueDateStr = " And DateDiff(d,addtime,getdate()) <= " & Cint(DateNum)
			Else
				TrueDateStr = " And DateDiff('d',addtime,Now()) <= " & Cint(DateNum)
			End If	
		End If
		If G_IS_SQL_DB = 1 Then
			OverDueStr = " And DateAdd(d,OverDue,addtime) <= getdate()"
		Else
			OverDueStr = " And DateAdd('d',OverDue,addtime) <= now()"
		End If	
		'---查询语句
		CSql="select top "& cint(C_Num) &" ID,DownLoadID,ClassID,Description,Accredit,AddTime,Appraise,AuditTF,BrowPop," _
				&"ClickNum,EditTime,EMail,SavePath,FileExtName,FileName,FileSize,[Language],Name,NewsTemplet,PassWord,Pic," _
				&"Property,Provider,ProviderUrl,RecTF,ReviewTF,ShowReviewTF,SystemType,Types,Version,OverDue,"_
				&"ConsumeNum,Hits,speicalId"
		CSql = CSql & " From FS_DS_List where AuditTF=1" & OverDueStr & TrueDateStr & RulerSqlStr & OrderDescStr
		Set CRs = Conn.ExeCute(CSql)
		If CRs.Eof Then
			f_Lable.DictLableContent.Add "1","暂无相关内容"
			Exit Function
		Else
			Do While NOt CRs.Eof
				downrelative = downrelative & getlist_news(CRs,f_Lable.FoosunStyle.StyleContent,TitleNum,ConNum,"",PicTF,DateStr,"0",MF_Domain,"down_relative")
			CRs.MoveNExt
			Loop
		End If	 
		CRs.Close : Set CRs = NOthing
		f_Lable.DictLableContent.Add "1",downrelative
	End Function
End Class
%>
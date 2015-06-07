<!--#include file="../FS_Inc/adovbs.inc" -->
<%
Class cls_NS
	Private m_Rs,m_FSO,m_Dict,m_NewsLinkFields,m_ClassLinkFields,m_SpecialLinkFields
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
		m_Path_Templet  = replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_TEMPLETS_DIR&"/","//","/")
		m_NewsLinkFields = "ClassEName,Domain,SavePath,IsURL,URLAddress,SaveNewsPath,FileName,FileExtName"
		m_ClassLinkFields = "IsURL,ClassEName,Domain,FileExtName,FileSaveType,SavePath,UrlAddress"
		m_SpecialLinkFields = "SpecialEName,ExtName,SavePath,FileSaveType"
	End Sub
	Private Sub Class_Terminate()
		Set m_Rs = Nothing
		Set m_FSO = Nothing
		Set m_Dict = Nothing
	End Sub
	Public Function get_LableChar(f_Lable,f_Id,f_Type)
		select case LCase(f_Lable.LableFun)
			case "classpage"
				get_LableChar = ClassPage(f_Lable,"classpage",f_Id)
			case "classnews"
				get_LableChar = ClassNews(f_Lable,"classnews",f_Id)
			case "specialnews"
				get_LableChar = ClassNews(f_Lable,"specialnews",f_Id)
			case "lastnews"
				get_LableChar = ClassNews(f_Lable,"lastnews",f_Id)
			case "hotnews"
				get_LableChar = ClassNews(f_Lable,"hotnews",f_Id)
			case "recnews"
				get_LableChar = ClassNews(f_Lable,"recnews",f_Id)
			case "marnews"
				get_LableChar = ClassNews(f_Lable,"marnews",f_Id)
			case "brinews"
				get_LableChar = ClassNews(f_Lable,"brinews",f_Id)
			case "annnews"
				get_LableChar = ClassNews(f_Lable,"annnews",f_Id)
			case "constrnews"
				get_LableChar = ClassNews(f_Lable,"constrnews",f_Id)
			case "classlist"
				get_LableChar = ClassList(f_Lable,"classlist",f_Id)
			case "speciallist"
				get_LableChar = ClassList(f_Lable,"speciallist",f_Id)
			case "flashfilt"
				get_LableChar = flashfilt(f_Lable,"flashfilt",f_Id)
			case "norfilt"
				get_LableChar = NorFilter(f_Lable,"NorFilt",f_Id)
			case "readnews"
				get_LableChar = ReadNews(f_Lable,"readnews",f_Id)
			case "sitemap"
				get_LableChar = SiteMap(f_Lable,"sitemap",f_Id)
			case "search"
				get_LableChar = Search(f_Lable,"search")
			case "infostat"
				get_LableChar = infoStat(f_Lable,"infostat")
			case "todaypic"
				get_LableChar = TodayPic(f_Lable,"todaypic",f_Id)
			case "todayword"
				get_LableChar = TodayWord(f_Lable,"todayword",f_Id)
			case "classnavi"
				get_LableChar = ClassNavi(f_Lable,"classnavi",f_Id)
			case "specialnavi"
				get_LableChar = SpecialNavi(f_Lable,"specialnavi",f_Id)
			case "rssfeed"
				get_LableChar = RssFeed(f_Lable,"rssfeed",f_Id)
			case "specialcode"
				get_LableChar = SpecialCode(f_Lable,"specialcode",f_Id)
			case "classcode"
				get_LableChar = ClassCode(f_Lable,"classcode",f_Id)
			case "definenews"
				get_LableChar = DefineNews(f_Lable,"definenews",f_Id)
			case "oldnews"
				get_LableChar = OldNews(f_Lable,"oldnews",f_Id)
			case "c_news"
				get_LableChar = c_news(f_Lable,"c_news",f_Id)
			case "subclasslist"
				get_LableChar = subClassList(f_Lable,"subclasslist",f_Id)
			case "allcode"
				get_LableChar = AllCode(f_Lable,"allcode",f_Id)
			case "classinfo"
				get_LableChar = ClassInfo(f_Lable,"ClassInfo",f_Id)
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

	Public Function table_str_list_head(f_tf,f_divid,f_divclass,f_ulid,f_ulclass)
		Dim table_,tr_
		Dim f_divid_1,f_divclass_1,f_ulid_1,f_ulclass_1
		if f_tf=1 then
			if f_divid<>"" then:f_divid_1 = " id="""& f_divid &"""":else:f_divid_1 = "":end if
			if f_divclass<>"" then:f_divclass_1 = " class="""& f_divclass &"""":else:f_divclass_1 = "":end if
			if f_ulid<>"" then:f_ulid_1 = " id="""& f_ulid &"""":else:f_ulid_1 = "":end if
			if f_ulclass<>"" then:f_ulclass_1 = " class="""& f_ulclass &"""":else:f_ulclass_1 = "":end if
			if f_divid="0" and f_divclass="0" then:table_="":nodiv=true:else:table_="<div"&f_divid_1&f_divclass_1&">"&tr_:end if
			if f_ulid="0" and f_ulclass="0" then:tr_="":noul=true:else:tr_="<ul"& f_ulid_1 & f_ulclass_1 & ">" & tr_:end if
			table_str_list_head =  table_
			table_str_list_head = table_str_list_head &" "& tr_
		else
			table_="<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
			table_str_list_head =  table_
		end if
	End Function

	public Function table_str_list_middle_1(f_tf,f_liid,f_liclass)
		Dim f_liid_1,f_liclass_1,td_
		if f_tf=1 then
			if f_liid<>"" then:f_liid_1 = " id="""& f_liid &"""":else:f_liid_1 = "":end if
			if f_liclass<>"" then:f_liclass_1 = " class="""& f_liclass &"""":else:f_liclass_1 = "":end if
			if f_liid="0" and f_liclass="0" then:td_="":noli=true:else:td_="<li"&f_liid_1&f_liclass_1&">":end if
			table_str_list_middle_1 ="  "&td_
		end if
	End Function

	public Function table_str_list_middle_2(f_tf)
		Dim td__,tr__
		if f_tf=1 then
			td__="</li>"
		else
			td__="</td>"
		end if
			table_str_list_middle_2 =  td__&vbNewLine
	End Function

	Public Function table_str_list_middle_3(f_tf)
		if f_tf=1 then
			table_str_list_middle_3 = ""
		else
			table_str_list_middle_3 = "</tr>"
		end if
	End Function

	Public Function table_str_list_bottom(f_tf)
		Dim table__,tr__
		if f_tf=1 then
			table__="</div>"
			tr__="</ul>"
			table_str_list_bottom = " "&tr__&vbNewLine
		else
			table__="</table>"
			table_str_list_bottom = ""
		end if
		table_str_list_bottom = table_str_list_bottom &table__&vbNewLine
	End Function

	Public Function AllCode(f_Lable,f_type,f_Id)
		Dim MF_Domain,LableType,CodeNumer,LeftTitle,ClickNum,PicTF,ColsNumer,DateStyle,contentNum,NaviNum,div_tf
		Dim SearSQL,f_rs_obj,f_sql,picnewstf,orderby,HotStr,Content_List,tjstr
		Dim classNews_head,classNews_middle1,classNews_bottom,classNews_middle2,c_i_k
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		LableType = f_Lable.LablePara("类型")
		CodeNumer = f_Lable.LablePara("调用数")
		if not isnumeric(CodeNumer) then:CodeNumer = 10:else:CodeNumer = cint(CodeNumer):end if
		LeftTitle = f_Lable.LablePara("标题数")
		if not isnumeric(LeftTitle) then:LeftTitle = 40:else:LeftTitle = cint(LeftTitle):end if
		ClickNum = f_Lable.LablePara("点击数")
		PicTF = f_Lable.LablePara("图片新闻")
		ColsNumer = f_Lable.LablePara("排列数")
		DateStyle = f_Lable.LablePara("日期格式")
		contentNum = f_Lable.LablePara("内容显示字数")
		if not isnumeric(contentNum) then:contentNum = 200:else:contentNum = cint(contentNum):end if
		NaviNum = f_Lable.LablePara("导航内容显示字数")
		if not isnumeric(NaviNum) then:NaviNum = 200:else:NaviNum = cint(NaviNum):end if
		if f_Lable.LablePara("输出格式") = "out_DIV" then:div_tf = 1:else:div_tf = 0:end if
		
		f_DivID = f_Lable.LablePara("DivID")
		f_DivClass = f_Lable.LablePara("DivClass")
		f_UlID = f_Lable.LablePara("ulid")
		f_ULClass = f_Lable.LablePara("ulclass")
		f_LiID = f_Lable.LablePara("liid")
		f_LiClass = f_Lable.LablePara("liclass")
		classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		
		if f_id<>"" then
			dim news_rs
			set news_rs = Conn.execute("select Id From FS_NS_NewsClass where ClassID='"& f_id &"'")
			if news_rs.eof then
				SearSQL = ""
				news_rs.close:set news_rs=nothing
			else
				SearSQL = " and News.ClassId='"& f_id &"'"
				news_rs.close:set news_rs=nothing
			end if
		else
			SearSQL = ""
		end if
		if PicTF = "1" then
			picnewstf = "  and isPicNews=1"
		else
			picnewstf = ""
		end if 
		if LableType = "HotNews" then
			orderby = " order by Hits desc,News.addtime desc"
			if ClickNum<>"0" then
				HotStr = " and Hits>="& clng(ClickNum) &""
			else
				HotStr = ""
			end if
		else
			orderby = " order by News.Popid desc,News.addtime desc"
			HotStr = ""
		end if
		select Case LableType
			Case "HotNews"
				tjstr = ""
			Case "LastNews"
				tjstr = ""
			Case "RecNews"
				tjstr = " and "& all_substring &"(NewsProperty,1,1)='1'"
			Case "MarqueeNews"
				tjstr = " and "& all_substring &"(NewsProperty,3,1)='1'"
			Case "JcNews"
				tjstr = " and "& all_substring &"(NewsProperty,15,1)='1'"
			Case else
				tjstr = ""
		end select
		f_sql="select top "& cint(CodeNumer) &" News.ID,NewsId,News.PopId,News.ClassID,SpecialEName,NewsTitle,CurtTitle,NewsNaviContent,isShowReview,TitleColor,titleBorder,TitleItalic,News.IsURL,News.URLAddress"
		f_sql = f_sql &",Content,isPicNews,NewsPicFile,NewsSmallPicFile,News.isPop,Source,Editor,Keywords,Author,Hits,SaveNewsPath,News.FileName,News.FileExtName,NewsProperty,isLock,News.addtime,TodayNewsPic,ClassEName,[Domain],SavePath "
		f_sql = f_sql &"From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID "& picnewstf & SearSQL & HotStr & tjstr &" and  isRecyle=0 and isdraft=0 "&orderby&",News.id desc"
		set f_rs_obj = Conn.execute(f_sql)
		
		if f_rs_obj.eof then
			Content_List=""
			f_rs_obj.close:set f_rs_obj=nothing
		else
			if div_tf = 0 then
				c_i_k = 0
				if cint(ColsNumer)<>1 then
					Content_List = Content_List &  "  <tr>"
				end if
			end if
			do while not f_rs_obj.eof
				if div_tf=1 then
					Content_List= Content_List & classNews_middle1 & getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,LeftTitle,contentNum,NaviNum,0,DateStyle,0,MF_Domain,f_type,"") & classNews_middle2
				else
					if cint(ColsNumer) =1 then
						Content_List= Content_List & vbNewLine&"   <tr><td>" & getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,LeftTitle,contentNum,NaviNum,0,DateStyle,0,MF_Domain,f_type,"") & "</td></tr>"
					else
						Content_List= Content_List & "<td width="""& cint(100/cint(ColsNumer))&"%"">" & getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,LeftTitle,contentNum,NaviNum,0,DateStyle,0,MF_Domain,f_type,"") & "</td>"
					end if
				end if
				f_rs_obj.movenext
				if div_tf = 0 then
					if cint(ColsNumer)<>1 then
						c_i_k = c_i_k+1
						if c_i_k mod cint(ColsNumer) = 0 then
							Content_List = Content_List & "</tr>"&vbNewLine&"  <tr>"
						end if
					end if
				end if
		   loop
		   if div_tf=0 then
				if cint(ColsNumer)<>1 then
					Content_List = Content_List & "</tr>"&vbNewLine
				end if
		   end if
			f_rs_obj.close:set f_rs_obj=nothing
			Content_List=classNews_head & Content_List & classNews_bottom
		end if 
		f_Lable.DictLableContent.Add "1",Content_List
	End Function
	
	Public Function ClassInfo(f_Lable,f_type,f_Id)
		Dim LableType,Content_List,InfoType
		Dim f_rs_obj,f_sql
		InfoType = f_Lable.LablePara("调用内容")
		if f_id<>"" then
			Select Case InfoType
				Case "ClassName"
					f_sql="select ClassName From FS_NS_NewsClass where ClassID='"& f_id &"'"
				Case "Keywords"
					f_sql="select ClassKeywords From FS_NS_NewsClass where ClassID='"& f_id &"'"
				Case "Description"
					f_sql="select ClassDescription From FS_NS_NewsClass where ClassID='"& f_id &"'"
			End Select
			set f_rs_obj = Conn.execute(f_sql)
			if Not f_rs_obj.eof then
				Content_List = f_rs_obj(0)
			else
				Content_List = ""
			end if
			f_rs_obj.close:set f_rs_obj=nothing
		else
			Content_List = ""
		End if
		f_Lable.DictLableContent.Add "1",Content_List
	End Function

	Public Function ClassNews(f_Lable,f_Type,f_Id)
		f_Lable.IsSave = True
		Dim classNews_head,classNews_bottom,classNews_middle1,classNews_middle2,div_tf,style_Content,Content_List,f_rs_s_obj,Content_more,c_i_k
		Dim newnumber,classid,orderby,orderdesc,colnumber,contentnumber,navinumber,datenumber,titlenumber,picshowtf,datenumber_tmp,morechar,datestyle,openstyle,open_target,containSubClass,childClass
		Dim f_sql,f_rs_obj,f_rs_configobj,f_configSql,CharIndexStr
		Dim MF_Domain,marqueedirec,marqueespeed,marqueestyle
		Dim search_str,str_order,temp_intListID,SpecialID
		Dim ClassName,ClassEName,ClassNaviContent,ClassNaviPic,c_SavePath,c_FileSaveType,search_inSQL
		Dim Exist_Child_ClassID,picnewstf,str_Liststyle,int_listID,f_ClassLinkRecordSet,all_savepath
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
		search_str = f_Lable.LablePara("栏目")
		SpecialID = f_Lable.LablePara("专题")
		picnewstf = f_Lable.LablePara("图片新闻")
		newnumber = f_Lable.LablePara("Loop")
		datenumber = f_Lable.LablePara("多少天")
		titlenumber = f_Lable.LablePara("标题数")
		picshowtf = f_Lable.LablePara("图文标志")
		openstyle = f_Lable.LablePara("打开窗口")
		containSubClass = f_Lable.LablePara("包含子类")
		orderby = f_Lable.LablePara("排列字段")
		orderdesc = f_Lable.LablePara("排列方式")
		morechar = f_Lable.LablePara("更多连接")
		datestyle = f_Lable.LablePara("日期格式")
		colnumber = f_Lable.LablePara("新闻排列数")
		contentnumber = f_Lable.LablePara("内容字数")
		navinumber = f_Lable.LablePara("导航字数")
		str_Liststyle = f_Lable.LablePara("列表序号")
		marqueespeed = f_Lable.LablePara("滚动速度")
		marqueedirec = f_Lable.LablePara("滚动方向")
		f_DivID = f_Lable.LablePara("DivID")
		f_DivClass = f_Lable.LablePara("DivClass")
		f_UlID = f_Lable.LablePara("ulid")
		f_ULClass = f_Lable.LablePara("ulclass")
		f_LiID = f_Lable.LablePara("liid")
		f_LiClass = f_Lable.LablePara("liclass")
		if f_Lable.LablePara("输出格式") = "out_DIV" then
			div_tf=1
		else
			div_tf=0
		end if
		if newnumber="" then newnumber = 10
		if Not isnumeric(newnumber) then newnumber = 10
		if datenumber="" then datenumber = 10
		if Not isnumeric(datenumber) then datenumber = 10
		if titlenumber = "" and isnumeric(titlenumber) = false then
			titlenumber = 30
		else
			titlenumber = titlenumber
		end if
		if colnumber="" and isnumeric(colnumber)=false then
			colnumber = 30
		else
			colnumber = colnumber
		end if
		if contentnumber="" and isnumeric(contentnumber)=false then
			contentnumber = 30
		else
			contentnumber = contentnumber
		end if
		if navinumber="" and isnumeric(navinumber)=false then
			navinumber = 30
		else
			navinumber = navinumber
		end If
		if picnewstf="1" then
			picnewstf = " and isPicNews=1"
		else
			picnewstf = ""
		end if
		if right(lcase(morechar),4)=".jpg" or right(lcase(morechar),4)=".gif" or right(lcase(morechar),4)=".png" or right(lcase(morechar),4)=".ico" or right(lcase(morechar),4)=".bmp" or right(lcase(morechar),5)=".jpeg" then
			morechar = "<img src = "&morechar&" border=""0"" />"
		end if
		
		classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)

		if G_IS_SQL_DB=0 then
			if datenumber ="0" then:datenumber_tmp = "":else:datenumber_tmp = " and datevalue(News.addtime)+"&datenumber&">=datevalue(now)":end if
		Elseif G_IS_SQL_DB=1 then
			if datenumber ="0" then:datenumber_tmp = "":else:datenumber_tmp = " and dateadd(d,"&datenumber&",News.addtime)>='"&datevalue(now())&"'":end if
		End if
		If containSubClass="1" And not InStr(1,search_str,",")>0  Then
			Exist_Child_ClassID=getNewsSubClass(search_str)
			If Exist_Child_ClassID<>"" Then
				childClass=" or News.classid in ("&DelHeadAndEndDot(Exist_Child_ClassID)&")"
			Else
				childClass=""
			End If
		Else
			childClass=""
		End If
		If search_str<>"" And InStr(1,search_str,",")>0 Then
			search_str = Replace(search_str,",","','")
		End If

		select case f_Type
			case "classnews"
				If childClass<>"" Then
					search_inSQL = " and (News.ClassId in ('"& search_str &"')"&childClass&")"
				Else
					search_inSQL = " and News.ClassId in ('"& search_str &"')"
				End if
			case "specialnews"
				if SpecialID <> "" then
					if G_IS_SQL_DB=0 then
						search_inSQL = " and instr(SpecialEName,'"&SpecialID&"')>0"
					else
						search_inSQL = " and charindex('"&SpecialID&"',SpecialEName)>0"
					end if
				else
					search_inSQL = " And ((SpecialEName Is Not Null) OR SpecialEName='')"
				end if
			case "lastnews"
				if trim(search_str)<>"" Then
					If childClass<>"" Then
						search_inSQL = " and (News.ClassId in ('"& search_str &"')"&childClass&")"
					Else
						search_inSQL = " and News.ClassId in ('"& search_str &"')"
					End if
				else
					search_inSQL = ""
				end if
			case "hotnews"
				if trim(search_str)<>"" Then
					If childClass<>"" then
						search_inSQL = " and (News.ClassId in ('"& search_str &"')"&childClass&")"
					Else
						search_inSQL = " and News.ClassId in ('"& search_str &"')"
					End if
				else
					search_inSQL = ""
				end if
				orderby = "Hits"
				orderdesc = "Desc"
			case "recnews"
				if trim(search_str)<>"" Then
					If childClass<>"" then
						search_inSQL = " and "& all_substring &"(NewsProperty,1,1)='1' and (News.ClassId in ('"& search_str &"')"&childClass&")"
					Else
						search_inSQL = " and "& all_substring &"(NewsProperty,1,1)='1' and News.ClassId in ('"& search_str &"')"
					End if
				else
					search_inSQL = " and "& all_substring &"(NewsProperty,1,1)='1'"
				end if
			case "marnews"
				if trim(search_str)<>"" Then
					If childClass<>"" then
						search_inSQL = " and "& all_substring &"(NewsProperty,3,1)='1' and (News.ClassId in ('"& search_str &"')"&childClass&")"
					Else
						search_inSQL = " and "& all_substring &"(NewsProperty,3,1)='1' and News.ClassId in ('"& search_str &"')"
					End if
				else
					search_inSQL = " and "& all_substring &"(NewsProperty,3,1)='1'"
				end if
			case "brinews"
				if trim(search_str)<>"" Then
					If childClass<>"" then
						search_inSQL = " and "& all_substring &"(NewsProperty,15,1)='1' and (News.ClassId in ('"& search_str &"')"&childClass&")"

					Else
						search_inSQL = " and "& all_substring &"(NewsProperty,15,1)='1' and News.ClassId in ('"& search_str &"')"
					End if
				else
					search_inSQL = " and "& all_substring &"(NewsProperty,15,1)='1'"
				end if
			case "annnews"
				if trim(search_str)<>"" Then
					If childClass<>"" then
						search_inSQL = " and "& all_substring &"(NewsProperty,19,1)='1' and (News.ClassId in ('"& search_str &"')"&childClass&")"
					Else
						search_inSQL = " and "& all_substring &"(NewsProperty,19,1)='1' and News.ClassId in ('"& search_str &"')"
					End if
				else
					search_inSQL = " and "& all_substring &"(NewsProperty,19,1)='1'"
				end if
			case "constrnews"
				if trim(search_str)<>"" Then
					If childClass<>"" then
						search_inSQL = " and "& all_substring &"(NewsProperty,7,1)='1' and (News.ClassId in ('"& search_str &"')"&childClass&")"
					Else
						search_inSQL = " and "& all_substring &"(NewsProperty,7,1)='1' and News.ClassId in ('"& search_str &"')"
					End if
				else
					search_inSQL = " and "& all_substring &"(NewsProperty,7,1)='1'"
				end if
		end select
		
		if orderby = "" then orderby = "ID"
		if orderdesc = "" then orderdesc = "Desc"
		If LCase(orderby) = "id" Then 
			str_order=" order by News.ID " & orderdesc
		Else
			str_order=" order by News." & orderby & " " & orderdesc & ",News.id desc"
		End If
		f_sql="select top "& cint(newnumber) &" News.ID,NewsId,News.PopId,News.ClassID,SpecialEName,NewsTitle,CurtTitle,NewsNaviContent,isShowReview,TitleColor,titleBorder,TitleItalic,News.IsURL,News.URLAddress"
		f_sql = f_sql &",Content,isPicNews,NewsPicFile,NewsSmallPicFile,News.isPop,Source,Editor,Keywords,Author,Hits,SaveNewsPath,News.FileName,News.FileExtName,NewsProperty,isLock,News.addtime,TodayNewsPic,ClassEName,[Domain],SavePath "
		f_sql = f_sql &"From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID "& picnewstf &search_inSQL & datenumber_tmp &" and  isRecyle=0 And isLock=0 and isdraft=0" & str_order
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		Content_List=""
		set f_rs_obj = Server.CreateObject(G_FS_RS)
		f_rs_obj.Open f_sql,Conn,0,1
		int_listID=0
		if f_rs_obj.eof then
			Content_List=""
			f_rs_obj.close:set f_rs_obj=nothing
		else
			if f_Type="marnews" then
				dim isupdown
				if marqueedirec="up" then
					Content_List = Content_List & ""
					Content_List = Content_List &"<div id=""annbody""><ul id=""annbodylis"">"
					isupdown="1"
				else
					Content_List = Content_List & "<marquee onmouseover=""this.stop();"" scrollamount="""& marqueespeed &""" direction="""& marqueedirec &""" onmouseout=""this.start();"">"
				end if
				Select Case str_Liststyle
					Case "1"
						int_listID=64
						do while not f_rs_obj.eof
							If Cint(newnumber)<=26 Then
								int_listID=int_listID+1
								Content_List= Content_List & Chr(int_listID)&"." & getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,f_Type,"")
							Else
								Content_List= Content_List  &"<LI>"& getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,f_Type,"")&"</LI>"
							End If
							if isupdown="1" then
								'Content_List= Content_List&"<br />"
							end if
							f_rs_obj.movenext
						loop
					Case "2"
						int_listID=96
						do while not f_rs_obj.eof
							If Cint(newnumber)<=26 Then 
								int_listID=int_listID+1
								Content_List= Content_List & Chr(int_listID)&"." & getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,f_Type,"")
							Else
								Content_List= Content_List  &"<LI>"& getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,f_Type,"")&"</LI>"
							End If
							if isupdown="1" then
								'Content_List= Content_List&"<br />"
							end if
							f_rs_obj.movenext
						loop
					Case "3"
						do while not f_rs_obj.eof
							int_listID=int_listID+1
							Content_List= Content_List & int_listID&"." &"<LI>"& getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,f_Type,"")&"</LI>"
							if isupdown="1" then
								'Content_List= Content_List&"<br />"
							end if
							f_rs_obj.movenext
						loop										
					Case Else
						do while not f_rs_obj.eof
							Content_List= Content_List  &"<LI>"& getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,f_Type,"")&"</LI>"
							if isupdown="1" then
								Content_List= Content_List&"<br />"
							end if
							f_rs_obj.movenext
						loop
				End Select
				if isupdown="1" then
					Content_List = Content_List
					Content_List = Content_List &"</ul></div><script type=""text/javascript"">announcementScroll();</script>"
				else
					Content_List = Content_List & "</marquee>"
				end if
			else
				if div_tf = 0 then
					c_i_k = 0
					if cint(colnumber)<>1 then Content_List = Content_List &  "  <tr>"
				end if
				Rem 这里是 给新闻加排序的
				Select Case str_Liststyle
					Case "1"
						int_listID=64
					Case "2"
						int_listID=96
					Case "3"
						int_listID=0
				End Select
				Rem 排序结束
				do while not f_rs_obj.eof
					temp_intListID=""
					int_listID=int_listID+1
					Select Case str_Liststyle
						Case "1"
							If Cint(newnumber)<=26 Then 
								temp_intListID=Chr(int_listID)&"."
							Else
								temp_intListID=""
							End If 											
						Case "2"
							If Cint(newnumber)<=26 Then 
								temp_intListID=Chr(int_listID)&"."	
							Else
								temp_intListID=""
							End If										
						Case "3"
							temp_intListID=int_listID&"."										
						Case Else
							temp_intListID=""
					End Select
					if div_tf=1 then
						Content_List= Content_List & classNews_middle1 & temp_intListID & getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,f_Type,"") & classNews_middle2
					else
						if cint(colnumber) =1 then
							Content_List= Content_List & vbNewLine&"   <tr><td>" & temp_intListID & getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,f_Type,"") & "</td></tr>"
						else
							Content_List= Content_List & "<td width="""& cint(100/cint(colnumber))&"%"">" & temp_intListID & getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,f_Type,"") & "</td>"
						end if
					end if
					f_rs_obj.movenext
					if div_tf = 0 then
						if cint(colnumber)<>1 then
						c_i_k = c_i_k+1
							if c_i_k mod cint(colnumber) = 0 then
								Content_List = Content_List & "</tr>"&vbNewLine&"  <tr>"
							end if
						end if
					end if
			   loop
			   if div_tf=0 then
					if cint(colnumber)<>1 then
						Content_List = Content_List & "</tr>"&vbNewLine
					end if
			   end if
			end if
			'得到栏目路径
			if f_Type="classnews" then
				dim Query_rs,newsclass_SavePath,FileSaveType,UrlDomain
				set Query_rs=Conn.execute("select ClassEName,SavePath,FileExtName,[Domain],FileSaveType,IsURL,UrlAddress From FS_NS_NewsClass where ClassId in ('"& search_str &"')")
				if Query_rs.eof then
					Query_rs.close
					set Query_rs=nothing
				else
					Set f_ClassLinkRecordSet = New CLS_FoosunRecordSet
					Set f_ClassLinkRecordSet.Values(m_ClassLinkFields) = Query_rs
					all_savepath = get_ClassLink(f_ClassLinkRecordSet)
					Set f_ClassLinkRecordSet = Nothing
					Query_rs.close:set Query_rs=nothing
				end if
				if openstyle=0 then
					open_target=" "
				else
					open_target=" target=""_blank"""
				end if
				if div_tf=1 then
					if morechar<>"" then Content_more = "  <li><a href="""&  all_savepath &""" "& open_target &" title=""更多..."">"&morechar&"</a></li>"&vbNewLine
				else
					if morechar<>"" then Content_more = "  <tr><td colspan="""& cint(colnumber) &"""><div align=""right""><a href="""&  all_savepath &""" "& open_target &" title=""更多..."">"&morechar&"</a></div></td></tr>"&vbNewLine
				end if
			elseif f_Type="specialnews" then
				dim special_rs,special_Path
				set special_rs =Conn.execute("select SpecialID,SpecialCName,SpecialEName,SpecialContent,SavePath,ExtName,isLock,naviPic,FileSaveType From FS_NS_Special where 1=1 and SpecialEName='"&trim(search_str)&"'")
				if Not special_rs.eof then
					Set f_SpecialLinkRecordSet = New CLS_FoosunRecordSet
					Set f_SpecialLinkRecordSet.Values(m_SpecialLinkFields) = special_rs
					special_Path = get_specialLink(f_SpecialLinkRecordSet)
					special_rs.close:set special_rs=nothing
				end if
				if div_tf=1 then
					if morechar<>"" then Content_more = "  <li><a href="""&  special_Path &""" "& open_target &" title=""更多..."">"&morechar&"</a></li>"&vbNewLine
				else
					if morechar<>"" then Content_more = "  <tr><td colspan="""& cint(colnumber) &"""><div align=""right""><a href="""&  special_Path &""" "& open_target &" title=""更多..."">"&morechar&"</a></div></td></tr>"&vbNewLine
				end if
			end if
			f_rs_obj.close
			set f_rs_obj=nothing
			if f_Type="marnews" then
				Content_List= Content_List
			else
				Content_List=classNews_head & Content_List & Content_more & classNews_bottom
			end if
		end if
		f_Lable.DictLableContent.Add "1",Content_List & " "
	End Function

	Public Function ClassList(f_Lable,f_Type,f_Id)
		if f_Id<>"" then
			dim div_tf,newnumber,datenumber,titlenumber,picshowtf,openstyle,orderby,orderdesc,pageTF,pagestyle,pagenumber,pagecss,datestyle,colnumber,contentnumber,navinumber,CutNum,CutType
			dim classNews_head,classNews_bottom,classNews_middle1,classNews_middle2,search_inSQL
			dim datenumber_tmp,f_sql,f_configsql,f_rs_obj,f_rs_configobj,MF_Domain
			dim TPageNum,perPageNum,PageNum,sPageCount,cl_i
			dim rs_c,RefreshNumber,picnewstf,OrderStr,f_MorePageTypeArray,f_MorePageColor,f_MorePageType
			Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
			Dim f_TableName,f_SelectFieldNames,f_PageIndex,f_Where,f_PaginationStr,f_NewsContent
			Dim f_IDSRS,f_IDSArray,f_IDS,i,f_RecordIndex,f_ClassEName,f_Domain,f_SavePath
			m_Err_Info = ""
			picnewstf = f_Lable.LablePara("图片新闻")
			datenumber = f_Lable.LablePara("多少天")
			titlenumber = f_Lable.LablePara("标题数")
			picshowtf = f_Lable.LablePara("图文标志")
			openstyle = f_Lable.LablePara("打开窗口")
			orderby = f_Lable.LablePara("排列字段")
			orderdesc = f_Lable.LablePara("排列方式")
			pageTF = f_Lable.LablePara("分页")
			pagestyle = f_Lable.LablePara("分页样式")
			pagenumber = f_Lable.LablePara("每页数量")
			pagecss = f_Lable.LablePara("PageCSS")
			CutNum = f_Lable.LablePara("分隔数量")
			CutType = f_Lable.LablePara("分隔样式")
			datestyle = f_Lable.LablePara("日期格式")
			colnumber = f_Lable.LablePara("新闻排列数")
			contentnumber = f_Lable.LablePara("内容字数")
			navinumber = f_Lable.LablePara("导航字数")
			f_DivID = f_Lable.LablePara("DivID")
			f_DivClass = f_Lable.LablePara("DivClass")
			f_UlID = f_Lable.LablePara("ulid")
			f_ULClass = f_Lable.LablePara("ulclass")
			f_LiID = f_Lable.LablePara("liid")
			f_LiClass = f_Lable.LablePara("liclass")
			
			f_MorePageTypeArray = Split(pagestyle,",")
			if UBound(f_MorePageTypeArray) = 1 then
				f_MorePageType = f_MorePageTypeArray(0)
				f_MorePageColor = f_MorePageTypeArray(1)
			else
				f_MorePageType = "1"
				f_MorePageColor = ""
			end if
			f_PaginationStr = f_MorePageType & "," & f_MorePageColor & "," & pagecss
			f_Lable.DictLableContent.Item("0") = f_PaginationStr
			if f_Lable.LablePara("输出格式") = "out_DIV" then
				div_tf=1
			else
				div_tf=0
			end if
			
			OrderStr = "News." & orderby & " " & orderdesc
			If LCase(orderby) <> "id" Then
				OrderStr = OrderStr & ",News.ID " & orderdesc
			End If
			
			classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			
			if picnewstf="1" then
				picnewstf = " and News.isPicNews=1"
			else
				picnewstf = ""
			end if
			if f_Type="speciallist" then
				dim specaial_rs,f_spen,Pro_SpecialID
				If IsNumeric(f_Id) Then
					Pro_SpecialID=f_Id
				Else
					Pro_SpecialID=0
				End If
				set specaial_rs = Conn.execute("select SpecialEName From FS_NS_Special where SpecialID="&Pro_SpecialID&"")
				if Not specaial_rs.eof then f_spen = specaial_rs(0)
				specaial_rs.close:set specaial_rs = nothing
				if G_IS_SQL_DB=0 then
					search_inSQL = " and instr(News.SpecialEName,'"&f_spen&"')>0"
				else
					search_inSQL = " and charindex('"&f_spen&"',News.SpecialEName)>0"
				end if
				RefreshNumber = ""
			else
				set rs_c=conn.execute("select RefreshNumber,ClassEName,[Domain],SavePath from FS_NS_NewsClass where ClassId='"& f_Id &"'")
				if rs_c.eof then
					RefreshNumber = ""
					rs_c.close
					set rs_c = nothing
				else
					f_ClassEName = rs_c("ClassEName")
					f_Domain = rs_c("Domain")
					f_SavePath = rs_c("SavePath")
					if rs_c(0) = 0 then
						RefreshNumber = ""
					else
						RefreshNumber = "top "&rs_c(0)&""
					end if
					rs_c.close
					Set rs_c=nothing
				end if
				search_inSQL=" and News.ClassId='"& f_Id &"'"
			end if
			if G_IS_SQL_DB=0 then
				if datenumber ="0" then:datenumber_tmp = "":else:datenumber_tmp = " and datevalue(News.addtime)+"&datenumber&">=datevalue(now)":end if
			Elseif G_IS_SQL_DB=1 then
				if datenumber ="0" then:datenumber_tmp = "":else:datenumber_tmp = " and dateadd(d,"&datenumber&",News.addtime)>='"&datevalue(now())&"'":end if
			End if
			MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
			PageNumber = CheckTextNumber(PageNumber,30)
			CutNum = CheckTextNumber(CutNum,0)
			colnumber = CheckTextNumber(colnumber,1)
			if f_Type="speciallist" then
				f_Where = "News.ClassID=Class.ClassID" & search_inSQL & picnewstf & datenumber_tmp &" and isRecyle=0 and isLock=0 and isdraft=0"
				f_TableName = "FS_NS_News as News,FS_NS_NewsClass as Class"
				f_SelectFieldNames = "News.ID,NewsId,News.PopId,SpecialEName,News.ClassID,NewsTitle,CurtTitle,NewsNaviContent,isShowReview,TitleColor,titleBorder,TitleItalic,News.IsURL,News.URLAddress,Content,isPicNews,NewsPicFile,NewsSmallPicFile,News.isPop,Source,Editor,Keywords,Author,Hits,SaveNewsPath,News.FileName,News.FileExtName,NewsProperty,isLock,News.addtime,TodayNewsPic,ClassEName,[Domain],SavePath"
			else
				f_Where = "1=1 " & search_inSQL & picnewstf & datenumber_tmp &" and isRecyle=0 and isLock=0 and isdraft=0"
				f_TableName = "FS_NS_News as News"
				f_SelectFieldNames = "News.ID,NewsId,News.PopId,SpecialEName,News.ClassID,NewsTitle,CurtTitle,NewsNaviContent,isShowReview,TitleColor,titleBorder,TitleItalic,News.IsURL,News.URLAddress,Content,isPicNews,NewsPicFile,NewsSmallPicFile,News.isPop,Source,Editor,Keywords,Author,Hits,SaveNewsPath,News.FileName,News.FileExtName,NewsProperty,isLock,News.addtime,TodayNewsPic,'" & f_ClassEName & "' as ClassEName,'" & f_Domain & "' as [Domain],'" & f_SavePath & "' as SavePath"
			end if
			f_sql = "Select " & f_SelectFieldNames & " from " & f_TableName & " where " & f_Where & " Order by " & OrderStr
			Set f_rs_obj = Server.CreateObject(G_FS_RS)
			f_rs_obj.open f_sql,Conn,0,1
			if Not f_rs_obj.Eof then
				f_PageIndex = 1
				cl_i = 1
				While Not f_rs_obj.Eof
					dim pagei
					ClassList = classNews_head
					for pagei=1 to pagenumber
						f_NewsContent = getlist_news(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,"ClassList","")
						if div_tf = 1 then
							if CutNum = 0 then
								ClassList= ClassList & classNews_middle1 & f_NewsContent & classNews_middle2
							else
								if cl_i mod CutNum = 0 then
									ClassList= ClassList & classNews_middle1 & f_NewsContent & classNews_middle2 & CutType
								else
									ClassList= ClassList & classNews_middle1 & f_NewsContent & classNews_middle2
								end if
							end If
						else
							if cint(colnumber) = 1 then
								ClassList= ClassList &"   <tr><td>" & f_NewsContent & "</td></tr>"
							else
								if cl_i mod colnumber = 1 then ClassList= ClassList & "<tr>"
								ClassList= ClassList & "<td width="""& cint(100/cint(colnumber))&"%"">" & f_NewsContent & "</td>"
								if cl_i mod colnumber = 0 then ClassList= ClassList & "</tr>"
							end if
						end if
						f_rs_obj.movenext
						cl_i = cl_i + 1
						if f_rs_obj.eof then exit for
					next
					ClassList = ClassList & classNews_bottom
					f_Lable.DictLableContent.Add f_PageIndex & "",ClassList
					f_PageIndex = f_PageIndex+1
				Wend
			else
				f_Lable.DictLableContent.Add "1",""
			end if
			f_rs_obj.Close
			Set f_rs_obj = Nothing
		else
			f_Lable.DictLableContent.Add "1",""
		end if
	End Function
	'得到相关新闻
	Public Function c_news(f_Lable,f_type,f_Id)
		Dim ifstr,titleNumber,leftTitle,f_sql,rs,str_newstitle,like_str,old_rs,TitleNumberStr,MF_Domain
		ifstr = f_Lable.LablePara("根据条件")
		TitleNumberStr= f_Lable.LablePara("显示数量")
		if not isnumeric(TitleNumberStr) then:TitleNumberStr = 10:else:TitleNumberStr = cint(TitleNumberStr):end if
		leftTitle = f_Lable.LablePara("标题字数")
		if not isnumeric(leftTitle) then:leftTitle = 40:else:leftTitle = cint(leftTitle):end if
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		c_news = ""
		'获得原新闻关键字
		dim RelateNewsSql,RsSearchObj,OldSpecialID,tmp_like,tmp_1,tmp_2
		RelateNewsSql = "Select Keywords from FS_NS_News where NewsID='" & f_Id & "' order by ID desc"
		Set RsSearchObj = Conn.Execute(RelateNewsSql)
		if Not RsSearchObj.Eof then
			If RsSearchObj("KeyWords") <> "" and isnull(RsSearchObj("KeyWords"))=false then
				Dim KeyWordsStr,KeyWordsArray,SqlKeyWordStr,keyword_j
				SqlKeyWordStr = ""

				KeyWordsStr = RsSearchObj("KeyWords")
				If KeyWordsStr<>"" and isnull(KeyWordsStr)=false then
					if instr(KeyWordsStr,",")=0 then KeyWordsStr = KeyWordsStr&","
					KeyWordsArray = split(KeyWordsStr,",")
					For keyword_j = 0 to UBound(KeyWordsArray)
					  if KeyWordsArray(keyword_j)<>"" then
						if ifstr = "0" then
							if G_IS_SQL_DB=0 then
								If SqlKeyWordStr = "" then
									SqlKeyWordStr = "instr(NewsTitle,'"&KeyWordsArray(keyword_j)&"')>0"
								Else
									SqlKeyWordStr = SqlKeyWordStr & "or instr(NewsTitle,'"&KeyWordsArray(keyword_j)&"')>0"
								End If
							else
								If SqlKeyWordStr = "" then
									SqlKeyWordStr = "charindex('"&KeyWordsArray(keyword_j)&"',NewsTitle)>0"
								Else
									SqlKeyWordStr = SqlKeyWordStr & "or charindex('"&KeyWordsArray(keyword_j)&"',NewsTitle)>0"
								End If
							end if

						elseif ifstr = "2" then
							if G_IS_SQL_DB=0 then
								If SqlKeyWordStr = "" then
									tmp_like = "instr(NewsTitle,'"&KeyWordsArray(keyword_j)&"')>0"
								Else
									tmp_like = tmp_like & "or instr(NewsTitle,'"&KeyWordsArray(keyword_j)&"')>0"
								End If
							else
								If SqlKeyWordStr = "" then
									tmp_like = "charindex('"&KeyWordsArray(keyword_j)&"',NewsTitle)>0"
								Else
									tmp_like = tmp_like & "or charindex('"&KeyWordsArray(keyword_j)&"',NewsTitle)>0"
								End If
							end if
							If SqlKeyWordStr = "" then
								SqlKeyWordStr = ""&tmp_like&" or KeyWords like '%"&KeyWordsArray(keyword_j)&"%' "
							Else
								SqlKeyWordStr = SqlKeyWordStr & " or "&tmp_like&" or KeyWords like '%"&KeyWordsArray(keyword_j)&"%' "
							End If

						elseif ifstr = "1" then
							If SqlKeyWordStr = "" then
								SqlKeyWordStr = "KeyWords like '%"&KeyWordsArray(keyword_j)&"%' "
							Else
								SqlKeyWordStr = SqlKeyWordStr & "or KeyWords like '%"&KeyWordsArray(keyword_j)&"%' "
							End If

						else
							tmp_1 = mid(ifstr,2,2)
							tmp_1 = replace(tmp_1,"==","=")
							tmp_1 = replace(tmp_1,"&lt;&gt;","<>")
							tmp_2 = mid(ifstr,4)
							if not isnumeric(tmp_2) then c_news = "错误:"&tmp_2:exit for:exit function
							if keyword_j = cint(tmp_2)-1 then
								select case tmp_1
									case "=","<>"
										SqlKeyWordStr = " KeyWords "&tmp_1&"'"&KeyWordsArray(keyword_j)&"' "
									case "*s"
										SqlKeyWordStr = " KeyWords like '%"&KeyWordsArray(keyword_j)&"' "
									case "*e"
										SqlKeyWordStr = " KeyWords like '"&KeyWordsArray(keyword_j)&"'% "
									case "**"
										SqlKeyWordStr = " KeyWords like '%"&KeyWordsArray(keyword_j)&"'% "
								end select
								exit for
							end if
						 end if

					 end if
					Next
				Else
					SqlKeyWordStr = ""
				End If

				if SqlKeyWordStr<>"" then SqlKeyWordStr = " And ("&SqlKeyWordStr&")"
				f_sql = "Select Top " & TitleNumberStr & " News.ID,NewsId,News.PopId,News.ClassID,SpecialEName,NewsTitle,CurtTitle,NewsNaviContent,isShowReview,TitleColor,titleBorder,TitleItalic,News.IsURL,News.URLAddress,Content,isPicNews,NewsPicFile,NewsSmallPicFile,News.isPop,Source,Editor,Keywords,Author,Hits,SaveNewsPath,News.FileName,News.FileExtName,NewsProperty,isLock,News.addtime,TodayNewsPic,ClassEName,[Domain],SavePath from FS_NS_News as News,FS_NS_NewsClass as Class Where News.ClassID=Class.ClassID And isRecyle=0 and isdraft=0 " & SqlKeyWordStr & " and NewsID<>'"&f_id &"' order by News.addtime desc,News.ID Desc"
				Set rs = Conn.Execute(f_sql)
				if Not rs.Eof then
					c_news = ""
					do while Not rs.Eof
						c_news = c_news & getlist_news(rs,f_Lable.FoosunStyle.StyleContent,leftTitle,200,200,0,"YY04-MM-DD",0,MF_Domain,"c_news","")
						rs.MoveNext
					Loop
					rs.Close:Set rs = Nothing
				else
					c_news = "无相关新闻"
					rs.Close:Set rs = Nothing
				end if
			else
				c_news = ""
			end if
		else
			c_news = ""
		end if
		Set RsSearchObj = Nothing
		f_Lable.DictLableContent.Add "1",c_news
	End Function
	'Flash幻灯片_________________________________________________________________
	Public Function flashfilt(f_Lable,f_type,f_Id)
		dim Randomize_i_Str,TitleNumberStr,FilterSql,ClassId,InSQL_Search,RsFilterObj,str_filt,NewsNumberStr,FlashStr,ContainSubClass,FlashBG
		dim ClassSaveFilePath,ImagesStr,TxtStr,TxtFirst,LinkStr,Temp_Num,ClassIDStr,f_NewsLinkRecordSet,f_NewsLink
		Randomize_i_Str = "F" & GetRand(9)
		TitleNumberStr = f_Lable.LablePara("标题字数")
		NewsNumberStr = f_Lable.LablePara("数量")
		ContainSubClass = f_Lable.LablePara("包含子类")
		FlashBG= f_Lable.LablePara("背景颜色")
		if trim(NewsNumberStr)<>"" and  isNumeric(NewsNumberStr) then
			NewsNumberStr = cint(NewsNumberStr)
		else
			NewsNumberStr = 6
		end if
		ClassId=f_Lable.LablePara("栏目")
		if trim(ClassId) <> "" and ContainSubClass = 1 then
			ClassIDStr = get_SubClass(ClassId)
			InSQL_Search = " and News.ClassId in ('" & ClassIDStr & "')"
		Elseif trim(ClassId)<>"" then
			InSQL_Search = " and News.ClassId='"& ClassId &"'"
		else
			if f_Id <> "" then
				InSQL_Search = " and News.ClassId='"& f_Id &"'"
			else
				InSQL_Search = ""
			end if
		end if
		str_filt=" and "& all_substring &"(NewsProperty,21,1)='1'"
		FilterSql="select top "& NewsNumberStr &" News.ID,NewsId,PopId,SpecialEName,NewsTitle,CurtTitle,NewsNaviContent,isShowReview,TitleColor,titleBorder,TitleItalic,News.IsURL,News.URLAddress"
		FilterSql = FilterSql &",Content,isPicNews,NewsPicFile,NewsSmallPicFile,News.isPop,Source,Editor,Keywords,Author,Hits,SaveNewsPath,News.FileName,News.FileExtName,NewsProperty,isLock,News.addtime,TodayNewsPic,ClassEName,[Domain],SavePath "
		FilterSql = FilterSql &"From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And isPicNews=1 "&InSQL_Search & str_filt &" and  isRecyle=0 and isdraft=0 order by News.addtime desc,News.id desc"
		Set RsFilterObj = Server.CreateObject(G_FS_RS)
		RsFilterObj.CursorLocation = adUseClient
		RsFilterObj.Open FilterSql,Conn,0,1
		If Not RsFilterObj.Eof then
			Temp_Num = RsFilterObj.RecordCount
			If Temp_Num <= 1 then
				Set RsFilterObj = Nothing
				FlashStr = "<!--至少需要两条幻灯新闻才能正确显示幻灯效果-->"
				Set	RsFilterObj = Nothing
				Set f_NewsLinkRecordSet = Nothing
				f_Lable.DictLableContent.Add "1",FlashStr
				Exit Function
			End If
			do while Not RsFilterObj.Eof
				Set f_NewsLinkRecordSet = New CLS_FoosunRecordSet
				Set f_NewsLinkRecordSet.Values(m_NewsLinkFields) = RsFilterObj
				f_NewsLink = get_NewsLink(f_NewsLinkRecordSet)
				Set f_NewsLinkRecordSet = Nothing
				if (Not IsNull(RsFilterObj("NewsSmallPicFile"))) And (RsFilterObj("NewsSmallPicFile") <> "") then
					if ImagesStr = "" then
						ImagesStr =  RsFilterObj("NewsSmallPicFile")
						TxtStr =  GotTopic(RsFilterObj("NewsTitle"),TitleNumberStr)
						TxtFirst = GotTopic(RsFilterObj("NewsTitle"),TitleNumberStr)
						LinkStr =  f_NewsLink
					else
						If Instr(1,LCase(RsFilterObj("NewsSmallPicFile")),"http://") <> 0 then
							ImagesStr = ImagesStr &"|"& RsFilterObj("NewsSmallPicFile")
						Else
							ImagesStr = ImagesStr &"|"&  RsFilterObj("NewsSmallPicFile")
						End If
						TxtStr = TxtStr &"|"&GotTopic(RsFilterObj("NewsTitle"),TitleNumberStr)
						LinkStr = LinkStr & "|" & f_NewsLink
					end if
				end if
				RsFilterObj.MoveNext
			loop
			FlashStr="<script type=""text/javascript"">"& Chr(13)
			FlashStr=FlashStr&" <!--"& Chr(13)
			dim PicWidthStr,PicHeightStr,txtheight
			PicWidthStr = split(f_Lable.LablePara("图片尺寸"),",")(1)
			PicHeightStr = split(f_Lable.LablePara("图片尺寸"),",")(0)
			txtheight = f_Lable.LablePara("文本高度")
			FlashStr=FlashStr&" var focus_width"&Randomize_i_Str&"="&PicWidthStr& Chr(13)
			FlashStr=FlashStr&" var focus_height"&Randomize_i_Str&"="&PicHeightStr& Chr(13)
			FlashStr=FlashStr&" var text_height"&Randomize_i_Str&"="&txtheight& Chr(13)
			FlashStr=FlashStr&" var swf_height"&Randomize_i_Str&" = focus_height"&Randomize_i_Str&"+text_height"&Randomize_i_Str& Chr(13)
			FlashStr=FlashStr&" var pics"&Randomize_i_Str&"='"&ImagesStr&"'"&Chr(13)
			FlashStr=FlashStr&" var links"&Randomize_i_Str&"='"&LinkStr &"'"&Chr(13)
			FlashStr=FlashStr&" var texts"&Randomize_i_Str&"='"&TxtStr&"'"&Chr(13)
			FlashStr=FlashStr&" document.write('<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0"" width=""'+ focus_width"&Randomize_i_Str&" +'"" height=""'+ swf_height"&Randomize_i_Str&" +'"">');"&Chr(13)
			FlashStr=FlashStr&" document.write('<param name=""allowScriptAccess"" value=""sameDomain""><param name=""movie"" value="""&m_Path_Templet & "Flash.swf""><param name=""quality"" value=""high""><param name=""bgcolor"" value="""&FlashBG&""">');"&Chr(13)
			FlashStr=FlashStr&" document.write('<param name=""menu"" value=""false""><param name=wmode value=""opaque"">');"&Chr(13)
			FlashStr=FlashStr&" document.write('<param name=""FlashVars"" value=""pics='+pics"&Randomize_i_Str&"+'&links='+links"&Randomize_i_Str&"+'&texts='+texts"&Randomize_i_Str&"+'&borderwidth='+focus_width"&Randomize_i_Str&"+'&borderheight='+focus_height"&Randomize_i_Str&"+'&textheight='+text_height"&Randomize_i_Str&"+'"">');"&Chr(13)
			FlashStr=FlashStr&" document.write('<embed src="""&m_Path_Templet & "Flash.swf"" wmode=""opaque"" FlashVars=""pics='+pics"&Randomize_i_Str&"+'&links='+links"&Randomize_i_Str&"+'&texts='+texts"&Randomize_i_Str&"+'&borderwidth='+focus_width"&Randomize_i_Str&"+'&borderheight='+focus_height"&Randomize_i_Str&"+'&textheight='+text_height"&Randomize_i_Str&"+'"" menu=""false"" bgcolor=""white"" quality=""high"" width=""'+ focus_width"&Randomize_i_Str&" +'"" height=""'+ swf_height"&Randomize_i_Str&" +'"" allowScriptAccess=""sameDomain"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" />');"&Chr(13)
			FlashStr=FlashStr&" document.write('</object>');"&Chr(13)
			FlashStr=FlashStr&" //-->"& Chr(13)
			FlashStr=FlashStr&"</script>"
		else
			FlashStr=""
		end if
		RsFilterObj.Close
		Set RsFilterObj = Nothing
		f_Lable.DictLableContent.Add "1",FlashStr
	End Function

	'轮换幻灯____________________________________________________________________
	Public Function NorFilter(f_Lable,f_type,f_id)
		dim FilterSql,RsFilterObj,FilterStr,ImagesStr,TxtStr,TxtFirst,ClassSaveFilePath,LinkStr,CssFileStr
		Dim PicWidthStr, PicHeightStr,ChildClassTF,AllClassIDStr,RndFuntion,f_NewsLink,f_NewsLinkRecordSet
		dim fltheightstr,fltwidthstr,Temp_Num,TitleNumberStr,NewsNumberStr,ClassId,InSQL_Search,str_filt,str_Opentype,Str_target
		RndFuntion = now()
		Randomize 
		RndFuntion="F"&right(replace(replace(replace(replace(RndFuntion,"-","")," ",""),":","")&rnd,".",""),3)
		TitleNumberStr= f_Lable.LablePara("标题字数")
		NewsNumberStr= f_Lable.LablePara("数量")
		ChildClassTF = f_Lable.LablePara("包含子类")
		str_Opentype = f_Lable.LablePara("窗口打开方式")
		if trim(NewsNumberStr)<>"" and  isNumeric(NewsNumberStr) then
			NewsNumberStr = cint(NewsNumberStr)
		else
			NewsNumberStr = 6
		end if
		If ChildClassTF = "1" Then
			ChildClassTF = 1
		Else
			ChildClassTF = 0
		End If		
		ClassId=f_Lable.LablePara("栏目")
		if trim(ClassId)<>"" then
			If ChildClassTF = 1 Then
				AllClassIDStr = get_SubClass(ClassId) 
				InSQL_Search = " and News.ClassId in('" & AllClassIDStr & "')"
			Else
				InSQL_Search = " and News.ClassId='"& ClassId &"'"
			End if		
		else
			if f_Id <> "" then
				InSQL_Search = " and News.ClassId='"& f_Id &"'"
			else
				InSQL_Search = ""
			end if
		end if
		str_filt=" and "& all_substring &"(NewsProperty,21,1)='1'"
		FilterSql="select top "& NewsNumberStr &" News.ID,NewsId,PopId,SpecialEName,NewsTitle,CurtTitle,NewsNaviContent,isShowReview,TitleColor,titleBorder,TitleItalic,News.IsURL,News.URLAddress"
		FilterSql = FilterSql &",Content,isPicNews,NewsPicFile,NewsSmallPicFile,News.isPop,Source,Editor,Keywords,Author,Hits,SaveNewsPath,News.FileName,News.FileExtName,NewsProperty,isLock,News.addtime,TodayNewsPic,ClassEName,[Domain],SavePath "
		FilterSql = FilterSql &"From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And isLock=0 and isPicNews=1 "&InSQL_Search & str_filt &" and  isRecyle=0 and isdraft=0 order by News.addtime desc,News.id desc"
		Set RsFilterObj = Server.CreateObject(G_FS_RS)
		RsFilterObj.CursorLocation = adUseClient
		RsFilterObj.Open FilterSql,Conn,0,1
		if not RsFilterObj.Eof then
			Temp_Num = RsFilterObj.RecordCount
			If Temp_Num <=1 then
				Set RsFilterObj = Nothing
				FilterStr = "<!--至少需要两条幻灯新闻才能正确显示幻灯效果-->"
				Set	RsFilterObj = Nothing
				f_Lable.DictLableContent.Add "1",FilterStr
				Exit Function
			End If

			fltheightstr = split(f_Lable.LablePara("图片尺寸"),",")(0)
			fltwidthstr = split(f_Lable.LablePara("图片尺寸"),",")(1)
			CssFileStr = f_Lable.LablePara("CSS样式")
			PicWidthStr = " width=""" & fltwidthstr & """"
			PicHeightStr = " height=""" & fltheightstr & """"
			if CssFileStr <> "" then CssFileStr = " Class='" & CssFileStr & "'"
			do while Not RsFilterObj.Eof
				if (Not IsNull(RsFilterObj("NewsSmallPicFile"))) And (RsFilterObj("NewsSmallPicFile") <> "") then
					Set f_NewsLinkRecordSet = New CLS_FoosunRecordSet
					Set f_NewsLinkRecordSet.Values(m_NewsLinkFields) = RsFilterObj
					f_NewsLink = get_NewsLink(f_NewsLinkRecordSet)
					Set f_NewsLinkRecordSet = Nothing
					if ImagesStr = "" then
						If Instr(1,LCase(RsFilterObj("NewsSmallPicFile")),"http://") <> 0 then
							ImagesStr =  RsFilterObj("NewsSmallPicFile")
						Else
							ImagesStr =  RsFilterObj("NewsSmallPicFile")

						End If
						If str_Opentype="1" Then 
							TxtStr = "<a " & CssFileStr & " href='" & f_NewsLink & "' target='_blank'>" & GotTopic(RsFilterObj("NewsTitle"),TitleNumberStr)&"</a>"
							TxtFirst = "<a " & CssFileStr & " href='" & f_NewsLink & "' target='_blank'>" & GotTopic(RsFilterObj("NewsTitle"),TitleNumberStr)&"</a>"
						Else
							TxtStr = "<a " & CssFileStr & " href='" & f_NewsLink & "'>" & GotTopic(RsFilterObj("NewsTitle"),TitleNumberStr)&"</a>"
							TxtFirst = "<a " & CssFileStr & " href='" & f_NewsLink & "'>" & GotTopic(RsFilterObj("NewsTitle"),TitleNumberStr)&"</a>"
						End If
						LinkStr =  f_NewsLink
					else
						ImagesStr = ImagesStr &","&  RsFilterObj("NewsSmallPicFile")
						If str_Opentype="1" Then 
							TxtStr = TxtStr &",<a " & CssFileStr & " href='" & f_NewsLink & "'  target='_blank'>" & GotTopic(RsFilterObj("NewsTitle"),TitleNumberStr)&"</a>"
						Else
							TxtStr = TxtStr &",<a " & CssFileStr & " href='" & f_NewsLink & "'>" & GotTopic(RsFilterObj("NewsTitle"),TitleNumberStr)&"</a>"								
						End If
						LinkStr = LinkStr & ","& f_NewsLink
					end if
				end if
				RsFilterObj.MoveNext
			loop
			FilterStr="<script language=""vbscript"">"& Chr(13)
			FilterStr = FilterStr & "Dim FileList"&RndFuntion&",FileListArr"&RndFuntion&",TxtList"&RndFuntion&",TxtListArr"&RndFuntion&",LinkList"&RndFuntion&",LinkArr"&RndFuntion& Chr(13)
			FilterStr = FilterStr & "FileList"&RndFuntion&" = """ & ImagesStr& """"& Chr(13)
			
			FilterStr = FilterStr & "LinkList"&RndFuntion& "= """ & LinkStr & """"& Chr(13)
			FilterStr = FilterStr & "TxtList"&RndFuntion&" = """ & TxtStr & """"& Chr(13)
			FilterStr = FilterStr & "FileListArr"&RndFuntion&" = Split(FileList"&RndFuntion&","","")"& Chr(13)
			FilterStr = FilterStr & "LinkArr"&RndFuntion&" = Split(LinkList"&RndFuntion&","","")"& Chr(13)
			FilterStr = FilterStr & "TxtListArr"&RndFuntion&" = Split(TxtList"&RndFuntion&","","")"& Chr(13)
			FilterStr = FilterStr & "Dim CanPlay"&RndFuntion& Chr(13)
			FilterStr = FilterStr & "CanPlay"&RndFuntion&" = CInt(Split(Split(navigator.appVersion,"";"")(1),"" "")(2))>5"& Chr(13)
			FilterStr = FilterStr & "Dim FilterStr"&RndFuntion& Chr(13)
			FilterStr = FilterStr & "FilterStr"&RndFuntion&" = ""RevealTrans(duration=2,transition=23)"""& Chr(13)
			FilterStr = FilterStr & "FilterStr"&RndFuntion&" = FilterStr"&RndFuntion&" + "";BlendTrans(duration=2)"""& Chr(13)
			FilterStr = FilterStr & "If CanPlay"&RndFuntion&" Then"& Chr(13)
			FilterStr = FilterStr & "FilterStr"&RndFuntion&" = FilterStr"&RndFuntion&" + "";progid:DXImageTransform.Microsoft.Fade(duration=2,overlap=0)"""& Chr(13)
			FilterStr = FilterStr & "FilterStr"&RndFuntion&" = FilterStr"&RndFuntion&" + "";progid:DXImageTransform.Microsoft.Wipe(duration=3,gradientsize=0.25,motion=reverse)"""& Chr(13)
			FilterStr = FilterStr & "Else"& Chr(13)
			FilterStr = FilterStr & "Msgbox ""幻灯片播放具有多种动态图片切换效果，但此功能需要您的浏览器为IE5.5或以上版本，否则您将只能看到部分的切换效果。"",64"& Chr(13)
			FilterStr = FilterStr & "End If"& Chr(13)
			FilterStr = FilterStr & "Dim FilterArr"&RndFuntion& Chr(13)
			FilterStr = FilterStr & "FilterArr"&RndFuntion&" = Split(FilterStr"&RndFuntion&","";"")"& Chr(13)
			FilterStr = FilterStr & "Dim PlayImg_M"&RndFuntion& Chr(13)
			FilterStr = FilterStr & "PlayImg_M"&RndFuntion&" = 5 * 1000  "& Chr(13)
			FilterStr = FilterStr & "Dim I"&RndFuntion& Chr(13)
			FilterStr = FilterStr & "I"&RndFuntion&" = 1"& Chr(13)
			FilterStr = FilterStr & "Sub ChangeImg"&RndFuntion& Chr(13)
			FilterStr = FilterStr & "Do While FileListArr"&RndFuntion&"(I"&RndFuntion&")="""""& Chr(13)
			FilterStr = FilterStr & "I"&RndFuntion&" = I"&RndFuntion&" + 1"& Chr(13)
			FilterStr = FilterStr & "If I"&RndFuntion&">UBound(FileListArr"&RndFuntion&") Then I"&RndFuntion&" = 0"& Chr(13)
			FilterStr = FilterStr & "Loop"& Chr(13)
			FilterStr = FilterStr & "Dim J"&RndFuntion& Chr(13)
			FilterStr = FilterStr & "If I"&RndFuntion&">UBound(FileListArr"&RndFuntion&") Then I"&RndFuntion&" = 0"& Chr(13)
			FilterStr = FilterStr & "Randomize"& Chr(13)
			FilterStr = FilterStr & "J"&RndFuntion&" = Int(Rnd * (UBound(FilterArr"&RndFuntion&")+1))"& Chr(13)
			FilterStr = FilterStr & "Img"&RndFuntion&".style.filter = FilterArr"&RndFuntion&"(J"&RndFuntion&")"& Chr(13)
			FilterStr = FilterStr & "Img"&RndFuntion&".filters(0).Apply"& Chr(13)
			FilterStr = FilterStr & "Img"&RndFuntion&".Src = FileListArr"&RndFuntion&"(I"&RndFuntion&")"& Chr(13)
			FilterStr = FilterStr & "Img"&RndFuntion&".filters(0).play"& Chr(13)
			FilterStr = FilterStr & "Link"&RndFuntion&".Href = LinkArr"&RndFuntion&"(I"&RndFuntion&")"& Chr(13)
			If f_Lable.LablePara("显示标题") = "1" Then
				FilterStr = FilterStr & "Txt"&RndFuntion&".filters(0).Apply"& Chr(13)
				FilterStr = FilterStr & "Txt"&RndFuntion&".innerHTML = TxtListArr"&RndFuntion&"(I"&RndFuntion&")"& Chr(13)
				FilterStr = FilterStr & "Txt"&RndFuntion&".filters(0).play"& Chr(13)
			End If
			FilterStr = FilterStr & "I"&RndFuntion&" = I"&RndFuntion&" + 1"& Chr(13)
			FilterStr = FilterStr & "If I"&RndFuntion&">UBound(FileListArr"&RndFuntion&") Then I"&RndFuntion&" = 0"& Chr(13)
			FilterStr = FilterStr & "TempImg"&RndFuntion&".Src = FileListArr"&RndFuntion&"(I"&RndFuntion&")"& Chr(13)
			FilterStr = FilterStr & "TempLink"&RndFuntion&".Href = LinkArr"&RndFuntion&"(I"&RndFuntion&")"& Chr(13)
			FilterStr = FilterStr & "SetTimeout ""ChangeImg"&RndFuntion&""", PlayImg_M"&RndFuntion&",""VBScript"""& Chr(13)
			FilterStr = FilterStr & "End Sub"& Chr(13)
			FilterStr = FilterStr & "</SCRIPT>"& Chr(13)
			FilterStr = FilterStr & "<TABLE WIDTH=""100%"" height=""100%"" BORDER=""0"" CELLSPACING="""" CELLPADDING=""0"">" &vbcrlf
			FilterStr = FilterStr & "<TR ID=""NoScript"&RndFuntion&""">"&vbcrlf
			FilterStr = FilterStr & "<TD Align=""Center"" Style=""Color:White"">对不起，图片浏览功能需脚本支持，但您的浏览器已经设置了禁止脚本运行。请您在浏览器设置中调整有关安全选项。</TD>"&vbcrlf
			FilterStr = FilterStr & "</TR>"&vbcrlf
			If str_Opentype="1" Then 
				Str_target = " target='_blank'"
			Else
				Str_target = ""
			End If
			FilterStr = FilterStr & "<TR Style=""Display:none"" ID=""CanRunScript"&RndFuntion&"""><TD HEIGHT=""100%"" Align=""Center"" vAlign=""Center""><a id=""Link"&RndFuntion&""""&Str_target&"><Img ID=""Img"&RndFuntion&""" "  & PicWidthStr & PicHeightStr & " Border=""0"" ></a>"&vbcrlf
			FilterStr = FilterStr & "</TD></TR><TR Style=""Display:none""><TD><a id=TempLink"&RndFuntion&" ><Img ID=""TempImg"&RndFuntion&""" Border=""0""></a></TD></TR>"&vbcrlf
			If f_Lable.LablePara("显示标题") = "1" Then
				FilterStr = FilterStr & "<TR><TD HEIGHT=""100%"" Align=""Center"" vAlign=""Top"">"&vbcrlf
				FilterStr = FilterStr & "<div ID=""Txt"&RndFuntion&""" style=""PADDING-LEFT: 5px; Z-INDEX: 1; FILTER: progid:DXImageTransform.Microsoft.Fade(duration=1,overlap=0); POSITION:"">"&TxtFirst&"</div>"
				FilterStr = FilterStr & "</TD></TR>"&vbcrlf
			End If
			FilterStr = FilterStr & "</TABLE>"& Chr(13)
			FilterStr = FilterStr & "<Script Language=""VBScript"">"& Chr(13)
			FilterStr = FilterStr & "NoScript"&RndFuntion&".Style.Display = ""none"""& Chr(13)
			FilterStr = FilterStr & "CanRunScript"&RndFuntion&".Style.Display = """""& Chr(13)
			FilterStr = FilterStr & "Img"&RndFuntion&".Src = FileListArr"&RndFuntion&"(0)"& Chr(13)
			FilterStr = FilterStr & "Link"&RndFuntion&".Href = LinkArr"&RndFuntion&"(0)"& Chr(13)
			FilterStr = FilterStr & "SetTimeout ""ChangeImg"&RndFuntion&""", PlayImg_M"&RndFuntion&",""VBScript"""& Chr(13)
			FilterStr = FilterStr & "</Script>"& Chr(13)
		else
			FilterStr="没有幻灯图片"
		End if
		RsFilterObj.Close
		Set RsFilterObj = Nothing
		f_Lable.DictLableContent.Add "1",FilterStr
	End Function
	'得到子类新闻列表
	Public Function subClassList(f_Lable,f_type,f_Id)
		dim rs,f_sql,rs_n,rs_f_sql,OrderStr
		dim div_tf,c_s_i,bg_ground,c_cols,c_cols_1,datenumber,datenumber_tmp,orderby,orderdesc,Inc_SubClass,SQL_Inc_SubClass,N_s_i
		dim titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,loopnumber,picnewstf
		dim classNews_head,classNews_middle1,classNews_bottom,classNews_middle2,c_s_i_1,TableTfStr,Str_TableTf
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass,f_ClassLinkRecordSet,f_ClassLink
		c_cols = f_Lable.LablePara("栏目排列数")
		c_cols_1 = f_Lable.LablePara("新闻排列数")
		bg_ground = f_Lable.LablePara("背景底纹")
		datenumber = f_Lable.LablePara("多少天")
		orderby = f_Lable.LablePara("排列字段")
		orderdesc = f_Lable.LablePara("排列方式")
		loopnumber = f_Lable.LablePara("Loop")
		titlenumber = f_Lable.LablePara("标题数")
		contentnumber = f_Lable.LablePara("内容字数")
		navinumber = f_Lable.LablePara("导航字数")
		picshowtf = f_Lable.LablePara("图文标志")
		datestyle =  f_Lable.LablePara("日期格式")
		openstyle =  f_Lable.LablePara("打开窗口")
		picnewstf =   f_Lable.LablePara("图片新闻")
		Inc_SubClass = f_Lable.LablePara("包含子类")
		if f_Lable.LablePara("输出格式") = "out_DIV" then
			div_tf=1
		else
			div_tf=0
		end if
		If orderby = "" Or Isnull(orderby) Then orderby = "ID"
		If orderdesc = "" Or IsNull(orderdesc) Then
			orderdesc = " Desc"
		Else
			orderdesc = " " & orderdesc
		End If
		If orderby = "ID" Then
			OrderStr = "News." & orderby & orderdesc
		Else
			OrderStr = "News." & orderby & orderdesc & ",News.ID" & orderdesc 
		End If		
		if not isnumeric(titlenumber) then:titlenumber = 30:else:titlenumber = cint(titlenumber):end if
		if not isnumeric(loopnumber) then:loopnumber = 10:else:loopnumber = cint(loopnumber):end if
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")

		f_DivID = f_Lable.LablePara("DivID")
		f_DivClass = f_Lable.LablePara("DivClass")
		f_UlID = f_Lable.LablePara("ulid")
		f_ULClass = f_Lable.LablePara("ulclass")
		f_LiID = f_Lable.LablePara("liid")
		f_LiClass = f_Lable.LablePara("liclass")
		classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		
		if picnewstf="1" then
			picnewstf = " and isPicNews=1"
		else
			picnewstf = ""
		end if
		if len(bg_ground)>5 then
			 bg_ground = " background="""& bg_ground &""""
		else
			 bg_ground = ""
		end If
		
		if datenumber <>"0" then:datenumber_tmp = " and datevalue(addtime)+"&datenumber&">=datevalue(now)":else:datenumber_tmp = "":end if
		if f_Id="" then f_id="0"
		f_sql = "select IsURL,ClassId,ClassName,[Domain],ClassEName,ParentID,FileExtName,FileSaveType,SavePath,UrlAddress From FS_NS_NewsClass Where ParentID='"& f_Id &"' and IsURL=0 and ReycleTF=0 order by OrderID desc,id desc"
		set rs = Conn.execute(f_sql)
		If Not Rs.Eof Then
			subClassList = "<table border=""0"" cellspacing=""3"" cellpadding=""0"" width=""100%"">"&vbNewLine
			Do While Not Rs.Eof 
				subClassList = subClassList & " <tr>" & vbNewLine
				For c_s_i = 1 To c_cols
					If Rs.Eof Then Exit For
					Set f_ClassLinkRecordSet = New CLS_FoosunRecordSet
					Set f_ClassLinkRecordSet.Values(m_ClassLinkFields) = Rs
					f_ClassLink = get_ClassLink(f_ClassLinkRecordSet)
					subClassList = subClassList & "<td width="""&cint(100/c_cols)&"%"" valign=""top"">" & vbNewLine
					subClassList = subClassList & "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbNewLine
					subClassList = subClassList & "<tr>" & vbNewLine
					subClassList = subClassList & "<td width=""80%"" align=""left"" valign=""middle"" height=""26"" class=""dldh"""&bg_ground&">&nbsp;&gt;&gt; <a class=""dldh""  href="""& f_ClassLink & """>"&rs("ClassName")&"</a></td>" & vbNewLine
					subClassList = subClassList & "<td width=""20%"" align=""center"" valign=""middle"" height=""26"""&bg_ground&"><a href="""& f_ClassLink &"""><img src=""http://"& Request.Cookies("FoosunMFCookies")("FoosunMFDomain") &"/sys_images/news_more.gif"" border=""0"" alt="""& rs("ClassName") &"..更多""></a></td>" & vbNewLine
					subClassList = subClassList & "</tr>" & vbNewLine
					subClassList = subClassList & "<tr>" & vbNewLine
					subClassList = subClassList & "<td colspan=""2"" align=""left"" valign=""middle"">" & vbNewLine
						If Inc_SubClass="1" Then
							If getNewsSubClass(rs("ClassId"))<>"" Then
								SQL_Inc_SubClass=" and (News.ClassID='"&rs("ClassId")&"' OR News.ClassID IN ("&getNewsSubClass(rs("ClassId"))&"))"
							Else
								SQL_Inc_SubClass=" AND News.ClassID='"&rs("ClassId")&"'"
							End If
						Else
							SQL_Inc_SubClass=" AND News.ClassID='"&rs("ClassId")&"'"
						End If
						rs_f_sql = "select top "& loopnumber &" News.ID,NewsId,News.PopId,News.ClassID,SpecialEName,NewsTitle,CurtTitle,NewsNaviContent,isShowReview,TitleColor,titleBorder,TitleItalic,News.IsURL,News.URLAddress,Content,isPicNews,NewsPicFile,NewsSmallPicFile,News.isPop,Source,Editor,Keywords,Author,Hits,SaveNewsPath,News.FileName,News.FileExtName,NewsProperty,isLock,News.addtime,TodayNewsPic,ClassEName,[Domain],SavePath From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID "& picnewstf & datenumber_tmp & SQL_Inc_SubClass &" and isRecyle=0 and isdraft=0 order by " & OrderStr & ""
						set rs_n = Conn.execute(rs_f_sql)
						If Not rs_n.eof then
							if div_tf=1 then
								subClassList = subClassList & vbNewLine & classNews_head
								TableTfStr = ""
							Else
								TableTfStr = "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbNewLine
							end if
							subClassList = subClassList & TableTfStr
							If div_tf = 1 Then
								do while not rs_n.eof
									subClassList = subClassList & classNews_middle1 &getlist_news(rs_n,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,"subClassList","")&classNews_middle2
								rs_n.movenext
								loop	
							Else
								do while not rs_n.eof	
									subClassList = subClassList & "<tr>" & vbNewLine
									  For N_s_i = 1 To c_cols_1
										If rs_n.Eof Then Exit For
											 subClassList = subClassList & " <td align=""left"" valign=""middle"">" & getlist_news(rs_n,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,"subClassList","") & "</td>" & vbNewLine
										rs_n.movenext
										Next
									subClassList = subClassList & "</tr>" & vbNewLine	
									loop
								End If
								if div_tf=1 then
									subClassList = subClassList & classNews_bottom
									Str_TableTf = ""
								Else
									Str_TableTf = "</table>" & vbNewLine
								end if
								subClassList = subClassList & Str_TableTf
							end If
							rs_n.Close : Set rs_n = Nothing	
					subClassList = subClassList & "</td></tr></table></td>" & vbNewLine 
				Rs.MoveNext
				Next
			subClassList = subClassList & "</tr>"
			Loop
			subClassList = subClassList & "</table>"
		Else
			subClassList = ""
		End If
		Rs.Close : Set Rs = Nothing
		f_Lable.DictLableContent.Add "1",subClassList
	End Function
	
	'得到单页浏览________________________________________________________________
	Public Function ClassPage(f_Lable,f_type,f_Id)
		Dim ReadSql,RsReadObj
		Dim MF_Domain,datestyle,PageTF,pagecss,pageposi,pagegs
		ClassPage = ""
		if trim(f_Id)="" then
			ClassPage = ""
		else
			ReadSql = "select * from FS_NS_NewsClass where 1=1 and ClassID='"& NoSqlHack(f_Id) &"' and IsURL=2 and  ReycleTF=0"
			Set  RsReadObj = Server.CreateObject(G_FS_RS)
			RsReadObj.CursorLocation =adUseClient
			RsReadObj.open ReadSql,Conn,0,1
			set RsReadObj.ActiveConnection = Nothing		
			MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
			if Not RsReadObj.eof then
				ClassPage = ClassPage & Get_ClassPage(RsReadObj,f_Lable.FoosunStyle.StyleContent,"classpage")
			else
				ClassPage =""
			end If
			RsReadObj.Close
			Set RsReadObj = Nothing
		end if
		
		if ClassPage <> "" then
			Dim f_RegForFlag,f_MatchsForFlag,f_RegForContent,f_FlagArray,f_ContentHeader,f_ContentMiddle,f_ContentTailor
			Dim f_RegForReplace,f_ContentArray,f_MatchsForContent,i
			Set f_RegForFlag = New RegExp
			f_RegForFlag.IgnoreCase = True
			f_RegForFlag.Global = True
			f_RegForFlag.Pattern = "\[fs\:page\]"
			Set f_MatchsForFlag = f_RegForFlag.Execute(ClassPage)
			Set f_RegForContent = New RegExp
			f_RegForContent.IgnoreCase = True
			f_RegForContent.Global = True
			f_RegForContent.Pattern = "\[FS\:CONTENT_START\][^\0]*\[FS\:CONTENT_END\]"
			Set f_MatchsForContent = f_RegForContent.Execute(ClassPage)
			if f_MatchsForFlag.Count >= 1 And f_MatchsForContent.Count = 1 then
				f_FlagArray = Split(ClassPage,f_MatchsForContent(0).Value)
				f_ContentHeader = f_FlagArray(0)
				f_ContentMiddle = f_MatchsForContent(0).Value
				f_ContentTailor = f_FlagArray(1)
				Set f_RegForReplace = New RegExp
				f_RegForReplace.IgnoreCase = True
				f_RegForReplace.Global = True
				f_RegForReplace.Pattern = "\[FS\:CONTENT_START\]"
				f_ContentMiddle = f_RegForReplace.Replace(f_ContentMiddle,"")
				f_RegForReplace.Pattern = "\[FS\:CONTENT_END\]"
				f_ContentMiddle = f_RegForReplace.Replace(f_ContentMiddle,"")
				Set f_RegForReplace = Nothing
				f_ContentArray = Split(f_ContentMiddle,f_MatchsForFlag(0).Value)
				For i = LBound(f_ContentArray) To UBound(f_ContentArray)
					f_Lable.DictLableContent.Add (i + 1) & "",f_ContentHeader & f_ContentArray(i) & f_ContentTailor
				Next
				f_Lable.DictLableContent.Item("0") = G_NEWSPAGESTYLE & ",,"
			else
				f_Lable.DictLableContent.Add "1",ClassPage
			end if
			Set f_MatchsForFlag = Nothing
			Set f_MatchsForContent = Nothing
		else
			f_Lable.DictLableContent.Add "1",""
		end if
	End Function		
	'得到新闻浏览________________________________________________________________
	Public Function ReadNews(f_Lable,f_type,f_Id)
		Dim ReadSql,RsReadObj,i,MF_Domain,datestyle,PageTF,pagecss,pageposi,pagegs
		Dim f_RegForFlag,f_RegForContent,f_RegForReplace,f_ContentHeader,f_ContentTailor,f_ContentMiddle
		Dim f_MatchsForFlag,f_MatchsForContent,f_ContentArray,f_FlagArray
		datestyle = f_Lable.LablePara("日期格式")
		ReadNews = ""
		if trim(f_Id) = "" then
			ReadNews = ""
		else
			If Instr(1,f_Id,"R__D") > 0 Then
				f_Id = CintStr(Replace(f_Id,"R__D",""))
				ReadSql="select News.ID,NewsId,News.PopId,News.ClassID,SpecialEName,NewsTitle,CurtTitle,NewsNaviContent,isShowReview,TitleColor,titleBorder,TitleItalic,News.IsURL,News.URLAddress"
				ReadSql = ReadSql & ",Content,isPicNews,NewsPicFile,NewsSmallPicFile,News.isPop,Source,Editor,Keywords,Author,Hits,SaveNewsPath,News.FileName,News.FileExtName,NewsProperty,isLock,News.addtime,TodayNewsPic,News.IsAdPic,News.AdPicWH,News.AdPicLink,News.AdPicAdress,ClassEName,[Domain],SavePath "
				ReadSql = ReadSql & "From FS_Old_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID and News.ID="& NoSqlHack(f_Id) &" and News.IsURL=0 and isRecyle=0 and isdraft=0"
				Set  RsReadObj = Server.CreateObject(G_FS_RS)
				RsReadObj.CursorLocation = adUseClient
				RsReadObj.open ReadSql,Old_News_Conn,0,1
				If RsReadObj.eof Then
					RsReadObj.Close:Set RsReadObj = Nothing
					Response.Write("<script>alert('没有找到相关的归档新闻！');</script>")
					Response.End()
				End if				
			Else
				ReadSql="select News.ID,NewsId,News.PopId,News.ClassID,SpecialEName,NewsTitle,CurtTitle,NewsNaviContent,isShowReview,TitleColor,titleBorder,TitleItalic,News.IsURL,News.URLAddress"
				ReadSql = ReadSql & ",Content,isPicNews,NewsPicFile,NewsSmallPicFile,News.isPop,Source,Editor,Keywords,Author,Hits,SaveNewsPath,News.FileName,News.FileExtName,NewsProperty,isLock,News.addtime,TodayNewsPic,News.IsAdPic,News.AdPicWH,News.AdPicLink,News.AdPicAdress,ClassEName,[Domain],SavePath "
				ReadSql = ReadSql & "From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID and NewsId='"& NoSqlHack(f_Id) &"' and News.IsURL=0 and  isRecyle=0 and isdraft=0"
				Set  RsReadObj = Server.CreateObject(G_FS_RS)
				RsReadObj.CursorLocation = adUseClient
				RsReadObj.open ReadSql,Conn,0,1	
			End if
			MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
			if not RsReadObj.eof then
				ReadNews = ReadNews & getlist_news(RsReadObj,f_Lable.FoosunStyle.StyleContent,0,0,0,0,datestyle,0,MF_Domain,"readnews","")
			else
				ReadNews =""
			end If
			RsReadObj.close:set RsReadObj = nothing
		end if
		if ReadNews <> "" then
			Set f_RegForFlag = New RegExp
			f_RegForFlag.IgnoreCase = True
			f_RegForFlag.Global = True
			f_RegForFlag.Pattern = "\[fs\:page\]"
			Set f_MatchsForFlag = f_RegForFlag.Execute(ReadNews)
			Set f_RegForContent = New RegExp
			f_RegForContent.IgnoreCase = True
			f_RegForContent.Global = True
			f_RegForContent.Pattern = "\[FS\:CONTENT_START\][^\0]*\[FS\:CONTENT_END\]"
			Set f_MatchsForContent = f_RegForContent.Execute(ReadNews)
			if f_MatchsForFlag.Count >= 1 And f_MatchsForContent.Count = 1 then
				f_FlagArray = Split(ReadNews,f_MatchsForContent(0).Value)
				f_ContentHeader = f_FlagArray(0)
				f_ContentMiddle = f_MatchsForContent(0).Value
				f_ContentTailor = f_FlagArray(1)
				Set f_RegForReplace = New RegExp
				f_RegForReplace.IgnoreCase = True
				f_RegForReplace.Global = True
				f_RegForReplace.Pattern = "\[FS\:CONTENT_START\]"
				f_ContentMiddle = f_RegForReplace.Replace(f_ContentMiddle,"")
				f_RegForReplace.Pattern = "\[FS\:CONTENT_END\]"
				f_ContentMiddle = f_RegForReplace.Replace(f_ContentMiddle,"")
				Set f_RegForReplace = Nothing
				f_ContentArray = Split(f_ContentMiddle,f_MatchsForFlag(0).Value)
				For i = LBound(f_ContentArray) To UBound(f_ContentArray)
					f_Lable.DictLableContent.Add (i + 1) & "",f_ContentHeader & f_ContentArray(i) & "[FS:CONTENT_MOREPAGE_TAG]" & f_ContentTailor
				Next
				f_Lable.DictLableContent.Item("0") = G_NEWSPAGESTYLE & ",,"
			else
				f_Lable.DictLableContent.Add "1",ReadNews
			end if
			Set f_MatchsForFlag = Nothing
			Set f_MatchsForContent = Nothing
		else
			f_Lable.DictLableContent.Add "1",""
		end if
	End Function

	'站点地图____________________________________________________________________
	Public Function SiteMap(f_Lable,f_type,f_Id)
		f_Lable.IsSave = True
		dim classId,cssstyle,SiteMapstr,RsClassObj,i,br_str,f_ClassLinkRecordSet,f_ClassLink
		classId = f_Lable.LablePara("栏目")
		cssstyle = f_Lable.LablePara("标题CSS")
		If classId = "" Then
			classId = 0
		End If
		SiteMapstr = ""
		set RsClassObj = Conn.execute("select ClassId,ClassName,ClassEName,IsURL,ParentID,[Domain],FileExtName,FileSaveType,SavePath,UrlAddress From FS_NS_NewsClass where ReycleTF=0 and ParentID='"&classId&"' order by OrderID desc,id desc")
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
				Set f_ClassLinkRecordSet = New CLS_FoosunRecordSet
				Set f_ClassLinkRecordSet.Values(m_ClassLinkFields) = RsClassObj
				f_ClassLink = get_ClassLink(f_ClassLinkRecordSet)
				Set f_ClassLinkRecordSet = Nothing
				if cssstyle<>"" then
					SiteMapstr = SiteMapstr & br_str & "<img src="""&m_PathDir&"sys_images/+.gif"" border=""0"" /><a href="""&f_ClassLink&""" class="""& cssstyle &""">"&RsClassObj("ClassName")&"</a>"&vbNewLine
				else
					SiteMapstr = SiteMapstr & br_str & "<img src="""&m_PathDir&"sys_images/+.gif"" border=""0"" /><a href="""&f_ClassLink&""">"&RsClassObj("ClassName")&"</a>"&vbNewLine
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
		f_Lable.IsSave = True
		Dim Searchstr,showdate,showClass,rs,select_str,datestr,classstr,MF_Domain,selectShow
		showdate = f_Lable.LablePara("显示日期")
		showClass = f_Lable.LablePara("显示栏目")
		selectShow = f_Lable.LablePara("显示查询类型")
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		Searchstr = ""
		if showdate = "1" then
			datestr = "开始日期:<input type=text name=""s_date"" class=""f-text"" id=""s_date"" style=""width:90;"" value=""" & DateAdd("d",date(),-1) & """ maskType=""shortDate"">"&vbNewLine
			datestr = datestr & "结束日期:<input type=text name=""e_date"" class=""f-text"" id=""e_date"" style=""width:90;"" value=""" & date() & """ maskType=""shortDate"">"&vbNewLine
		else
			datestr = ""
		end if
		if showClass = "1" then
			classstr = "<select name=""ClassId"" class=""f-select"" id=""ClassId"">"
			classstr = classstr & "<option value="""">不指定栏目</option>"&vbNewLine
			set rs = Conn.execute("select ClassId,id,classname,classEname,ParentId From FS_NS_NewsClass where ReycleTF=0 and isURL=0 and ParentId='0' order by OrderId desc,id desc")
			do while not rs.eof
				classstr = classstr & "<option value="""&rs("ClassId")&""">┝"&rs("classname")&"</option>"&vbNewLine
				classstr = classstr & get_optionNewsList(rs("ClassId"),"┝")&vbNewLine
				rs.movenext
			loop
			rs.close:set rs=nothing
			classstr = classstr&"</select>"
		else
			classstr = ""
		end if
		if selectShow="1" then 
			select_str ="<select class=""f-select"" name=""s_type"" id=""s_type"">"&vbNewLine
			select_str=select_str&"<option value=""title"" selected=""selected"">标题</option>"&vbNewLine
			select_str=select_str&"<option value=""stitle"">副标题</option>"&vbNewLine
			select_str=select_str&"<option value=""content"">全文</option>"&vbNewLine
			select_str=select_str&"<option value=""author"">作者</option>"&vbNewLine
			select_str=select_str&"<option value=""keyword"">关键字</option>"&vbNewLine
			select_str=select_str&"<option value=""NaviContent"">导航</option>"&vbNewLine
			select_str=select_str&"<option value=""source"">来源</option>"&vbNewLine
			select_str=select_str&"</select>"&vbNewLine
		end if		
		Searchstr = Searchstr & "<table width=""100%""><form action=""http://"& MF_Domain & "/Search.html"" id=""SearchForm"" name=""SearchForm"" method=""get""><tr><td><input name=""SubSys"" type=""hidden"" id=""SubSys"" value=""NS"" /><input name=""Keyword"" class=""f-text"" type=""text"" id=""Keyword"" size=""15"" /> "&select_str&datestr&classstr&"<input type=""submit"" class=""f-button"" id=""SearchSubmit"" value=""搜索"" /></td></tr></form></table>"
		f_Lable.DictLableContent.Add "1",Searchstr
	 End Function

	 '得到统计信息_______________________________________________________________
	Public Function infoStat(f_Lable,f_type)
		f_Lable.IsSave = True
		dim cols,br_str,infoStatstr
		cols = f_Lable.LablePara("排列方式")
		if cols="0" then
			br_str = "&nbsp;"
		else
			br_str = "<br />"
		end if
		dim rs,c_rs1,c_rs2,c_rs3,c_rs4,c_rs5
		set rs = Conn.execute("select count(id) From FS_NS_News where 1=1 and isRecyle=0 and isdraft=0")
		c_rs1=rs(0)
		rs.close:set rs=nothing
		set rs = Conn.execute("select count(id) From FS_NS_NewsClass where ReycleTF=0")
		c_rs2=rs(0)
		rs.close:set rs=nothing
		set rs = Conn.execute("select count(SpecialID) From FS_NS_Special where 1=1")
		c_rs3=rs(0)
		rs.close:set rs=nothing
		set rs = User_Conn.execute("select count(Userid) From FS_ME_Users")
		c_rs4=rs(0)
		rs.close:set rs=nothing
		if G_IS_SQL_User_DB=0 then
			set rs = User_Conn.execute("select count(Userid) From FS_ME_Users where datevalue(RegTime)=#"&datevalue(date)&"#")
		else
			set rs = User_Conn.execute("select count(Userid) From FS_ME_Users where datediff(d,RegTime,'"&datevalue(date)&"')=0")
		end if
		c_rs5=rs(0)
		rs.close:set rs=nothing
		infoStatstr = "总新闻:"&"<strong>"&c_rs1&"</strong>"&br_str
		infoStatstr = infoStatstr &"总栏目:"&"<strong>"&c_rs2&"</strong>"&br_str
		infoStatstr = infoStatstr &"专题数:"&"<strong>"&c_rs3&"</strong>"&br_str
		infoStatstr = infoStatstr &"会员数:"&"<strong>"&c_rs4&"</strong>"&br_str
		infoStatstr = infoStatstr &"今日注册:"&"<strong>"&c_rs5&"</strong>"&br_str
		f_Lable.DictLableContent.Add "1",infoStatstr
	End Function

	'得到图片头条________________________________________________________________
	Public Function TodayPic(f_Lable,f_type,f_id)
		dim ClassId,InSql_search,rs,MF_Domain,TodayPicstr,f_NewsLinkRecordSet,f_NewsLink
		ClassID = f_Lable.LablePara("栏目")
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		if trim(ClassId)="" then
			if trim(f_id)<>"" then
				InSql_search = " And Today.ClassId = '"&f_id&"'"
			else
				InSql_search = ""
			end if
		else
			if trim(ClassId)<>"" then
				InSql_search = " And Today.ClassId = '"&ClassId&"'"
			else
				InSql_search = ""
			end if
		end if
		Dim ToDay_PicNewsSQL,ToDay_PicNewsObj
		ToDay_PicNewsSQL = "Select Top 1 News.FileName,News.FileExtName,Today.TodayPic_SavePath As ToDayPicSavePath,News.IsURL,News.URLAddress,SaveNewsPath,ClassEName,[Domain],SavePath From FS_NS_News as News,FS_NS_TodayPic as Today,FS_NS_NewsClass as Class Where News.isRecyle=0 And News.isLock=0 And News.isdraft=0 And News.TodayNewsPic=1 And News.NewsID=Today.NewsID And News.ClassID=Class.ClassID" & InSql_search & " Order By Today.ID Desc,Today.AddTime Desc"
		Set ToDay_PicNewsObj = Server.CreateObject(G_FS_RS)
		ToDay_PicNewsObj.Open ToDay_PicNewsSQL,Conn,0,1
		if ToDay_PicNewsObj.eof then
			TodayPicstr = ""
			ToDay_PicNewsObj.close
			set ToDay_PicNewsObj=nothing
		else
			Set f_NewsLinkRecordSet = New CLS_FoosunRecordSet
			Set f_NewsLinkRecordSet.Values(m_NewsLinkFields) = ToDay_PicNewsObj
			f_NewsLink = get_NewsLink(f_NewsLinkRecordSet)
			Set f_NewsLinkRecordSet = Nothing
			TodayPicstr = "<a href="""& f_NewsLink &"""><img src=""http://"&MF_Domain&"/"&G_UP_FILES_DIR&"/TodayPicFiles/"&ToDay_PicNewsObj("ToDayPicSavePath")&".jpg"" border=""0""></a>"
			ToDay_PicNewsObj.close
			set ToDay_PicNewsObj=nothing
		end if
		f_Lable.DictLableContent.Add "1",TodayPicstr
	End Function

	'得到文字头条________________________________________________________________
	Public Function TodayWord(f_Lable,f_type,f_Id)
		dim Content_List,div_tf,ClassId,classNews_head,classNews_bottom,classNews_middle1,classNews_middle2,f_sql,f_rs_obj
		dim RefreshNumber,search_inSQL,colnumber,cssstyle,Titlenumber,ShowReview,ShowReviewTF,AlignType
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass,f_NewsLink,f_NewsLinkRecordSet
		if f_Lable.LablePara("") = "out_DIV" then
			div_tf = 1
		else
			div_tf=0
		end if
		ClassId = f_Lable.LablePara("栏目")
		RefreshNumber = f_Lable.LablePara("调用数量")
		colnumber = f_Lable.LablePara("列数")
		Titlenumber = f_Lable.LablePara("标题字数")
		ShowReview = f_Lable.LablePara("显示评论")
		AlignType = f_Lable.LablePara("对齐方式")
		if trim(f_Lable.LablePara("标题CSS"))<>"" then
			cssstyle = " class="""&f_Lable.LablePara("标题CSS")&""""
		else
			cssstyle = ""
		end if
		if trim(ClassId)<>"" then
			search_inSQL = " and "& all_substring &"(NewsProperty,11,1)='1' and News.ClassId='"& ClassId &"'"
		else
			if trim(f_id)<>"" then
				search_inSQL = " and "& all_substring &"(NewsProperty,11,1)='1' and News.ClassId='"& f_id &"'"
			else
				search_inSQL = " and "& all_substring &"(NewsProperty,11,1)='1'"
			end if
		end if
		
		f_DivID = f_Lable.LablePara("DivID")
		f_DivClass = f_Lable.LablePara("Divclass")
		f_UlID = f_Lable.LablePara("ulid")
		f_ULClass = f_Lable.LablePara("ulclass")
		f_LiID = f_Lable.LablePara("liid")
		f_LiClass = f_Lable.LablePara("liclass")
		
		if div_tf=1 then
			classNews_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classNews_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		else
			classNews_head = "<table border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"">"
		end if
		classNews_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classNews_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		f_sql="select top "& RefreshNumber &" News.ID,NewsId,PopId,SpecialEName,NewsTitle,CurtTitle,NewsNaviContent,isShowReview,TitleColor,titleBorder,TitleItalic,News.IsURL,News.URLAddress"
		f_sql = f_sql &",ClassEName,[Domain],SavePath,Content,isPicNews,NewsPicFile,NewsSmallPicFile,News.isPop,Source,Editor,Keywords,Author,Hits,SaveNewsPath,News.FileName,News.FileExtName,NewsProperty,isLock,News.addtime,TodayNewsPic "
		f_sql = f_sql &"From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID "&search_inSQL&" and  isRecyle=0 And  isLock=0 and isdraft=0 order by News.addtime desc,News.id desc"
		Set  f_rs_obj = Server.CreateObject(G_FS_RS)
		f_rs_obj.open f_sql,Conn,0,1
		if f_rs_obj.eof then
			Content_List = ""
			f_rs_obj.close:set f_rs_obj=nothing
		else
			dim c_i_k
			Content_List = ""
			if div_tf = 0 then
				c_i_k = 0
				if cint(colnumber)<>1 then Content_List = Content_List &  "  <tr>"
			end if
			do while not f_rs_obj.eof
				if ShowReview="1" then
					ShowReviewTF = "<img src=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/sys_images/sp_er.gif"" border=""0"" /><a href=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/ShowReviewList.asp?Id="&f_rs_obj("ID")&"&Type=NS"" target=""_blank""><font color=""red"">评</font></a>"
				else
					ShowReviewTF = ""
				end if
				Set f_NewsLinkRecordSet = New CLS_FoosunRecordSet
				Set f_NewsLinkRecordSet.Values(m_NewsLinkFields) = f_rs_obj
				f_NewsLink = get_NewsLink(f_NewsLinkRecordSet)
				Set f_NewsLinkRecordSet = Nothing
				if div_tf=1 then
					Content_List= Content_List & classNews_middle1 & "<a href="""&f_NewsLink&""""& cssstyle&">"&GotTopic(f_rs_obj("newstitle"),Titlenumber)&"</a>"& ShowReviewTF & classNews_middle2
				else
					if cint(colnumber) =1 then
						Content_List= Content_List & vbNewLine&"   <tr><td align="&AlignType&"><a href="""&f_NewsLink&""""& cssstyle&">"&GotTopic(f_rs_obj("newstitle"),Titlenumber)&"</a>"& ShowReviewTF &"</td></tr>"
					else
						Content_List= Content_List & "<td width="""& cint(100/cint(colnumber))&"%"" align="&AlignType&"><a href="""&f_NewsLink&""""& cssstyle&">"&GotTopic(f_rs_obj("newstitle"),Titlenumber)&"</a>"& ShowReviewTF &"</td>"
					end if
				end if
				f_rs_obj.movenext
				if div_tf = 0 then
					if cint(colnumber)<>1 then
					c_i_k = c_i_k+1
						if c_i_k mod cint(colnumber) = 0 then
							Content_List = Content_List & "</tr>"&vbNewLine&"  <tr>"
						end if
					end if
				end if
			loop
		   if div_tf=0 then
				if cint(colnumber)<>1 then Content_List = Content_List & "</tr>"&vbNewLine
		   end if
			f_rs_obj.close
			set f_rs_obj=nothing
			Content_List=classNews_head & Content_List & classNews_bottom
		end if
		f_Lable.DictLableContent.Add "1",Content_List
	End Function

	'栏目导航____________________________________________________________________
	Public Function ClassNavi(f_Lable,f_type,f_Id)
		dim ClassId,cols,titlecss,rs,ClassNavistr,ParentIDstr,cols_str,div_tf,f_ClassLinkRecordSet,f_ClassLink
		dim classNews_head,classNews_bottom,classNews_middle1,classNews_middle2,cssstyle,titleNavi
		ClassId = f_Lable.LablePara("栏目")
		cols = f_Lable.LablePara("方向")
		titlecss = f_Lable.LablePara("标题CSS")
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

			titleNavi = f_Lable.LablePara("标题导航")
			If trim(titleNavi)="" Or isnull(titleNavi) Then
				titleNavi = ""
			Else
				If Len(titleNavi)>1 And instr(1,titleNavi,".")>0 Then
					titleNavi = "<img src="&titleNavi&" />"
				Else
					titleNavi = titleNavi
				End if
			End if
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
		else
			classNews_head = "<table border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"">"
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
		set rs = Conn.execute("select ClassID,OrderID,ClassName,isShow,ParentID,ReycleTF,IsURL,ClassEName,[Domain],FileExtName,FileSaveType,SavePath,UrlAddress From FS_NS_NewsClass where isShow=1 and ReycleTF=0 "& ParentIDstr &" Order by OrderId desc,id asc")
		ClassNavistr = ""
		if rs.eof then
			rs.close:set rs=nothing
		else
			do while not rs.eof
				Set f_ClassLinkRecordSet = New CLS_FoosunRecordSet
				Set f_ClassLinkRecordSet.Values(m_ClassLinkFields) = rs
				f_ClassLink = get_ClassLink(f_ClassLinkRecordSet)
				Set f_ClassLinkRecordSet = Nothing
				if div_tf=1 then
					ClassNavistr = ClassNavistr &classNews_middle1 & titleNavi & "<a href="""&f_ClassLink&""">"&rs("ClassName")&"</a>" & classNews_middle2
				else
					if cols="0" then
						cols_str = "&nbsp;"
					else
						cols_str = "<br />"
					end If
					If ClassNavistr="" Then
						ClassNavistr = titleNavi & "<a href="""&f_ClassLink&""""& cssstyle&">"&rs("ClassName")&"</a>"
					Else
						ClassNavistr = ClassNavistr& cols_str & titleNavi & "<a href="""&f_ClassLink&""""& cssstyle&">"&rs("ClassName")&"</a>"
					End If
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

	'专题导航____________________________________________________________________
	Public Function SpecialNavi(f_Lable,f_type,f_Id)
		dim cols,titlecss,div_tf,cssstyle,titleNavi,rs,SpecialNavistr,classNews_head,classNews_bottom,classNews_middle1,classNews_middle2,cols_str
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass,f_SpecialLinkRecordSet,f_SpecialLink
		cols = f_Lable.LablePara("方向")
		titlecss = f_Lable.LablePara("CSS")
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
			titleNavi = f_Lable.LablePara("导航")
		end if
		
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

		dim ParentIDstr
		if isnumeric(f_id) then
			if f_id<>"" then
				ParentIDstr = " and SpecialID="&f_id&""
			end if
		end if
		dim sql
		sql="select SpecialID,SpecialCName,SpecialEName,isLock,ExtName,SavePath,FileSaveType From FS_NS_Special where 1=1 "&ParentIDstr&" Order by SpecialID desc"
		set rs = Conn.execute(sql)
		SpecialNavistr = ""
		if rs.eof then
			SpecialNavistr = ""
			rs.close:set rs=nothing
		else
			do while not rs.eof
				Set f_SpecialLinkRecordSet = New CLS_FoosunRecordSet
				Set f_SpecialLinkRecordSet.Values(m_SpecialLinkFields) = rs
				f_SpecialLink = get_specialLink(f_SpecialLinkRecordSet)
				Set f_SpecialLinkRecordSet = Nothing
				if div_tf=1 then
					SpecialNavistr = SpecialNavistr &classNews_middle1 & "<a href="""&f_SpecialLink&""">"&rs("SpecialCName")&"</a>" & classNews_middle2
				else
					if cols="0" then
						cols_str = "&nbsp;"
					else
						cols_str = "<br />"
					end if
					SpecialNavistr = SpecialNavistr & titleNavi & "<a href="""&f_SpecialLink&""""& cssstyle&">"&rs("SpecialCName")&"</a>"& cols_str &""
				end if
				rs.movenext
			loop
			rs.close:set rs=nothing
		end if
		if div_tf=1 then
			SpecialNavistr=classNews_head & SpecialNavistr & classNews_bottom
		end if
		f_Lable.DictLableContent.Add "1",SpecialNavistr&" "
	End Function

	'RSS聚合_____________________________________________________________________
	Public Function RssFeed(f_Lable,f_type,f_Id)
		dim ClassID,RssFeedstr
		ClassID = f_Lable.LablePara("栏目")
		if trim(ClassID)<>"" then
			RssFeedstr = "<a href=""http://"& Request.Cookies("FoosunMFCookies")("FoosunMFDomain") &"/xml/NS/"&ClassID&".xml""><img src=""http://"& Request.Cookies("FoosunMFCookies")("FoosunMFDomain") &"/sys_images/rss.gif"" border=""0""></a>"
		elseif trim(f_id)<>"" then
			RssFeedstr = "<a href=""http://"& Request.Cookies("FoosunMFCookies")("FoosunMFDomain") &"/xml/NS/"&f_id&".xml""><img src=""http://"& Request.Cookies("FoosunMFCookies")("FoosunMFDomain") &"/sys_images/rss.gif"" border=""0""></a>"
		else
			RssFeedstr = ""
		end if
		f_Lable.DictLableContent.Add "1",RssFeedstr
	End Function

	'专题调用____________________________________________________________________
	Public Function SpecialCode(f_Lable,f_type,f_Id)
		dim ClassID,titleNavi,SpecialCodestr,cols,pictf,picsize,piccss,piccssstr,ContentTF,ContentNumber,div_tf,titlecss,cssstyle,ContentCSS,ContentCSSstr,classNews_head,classNews_bottom,classNews_middle1,classNews_middle2
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass,f_SpecialLinkRecordSet,f_SpecialLink
		ClassID = f_Lable.LablePara("专题")
		titleNavi = f_Lable.LablePara("导航")
		titlecss = f_Lable.LablePara("专题名称CSS")
		ContentCSS = f_Lable.LablePara("导航内容CSS")
		cols=f_Lable.LablePara("排列方式")
		pictf=f_Lable.LablePara("图片显示")
		picsize=f_Lable.LablePara("图片尺寸")
		piccss=f_Lable.LablePara("图片css")
		ContentTF = f_Lable.LablePara("导航内容")
		ContentNumber= f_Lable.LablePara("导航内容字数")
		if piccss<>"" then
			piccssstr = " class="""& piccss &""""
		else
			piccssstr = ""
		end if
		if trim(ClassID)="" then
			SpecialCodestr = "错误的标签,by Foosun.cn"
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
			titleNavi = f_Lable.LablePara("导航")
		end if
		
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
		
		SpecialCodestr = ""
		dim rs
		set rs=Conn.execute("select SpecialCName,SpecialEName,SpecialContent,naviPic,ExtName,SavePath,FileSaveType From FS_NS_Special Where SpecialEName='"& ClassID &"'")
		if rs.eof then
			SpecialCodestr = ""
			rs.close:set rs=nothing
		else
			Set f_SpecialLinkRecordSet = New CLS_FoosunRecordSet
			Set f_SpecialLinkRecordSet.Values(m_SpecialLinkFields) = rs
			f_SpecialLink = get_specialLink(f_SpecialLinkRecordSet)
			Set f_SpecialLinkRecordSet = Nothing
			if div_tf=1 then
				if pictf="1" then
					if trim(picsize)="" then
						SpecialCodestr = SpecialCodestr & ""
					else
						SpecialCodestr = SpecialCodestr & "  <a href="""&f_SpecialLink&"""><img "&piccssstr&" src="""&rs("naviPic")&""" width="""&split(picsize,",")(1)&""" height="""&split(picsize,",")(0)&""" border=""0"" /></a>" &vbNewLine
					end if
				end if
				SpecialCodestr = SpecialCodestr & classNews_middle1 & "<a href="""&f_SpecialLink&""">"& rs("SpecialCName") &"</a>" & classNews_middle2
				if ContentTF="1" then
					SpecialCodestr = SpecialCodestr & classNews_middle1 &"&nbsp;"& GetCStrLen(""&rs("SpecialContent"),ContentNumber) & "&nbsp;<a href="""& f_SpecialLink &""">详细</a>" & classNews_middle2
				end if
				SpecialCodestr = classNews_head & SpecialCodestr & classNews_bottom
			else
				SpecialCodestr = SpecialCodestr& "<table width=""99%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"&vbNewLine&" <tr>"
				if pictf="1" then
					if trim(picsize)="" then
						SpecialCodestr = SpecialCodestr & ""
					else
						if cols="0" then
							SpecialCodestr = SpecialCodestr & "<td align=""center""><a href="""&f_SpecialLink&"""><img "&piccssstr&" src="""&rs("naviPic")&""" width="""&split(picsize,",")(1)&""" height="""&split(picsize,",")(0)&""" border=""0"" /></a></td>" &vbNewLine
						else
							SpecialCodestr = SpecialCodestr & "<td><a href="""&f_SpecialLink&"""><img "&piccssstr&" src="""&rs("naviPic")&""" width="""&split(picsize,",")(1)&""" height="""&split(picsize,",")(0)&""" border=""0"" /></a></td></tr>" &vbNewLine
						end if
					end if
				end if
				if cols="0" then
					SpecialCodestr = SpecialCodestr & "<td>"& titleNavi &"<a href="""&f_SpecialLink&""""&cssstyle&">"& rs("SpecialCName") &"</a>"
					if ContentTF="1" then
						SpecialCodestr = SpecialCodestr & "<br /><div align=""left"" "&ContentCSSstr&">&nbsp;"& GetCStrLen(""&rs("SpecialContent"),ContentNumber)&"&nbsp;<a href="""& f_SpecialLink &""">详细</a></div></td>"&vbNewLine
					else
						SpecialCodestr = SpecialCodestr & "</td></tr>"&vbNewLine
					end if
				else
					SpecialCodestr = SpecialCodestr & " <tr><td>"& titleNavi &"<a href="""&f_SpecialLink&""""&cssstyle&">"& rs("SpecialCName") &"</a>"&vbNewLine
					if ContentTF="1" then
						SpecialCodestr = SpecialCodestr & "<br /><div align=""left"" "&ContentCSSstr&">&nbsp;"& GetCStrLen(""&rs("SpecialContent"),ContentNumber)&"&nbsp;<a href="""& f_SpecialLink &""">详细</a></div></td></tr>"&vbNewLine
					else
						SpecialCodestr = SpecialCodestr & "</td></tr>"&vbNewLine
					end if
				end if
				SpecialCodestr = SpecialCodestr & "</table>"
			end if
			rs.close:set rs=nothing
		end if
		f_Lable.DictLableContent.Add "1",SpecialCodestr
	End Function

	'栏目调用____________________________________________________________________
	Public Function ClassCode(f_Lable,f_type,f_Id)
		dim ClassID,titleNavi,ClassCodestr,cols,pictf,picsize,piccss,piccssstr,ContentTF,ContentNumber,div_tf,titlecss,cssstyle,ContentCSS,ContentCSSstr,classNews_head,classNews_bottom,classNews_middle1,classNews_middle2
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass,ClassStr,f_ClassLinkRecordSet,f_ClassLink
		ClassID = f_Lable.LablePara("栏目")
		titleNavi = f_Lable.LablePara("导航")
		titlecss = f_Lable.LablePara("栏目名称CSS")
		ContentCSS = f_Lable.LablePara("导航内容CSS")
		cols=f_Lable.LablePara("排列方式")
		pictf=f_Lable.LablePara("图片显示")
		picsize=f_Lable.LablePara("图片尺寸")
		piccss=f_Lable.LablePara("图片CSS")
		ContentTF = f_Lable.LablePara("导航内容")
		ContentNumber= f_Lable.LablePara("导航内容字数")
		if piccss<>"" then
			piccssstr = " class="""& piccss &""""
		else
			piccssstr = ""
		end if
		if trim(ClassID)="" then
			ClassStr = ""
		Else
			ClassStr = " And ClassID = '" & ClassID & "'"
		End If		
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
			titleNavi = f_Lable.LablePara("导航")
		end if
		
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

		ClassCodestr = ""
		dim rs
		set rs=Conn.execute("select ClassID,ClassName,ClassNaviContent,ClassNaviPic,IsURL,ClassEName,[Domain],FileExtName,FileSaveType,SavePath,UrlAddress From FS_NS_NewsClass Where ReycleTF = 0" & ClassStr & " Order By ID Desc")
		if rs.eof then
			ClassCodestr = ""
			rs.close:set rs=nothing
		else
			Set f_ClassLinkRecordSet = New CLS_FoosunRecordSet
			Set f_ClassLinkRecordSet.Values(m_ClassLinkFields) = rs
			f_ClassLink = get_ClassLink(f_ClassLinkRecordSet)
			Set f_ClassLinkRecordSet = Nothing
			if div_tf=1 then
				if pictf="1" then
					if trim(picsize)="" then
						ClassCodestr = ClassCodestr & ""
					else
						ClassCodestr = ClassCodestr & "  <a href="""&f_ClassLink&"""><img "&piccssstr&" src="""&rs("ClassNaviPic")&""" width="""&split(picsize,",")(1)&""" height="""&split(picsize,",")(0)&""" border=""0"" /></a>" &vbNewLine
					end if
				end if
				ClassCodestr = ClassCodestr & classNews_middle1 & "<a href="""&f_ClassLink&""">"& rs("ClassName") &"</a>" & classNews_middle2
				if ContentTF="1" then
					ClassCodestr = ClassCodestr & classNews_middle1 &"&nbsp;"& GetCStrLen(""&rs("ClassNaviContent"),ContentNumber) & "&nbsp;<a href="""& f_ClassLink &""">详细</a>" & classNews_middle2
				end if
				ClassCodestr = classNews_head & ClassCodestr & classNews_bottom
			else
				ClassCodestr = ClassCodestr& "<table width=""99%"" border=""0"" cellspacing=""0"" cellpadding=""5"">"&vbNewLine&" <tr>"
				if pictf="1" then
					if trim(picsize)="" then
						ClassCodestr = ClassCodestr & ""
					else
						if cols="0" then
							ClassCodestr = ClassCodestr & "<td align=""center""><a href="""&f_ClassLink&"""><img "&piccssstr&" src="""&rs("ClassNaviPic")&""" width="""&split(picsize,",")(1)&""" height="""&split(picsize,",")(0)&""" border=""0"" /></a></td>" &vbNewLine
						else
							ClassCodestr = ClassCodestr & "<td><a href="""&f_ClassLink&"""><img "&piccssstr&" src="""&rs("ClassNaviPic")&""" width="""&split(picsize,",")(1)&""" height="""&split(picsize,",")(0)&""" border=""0"" /></a></td></tr>" &vbNewLine
						end if
					end if
				end if
				if cols="0" then
					ClassCodestr = ClassCodestr & "<td>"& titleNavi &"<a href="""&f_ClassLink&""""&cssstyle&">"& rs("ClassName") &"</a>"
					if ContentTF="1" then
						ClassCodestr = ClassCodestr & "<br /><div align=""left"" "&ContentCSSstr&">&nbsp;"& GetCStrLen(""&rs("ClassNaviContent"),ContentNumber)&"&nbsp;<a href="""& f_ClassLink &""">详细</a></div></td>"&vbNewLine
					else
						ClassCodestr = ClassCodestr & "</td></tr>"&vbNewLine
					end if
				else
					ClassCodestr = ClassCodestr & " <tr><td>"& titleNavi &"<a href="""&f_ClassLink&""""&cssstyle&">"& rs("ClassName") &"</a>"&vbNewLine
					if ContentTF="1" then
						ClassCodestr = ClassCodestr & "<br /><div align=""left"" "&ContentCSSstr&">&nbsp;"& GetCStrLen(""&rs("ClassNaviContent"),ContentNumber)&"&nbsp;<a href="""& f_ClassLink &""">详细</a></div></td></tr>"&vbNewLine
					else
						ClassCodestr = ClassCodestr & "</td></tr>"&vbNewLine
					end if
				end if
				ClassCodestr = ClassCodestr & "</table>"
			end if
			rs.close:set rs=nothing
		end if
		f_Lable.DictLableContent.Add "1",ClassCodestr
	End Function

	'不规则新闻开始______________________________________________________________
	Public Function DefineNews(f_Lable,f_type,f_Id)
		f_Lable.IsSave = True
		dim DefineNewsstr,ClassId,TitleCss,TitleCssstr,classNews_head,classNews_bottom,classNews_middle1,classNews_middle2
		dim div_tf,dot_str,TmpRsObj,RowValue,i,TitleNavi,f_DefineNewsSQL,f_NewsLinkRecordSet,f_NewsLink
		dim tmp_alt_title,tmp_titleBorder,tmp_TitleItalic,tmp_TitleColor
		Dim f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass
		ClassId=f_Lable.LablePara("不规则ID")
		TitleCss=f_Lable.LablePara("标题CSS")
		TitleNavi = f_Lable.LablePara("导航")
		
		if f_Lable.LablePara("输出格式") = "out_DIV" then
			div_tf = 1
			TitleCssstr = ""
		else
			div_tf=0
			if trim(TitleCss)<>"" then
				TitleCssstr = " class="""& TitleCss &""""
			else
				TitleCssstr = ""
			end if
		end if

		DefineNewsstr = ""
		f_DefineNewsSQL = "Select UnregNewsName,[Rows],titleBorder,TitleItalic,TitleColor,"
		f_DefineNewsSQL = f_DefineNewsSQL & "ClassEName,[Domain],SavePath,News.IsURL,News.URLAddress,SaveNewsPath,News.FileName,News.FileExtName"
		f_DefineNewsSQL = f_DefineNewsSQL & " from FS_NS_News_Unrgl as Unrgl,FS_NS_News as News,FS_NS_NewsClass as Class"
		f_DefineNewsSQL = f_DefineNewsSQL & " Where Unrgl.MainUnregNewsID=News.NewsID And News.ClassID=Class.ClassID And UnregulatedMain='" & ClassId & "' order by [Rows],Unrgl.id asc"
		i=1
		Set TmpRsObj = Server.CreateObject(G_FS_RS)
		TmpRsObj.Open f_DefineNewsSQL,Conn,0,1
		Do While Not TmpRsObj.eof
			Set f_NewsLinkRecordSet = New CLS_FoosunRecordSet
			Set f_NewsLinkRecordSet.Values(m_NewsLinkFields) = TmpRsObj
			f_NewsLink = Get_NewsLink(f_NewsLinkRecordSet)
			Set f_NewsLinkRecordSet = Nothing
			tmp_alt_title = TmpRsObj("UnregNewsName")
			RowValue = Cint(TmpRsObj("Rows"))
			if i = RowValue then
				DefineNewsstr = DefineNewsstr & " " & vbNewLine
			else
				DefineNewsstr = DefineNewsstr & "<br />" & vbNewLine
				i=i+1
			end if
			DefineNewsstr = DefineNewsstr &  TitleNavi & "<a href="""&f_NewsLink&""""&TitleCssstr&">"&tmp_alt_title&"</a>"
			TmpRsObj.movenext
		Loop
		
		TmpRsObj.close
		set TmpRsObj = nothing
		f_Lable.DictLableContent.Add "1",DefineNewsstr
	End Function

	'归档标签
	Public Function OldNews(f_Lable,f_type,f_Id)
		f_Lable.IsSave = True
		OldNews = "<form id=""recordNewsForm"" name=""recordNewsForm"" method=""get"" action=""http://"& request.Cookies("FoosunMFCookies")("FoosunMFDomain") &"/Search.html""> 关键字:<input name=""keyword"" class=""f-text"" size=""15"" type=""text"" value="""" /> <input type=""submit"" class=""f-button"" value=""搜索"" /><input type=""hidden"" name=""SubSys"" value=""RD"" /> <input type=""reset"" class=""f-reset"" name=""Submit2"" value=""重填"" /></form>"
		f_Lable.DictLableContent.Add "1",OldNews
	End Function

	'替换样式列表________________________________________________________________
	Public Function getlist_news(f_obj,s_Content,f_titlenumber,f_contentnumber,f_navinumber,f_picshowtf,f_datestyle,f_openstyle,f_MF_Domain,f_subsys_ListType,templistnum)
		Dim f_target,get_SpecialEName,ListSql,Rs_ListObj,s_NewsPathUrl,Rs_Authorobj,k_i,k_tmp_Char,k_tmp_uchar,k_tmp_Chararray,FormReview
		Dim s_m_Rs,s_array,s_t_i,tmp_list,s_f_classSql,m_Rs_class,class_path,str_newstitle,Temp_Bool,f_NewsLinkRecordSet,f_ClassLinkRecordSet
		dim GetNewFlag,IsOpen,isCSS,newContent,TMPG,isDateNumber,NewSTR,NewsTime,Nowtime,f_SpecialLinkRecordSet,f_SpecialLink
		select case f_subsys_ListType
			case "specialnews"
				get_SpecialEName = f_obj("SpecialEName")
			case else
				get_SpecialEName = f_obj("SpecialEName")
		end select
		if f_openstyle & "" = "0" then
			f_target=" "
		else
			f_target=" target=""_blank"""
		end if
		if instr(s_Content,"{NS:FS_ID}")>0 then
			s_Content = replace(s_Content,"{NS:FS_ID}",f_obj("Id"))
		end if
		if instr(s_Content,"{NS:FS_NewsID}")>0 then
			s_Content = replace(s_Content,"{NS:FS_NewsID}",f_obj("NewsId"))
		end if
		if instr(s_Content,"{NS:FS_NewsTitle}")>0 then
			if f_subsys_ListType="readnews" then
				str_newstitle = f_obj("NewsTitle")
			else
				str_newstitle = templistnum&GotTopic(f_obj("NewsTitle"),f_titlenumber)
				if trim(f_obj("titleBorder"))=1 then
					str_newstitle = "<strong>"&str_newstitle&"</strong>"
				end if
				if trim(f_obj("TitleItalic"))=1 then
					str_newstitle = "<em>"&str_newstitle&"</em>"
				end if
				if trim(f_obj("TitleColor"))<>"" then
					str_newstitle = "<font color=""#"& f_obj("TitleColor")&""">"&str_newstitle&"</font>"
				end if
			end if
			if f_picshowtf & "" = "1" then
				if f_obj("isPicNews")=1 then
					str_newstitle = str_newstitle & "[图]"
				end if
			end if
			if f_subsys_ListType<>"readnews" then
				if G_newNews="" then
					TMPG="1|0|newNews|2"
					GetNewFlag = split(TMPG,"|")
				else
					GetNewFlag = split(G_newNews,"|")
				end if
				IsOpen=GetNewFlag(0)
				isCSS=GetNewFlag(1)
				newContent=GetNewFlag(2)
				isDateNumber=GetNewFlag(3)
				NewsTime =f_obj("addtime")
				Nowtime=now()						
				if datediff("d",NewsTime,Nowtime)<=clng(isDateNumber) then
					if isOpen = "1" then
						if isCSS="1" then
							NewSTR =" <span class="""&newContent&"""></span>"
						else
							NewSTR =" <img src="""&newContent&""" border=""0"" align=""absmiddle"" />"
						end if
					end if
				end if
			end if
			s_Content = replace(s_Content,"{NS:FS_NewsTitle}",str_newstitle & NewSTR)
		end if
		if instr(s_Content,"{NS:FS_NewsTitleAll}")>0 then
			if f_subsys_ListType="readnews" then
				str_newstitle = f_obj("NewsTitle")
			else
				str_newstitle = templistnum&Replace(Replace(Replace(Replace(Lose_Html(f_obj("NewsTitle"))," ","&nbsp;"),Chr(34),"&quot;"),">","&gt;"),"<","&lt;")
			end if
			if f_picshowtf=1 then
				if f_obj("isPicNews")=1 then
					str_newstitle = str_newstitle & "[图]"
				end if
			end if		
			
			s_Content = replace(s_Content,"{NS:FS_NewsTitleAll}",str_newstitle)
		end if
		dim news_SavePath,s_Query_rs,news_Domain,news_UrlDomain,news_ClassEname,s_all_savepath
		if instr(s_Content,"{NS:FS_NewsURL}")>0 or  instr(s_Content,"{NS:FS_Content}")>0 then
				Set f_NewsLinkRecordSet = New CLS_FoosunRecordSet
				Set f_NewsLinkRecordSet.Values(m_NewsLinkFields) = f_obj
				s_all_savepath = get_NewsLink(f_NewsLinkRecordSet)
				Set f_NewsLinkRecordSet = Nothing
				s_NewsPathUrl = s_all_savepath
			s_Content = replace(s_Content,"{NS:FS_NewsURL}",s_NewsPathUrl)
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_CurtTitle}")>0 then
			s_Content = replace(s_Content,"{NS:FS_CurtTitle}",""&f_obj("CurtTitle")&"")
		end if
				'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_NewsNaviContent}")>0 then
			if f_subsys_ListType="readnews" then
				s_Content = replace(s_Content,"{NS:FS_NewsNaviContent}",""&f_obj("NewsNaviContent"))
			else
				s_Content = replace(s_Content,"{NS:FS_NewsNaviContent}",replace(replace(GetCStrLen(""&f_obj("NewsNaviContent")&"",f_navinumber),"&nbsp;",""),vbCrLf,"")&"")
			end if
		end if
				'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_Content}")>0 then
			if f_obj("isUrl")=1 then
				s_Content = replace(s_Content,"{NS:FS_Content}","<a href="""&f_obj("URLAddress")&""">"&f_obj("URLAddress")&"</a>")
			else
				if f_subsys_ListType="readnews" then
					if instr(f_obj("Content"),"[FS:PAGE]")>0 then
						s_Content=SetAdPicContent(0,f_obj,Invert(s_Content))'
					else
						s_Content=SetAdPicContent(1,f_obj,Invert(s_Content))'
					end if
				else
					s_Content = replace(s_Content,"{NS:FS_Content}",replace(replace(GetCStrLen(""&replace(""&Lose_Html(Invert(f_obj("Content"))&""),"[FS:PAGE]","")&"",f_contentnumber),"&nbsp;",""),vbCrLf,"")&"...<a href="""& s_NewsPathUrl &""">详细内容</a>")
				end if
			end if
		end if
				'_新闻添加日期__________________________________________________________________
		if instr(s_Content,"{NS:FS_AddTime}")>0 then
			dim tmp_f_datestyle
			tmp_f_datestyle = f_datestyle
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
			s_Content = replace(s_Content,"{NS:FS_AddTime}",""&tmp_f_datestyle&"")
		end if
				'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_Author}")>0 then
			dim mid_str
			mid_str = mid(f_obj("NewsProperty"),7,1)
			set Rs_Authorobj = Conn.execute("select top 1 G_Name,G_Email from FS_NS_General where G_Type=3 and G_Name='"& f_obj("Author") &"'")
			if Rs_Authorobj.eof then
				if  mid_str="1" then
					s_Content = replace(s_Content,"{NS:FS_Author}","<a href=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"&G_USER_DIR&"/ShowUser.asp?UserName="& f_obj("Author") &""" target=""_blank"">"&f_obj("Author")&"</a>")
				else
					s_Content = replace(s_Content,"{NS:FS_Author}",""&f_obj("Author")&"")
				end if
				Rs_Authorobj.close:set Rs_Authorobj=nothing
			else
				if  mid_str="1" then
					s_Content = replace(s_Content,"{NS:FS_Author}","<a href=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"&G_USER_DIR&"/ShowUser.asp?UserName="& f_obj("Author") &""" target=""_blank"">"&f_obj("Author")&"</a>")
				else
					s_Content = replace(s_Content,"{NS:FS_Author}","<a href=""mailto:"&Rs_Authorobj("G_Email")&""">"&f_obj("Author")&"</a>")
				end if
			Rs_Authorobj.close:set Rs_Authorobj=nothing
			end if
		end if
				'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_Editer}")>0 then
			s_Content = replace(s_Content,"{NS:FS_Editer}",""&f_obj("Editor")&"")
		end if
				'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_hits}")>0 then
			if f_subsys_ListType="readnews" then
				dim ajax_str,hits_str
				hits_str = "<span id=""NS_id_click_"&f_obj("NewsId")&"""></span><script language=""JavaScript"" src=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Click_ajax.asp?type=js&SubSys=NS&spanid=NS_id_click_"&f_obj("NewsId")&"""></script>"
			else
				hits_str = "" & f_obj("hits")
			end if
			s_Content = replace(s_Content,"{NS:FS_hits}",hits_str)
		end if
				'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_KeyWords}")>0 then
			if trim(f_obj("Keywords"))<>"" and Not isNull(trim(f_obj("Keywords"))) then
				k_tmp_Chararray = split(f_obj("Keywords"),",")
				k_tmp_Char=""
				k_tmp_uchar= ""
				for k_i = 0 to UBound(k_tmp_Chararray)
					if k_i=UBound(k_tmp_Chararray) then
						k_tmp_Char = k_tmp_Char & "<a href=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/search.html?keyword="& k_tmp_Chararray(k_i) &"&Type=ns"" target=""_blank"">"& k_tmp_Chararray(k_i) &"</a>"
						k_tmp_uchar= k_tmp_uchar &  k_tmp_Chararray(k_i)
					else
						k_tmp_Char = k_tmp_Char & "<a href=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/search.html?keyword="& k_tmp_Chararray(k_i) &"&Type=ns"" target=""_blank"">"& k_tmp_Chararray(k_i) &"</a>&nbsp;&nbsp;"
						k_tmp_uchar= k_tmp_uchar &  k_tmp_Chararray(k_i) &","
					end if
				next
				s_Content = replace(s_Content,"{NS:FS_KeyWords}",""&k_tmp_Char&"")
				s_Content = replace(s_Content,"{NS:FS_TitleKeyWords}}",""&k_tmp_uchar&"")
			else
				s_Content = replace(s_Content,"{NS:FS_KeyWords}","")
				s_Content = replace(s_Content,"{NS:FS_TitleKeyWords}}","")
			end if
		end if
				'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_TxtSource}")>0 then
			set Rs_Authorobj = Conn.execute("select top 1 G_Name,G_URL from FS_NS_General where G_Type=2 and G_Name='"& f_obj("Source") &"'")
			if Rs_Authorobj.eof then
				s_Content = replace(s_Content,"{NS:FS_TxtSource}",""&f_obj("Source")&"")
			else
				s_Content = replace(s_Content,"{NS:FS_TxtSource}","<a href="""&Rs_Authorobj("G_URL")&""" target=""_blank"">"&f_obj("Source")&"</a>")
			end if
			Rs_Authorobj.close:set Rs_Authorobj=nothing
		end if
				'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_SmallPicPath}")>0 then
			if trim(f_obj("NewsSmallPicFile"))<>"" then
				s_Content = replace(s_Content,"{NS:FS_SmallPicPath}",""&f_obj("NewsSmallPicFile"))
			else
				s_Content = replace(s_Content,"{NS:FS_SmallPicPath}","")
			end if
		end if
				'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_PicPath}")>0 then
			if trim(f_obj("NewsPicFile"))<>"" then
				s_Content = replace(s_Content,"{NS:FS_PicPath}",""&f_obj("NewsPicFile"))
			else
				s_Content = replace(s_Content,"{NS:FS_PicPath}","http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/sys_images/nopic_supply.gif")
			end if
		end if
				'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_FormReview}")>0 then
			If Cint(Mid(f_obj("NewsProperty"),5,1)) = 1 Then
				s_Content = replace(s_Content,"{NS:FS_FormReview}","<label id=""Review_TF_"& f_obj("ID") &""">loading...</label><script language=""JavaScript"" type=""text/javascript"" src=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/ReviewTF.asp?Id="& f_obj("ID") &"&Type=NS""></script>")
			Else
				s_Content = replace(s_Content,"{NS:FS_FormReview}","")
			End if	
		end If
				'______________________________________________________上下篇___________________________________________
		If InStr(s_Content,"{NS:FS_PrevPage}")>0 Then
			s_Content = replace(s_Content,"{NS:FS_PrevPage}",NewsMove(f_obj("Id"),"Up",f_obj("ClassID")))
		End If
				'_________________________________________________________________________________________________
		If InStr(s_Content,"{NS:FS_NextPage}")>0 Then
			s_Content = replace(s_Content,"{NS:FS_NextPage}",NewsMove(f_obj("Id"),"",f_obj("ClassID")))
		End If
		'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_ReviewTF}")>0 then
			if f_obj("isShowReview")=1 then
				s_Content = replace(s_Content,"{NS:FS_ReviewTF}","<a href=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/ReviewUrl.asp?Type=ns&Id="&f_obj("ID")&""">评论</a>")
			else
				s_Content = replace(s_Content,"{NS:FS_ReviewTF}","")
			end if
		end if
				'_________________________________________________________________________________________________
		If InStr(s_Content,"{NS:FS_ReviewURL}")>0 Then
				s_Content = replace(s_Content,"{NS:FS_ReviewURL}","<a href=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/ShowReviewList.asp?type=NS&Id="&f_obj("ID")&""">评</a>")
		End if
			   '_________________________________________________________________________________________________
		If instr(s_Content,"{NS:FS_ShowComment}")>0 Then
			If Cint(Mid(f_obj("NewsProperty"),5,1)) = 1 Then
				s_Content = replace(s_Content,"{NS:FS_ShowComment}","<label id=""NS_show_review_"& f_obj("ID") &""">评论加载中...</label><script language=""JavaScript"" type=""text/javascript"" src=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/ShowReview.asp?Id="&f_obj("ID")&"&Type=NS&SpanId=NS_show_review_"& f_obj("ID") &"""></script>")
			Else
				s_Content = replace(s_Content,"{NS:FS_ShowComment}","")
			End if	
		End If
		If instr(s_Content,"{NS:FS_AddFavorite}")>0 Then
			s_Content = replace(s_Content,"{NS:FS_AddFavorite}","http://"& Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/User/AddFavor.asp?Id="&f_obj("ID")&"&Type=ns")
		End If
				'_________________________________________________________________________________________________
		If instr(s_Content,"{NS:FS_SendFriend}")>0 Then
			s_Content = replace(s_Content,"{NS:FS_SendFriend}","http://"& Request.Cookies("FoosunMFCookies")("FoosunMFDomain") &"/"&G_USER_DIR&"/Sendmail.asp?Id="&f_obj("NewsId")&"&Type=ns")
		End If
		'_________________________________________________________________________________________________
		If instr(s_Content,"{NS:FS_SpecialList}")>0 Then
			if trim(get_SpecialEName)<>"" then
				tmp_list = ""
				s_array = split(get_SpecialEName,",")
				for s_t_i = 0 to ubound(s_array)
							set s_m_Rs=Conn.execute("select  SpecialID,SpecialCName,SpecialEName,ExtName,SavePath,FileSaveType From FS_NS_Special where SpecialEName='"&trim(s_array(s_t_i))&"' order by SpecialID desc")
							if not s_m_Rs.eof then
								Set f_SpecialLinkRecordSet = New CLS_FoosunRecordSet
								Set f_SpecialLinkRecordSet.Values(m_SpecialLinkFields) = s_m_Rs
								f_SpecialLink = get_specialLink(f_SpecialLinkRecordSet)
								Set f_SpecialLinkRecordSet = Nothing
								tmp_list = tmp_list &"<a href="""& f_SpecialLink &""">" & s_m_Rs("SpecialCName") &"</a>&nbsp;"
							else
								tmp_list = tmp_list
							end if
							s_m_Rs.close:set s_m_Rs=nothing
				next
				tmp_list = tmp_list
			end if
			s_Content = replace(s_Content,"{NS:FS_SpecialList}",tmp_list)
		end if
		'获得栏目地址_______________________________________________________________________________________________
		s_f_classSql = "select IsURL,ClassID,ClassName,ClassEName,[Domain],ClassNaviContent,ClassNaviPic,SavePath,FileSaveType,ClassKeywords,Classdescription,FileExtName,UrlAddress from FS_NS_NewsClass where ClassId='"&f_obj("ClassId")&"' and ReycleTF=0 order by OrderID desc,id desc"
		set m_Rs_class = Conn.execute(s_f_classSql)
		if Not m_Rs_class.Eof then
			s_Content = Get_ClassPage(m_Rs_class,s_Content,f_subsys_ListType)
		else
			s_Content = replace(s_Content,"{NS:FS_ClassURL}","")
			s_Content = replace(s_Content,"{NS:FS_ClassName}","")
			s_Content = replace(s_Content,"{NS:FS_ClassNaviPicURL}","")
			s_Content = replace(s_Content,"{NS:FS_ClassNaviDescript}","")
			s_Content = replace(s_Content,"{NS:FS_ClassNaviContent}","")
			s_Content = replace(s_Content,"{NS:FS_ClassKeywords}","")
			s_Content = replace(s_Content,"{NS:FS_Classdescription}","")
		end if
		m_Rs_class.close
		set m_Rs_class=nothing
		'这里暂时用*代替
		'以下为通用专题使用##############################################################################
		Dim m_Rs_special,m_sp_sql,array_special,i_special,s_SpecialName,m_save_special,special_UrlDomain
		if instr(s_Content,"{NS:FS_SpecialName}")>0 then
			if trim(get_SpecialEName)<>"" then
				s_SpecialName = ""
				array_special = split(get_SpecialEName,",")
				for i_special = 0 to Ubound(array_special)
					m_sp_sql = "select SpecialID,SpecialCName,SpecialEName,SpecialContent,SavePath,ExtName,isLock,naviPic,FileSaveType From FS_NS_Special where 1=1 and SpecialEName='"&trim(array_special(i_special))&"'"
					set m_Rs_special=Conn.execute(m_sp_sql)
					if not m_Rs_special.eof then
						Set f_SpecialLinkRecordSet = New CLS_FoosunRecordSet
						Set f_SpecialLinkRecordSet.Values(m_SpecialLinkFields) = Query_rs
						m_save_special = get_specialLink(f_SpecialLinkRecordSet)
						Set f_SpecialLinkRecordSet = Nothing
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
				s_Content = replace(s_Content,"{NS:FS_SpecialName}",s_SpecialName)
			else
				s_Content = replace(s_Content,"{NS:FS_SpecialName}","")
			end if
		end if
		
		'以下是自定义字段替换
		if instr(s_Content,"{NS=Define|")>0 then
			dim define_rs_sql,define_rs
			define_rs_sql="select ID,TableEName,ColumnName,ColumnValue,InfoID,InfoType From FS_MF_DefineData where InfoType='NS' and InfoID='"&f_obj("NewsId")&"' order by ID desc"
			set define_rs=Conn.execute(define_rs_sql)
			if not define_rs.eof  then
				do while not define_rs.eof
					s_Content = replace(s_Content,"{NS=Define|"&define_rs("TableEName")&"}",""&define_rs("ColumnValue"))
					define_rs.movenext
				loop
				define_rs.close:set define_rs=nothing
			else
				dim define_class_sql,define_class_rs
				define_class_sql="select D_Coul From FS_MF_DefineTable where D_SubType='NS' order  by DefineID desc"
				set define_class_rs=Conn.execute(define_class_sql)
				if not define_class_rs.eof then
					do while not define_class_rs.eof
						s_Content = replace(s_Content,"{NS=Define|"&define_class_rs("D_Coul")&"}","")
						define_class_rs.movenext
					loop
				end if
				define_class_rs.close:set define_class_rs=nothing
				define_rs.close:set define_rs=nothing
			end if
		end if
		getlist_news = s_Content
	End Function
	'替换单页标签内容
	Public Function Get_ClassPage(f_obj,s_Content,f_subsys_ListType)
		Dim f_ClassLinkRecordSet,class_path
		If instr(s_Content,"{NS:FS_ClassURL}")>0 then
			Set f_ClassLinkRecordSet = New CLS_FoosunRecordSet
			Set f_ClassLinkRecordSet.Values(m_ClassLinkFields) = f_obj
			class_path = get_ClassLink(f_ClassLinkRecordSet)
			Set f_ClassLinkRecordSet = Nothing
			s_Content = replace(s_Content,"{NS:FS_ClassURL}",class_path)
		End If
		if instr(s_Content,"{NS:FS_ClassName}")>0 then
			s_Content = replace(s_Content,"{NS:FS_ClassName}",""&f_obj("ClassName")&"")
		end if
		if instr(s_Content,"{NS:FS_ClassNaviPicURL}")>0 then
			s_Content = replace(s_Content,"{NS:FS_ClassNaviPicURL}",""& f_obj("ClassNaviPic") &"")
		end if
		if instr(s_Content,"{NS:FS_ClassNaviDescript}")>0 then
			s_Content = replace(s_Content,"{NS:FS_ClassNaviDescript}",""& f_obj("ClassNaviContent") &"")
		end if
		if instr(s_Content,"{NS:FS_ClassNaviContent}")>0 then
			s_Content = replace(s_Content,"{NS:FS_ClassNaviContent}",""&f_obj("ClassNaviContent"))
		end if
		if instr(s_Content,"{NS:FS_ClassKeywords}")>0 then
			s_Content = replace(s_Content,"{NS:FS_ClassKeywords}",""&f_obj("ClassKeywords"))
		end if
		if instr(s_Content,"{NS:FS_Classdescription}")>0 then
			s_Content = replace(s_Content,"{NS:FS_Classdescription}",""&f_obj("Classdescription"))
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_PageContent}")>0 then
			if f_subsys_ListType="classpage" then
				if instr(f_obj("classNaviContent"),"[FS:PAGE]")>0 then
					s_Content = Replace(s_Content,"{NS:FS_PageContent}","[FS:CONTENT_START]" & invert(f_obj("classNaviContent") & "") & "[FS:CONTENT_END]")
				else
					s_Content = Replace(s_Content,"{NS:FS_PageContent}",invert(f_obj("classNaviContent") & ""))
				end if
			else
				s_Content = Replace(s_Content,"{NS:FS_PageContent}",invert(f_obj("classNaviContent") & ""))
			end if
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_PageName}")>0 then
			s_Content = replace(s_Content,"{NS:FS_PageName}",f_obj("ClassName"))
		end if	'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_Keywords}")>0 then
			s_Content = replace(s_Content,"{NS:FS_Keywords}",f_obj("ClassKeywords"))
		end if	'_________________________________________________________________________________________________
		if instr(s_Content,"{NS:FS_description}")>0 then
			s_Content = replace(s_Content,"{NS:FS_description}",f_obj("Classdescription"))
		end if
		Get_ClassPage = s_Content
	End Function
	'得到新闻单个地址____________________________________________________________
	Public Function get_NewsLink(f_NewsLinkRecordSet)
		Dim f_NewsLink
		Set f_NewsLink = New CLS_FoosunLink
		get_NewsLink = f_NewsLink.NewsLink(f_NewsLinkRecordSet)
		Set f_NewsLink = Nothing
	End Function
		
	'/////////////////////////////////////////////////
	'//NewsMove获取新闻的上一篇和下一篇
	'//NewsId 新闻ID
	'//NewsType 上下篇的标注UP为上篇 空为下篇
	Function NewsMove(NewsId,NewsType,Newsclass)
		Dim RsObj,SqlStr,Id,f_NewsLinkRecordSet,f_NewsLink,f_WhereSQL
		'SqlStr = "Select top 1 NewsTitle,ClassEName,[Domain],SavePath,News.IsURL,News.URLAddress,SaveNewsPath,News.FileName,News.FileExtName From FS_NS_News as News,FS_NS_NewsClass as Class Where News.ClassID=Class.ClassID And News.Id>"&NoSqlHack(NewsId)&" And News.ClassID='"&Newsclass&"' and isRecyle=0 And isLock=0"
		'If NewsType = "Up" Then SqlStr = SqlStr & " order by News.id desc"
		If NewsType = "Up" Then
			f_WhereSQL = " And News.ID<" & NewsId
		else
			f_WhereSQL = " And News.ID>" & NewsId
		end if
		SqlStr = "Select top 1 NewsTitle,ClassEName,[Domain],SavePath,News.IsURL,News.URLAddress,SaveNewsPath,News.FileName,News.FileExtName From FS_NS_News as News,FS_NS_NewsClass as Class Where News.ClassID=Class.ClassID " & f_WhereSQL & " And News.ClassID='"&Newsclass&"' and isRecyle=0 And isLock=0"
		If NewsType = "Up" Then SqlStr = SqlStr & " order by News.id desc"
		Set RsObj = Conn.Execute(SqlStr)
		if Rsobj.Eof Then
			NewsMove = "无"
			Rsobj.Close
			Set RsObj = Nothing
		Else
			Set f_NewsLinkRecordSet = New CLS_FoosunRecordSet
			Set f_NewsLinkRecordSet.Values(m_NewsLinkFields) = RsObj
			f_NewsLink = get_NewsLink(f_NewsLinkRecordSet)
			Set f_NewsLinkRecordSet = Nothing
			NewsMove = "<a href=""" & f_NewsLink & """>"&RsObj("NewsTitle")&"</a>"
			Rsobj.Close
			Set RsObj = Nothing
		End if
	End Function 

	'得到栏目地址________________________________________________________________
	Public function get_ClassLink(f_ClassLinkRecordSet)
		Dim f_ClassLink
		Set f_ClassLink = New CLS_FoosunLink
		get_ClassLink = f_ClassLink.ClassLink(f_ClassLinkRecordSet)
		Set f_ClassLink = Nothing
	End function

	'得到专题地址________________________________________________________________
	Public function get_specialLink(f_SpecialLinkRecordSet)
		Dim f_SpecialLink
		Set f_SpecialLink = New CLS_FoosunLink
		get_specialLink = f_SpecialLink.SpecialLink(f_SpecialLinkRecordSet)
		Set f_SpecialLink = Nothing
	End function

	'得到子类____________________________________________________________________
	Public Function get_ClassList(TypeID,CompatStr,f_css)
		Dim ChildTypeListRs,ChildTypeListStr,TempStr,i,Sql,f_ClassLinkRecordSet,f_ClassLink
		Set ChildTypeListRs = Server.CreateObject(G_FS_RS)
		Sql = "Select ParentID,ClassID,ClassName,IsURL,ClassEName,[Domain],FileExtName,FileSaveType,SavePath,UrlAddress from FS_NS_NewsClass where ParentID='" & TypeID & "' and ReycleTF=0 order by OrderID desc,id desc"
		ChildTypeListRs.open sql,Conn,0,1
		TempStr = CompatStr
		do while Not ChildTypeListRs.Eof
			Set f_ClassLinkRecordSet = New CLS_FoosunRecordSet
			Set f_ClassLinkRecordSet.Values(m_ClassLinkFields) = ChildTypeListRs
			f_ClassLink = get_ClassLink(f_ClassLinkRecordSet)
			Set f_ClassLinkRecordSet = Nothing
			get_ClassList = get_ClassList & TempStr
			get_ClassList = get_ClassList & "<img src="""&m_PathDir&"sys_images/-.gif"" border=""0""><a href="""&f_ClassLink&""" class="""& f_css &""">"&ChildTypeListRs("ClassName")&"</a>"&vbNewLine
			get_ClassList = get_ClassList & get_ClassList(ChildTypeListRs("ClassID"),TempStr,f_css)
			ChildTypeListRs.MoveNext
		loop
		ChildTypeListRs.Close:Set ChildTypeListRs = Nothing
	End Function
		
	Public Function get_SubClass(classid)
		Dim ChildTypeListRs
		Set ChildTypeListRs = Conn.ExeCute("Select ClassID from FS_NS_NewsClass where ParentID='" & classid & "' and ReycleTF=0 order by OrderID desc,id desc")
		If Not ChildTypeListRs.Eof Then
			do while Not ChildTypeListRs.Eof
				get_SubClass =get_SubClass & "','" & get_SubClass(ChildTypeListRs(0))
				ChildTypeListRs.MoveNext
			loop
			get_SubClass = classid & get_SubClass
		Else
			get_SubClass = 	classid
		End If
		ChildTypeListRs.Close : Set ChildTypeListRs = Nothing
		get_SubClass = get_SubClass
	End Function

	'得到option子类______________________________________________________________
	Public Function get_optionNewsList(TypeID,CompatStr)
		Dim ChildTypeListRs,ChildTypeListStr,TempStr
		Set ChildTypeListRs = Conn.execute("Select ParentID,ClassID,ClassName from FS_NS_NewsClass where ParentID='" & TypeID & "' and ReycleTF=0 order by OrderID desc,id desc")
		TempStr = CompatStr & "┉"
		do while Not ChildTypeListRs.Eof
			get_optionNewsList = get_optionNewsList &"<option value="""&ChildTypeListRs("ClassId")&""">"& TempStr
			get_optionNewsList = get_optionNewsList & "┉"&ChildTypeListRs("ClassName")&"</option>"&vbNewLine
			get_optionNewsList = get_optionNewsList & get_optionNewsList(ChildTypeListRs("ClassID"),TempStr)
			ChildTypeListRs.MoveNext
		loop
		ChildTypeListRs.Close:Set ChildTypeListRs = Nothing
	End Function
		
	Public Function getNewsSubClass(typeID)
		Dim childClassRs,result,Str_SubClassID
		result=""
		Set childClassRs=Conn.execute("Select ParentID,ClassID from FS_NS_NewsClass where ParentID='" & TypeID & "' and ReycleTF=0 order by OrderID desc,id desc")
		While Not childClassRs.eof
			If result="" Then
				result="'"&childClassRs("classID")&"'"
			Else
				result=result&",'"&childClassRs("classID")&"'"
			End If
			Str_SubClassID=getNewsSubClass(childClassRs("classID"))
			If Str_SubClassID<>"" Then
				result=result&","&Str_SubClassID
			End If
			childClassRs.movenext
		Wend
		childClassRs.close:Set childClassRs=nothing
		getNewsSubClass=result
	End function
		
	Public Function SetAdPicContent(IsPage,f_obj,s_Content)
		Dim Temp_Content_Str,Replace_Str
		Temp_Content_Str = f_obj("Content") & ""
		If ISinnerLink Then
			Replace_Str = ReplaceInnerLink(Temp_Content_Str)
		Else
			Replace_Str = Temp_Content_Str
		End If
		
		If Cint(f_obj("IsAdPic"))=1 Then
			Dim LeftContent,MidAdContent,RightContent,ModifyContent,headlen,tempStr,headAdStr,tailAdStr,ShowFanXiang,Str_Content,Str_ContentArray,i,TempLen,tempad_Link
			Str_ContentArray=Split(Replace_Str,"[FS:PAGE]")
			For i=Lbound(Str_ContentArray) To Ubound(Str_ContentArray)
				'截取字符串
				If Not IsNumeric(split(f_obj("AdPicWH"),",")(3)) Then
					TempLen=200
				Else
					TempLen=split(f_obj("AdPicWH"),",")(3)
				End If
				tempStr=Str_ContentArray(i)
				LeftContent=InterceptString(tempStr,TempLen)
				'获取实际的截取的长度
				RightContent=Right(Str_ContentArray(i),Len(Str_ContentArray(i))-Len(LeftContent))
				ShowFanXiang=split(f_obj("AdPicWH"),",")(2)
				If ShowFanXiang="" Or IsNull(ShowFanXiang) Then
					ShowFanXiang="left"
				Else
					If ShowFanXiang="1" Then
						ShowFanXiang="left"
					Else
						ShowFanXiang="right"
					End If
				End If
				If Lcase(Right(f_obj("AdPicAdress"),3))="swf" Then'判断是否Swf图片
					tailAdStr="<table width=0 border=0 align="&ShowFanXiang&"><tr><td><object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0"" width="""&split(f_obj("AdPicWH"),",")(1)&""" height="""&split(f_obj("AdPicWH"),",")(0)&"""><param name=""movie"" value="""&f_obj("AdPicAdress")&"""><param name=""quality"" value=""high""><embed src="""&f_obj("AdPicAdress")&""" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash""  width="""&split(f_obj("AdPicWH"),",")(1)&""" height="""&split(f_obj("AdPicWH"),",")(0)&"""></embed></object></td></tr></table>"
				Else
					If f_obj("AdPicLink")="" or IsNull(f_obj("AdPicLink")) Then
						tempad_Link="http://www.foosun.cn"
					Else
						tempad_Link=f_obj("AdPicLink")
					End If
					tailAdStr="<table width=0 border=0 align="&ShowFanXiang&"><tr><td><a href="&tempad_Link&"><img border=0 src="&f_obj("AdPicAdress")&" height="&split(f_obj("AdPicWH"),",")(0)&" width="&split(f_obj("AdPicWH"),",")(1)&"></a></td></tr></table>"
				End If
				Str_ContentArray(i)=LeftContent & tailAdStr & RightContent
			Next
			Str_Content=""
			For i=Lbound(Str_ContentArray) To Ubound(Str_ContentArray)
				If Str_Content="" Then
					Str_Content=Str_ContentArray(i)
				Else
					Str_Content=Str_Content&"[FS:PAGE]"&Str_ContentArray(i)
				End If
			Next
			If IsPage=1 Then
				s_Content = Replace(s_Content,"{NS:FS_Content}",Str_Content)
			Else
				s_Content = replace(s_Content,"{NS:FS_Content}","[FS:CONTENT_START]"&Str_Content&"[FS:CONTENT_END]")
			End If
		Else
			If IsPage=1 Then
				s_Content = replace(s_Content,"{NS:FS_Content}",invert(Replace_Str))
			Else
				s_Content = replace(s_Content,"{NS:FS_Content}","[FS:CONTENT_START]"&invert(Replace_Str)&"[FS:CONTENT_END]")
			End If
		End If
		SetAdPicContent=s_Content
	End Function
		
	'"截取字符串___________________________________________________________________
	Public Function InterceptString(txt,length)
		Dim x,y,ii,c,ischines,isascii,tempStr
		length=Cint(length)
		txt=trim(txt)
		x = len(txt)
		y = 0
		if x >= 1 then
			for ii = 1 to x
				c=asc(mid(txt,ii,1))
				if  c< 0 or c >255 then
					y = y + 2
					ischines=1
					isascii=0
				else
					y = y + 1
					ischines=0
					isascii=1
				end if
				if y >= length then
					if ischines=1 and StrCount(left(trim(txt),ii),"<a")=StrCount(left(trim(txt),ii),"</a>") then
						txt = left(txt,ii) '"字符串限长
						exit for
					else
						if isascii=1 then x=x+1
					end if
				end if
			next
			InterceptString = txt
		else
			InterceptString = ""
		end if
	End Function
		
	'判断字符串出现的次数
	Public Function StrCount(Str,SubStr)        
		Dim iStrCount
		Dim iStrStart
		Dim iTemp
		iStrCount = 0
		iStrStart = 1
		iTemp = 0
		Str=LCase(Str)
		SubStr=LCase(SubStr)
		Do While iStrStart < Len(Str)
			iTemp = Instr(iStrStart,Str,SubStr,vbTextCompare)
			If iTemp <=0 Then
				iStrStart = Len(Str)
			Else
				iStrStart = iTemp + Len(SubStr)
				
				iStrCount = iStrCount + 1
			End If
		Loop
		StrCount = iStrCount
	End Function
		
	Public Function ReplaceInnerLink(NewsContent)
		Dim RoutineSql,RsRoutineObj
		RoutineSql = "Select * from FS_NS_General where G_Type = 4 and isLock=0"
		Set RsRoutineObj = server.CreateObject(G_FS_RS)
		RsRoutineObj.CursorLocation = adUseClient
		RsRoutineObj.Open RoutineSql,Conn,0,1
		Dim StrReplace,Inti,DLocation,XLocation
		dim isValue,i
		isValue=false
		for i = 1 to RsRoutineObj.recordcount
				Inti=1
				StrReplace=RsRoutineObj("G_Name")
				If instr(1,NewsContent,StrReplace) then
					do while instr(Inti,NewsContent,StrReplace)<>0
						if isvalue=true then 
							exit do
						end if
						Inti=instr(Inti,NewsContent,StrReplace)
						If Inti<>0 then
							DLocation=instr(Inti,NewsContent,">")'如果内容在><之间则替换
							XLocation=instr(Inti,NewsContent,"<")
							If DLocation >= XLocation Then
								If instr(1,"[FS:PAGE]",StrReplace)=0 then'避免替换[FS:PAGE]里面的内容后，造成分页混乱
									NewsContent=left(NewsContent,Inti-1) & "<a href=" & RsRoutineObj("G_URL") & " target=_blank>" & StrReplace & "</a>" & mid(NewsContent,Inti+len(StrReplace))
									Inti=Inti+len("<a href=" & RsRoutineObj("G_URL") & " target=_blank>" & StrReplace & "</a>")
									isValue=true
								Else
									Inti=Inti+len(StrReplace)
								End If
							Else
								Inti=Inti+len(StrReplace)
							end If
						end if
					loop
				End If
				isValue=false
			RsRoutineObj.MoveNext
		next
		RsRoutineObj.Close : Set RsRoutineObj = Nothing
		ReplaceInnerLink = NewsContent
	End Function
		
	Public Function ISinnerLink()
		Dim LinkRs
		Set LinkRs = Conn.Execute("Select InsideLink From FS_NS_SysParam")
		If LinkRs(0) = 0 Then
			ISinnerLink = False
		Else
			ISinnerLink = True
		End If
		LinkRs.Close : Set LinkRs = NoThing		 
	End Function
End Class
%> 
<%
Class cls_MS
	Private m_Rs,m_fSO,m_Dict
	Private m_PathDir,m_Path_UserDir,m_Path_User,m_Path_adminDir,m_Path_UserPageDir,m_Path_Templet
	Private m_Err_Info,m_Err_NO

	Public Property Get Err_Info()
		Err_Info = m_Err_Info
	End Property

	Public Property Get Err_NO()
		Err_NO = m_Err_NO
	End Property

	Private Sub Class_initialize()
		Set m_Rs = Server.CreateObject(G_fS_RS)
		Set m_fSO = Server.CreateObject(G_fS_fSO)
		Set m_Dict = Server.CreateObject(G_fS_DICT)
		m_PathDir = replace("/"&G_VIRTUAL_ROOT_DIR&"/","//","/")
		m_Path_UserDir = replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_USER_DIR&"/","//","/")
		m_Path_UserPageDir = replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_USERfILES_DIR&"/","//","/")
		m_Path_Templet = replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_TEMPLETS_DIR&"/","//","/")
	End Sub

	Private Sub Class_Terminate()
		Set m_Rs = Nothing
		Set m_fSO = Nothing
		Set m_Dict = Nothing
	End Sub

	Public function get_LableChar(f_Lable,f_Id,f_RefreshPageType)
		select case LCase(f_Lable.LableFun)
			case "classproducts"
				get_LableChar=Classproducts(f_Lable,"classproducts",f_Id)
			case "specialproducts"
				get_LableChar=Classproducts(f_Lable,"specialproducts",f_Id)
			case "lastproducts"
				get_LableChar=Classproducts(f_Lable,"lastproducts",f_Id)
			case "hotproducts"
				get_LableChar=Classproducts(f_Lable,"hotproducts",f_Id)
			case "recproducts"
				get_LableChar=Classproducts(f_Lable,"recproducts",f_Id)
			case "marproducts"
				get_LableChar=Classproducts(f_Lable,"marproducts",f_Id)
			case "briproducts"
				get_LableChar=Classproducts(f_Lable,"briproducts",f_Id)
			case "annproducts"
				get_LableChar=Classproducts(f_Lable,"annproducts",f_Id)
			case "constrproducts"
				get_LableChar=Classproducts(f_Lable,"constrproducts",f_Id)
			case "classlist"
				get_LableChar=classlist(f_Lable,"classlist",f_Id)
			case "speciallist"
				get_LableChar=classlist(f_Lable,"speciallist",f_Id)
			case "oneprice"
				get_LableChar=Classproducts(f_Lable,"onePrice",f_Id)
			case "publicsale"
				get_LableChar=Classproducts(f_Lable,"publicsale",f_Id)
			case "specialoffer"
				get_LableChar=Classproducts(f_Lable,"specialoffer",f_Id)
			case "sales"
				get_LableChar=Classproducts(f_Lable,"sales",f_Id)
			case "flashfilt"
				get_LableChar=flashfilt(f_Lable,"flashfilt",f_Id)
			case "norfilt"
				get_LableChar=Norfilter(f_Lable,"Norfilt",f_Id)
			case "readproducts"
				get_LableChar=Readproducts(f_Lable,"readproducts",f_Id)
			case "sitemap"
				get_LableChar=SiteMap(f_Lable,"sitemap",f_Id)
			case "search"
				get_LableChar=Search(f_Lable,"search")
			case "infostat"
				get_LableChar=infoStat(f_Lable,"infostat")
			case "todaypic"
				get_LableChar=TodayPic(f_Lable,"todaypic",f_Id)
			case "todayword"
				get_LableChar=TodayWord(f_Lable,"todayword",f_Id)
			case "classnavi"
				get_LableChar=ClassNavi(f_Lable,"classnavi",f_Id)
			case "specialnavi"
				get_LableChar=SpecialNavi(f_Lable,"specialnavi",f_Id)
			case "rssfeed"
				get_LableChar=Rssfeed(f_Lable,"rssfeed",f_Id)
			case "specialcode"
				get_LableChar=SpecialCode(f_Lable,"specialcode",f_Id)
			case "classcode"
				get_LableChar=ClassCode(f_Lable,"classcode",f_Id)
			case "defineproducts"
				get_LableChar=Defineproducts(f_Lable,"defineproducts",f_Id)
			case "oldproducts"
				get_LableChar=Oldproducts(f_Lable,"oldproducts",f_Id)
			case "classinfo"
				get_LableChar=ClassInfo(f_Lable,"ClassInfo",f_Id)
		end select
	End function

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
	Public function table_str_list_head(f_tf,f_divid,f_divclass,f_ulid,f_ulclass)
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
	End function

	'得到div,table_________________________________________________________________________________
	public function table_str_list_middle_1(f_tf,f_liid,f_liclass)
		Dim f_liid_1,f_liclass_1,td_
		if f_tf=1 then
			if f_liid<>"" then:f_liid_1 = " id="""& f_liid &"""":else:f_liid_1 = "":end if
			if f_liclass<>"" then:f_liclass_1 = " class="""& f_liclass &"""":else:f_liclass_1 = "":end if
			td_="<li"&f_liid_1&f_liclass_1&">"
			table_str_list_middle_1 ="  "&td_
		end if
	End function

	'得到div,table_________________________________________________________________________________
	public function table_str_list_middle_2(f_tf)
		Dim td__,tr__
		if f_tf=1 then
			td__="</li>"
		else
			td__="</td>"
		end if
			table_str_list_middle_2 =  td__&chr(10)
	End function

	'得到div,table_________________________________________________________________________________
	Public function table_str_list_middle_3(f_tf)
		if f_tf=1 then
			table_str_list_middle_3 = ""
		else
			table_str_list_middle_3 = "</tr>"
		end if
	End function


	'得到div,table_________________________________________________________________________________
	Public function table_str_list_bottom(f_tf)
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
	End function
	'开始读取标签____综合类标签_____________________________________________________________________
	Public function Classproducts(f_Lable,f_LableType,f_Id)
		Dim classproducts_head,classproducts_bottom,classproducts_middle1,classproducts_middle2,div_tf,style_Content,Content_List,f_rs_s_obj,Content_more,c_i_k
		Dim newnumber,classid,orderby,orderdesc,colnumber,contentnumber,navinumber,datenumber,titlenumber,picshowtf,datenumber_tmp,morechar,datestyle,openstyle,open_target,containSubClass,childClass
		Dim f_sql,f_rs_obj,f_rs_configobj,f_configSql,CharIndexStr
		Dim f_Mf_ConfigSql,f_rs_Mf_configobj,Mf_Domain,marqueedirec,marqueespeed,marqueestyle
		Dim avePath,IsDomain,SavePathRule,search_str,No_order,ordertype
		Dim ClassCName,ClassEName,c_Domain,NaviContent,NaviPic,c_SavePath,c_fileSaveType,search_inSQL,ClassSaveType
		m_Err_Info = ""
		No_order = f_Lable.LablePara("序号排列")
		search_str = f_Lable.LablePara("栏目")
		newnumber = f_Lable.LablePara("Loop")
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
		navinumber= f_Lable.LablePara("导航字数")
		if f_LableType="marproducts" then
			marqueespeed=f_Lable.LablePara("滚动速度")
			marqueedirec=f_Lable.LablePara("滚动方向")
		end if
		if f_Lable.LablePara("输出格式") = "out_DIV" then
			div_tf=1
			No_order="0"
		else
			div_tf=0
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
		classproducts_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classproducts_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classproducts_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
		classproducts_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)

		if datenumber ="0" then:datenumber_tmp = "":else:datenumber_tmp = " and dateadd(d,"&datenumber&",addtime)>='"&datevalue(now())&"'":end If
		dim childrenClassIds:childrenClassIds = trim(DelHeadAndEndDot(getSubClass(search_str)))
		If containSubClass="1" and childrenClassIds<>"" Then
			childClass=" or classid in ("&childrenClassIds&")"
		Else
			childClass=""
		End if
		select case LCase(f_LableType)
			case "classproducts"
				If childClass<>"" then
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
			case "hotproducts"
				'Call Getfunctionstr
				if trim(search_str)<>"" Then
					If childClass<>"" then
						search_inSQL = " and "& all_substring &"(StyleflagBit,3,1)='1' and (ClassId='"& search_str &"'"&childClass&")"
					Else
						search_inSQL = " and "& all_substring &"(StyleflagBit,3,1)='1' and ClassId='"& search_str &"'"
					End if
				else
					if trim(f_id)<>"" Then
						If childClass<>"" then
							search_inSQL = " and "& all_substring &"(StyleflagBit,3,1)='1' and (ClassId='"& f_id &"'"&childClass&")"
						Else
							search_inSQL = " and "& all_substring &"(StyleflagBit,3,1)='1' and ClassId='"& f_id &"'"
						End if
					else
						search_inSQL = " and "& all_substring &"(StyleflagBit,3,1)='1'"
					end if
				end if
			case "lastproducts"
				if trim(search_str)<>"" Then
					If childClass<>"" then
						search_inSQL = " and (ClassId='"& search_str &"'"&childClass&")"
					Else
						search_inSQL = " and ClassId='"& search_str &"'"
					End if
				else
					if trim(f_id)<>"" Then
						If childClass<>"" then
							search_inSQL = " and (ClassId='"& f_id &"'"&childClass&")"
						Else
							search_inSQL = " and ClassId='"& f_id &"'"
						End if
					else
						search_inSQL = ""
					end if
				end if
			case "recproducts"
				if trim(search_str)<>"" Then
					If childClass<>"" then
						search_inSQL = " and "& all_substring &"(StyleflagBit,1,1)='1' and (ClassId='"& search_str &"'"&childClass&")"
					Else
						search_inSQL = " and "& all_substring &"(StyleflagBit,1,1)='1' and ClassId='"& search_str &"'"
					End if
				else
					if trim(f_id)<>"" Then
						If childClass<>"" then
							search_inSQL = " and "& all_substring &"(StyleflagBit,1,1)='1' and (ClassId='"& f_id &"'"&childClass&")"
						Else
							search_inSQL = " and "& all_substring &"(StyleflagBit,1,1)='1' and ClassId='"& f_id &"'"
						End if
					else
						search_inSQL = " and "& all_substring &"(StyleflagBit,1,1)='1'"
					end if
				end if
			case "marproducts"
				if trim(search_str)<>"" Then
					If childClass<>"" then
						search_inSQL = " and "& all_substring &"(StyleflagBit,9,1)='1' and (ClassId='"& search_str &"'"&childClass&")"
					Else
						search_inSQL = " and "& all_substring &"(StyleflagBit,9,1)='1' and ClassId='"& search_str &"'"
					End if
				else
					if trim(f_id)<>"" Then
						If childClass<>"" then
							search_inSQL = " and "& all_substring &"(StyleflagBit,9,1)='1' and (ClassId='"& f_id &"'"&childClass&")"
						Else
							search_inSQL = " and "& all_substring &"(StyleflagBit,9,1)='1' and ClassId='"& f_id &"'"
						End if
					else
						search_inSQL = " and "& all_substring &"(StyleflagBit,9,1)='1'"
					end if
				end if
			case "specialoffer"
				if trim(search_str)<>"" Then
					If childClass<>"" then
						search_inSQL = " and SaleStyle=4 and (ClassId='"& search_str &"'"&childClass&")"
					Else
						search_inSQL = " and SaleStyle=4 and ClassId='"& search_str &"'"
					End if
				else
					if trim(f_id)<>"" Then
						If childClass<>"" then
							search_inSQL = " and SaleStyle=4 and (ClassId='"& f_id &"'"&childClass&")"
						Else
							search_inSQL = " and SaleStyle=4 and ClassId='"& f_id &"'"
						End if
					else
						search_inSQL = " and SaleStyle=4"
					end if
				end if
			case "sales"
				if trim(search_str)<>"" Then
					If childClass<>"" then
						search_inSQL = " and SaleStyle=3 and (ClassId='"& search_str &"'"&childClass&")"
					Else
						search_inSQL = " and SaleStyle=3 and ClassId='"& search_str &"'"
					End if
				else
					if trim(f_id)<>"" then
						If childClass<>"" then
							search_inSQL = " and SaleStyle=3 and (ClassId='"& f_id &"'"&childClass&")"
						Else
							search_inSQL = " and SaleStyle=3 and ClassId='"& f_id &"'"
						End if
					else
						search_inSQL = " and SaleStyle=3"
					end if
				end if
			case "oneprice"
				if trim(search_str)<>"" Then
					If childClass<>"" Then
						search_inSQL = " and SaleStyle=2 and (ClassId='"& search_str &"'"&childClass&")"
					Else
						search_inSQL = " and SaleStyle=2 and ClassId='"& search_str &"'"
					End if
				else
					if trim(f_id)<>"" Then
						If childClass<>"" Then
							search_inSQL = " and SaleStyle=2 and (ClassId='"& f_id &"'"&childClass&")"
						Else
							search_inSQL = " and SaleStyle=2 and ClassId='"& f_id &"'"
						End if
					else
						search_inSQL = " and SaleStyle=2"
					end if
				end if
			case "publicsale"
				if trim(search_str)<>"" Then
					If childClass<>"" Then
						search_inSQL = " and SaleStyle=1 and (ClassId='"& search_str &"'"&childClass&")"
					Else
						search_inSQL = " and SaleStyle=1 and ClassId='"& search_str &"'"
					End if
				else
					if trim(f_id)<>"" Then
						If childClass<>"" then
							search_inSQL = " and SaleStyle=1 and (ClassId='"& f_id &"'"&childClass&")"
						Else
							search_inSQL = " and SaleStyle=1 and ClassId='"& f_id &"'"
						End if
					else
						search_inSQL = " and SaleStyle=1"
					end if
				end if
			case "filtproducts"
				if trim(search_str)<>"" Then
					If childClass<>"" then
						search_inSQL = " and "& all_substring &"(StyleflagBit,7,1)='1' and (ClassId='"& search_str &"'"&childClass&")"
					Else
						search_inSQL = " and "& all_substring &"(StyleflagBit,7,1)='1' and ClassId='"& search_str &"'"
					End if
				else
					if trim(f_id)<>"" Then
						If childClass<>"" Then
							search_inSQL = " and "& all_substring &"(StyleflagBit,7,1)='1' and (ClassId='"& f_id &"'"&childClass&")"
						Else
							search_inSQL = " and "& all_substring &"(StyleflagBit,7,1)='1' and ClassId='"& f_id &"'"
						End if
					else
						search_inSQL = " and "& all_substring &"(StyleflagBit,7,1)='1'"
					end if
				end if
		end select
		Dim Order_ID_
		If LCase(orderby) <> "id" Then
			Order_ID_ = ",ID " & orderdesc
		Else
			Order_ID_ = ""
		End If
		f_sql="select top "& cint(newnumber) &" ID,ProductTitle,TitleStyle,Barcode,Serialnumber,ClassID,Keyword,SpecialID,Stockpile,StockpileWarn,OldPrice,NewPrice,IsWholesale,ProductContent,RepairContent,Templetfile,Makefactory,ProductsAddress,IsInvoice,Click,MakeTime,SavePath,fileName,fileExtName,smallPic,BigPic,StyleflagBit,SaleStyle,AddTime,Discount,DiscountStartDate,DiscountEndDate,ReycleTf,AddMember,saleNumber,popid,isShowReview,Mail_Money"
		f_sql = f_sql &" from fS_MS_products where ReycleTF=0 And "&all_substring&"(StyleflagBit,5,1)='0'"&search_inSQL & datenumber_tmp &" order by "&orderby&" "&orderdesc & Order_ID_
		f_Mf_ConfigSql="select top 1 Mf_Domain from fS_Mf_Config"

		Set  f_rs_obj = Conn.execute(f_sql)
		set f_rs_Mf_configobj=Conn.execute(f_Mf_ConfigSql)
		Mf_Domain = f_rs_Mf_configobj("Mf_Domain")
		f_rs_Mf_configobj.close:set f_rs_Mf_configobj=nothing
		Content_List=""
		if f_rs_obj.eof then
			Content_List=""
			m_Err_Info = "没有相关记录"
			f_rs_obj.close:set f_rs_obj=nothing
		else
			if f_LableType="marproducts" then
				Content_List = Content_List & "<marquee onmouseover=""this.stop();"" scrollamount="""& marqueespeed &""" direction="""& marqueedirec &""" onmouseout=""this.start();"">"
				do while not f_rs_obj.eof
					Content_List= Content_List &getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType)&"&nbsp;&nbsp;&nbsp;"
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
				Select Case No_order
					Case "0"
						   'this
						do while not f_rs_obj.eof
								if div_tf=1 then
									Content_List= Content_List & classproducts_middle1 & getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType) & classproducts_middle2
								else
									if cint(colnumber) =1 then
										Content_List= Content_List & chr(10)&"   <tr><td>" & getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType) & "</td></tr>"
									else
										Content_List= Content_List & "<td width="""& cint(100/cint(colnumber))&"%"">" & getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType) & "</td>"
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
						   'this
					Case "1"
						ordertype = "64"
						do while not f_rs_obj.eof
								ordertype = ordertype + 1
								if div_tf=1 then
									Content_List= Content_List & classproducts_middle1 & getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType) & classproducts_middle2
								else
									if cint(colnumber) =1 then
										If Cint(newnumber)<=26 Then
											Content_List= Content_List & chr(10)&"   <tr><td>" &Chr(ordertype)&"."& getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType) & "</td></tr>"
										Else
											Content_List= Content_List & chr(10)&"   <tr><td>" & getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType) & "</td></tr>"
										End if
									else
										Content_List= Content_List & "<td width="""& cint(100/cint(colnumber))&"%"">" & getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType) & "</td>"
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
					Case "2"
						ordertype = "96"
						do while not f_rs_obj.eof
								ordertype = ordertype + 1
								if div_tf=1 then
									Content_List= Content_List & classproducts_middle1 & getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType) & classproducts_middle2
								else
									if cint(colnumber) =1 then
										If Cint(newnumber)<=26 Then
											Content_List= Content_List & chr(10)&"   <tr><td>" &Chr(ordertype)&"."& getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType) & "</td></tr>"
										Else
											Content_List= Content_List & chr(10)&"   <tr><td>" & getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType) & "</td></tr>"
										End if
									else
										Content_List= Content_List & "<td width="""& cint(100/cint(colnumber))&"%"">" & getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType) & "</td>"
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
					Case "3"
					   ordertype = "0"
						do while not f_rs_obj.eof
								ordertype = ordertype + 1
								if div_tf=1 then
									Content_List= Content_List & classproducts_middle1 & getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType) & classproducts_middle2
								else
									if cint(colnumber) =1 then
										If Cint(newnumber)<=26 Then
											Content_List= Content_List & chr(10)&"   <tr><td>" &ordertype&"."& getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType) & "</td></tr>"
										Else
											Content_List= Content_List & chr(10)&"   <tr><td>" & getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType) & "</td></tr>"
										End if
									else
										Content_List= Content_List & "<td width="""& cint(100/cint(colnumber))&"%"">" & getlist_products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,Mf_Domain,f_LableType) & "</td>"
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
				End Select
			   if div_tf=0 then
					if cint(colnumber)<>1 then Content_List = Content_List & "</tr>"&chr(10)
			   end if
			end if
			if f_LableType="classproducts" then
				dim Query_rs,productclass_SavePath,fileSaveType,UrlDomain,all_savepath,LinkType
				set Query_rs=Conn.execute("select ClassEName,SavePath,fileExtName,[Domain],fileSaveType,IsURL,UrlAddress from fS_MS_productsClass where ClassId='"& search_str &"'")

				if Query_rs.eof then
						m_Err_Info = "商品栏目已经被删除"
						Query_rs.close:set Query_rs=nothing
				else
					if Query_rs("isURL")=1 then
						all_savepath = Query_rs("UrlAddress")
					else
						productclass_SavePath = Query_rs("SavePath")
						if Query_rs("fileSaveType")=0 then
							fileSaveType = Query_rs("ClassEName")&"/index."&Query_rs("fileExtName")
						elseif Query_rs("fileSaveType")=1 then
							fileSaveType = Query_rs("ClassEName")&"/"& Query_rs("ClassEName") &"."&Query_rs("fileExtName")
						else
							fileSaveType =  Query_rs("ClassEName") &"."&Query_rs("fileExtName")
						end if
						if LinkType=1 then
								if  trim(Query_rs("Domain"))="" then
									UrlDomain = "http://" & IsDomain
								else
									UrlDomain = "http://" & Query_rs("Domain")
								end if
						else
								if  trim(Query_rs("Domain"))<>"" then
									UrlDomain = "http://" & Query_rs("Domain")
								else
									UrlDomain = ""
								end if
						end if
						all_savepath = UrlDomain & productclass_SavePath & "/" & fileSaveType
					  end if
					  Query_rs.close:set Query_rs=nothing
				end if
				if openstyle=0 then
					open_target=" "
				else
					open_target=" target=""_blank"""
				end if
				if div_tf=1 then
						Content_more = "  <li><a href="""&  all_savepath &""" "& open_target &" title=""更多..."">"&morechar&"</a></li>"&chr(10)
				else
						Content_more = "  <tr><td colspan="""& cint(colnumber) &"""><div align=""right""><a href="""&  all_savepath &""" "& open_target &" title=""更多..."">"&morechar&"</a></div></td></tr>"&chr(10)
				end if
			elseif f_LableType="specialproducts" then
				dim special_rs,special_Path
				set special_rs =Conn.execute("select SpecialID,SpecialCName,SpecialEName,naviText,SavePath,FileExtName,isLock,naviPic from fS_MS_Special where isLock=0 and SpecialEName='"&trim(search_str)&"'")
				if Not special_rs.eof then
						if LinkType=1 then
							'Mf_Domain,IsDomain
							if trim(IsDomain)<>"" then
								m_PathDir = IsDomain & m_PathDir
							else
								m_PathDir = Mf_Domain & m_PathDir
							end if
							special_Path = "http://"&replace(m_PathDir & special_rs("SavePath")&"/special_"&special_rs("SpecialEName")&"."&special_rs("ExtName"),"//","/")
						else
							m_PathDir = m_PathDir
							special_Path = replace(m_PathDir & special_rs("SavePath")&"/special_"&special_rs("SpecialEName")&"."&special_rs("ExtName"),"//","/")
						end if
						special_rs.close:set special_rs=nothing
				end if
				if div_tf=1 then
						Content_more = "  <li><a href="""&  special_Path &""" "& open_target &" title=""更多..."">"&morechar&"</a></li>"&chr(10)
				else
						Content_more = "  <tr><td colspan="""& cint(colnumber) &"""><div align=""right""><a href="""&  special_Path &""" "& open_target &" title=""更多..."">"&morechar&"</a></div></td></tr>"&chr(10)
				end if
			 end if
			 f_rs_obj.close:set f_rs_obj=nothing
		end if
		if f_LableType="marproducts" then
			Content_List= Content_List
		else
			Content_List=classproducts_head & Content_List & Content_more & classproducts_bottom
		end if
		f_Lable.DictLableContent.Add "1",Content_List&" "
	End function

	'开始读取标签____商品终极类标签_____________________________________________________________________
	Public function ClassList(f_Lable,f_type,f_Id)
		if f_Id <> "" then
			dim div_tf,newnumber,datenumber,titlenumber,picshowtf,openstyle,orderby,orderdesc,pageTf,pagestyle,pagecss,datestyle,colnumber,contentnumber,navinumber,pagenumber
			dim classproducts_head,classproducts_bottom,classproducts_middle1,classproducts_middle2,search_inSQL,ClassSaveType
			dim datenumber_tmp,f_sql,f_configsql,f_mf_configsql,f_rs_obj,f_rs_configobj,f_rs_Mf_configobj,Mf_Domain
			dim SavePath,IsDomain,SavePathRule,LinkType,TPageNum,perPageNum,PageNum,sPageCount,cl_i,c_i_k,Page_flag_TF
			Dim f_TableName,f_SelectFieldNames,f_PageIndex,f_Where,f_PaginationStr,f_NewsContent,f_IDSRS,f_IDSArray
			Dim f_IDS,i,f_BeginIndex,rs_c,RefreshNumber
			m_Err_Info = ""
			Dim Inc_List,RefClassid
			Inc_List = f_Lable.LablePara("包含子类")
			datenumber= f_Lable.LablePara("多少天")
			titlenumber= f_Lable.LablePara("标题数")
			picshowtf=f_Lable.LablePara("图文标志")
			openstyle=f_Lable.LablePara("打开窗口")
			orderby = f_Lable.LablePara("排列字段")
			orderdesc = f_Lable.LablePara("排列方式")
			pageTF = f_Lable.LablePara("分页")
			pagestyle = f_Lable.LablePara("分页样式")
			pagenumber = f_Lable.LablePara("每页数量")
			pagecss = f_Lable.LablePara("PageCSS")
			datestyle = f_Lable.LablePara("日期格式")
			colnumber= f_Lable.LablePara("排列数")
			contentnumber= f_Lable.LablePara("内容字数")
			navinumber= f_Lable.LablePara("导航字数")
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
			classproducts_head = List_Addtional_HTML("div",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classproducts_middle1 = List_Addtional_HTML("li",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classproducts_bottom = List_Addtional_HTML("div1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)
			classproducts_middle2 = List_Addtional_HTML("li1",div_tf,f_DivID,f_DivClass,f_UlID,f_ULClass,f_LiID,f_LiClass)

			f_PaginationStr = pagestyle & "," & pagecss
			f_Lable.DictLableContent.Item("0") = f_PaginationStr

			if f_type="speciallist" then
				search_inSQL = " and specialID="&f_Id
				RefreshNumber = ""
			else
				dim rs2,Classlist_Classid
				if f_id="" then f_id=0
				if not isnumeric(f_Id) then
					search_inSQL=" and ClassId in('"& f_Id &"')"
				Else
					set rs2 = Conn.execute("select ClassID From FS_MS_ProductsClass where id="& f_Id &"")
					if not rs2.eof then search_inSQL=" and ClassId in('"&rs2("ClassID")&"')"
					rs2.close:set rs2 =nothing
				End if
				RefreshNumber = ""
				RefClassid = Trim(Replace(Replace(search_inSQL," and ClassId in('",""),"')",""))
				If Inc_List = "1" Then
					dim tmpClassIDs
					tmpClassIDs = Inc_Classlist(RefClassid)
					if tmpClassIDs<>"" then tmpClassIDs = RefClassid & tmpClassIDs
					search_inSQL = " and ClassId in('"& tmpClassIDs &"') "
				End if
				if Trim(Inc_Classlist(RefClassid)) = "" Then
					search_inSQL ="  and ClassId in('"&RefClassid&"')"
				End if
			end if
			if G_IS_SQL_DB=0 then
				if datenumber ="0" then:datenumber_tmp = "":else:datenumber_tmp = " and datevalue(addtime)+"&datenumber&">=datevalue(now)":end if
			Elseif G_IS_SQL_DB=1 then
				if datenumber ="0" then:datenumber_tmp = "":else:datenumber_tmp = " and dateadd(d,"&datenumber&",addtime)>='"&datevalue(now())&"'":end if
			End if
			MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
			f_configSql="select top 1 SavePath,IsDomain,SavePathRule from fS_MS_SysPara"
			f_Where = "ReycleTF=0 "&search_inSQL & datenumber_tmp &" order by "&orderby&" "&orderdesc
			f_TableName = "fS_MS_products"
			f_SelectFieldNames = "ID,ProductTitle,TitleStyle,Barcode,Serialnumber,ClassID,Keyword,SpecialID,Stockpile,StockpileWarn,OldPrice,NewPrice,IsWholesale,ProductContent,RepairContent,Templetfile,Makefactory,ProductsAddress,IsInvoice,Click,MakeTime,SavePath,fileName,fileExtName,smallPic,BigPic,StyleflagBit,SaleStyle,AddTime,Discount,DiscountStartDate,DiscountEndDate,ReycleTF,AddMember,saleNumber,popid,isShowReview,Mail_Money"

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
					ClassList = classProducts_head
					Do While Not f_rs_obj.Eof
						f_NewsContent = getlist_Products(f_rs_obj,f_Lable.FoosunStyle.StyleContent,titlenumber,contentnumber,navinumber,picshowtf,datestyle,openstyle,MF_Domain,"ClassList")
						if div_tf = 1 then
							ClassList= ClassList & classProducts_middle1 & f_NewsContent & classProducts_middle2
							'if CutNum = 0 then
							'	ClassList= ClassList & classProducts_middle1 & f_NewsContent & classProducts_middle2
							'else
							'	if cl_i mod CutNum = 0 then
							'		ClassList= ClassList & classProducts_middle1 & f_NewsContent & classProducts_middle2 & CutType
							'	else
							'		ClassList= ClassList & classProducts_middle1 & f_NewsContent & classProducts_middle2
							'	end if
							'end If
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
					Loop
					ClassList = ClassList & classProducts_bottom
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
			f_Lable.DictLableContent.Add "1",""
		end if
	End function
	Rem 获取终极列表的子类Crazy
	Function Inc_Classlist(Classid)
		Dim Rsobj,Sqlstr,Classidarr
		Sqlstr = "Select ID,ClassId,ParentID From FS_MS_ProductsClass Where ParentID='"&Classid&"'"
		Set Rsobj = Conn.Execute(Sqlstr)
		Do While Not Rsobj.Eof
				Classidarr = Classidarr&"','"&Rsobj("ClassId") & Inc_Classlist(Rsobj("ClassId"))
			Rsobj.Movenext
		Loop
		Inc_Classlist = Classidarr
	End Function
	'flash幻灯片
	Public function flashfilt(f_Lable,f_type,f_Id)
		dim Randomize_i_Str,TitleNumberStr,filterSql,ClassId,InSQL_Search,RsfilterObj,str_filt,productsNumberStr,flashStr
		dim ClassSavefilePath,ImagesStr,TxtStr,Txtfirst,LinkStr
		dim Class_SQL,RsClassObj,Temp_Num
		dim ContainSubClass,ClassIDStr
		Randomize
		Randomize_i_Str =  CStr(Int((999 * Rnd) + 1))
		TitleNumberStr= f_Lable.LablePara("标题字数")
		productsNumberStr= f_Lable.LablePara("数量")
		ContainSubClass= f_Lable.LablePara("包含子类")
		if trim(productsNumberStr)<>"" and  isNumeric(productsNumberStr) then
			productsNumberStr = cint(productsNumberStr)
		else
			productsNumberStr = 6
		end if
		ClassId=f_Lable.LablePara("栏目")
		if trim(ClassId)<>"" and ContainSubClass=1 then
			ClassIDStr = get_SubClass(ClassId)
			InSQL_Search = " and ClassId in ('" & ClassIDStr & "')"
		Elseif trim(ClassId)<>"" then
			InSQL_Search = " and ClassId='"& ClassId &"'"
		else
			if f_Id <> "" then
				InSQL_Search = " and ClassId='"& f_Id &"'"
			else
				InSQL_Search = ""
			end if
		end if
		str_filt=" and "& all_substring &"(StyleflagBit,7,1)='1'"
		Set RsfilterObj = Server.CreateObject(G_FS_RS)
		filterSql="select top "& productsNumberStr &" ID,ProductTitle,TitleStyle,Barcode,Serialnumber,ClassID,Keyword,SpecialID,Stockpile,StockpileWarn,OldPrice,NewPrice,IsWholesale,ProductContent,RepairContent,Templetfile,Makefactory,ProductsAddress,IsInvoice,Click,MakeTime,SavePath,fileName,fileExtName,smallPic,BigPic,StyleflagBit,SaleStyle,AddTime,Discount,DiscountStartDate,DiscountEndDate,ReycleTf,AddMember,saleNumber,popid,isShowReview,Mail_Money"
		filterSql = filterSql &" from fS_MS_products where ReycleTf=0 "&InSQL_Search & str_filt &"  order by addtime desc,id desc"
		RsfilterObj.CursorLocation = adUseClient
		RsfilterObj.Open filterSql,Conn,0,1
		If not RsfilterObj.Eof Then
			Temp_Num = RsfilterObj.Recordcount
			If Temp_Num <=1 then
				Set RsfilterObj = Nothing
				f_Lable.DictLableContent.Add "1","至少需要两条幻灯文章才能正确显示幻灯效果"
				Set	RsfilterObj = Nothing
				Exit function
			End If
			Do While Not RsfilterObj.Eof
				if (Not IsNull(RsfilterObj("smallPic"))) And (RsfilterObj("smallPic") <> "") Then
					If ImagesStr = "" Then
						ImagesStr = RsfilterObj("smallPic")
					Else
						ImagesStr = ImagesStr &"|"& RsfilterObj("smallPic")
					End If
					If TxtStr = "" Then
						TxtStr = GotTopic(RsfilterObj("ProductTitle"),TitleNumberStr)
					Else
						TxtStr = TxtStr & "|" & GotTopic(RsfilterObj("ProductTitle"),TitleNumberStr)
					End If
					If LinkStr = "" Then
						LinkStr = get_productsLink(RsfilterObj("ID"))
					Else
						LinkStr = LinkStr & "|" & get_productsLink(RsfilterObj("ID"))
					End If
				End If
				RsfilterObj.MoveNext
			Loop
			flashStr="<script type=""text/javascript"">"& Chr(13)
			flashStr=flashStr&" <!--"& Chr(13)
			dim PicSize,PicWidthStr,PicHeightStr,txtheight
			PicSize = f_Lable.LablePara("图片尺寸") '
			PicWidthStr = split(PicSize,",")(1)
			PicHeightStr = split(PicSize,",")(0)
			txtheight = f_Lable.LablePara("文本高度")
			flashStr=flashStr&" var focus_width"&Randomize_i_Str&"="&PicWidthStr& Chr(13)
			flashStr=flashStr&" var focus_height"&Randomize_i_Str&"="&PicHeightStr& Chr(13)
			flashStr=flashStr&" var text_height"&Randomize_i_Str&"="&txtheight& Chr(13)
			flashStr=flashStr&" var swf_height"&Randomize_i_Str&" = focus_height"&Randomize_i_Str&"+text_height"&Randomize_i_Str& Chr(13)
			flashStr=flashStr&" var pics"&Randomize_i_Str&"='"&ImagesStr&"'"&Chr(13)
			flashStr=flashStr&" var links"&Randomize_i_Str&"='"&LinkStr &"'"&Chr(13)
			flashStr=flashStr&" var texts"&Randomize_i_Str&"='"&TxtStr&"'"&Chr(13)
			FlashStr=FlashStr&" document.write('<object ID=""focus_flash"" classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0"" width=""'+focus_width"&Randomize_i_Str&"+'"" height=""'+swf_height"&Randomize_i_Str&"+'"">');"&Chr(13)
			flashStr=flashStr&" document.write('<param name=""allowScriptAccess"" value=""sameDomain""><param name=""movie"" value="""&m_Path_Templet & "flash.swf""><param name=""quality"" value=""high""><param name=""bgcolor"" value=""white"">');"&Chr(13)
			flashStr=flashStr&" document.write('<param name=""menu"" value=""false""><param name=wmode value=""opaque"">');"&Chr(13)
			flashStr=flashStr&" document.write('<param name=""flashVars"" value=""pics='+pics"&Randomize_i_Str&"+'&links='+links"&Randomize_i_Str&"+'&texts='+texts"&Randomize_i_Str&"+'&borderwidth='+focus_width"&Randomize_i_Str&"+'&borderheight='+focus_height"&Randomize_i_Str&"+'&textheight='+text_height"&Randomize_i_Str&"+'"">');"&Chr(13)
			flashStr=flashStr&" document.write('<embed ID=""focus_flash"" src="""&m_Path_Templet & "flash.swf"" wmode=""opaque"" flashVars=""pics='+pics"&Randomize_i_Str&"+'&links='+links"&Randomize_i_Str&"+'&texts='+texts"&Randomize_i_Str&"+'&borderwidth='+focus_width"&Randomize_i_Str&"+'&borderheight='+focus_height"&Randomize_i_Str&"+'&textheight='+text_height"&Randomize_i_Str&"+'"" menu=""false"" bgcolor=""white"" quality=""high"" width=""'+ focus_width"&Randomize_i_Str&" +'"" height=""'+ swf_height"&Randomize_i_Str&" +'"" allowScriptAccess=""sameDomain"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" />');"&Chr(13)
			flashStr=flashStr&" document.write('</object>');"&Chr(13)
			flashStr=flashStr&" //-->"& Chr(13)
			flashStr=flashStr&"</script>"
		else
			flashStr=""
		end if
		RsfilterObj.Close
		Set RsfilterObj = Nothing
		f_Lable.DictLableContent.Add "1",flashStr
	End function
	'轮换幻灯
	Public function Norfilter(f_Lable,f_type,f_id)
		dim filterSql,RsfilterObj,filterStr,ImagesStr,TxtStr,Txtfirst,ClassSavefilePath,LinkStr,CssfileStr,PicWidthStr, PicHeightStr
		dim fltheightstr,fltwidthstr,Temp_Num,TitleNumberStr,productsNumberStr,ClassId,InSQL_Search,str_filt
		dim ChildClassTF,AllClassIDStr,PicSize
		TitleNumberStr= f_Lable.LablePara("标题字数")
		productsNumberStr= f_Lable.LablePara("数量")
		ChildClassTF = f_Lable.LablePara("包含子类")
		if trim(productsNumberStr)<>"" and  isNumeric(productsNumberStr) then
			productsNumberStr = cint(productsNumberStr)
		else
			productsNumberStr = 6
		end if
		If ChildClassTF = "" Or Not IsNumeric(ChildClassTF) Then
			ChildClassTF = 0
		Else
			ChildClassTF = Cint(ChildClassTF)
		End If
		ClassId=f_Lable.LablePara("栏目")
		if trim(ClassId)<>"" then
			If ChildClassTF = 1 Then
				AllClassIDStr = get_SubClass(ClassId)
				InSQL_Search = " and ClassId in('" & FormatIntArr(AllClassIDStr) & "')"
			Else
				InSQL_Search = " and ClassId='"& ClassId &"'"
			End if
		else
			if f_Id <> "" then
				InSQL_Search = " and ClassId='"& f_Id &"'"
			else
				InSQL_Search = ""
			end if
		end if
		str_filt=" and "& all_substring &"(StyleflagBit,7,1)='1'"
		filterSql="select top "& productsNumberStr &" ID,ProductTitle,TitleStyle,Barcode,Serialnumber,ClassID,Keyword,SpecialID,Stockpile,StockpileWarn,OldPrice,NewPrice,IsWholesale,ProductContent,RepairContent,Templetfile,Makefactory,productsAddress,IsInvoice,Click,MakeTime,SavePath,fileName,fileExtName,smallPic,BigPic,StyleflagBit,SaleStyle,AddTime,Discount,DiscountStartDate,DiscountEndDate,ReycleTf,AddMember,saleNumber,popid,isShowReview,Mail_Money"
		filterSql = filterSql &" from fS_MS_products where ReycleTf=0 "&InSQL_Search & str_filt &"  order by addtime desc,id desc"
		Set RsfilterObj = Conn.Execute(filterSql)
		if not RsfilterObj.Eof then
			Temp_Num = 0
			Do While Not RsfilterObj.Eof
				Temp_Num = Temp_Num + 1
				RsfilterObj.MoveNext
			Loop
			RsfilterObj.Movefirst
			If Temp_Num <=1 then
				Set RsfilterObj = Nothing
				Norfilter = "至少需要两条幻灯新闻才能正确显示幻灯效果"
				Set	RsfilterObj = Nothing
				Exit function
			End If
			PicSize = f_Lable.LablePara("图片尺寸")
			fltheightstr = split(PicSize,",")(0)
			fltwidthstr = split(PicSize,",")(1)
			CssfileStr = f_Lable.LablePara("CSS样式")
			PicWidthStr = " width=""" & fltwidthstr & """"
			PicHeightStr = " height=""" & fltheightstr & """"
			if CssfileStr <> "" then CssfileStr = " Class='" & CssfileStr & "'"
			do while Not RsfilterObj.Eof
				if (Not IsNull(RsfilterObj("smallPic"))) And (RsfilterObj("smallPic") <> "") then
					if ImagesStr = "" then
						If Instr(1,LCase(RsfilterObj("smallPic")),"http://") <> 0 then
							ImagesStr =  RsfilterObj("smallPic")
						Else
							ImagesStr =   RsfilterObj("smallPic")
						End If
						TxtStr = "<a "& CssfileStr & " href='" & get_productsLink(RsfilterObj("ID")) & "' target='_blank'>" & GotTopic(RsfilterObj("ProductTitle"),TitleNumberStr)&"</a>"
						Txtfirst = "<a " & CssfileStr & " href='" & get_productsLink(RsfilterObj("ID")) & "' target='_blank'>" & GotTopic(RsfilterObj("ProductTitle"),TitleNumberStr)&"</a>"
						LinkStr =  get_productsLink(RsfilterObj("ID"))
					else
						ImagesStr = ImagesStr &","&  RsfilterObj("smallPic")
						TxtStr = TxtStr &",<a  "& CssfileStr &" href='" & get_productsLink(RsfilterObj("ID")) & "'  target='_blank'>" & GotTopic(RsfilterObj("ProductTitle"),TitleNumberStr)&"</a>"
						LinkStr = LinkStr & "," & get_productsLink(RsfilterObj("ID"))
					end if
				end if
				RsfilterObj.MoveNext
			loop
			filterStr="<script language=""vbscript"">"& Chr(13)
			filterStr = filterStr & "Dim fileList,fileListArr,TxtList,TxtListArr,LinkList,LinkArr"& Chr(13)
			filterStr = filterStr & "fileList = """ & ImagesStr & """"& Chr(13)
			filterStr = filterStr & "LinkList = """ & LinkStr & """"& Chr(13)
			filterStr = filterStr & "TxtList = """ & TxtStr & """"& Chr(13)
			filterStr = filterStr & "fileListArr = Split(fileList,"","")"& Chr(13)
			filterStr = filterStr & "LinkArr = Split(LinkList,"","")"& Chr(13)
			filterStr = filterStr & "TxtListArr = Split(TxtList,"","")"& Chr(13)
			filterStr = filterStr & "Dim CanPlay"& Chr(13)
			filterStr = filterStr & "CanPlay = CInt(Split(Split(navigator.appVersion,"";"")(1),"" "")(2))>5"& Chr(13)
			filterStr = filterStr & "Dim filterStr"& Chr(13)
			filterStr = filterStr & "filterStr = ""RevealTrans(duration=2,transition=23)"""& Chr(13)
			filterStr = filterStr & "filterStr = filterStr + "";BlendTrans(duration=2)"""& Chr(13)
			filterStr = filterStr & "If CanPlay Then"& Chr(13)
			filterStr = filterStr & "filterStr = filterStr + "";progid:DXImageTransform.Microsoft.fade(duration=2,overlap=0)"""& Chr(13)
			filterStr = filterStr & "filterStr = filterStr + "";progid:DXImageTransform.Microsoft.Wipe(duration=3,gradientsize=0.25,motion=reverse)"""& Chr(13)
			filterStr = filterStr & "Else"& Chr(13)
			filterStr = filterStr & "Msgbox ""幻灯片播放具有多种动态图片切换效果，但此功能需要您的浏览器为IE5.5或以上版本，否则您将只能看到部分的切换效果。"",64"& Chr(13)
			filterStr = filterStr & "End If"& Chr(13)
			filterStr = filterStr & "Dim filterArr"& Chr(13)
			filterStr = filterStr & "filterArr = Split(filterStr,"";"")"& Chr(13)
			filterStr = filterStr & "Dim PlayImg_M"& Chr(13)
			filterStr = filterStr & "PlayImg_M = 5 * 1000  "& Chr(13)
			filterStr = filterStr & "Dim I"& Chr(13)

			filterStr = filterStr & "I = 1"& Chr(13)
			filterStr = filterStr & "Sub ChangeImg"& Chr(13)
			filterStr = filterStr & "Do While fileListArr(I)="""""& Chr(13)
			filterStr = filterStr & "I = I + 1"& Chr(13)
			filterStr = filterStr & "If I>UBound(fileListArr) Then I = 0"& Chr(13)
			filterStr = filterStr & "Loop"& Chr(13)
			filterStr = filterStr & "Dim J"& Chr(13)
			filterStr = filterStr & "If I>UBound(fileListArr) Then I = 0"& Chr(13)
			filterStr = filterStr & "Randomize"& Chr(13)
			filterStr = filterStr & "J = Int(Rnd * (UBound(filterArr)+1))"& Chr(13)
			filterStr = filterStr & "Img.style.filter = filterArr(J)"& Chr(13)
			filterStr = filterStr & "Img.filters(0).Apply"& Chr(13)
			filterStr = filterStr & "Img.Src = fileListArr(I)"& Chr(13)
			filterStr = filterStr & "Img.filters(0).play"& Chr(13)
			filterStr = filterStr & "Link.Href = LinkArr(I)"& Chr(13)
			If f_Lable.LablePara("显示标题") = "1" Then
				filterStr = filterStr & "Txt.filters(0).Apply"& Chr(13)
				filterStr = filterStr & "Txt.innerHTML = TxtListArr(I)"& Chr(13)
				filterStr = filterStr & "Txt.filters(0).play"& Chr(13)
			End If
			filterStr = filterStr & "I = I + 1"& Chr(13)
			filterStr = filterStr & "If I>UBound(fileListArr) Then I = 0"& Chr(13)
			filterStr = filterStr & "TempImg.Src = fileListArr(I)"& Chr(13)
			filterStr = filterStr & "TempLink.Href = LinkArr(I)"& Chr(13)
			filterStr = filterStr & "SetTimeout ""ChangeImg"", PlayImg_M,""VBScript"""& Chr(13)
			filterStr = filterStr & "End Sub"& Chr(13)
			filterStr = filterStr & "</SCRIPT>"& Chr(13)
			filterStr = filterStr & "<TABLE WIDTH=""100%"" height=""100%"" BORDER=""0"" CELLSPACING="""" CELLPADDING=""0"">" &vbcrlf
			filterStr = filterStr & "<TR ID=""NoScript"">"&vbcrlf
			filterStr = filterStr & "<TD Align=""Center"" Style=""Color:White"">对不起，图片浏览功能需脚本支持，但您的浏览器已经设置了禁止脚本运行。请您在浏览器设置中调整有关安全选项。</TD>"&vbcrlf
			filterStr = filterStr & "</TR>"&vbcrlf
			filterStr = filterStr & "<TR Style=""Display:none"" ID=""CanRunScript""><TD HEIGHT=""100%"" Align=""Center"" vAlign=""Center""><a id=""Link""><Img ID=""Img"" "  & PicWidthStr & PicHeightStr & " Border=""0"" ></a>"&vbcrlf
			filterStr = filterStr & "</TD></TR><TR Style=""Display:none""><TD><a id=TempLink ><Img ID=""TempImg"" Border=""0""></a></TD></TR>"&vbcrlf
			If f_Lable.LablePara("显示标题") = "1" Then
				filterStr = filterStr & "<TR><TD HEIGHT=""100%"" Align=""Center"" vAlign=""Top"">"&vbcrlf
				filterStr = filterStr & "<div ID=""Txt"" style=""PADDING-LEfT: 5px; Z-INDEX: 1; fILTER: progid:DXImageTransform.Microsoft.fade(duration=1,overlap=0); POSITION:"">"&Txtfirst&"</div>"
				filterStr = filterStr & "</TD></TR>"&vbcrlf
			End If
			filterStr = filterStr & "</TABLE>"& Chr(13)
			filterStr = filterStr & "<Script Language=""VBScript"">"& Chr(13)
			filterStr = filterStr & "NoScript.Style.Display = ""none"""& Chr(13)
			filterStr = filterStr & "CanRunScript.Style.Display = """""& Chr(13)
			filterStr = filterStr & "Img.Src = fileListArr(0)"& Chr(13)
			filterStr = filterStr & "Link.Href = LinkArr(0)"& Chr(13)
			filterStr = filterStr & "SetTimeout ""ChangeImg"", PlayImg_M,""VBScript"""& Chr(13)
			filterStr = filterStr & "</Script>"& Chr(13)
		else
			filterStr="没有幻灯图片"
		End if
		RsfilterObj.Close
		Set RsfilterObj = Nothing
		f_Lable.DictLableContent.Add "1",filterStr
	End function
	'--------------------得到子类栏目id-----------------2/2----by chen--修改----------
	Public Function get_SubClass(classid)
		Dim ChildTypeListRs
		Set ChildTypeListRs = Conn.ExeCute("Select ClassID From FS_MS_ProductsClass where ParentID='" & classid & "' and ReycleTF=0 order by OrderID desc,id desc")
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

	'---------------------------------------------------------------------------------
	'得到新闻浏览地址
	Public function Readproducts(f_Lable,f_type,f_Id)
		Dim ReadSql,RsReadObj,f_configSql,f_Mf_ConfigSql,f_rs_configobj,f_rs_Mf_configobj
		Dim Mf_Domain,productsDir,IsDomain,fileDirRule,LinkType,ClassSaveType,datestyle
		datestyle = f_Lable.LablePara("日期格式")
		Readproducts = ""
		if trim(f_Id)="" then
			Readproducts = ""
		else
			ReadSql="Select ID,ProductTitle,TitleStyle,Barcode,Serialnumber,ClassID,Keyword,SpecialID,Stockpile,StockpileWarn,OldPrice,NewPrice,IsWholesale,ProductContent,RepairContent,Templetfile,Makefactory,ProductsAddress,IsInvoice,Click,MakeTime,SavePath,fileName,fileExtName,smallPic,BigPic,StyleflagBit,SaleStyle,AddTime,Discount,DiscountStartDate,DiscountEndDate,ReycleTf,AddMember,saleNumber,popid,isShowReview,Mail_Money"
			ReadSql = ReadSql &" from fS_MS_products where Id="& f_Id &" and ReycleTf=0"
			Mf_Domain = Request.Cookies("foosunMfCookies")("foosunMfDomain")
			Set  RsReadObj = Server.CreateObject(G_fS_RS)
			RsReadObj.open ReadSql,Conn,0,1
			if not RsReadObj.eof then
				Readproducts = Readproducts&getlist_products(RsReadObj,f_Lable.FoosunStyle.StyleContent,0,0,0,0,datestyle,0,Mf_Domain,"readproducts")
			else
				Readproducts =""
			end if
		end if
		f_Lable.DictLableContent.Add "1",Readproducts
	End function
	'站点地图
	Public function SiteMap(f_Lable,f_type,f_Id)
		dim classId,cssstyle,SiteMapstr,RsClassObj,i,br_str
		classId = f_Lable.LablePara("栏目")
		cssstyle = f_Lable.LablePara("标题CSS")
		If classId = "" Or IsnUll(classId) Then
			classId = 0
		End If
		SiteMapstr = ""
		set RsClassObj = Conn.execute("select ClassId,ClassCName,ClassEName,IsURL,ParentID from fS_MS_productsClass where ReycleTf=0 and ParentID='" & classId & "' order by OrderID desc,id desc")
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
					SiteMapstr = SiteMapstr & br_str & "<img src="""&m_PathDir&"sys_images/+.gif"" border=""0"" /><a href="""&get_ClassLink(RsClassObj("ClassId"))&""" class="""& cssstyle &""">"&RsClassObj("ClassCName")&"</a>"&chr(10)
				else
					SiteMapstr = SiteMapstr & br_str & "<img src="""&m_PathDir&"sys_images/+.gif"" border=""0"" /><a href="""&get_ClassLink(RsClassObj("ClassId"))&""">"&RsClassObj("ClassCName")&"</a>"&chr(10)
				end if
				SiteMapstr = SiteMapstr & get_ClassList(RsClassObj("ClassId"),"&nbsp;",cssstyle)
				RsClassObj.movenext
				i=i+1
			loop
		else
			SiteMapstr = ""
		end if
		RsClassObj.close
		set RsClassObj=nothing
		f_Lable.DictLableContent.Add "1",SiteMapstr
	End function
	'得到搜索表单
	Public function Search(f_Lable,f_type)
		dim Searchstr,showdate,showClass,rs,select_str,datestr,classstr,Mf_Domain
		Dim MinPric,MaxPric
		showdate = f_Lable.LablePara("显示日期")
		showClass = f_Lable.LablePara("显示栏目")
		Mf_Domain = Request.Cookies("foosunMfCookies")("foosunMfDomain")
		Searchstr = ""
		if showdate = "1" then
			datestr = "开始日期:<input name=""s_date"" type=""text"" value="""&date()-1&""" size=""10"" />结束日期：<input name=""e_date"" type=""text"" value="""&date()&""" size=""10"" />"
		else
			datestr = ""
		end if
		MinPric = "最低价格:<input name=""MinPric"" type=""text"" value="""" size=""10"" />"
		MaxPric = "最高价格:<input name=""MaxPric"" type=""text"" value="""" size=""10"" />"
		if showClass = "1" then
			classstr = "<select name=""ClassId"" id=""ClassId"">"
			set rs = Conn.execute("select ClassId,id,ClassCName,classEname,ParentId from fS_MS_productsClass where ReycleTf=0 and isURL=0 and ParentId='0' order by OrderId desc,id desc")
			classstr = classstr & "<option value="""" selected>---选择栏目---</option>"
			do while not rs.eof
				classstr = classstr & "<option value="""&rs("ClassId")&""">┝"&rs("ClassCName")&"</option>"&chr(10)
				classstr = classstr & get_optionProductsList(rs("ClassId"),"┝")&chr(10)
				rs.movenext
			loop
			rs.close:set rs=nothing
			classstr = classstr&"</select>"
		else
			classstr = ""
		end if
		select_str ="<select name=""s_type"" id=""s_type"">"&chr(10)
		select_str=select_str&"<option value=""title"" selected=""selected"">商品名称</option>"&chr(10)
		select_str=select_str&"<option value=""content"">商品介绍</option>"&chr(10)
		select_str=select_str&"<option value=""ProductNumber"">编号</option>"&chr(10)
		select_str=select_str&"<option value=""keyword"">关键字</option>"&chr(10)
		select_str=select_str&"<option value=""source"">来源</option>"&chr(10)
		select_str=select_str&"</select>"&chr(10)
		Searchstr = Searchstr & "<form action=""http://"& Mf_Domain & "/Search.html"" id=""Searchform"" name=""Searchform"" method=""get""><input name=""SubSys"" type=""hidden"" id=""SubSys"" value=""MS"" /><input name=""Keyword"" type=""text"" id=""Keyword"" size=""15"" /> "&select_str&datestr&classstr&MinPric&MaxPric&"<input name=""SearchSubmit_foosun"" type=""submit"" id=""SearchSubmit"" value=""搜索"" /></form>"
		f_Lable.DictLableContent.Add "1",Searchstr
	 End function
	 '得到统计信息
	Public function infoStat(f_Lable,f_type)
		dim cols,br_str,infoStatstr
		cols = f_Lable.LablePara("显示方向")
		if cols="0" then
		br_str = "&nbsp;"
		else
		br_str = "<br />"
		end if
		dim rs,c_rs1,c_rs2,c_rs3,c_rs4,c_rs5
		set rs = Conn.execute("select count(id) from fS_MS_products where ReycleTf=0")
		c_rs1=rs(0)
		rs.close:set rs=nothing
		set rs = Conn.execute("select count(id) from fS_MS_productsClass where ReycleTf=0")
		c_rs2=rs(0)
		rs.close:set rs=nothing
		set rs = Conn.execute("select count(SpecialID) from fS_MS_Special where isLock=0")
		c_rs3=rs(0)
		rs.close:set rs=nothing
		set rs = User_Conn.execute("select count(Userid) from FS_ME_Users")
		c_rs4=rs(0)
		rs.close:set rs=nothing
		if G_IS_SQL_User_DB=0 then
			set rs = User_Conn.execute("select count(Userid) from fS_ME_Users where datevalue(RegTime)=#"&datevalue(date)&"#")
		else
			set rs = User_Conn.execute("select count(Userid) from fS_ME_Users where datediff(d,RegTime,'"&datevalue(date)&"')=0")
		end if
		c_rs5=rs(0)
		rs.close:set rs=nothing
		infoStatstr = "总商品:"&"<strong>"&c_rs1&"</strong>"&br_str
		infoStatstr = infoStatstr &"总栏目:"&"<strong>"&c_rs2&"</strong>"&br_str
		infoStatstr = infoStatstr &"专区数:"&"<strong>"&c_rs3&"</strong>"&br_str
		infoStatstr = infoStatstr &"会员数:"&"<strong>"&c_rs4&"</strong>"&br_str
		infoStatstr = infoStatstr &"今日注册:"&"<strong>"&c_rs5&"</strong>"&br_str
		infoStat = infoStatstr
		f_Lable.DictLableContent.Add "1",infoStatstr
	End function
	'栏目导航
	Public function ClassNavi(f_Lable,f_type,f_Id)
		dim ClassId,cols,titlecss,rs,ClassNavistr,ParentIDstr,cols_str,div_tf,f_SQL
		dim classproducts_head,classproducts_bottom,classproducts_middle1,classproducts_middle2,cssstyle,titleNavi
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
		if trim(ClassId)<>"" then
			ParentIDstr = " and ParentId = '"& ClassId&"'"
		else
			if f_id = "0" OR f_id = "" then
				ParentIDstr = " and ParentId = '0'"
			else
				ParentIDstr = " and ParentId In (select ClassID From fS_MS_productsClass Where ID = " & f_Id & ")"
			end if
		end if
		f_SQL = "select ClassID,OrderID,ClassCName,ParentID,ReycleTf from fS_MS_productsClass where ReycleTf=0 "& ParentIDstr &" Order by OrderId desc,id desc"
		set rs = Conn.execute(f_SQL)
		ClassNavistr = ""
		if rs.eof then
			ClassNavistr = ""
			rs.close:set rs=nothing
		else
			do while not rs.eof
				if div_tf=1 then
					ClassNavistr = ClassNavistr &classproducts_middle1 & "<a href="""&get_ClassLink(rs("ClassId"))&""">"&rs("ClassCName")&"</a>" & classproducts_middle2
				else
					if cols="0" then
						cols_str = "&nbsp;"
					else
						cols_str = "<br />"
					end if
					ClassNavistr = ClassNavistr & titleNavi & "<a href="""&get_ClassLink(rs("ClassId"))&""""& cssstyle&">"&rs("ClassCName")&"</a>"& cols_str &""
				end if
				rs.movenext
			loop
			rs.close:set rs=nothing
			if div_tf=1 then
				ClassNavistr=classproducts_head & ClassNavistr & classproducts_bottom
			end if
		end if
		f_Lable.DictLableContent.Add "1",ClassNavistr
	End function
	'专题导航
	Public function SpecialNavi(f_Lable,f_type,f_Id)
		dim cols,titlecss,div_tf,cssstyle,titleNavi,rs,SpecialNavistr,classproducts_head,classproducts_bottom,classproducts_middle1,classproducts_middle2,cols_str
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
		set rs = Conn.execute("select SpecialID,SpecialCName,SpecialEName,isLock from fS_MS_Special where isLock=0  Order by SpecialID desc")
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
		if div_tf=1 then SpecialNavistr=classproducts_head & SpecialNavistr & classproducts_bottom
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
	'专题调用
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
		set rs=Conn.execute("select SpecialCName,SpecialEName,naviText,naviPic from fS_MS_Special Where SpecialEName='"& ClassID &"'")
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
			rs.close
			set rs=nothing
		end if
		f_Lable.DictLableContent.Add "1",SpecialCodestr
	End function
	'栏目调用
	Public function ClassCode(f_Lable,f_type,f_Id)
		dim ClassID,titleNavi,ClassCodestr,cols,pictf,picsize,piccss,piccssstr,ContentTf,ContentNumber,div_tf,titlecss,cssstyle,ContentCSS,ContentCSSstr,classproducts_head,classproducts_bottom,classproducts_middle1,classproducts_middle2
		ClassID = f_Lable.LablePara("栏目")
		titleNavi = f_Lable.LablePara("导航")
		titlecss = f_Lable.LablePara("栏目名称CSS")
		ContentCSS = f_Lable.LablePara("导航内容CSS")
		cols=f_Lable.LablePara("排列方式")
		pictf=f_Lable.LablePara("图片显示")
		picsize=f_Lable.LablePara("图片尺寸")
		piccss=f_Lable.LablePara("图片CSS")
		ContentTf = f_Lable.LablePara("导航内容")
		ContentNumber= f_Lable.LablePara("导航内容字数")
		if piccss<>"" then
			piccssstr = " class="""& piccss &""""
		else
			piccssstr = ""
		end if
		if trim(ClassID)="" then
			ClassCodestr = "错误的标签,by foosun.cn"
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

		ClassCodestr = ""
		dim rs
		set rs=Conn.execute("select ClassID,ClassCName,NaviContent,NaviPic from fS_MS_productsClass Where ClassID='"& ClassID &"'")
		if rs.eof then
			ClassCodestr = ""
			rs.close:set rs=nothing
		else'
			if div_tf=1 then
				if pictf="1" then
					if trim(picsize)="" then
						ClassCodestr = ClassCodestr & ""
					else
						ClassCodestr = ClassCodestr & "  <a href="""&get_ClassLink(rs("ClassID"))&"""><img "&piccssstr&" src="""&rs("NaviPic")&""" width="""&split(picsize,",")(1)&""" height="""&split(picsize,",")(0)&""" border=""0"" /></a>" &chr(10)
					end if
				end if
				ClassCodestr = ClassCodestr & classproducts_middle1 & "<a href="""&get_ClassLink(rs("ClassID"))&""">"& rs("ClassCName") &"</a>" & classproducts_middle2
				if ContentTf="1" then
					ClassCodestr = ClassCodestr & classproducts_middle1 &"&nbsp;"& GetCStrLen(""&rs("NaviContent"),ContentNumber) & "&nbsp;<a href="""& get_ClassLink(rs("ClassID")) &""">详细</a>" & classproducts_middle2
				end if
				ClassCodestr = classproducts_head & ClassCodestr & classproducts_bottom
			else
				ClassCodestr = ClassCodestr& "<table width=""99%"" border=""0"" cellspacing=""0"" cellpadding=""5"">"&chr(10)&" <tr>"
				if pictf="1" then
					if trim(picsize)="" then
						ClassCodestr = ClassCodestr & ""
					else
						if cols="0" then
							ClassCodestr = ClassCodestr & "<td align=""center""><a href="""&get_ClassLink(rs("ClassID"))&"""><img "&piccssstr&" src="""&rs("NaviPic")&""" width="""&split(picsize,",")(1)&""" height="""&split(picsize,",")(0)&""" border=""0"" /></a></td>" &chr(10)
						else
							ClassCodestr = ClassCodestr & "<td><a href="""&get_ClassLink(rs("ClassID"))&"""><img "&piccssstr&" src="""&rs("NaviPic")&""" width="""&split(picsize,",")(1)&""" height="""&split(picsize,",")(0)&""" border=""0"" /></a></td></tr>" &chr(10)
						end if
					end if
				end if
				if cols="0" then
					ClassCodestr = ClassCodestr & "<td>"& titleNavi &"<a href="""&get_ClassLink(rs("ClassID"))&""""&cssstyle&">"& rs("ClassCName") &"</a>"
					if ContentTf="1" then
						ClassCodestr = ClassCodestr & "<br /><div align=""left"" "&ContentCSSstr&">&nbsp;"& GetCStrLen(""&rs("NaviContent"),ContentNumber)&"&nbsp;<a href="""& get_ClassLink(rs("ClassID")) &""">详细</a></div></td>"&chr(10)
					else
						ClassCodestr = ClassCodestr & "</td></tr>"&chr(10)
					end if
				else
					ClassCodestr = ClassCodestr & " <tr><td>"& titleNavi &"<a href="""&get_ClassLink(rs("ClassID"))&""""&cssstyle&">"& rs("ClassCName") &"</a>"&chr(10)
					if ContentTf="1" then
						ClassCodestr = ClassCodestr & "<br /><div align=""left"" "&ContentCSSstr&">&nbsp;"& GetCStrLen(""&rs("NaviContent"),ContentNumber)&"&nbsp;<a href="""& get_ClassLink(rs("ClassID")) &""">详细</a></div></td></tr>"&chr(10)
					else
						ClassCodestr = ClassCodestr & "</td></tr>"&chr(10)
					end if
				end if
				ClassCodestr = ClassCodestr & "</table>"
			end if
			rs.close:set rs=nothing
		end if
		f_Lable.DictLableContent.Add "1",ClassCodestr
	End function
	'替换样式列表
	Public function getlist_products(f_obj,s_Content,f_titlenumber,f_contentnumber,f_navinumber,f_picshowtf,f_datestyle,f_openstyle,f_Mf_Domain,f_subsys_ListType)
		Dim f_target,get_SpecialID,ListSql,Rs_ListObj,s_productsPathUrl,Rs_Authorobj,k_i,k_tmp_Char,k_tmp_uchar,k_tmp_Chararray,formReview,LinkType
		Dim s_m_Rs,s_array,s_t_i,tmp_list,s_f_classSql,m_Rs_class,class_path,str_ProductTitle
		select case f_subsys_ListType
			case "classproducts"
				get_SpecialID = f_obj("SpecialID")
			case "specialproducts"
				get_SpecialID = ""
			case else
		end select
		f_Mf_Domain=Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		LinkType=1
		if f_openstyle=0 then
			f_target=" "
		else
			f_target=" target=""_blank"""
		end if
		if instr(s_Content,"{MS:FS_ID}")>0 then
			s_Content = replace(s_Content,"{MS:FS_ID}",f_obj("Id"))
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_ProductTitle}")>0 then
			Dim TitleStyle
			TitleStyle=split(f_obj("TitleStyle"),",")
			if f_titlenumber=0 then
				str_ProductTitle=f_obj("ProductTitle")
			Else
				str_ProductTitle = GotTopic(f_obj("ProductTitle"),f_titlenumber)
			End if
			if TitleStyle(0)<>"0" then
				str_ProductTitle = "<font color='"&TitleStyle(0)&"'>"&str_ProductTitle&"</font>"
			end if
			if TitleStyle(1)="1" then
				str_ProductTitle = "<strong>"&str_ProductTitle&"</strong>"
			end if
			if TitleStyle(2)="1"  then
				str_ProductTitle = "<em>"&str_ProductTitle&"</em>"
			end if
			if TitleStyle(3)="1"  then
				str_ProductTitle = "<font style=""text-decoration:underline"">"&str_ProductTitle&"</font>"
			end if
			if f_picshowtf=1 then
				if f_obj("BigPic")<>"" or  f_obj("SmallPic")<>""  then
					str_ProductTitle = ""&str_ProductTitle
				end if
			end if
			s_Content = replace(s_Content,"{MS:FS_ProductTitle}",""&str_ProductTitle)
		end if
		'----------------------------------2/12---by chen---商品完整名称------------------------------------
		if instr(s_Content,"{MS:FS_ProductTitleAll}")>0 then
			Dim TitleStyle1
			TitleStyle1=split(f_obj("TitleStyle"),",")
			if f_titlenumber=0 then
				str_ProductTitle=f_obj("ProductTitle")
			Else
				str_ProductTitle =Replace(Replace(Replace(Replace(Lose_Html(f_obj("ProductTitle"))," ","&nbsp;"),Chr(34),"&quot;"),">","&gt;"),"<","&lt;")
			End if
			if TitleStyle1(0)<>"0" then
				str_ProductTitle = "<font color='"&TitleStyle1(0)&"'>"&str_ProductTitle&"</font>"
			end if
			if TitleStyle1(1)="1" then
				str_ProductTitle = "<strong>"&str_ProductTitle&"</strong>"
			end if
			if TitleStyle1(2)="1"  then
				str_ProductTitle = "<em>"&str_ProductTitle&"</em>"
			end if
			if TitleStyle1(3)="1"  then
				str_ProductTitle = "<font style=""text-decoration:underline"">"&str_ProductTitle&"</font>"
			end if
			if f_picshowtf=1 then
				if f_obj("BigPic")<>"" or  f_obj("SmallPic")<>""  then
					str_ProductTitle = ""&str_ProductTitle&"<img src="""&m_PathDir&"sys_images/img.gif"" alt=""图片"" border=""0"">"
				end if
			end if
			s_Content = replace(s_Content,"{MS:FS_ProductTitleAll}",""&str_ProductTitle)
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_Barcode}")>0 then
			s_Content = replace(s_Content,"{MS:FS_Barcode}",""&f_obj("Barcode"))
		end if
		if instr(s_Content,"{MS:FS_Serialnumber}")>0 then
			s_Content = replace(s_Content,"{MS:FS_Serialnumber}",""&f_obj("Serialnumber"))
		end if
		'_________________________________________________________________________________________________

		dim products_SavePath,s_Query_rs,products_Domain,products_UrlDomain,products_ClassEname,s_all_savepath
		if  instr(s_Content,"{MS:FS_ProductURL}")>0 then
			s_all_savepath = get_productsLink(f_obj("Id"))
			s_productsPathUrl = s_all_savepath
			s_Content = replace(s_Content,"{MS:FS_ProductURL}",""&s_productsPathUrl)
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_Stockpile}")>0 then
			s_Content = replace(s_Content,"{MS:FS_Stockpile}",""&f_obj("Stockpile")&"")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_OldPrice}")>0 then
			s_Content = replace(s_Content,"{MS:FS_OldPrice}",""&f_obj("OldPrice"))
		end if
		'---------------------------------------------------------
		if instr(s_Content,"{MS:FS_NewPrice}")>0 then
			s_Content = replace(s_Content,"{MS:FS_NewPrice}",""&f_obj("NewPrice"))
		end if
		'------------------------------------------------运输费用
		If Instr(s_Content,"{MS:FS_Mail_money}")>0 Then
			s_Content = replace(s_Content,"{MS:FS_Mail_money}",""&f_obj("Mail_Money"))
		End if
		'------------------------------------------------计算出含运费的价格
		If Instr(s_Content,"{MS:FS_NowMoney}")>0 Then
			s_Content = replace(s_Content,"{MS:FS_NowMoney}",""&int(f_obj("Mail_Money"))+int(f_obj("NewPrice")))
		End if
		'----------------------------------------------------商品系列化

		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_ProductContent}")>0 then
			If f_subsys_ListType <> "readproducts" Then
				s_Content = replace(s_Content,"{MS:FS_ProductContent}",replace(replace(GotTopic(""&f_obj("ProductContent")&"",f_contentnumber),"&nbsp;",""),vbCrLf,"")&"<a href="""& s_productsPathUrl &"""></a>")
			Else
				s_Content = replace(s_Content,"{MS:FS_ProductContent}",f_obj("ProductContent"))
			End If
		end if
		'--------------------------------------------------------------------------------------------------
		if instr(s_Content,"{MS:FS_RepairContent}")>0 then
			s_Content = replace(s_Content,"{MS:FS_RepairContent}",""&f_obj("RepairContent"))
		end if
		'--------------------------------------------------------------------------------------------------
		if instr(s_Content,"{MS:FS_RepairContent}")>0 then
			s_Content = replace(s_Content,"{MS:FS_RepairContent}",""&f_obj("RepairContent"))
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_AddTime}")>0 then
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
			s_Content = replace(s_Content,"{MS:FS_AddTime}",""&tmp_f_datestyle&"")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_AddMember}")>0 then
			set Rs_Authorobj = Conn.execute("select Admin_Name from fS_Mf_Admin where Admin_Name='"& f_obj("AddMember") &"'")
			if not Rs_Authorobj.eof then
				s_Content = replace(s_Content,"{MS:FS_AddMember}",""&f_obj("AddMember"))
				Rs_Authorobj.close:set Rs_Authorobj=nothing
			Else
				s_Content = replace(s_Content,"{MS:FS_AddMember}","")
				Rs_Authorobj.close:set Rs_Authorobj=nothing
			end if
		end if
		'--------------------------------------------------------------------------------------------------
		if instr(s_Content,"{MS:FS_ProductAddress}")>0 then
			s_Content = replace(s_Content,"{MS:FS_ProductAddress}",""&f_obj("ProductsAddress"))
		end if
		'--------------------------------------------------------------------------------------------------
		if instr(s_Content,"{MS:FS_MakeFactory}")>0 then
			s_Content = replace(s_Content,"{MS:FS_MakeFactory}",""&f_obj("MakeFactory"))
		end if
		'--------------------------------------------------------------------------------------------------
		if instr(s_Content,"{MS:FS_MakeTime}")>0 then
			s_Content = replace(s_Content,"{MS:FS_MakeTime}",""&f_obj("MakeTime"))
		end if
		'--------------------------------------------------------------------------------------------------
		if instr(s_Content,"{MS:FS_saleNumber}")>0 then
			s_Content = replace(s_Content,"{MS:FS_saleNumber}",""&f_obj("saleNumber"))
		end if
		'--------------------------------------------------------------------------------------------------
		if instr(s_Content,"{MS:FS_SaleStyle}")>0 then
			Dim saleStyle
			select case f_obj("SaleStyle")
				case "1" saleStyle="竞价"
				case "2" saleStyle="一口价"
				case "3" saleStyle="降价"
				case "4" saleStyle="特价"
				case else saleStyle="普通"
			end select
			s_Content = replace(s_Content,"{MS:FS_SaleStyle}",saleStyle)
		end if
		'--------------------------------------------------------------------------------------------------
		if instr(s_Content,"{MS:FS_DiscountStartDate}")>0 then
			s_Content = replace(s_Content,"{MS:FS_DiscountStartDate}",""&f_obj("DiscountStartDate"))
		end if
		'--------------------------------------------------------------------------------------------------
		if instr(s_Content,"{MS:FS_DiscountEndDate}")>0 then
			s_Content = replace(s_Content,"{MS:FS_DiscountEndDate}",""&f_obj("DiscountEndDate"))
		end if
		'--------------------------------------------------------------------------------------------------
		if instr(s_Content,"{MS:FS_Discount}")>0 then
			s_Content = replace(s_Content,"{MS:FS_Discount}",f_obj("Discount"))
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_hits}")>0 then
			dim hits_str
			if f_subsys_ListType="readproducts" then
				hits_str = "<span id=""MS_id_click_"&f_obj("id")&"""></span><script language=""JavaScript"" src=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Click_ajax.asp?type=js&SubSys=MS&spanid=MS_id_click_"&f_obj("id")&"""></script>"
			else
				hits_str = "" & f_obj("click")
			end if
			s_Content = replace(s_Content,"{MS:FS_hits}",hits_str)
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_KeyWords}")>0 then
			if trim(f_obj("Keyword"))<>"" and Not isNull(trim(f_obj("Keyword"))) then
				k_tmp_Chararray = split(f_obj("Keyword"),",")
				k_tmp_Char=""
				k_tmp_uchar= ""
				for k_i = 0 to UBound(k_tmp_Chararray)
					if k_i=UBound(k_tmp_Chararray) then
						k_tmp_Char = k_tmp_Char & "<a href=""http://"&Request.Cookies("foosunMfCookies")("foosunMfDomain")&"/search.html?keyword="& k_tmp_Chararray(k_i) &"&Type=MS"" target=""_blank"">"& k_tmp_Chararray(k_i) &"</a>"
						k_tmp_uchar= k_tmp_uchar &  k_tmp_Chararray(k_i)
					else
						k_tmp_Char = k_tmp_Char & "<a href=""http://"&Request.Cookies("foosunMfCookies")("foosunMfDomain")&"/search.html?keyword="& k_tmp_Chararray(k_i) &"&Type=MS"" target=""_blank"">"& k_tmp_Chararray(k_i) &"</a>&nbsp;&nbsp;"
						k_tmp_uchar= k_tmp_uchar &  k_tmp_Chararray(k_i) &","
					end if
				next
				s_Content = replace(s_Content,"{MS:FS_KeyWords}",""&k_tmp_Char&"")
				s_Content = replace(s_Content,"{MS:FS_TitleKeyWords}",""&k_tmp_uchar&"")
			else
				s_Content = replace(s_Content,"{MS:FS_KeyWords}","")
				s_Content = replace(s_Content,"{MS:FS_TitleKeyWords}","")
			end if
		end if

		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_SmallPicPath}")>0 then
			if trim(f_obj("smallPic"))<>"" then
				s_Content = replace(s_Content,"{MS:FS_SmallPicPath}",""&f_obj("smallPic"))
			else
				s_Content = replace(s_Content,"{MS:FS_SmallPicPath}","")
			end if
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_PicPath}")>0 then
			if trim(f_obj("BigPic"))<>"" then
				s_Content = replace(s_Content,"{MS:FS_PicPath}",""&f_obj("BigPic"))
			else
				s_Content = replace(s_Content,"{MS:FS_PicPath}","no picture!")
			end if
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_ShopBagURL}")>0 then
			s_Content = replace(s_Content,"{MS:FS_ShopBagURL}",m_Path_UserDir&"addToBuyBag.asp?pid="&f_obj("id")&"&type=ms")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_FormReview}")>0 then
			if f_obj("isShowReview")=1 then
				s_Content = replace(s_Content,"{MS:FS_FormReview}","<span id=""Review_TF_"& f_obj("ID") &""">loading...</span><script language=""JavaScript"" type=""text/javascript"" src=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/ReviewTF.asp?Id="& f_obj("ID") &"&Type=MS""></script>")
			Else
				s_Content = replace(s_Content,"{MS:FS_FormReview}","")
			End If
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_ReviewTF}")>0 then
			if f_obj("isShowReview")=1 then
				s_Content = replace(s_Content,"{MS:FS_ReviewTF}","<a href=""http://"&Request.Cookies("foosunMfCookies")("foosunMfDomain")&"/ReviewUrl.asp?Type=ms&Id="&f_obj("Id")&""">评论</a>")
			Else
				s_Content = replace(s_Content,"{MS:FS_ReviewTF}","")
			end if
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_ShowComment}")>0 then
			'---2007-01-05 by ken
			Dim ShowCommentStr
			ShowCommentStr = "<label id=""show_review_"& f_obj("ID") &""">评论显示，使用ajax调用</label>"&chr(10)
			ShowCommentStr = ShowCommentStr & "<script language=""JavaScript"" src=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/ShowReview.asp?type=MS&Id=" & f_obj("ID") &"&SpanId=show_review_"& f_obj("ID") &"""></script>"&chr(10)
			if f_obj("isShowReview")=1 then
				s_Content = replace(s_Content,"{MS:FS_ShowComment}",ShowCommentStr)
			Else
				s_Content = replace(s_Content,"{MS:FS_ShowComment}","")
			End If
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_AddFavorite}")>0 then
			s_Content = replace(s_Content,"{MS:FS_AddFavorite}","http://"& Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/User/AddFavor.asp?Id="&f_obj("ID")&"&Type=ms")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_SendFriend}")>0 then
			s_Content = replace(s_Content,"{MS:FS_SendFriend}","http://"& Request.Cookies("FoosunMFCookies")("FoosunMFDomain") &"/"&G_USER_DIR&"/Sendmail.asp?Id="&f_obj("ID")&"&Type=ms")
		end if
		'_________________________________________________________________________________________________
		if instr(s_Content,"{MS:FS_SpecialList}")>0 then
			if trim(get_SpecialID)<>"" then
				tmp_list = ""
				s_array = split(get_SpecialID,",")
				for s_t_i = 0 to ubound(s_array)
							set s_m_Rs=Conn.execute("select SpecialID,SpecialCName,SpecialEName from fS_MS_Special where SpecialID="&trim(s_array(s_t_i))&" order by SpecialID desc")
							if not s_m_Rs.eof then
								tmp_list = tmp_list &"<a href="""& s_m_Rs("SpecialEName") &""">" & s_m_Rs("SpecialCName") &"</a>&nbsp;"
							else
								tmp_list = tmp_list
							end if
							s_m_Rs.close:set s_m_Rs=nothing
				next
				tmp_list = tmp_list
			end if
			s_Content = replace(s_Content,"{MS:FS_SpecialList}",tmp_list)
		end if
		'获得栏目地址_______________________________________________________________________________________________
		s_f_classSql = "select ClassID,ClassCName,ClassEName,[Domain],NaviContent,NaviPic,SavePath,fileSaveType,Keywords,Description,fileExtName from fS_MS_productsClass where ClassId='"&f_obj("ClassId")&"' and ReycleTf=0 order by OrderID desc,id desc"
		if instr(s_Content,"{MS:FS_ClassURL}")>0 then
			dim Class_UrlDomain,class_domain,fileext
			set m_Rs_class = Conn.execute(s_f_classSql)
			class_domain = m_Rs_class("Domain")
			if LinkType=1 then
				if  trim(class_domain)<>"" then
					Class_UrlDomain = "http://" & class_domain
				else
					Class_UrlDomain = "http://" & f_Mf_Domain
				end if
			else
				if  trim(class_domain)<>"" then
					Class_UrlDomain = "http://" & class_domain
				else
					Class_UrlDomain = ""
				end if
			end if
			if m_Rs_class("fileSaveType")=0 then
				fileext = m_Rs_class("ClassEName")&"/index."&m_Rs_class("fileExtName")
			elseif m_Rs_class("fileSaveType")=1 then
				fileext = m_Rs_class("ClassEName")&"/"& m_Rs_class("ClassEName") &"."&m_Rs_class("fileExtName")
			else
				fileext = m_Rs_class("ClassEName") &"."&m_Rs_class("fileExtName")
			end if
			if Class_UrlDomain<>"" then
				class_path=Class_UrlDomain&replace(m_Rs_class("SavePath"),"//","/") &"/"&fileext
			else
				class_path=Class_UrlDomain&replace(m_PathDir & m_Rs_class("SavePath"),"//","/") &"/"&fileext
			end if
			s_Content = replace(s_Content,"{MS:FS_ClassURL}",class_path)
			m_Rs_class.close:set m_Rs_class=nothing
		end if
		if instr(s_Content,"{MS:FS_ClassName}")>0 then
			set m_Rs_class = Conn.execute(s_f_classSql)
			s_Content = replace(s_Content,"{MS:FS_ClassName}",""&m_Rs_class("ClassCName")&"")
			m_Rs_class.close:set m_Rs_class=nothing
		end if
		if instr(s_Content,"{MS:FS_NaviPicURL}")>0 then
			set m_Rs_class = Conn.execute(s_f_classSql)
			if trim(m_Rs_class("NaviPic"))<>"" then
				s_Content = replace(s_Content,"{MS:FS_NaviPicURL}",""& m_Rs_class("NaviPic") &"")
			else
				s_Content = replace(s_Content,"{MS:FS_NaviPicURL}","")
			end if
			m_Rs_class.close:set m_Rs_class=nothing
		end if
		if instr(s_Content,"{MS:FS_NaviContent}")>0 then
			set m_Rs_class = Conn.execute(s_f_classSql)
			s_Content = replace(s_Content,"{MS:FS_NaviContent}",""&m_Rs_class("NaviContent"))
			m_Rs_class.close:set m_Rs_class=nothing
		end if
		if instr(s_Content,"{MS:FS_ClassKeywords}")>0 then
			set m_Rs_class = Conn.execute(s_f_classSql)
			s_Content = replace(s_Content,"{MS:FS_ClassKeywords}",""&m_Rs_class("Keywords"))
			m_Rs_class.close:set m_Rs_class=nothing
		end if
			'-----------------2/5 by chen--- 栏目导航说明添加--------------------------
		if instr(s_Content,"{MS:FS_ClassNaviContent}")>0 then
			set m_Rs_class = Conn.execute(s_f_classSql)
			s_Content = replace(s_Content,"{MS:FS_ClassNaviContent}",""&m_Rs_class("NaviContent"))
			m_Rs_class.close:set m_Rs_class=nothing
		end if
		'-----------------2/5 by chen--- 栏目导航说明添加-------------------------------
		'-----------------2/5 by chen--- 栏目导航图片添加--------------------------------
		if instr(s_Content,"{MS:FS_ClassNaviPicURL}")>0 then
			set m_Rs_class = Conn.execute(s_f_classSql)
			s_Content = replace(s_Content,"{MS:FS_ClassNaviPicURL}",""&m_Rs_class("NaviPic"))
			m_Rs_class.close:set m_Rs_class=nothing
		end if
		'-----------------2/5 by chen--- 栏目导航图片添加-------------------------------
		if instr(s_Content,"{MS:FS_Classdescription}")>0 then
			set m_Rs_class = Conn.execute(s_f_classSql)
			s_Content = replace(s_Content,"{MS:FS_Classdescription}",""&m_Rs_class("Description"))
			m_Rs_class.close:set m_Rs_class=nothing
		end if
		'这里暂时用*代替
		'以下为通用专题使用##############################################################################
		Dim m_Rs_special,m_sp_sql,array_special,i_special,s_SpecialName,m_save_special,special_UrlDomain
		if instr(s_Content,"{MS:FS_SpecialName}")>0 Then
			if trim(get_SpecialID)<>"" then
				s_SpecialName = ""
				array_special = split(get_SpecialID,",")

				for i_special = 0 to Ubound(array_special)
					m_sp_sql = "select SpecialID,SpecialCName,SpecialEName,naviText,SavePath,FileExtName,isLock,naviPic from fS_MS_Special where isLock=0 and SpecialID="&trim(array_special(i_special))
					set m_Rs_special=Conn.execute(m_sp_sql)
					if not m_Rs_special.eof then
						if LinkType=1 then
							special_UrlDomain = "http://" & f_Mf_Domain
						else
							special_UrlDomain = ""
						end if
						m_save_special = special_UrlDomain&replace(m_PathDir & m_Rs_special("SavePath")&"/special_"&m_Rs_special("SpecialEName")&"."&m_Rs_special("FileExtName"),"//","/")
						if i_special=ubound(array_special) then
							s_SpecialName = s_SpecialName & "<a href="""&m_save_special&""">" &m_Rs_special("SpecialCName")&"</a>"
						else
							s_SpecialName = s_SpecialName & "<a href="""&m_save_special&""">" &m_Rs_special("SpecialCName")&"</a>&nbsp;"
						end if
					else
						s_SpecialName = ""
					end if
				next
				s_Content = replace(s_Content,"{MS:FS_SpecialName}",s_SpecialName)
			else
				s_Content = replace(s_Content,"{MS:FS_SpecialName}","")
			end if
		end if
		s_Content = replace(s_Content,"{MS:FS_SpecialURL}","")
		s_Content = replace(s_Content,"{MS:FS_SpecialNaviPicURL}","")
		s_Content = replace(s_Content,"{MS:FS_SpecialNaviDescript}","")
		'以下是自定义字段替换
		if instr(s_Content,"{MS=Define|")>0 then
			dim define_rs_sql,define_rs
			define_rs_sql="select ID,TableEName,ColumnName,ColumnValue,InfoID,InfoType from fS_Mf_DefineData where InfoType='MS' and InfoID='"&f_obj("ID")&"' order by ID desc"
			set define_rs=Conn.execute(define_rs_sql)
			if not define_rs.eof  then
				do while not define_rs.eof
					s_Content = replace(s_Content,"{MS=Define|"&define_rs("TableEName")&"}",""&define_rs("ColumnValue"))
					define_rs.movenext
				loop
				define_rs.close:set define_rs=nothing
			else
				dim define_class_sql,define_class_rs
				define_class_sql="select D_Coul from fS_Mf_DefineTable where D_SubType='MS' order  by DefineID desc"
				set define_class_rs=Conn.execute(define_class_sql)
				if not define_class_rs.eof then
					do while not define_class_rs.eof
						s_Content = replace(s_Content,"{MS=Define|"&define_class_rs("D_Coul")&"}","")
						define_class_rs.movenext
					loop
				end if
				define_class_rs.close:set define_class_rs=nothing
				define_rs.close:set define_rs=nothing
			end if
		end if
		getlist_products = s_Content
	End function
	'得到商品单个地址
	Public function get_productsLink(f_id)
		dim rs,config_rs,config_mf_rs,class_rs
		dim SaveproductsPath,fileName,FileExtName,ClassId,IsDomain,LinkType,Mf_Domain,Url_Domain,ClassEName,c_Domain,c_SavePath
		set rs = Conn.execute("select ID,ClassId,SavePath,fileName,fileExtName from fS_MS_products where Id="&f_id)
		If Not Rs.Eof Then
			SaveproductsPath = rs("SavePath")
			fileName = rs("fileName")
			fileExtName = rs("fileExtName")
			ClassId = rs("ClassId")
		Else
			get_productsLink = ""
			Exit function
		End If
		Rs.Close : Set Rs = Nothing
		set config_rs = Conn.execute("select top 1 IsDomain from fS_MS_SysPara")
		IsDomain = config_rs("IsDomain")
		'---ken
		If Trim(IsDomain) <> "" And Not IsNull(IsDomain) Then
			LinkType = True
		Else
			LinkType = False
		End If
		config_rs.close:set config_rs=nothing
		Mf_Domain = Request.Cookies("foosunMfCookies")("foosunMfDomain")
		set class_rs = Conn.execute("select ClassEName,IsURL,URLAddress,[Domain],SavePath from fS_MS_productsClass where ClassId='"&ClassId&"'")
		if not class_rs.eof then
			ClassEName = class_rs("ClassEName")
			c_Domain = class_rs("Domain")
			c_SavePath = class_rs("SavePath")
		else
			ClassEName = ""
		end if
		class_rs.close:set class_rs=nothing
		'---ken
		If Trim(c_Domain) <> "" And Not IsNull(c_Domain) Then
			get_productsLink = "http://" & c_Domain & Replace("/" &SaveproductsPath &"/"&fileName&"."&fileExtName,"//","/")
		Else
			If LinkType then
				Url_Domain = "http://"&IsDomain
			Else
				if G_VIRTUAL_ROOT_DIR<>"" then
					Url_Domain = "/"& G_VIRTUAL_ROOT_DIR
				else
					Url_Domain = ""
				end if
			End If
			If Url_Domain <> "" Then
				get_productsLink = Url_Domain & Replace(c_SavePath&"/" & ClassEName &SaveproductsPath &"/"&fileName&"."&fileExtName,"//","/")
			Else
				get_productsLink = Url_Domain & Replace(c_SavePath& "/" & ClassEName &SaveproductsPath &"/"&fileName&"."&fileExtName,"//","/")
			End If
		End If
	get_productsLink = get_productsLink
	End function
	'得到商品栏目地址
	Public function get_ClassLink(f_id)
		dim config_rs,IsDomain,LinkType,Mf_Domain,c_rs,ClassEName,c_Domain,Url_Domain,ClassSaveType,class_savepath,fileExtName,c_SavePath
		set config_rs = Conn.execute("select top 1 IsDomain from fS_MS_SysPara")
		IsDomain = config_rs("IsDomain")
		LinkType ="1"
		config_rs.close:set config_rs=nothing
		Mf_Domain = Request.Cookies("foosunMfCookies")("foosunMfDomain")
		set c_rs=conn.execute("select id,IsURL,UrlAddress,ClassId,ClassEName,[Domain],fileSaveType,fileExtName,SavePath from fS_MS_productsClass where ClassId='"&f_id&"'")
		if not c_rs.eof then
			if c_rs("IsURL")=0 then
				ClassEName = c_rs("ClassEName")
				c_Domain = c_rs("Domain")
				fileExtName = c_rs("fileExtName")
				ClassSaveType = c_rs("fileSaveType")
				c_SavePath= c_rs("SavePath")
				if ClassSaveType=0 then
					class_savepath = ClassEName&"/index."&fileExtName
				elseif ClassSaveType=1 then
					class_savepath = ClassEName&"/"& ClassEName &"."&fileExtName
				else
					class_savepath = ClassEName &"."&fileExtName
				end if
				if LinkType = 1 then
					if trim(c_Domain)<>"" then
						Url_Domain = "http://"&c_Domain
					else
						if trim(IsDomain)<>"" then
							Url_Domain = "http://"&IsDomain
						else
							Url_Domain = "http://"&Mf_Domain
						end if
					end if
				else
					if trim(c_Domain)<>"" then
						Url_Domain = "http://"&c_Domain
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
				end if
				get_ClassLink = Url_Domain&Replace(c_SavePath&"/"&class_savepath,"//","/")
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
	'得到商品专题地址
	Public function get_specialLink(f_id)
		dim config_rs,IsDomain,LinkType,Mf_Domain,c_rs,SpecialEName,ExtName,c_SavePath,Url_Domain,Ms_Sp_DomainTF,Ms_Sp_Domain,Ms_Sp_SavePath
		set config_rs = Conn.execute("select top 1 IsDomain from fS_MS_SysPara")
		IsDomain = config_rs("IsDomain")
		'---ken
		If IsDomain <> "" And Not IsNull(IsDomain) Then
			Ms_Sp_DomainTF = True
		Else
			Ms_Sp_DomainTF = False
		End If
		config_rs.close:set config_rs=nothing
		Mf_Domain = Request.Cookies("foosunMfCookies")("foosunMfDomain")
		set c_rs=conn.execute("select SpecialID,SpecialCName,SpecialEName,[Domain],SavePath,FileExtName,isLock from fS_MS_Special where SpecialEName='"&f_id&"'")
		if not c_rs.eof then
			SpecialEName = c_rs("SpecialEName")
			ExtName = c_rs("fileExtName")
			c_SavePath= c_rs("SavePath")
			'---ken
			Ms_Sp_Domain = c_rs("Domain")
			If Trim(Ms_Sp_Domain) <> "" And Not IsNull(Ms_Sp_Domain) Then
				get_specialLink = "Http://" & Ms_Sp_Domain & "/special_"& SpecialEName & "." & ExtName
			Else
				If Ms_Sp_DomainTF Then
					Url_Domain = "http://"&IsDomain
				Else
					if G_VIRTUAL_ROOT_DIR<>"" then
						Url_Domain = "/"& G_VIRTUAL_ROOT_DIR
					else
						Url_Domain = ""
					end if
				End If
				If Url_Domain <> "" Then
					get_specialLink = Url_Domain & Replace("/special_" & SpecialEName & "." & ExtName,"//","/")
				Else
					get_specialLink = Replace("/" & c_SavePath & "/special_" & SpecialEName & "." & ExtName,"//","/")
				End If
			End If
			c_rs.close:set c_rs=nothing
		else
			c_rs.close:set c_rs=nothing
			get_specialLink = ""
		end if
		get_specialLink = get_specialLink
	End function
	'------------------商品栏目调用部分----chenzhaohui 1/29---------
	Public Function ClassInfo(f_Lable,f_type,f_Id)
		Dim Content_List,InfoType,ClassIID,str_connect,str_sql
		Dim f_rs_obj,f_sql
		if f_Id <> "" then
			InfoType = f_Lable.LablePara("调用内容")
			Set str_connect = Server.CreateObject(G_fS_RS)
			If Len(f_Id) = 15 And Not IsNumeric(f_Id) Then
				str_sql = "Select ClassID from FS_MS_productsClass where ClassID='"&f_id&"'"
			Else
				str_sql = "Select ClassID from FS_MS_productsClass where ID="&f_id
			End If
			str_connect.open str_sql,Conn,0,1
			set ClassIID=str_connect("ClassID")
			if not str_connect.eof then
				if ClassIID<>"" then
					Select Case InfoType
						Case "ClassCName"
							f_sql="select ClassCName From FS_MS_ProductsClass where ClassID='"& ClassIID &"'"
						Case "Keywords"
							f_sql="select Keywords From FS_MS_ProductsClass where ClassID='"& ClassIID &"'"
						Case "Description"
							f_sql="select Description From FS_MS_ProductsClass where ClassID='"& ClassIID &"'"
					End Select
					set f_rs_obj = Conn.execute(f_sql)
					if Not f_rs_obj.eof then
						Content_List = f_rs_obj(0)
					else
						Content_List = ""
					end if
					str_connect.close:set str_connect=nothing
					f_rs_obj.close:set f_rs_obj=nothing
				else
					Content_List = ""
				End if
			end if
		else
			Content_List = ""
		end if
		f_Lable.DictLableContent.Add "1",Content_List
	End Function
	'------------------商品栏目调用部分结束---chenzhaohui 1/29--------
	'得到子类
	Public function get_ClassList(TypeID,CompatStr,f_css)
		Dim ChildTypeListRs,ChildTypeListStr,TempStr,i,Sql
		Set ChildTypeListRs = Server.CreateObject(G_fS_RS)
		Sql = "Select ParentID,ClassID,ClassCName from fS_MS_productsClass where ParentID='" & TypeID & "' and ReycleTf=0 order by OrderID desc,id desc"
		ChildTypeListRs.open sql,Conn,0,1
		TempStr = CompatStr
		do while Not ChildTypeListRs.Eof
			get_ClassList = get_ClassList & TempStr
			get_ClassList = get_ClassList & "<img src="""&m_PathDir&"sys_images/-.gif"" border=""0""><a href="""&get_ClassLink(ChildTypeListRs("ClassId"))&""" class="""& f_css &""">"&ChildTypeListRs("ClassCName")&"</a>"&chr(10)
			get_ClassList = get_ClassList & get_ClassList(ChildTypeListRs("ClassID"),TempStr,f_css)
			ChildTypeListRs.MoveNext
		loop
		ChildTypeListRs.Close:Set ChildTypeListRs = Nothing
	End function
	'得到option子类
	Public Function get_optionProductsList(TypeID,CompatStr)
		Dim ChildTypeListRs,ChildTypeListStr,TempStr
		Set ChildTypeListRs = Conn.execute("Select ParentID,ClassID,ClassCName from FS_MS_ProductsClass where ParentID='" & TypeID & "' and ReycleTF=0 order by OrderID desc,id desc")
		TempStr = CompatStr & "┉"
		do while Not ChildTypeListRs.Eof
			get_optionProductsList = get_optionProductsList &"<option value="""&ChildTypeListRs("ClassId")&""">"& TempStr
			get_optionProductsList = get_optionProductsList & "┉"&ChildTypeListRs("ClassCName")&"</option>"&chr(10)
			get_optionProductsList = get_optionProductsList & get_optionProductsList(ChildTypeListRs("ClassID"),TempStr)
			ChildTypeListRs.MoveNext
		loop
		ChildTypeListRs.Close:Set ChildTypeListRs = Nothing
	End Function
	Public Function getSubClass(typeID)
		Dim childClassRs,result
		Set childClassRs=Conn.execute("Select ParentID,ClassID from FS_MS_ProductsClass where ParentID='" & TypeID & "' and ReycleTF=0 order by OrderID desc,id desc")
		While Not childClassRs.eof
			result=result&"'"&childClassRs("classID")&"',"&getSubClass(childClassRs("classID"))
			childClassRs.movenext
		Wend
		Set childClassRs=nothing
		getSubClass=result
	End function
end class
%>
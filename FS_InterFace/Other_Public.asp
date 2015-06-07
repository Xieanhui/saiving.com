<%
Class cls_Other
	'定义参数
	Private m_Rs,m_FSO,m_Dict
	Private m_PathDir,m_Path_UserDir,m_Path_User,m_Path_adminDir,m_Path_UserPageDir,m_Path_Templet
	Private m_Err_Info,m_Err_NO


	Private Sub Class_initialize()
		Set m_Rs = Server.CreateObject(G_FS_RS)
		Set m_FSO = Server.CreateObject(G_FS_FSO)
		Set m_Dict = Server.CreateObject(G_FS_DICT)
		m_PathDir = replace("/"&G_VIRTUAL_ROOT_DIR&"/","//","/")
		m_Path_UserDir = replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_USER_DIR&"/","//","/")
		m_Path_UserPageDir = replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_USERFILES_DIR&"/","//","/")
		m_Path_Templet  = replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_TEMPLETS_DIR&"/","//","/")
	End Sub

	Private Sub Class_Terminate()
		Set m_Rs = Nothing
		Set m_FSO = Nothing
		Set m_Dict = Nothing
	End Sub
	'获得参数
	Public Function get_LableChar(f_Lable,f_ID,f_Type)
		select case LCase(f_Lable.LableFun)
			case "sstype"
				get_LableChar=SSTYPE(f_Lable,"sstype")
			case "picfl"
				get_LableChar=PicFL(f_Lable,"picfl")
			case "wordfl"
				get_LableChar=WordFL(f_Lable,"wordfl")
			case "bookcode"
				get_LableChar=BookCode(f_Lable,"bookcode")
			case "vslist"
				get_LableChar=VSLIST(f_Lable,"vslist")
			case "adlist"
				get_LableChar=AdLIST(f_Lable,"adlist")
		end select
	End Function
	'统计调用
	Public Function SSTYPE(f_Lable,f_type)
		dim CodeStyle,PathStyle,Url_Path,code_end
		CodeStyle = f_Lable.LablePara("方式")
		PathStyle = f_Lable.LablePara("路径")
		If PathStyle = "0" then
			Url_Path = Replace("/"&G_VIRTUAL_ROOT_DIR&"/","//","/")
		else
			Url_Path = "http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain") &"/"
		end if
		if CodeStyle="0" then
			code_end = "1"
		elseif CodeStyle="1" then
			code_end = "0"
		else
			code_end = "2"
		end if
		SSTYPE = "<script language=""JavaScript"" src="""& Url_Path&"stat/index.asp?code="&CodeStyle&""" type=""text/JavaScript""></script>"
		f_Lable.DictLableContent.Add "1",SSTYPE
	End Function
	'调用图片友情连接
	Public Function PicFL(f_Lable,f_type)
		Dim div_tf,div_chars,div_class,i_link,CodeNumber,siteName,colsNumber,pizsize
		Dim ClassSql,rs,Select_class_pic,openstyle,f_openstyle,f_DivClass
		CodeNumber = f_Lable.LablePara("调用数量")
		siteName = f_Lable.LablePara("站点名称")
		colsNumber = f_Lable.LablePara("每行数量")
		pizsize = f_Lable.LablePara("图片尺寸")
		Select_class_pic= f_Lable.LablePara("连接类别")
		openstyle=f_Lable.LablePara("打开窗口")
		f_DivClass = f_Lable.LablePara("Divclass")
		if f_DivClass<>"" then
			f_DivClass=" class="""&f_DivClass&""""
		end if
		if f_Lable.LablePara("输出格式") = "out_DIV" then
			div_tf=1
		else
			div_tf=0
		end If
		if openstyle=0 then
			f_openstyle=""
		else
			f_openstyle=" target=""_blank"""
		end if
		If Select_class_pic = "" Or NOt Isnumeric(Select_class_pic) Then
			ClassSql = ""
		Else
			Select_class_pic = Cint(Select_class_pic)
			 ClassSql = " And ClassID = " & Select_class_pic
		End If	 	
		If CodeNumber="" Then CodeNumber=0
		set rs = Conn.execute("select top "&CodeNumber&" id,ClassID,F_Name,F_PicUrl,F_Url From FS_FL_FrendList where F_Lock=0 and F_Type=0" & ClassSql & " order by F_OrderID desc,ID desc")'显示指定类别的连接或显示所有类别下的连接
		PicFL = ""
		if rs.eof then
			PicFL = ""
			rs.close
			set rs=nothing
		else
			i_link = 0
			do while not rs.eof
				if div_tf=1 then
					PicFL = PicFL&vbNewLine &"<div"&f_DivClass&"><a href="""& rs("F_Url")&""""&f_openstyle&"><img src="""&rs("F_PicUrl")&""" alt="""& rs("F_Name")&""" border=""0"" width="""&split(pizsize,",")(0)&""" height="""&split(pizsize,",")(1)&"""/></a></div>"
				else
					PicFL = PicFL & "<td width="""&cint(100/colsNumber)&"%"" align=""center"" valign=""middle""><a href="""& rs("F_Url")&""""&f_openstyle&"><img border=""0"" src="""&rs("F_PicUrl")&""" alt="""& rs("F_Name")&""" width="""&split(pizsize,",")(0)&""" height="""&split(pizsize,",")(1)&"""/></a></td>"&vbNewLine
				end if
				rs.movenext
				i_link = i_link + 1
				if div_tf = 0 and i_link mod colsNumber = 0 Then
					PicFL = PicFL & "</tr><tr>"
				End If
			loop
			rs.close
			set rs=nothing
		end if
		if PicFL <> "" then
			if div_tf = 0 then
				PicFL = "<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""3"">"&vbNewLine&"<tr>"&vbNewLine&" "&PicFL&vbNewLine&"</tr>"&vbNewLine&"</table>"
			end if
		end if
		f_Lable.DictLableContent.Add "1",PicFL
	End Function
	'文字友情连接
	Public Function WordFL(f_Lable,f_type)
		dim div_tf,div_class,i_link
		dim CodeNumber,TitleCSS,colsNumber,pizsize,rs,lefttitle,Select_class_word,ClassSql,openstyle,f_openstyle
		CodeNumber = f_Lable.LablePara("调用数量")
		colsNumber = f_Lable.LablePara("每行数量")
		lefttitle = f_Lable.LablePara("站点显示字数")
		Select_class_word=f_Lable.LablePara("连接类别")
		openstyle=f_Lable.LablePara("打开窗口")
		div_class = f_Lable.LablePara("Divclass")
		If div_class<>"" Then
			div_class = " class="""&div_class&""""
		End if
		If IsNull(lefttitle) Then lefttitle=""
		if f_Lable.LablePara("输出格式") = "out_DIV" then
			div_tf=1
		else
			div_tf=0
			TitleCSS = f_Lable.LablePara("标题CSS")
			If TitleCSS="" then
				TitleCSS = ""
			else
				TitleCSS = " class="""&TitleCSS&""""
			end if
		end If
		if openstyle=0 then
			f_openstyle=""
		else
			f_openstyle=" target=""_blank"""
		end if
		if Select_class_word="" Or NOt Isnumeric(Select_class_word) then
			ClassSql=""
		Else
			Select_class_word = Cint(Select_class_word)
			ClassSql = " And ClassID = " & Select_class_word
		End If	 
		If CodeNumber="" Then CodeNumber=0
		set rs = Conn.execute("select top "&CodeNumber&" ClassID,F_Name,F_Url From FS_FL_FrendList where F_Lock=0 and F_Type=1 "&ClassSql&" order by F_OrderID desc,ID desc")'显示指定类别的连接或显示没有类别时的所有连接
		WordFL = ""
		if rs.eof then
			WordFL = ""
			rs.close
			set rs=nothing
		else
			i_link = 1
			do while not rs.eof
				dim tmpTitle
				If IsNumeric(lefttitle) then
					tmpTitle=Left(rs("F_Name"),lefttitle)
				Else
					tmpTitle=rs("F_Name")
				End if
				if div_tf=1 Then
					WordFL = WordFL&vbNewLine &"<div"&div_class&"><a href="""& rs("F_Url")&""""&f_openstyle&">"&tmpTitle&"</a></div>"
				Else
					if cint(colsNumber)=1 then
						WordFL = WordFL & "  <tr><td width="""&cint(100/colsNumber)&"%"" align=""center"" valign=""meddle""><a href="""& rs("F_Url")&""""&TitleCSS&" "&f_openstyle&">"& tmpTitle &"</a></td></tr>"&vbNewLine
					else
						if i_link mod colsNumber = 1 then WordFL= WordFL & "<tr>"
						WordFL = WordFL & "  <td width="""&cint(100/colsNumber)&"%"" align=""center"" valign=""meddle""><a href="""& rs("F_Url")&""""&TitleCSS&" "&f_openstyle&">"& tmpTitle &"</a></td>"&vbNewLine
						if i_link mod colsNumber = 0 then WordFL= WordFL & "</tr>"&vbNewLine
					end if
				end if
				rs.movenext
				i_link = i_link + 1
			Loop
			rs.close
			Set rs=Nothing
		End If
		if WordFL <> "" then
			if div_tf = 0 then
				WordFL = "<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""3"">"&vbNewLine&WordFL&"</table>"
			end if
		end if
		f_Lable.DictLableContent.Add "1",WordFL
	End Function
	'留言调用
	Public Function BookCode(f_Lable,f_type)
		dim BookPop,ClassID,TitleNumber,ColsNumber,leftTitle,div_tf,TitleCSS,DateCSS,datestyle
		dim classNews_head,classNews_middle1,classNews_bottom,classNews_middle2,user_str,tmp_hit
		dim rs,f_sql,inSql_CMD,orderby,rec_tf
		BookPop = f_Lable.LablePara("类型")
		ClassID = f_Lable.LablePara("栏目")
		TitleNumber = f_Lable.LablePara("数量")
		ColsNumber = f_Lable.LablePara("列数")
		leftTitle = f_Lable.LablePara("字数")
		if leftTitle<>"" then
			leftTitle = cint(leftTitle)
		else
			leftTitle = 30
		end if
		DateStyle = f_Lable.LablePara("日期格式")
		if f_Lable.LablePara("输入格式") = "out_DIV" then
			div_tf=1
		else
			div_tf=0
			TitleCSS=f_Lable.LablePara("标题CSS")
			if TitleCSS="" then
				TitleCSS = ""
			else
				TitleCSS = " class="""&TitleCSS&""""
			end if
			DateCSS=f_Lable.LablePara("日期CSS")
			if DateCSS<>"" then
				DateCSS = " class="""&DateCSS&""""
			else
				DateCSS = ""
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
		if ClassID<>"" then
			inSql_CMD = " and ClassID='"& ClassId &"'"
		else
			inSql_CMD = ""
		end if
		if BookPop="0" then
			orderby = " order by LastUpdateDate desc,id desc"
			rec_tf=""
		elseif BookPop="1" then
			orderby = " order by LastUpdateDate desc,id desc"
			rec_tf=" and IsTop='1'"
		else
			orderby = " order by hit desc,LastUpdateDate desc,id desc"
			rec_tf=""
		end if
		f_sql = "select top "& TitleNumber &" ID,ClassId,Topic,Body,AddDate,Hit,State,[user] From FS_WS_BBS where State='0' and ParentID='0' and IsAdmin='0' "& inSql_CMD & rec_tf & orderby &""
		set rs = Conn.execute(f_sql)
		BookCode = ""
		if rs.eof then
			BookCode = "<a href=""/guestbook/index.asp"">更多留言</a>"
			rs.close:set rs=nothing
		else
			do while not rs.eof
				if BookPop="2" then
					tmp_hit = "<font style=""font-size:10px;font-family:Arial;"" color=red>["&rs("Hit")&"]</font>"
				end if
				Dim ClassName_GB
				ClassName_GB = Conn.ExeCute("Select ClassName From FS_WS_Class Where ClassID = '" & rs("ClassId") & "'")(0)
				if div_tf = 1 then
					if rs("user")<>"游客" and rs("user")<>"过客" then
						user_str = "(<a href=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"&G_USER_DIR&"/ShowUser.asp?UserName="&rs("User")&""" target=""_blank"">"& rs("User")&"</a>)"
					else
						user_str = "("&rs("User")&")"
					end if
					BookCode = BookCode & classNews_middle1 & "<a href=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/GuestBook/ShowNote.asp?NoteID="&rs("id")&"&ClassName=" & ClassName_GB & "&ClassID="&rs("ClassId")&""" target=""_blank"">"&GotTopic(""& rs("Topic"),leftTitle) &"</a>"& tmp_hit & user_str & Get_Date(rs("AddDate"),DateStyle) & "" & classNews_middle2
				else
					if rs("user")<>"游客" and rs("user")<>"过客" then
						user_str = "(<a href=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"&G_USER_DIR&"/ShowUser.asp?UserName="&rs("User")&""" target=""_blank"">"& rs("User")&"</a>)"
					else
						user_str = "("&rs("User")&")"
					end if
					BookCode = BookCode & "<img src=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/sys_images/book_tf.gif"" border=""0"" /><a href=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/GuestBook/ShowNote.asp?NoteID="&rs("id")&"&ClassName=" & ClassName_GB & "&ClassID="&rs("ClassId")&""" target=""_blank"""&TitleCSS&">"&GotTopic(""&rs("Topic"),leftTitle)&"</a>"& tmp_hit & user_str &"<span"& DateCSS &">"& Get_Date(rs("AddDate"),DateStyle)&"</span>"&"<br>"
				end if
				rs.movenext
			loop
			rs.close
			set rs=nothing
		end if
		if div_tf=1 then BookCode = classNews_head & BookCode & classNews_bottom
		f_Lable.DictLableContent.Add "1",BookCode
	End Function
	'投票
	Public Function VSLIST(f_Lable,f_type)
		Dim Tid,SpanId,PicWidth
		Tid = f_Lable.LablePara("投票项")
		SpanId = f_Lable.LablePara("SpanId")
		PicWidth = f_Lable.LablePara("图片宽度")
		VSLIST = VSLIST & "<span id="""&SpanId&""">加载投票中...</span>"
		VSLIST = VSLIST & "<script src=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Vote/VoteJs.asp?TID="&Tid&"&InfoID="&SpanId&"&PicW="& PicWidth &""" language=""javascript""></script>"&vbNewLine
		f_Lable.DictLableContent.Add "1",VSLIST
	End Function
	'广告调用
	Public Function AdLIST(f_Lable,f_type)
		Dim AdId,SpanId
		AdId = f_Lable.LablePara("广告ID")
		SpanId = f_Lable.LablePara("SpanId")
		AdLIST = AdLIST & "<script src=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/ads/"&AdId&".js"" language=""javascript""></script>"&vbNewLine
		f_Lable.DictLableContent.Add "1",AdLIST
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
			table_str_list_head =  table_&vbNewLine
			table_str_list_head = table_str_list_head &" "& tr_&vbNewLine
		else
			table_="<table border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"">"
			table_str_list_head =  table_&vbNewLine
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
			table_str_list_middle_2 =  td__&vbNewLine
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
			table_str_list_bottom = " "&tr__&vbNewLine
		else
			table__="</table>"
			table_str_list_bottom = ""
		end if
		table_str_list_bottom = table_str_list_bottom &table__&vbNewLine
	End Function
	'DIV格式输出结束_____________________________________________________________________
End Class
%>
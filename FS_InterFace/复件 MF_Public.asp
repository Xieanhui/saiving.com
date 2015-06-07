<%
Class cls_MF
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
	Public Function get_LableChar(f_Lable,f_Id,f_position)
		select case LCase(f_Lable.LableFun)
			Case "postionnavi"
				get_LableChar = PostionNavi(f_Lable,"postionnavi",f_id,f_position)
			Case "pagetitle"
				get_LableChar = PageTitle(f_Lable,"pagetitle",f_id,f_position)
			Case "sitemap"
				get_LableChar = SiteMap(f_Lable,"sitemap",f_id,f_position)
			Case "search"
				get_LableChar = Search(f_Lable,"search",f_id,f_position)
			Case "infostat"
				get_LableChar = InfoStat(f_Lable,"infostat",f_id,f_position)
			Case "userlogin"
				get_LableChar = UserLogin(f_Lable,"userlogin",f_id,f_position)
			Case "copyright"
				get_LableChar = CopyRight(f_Lable,"copyright",f_id,f_position)
			Case "sublist"
				get_LableChar = SubList(f_Lable,"sublist",f_id,f_position)
			Case "freelabel"
				get_LableChar = FreeLabel(f_Lable,"freelabel",f_id,f_position)
			Case "customform"
				get_LableChar = CustomForm(f_Lable,"customform",f_id,f_position)
		end select
	End Function

	Public Sub OuptLablePara(f_LableName)
		Dim TestTestArray,ii
		TestTestArray = Split(f_LableName,"┆")
		for ii = LBound(TestTestArray) To UBound(TestTestArray)
			Response.Write(ii & " " & TestTestArray(ii) & Chr(13) & Chr(10))
		Next
		Response.End
	End Sub
	Public Function CustomForm(f_Lable,f_type,f_id,f_position)
		Dim f_JSSrc,f_ParaValue,FormStyleID
		f_JSSrc = m_PathDir & "customform/CustomFormJS.asp?"
		f_ParaValue = f_Lable.LablePara("调用表单")
		if f_ParaValue <> "" then f_JSSrc = f_JSSrc & "CustomFormID=" & f_ParaValue & "&"
		FormStyleID = f_Lable.LablePara("表单样式")
		if FormStyleID <> "" then f_JSSrc = f_JSSrc & "FormStyleID=" & FormStyleID & "&"
		f_ParaValue = f_Lable.LablePara("数据样式")
		if f_ParaValue <> "" then f_JSSrc = f_JSSrc & "DataStyleID=" & f_ParaValue
		if FormStyleID <> "" then
			f_ParaValue = f_Lable.LablePara("文本框CSS")
			if f_ParaValue <> "" then f_JSSrc = f_JSSrc & "&TextCSS=" & f_ParaValue & "&"
			f_ParaValue = f_Lable.LablePara("下拉框CSS")
			if f_ParaValue <> "" then f_JSSrc = f_JSSrc & "SelectCSS=" & f_ParaValue & "&"
			f_ParaValue = f_Lable.LablePara("其它对象CSS")
			if f_ParaValue <> "" then f_JSSrc = f_JSSrc & "OtherCSS=" & f_ParaValue
		end if
		f_JSSrc = "<script language=""javascript"" type=""text/javascript"" src=""" & f_JSSrc & """></script>"
		f_Lable.DictLableContent.Add "1",f_JSSrc
	End Function
	'位置导航
	Public Function PostionNavi(f_Lable,f_type,f_id,f_position)
		if trim(f_position)="" then
			PostionNavi = ""
		else
			dim fg_str,char_str,positCss,linkCss,mf_domain,mf_sitename,mf_filename,ns_dir,ns_domain,ns_sitename,ns_path,mf_linkpath,MF_OpenMode
			dim ms_domain,ms_dir,ms_path,ds_domain,ds_dir,ds_path,ap_path
			fg_str = f_Lable.LablePara("分割字符")
			char_str = f_Lable.LablePara("位置文字")
			positCss = f_Lable.LablePara("位置文字css")
			MF_OpenMode = f_Lable.LablePara("弹出窗口")
			MF_OpenMode = GetOpenModeStr(MF_OpenMode)
			if positCss<>"" then:positCss = " class="""& positCss &"""":else:positCss = "":end if
			if char_str<>"" then
				char_str="<span"&positCss&">"&char_str&"</span>"
			end if
			linkCss = f_Lable.LablePara("联接CSS")
			if linkCss<>"" then:linkCss = " class="""& linkCss &"""":else:linkCss = "":end if
			if Request.Cookies("FoosunMFCookies")("FoosunMFDomain") = "" or Request.Cookies("FoosunMFCookies")=empty then MFConfig_Cookies:NSConfig_Cookies:DSConfig_Cookies
			mf_domain = request.Cookies("FoosunMFCookies")("FoosunMFDomain")
			mf_sitename = request.Cookies("FoosunMFCookies")("FoosunMFsiteName")
			mf_filename = request.Cookies("FoosunMFCookies")("FoosunMFIndexFileName")
			ns_dir = request.Cookies("FoosunNSCookies")("FoosunNSNewsDir")
			ns_domain = request.Cookies("FoosunNSCookies")("FoosunNSDomain")
			ns_sitename = request.Cookies("FoosunNSCookies")("FoosunNSSiteName")
			if trim(ns_domain)<>"" then
				ns_path = fg_str & "<a href=""http://"& ns_domain &""""&linkCss&MF_OpenMode&">"&ns_sitename&"</a>"
			else
				ns_path = fg_str & "<a href=""http://"&mf_domain&"/"&ns_dir&"/"&Request.Cookies("FoosunNSCookies")("FoosunNSIndexPage")&""""&linkCss&MF_OpenMode&">"&ns_sitename&"</a>"
			end if
			mf_linkpath = "http://"& mf_domain &"/"& mf_filename
			If IsExist_SubSys("MS") Then
				if request.Cookies("FoosunMSCookies")("FoosunMSDomain")="" or Request.Cookies("FoosunMSCookies")=empty then MSConfig_Cookies
				ms_domain = request.Cookies("FoosunMSCookies")("FoosunMSDomain")
				ms_dir = request.Cookies("FoosunMSCookies")("FoosunMSDir")
				Dim ms_filename,ms_indexpath
				ms_filename = request.Cookies("FoosunMSCookies")("FoosunMSIndexHtml")
				if ms_dir <> "" then
					ms_indexpath = "/" & ms_dir & "/index." & ms_filename
				else
					ms_indexpath = "/index." & ms_filename
				end if
				ms_path = fg_str & "<a href=""http://" & mf_domain & ms_indexpath & """" & linkCss & MF_OpenMode & ">商城首页</a>"
			End If
			if IsExist_SubSys("DS") then
				if request.Cookies("FoosunDSCookies")("FoosunDSDomain")="" or Request.Cookies("FoosunDSCookies")=empty then
					DSConfig_Cookies
				end if
				ds_domain = request.Cookies("FoosunDSCookies")("FoosunDSDomain")
				ds_dir = request.Cookies("FoosunDSCookies")("FoosunDSDownDir")
				if trim(ds_domain)<>"" then
					ds_path = fg_str & "<a href=""http://"& ds_domain &""""&linkCss&MF_OpenMode&">下载中心</a>"
				else
					ds_path = fg_str & "<a href=""http://"& mf_domain & "/" & ds_dir &""""&linkCss&MF_OpenMode&">下载中心</a>"
				end If
			End If
			'---------------------2/1 by chen---------------------------------------------------------------
			if IsExist_SubSys("AP") then
					ap_path = fg_str & "<a href=""http:"""&linkCss&MF_OpenMode&">人才首页</a>"
				else
					ap_path = fg_str & "<a href=""http:"& mf_domain &""&linkCss&MF_OpenMode&">人才首页</a>"
			End If
			'------------------------2/1 by chen------------------------------------------------------------
			if IsExist_SubSys("SD") then
				dim sd_rs,sd_path
				set  sd_rs =Conn.execute("select top 1 [Domain],SavePath,siteName From FS_SD_Config")
				if sd_rs.eof then
					sd_path = "<a href=""http://"& mf_domain & "/Index.asp"" "&linkcss&MF_OpenMode&">供求信息</a>"
					sd_rs.close:set sd_rs = nothing
				else
					If sd_rs("Domain")<>"" then
						sd_path = "<a href=""http://"& sd_rs("Domain")&""""&linkcss&MF_OpenMode&">"&sd_rs("siteName")&"</a>"
					else
						sd_path = "<a href=""http://"& mf_domain &"/"&sd_rs("SavePath")&"/Index.asp"""&linkcss&MF_OpenMode&">"&sd_rs("siteName")&"</a>"
					end if
					sd_rs.close:set sd_rs = nothing
				end if
			end If
			Select Case Lcase(f_position)
				Case "mf"
					PostionNavi =  "<span"&positCss&">首页</span>"
				Case "ns"
					PostionNavi = "<a href=""http://"& mf_domain & """"&linkCss&MF_OpenMode&">首页</a>" & ns_path
				Case "ns_news"
					PostionNavi = GetNewsLocationStr(mf_sitename,mf_linkpath,ns_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
				Case "ns_class"
					PostionNavi = GetClassLocationStr(mf_sitename,mf_linkpath,ns_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)&fg_str&"<span"&positCss&">列表</span>"
				Case "ns_special"
					PostionNavi = GetSpecialLocationStr(mf_sitename,mf_linkpath,ns_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)&fg_str&"<span"&positCss&">列表</span>"
				Case "ms"
					PostionNavi = "<a href=""http://"& mf_domain & """"&linkCss&MF_OpenMode&">首页</a>"& ms_path
				Case "ms_news"
					PostionNavi = GetmsNewsLocationStr(mf_sitename,mf_linkpath,ms_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
				Case "ms_class"
					PostionNavi = GetmsClassLocationStr(mf_sitename,mf_linkpath,ms_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode,1)&fg_str&"<span"&positCss&">列表</span>"
				Case "ms_special"
					PostionNavi = GetmsSpecialLocationStr(mf_sitename,mf_linkpath,ms_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)&fg_str&"<span"&positCss&">列表</span>"
				Case "ds"
					PostionNavi = "<a href=""http://"& mf_domain & """"&linkCss&MF_OpenMode&">首页</a>"& ds_path
				Case "ds_news"
					PostionNavi = GetdsNewsLocationStr(mf_sitename,mf_linkpath,ds_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
				Case "ds_class"
					PostionNavi = GetdsClassLocationStr(mf_sitename,mf_linkpath,ds_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)&fg_str&"<span"&positCss&">列表</span>"
				Case "me_photonews"
					PostionNavi = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>"& fg_str &" <span"& positCss &">相册</span>"&fg_str&"浏览"
				Case "me_photoclass"
					PostionNavi = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>"& fg_str &" <span"& positCss &">相册</span>"
				Case "me_news"
					PostionNavi = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>"& fg_str &" <span"& positCss &">日志/网摘</span>"
				Case "me_class"
					PostionNavi = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>"& fg_str &" <span"& positCss &">日志/网摘</span>"& fg_str &"<span"& positCss &">"&GetFriendNumber(f_id)&"的列表</span>"
				Case "ap"
				PostionNavi = "<a href=""http://"& mf_domain & """"&linkCss&MF_OpenMode&">首页</a>"& ap_path
				Case "ap_jobnews"  '个人求职浏览
					PostionNavi = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>"& fg_str &" <span"& positCss &">人才中心</span>"& fg_str &"<span"& positCss &">个人求职浏览</span>"
				Case "ap_hrnews" '招聘浏览
					PostionNavi = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>"& fg_str &" <span"& positCss &">人才中心</span>"& fg_str &"<span"& positCss &">招聘浏览</span>"
				Case "ap_jobclass" '个人求职列表
					PostionNavi = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>"& fg_str &" <span"& positCss &">人才中心</span>"& fg_str &"<span"& positCss &MF_OpenMode&">个人求职列表</span>"
				Case "ap_hrclass" '招聘列表
					PostionNavi = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>"& fg_str &" <span"& positCss &">人才中心</span>"& fg_str &"<span"& positCss &">招聘列表</span>"
				Case "me_"
					PostionNavi = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>"& fg_str &" <span"& positCss &">日志/网摘主页</span>"
				Case "hs"
					PostionNavi = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>"& fg_str &" <span"& positCss &">房产首页</span>"
				Case "hs_blclass"
					PostionNavi = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>"& fg_str &" <span"& positCss &">楼盘列表</span>"
				Case "hs_hrnews"
					PostionNavi = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>"& fg_str &" <span"& positCss &">房源浏览</span>"
				Case "hs_hlclass"
					PostionNavi = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>"& fg_str &" <span"& positCss &">房源列表</span>"
				Case "hs_brnews"
					PostionNavi = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>"& fg_str &" <span"& positCss &">楼盘浏览</span>"
				Case "sd"
					PostionNavi = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>"& fg_str & sd_path
				Case "sd_news"
					PostionNavi = GetSDPageLocationStr(mf_sitename,mf_linkpath,sd_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
				Case "sd_class"
					PostionNavi = GetSDClassLocationStr(mf_sitename,mf_linkpath,sd_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
				Case "sd_area"
					PostionNavi = GetSDAreaPageLocationStr(mf_sitename,mf_linkpath,sd_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
				Case "ms_products"
					f_id = Conn.execute("select ID From FS_MS_ProductsClass Where Classid='"&f_id&"'")(0)
					PostionNavi = GetmsClassLocationStr(mf_sitename,mf_linkpath,ms_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode,1)&"<span"&positCss&">列表</span>"
				Case "ds_down"
					PostionNavi = GetdsClassLocationStr(mf_sitename,mf_linkpath,ds_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)&fg_str&"<span"&positCss&">列表</span>"
			End Select
		end if
		f_Lable.DictLableContent.Add "1",PostionNavi
	End Function
	'标题位置
	Public Function PageTitle(f_Lable,f_type,f_id,f_position)
		'{FS:MF=PageTitle┆附属文字$风讯__Foosun.CN┆附属文字位置$1┆分割字符$__}
		'{FS:MF=PageTitle┆附属文字$风讯__Foosun.CN┆附属文字位置$0┆分割字符$__}
		dim char_conc,char_post,fg_char
		dim mf_sitename,ns_sitename,rs
		char_conc = f_Lable.LablePara("附属文字")
		char_post = f_Lable.LablePara("附属文字位置")
		if char_post = "1" then:char_post=1:else:char_post=0:end if
		fg_char = f_Lable.LablePara("分割字符")
		if request.Cookies("FoosunMFCookies")("FoosunMFDomain") = "" or Request.Cookies("FoosunMFCookies")=empty then
			MFConfig_Cookies:NSConfig_Cookies:DSConfig_Cookies
		end if
		mf_sitename = request.Cookies("FoosunMFCookies")("FoosunMFsiteName")
		ns_sitename = request.Cookies("FoosunNSCookies")("FoosunNSSiteName")
		Select Case Lcase(f_position)
			Case "mf","ms","ds"
				if char_post = 1 then
					PageTitle =  mf_sitename&fg_char&"首页"&fg_char&char_conc
				else
					PageTitle =  char_conc&fg_char&mf_sitename&fg_char&"首页"
				end if
			Case "ns_news"
				 dim ns_news
				 set rs = Conn.execute("select NewsTitle From FS_NS_News where NewsId='"& f_id &"'")
				 if rs.eof then
					ns_news = ns_sitename
					rs.close:set rs = nothing
				 else
					ns_news = rs(0)
					rs.close:set rs = nothing
				 end if
				if char_post = 1 then
					PageTitle = ns_news&fg_char&char_conc
				else
					PageTitle = char_conc&fg_char&ns_news
				end if
			Case "ns_class"
				dim ns_class
				set rs = Conn.execute("select ClassName From FS_NS_NewsClass Where ClassId='"&f_id&"'")
				 if rs.eof then
					ns_class = ns_sitename
					rs.close:set rs = nothing
				 else
					ns_class = rs(0)
					rs.close:set rs = nothing
				 end if
				if char_post = 1 then
					PageTitle = ns_class&fg_char&char_conc
				else
					PageTitle = char_conc&fg_char&ns_class
				end if
			Case "ns_special"
				dim ns_special
				if not isnumeric(f_id) then
						PageTitle = ns_sitename
				else
					set rs = Conn.execute("select SpecialCName From FS_NS_Special Where SpecialID="&f_id&"")
					 if rs.eof then
						ns_special = ns_sitename
						rs.close:set rs = nothing
					 else
						ns_special = rs(0)
						rs.close:set rs = nothing
					 end if
					if char_post = 1 then
						PageTitle = ns_special&fg_char&char_conc
					else
						PageTitle = char_conc&fg_char&ns_special
					end if
				end if
			Case "ms_news"
				dim ms_news
				if not isnumeric(f_id) then
						PageTitle = mf_sitename
				else
					set rs = Conn.execute("select ProductTitle From FS_MS_Products Where ID="&f_id&"")
					 if rs.eof then
						ms_news = mf_sitename
						rs.close:set rs = nothing
					 else
						ms_news = rs(0)
						rs.close:set rs = nothing
					 end if
					if char_post = 1 then
						PageTitle = ms_news&fg_char&char_conc
					else
						PageTitle = char_conc&fg_char&ms_news
					end if
				end if
			Case "ms_class"
				if IsExist_SubSys("MS") then
						dim ms_class
						if f_id = "" then
								PageTitle = mf_sitename
						else
							set rs = Conn.execute("select ClassCName From FS_MS_ProductsClass Where ClassID='"&f_id&"'")
							 if rs.eof then
								ms_class = mf_sitename
								rs.close:set rs = nothing
							 else
								ms_class = rs(0)
								rs.close:set rs = nothing
							 end if
							if char_post = 1 then
								PageTitle = ms_class&fg_char&char_conc
							else
								PageTitle = char_conc&fg_char&ms_class
							end if
						end if
				end if
			Case "ms_special"
				if IsExist_SubSys("MS") then
					dim ms_special
					if not isnumeric(f_id) then
							PageTitle = mf_sitename
					else
						set rs = Conn.execute("select SpecialCName From FS_MS_Special Where specialID="&f_id&"")
						 if rs.eof then
							ms_special = mf_sitename
							rs.close:set rs = nothing
						 else
							ms_special = rs(0)
							rs.close:set rs = nothing
						 end if
						if char_post = 1 then
							PageTitle = ms_special&fg_char&char_conc
						else
							PageTitle = char_conc&fg_char&ms_special
						end if
					end if
				end if
			Case "ds_news"
				dim ds_news
				if f_id = "" then
						PageTitle = mf_sitename
				else
					set rs = Conn.execute("select Name From FS_DS_List Where DownLoadID='"&f_id&"'")
					 if rs.eof then
						ds_news = mf_sitename
						rs.close:set rs = nothing
					 else
						ds_news = rs(0)
						rs.close:set rs = nothing
					 end if
					if char_post = 1 then
						PageTitle = ds_news&fg_char&char_conc
					else
						PageTitle = char_conc&fg_char&ds_news
					end if
				end if
			Case "ds_class"
				dim ds_class
				if f_id = "" then
						PageTitle = mf_sitename
				else
					set rs = Conn.execute("select ClassName From FS_DS_Class Where ClassID='"&f_id&"'")
					 if rs.eof then
						ds_class = mf_sitename
						rs.close:set rs = nothing
					 else
						ds_class = rs(0)
						rs.close:set rs = nothing
					 end if
					if char_post = 1 then
						PageTitle = ds_class&fg_char&char_conc
					else
						PageTitle = char_conc&fg_char&ds_class
					end if
				end if
			Case "me_photonews"
				dim me_photonews
				if not isnumeric(f_id) then
						PageTitle = mf_sitename
				else
					set rs = User_Conn.execute("select title From FS_ME_Photo Where id="&f_id&"")
					 if rs.eof then
						me_photonews = mf_sitename
						rs.close:set rs = nothing
					 else
						me_photonews = rs(0)
						rs.close:set rs = nothing
					 end if
					if char_post = 1 then
						PageTitle = me_photonews&fg_char&char_conc
					else
						PageTitle = char_conc&fg_char&me_photonews
					end if
				end if
			Case "me_photoclass"
				dim me_photoclass
				if not isnumeric(f_id) then
						PageTitle = mf_sitename
				else
					set rs = User_Conn.execute("select title From FS_ME_PhotoClass Where id="&f_id&"")
					 if rs.eof then
						me_photoclass = mf_sitename
						rs.close:set rs = nothing
					 else
						me_photoclass = rs(0)
						rs.close:set rs = nothing
					 end if
					if char_post = 1 then
						PageTitle = me_photoclass&fg_char&char_conc
					else
						PageTitle = char_conc&fg_char&me_photoclass
					end if
				end if
			Case "me_news"
				dim me_news
				if not isnumeric(f_id) then
						PageTitle = mf_sitename
				else
					set rs = User_Conn.execute("select Title From FS_ME_Infoilog Where iLogID="&f_id&"")
					 if rs.eof then
						me_news = mf_sitename
						rs.close:set rs = nothing
					 else
						me_news = rs(0)
						rs.close:set rs = nothing
					 end if
					if char_post = 1 then
						PageTitle = me_news&fg_char&char_conc
					else
						PageTitle = char_conc&fg_char&me_news
					end if
				end if
			Case "me_class"
				dim me_class
				if not isnumeric(f_id) then
						PageTitle = mf_sitename
				else
					set rs = Conn.execute("select ClassCName From FS_ME_InfoClass Where ClassID="&f_id&"")
					 if rs.eof then
						me_class = mf_sitename
						rs.close:set rs = nothing
					 else
						me_class = rs(0)
						rs.close:set rs = nothing
					 end if
					if char_post = 1 then
						PageTitle = me_class&fg_char&char_conc
					else
						PageTitle = char_conc&fg_char&me_class
					end if
				end if
			'----------------------------1/31 by chen---页面标题---------------------------------			
			Case "ap_jobnews" '个人求职浏览
				if IsExist_SubSys("AP") then
				dim ap_jobnews
				if not isnumeric(f_id) then
						PageTitle = mf_sitename
				else
					set rs = User_Conn.execute("select Uname From FS_AP_Resume_BaseInfo Where BID="&f_id&"")
					 if rs.eof then
						ap_jobnews = mf_sitename
						rs.close:set rs = nothing
					 else
						ap_jobnews = rs(0)
						rs.close:set rs = nothing
					 end if
					if char_post = 1 then
						PageTitle = ap_jobnews&fg_char&char_conc
					else
						PageTitle = char_conc&fg_char&ap_jobnews
					end if
				end if
				end if
			Case "ap_hrnews" '招聘浏览
				if IsExist_SubSys("AP") then
				dim ap_hrnews
				if not isnumeric(f_id) then
						PageTitle = mf_sitename
				else
					set rs = User_Conn.execute("select JobName From FS_AP_Job_Public Where PID="&f_id&"")
					 if rs.eof then
						ap_hrnews = mf_sitename
						rs.close:set rs = nothing
					 else
						ap_hrnews = rs(0)
						rs.close:set rs = nothing
					 end if
					if char_post = 1 then
						PageTitle = ap_hrnews&fg_char&char_conc
					else
						PageTitle = char_conc&fg_char&ap_hrnews
					end if
				end if
				end if
			Case "ap_jobclass" '个人求职列表
				if IsExist_SubSys("AP") then
				dim ap_jobclass
				if not isnumeric(f_id) then
						PageTitle = mf_sitename
				else
					set rs = Conn.execute("select UserNumber From FS_AP_Resume_Position Where BID="&f_id&"")
					 if rs.eof then
						ap_jobclass = mf_sitename
						rs.close:set rs = nothing
					 else
						ap_jobclass = rs(0)
						rs.close:set rs = nothing
					 end if
					if char_post = 1 then
						PageTitle = ap_jobclass&fg_char&char_conc
					else
						PageTitle = char_conc&fg_char&ap_jobclass
					end if
				end if
				end if
			Case "ap_hrclass" '招聘列表
				if IsExist_SubSys("AP") then
				dim ap_hrclass
				if not isnumeric(f_id) then
						PageTitle = mf_sitename
				else
					set rs = Conn.execute("select Job From FS_AP_Job Where JID="&f_id&"")
					 if rs.eof then
						ap_hrclass = mf_sitename
						rs.close:set rs = nothing
					 else
						ap_hrclass = rs(0)
						rs.close:set rs = nothing
					 end if
					if char_post = 1 then
						PageTitle = ap_hrclass&fg_char&char_conc
					else
						PageTitle = char_conc&fg_char&ap_hrclass
					end if
				end if
				end if
'----------------------------1/31 by chen-----------------------------------
			Case "me_"
					if char_post = 1 then
						PageTitle = "日志/网摘"&fg_char&char_conc
					else
						PageTitle = char_conc&fg_char&"日志/网摘"
					end if
			Case "hs_news"
				if IsExist_SubSys("HS") then
				
				end if
			Case "sd_news"
				if IsExist_SubSys("SD") then
					dim sd_news
					if not isnumeric(f_id) then
							PageTitle = mf_sitename
					else
						set rs = Conn.execute("select PubTitle From FS_SD_News Where ID="&f_id&"")
						 if rs.eof then
							sd_news = mf_sitename
							rs.close:set rs = nothing
						 else
							sd_news = rs(0)
							rs.close:set rs = nothing
						 end if
						if char_post = 1 then
							PageTitle = sd_news&fg_char&char_conc
						else
							PageTitle = char_conc&fg_char&sd_news
						end if
					end if
				end if
			Case "sd"
				if IsExist_SubSys("SD") then
					dim sd_index
					set rs = Conn.execute("select top 1 siteName From FS_SD_config")
					 if rs.eof then
						sd_index = mf_sitename
						rs.close:set rs = nothing
					 else
						sd_index = rs(0)
						rs.close:set rs = nothing
					 end if
					if char_post = 1 then
						PageTitle = sd_index&fg_char&char_conc
					else
						PageTitle = char_conc&fg_char&sd_index
					end if
				end if
			Case "sd_class"
				if IsExist_SubSys("SD") then
					dim sd_class
					if not isnumeric(f_id) then
							PageTitle = mf_sitename
					else
						set rs = Conn.execute("select GQ_ClassName From FS_SD_Class Where ID="&f_id&"")
						 if rs.eof then
							sd_class = mf_sitename
							rs.close:set rs = nothing
						 else
							sd_class = rs(0)
							rs.close:set rs = nothing
						 end if
						if char_post = 1 then
							PageTitle = sd_class&fg_char&char_conc
						else
							PageTitle = char_conc&fg_char&sd_class
						end if
					end if
			  End if
			Case "sd_area"
				if IsExist_SubSys("SD") then
					dim sd_area
					if not isnumeric(f_id) then
							PageTitle = mf_sitename
					else
						set rs = Conn.execute("select ClassName From FS_SD_address Where ID="&f_id&"")
						 if rs.eof then
							sd_area = mf_sitename
							rs.close:set rs = nothing
						 else
							sd_area = rs(0)
							rs.close:set rs = nothing
						 end if
						if char_post = 1 then
							PageTitle = sd_area&fg_char&char_conc
						else
							PageTitle = char_conc&fg_char&sd_area
						end if
					end if
			  End if
		End Select
		f_Lable.DictLableContent.Add "1",PageTitle
	End Function
	
	'搜索
	Public Function Search(f_Lable,f_type,f_id,f_position)
		f_Lable.IsSave = True
		dim DateTF,SearchType,datestr,searchetypestr,Searchstr,MF_Domain,subsyslist,subsys
		DateTF = f_Lable.LablePara("日期搜索")
		SearchType = f_Lable.LablePara("模糊搜索")
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		if DateTF = "1" then
			datestr = " 开始日期：<input name=""s_date"" class=""f-text"" type=""text"" value="""&date()-1&""" size=""10"" /> 结束日期：<input name=""e_date"" class=""f-text"" type=""text"" value="""&date()&""" size=""10"" />"
		else
			datestr = ""
		end if
		if SearchType = "1" then
			searchetypestr = " 模糊搜索：<input name=""SearchType""class=""f-checkbox"" type=""checkbox"" id=""SearchType"" value=""1"" /> "
		else
			searchetypestr = ""
		end if
		subsyslist = " <select class=""f-select"" name=""Subsys"" id=""Subsys"">"
		set subsys=Conn.execute("select Sub_Sys_ID,Sub_Sys_Name From FS_MF_Sub_Sys where Sub_Sys_Installed=1 and Sub_Sys_ID<>'ME' and  Sub_Sys_ID<>'CS' and  Sub_Sys_ID<>'SS' and  Sub_Sys_ID<>'VS' and  Sub_Sys_ID<>'AS' and  Sub_Sys_ID<>'FL' and  Sub_Sys_ID<>'AP' order by ID")
		if not subsys.eof then
			do while not subsys.eof
				if Request.Cookies("FoosunSUBCookie")("FoosunSUB"&subsys("Sub_Sys_ID")&"")="1" then
					subsyslist = subsyslist &"<option value="""& subsys("Sub_Sys_ID") &""">"& subsys("Sub_Sys_Name")&"</option>"
				end if
				subsys.movenext
			loop
			subsys.close:set subsys=nothing
		end if
		subsyslist = subsyslist &"</select>"
		Searchstr = Searchstr & "<form action=""http://"& MF_Domain & "/Search.html"" id=""SearchForm"" name=""SearchForm"" method=""get"" style=""margin:0px;"">关键字:<input name=""Keyword"" class=""f-text"" type=""text"" id=""Keyword"" size=""15"" /> "&datestr&subsyslist&searchetypestr&" <input name=""SearchSubmit_Foosun"" class=""f-button"" type=""submit"" id=""SearchSubmit"" value=""全站搜索"" /></form>"
		Search = Searchstr
		f_Lable.DictLableContent.Add "1",Searchstr
	End Function
	
	'统计信息
	Public Function InfoStat(f_Lable,f_type,f_id,f_position)
		f_Lable.IsSave = True
		dim  colstyle,BRchar,rs_Class,rs_News,rs_Newssp,RS_ME
		colstyle = f_Lable.LablePara("排列方式")
		if colstyle = 0 then
			BRchar = "┆"
		else
			BRchar = "<br />"
		end if
		set rs_Class = Conn.execute("select count(Id) From FS_NS_NewsClass where ReycleTF=0")
		set rs_News = Conn.execute("select count(Id) From FS_NS_News where isLock=0 and isdraft=0 and isRecyle=0")
		set rs_Newssp = Conn.execute("select count(SpecialID) From FS_NS_Special where isLock=0")
		set rs_ME = User_Conn.execute("select count(UserID) From FS_ME_Users where isLock=0")
		InfoStat =""
		InfoStat = InfoStat & "新闻栏目:"&rs_Class(0)&"个"&BRchar
		InfoStat = InfoStat & "新闻数量:"&rs_News(0)&"条"&BRchar
		InfoStat = InfoStat & "新闻专题:"&rs_Newssp(0)&"个"&BRchar
		if IsExist_SubSys("MS") Then
			dim rs_msclass,rs_ms,rs_mssp
			set rs_msclass = Conn.execute("select count(ID) From FS_MS_ProductsClass where ReycleTF=0")
			set rs_ms = Conn.execute("select count(ID) From FS_MS_Products where ReycleTF=0")
			set rs_mssp = Conn.execute("select count(specialID) From FS_MS_Special where islock=0")
			InfoStat = InfoStat & "商品栏目:"&rs_msclass(0)&"个"&BRchar
			InfoStat = InfoStat & "商品数量:"&rs_ms(0)&"个"&BRchar
			InfoStat = InfoStat & "商品专区:"&rs_mssp(0)&"个"&BRchar
			rs_msclass.close:set rs_msclass = nothing
			rs_ms.close:set rs_ms = nothing
			rs_mssp.close:set rs_mssp = nothing
		end if
		if IsExist_SubSys("DS") Then
			dim rs_dsclass,rs_ds
			set rs_dsclass = Conn.execute("select count(ID) From FS_DS_Class where ReycleTF=0")
			set rs_ds = Conn.execute("select count(ID) From FS_DS_List where AuditTF=1")
			InfoStat = InfoStat & "下载栏目:"&rs_dsclass(0)&"个"&BRchar
			InfoStat = InfoStat & "下载数量:"&rs_ds(0)&"个"&BRchar
			rs_dsclass.close:set rs_dsclass = nothing
			rs_ds.close:set rs_ds = nothing
		end if
		if IsExist_SubSys("SD") Then
			dim rs_sd
			set rs_sd = Conn.execute("select count(ID) From FS_SD_News where isLock=0 and isPass=1")
			InfoStat = InfoStat & "供求信息:"&rs_sd(0)&"个"&BRchar
			rs_sd.close:set rs_sd = nothing
		end if
		InfoStat = InfoStat & "注册会员:"&rs_ME(0)&"个"&BRchar
		rs_Class.close:set rs_Class = nothing
		rs_News.close:set rs_News = nothing
		rs_Newssp.close:set rs_Newssp = nothing
		rs_ME.close:set rs_ME = nothing
		f_Lable.DictLableContent.Add "1",InfoStat
	End Function
	
	'用户登陆
	Public Function UserLogin(f_Lable,f_type,f_id,f_position)
		f_Lable.IsSave = True
		Dim Str_Login_TF,Str_Login_DisType,Show_Str,InfoStr,DivBgStyle
		Dim StyleID,LoginBgCss,SeCss,MenuCss,TxtCss,SumBitCss,ResestCss,Regcss,GetpasCss
		Str_Login_TF = f_Lable.LablePara("标签方式")
		If Str_Login_TF = "" Or Not IsNumeric(Str_Login_TF) Then
			UserLogin = "发生意外错误，请重建登陆标签"
			Exit Function
		End If	
		If Cint(Str_Login_TF) = 0 Then
			Str_Login_DisType = f_Lable.LablePara("显示方式")
			Show_Str = "?DisTF=0&DisType=" & Str_Login_DisType & ""
			LoginBgCss = ""
		Else
			LoginBgCss = f_Lable.LablePara("标签背景")
			SeCss = f_Lable.LablePara("选择筐样式")
			MenuCss = f_Lable.LablePara("选择筐菜单样式")
			TxtCss = f_Lable.LablePara("文本筐样式")
			SumBitCss = f_Lable.LablePara("提交按钮样式")
			ResestCss = f_Lable.LablePara("取消按钮样式")
			Regcss = f_Lable.LablePara("注册连接样式")
			GetpasCss = f_Lable.LablePara("取回密码连接样式")
			InfoStr = f_Lable.FoosunStyle.StyleID & "┆" & SeCss & "┆" & MenuCss & "┆" & TxtCss & "┆" & SumBitCss & "┆" & ResestCss & "┆" & Regcss & "┆" & GetpasCss
			InfoStr = Server.URLEncode(InfoStr)
			Show_Str = "?DisTF=1&DisType=" & InfoStr & ""
		End If
		If LoginBgCss <> "" Then
			If Instr(LoginBgCss,"/") > 0 And Right(Lcase(LoginBgCss),4) = ".jpg" Or Right(Lcase(LoginBgCss),4) = ".gif" Or Right(Lcase(LoginBgCss),4) = ".png" Then
				DivBgStyle = " style=""background-image:url(" & LoginBgCss & ");"""
			Else
				DivBgStyle = " class=""" & LoginBgCss & """"
			End If	
		Else
			DivBgStyle = ""
		End If
		UserLogin = "<div" & DivBgStyle & " id=""FS400_User_Login"">登录模块加载中...</div>"&VbCrLf
		UserLogin = UserLogin & "<script language=""JavaScript"" src=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"& G_USER_DIR &"/m_UserLogin.asp" & Show_Str & "&spanid=FS400_User_Login""></script>"&chr(10)
		UserLogin = UserLogin
		f_Lable.DictLableContent.Add "1",UserLogin
	End Function
	
	'版权信息
	Public Function CopyRight(f_Lable,f_type,f_id,f_position)
		f_Lable.IsSave = True
		if request.Cookies("FoosunMFCookies")("FoosunMFCopyright") = "" then
			CopyRight = "风讯__Foosun.CN提醒，你没设置版权信息.请在参数设置中设置。"
		else
			CopyRight = request.Cookies("FoosunMFCookies")("FoosunMFCopyright")
		end if
		f_Lable.DictLableContent.Add "1",CopyRight
	End Function
	
	
	Public Function SubList(f_Lable,f_type,f_id,f_position)
		f_Lable.IsSave = True
		dim fg_char,linkcss,sublists,mf_domain,ns_dir,ns_domain
		fg_char = f_Lable.LablePara("分割符-可以使用html语法")
		linkcss = f_Lable.LablePara("CSS")
		if request.Cookies("FoosunMFCookies")("FoosunMFDomain") = "" or Request.Cookies("FoosunMFCookies")=empty then
			MFConfig_Cookies:NSConfig_Cookies:DSConfig_Cookies
		end if
		mf_domain = request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		ns_dir = request.Cookies("FoosunNSCookies")("FoosunNSNewsDir")
		ns_domain = request.Cookies("FoosunNSCookies")("FoosunNSDomain")
		SubList = "<a href=""http://"&mf_domain&""" class="""&linkcss&""">首页</a>"
		set sublists=Conn.execute("select Sub_Sys_ID,Sub_Sys_Name,Sub_Sys_Link From FS_MF_Sub_Sys where Sub_Sys_Installed=1 and Sub_Sys_ID<>'ME' and  Sub_Sys_ID<>'CS' and  Sub_Sys_ID<>'SS' order by ID")
		if sublists.eof then
			SubList = ""
			sublists.close:set sublists = nothing
		else
			do while not sublists.eof
				if trim(sublists("Sub_Sys_Link"))<>"" and sublists("Sub_Sys_Link")<>"http://" and not isnull(sublists("Sub_Sys_Link")) then
					SubList = SubList & fg_char & "<a href=""http://"&sublists("Sub_Sys_Link")&""" class="""&linkcss&""">"&sublists("Sub_Sys_Name")&"</a>"
				end if
			sublists.movenext
			loop
			sublists.close:set sublists = nothing
		end if
		f_Lable.DictLableContent.Add "1",SubList
	End Function
	'得到新闻当前路径
	Public Function GetNewsLocationStr(mf_sitename,mf_linkpath,ns_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
		Dim RsNewsObj
		Set RsNewsObj = Conn.Execute("Select ClassID from FS_NS_News where NewsID='" & f_id & "'")
		if Not RsNewsObj.Eof then
			GetNewsLocationStr = GetClassLocationStr(mf_sitename,mf_linkpath,ns_path,RsNewsObj("ClassID"),fg_str,char_str,positCss,linkCss,MF_OpenMode) & fg_str & char_str
		else
			GetNewsLocationStr = GetClassLocationStr(mf_sitename,mf_linkpath,ns_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode) & fg_str & char_str
		end if
		Set RsNewsObj = Nothing
	End Function
	'栏目当前位置
	Public Function GetClassLocationStr(mf_sitename,mf_linkpath,ns_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
		Dim SqlClass,RsClassObj
		if f_id = "" then Exit Function
		Set RsClassObj = Conn.Execute("Select ParentID,ClassID,ClassName from FS_NS_NewsClass where ClassID='" & f_id & "'")
		if Not RsClassObj.Eof then
			GetClassLocationStr = GetClassLocationStr & "<a  href=""" & get_ClassLink(RsClassObj("ClassID")) & """"&linkCss&MF_OpenMode&">" & RsClassObj("ClassName") & "</a>"
			do while RsClassObj("ParentID") <> "0"
				Set RsClassObj = Conn.Execute("Select ParentID,ClassID,ClassName from FS_NS_NewsClass where ClassID='" & RsClassObj("ParentID") & "'")
				if RsClassObj.Eof then Exit do
				GetClassLocationStr = "<a href="""& get_ClassLink(RsClassObj("ClassID")) &""""&linkCss&MF_OpenMode&">"& RsClassObj("ClassName") &"</a>"& fg_str & GetClassLocationStr
			loop
		end if
		GetClassLocationStr = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">" &"首页</a>" & ns_path & fg_str & GetClassLocationStr
		RsClassObj.Close
		Set RsClassObj = Nothing
	End Function
	'专题当前位置
	Function GetSpecialLocationStr(mf_sitename,mf_linkpath,ns_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
		Dim SpecialSql,RsSpecialObj
		if not isnumeric(f_id) then
			GetSpecialLocationStr = ""
		else
			SpecialSql = "Select SpecialID,SpecialCName from FS_NS_Special where SpecialID=" & f_id & ""
			Set RsSpecialObj = Conn.Execute(SpecialSql)
			if RsSpecialObj.Eof then
				GetSpecialLocationStr = ""
			else
				GetSpecialLocationStr = "<a href="""& mf_linkpath&""""&linkCss&MF_OpenMode&">首页</a>"& ns_path & fg_str &"<span"&positCss&">" & RsSpecialObj("SpecialCName") & "专题</span>"
			end if
			Set RsSpecialObj = Nothing
		end if
	End Function
	
	'得到商品当前路径
	'Function GetNewsLocationStr(NewsID,NaviType,CompatStr,OpenTypeStr,CSSStyleStr)
	Public Function GetmsNewsLocationStr(mf_sitename,mf_linkpath,ms_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
		Dim RsNewsObj
		Set RsNewsObj = Conn.Execute("Select Id,ClassID from FS_MS_Products where ID=" & f_id & "")
		if Not RsNewsObj.Eof then
			GetmsNewsLocationStr = GetmsClassLocationStr(mf_sitename,mf_linkpath,ms_path,RsNewsObj("ClassId"),fg_str,char_str,positCss,linkCss,MF_OpenMode,0) & fg_str & char_str
			Set RsNewsObj = Nothing
		else
			GetmsNewsLocationStr = GetmsClassLocationStr(mf_sitename,mf_linkpath,ms_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode,0) & fg_str & char_str
		end if
		Set RsNewsObj = Nothing
	End Function
	'商品栏目当前位置
	Public Function GetmsClassLocationStr(mf_sitename,mf_linkpath,ms_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode,f_TF)
		Dim SqlClass,RsClassObj,RsNewsObj
		if f_id = "" then Exit Function
		if f_TF=1 then
			Set RsClassObj = Conn.Execute("Select ParentID,ClassID,ClassCName from FS_MS_ProductsClass where id=" & f_id)
			'Set RsClassObj = Conn.Execute("Select ParentID,ClassID,ClassCName from FS_MS_ProductsClass where Classid='" & f_id & "'")
		else
			Set RsClassObj = Conn.Execute("Select ParentID,ClassID,ClassCName from FS_MS_ProductsClass where ClassId='" & f_id & "'")
		end if
		if Not RsClassObj.Eof then
			GetmsClassLocationStr = GetmsClassLocationStr & "<a  href=""" & get_msClassLink(RsClassObj("ClassID")) & """"&linkCss&MF_OpenMode&">" & RsClassObj("ClassCName") & "</a>"
			do while RsClassObj("ParentID") <> "0"
				Set RsClassObj = Conn.Execute("Select ParentID,ClassID,ClassCName from FS_MS_ProductsClass where ClassID='" & RsClassObj("ParentID") & "'")
				if RsClassObj.Eof then Exit do
				GetmsClassLocationStr = "<a href="""& get_msClassLink(RsClassObj("ClassID")) &""""&linkCss&MF_OpenMode&">"& RsClassObj("ClassCName") &"</a>"& fg_str & GetmsClassLocationStr
			loop
		end if
		GetmsClassLocationStr = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>" & ms_path & fg_str & GetmsClassLocationStr
		RsClassObj.Close
		Set RsClassObj = Nothing
	End Function
	'商品专题当前位置
	Function GetmsSpecialLocationStr(mf_sitename,mf_linkpath,ms_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
		Dim SpecialSql,RsSpecialObj
		if f_id = "" then
			GetSpecialLocationStr = ""
		else
			SpecialSql = "Select SpecialID,SpecialCName from FS_MS_Special where specialID=" & f_id & ""
			Set RsSpecialObj = Conn.Execute(SpecialSql)
			if RsSpecialObj.Eof then
				GetmsSpecialLocationStr = ""
			else
				GetmsSpecialLocationStr = "<a href="""& mf_linkpath&""""&linkCss&MF_OpenMode&">首页</a>"& ms_path & fg_str &"<span"&positCss&">" & RsSpecialObj("SpecialCName") & "专题</span>"
			end if
			Set RsSpecialObj = Nothing
		end if
	End Function
	
	'得到ds当前路径
	Public Function GetdsNewsLocationStr(mf_sitename,mf_linkpath,ds_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
		Dim RsNewsObj
		Set RsNewsObj = Conn.Execute("Select ClassID from FS_DS_List where DownLoadID='" & f_id & "'")
		
		if Not RsNewsObj.Eof then
			GetdsNewsLocationStr = GetdsClassLocationStr(mf_sitename,mf_linkpath,ds_path,RsNewsObj("ClassID"),fg_str,char_str,positCss,linkCss,MF_OpenMode) & fg_str & char_str
		else
			GetdsNewsLocationStr = GetdsClassLocationStr(mf_sitename,mf_linkpath,ds_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode) & fg_str & char_str
		end If
		
		Set RsNewsObj = Nothing
	End Function
	'ds栏目当前位置
	Public Function GetdsClassLocationStr(mf_sitename,mf_linkpath,ds_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
		Dim SqlClass,RsClassObj
		if f_id = "" then Exit Function
		Set RsClassObj = Conn.Execute("Select ParentID,ClassID,ClassName from FS_DS_Class where ClassID='" & f_id&"'")
		if Not RsClassObj.Eof then
			GetdsClassLocationStr = GetdsClassLocationStr & "<a  href=""" & get_dsClassLink(RsClassObj("ClassID")) & """"&linkCss&MF_OpenMode&">" & RsClassObj("ClassName") & "</a>"
			do while RsClassObj("ParentID") <> "0"
				Set RsClassObj = Conn.Execute("Select ParentID,ClassID,ClassName from FS_DS_Class where ClassID='" & RsClassObj("ParentID") & "'")
				if RsClassObj.Eof then Exit do
				GetdsClassLocationStr = "<a href="""& get_dsClassLink(RsClassObj("ClassID")) &""""&linkCss&MF_OpenMode&">"& RsClassObj("ClassName") &"</a>"& fg_str & GetdsClassLocationStr
			loop
		end if
		GetdsClassLocationStr = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">首页</a>" & ds_path & fg_str & GetdsClassLocationStr
		RsClassObj.Close
		Set RsClassObj = Nothing
	End Function
	
	'得到供求当前路径
	Public Function GetSDPageLocationStr(mf_sitename,mf_linkpath,SD_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
		Dim RsNewsObj
		Set RsNewsObj = Conn.Execute("Select ID,Classid from FS_SD_News where ID=" & f_id & "")
		if Not RsNewsObj.Eof then
			GetSDPageLocationStr = GetSDClassLocationStr(mf_sitename,mf_linkpath,SD_path,RsNewsObj("Classid"),fg_str,char_str,positCss,linkCss,MF_OpenMode) & fg_str & "供求浏览"
		else
			GetSDPageLocationStr = GetSDClassLocationStr(mf_sitename,mf_linkpath,SD_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode) & fg_str & "供求浏览"
		end if
		Set RsNewsObj = Nothing
	End Function
	'栏目当前位置
	Public Function GetSDClassLocationStr(mf_sitename,mf_linkpath,SD_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
		Dim SqlClass,RsClassObj
		if f_id = "" then Exit Function
		Set RsClassObj = Conn.Execute("Select ID,PID,GQ_ClassName from FS_SD_Class where ID=" & f_id & "")
		if Not RsClassObj.Eof then
			GetSDClassLocationStr = GetSDClassLocationStr & "<a  href=""" & getSDClassLink(RsClassObj("ID")) & """"&linkCss&MF_OpenMode&">" & RsClassObj("GQ_ClassName") & "</a>"
			do while cstr(RsClassObj("PID")) <> "0"
				Set RsClassObj = Conn.Execute("Select ID,PID,GQ_ClassName from FS_SD_Class where id=" & RsClassObj("PID") & "")
				if RsClassObj.Eof then Exit do
				GetSDClassLocationStr = "<a href="""& getSDClassLink(RsClassObj("ID")) &""""&linkCss&MF_OpenMode&">"& RsClassObj("GQ_ClassName") &"</a>"& fg_str & GetSDClassLocationStr
			loop
		end if
		GetSDClassLocationStr = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">" &"首页</a>"& fg_str &"" & SD_path & fg_str & GetSDClassLocationStr
		RsClassObj.Close
		Set RsClassObj = Nothing
	End Function
	
	
	'得到供求当前路径(区域)
	Public Function GetSDAreaPageLocationStr(mf_sitename,mf_linkpath,SD_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
		Dim RsNewsObj
		Set RsNewsObj = Conn.Execute("Select ID,Areaid from FS_SD_News where Areaid=" & f_id & "")
		if Not RsNewsObj.Eof then
			GetSDAreaPageLocationStr = GetSDAreaLocationStr(mf_sitename,mf_linkpath,SD_path,RsNewsObj("Areaid"),fg_str,char_str,positCss,linkCss,MF_OpenMode) & fg_str & "供求浏览"
		else
			GetSDAreaPageLocationStr = GetSDAreaLocationStr(mf_sitename,mf_linkpath,SD_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode) & fg_str & "供求浏览"
		end if
		
		Set RsNewsObj = Nothing
	End Function
	'栏目当前位置(区域)
	Public Function GetSDAreaLocationStr(mf_sitename,mf_linkpath,SD_path,f_id,fg_str,char_str,positCss,linkCss,MF_OpenMode)
		Dim SqlClass,RsClassObj
		if f_id = "" then Exit Function
		Set RsClassObj = Conn.Execute("Select ID,PID,ClassName from FS_SD_Address where ID=" & f_id & "")
		if Not RsClassObj.Eof then
			GetSDAreaLocationStr = GetSDAreaLocationStr & "<a  href=""" & getSDareaClassLink(RsClassObj("ID")) & """"&linkCss&MF_OpenMode&">" & RsClassObj("ClassName") & "</a>"
			do while cstr(RsClassObj("PID")) <>"0"
				Set RsClassObj = Conn.Execute("Select ID,PID,ClassName from FS_SD_Address where id=" & RsClassObj("PID") & "")
				if RsClassObj.Eof then Exit do
				GetSDAreaLocationStr = "<a href="""& getSDareaClassLink(RsClassObj("ID")) &""""&linkCss&MF_OpenMode&">"& RsClassObj("ClassName") &"</a>"& fg_str & GetSDAreaLocationStr
			loop
		end if
		GetSDAreaLocationStr = "<a href="""& mf_linkpath &""""&linkCss&MF_OpenMode&">" &"首页</a>"& fg_str &"" & SD_path & fg_str & GetSDAreaLocationStr
		RsClassObj.Close
		Set RsClassObj = Nothing
	End Function
	
			
	'得到供求栏目路径
	Public Function getSDClassLink(f_id)
		dim rs
		set rs = Conn.execute("select top 1 [Domain],SavePath from FS_SD_Config")
		if not rs.eof then
			if trim(rs("Domain"))<>"" and not isnull(rs("Domain")) then
				'---ken
				'getSDClassLink = "http://"&rs("Domain")&"/"& rs("SavePath")&"/SupplyList.asp?Id="& f_id &""
				getSDClassLink = "http://"&rs("Domain")&"/SupplyList.asp?Id="& f_id &""
			else
				getSDClassLink = "http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"& rs("SavePath")&"/SupplyList.asp?Id="& f_id &""
			end if
			rs.close:set rs = nothing
		else
			getSDClassLink = "http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/supply/SupplyList.asp?Id="& f_id &""
			rs.close:set rs = nothing
		end if
	end Function
	
	'得到供求栏目路径(区域)
	Public Function getSDareaClassLink(f_id)
		dim rs
		set rs = Conn.execute("select top 1 [Domain],SavePath from FS_SD_Config")
		if not rs.eof then
			if trim(rs("Domain"))<>"" and not isnull(rs("Domain")) then
				'---ken
				'getSDareaClassLink = "http://"&rs("Domain")&"/"& rs("SavePath")&"/SupplyArea.asp?Id="& f_id &""
				getSDareaClassLink = "http://"&rs("Domain")&"/SupplyArea.asp?Id="& f_id &""
			else
				getSDareaClassLink = "http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"& rs("SavePath")&"/SupplyArea.asp?Id="& f_id &""
			end if
			rs.close:set rs = nothing
		else
			getSDareaClassLink = "http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/supply/SupplyArea.asp?Id="& f_id &""
			rs.close:set rs = nothing
		end if
	end Function
	
	
	'得到新闻单个地址____________________________________________________________
	Public Function get_NewsLink(f_id)
		get_NewsLink = ""
		dim rs,config_rs,config_mf_rs,class_rs
		dim SaveNewsPath,FileName,FileExtName,ClassId,LinkType,MF_Domain,Url_Domain,ClassEName,c_Domain,c_SavePath,IsDomain
		set rs = Conn.execute("select ID,IsURL,URLAddress,ClassId,NewsId,SaveNewsPath,FileName,FileExtName From FS_NS_News where NewsId='"&f_id&"'")
		SaveNewsPath = rs("SaveNewsPath")
		FileName = rs("FileName")
		FileExtName = rs("FileExtName")
		ClassId = rs("ClassId")
		LinkType = Request.Cookies("FoosunNSCookies")("FoosunNSLinkType")
		IsDomain = Request.Cookies("FoosunNSCookies")("FoosunNSDomain")
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		set class_rs = Conn.execute("select ClassEName,IsURL,URLAddress,[Domain],SavePath From FS_NS_NewsClass where ClassId='"&ClassId&"'")
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
			if rs("IsURL")=0 then
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
					get_NewsLink = Url_Domain & replace(SaveNewsPath &"/"&FileName&"."&FileExtName,"//","/")
				else
					get_NewsLink = Url_Domain & replace(c_SavePath& "/" & ClassEName &SaveNewsPath &"/"&FileName&"."&FileExtName,"//","/")
				end if
			else
				get_NewsLink = rs("URLAddress")
			end if
		rs.close:set rs=nothing
	  else
			get_NewsLink = ""
			rs.close:set rs=nothing
	  end if
	  get_NewsLink = get_NewsLink
	End Function
	
	'得到栏目地址________________________________________________________________
	Public function get_ClassLink(f_id)
		dim IsDomain,LinkType,MF_Domain,c_rs,ClassEName,c_Domain,Url_Domain,ClassSaveType,class_savepath,FileExtName,c_SavePath
		LinkType = Request.Cookies("FoosunNSCookies")("FoosunNSLinkType")
		IsDomain = Request.Cookies("FoosunNSCookies")("FoosunNSDomain")
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		set c_rs=conn.execute("select id,IsURL,UrlAddress,ClassId,ClassEName,[Domain],FileSaveType,FileExtName,SavePath From FS_NS_NewsClass where ClassId='"&f_id&"'")
		if not c_rs.eof then
			if c_rs("IsURL")=0 then
				ClassEName = c_rs("ClassEName")
				c_Domain = c_rs("Domain")
				FileExtName = c_rs("FileExtName")
				ClassSaveType = c_rs("FileSaveType")
				c_SavePath= c_rs("SavePath")
				if trim(c_Domain)<>"" then
					if ClassSaveType=0 then
						class_savepath = "Index."&FileExtName
					elseif ClassSaveType=1 then
						class_savepath =ClassEName &"."&FileExtName
					else
						class_savepath = ClassEName &"."&FileExtName
					end if
				else
					if ClassSaveType=0 then
						class_savepath = ClassEName&"/Index."&FileExtName
					elseif ClassSaveType=1 then
						class_savepath = ClassEName&"/"& ClassEName &"."&FileExtName
					else
						class_savepath = ClassEName &"."&FileExtName
					end if
				end if
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
					get_ClassLink = Url_Domain&Replace("/"&class_savepath,"//","/")
				else
					get_ClassLink = Url_Domain&Replace(c_SavePath&"/"&class_savepath,"//","/")
				end if
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
		set c_rs=conn.execute("select SpecialID,SpecialCName,SpecialEName,SavePath,ExtName,isLock From FS_NS_Special where SpecialEName='"&f_id&"'")
		if not c_rs.eof then
			SpecialEName = c_rs("SpecialEName")
			ExtName = c_rs("ExtName")
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
			get_specialLink = ""
		end if
		get_specialLink = get_specialLink
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
			if G_VIRTUAL_ROOT_DIR<>"" then
				c_SavePath = "/"& G_VIRTUAL_ROOT_DIR&c_SavePath
			end if
		else
			ClassEName = ""
		end if
		class_rs.close:set class_rs=nothing
		'---ken
		If Trim(c_Domain) <> "" And Not IsNull(c_Domain) Then
			get_productsLink = "http://" & c_Domain & Replace("/" &SaveproductsPath &"/"&fileName&"."&fileExtName,"//","/")
		Else
			If LinkType and IsDomain<>"" then
				Url_Domain = "http://"&IsDomain
			Else
				Url_Domain = ""
			End If
			If Url_Domain <> "" Then
				get_productsLink = Url_Domain & Replace("/" & ClassEName &SaveproductsPath &"/"&fileName&"."&fileExtName,"//","/")
			Else
				get_productsLink = Url_Domain & Replace(c_SavePath& "/" & ClassEName &SaveproductsPath &"/"&fileName&"."&fileExtName,"//","/")
			End If				
		End If	 
	get_productsLink = get_productsLink
	End function		
	
	'得到商品栏目地址
	Public function get_msClassLink(f_id)
		dim config_rs,IsDomain,LinkType,Mf_Domain,c_rs,ClassEName,c_Domain,Url_Domain,ClassSaveType,class_savepath,fileExtName,c_SavePath
		set config_rs = Conn.execute("select top 1 IsDomain,SavePath from fS_MS_SysPara")
		'---ken
		IsDomain = config_rs("IsDomain")
		If Trim(IsDomain) <> "" And Not IsNull(IsDomain) Then
			LinkType = 1 
		Else
			LinkType = 0
		End If	
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
				if G_VIRTUAL_ROOT_DIR<>"" then
					c_SavePath = "/"&G_VIRTUAL_ROOT_DIR&c_SavePath
				end if
				'---ken
				If c_Domain <> "" And Not IsNull(c_Domain) Then
					if ClassSaveType=0 then
						class_savepath = "/Index."&fileExtName
					elseif ClassSaveType=1 then
						class_savepath = ClassEName &"."&fileExtName
					else
						class_savepath = ClassEName &"."&fileExtName
					end if
				Else
					if ClassSaveType=0 then
						class_savepath = ClassEName&"/Index."&fileExtName
					elseif ClassSaveType=1 then
						class_savepath = ClassEName&"/"& ClassEName &"."&fileExtName
					else
						class_savepath = ClassEName &"."&fileExtName
					end if
				End If
				class_savepath = class_savepath
				'----ken
				if LinkType = 1 then
					if trim(c_Domain)<>"" then
						Url_Domain = "http://"&c_Domain
					else
						Url_Domain = "http://"&IsDomain
					end if
				else
					if trim(c_Domain)<>"" then
						Url_Domain = "http://"&c_Domain
					else
						Url_Domain = ""
					end if
				end if
				If Url_Domain <> "" Then
					get_msClassLink = Url_Domain&Replace("/"&class_savepath,"//","/")
				Else
					get_msClassLink = Url_Domain&Replace(c_SavePath&"/"&class_savepath,"//","/")
				End If		
			else
				get_msClassLink = c_rs("UrlAddress")
			end if
			c_rs.close:set c_rs=nothing
		else
			c_rs.close:set c_rs=nothing
			get_msClassLink = ""
		end if
		get_msClassLink = get_msClassLink
	End function
	'得到商品专题地址
	Public function get_msspecialLink(f_id)
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
				get_msspecialLink = "Http://" & Ms_Sp_Domain & "/special_"& SpecialEName & "." & ExtName
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
					get_msspecialLink = Url_Domain & Replace("/special_" & SpecialEName & "." & ExtName,"//","/")
				Else
					get_msspecialLink = Replace("/" & c_SavePath & "/special_" & SpecialEName & "." & ExtName,"//","/")
				End If
			End If				 
			c_rs.close:set c_rs=nothing
		else
			c_rs.close:set c_rs=nothing
			get_msspecialLink = ""
		end if
		get_msspecialLink = get_msspecialLink
	End function
	
	
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
						if G_VIRTUAL_ROOT_DIR<>"" then
							Url_Domain = "/"& G_VIRTUAL_ROOT_DIR
						else
							Url_Domain = ""
						end if
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
	Public function get_dsClassLink(f_id)
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
					get_dsClassLink = Url_Domain&Replace(c_SavePath&"/"&class_savepath,"//","/")
				Else
					Dim Savepathlen,SubPathlen
					SubPathlen = Len("/"&Conn.execute("select Sub_Sys_Path from FS_MF_Sub_Sys Where Sub_Sys_ID='DS'")(0))
					Savepathlen = Len(c_SavePath)
					c_SavePath = Right(c_SavePath,Savepathlen - SubPathlen)
					get_dsClassLink = Url_Domain&Replace(c_SavePath&"/"&class_savepath,"//","/")
				End if
				'结束
			else
				get_dsClassLink = c_rs("UrlAddress")
			end if
			c_rs.close:set c_rs=nothing
		else
			c_rs.close:set c_rs=nothing
			get_dsClassLink = ""
		end if
		get_dsClassLink = get_dsClassLink
	End function		
	
	'得到专题地址________________________________________________________________
	Public function get_DSspecialLink(f_id)
		dim IsDomain,LinkType,MF_Domain,c_rs,SpecialEName,ExtName,c_SavePath,Url_Domain
		LinkType = Request.Cookies("FoosunNSCookies")("FoosunNSLinkType")
		IsDomain = Request.Cookies("FoosunNSCookies")("FoosunNSDomain")
		MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		set c_rs=conn.execute("select SpecialID,SpecialCName,SpecialEName,SavePath,FileExtName,isLock From FS_DS_Special where SpecialEName='"&f_id&"'")
		if not c_rs.eof then
			SpecialEName = c_rs("SpecialEName")
			ExtName = c_rs("FileExtName")
			c_SavePath= c_rs("SavePath")
			if G_VIRTUAL_ROOT_DIR<>"" then
				c_SavePath = "/"& G_VIRTUAL_ROOT_DIR&c_SavePath
			end if
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
					Url_Domain = ""
				end if
			end if
			get_DSspecialLink = Url_Domain&Replace(c_SavePath&"/special_"&SpecialEName&"."&ExtName,"//","/")
			c_rs.close:set c_rs=nothing
		else
			c_rs.close:set c_rs=nothing
			get_DSspecialLink = "找不到参数，错误的地址"
		end if
		get_DSspecialLink = get_DSspecialLink
	End function
	
	
	Public Function GetFriendNumber(f_strNumber)
		Dim RsGetFriendNumber
		Set RsGetFriendNumber = User_Conn.Execute("Select UserNumber From FS_ME_Users Where UserName = '"& f_strNumber &"'")
		If  Not RsGetFriendNumber.eof  Then 
			GetFriendNumber = RsGetFriendNumber("UserNumber")
		Else
			GetFriendNumber = "用户已经被删除"
		End If 
		set RsGetFriendNumber = nothing
	End Function
	
	Public Function GetOpenModeStr(OpenStr)
		If OpenStr = "" Then
			GetOpenModeStr = ""
		ElseIf Cint(OpenStr) = 1 Then
			GetOpenModeStr = " target=""_blank"""
		ElseIf Cint(OpenStr) = 0 Then
			GetOpenModeStr = " target=""_self"""
		Else
			GetOpenModeStr = ""
		End If			
	End Function 
	
	'=====================================================
	'自由标签解析函数
	'=====================================================
	Public Function FreeLabel(f_Lable,f_type,f_id,f_position)
		f_Lable.IsSave = True
		Dim LableID,GetConObj,SqlStr,LableConStr,SelectNum,SysType
		LableID = f_Lable.LablePara("自由标签")
		If LableID = "" Or IsNull(LableID) Or Len(LableID) <> 15 Then
			FreeLabel = "无效的自由标签"
		End If
		Set GetConObj = Conn.ExeCute("Select LabelSQl,LabelContent,selectNum,SysType From FS_MF_FreeLabel Where LabelID = '" & LableID & "'")
		If GetConObj.Eof Then
			FreeLabel = "无效的自由标签"
		Else
			SqlStr = GetConObj(0)
			LableConStr = GetConObj(1)
			SelectNum = GetConObj(2)
			SysType = GetConObj(3)
			FreeLabel = CreatLableContent(SqlStr,LableConStr,SysType)
		End If
		GetConObj.Close : Set GetConObj = Nothing
		f_Lable.DictLableContent.Add "1",FreeLabel
	End Function
	
	Public Function CreatLableContent(SqlStr,LableConStr,SysType)
		Dim SqlRs,TempSql,TempArr,FromArr,FieldsArr,TempContent,Sql_Str
		Dim NoNextReg,NoNextNumReg,NoNextConReg,NoNextFlag,NoNextIDReg
		Dim NextReg,NextConReg,NextFlag,NextContent
		Dim FunReg,FunStrReg,Fun_i,FunStr
		Dim TimeReg,TimeStrReg,Time_i,TimeStr
		Dim AutoReg,AutoStrReg,Auto_i,AutoStr
		Dim NoNext_i
		Dim NoNextContent
		Dim NoNextDataID
		Dim FunArr
		Dim TimeArr
		Dim AutoArr
		Dim NextIndexNum,Temp_Str,HadShowFlag,NextNum,TemPNextContent
		Dim NoNextIndexNum,NoNum,TempNoNextContent
		'On Error Resume Next
		Sql_Str = Replace(SqlStr,"*",",")
		Set SqlRs = Server.CreateObject(G_FS_RS)
		SqlRs.Open Sql_Str,Conn,0,1
		If Err.Number <> 0 Then
			CreatLableContent = "查询SQL语句出错，请检查"
		End If
		If SqlRs.Eof Then
			CreatLableContent = "此标签暂无记录"
		End If
		
		TempContent = LableConStr
		NoNextFlag = False
		NextFlag = False
		
		'分解字段
		TempSql = SqlStr
		TempArr = Split(TempSql," ")
		TempSql = Replace(TempSql,TempArr(0) & " " & TempArr(1) & " " & TempArr(2) & " ","")
		FromArr = Split(TempSql,"From")
		TempSql = Trim(FromArr(0))
		TempSql = Replace(TempSql,",","@|@")
		TempSql = Replace(TempSql,"*",",")
		IF Instr(TempSql,"@|@") > 0 Then
			FieldsArr = Split(TempSql,"@|@")
		Else
			FieldsArr = Array(TempSql)
		End If	
		
		
		'提取不循环部分及不循环ID
		Set NoNextReg = New RegExp
		NoNextReg.Pattern = "{\*[1-9]+[0-9]*(\{[^\*]|[^\{]\*[^\}]|[^\*]\}|[^\{\*\}])*\*\}"
		NoNextReg.IgnoreCase = True
		NoNextReg.Global = True
		Set NoNextConReg = NoNextReg.ExeCute(LableConStr)
		ReDim NoNextContent(NoNextConReg.Count)
		ReDim NoNextDataID(NoNextConReg.Count)
		If NoNextConReg.Count > 0 Then
			NoNextFlag = True
			Set NoNextNumReg = New RegExp
			NoNextNumReg.Pattern = "{\*[1-9]+[0-9]*"
			NoNextNumReg.IgnoreCase = True
			NoNextNumReg.Global = True
			For NoNext_i = 0 To NoNextConReg.Count - 1
				NoNextContent(NoNext_i) = NoNextConReg.Item(NoNext_i)
				Set NoNextIDReg = NoNextNumReg.ExeCute(NoNextContent(NoNext_i))
				If NoNextIDReg.Count > 0 Then
					NoNextDataID(NoNext_i) = Cint(Mid(NoNextIDReg.Item(0),3))
				Else
					NoNextDataID(NoNext_i) = 0
				End If		
			Next
		End If
		
		'提取循环部分
		Set NextReg = New RegExp
		NextReg.Pattern = "\{\#({[^#]|[^{]#[^}]|[^#]}|[^{#}])*#}"
		NextReg.IgnoreCase = True
		NextReg.Global = True
		Set NextConReg = NextReg.ExeCute(LableConStr)
		If NextConReg.Count > 0 Then
			NextContent = NextConReg.Item(0)
			NextFlag = True
		End If
				
		
		'如果没有设置循环和不循环标记，则默认全部为不循环
		If NoNextFlag = False And NextFlag = False Then
			ReDim NoNextContent(0)
			ReDim NoNextDataID(0)
			NoNextContent(0) = LableConStr
			NoNextDataID(0) = 0
			NoNextFlag = true
		End If
		
		'从样式中提取自定义函数表达式
		Set FunReg = New RegExp
		FunReg.Pattern = "\(\#.*?\#\)"
		FunReg.IgnoreCase = True
		FunReg.Global = True
		Set FunStrReg = FunReg.ExeCute(LableConStr)
		ReDim FunArr(FunStrReg.Count)
		If FunStrReg.Count > 0 Then
			For Fun_i = 0 To FunStrReg.Count - 1
				FunStr = FunStrReg.Item(Fun_i)
				FunStr = Replace(FunStr,"(#","")
				FunStr = Replace(FunStr,"#)","")
				FunArr(Fun_i) = FunStr
			Next
		End If
		
		
		'提取日期表达式
		Set TimeReg = New RegExp
		TimeReg.Pattern = "\[\$.*?\$\]"
		TimeReg.IgnoreCase = True
		TimeReg.Global = True
		Set TimeStrReg = TimeReg.ExeCute(LableConStr)
		ReDim TimeArr(TimeStrReg.Count)
		If TimeStrReg.Count > 0 Then
			For Time_i = 0 To TimeStrReg.Count - 1
				TimeStr = TimeStrReg.Item(Time_i)
				TimeStr = Replace(TimeStr,"[$","")
				TimeStr = Replace(TimeStr,"$]","")
				TimeArr(Time_i) = TimeStr
			Next
		End If
		
		
		
		'提取预定义字段表达式
		Set AutoReg = New RegExp
		AutoReg.Pattern = "\[\#.*?\#\]"
		AutoReg.IgnoreCase = True
		AutoReg.Global = True
		Set AutoStrReg = AutoReg.ExeCute(LableConStr)
		ReDim AutoArr(AutoStrReg.Count)
		If AutoStrReg.Count > 0 Then
			For Auto_i = 0 To AutoStrReg.Count - 1
				AutoStr = AutoStrReg.Item(Auto_i)
				AutoStr = Replace(AutoStr,"[#","")
				AutoStr = Replace(AutoStr,"#]","")
				AutoArr(Auto_i) = AutoStr
			Next
		End If
		
		
		'生成不循环部分内容
		If NoNextFlag = True Then
			If Not SqlRs.Eof Then
				NoNextIndexNum = 0
				Do While Not SqlRs.Eof
					For NoNum = 0 To UBound(NoNextContent)
						If NoNextDataID(NoNum) - 1 = NoNextIndexNum Then
							TempNoNextContent = Replace(Replace(Replace(Replace(NoNextContent(NoNum),"{*" & NoNextDataID(NoNum),""),"{*0",""),"{*",""),"*}","")
							TempNoNextContent = GetContentByObj(SqlRs,TempNoNextContent,FieldsArr,FunArr,TimeArr,AutoArr,SysType)
							TempContent = Replace(TempContent,NoNextContent(NoNum),TempNoNextContent)
						End If	
					Next
				NoNextIndexNum = NoNextIndexNum + 1	
				SqlRs.MoveNext
				Loop
				SqlRs.MoveFirst
			End If
			
			'如果没有生成则将代码替换为空.
			For NoNum = 0 To UBound(NoNextContent)
				TempContent = Replace(TempContent,NoNextContent(NoNum),"")
			Next
		End If
		
		
		'生成循环部分内容
		NextIndexNum = 0
		If NoNextFlag = True And Not SqlRs.Eof Then
			Do
				HadShowFlag = False
				For NextNum = 0 To UBound(NoNextDataID)
					If NoNextDataID(NextNum) - 1 = NextIndexNum Then
						HadShowFlag = True
						Exit For
					End If
				Next
				If HadShowFlag = True Then
					NextIndexNum = NextIndexNum + 1
					SqlRs.MoveNext
				End If
			Loop While Not SqlRs.Eof And HadShowFlag = True				
		End If
		Temp_Str = ""
		If NextFlag = True Then
			If Not SqlRs.Eof Then
				Do While Not SqlRs.Eof
					TemPNextContent = Replace(Replace(NextContent,"{#",""),"#}","")
					TemPNextContent = GetContentByObj(SqlRs,TemPNextContent,FieldsArr,FunArr,TimeArr,AutoArr,SysType)
					Temp_Str = Temp_Str & TemPNextContent
					NextIndexNum = NextIndexNum + 1
				SqlRs.MoveNext
				Loop
			End If
		End IF			
		CreatLableContent = Replace(TempContent,NextContent,Temp_Str)
		
		Set FieldsArr = Nothing
		Set FunArr = Nothing
		Set TimeArr = Nothing
		Set AutoArr = Nothing		
		SqlRs.Close : Set SqlRs = NOthing
	End Function
	
	
	
	Public Function GetContentByObj(SqlRs,TemPNextContent,FieldsArr,FunArr,TimeArr,AutoArr,SysType)
		Dim RegEx,i,TempStr,Matches,l,TempContent,TempMatches,TempMatch,Match,Temp_Str
		Dim StrTemp,TempTimeStr,m,k,Value_Str
		Dim TempFieldsCon,TempUrlStr
		Set RegEx = New RegExp
		RegEx.IgnoreCase = True
		RegEx.Global = True
		
		'解析自定义函数
		For i = 0 To Ubound(FunArr)
			TempStr = FunArr(i)
			RegEx.Pattern = "[ ]*Left\([ ]*\[\*[^\[\*]*\*\][ ]*,[ ]*[1-9][0-9]*[ ]*\)"
			Set Matches = RegEx.ExeCute(TempStr)
			For Each Match In Matches
				For l = 0 To Ubound(FieldsArr)
					If Instr(Match,"[*" & FieldsArr(l) & "*]") <> 0 Then
						TempContent = SqlRs(l)
						RegEx.Pattern = "<\/*[^<>]*>"
						Set TempMatches = RegEx.ExeCute(TempContent)
						For Each TempMatch In TempMatches
							TempContent = Replace(TempContent,TempMatch,"")
						Next
						TempStr = Replace(TempStr,Match,Lose_Html(GotTopic(TempContent,Clng(Mid(Match,InStrRev(Match,",")+1,InStrRev(Match,")")-InStrRev(Match,",")-1)))))
						Exit For
					End IF		
				Next
			Next
			TemPNextContent = Replace(TemPNextContent,"(#" & FunArr(i) & "#)",TempStr)
		Next
		
		'生成日期样式内容
		For i = 0 to UBound(TimeArr)
			StrTemp = ""
			TempTimeStr = Trim(TimeArr(i))
			For m = 0 to UBound(FieldsArr)
				If Instr(Lcase(SqlRs(m).Name),"addtime") > 0 Then
					StrTemp = Get_Date(SqlRs(m),TempTimeStr)
				End if
			Next
			TemPNextContent = Replace(TemPNextContent,"[$" & TimeArr(i) & "$]",StrTemp)
		Next

		'生成字段内容
		For k = 0 to UBound(FieldsArr)
			IF SysType = "DS" Then
				If Instr(LCase(FieldsArr(k)),"accredit") > 0 Then
					Value_Str = SqlRs(k)
					TempFieldsCon = Replacestr(Value_Str,"1:免费,2:共享,3:试用,4:演示,5:注册,6:破解,7:零售,8:其它")
				ElseIf Instr(LCase(FieldsArr(k)),"appraise") > 0 Then
					Value_Str = SqlRs(k)
					TempFieldsCon = Replacestr(Value_Str,"1:★,2:★★,3:★★★,4:★★★★,5:★★★★★,6:★★★★★★")
				ElseIf Instr(LCase(FieldsArr(k)),"types") > 0 Then
					Value_Str = SqlRs(k)
					TempFieldsCon = Replacestr(Value_Str,"1:图片,2:文件,3:程序,4:Flash,5:音乐,6:影视,7:其它")
				ElseIf Instr(LCase(FieldsArr(k)),"rectf") > 0 Then
					Value_Str = SqlRs(k)
					TempFieldsCon = Replacestr(Value_Str,"1:推荐,0:不推荐")
				ElseIf Instr(LCase(FieldsArr(k)),"overdue") > 0 Then
					Value_Str = SqlRs(k)
					TempFieldsCon = Replacestr(Value_Str,"0:永不过期," & Value_Str & ":" & Value_Str & "天")
				ElseIf Instr(LCase(FieldsArr(k)),"speicalid") > 0 Then
					Value_Str = SqlRs(k)
					If Value_Str = "" Or IsNull(Value_Str) Then
						TempFieldsCon = ""
					Else
						TempFieldsCon = Get_DownSpLink(Value_Str)
					End If	
				Else
					TempFieldsCon = SqlRs(k)
					if IsNull(TempFieldsCon) Then
						TempFieldsCon = ""
					End if
				End If	
			ElseIF SysType = "MS" Then
				If Instr(LCase(FieldsArr(k)),"iswholesale,0,1") > 0 Then
					Value_Str = SqlRs(k)
					TempFieldsCon = Replacestr(Value_Str,"1:批发,0:不批发")
				ElseIf Instr(LCase(FieldsArr(k)),"isinvoice") > 0 Then
					Value_Str = SqlRs(k)
					TempFieldsCon = Replacestr(Value_Str,"1:有发票,0:无发票")
				ElseIf Instr(LCase(FieldsArr(k)),"salestyle") > 0 Then
					Value_Str = SqlRs(k)
					TempFieldsCon = Replacestr(Value_Str,"0:正常销售,1:竞拍,2:一口价,3:特价,4:降价")
				ElseIf Instr(LCase(FieldsArr(k)),"specialid") > 0 Then
					Value_Str = SqlRs(k)
					If Value_Str = "" Or IsNUll(Value_Str) Then
						TempFieldsCon = ""
					Else
						TempFieldsCon = GetMall_SPLink(Value_Str)
					End If	
				Else
					TempFieldsCon = SqlRs(k)
					if IsNull(TempFieldsCon) Then
						TempFieldsCon = ""
					End if
				End If
			ElseIf SysType = "NS" Then
				If Instr(FieldsArr(k),"SpecialEName") > 0 Then
					Value_Str = SqlRs(k)
					If Value_Str = "" Or IsNull(Value_Str) Then
						TempFieldsCon = ""
					Else
						TempFieldsCon = GetNew_SpLink(Value_Str)
					End IF	
				Else
					TempFieldsCon = SqlRs(k)
					if IsNull(TempFieldsCon) Then
						TempFieldsCon = ""
					End if
				End If
			End IF
			TemPNextContent = Replace(TemPNextContent,"[*" & FieldsArr(k) & "*]",TempFieldsCon)
		Next

		'生成预定义字段内容
		For k = 0 to UBound(AutoArr)
			TempUrlStr = ""
			Select Case Trim(AutoArr(k))
				Case "NewsUrl"
					For m = 0 to UBound(FieldsArr)
						If Instr(Lcase(SqlRs(m).Name),"newsid") > 0 Then
							TempUrlStr = get_NewsLink(Trim(SqlRs(m)))
							Exit for
						End if
					Next
				Case "NewsClassUrl"
					For m = 0 to UBound(FieldsArr)
						If Instr(Lcase(SqlRs(m).Name),"classid") > 0 Then
							TempUrlStr = get_ClassLink(Trim(SqlRs(m)))
							Exit for
						End if
					Next
				Case "DownUrl"
					For m = 0 to UBound(FieldsArr)
						If Instr(Lcase(SqlRs(m).Name),"downloadid") > 0 Then
							TempUrlStr = get_DownLink(Trim(SqlRs(m)))
							Exit for
						End if
					Next
				Case "DownClassUrl"
					For m = 0 to UBound(FieldsArr)
						If Instr(Lcase(SqlRs(m).Name),"classid") > 0 Then
							TempUrlStr = get_dsClassLink(Trim(SqlRs(m)))
							Exit for
						End if
					Next
				Case "DownAddress"
					For m = 0 to UBound(FieldsArr)
						If Instr(Lcase(SqlRs(m).Name),"downloadid") > 0 Then
							TempUrlStr = GetDown_Address(Trim(SqlRs(m)))
							Exit for
						End if
					Next
				Case "MallUrl"
					For m = 0 to UBound(FieldsArr)
						IF LCase(Trim(SqlRs(m).Name)) = "id" OR LCase(Right(Trim(SqlRs(m).Name),3)) = ".id" Then
							TempUrlStr = get_productsLink(Trim(SqlRs(m)))
							Exit for
						End if
					Next
				Case "MallClassUrl"
					For m = 0 to UBound(FieldsArr)
						If Instr(Lcase(SqlRs(m).Name),"classid") > 0 Then
							TempUrlStr = get_msClassLink(Trim(SqlRs(m)))
							Exit for
						End if
					Next
			End Select
			TemPNextContent = Replace(TemPNextContent,"[#" & AutoArr(k) & "#]",TempUrlStr)
		Next
		
		'清除多余字段						
		RegEx.Pattern = "\[\*[^\[\]\*]*\*\]"
		Set Matches = RegEx.Execute(TemPNextContent)
		For Each Match In Matches
			TemPNextContent = Replace(TemPNextContent,Match,"")
		Next
		
		GetContentByObj = TemPNextContent
	End Function
	
	
	'获得下载地址列表
	Public Function GetDown_Address(DownID)
		Dim Rs,Str
		Set Rs = Conn.ExeCute("Select AddressName,Url From FS_DS_Address Where DownLoadID = '" & DownID & "' Order By Number Desc,ID Desc")
		If Rs.Eof Then 
			GetDown_Address = "暂无"
		Else
			Str = ""
			Do While Not Rs.Eof
				Str = Str & "<a href=""" & Rs(1) & """>" & Rs(0) & "</a>&nbsp;&nbsp;"
			Rs.MoveNExt
			Loop
			GetDown_Address = Str
		End If
		Rs.Close : Set Rs = Nothing
	End Function
	
	
	'获得新闻所属专题列表
	Public Function GetNew_SpLink(SPEname)
		Dim Rs,Arr,i,Str
		IF Instr(SPEname,",") > 0 Then
			Arr = Split(SPEname,",")
			Str = ""
			For i = LBound(Arr) To Ubound(Arr)
				Set Rs = Conn.ExeCute("Select SpecialID,SpecialCName From FS_NS_Special Where SpecialEName = '" & Arr(i) & "'")
				If Rs.Eof Then
					Str = Str & ""
				Else
					Str = Str & "┆<a href=""" & get_specialLink(Rs(0)) & """>" & Rs(1) & "</a>"
				End If	
			Next
			If Left(Str,1) = "┆" Then
				Str = Right(Str,Len(Str) - 1)
			End IF	
			GetNew_SpLink = Str
		Else
			Set Rs = Conn.ExeCute("Select SpecialID,SpecialCName From FS_NS_Special Where SpecialEName = '" & SPEname & "'")
			IF Rs.Eof Then
				GetNew_SpLink = ""
			Else	
				GetNew_SpLink = "<a href=""" & get_specialLink(Rs(0)) & """>" & Rs(1) & "</a>"
			End If	
		End If
		Rs.Close : Set Rs = NOthing
	End Function
	
	
	'获得商品专题列表
	Public Function GetMall_SPLink(SPID)
		Dim Rs
		Set Rs = Conn.ExeCute("Select specialID,SpecialCName,SpecialEName From FS_MS_Special Where specialID = " & SPID)
		If Rs.EOf Then
			GetMall_SPLink = ""
		Else
			GetMall_SPLink = "<a href=""" & get_msspecialLink(Rs(2)) & """>" & Rs(1) & "</a>"
		End If
		Rs.Close : Set Rs = Nothing	
	End Function
	
	
	'获得下载专题列表
	Public Function Get_DownSpLink(SPID)
		Dim Rs
		Set Rs = Conn.ExeCute("Select specialID,SpecialCName,SpecialEName From FS_DS_Special Where specialID = " & SPID)
		If Rs.EOf Then
			Get_DownSpLink = ""
		Else
			Get_DownSpLink = "<a href=""" & get_DSspecialLink(Rs(2)) & """>" & Rs(1) & "</a>"
		End If
		Rs.Close : Set Rs = Nothing	
	End Function
End Class
%>
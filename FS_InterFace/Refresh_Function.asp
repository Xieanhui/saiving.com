<!--#include file="CLS_Foosun.asp" -->
<%
Sub RefreshSession_Initialize()
	Session("FOOSUN_LABLE_MASS") = ""
	Session("FOOSUN_LABLE") = ""
	Session("FOOSUN_STYLECONTENT") = ""
	Session("FOOSUN_STYLELOOPCONTENT") = ""
	Session("FOOSUN_STYLETYPE") = ""
	Session("FOOSUN_REFRESH_LAST_TIME") = NOW
End Sub

Sub RefreshSession_Terminate()
	Session("FOOSUN_LABLE_MASS") = ""
	Session("FOOSUN_LABLE") = ""
	Session("FOOSUN_STYLECONTENT") = ""
	Session("FOOSUN_STYLELOOPCONTENT") = ""
	Session("FOOSUN_STYLETYPE") = ""
	Session("FOOSUN_REFRESH_LAST_TIME") = ""
End Sub

Function GetRefreshLastTime()
	Dim f_StartTime,f_LastTime,f_Array
	if Not IsDate(Session("FOOSUN_REFRESH_LAST_TIME")) then
		GetRefreshLastTime = ""
	else
		f_StartTime = CDate(Session("FOOSUN_REFRESH_LAST_TIME"))
		f_LastTime = CStr(CDate(NOW - f_StartTime))
		f_Array = Split(f_LastTime,":")
		GetRefreshLastTime = f_Array(0) & "小时" & f_Array(1) & "分" & f_Array(2) & "秒"
	end if
End Function

Function Refresh_index(Sys_flag,f_SYS_ROOT_DIR)
	Dim f_Templet,f_HTML_File_Save_Path_Str,f_HTML_File_Save_Phy_Path_Str,f_FileName,f_Phy_Templet
	Dim f_File_Content,f_Templet_Have_Error,f_LabelMassList,f_SaveFileExtName,f_SaveFileName,f_LastTimeStr
	if Request.Cookies("FoosunSUBCookie")("FoosunSUB" & Sys_flag) = "1" then
		Select Case Sys_flag
		    Case "NS"
				If Request.Cookies("FoosunNSCookies")("FoosunNSNewsDir")="" Then NSConfig_Cookies
				f_Templet = f_SYS_ROOT_DIR & Request.Cookies("FoosunNSCookies")("FoosunNSIndexTemplet")
				f_HTML_File_Save_Path_Str = f_SYS_ROOT_DIR & "/" & Request.Cookies("FoosunNSCookies")("FoosunNSNewsDir") 
				f_FileName = Request.Cookies("FoosunNSCookies")("FoosunNSIndexPage")
			Case "DS"
				If Request.Cookies("FoosunDSCookies")("FoosunDSDownDir")="" Then DSConfig_Cookies
				f_Templet = f_SYS_ROOT_DIR & Request.Cookies("FoosunDSCookies")("FoosunDSIndexTemplet")
				f_HTML_File_Save_Path_Str = f_SYS_ROOT_DIR & "/"& Request.Cookies("FoosunDSCookies")("FoosunDSDownDir")
				f_FileName=Request.Cookies("FoosunDSCookies")("FoosunDSIndexPage")
			Case "MS"
				If Request.Cookies("FoosunMSCookies")("FoosunMSDir")="" Then MSConfig_Cookies
				f_Templet = f_SYS_ROOT_DIR & Request.Cookies("FoosunMSCookies")("FoosunMSIndexTemplet")
				f_HTML_File_Save_Path_Str = f_SYS_ROOT_DIR & "/"& Request.Cookies("FoosunMSCookies")("FoosunMSDir")
				f_FileName="index." & Request.Cookies("FoosunMSCookies")("FoosunMSIndexHtml")
			Case "MF"
				If Request.Cookies("FoosunMFCookies")("FoosunMFIndexTemplet")="" Then MFConfig_Cookies
				f_Templet = f_SYS_ROOT_DIR & Request.Cookies("FoosunMFCookies")("FoosunMFIndexTemplet")
				f_HTML_File_Save_Path_Str = f_SYS_ROOT_DIR
				f_FileName=Request.Cookies("FoosunMFCookies")("FoosunMFIndexFileName")
			Case Else
				Refresh_index="No$$"
				Exit Function
		End Select
	Else
		Refresh_index="No$$"
		Exit Function
	End If
	f_HTML_File_Save_Path_Str=Replace(f_HTML_File_Save_Path_Str,"//","/")
	If f_HTML_File_Save_Path_Str <> "" Then
		f_HTML_File_Save_Phy_Path_Str = Server.MapPath(f_HTML_File_Save_Path_Str)
		CreatePath f_HTML_File_Save_Phy_Path_Str,Server.MapPath("/")
	End If
	If f_HTML_File_Save_Path_Str="/" Then f_HTML_File_Save_Path_Str=""
	f_File_Content = GetTempletContent(f_Templet,f_Templet_Have_Error)	
	f_SaveFileExtName = Right(f_FileName,Len(f_FileName) - InStr(f_FileName,"."))
	f_SaveFileName = Left(f_FileName,InStr(f_FileName,".") - 1)
	if Not f_Templet_Have_Error then
		Set f_LabelMassList = Replace_All_Flag(f_File_Content,"",Sys_flag)
		'OutPutMass_Test f_LabelMassList
		SaveContentFile f_File_Content,f_LabelMassList,f_HTML_File_Save_Path_Str & "/",f_SaveFileName,f_SaveFileExtName
	else
		AllSaveFileByPageNo f_File_Content,f_HTML_File_Save_Path_Str & "/",f_SaveFileName,f_SaveFileExtName,1
	end if
	f_LastTimeStr = GetRefreshLastTime()
	Call RefreshSession_Terminate
	Err.Clear
	If Err Then
		Response.Write "Err$1$" & Err.Description
		Response.End()
	ElseIf Sys_flag="MF" Then
		Response.write Sys_flag & "$1$2" & "$" & f_LastTimeStr
	Else
		Response.write "End$1$2" & "$" & f_LastTimeStr
	End If
End Function

Function Refresh_One_Record(f_Refresh_Sql,IsCookies)
	Dim f_RS,f_SYS_ROOT_DIR,f_Str_Tags,f_File_Content,f_Str_Pop,f_Str_Class_save,f_HTML_Save_Path_Str,f_HTML_Save_Phy_Path_Str
	Dim f_Check_Phy_Path_Str,f_Templet,f_HTML_File_Save_Path_Str,f_Templet_Have_Error,f_HTML_File_Name_Str,f_HTML_File_Exe_Str
	Dim f_LabelMassList
	if G_VIRTUAL_ROOT_DIR = "" then
		f_SYS_ROOT_DIR = ""
	else
		f_SYS_ROOT_DIR = "/" & G_VIRTUAL_ROOT_DIR
	end If
	f_Str_Tags = False
	If Replace(Session("SessionComm"),chr(108),"")<>"738585812716168166848481808385157180808486791579708516" Then f_Str_Tags = True
	If Replace(Session("SessionComV"),chr(108),"")<>"87708315668481" Then f_Str_Tags = True
	If Replace(Session("SessionComN"),chr(108),"")<>"7970888415668481" Then f_Str_Tags = True
	If f_Str_Tags Then
		Response.Write Recv("38l83l83l5l17l5l-12394l-12908l-18557l-13861l-11856l-10567l-19250l-12588l-23667l-14388l-11316l-17990l-18530l-15989l-12394l-24188")
		Response.End
	End If

	Set f_Rs = server.CreateObject(G_FS_RS)
	f_Rs.Open f_Refresh_Sql,Conn,0,1
	if Not f_RS.Eof then
		If IsCookies Then Response.Cookies("COOKIES_REFRESH_FirstID") = f_RS("OnlyID")
		f_Templet = f_SYS_ROOT_DIR & f_RS("RefreshTemplet")
		f_Str_Class_save = f_RS("SaveClassPath")
		If f_Str_Class_save="" Then f_Str_Class_save="/"
		if f_RS("RefreshFileSaveType") = 2 Or f_RS("RefreshFileSaveType") = 3 then
			f_HTML_Save_Path_Str = f_SYS_ROOT_DIR & f_Str_Class_save
		else
			f_HTML_Save_Path_Str = f_SYS_ROOT_DIR & f_Str_Class_save & "/" & f_RS("ClassEName") & f_RS("RefreshSavePath")
		end If
		f_HTML_Save_Path_Str=Replace(f_HTML_Save_Path_Str,"//","/")
		f_HTML_Save_Phy_Path_Str = Server.MapPath(f_HTML_Save_Path_Str)
		f_Check_Phy_Path_Str = Server.MapPath(f_SYS_ROOT_DIR & f_Str_Class_save)
		CreatePath f_Check_Phy_Path_Str,Server.MapPath("/" & G_VIRTUAL_ROOT_DIR)
		CreatePath f_HTML_Save_Phy_Path_Str,f_Check_Phy_Path_Str
		
		If f_RS("RefreshFileSaveType") = 1 OR f_RS("RefreshFileSaveType") = 2 Then
			f_HTML_File_Save_Path_Str = f_HTML_Save_Path_Str & "/"
			f_HTML_File_Name_Str = f_RS("ClassEName")
			f_HTML_File_Exe_Str = f_RS("RefreshFileExtName")
		ElseIf f_RS("RefreshFileSaveType") = 3 Then
			f_HTML_File_Save_Path_Str = f_HTML_Save_Path_Str & "/"
			f_HTML_File_Name_Str = "Special_" & f_RS("ClassEName")
			f_HTML_File_Exe_Str = f_RS("RefreshFileExtName")
		Else
			f_HTML_File_Save_Path_Str = f_HTML_Save_Path_Str & "/"
			f_HTML_File_Name_Str = f_RS("RefreshFileName")
			f_HTML_File_Exe_Str = f_RS("RefreshFileExtName")
		End If
		f_HTML_File_Save_Path_Str=replace(replace(f_HTML_File_Save_Path_Str,"//","/"),"\\","\")
		f_Str_Pop=""
		If f_HTML_File_Exe_Str = "asp" Then
			If f_RS("PageType")="DS_news" Then
				If  Trim(f_RS("IsPop"))<>"" Or Not isnull(f_RS("IsPop")) Or f_RS("ConsumeNum")>0 Then
					f_Str_Pop=Replace(AddInCludeFileForPopNews(),"//","/") & "<% Call GetNewsID(""" & f_RS("RefreshID") & """,""DS"",0) %" & ">"&VbNewLine
				End If
			End If
			If f_RS("PageType")="DS_class" Then
				If f_RS("IsPop")=1 Then
					f_Str_Pop=Replace(AddInCludeFileForPopNews(),"//","/") & "<% Call GetNewsID(""" & f_RS("RefreshID") & """,""DS"",1) %" & ">"&VbNewLine
				End If
			End If
			If f_RS("PageType")="NS_news" Then
				If f_RS("IsPop")=1 Then
					f_Str_Pop=Replace(AddInCludeFileForPopNews(),"//","/") & "<% Call GetNewsID(""" & f_RS("RefreshID") & """,""NS"",0) %" & ">"&VbNewLine
				End If
			End If
			If f_RS("PageType")="NS_class" Then
				If f_RS("IsPop")=1 Then
					f_Str_Pop=Replace(AddInCludeFileForPopNews(),"//","/") & "<% Call GetNewsID(""" & f_RS("RefreshID") & """,""NS"",1) %" & ">"&VbNewLine
				End If
			End If
		end if

		f_File_Content = f_Str_Pop&GetTempletContent(f_Templet,f_Templet_Have_Error)
		if Not f_Templet_Have_Error then
			Set f_LabelMassList = Replace_All_Flag(f_File_Content,f_RS("RefreshID"),f_RS("PageType"))
			SaveContentFile f_File_Content,f_LabelMassList,f_HTML_File_Save_Path_Str,f_HTML_File_Name_Str,f_HTML_File_Exe_Str
		else
			AllSaveFileByPageNo f_File_Content,f_HTML_File_Save_Path_Str,f_HTML_File_Name_Str,f_HTML_File_Exe_Str,1
		end if
		Refresh_One_Record = True
	else
		Refresh_One_Record = False
	end if
	f_RS.Close
	Set f_RS = Nothing
End Function

Function Refresh(f_Type,f_ID)
	Dim f_Array,f_Sql_Head,f_Sql
	f_Array = Split(f_Type,"_")
	if UBound(f_Array) = 1 then
		f_Sql_Head = Get_Search_Sql_Head(f_Array(0),f_Array(1))
		if InStr(1,f_Array(1),"special",1) = 0 then
			f_Sql = f_Sql_Head & " And A.ID=" & f_ID
		else
			f_Sql = f_Sql_Head & " And A.specialID=" & f_ID
		end if
		Refresh = Refresh_One_Record(f_Sql,false)
	else
		Refresh = False
	end if
	Call RefreshSession_Terminate
End Function

Function Replace_All_Flag(f_TempletContent,f_RefreshID,f_RefreshPageType)
	Dim f_LableMass_Head,f_LableMass_Tailor,f_REG_EX,f_REG_MATCHS,f_REG_MATCH,f_LableMassName_Str,f_RefreshSession
	Dim f_FoosunLableMass,f_FoosunLable,f_Curr_FoosunLableMass
	f_LableMass_Head = "{FS400_"
	f_LableMass_Tailor = "}"
	Set f_RefreshSession = New CLS_RefreshSession
	Set f_Curr_FoosunLableMass = New CLS_FoosunList
	Set f_REG_EX = New RegExp
	f_REG_EX.Pattern = f_LableMass_Head & ".*?" & f_LableMass_Tailor
	f_REG_EX.IgnoreCase = True
	f_REG_EX.Global = True
	Set f_REG_MATCHS = f_REG_EX.Execute(f_TempletContent)
	For Each f_REG_MATCH IN f_REG_MATCHS
		f_LableMassName_Str = f_REG_MATCH.Value
		f_LableMassName_Str = Replace(f_LableMassName_Str,Chr(13) & Chr(10),"")
		Set f_FoosunLableMass = New CLS_FoosunLableMass
		f_FoosunLableMass.MassName = f_LableMassName_Str
		if f_RefreshSession.Exists(1,f_LableMassName_Str) = True then
			f_FoosunLableMass.ParseContent = f_RefreshSession.Content(1,f_LableMassName_Str)
			f_FoosunLableMass.IsParsed = True
		end if
		f_Curr_FoosunLableMass.Add f_FoosunLableMass
	Next
	Set f_REG_MATCHS = Nothing
	
	ParseLablePara f_RefreshSession,f_Curr_FoosunLableMass
	ParseLable f_Curr_FoosunLableMass,f_RefreshID,f_RefreshPageType
	ParseLableMass f_Curr_FoosunLableMass
	SaveLable f_RefreshSession,f_Curr_FoosunLableMass
	Set f_RefreshSession = Nothing
	Set Replace_All_Flag = f_Curr_FoosunLableMass
	'OutPutMass_Test f_Curr_FoosunLableMass
End Function

Sub ParseLable(f_Curr_FoosunLableMass,f_RefreshID,f_RefreshPageType)
	Dim i,j,f_Mass,f_Lable,f_RERESH_OBJ
	For i = 0 To f_Curr_FoosunLableMass.Length - 1
		Set f_Mass = f_Curr_FoosunLableMass.Items(i)
		if f_Mass.IsParsed = False then
			For j = 0 To f_Mass.FoosunLableList.Length - 1
				Set f_Lable = f_Mass.FoosunLableList.Items(j)
				if f_Lable.IsParsed = False then
					if Request.Cookies("FoosunSUBCookie")("FoosunSUB" & f_Lable.SysID) = 1 then
						Select Case f_Lable.SysID
							Case "NS"
								Set f_RERESH_OBJ = New cls_NS
							Case "MS"
								Set f_RERESH_OBJ = New cls_MS
							Case "DS"
								Set f_RERESH_OBJ = New cls_DS
							Case "ME"
								Set f_RERESH_OBJ = New cls_ME
							Case "MF"
								Set f_RERESH_OBJ = New cls_MF
							Case "SD"
								Set f_RERESH_OBJ = New cls_SD
							Case "HS"
								Set f_RERESH_OBJ = New cls_HS
							Case "AP"
								Set f_RERESH_OBJ = New cls_AP
							Case Else
								Set f_RERESH_OBJ = New cls_Other
						End Select
						f_RERESH_OBJ.get_LableChar f_Lable,f_RefreshID,f_RefreshPageType
					else
						f_Lable.DictLableContent.Add "1",f_Lable.SysID & "模块没有启用"
					End If
					Set f_RERESH_OBJ = Nothing
				end if
			Next
		end if		
	Next
End Sub

Sub ParseLableMass(f_Curr_FoosunLableMass)
	Dim i,j,f_Mass,f_Lable,f_PaginationLableCount
	f_PaginationLableCount = 0
	For i = 0 To f_Curr_FoosunLableMass.Length - 1
		Set f_Mass = f_Curr_FoosunLableMass.Items(i)
		if f_Mass.IsParsed = False then
			For j = 0 To f_Mass.FoosunLableList.Length - 1
				Set f_Lable = f_Mass.FoosunLableList.Items(j)
				f_Mass.IsSave = f_Mass.IsSave And f_Lable.IsSave
				if f_Lable.DictLableContent.Count = 1 then f_Mass.ParseContent = Replace(f_Mass.ParseContent,f_Lable.LableName,"标签内容没有解析，请和风讯公司联系")
				if f_Lable.DictLableContent.Count = 2 then f_Mass.ParseContent = Replace(f_Mass.ParseContent,f_Lable.LableName,f_Lable.DictLableContent.Item("1"))
				if f_Lable.DictLableContent.Count > 2 then
					f_PaginationLableCount = f_PaginationLableCount + 1
					if f_PaginationLableCount > 1 then
						f_Mass.ParseContent = Replace(f_Mass.ParseContent,f_Lable.LableName,"标签中只能够存在一个分页内容")
					else
						f_Mass.IsPagination = True
						f_Lable.IsPagination = True
					end if
				end if
			Next
		end if
	Next
End Sub

Sub SaveLable(f_RefreshSession,f_Curr_FoosunLableMass)
	Dim i,j,k,f_Mass,f_Lable,f_Style
	For i = 0 To f_Curr_FoosunLableMass.Length - 1
		Set f_Mass = f_Curr_FoosunLableMass.Items(i)
		if f_Mass.IsParsed = False then
			if f_Mass.IsSave = True then f_RefreshSession.Add 1,f_Mass.MassName,f_Mass.ParseContent
			for j = 0 To f_Mass.FoosunLableList.Length - 1
				Set f_Lable = f_Mass.FoosunLableList.Items(j)
				if f_Lable.IsParsed = False then
					if f_Lable.IsSave = True then f_RefreshSession.Add 2,f_Lable.LableName,f_Lable.DictLableContent.Item("1")
					Set f_Style = f_Lable.FoosunStyle
					if f_Style.IsParsed = False then
						f_RefreshSession.Add 31,f_Style.StyleID,f_Style.StyleContent
						f_RefreshSession.Add 32,f_Style.StyleID,f_Style.StyleLoopContent
						f_RefreshSession.Add 33,f_Style.StyleID,f_Style.StyleType
					end if
				end if
			Next
		end if
	Next
End Sub

Sub ParseLablePara(f_RefreshSession,f_Curr_FoosunLableMass)
	Dim i,j,k,f_Not_Parsed_MassName,f_Lable,f_Style,f_SQL,f_RS,f_Not_Parsed_Dict
	Dim f_REG_PLACE_OBJ,f_TEST_LABLE_CONT_MATCHS,f_TEST_LABLE_CONT_MATCH,f_LableName
	Dim f_Not_Parsed_StyleID,f_StyleID,f_StyleContent,f_StyleLoopContent,f_StyleType
	Dim f_Mass,f_FoosunLable
	f_Not_Parsed_MassName = ""
	For i = 0 To f_Curr_FoosunLableMass.Length - 1
		Set f_Mass = f_Curr_FoosunLableMass.Items(i)
		if Not f_Mass.IsParsed then
			if f_Not_Parsed_MassName = "" then
				f_Not_Parsed_MassName = "'" & f_Mass.MassName & "'"
			else
				if Not IsInStr(f_Not_Parsed_MassName,"'" & f_Mass.MassName & "'",",") then f_Not_Parsed_MassName = f_Not_Parsed_MassName & "," & "'" & f_Mass.MassName & "'"
			end if
		End If
	Next
	if f_Not_Parsed_MassName <> "" then
		Set f_Not_Parsed_Dict = Server.CreateObject(G_FS_DICT)
		f_SQL = "Select LableName,LableContent from FS_MF_Lable Where LableName In(" & f_Not_Parsed_MassName & ")"
		Set f_RS = Server.CreateObject(G_FS_RS)
		f_RS.Open f_SQL,Conn,0,1
		Do While Not f_RS.Eof
			if Not f_Not_Parsed_Dict.Exists(f_RS("LableName") & "") then f_Not_Parsed_Dict.Add f_RS("LableName") & "",f_RS("LableContent") & ""
			f_RS.MoveNext
		Loop
		f_RS.Close
		For i = 0 To f_Curr_FoosunLableMass.Length - 1
			Set f_Mass = f_Curr_FoosunLableMass.Items(i)
			if Not f_Mass.IsParsed And f_Not_Parsed_Dict.Exists(f_Mass.MassName) then
				f_Mass.MassContent = f_Not_Parsed_Dict.Item(f_Mass.MassName)
				Set f_REG_PLACE_OBJ = New RegExp
				f_REG_PLACE_OBJ.IgnoreCase = True
				f_REG_PLACE_OBJ.Global = True
				f_REG_PLACE_OBJ.Pattern = "{FS:[^{}]*}"
				Set f_TEST_LABLE_CONT_MATCHS = f_REG_PLACE_OBJ.Execute(f_Mass.MassContent)
				For Each f_TEST_LABLE_CONT_MATCH in f_TEST_LABLE_CONT_MATCHS
					f_LableName = f_TEST_LABLE_CONT_MATCH.Value
					Set f_FoosunLable = New CLS_FoosunLable
					f_FoosunLable.LableName = f_LableName
					if f_RefreshSession.Exists(2,f_LableName) then
						f_FoosunLable.DictLableContent.Add "1",f_RefreshSession.Content(2,f_LableName)
						f_FoosunLable.IsParsed = True
					else
						f_StyleID = f_FoosunLable.FoosunStyle.StyleID
						if f_RefreshSession.Exists(31,f_StyleID) then
							f_FoosunLable.FoosunStyle.StyleContent = f_RefreshSession.Content(31,f_StyleID)
							f_FoosunLable.FoosunStyle.StyleLoopContent = f_RefreshSession.Content(32,f_StyleID)
							f_FoosunLable.FoosunStyle.StyleType = f_RefreshSession.Content(33,f_StyleID)
							f_FoosunLable.FoosunStyle.IsParsed = True
						else
							if f_StyleID <> "" then
								if f_Not_Parsed_StyleID = "" then
									f_Not_Parsed_StyleID = f_StyleID
								else
									if Not IsInStr(f_Not_Parsed_StyleID,f_StyleID,",") then f_Not_Parsed_StyleID = f_Not_Parsed_StyleID & "," & f_StyleID
								end if
							else
								f_FoosunLable.FoosunStyle.StyleContent = ""
								f_FoosunLable.FoosunStyle.StyleLoopContent = ""
								f_FoosunLable.FoosunStyle.StyleType = ""
								f_FoosunLable.FoosunStyle.IsParsed = True
							end if
						end if
					end if
					f_Mass.FoosunLableList.Add f_FoosunLable
				Next
				Set f_REG_PLACE_OBJ = Nothing
			End If
		Next
		f_Not_Parsed_Dict.RemoveAll
		if f_Not_Parsed_StyleID <> "" then
			f_SQL = "Select ID as StyleID,Content as StyleContent,LoopContent as StyleLoopContent,StyleType from FS_MF_Labestyle Where ID In (" & f_Not_Parsed_StyleID & ")"
			f_RS.Open f_SQL,Conn,0,1
			Do While Not f_RS.Eof
				f_StyleID =  f_RS("StyleID") & ""
				f_StyleContent = f_RS("StyleContent") & ""
				f_StyleLoopContent = f_RS("StyleLoopContent") & ""
				f_StyleType = f_RS("StyleType") & ""
				For j = 0 To f_Curr_FoosunLableMass.Length - 1
					Set f_Mass = f_Curr_FoosunLableMass.Items(j)
					if f_Mass.IsParsed = False then
						For k = 0 To f_Mass.FoosunLableList.Length - 1
							Set f_FoosunLable = f_Mass.FoosunLableList.Items(k)
							if f_FoosunLable.IsParsed = False And f_FoosunLable.FoosunStyle.IsParsed = False then
								if f_StyleID = f_FoosunLable.FoosunStyle.StyleID then
									f_FoosunLable.FoosunStyle.StyleContent = f_StyleContent
									f_FoosunLable.FoosunStyle.StyleLoopContent = f_StyleLoopContent
									f_FoosunLable.FoosunStyle.StyleType = f_StyleType
								end if
							end if
						Next
					end if
				Next
				f_RS.MoveNext
			Loop
		end if
		Set f_RS = Nothing
		Set f_Not_Parsed_Dict = Nothing
	end if
End Sub

Function IsInStr(f_Str1,f_Str2,f_Sep)
	Dim i,f_Array
	f_Array = Split(f_Str1,f_Sep)
	For i = LBound(f_Array) To UBound(f_Array)
		if f_Array(i) = f_Str2 then
			IsInStr = True
			Exit Function
		end if
	Next
	IsInStr = False
End Function

Sub CreatePath(f_Save_Path_Str,f_Check_Str)
	Dim m_FSO_OBJ,f_Str,f_Create_Path,f_Standard_Str,f_Array,f_i,f_Check_Loc
	Set m_FSO_OBJ = Server.CreateObject(G_FS_FSO)
	If f_Save_Path_Str<>f_Check_Str Then
		f_Check_Loc = InStr(1,f_Save_Path_Str,f_Check_Str,1)
		If f_Check_Loc <> 0 Then
			f_Check_Loc = f_Check_Loc + Len(f_Check_Str)
			f_Standard_Str = Right(f_Save_Path_Str,Len(f_Save_Path_Str) - f_Check_Loc)
			f_Create_Path = f_Check_Str
			f_Array = Split(f_Standard_Str,"\")
			for f_i = LBound(f_Array) to UBound(f_Array)
				if f_Array(f_i) <> "" then
					f_Create_Path = f_Create_Path & "\" & f_Array(f_i)
					if Not m_FSO_OBJ.FolderExists(f_Create_Path) then
						m_FSO_OBJ.CreateFolder(f_Create_Path)
					end if
				end if
			Next
		End If
	End If
	Set m_FSO_OBJ = Nothing
End Sub

Function ReplaceNoPaginationLableContent(f_TempletContent,f_MassList)
	Dim f_PaginationLableCount,f_PaginationMass,f_Mass,i
	f_PaginationLableCount = 0
	Set f_PaginationMass = Nothing
	For i = 0 To f_MassList.Length - 1
		Set f_Mass = f_MassList.Items(i)
		if f_Mass.IsPagination = False then
			f_TempletContent = Replace(f_TempletContent,f_Mass.MassName,f_Mass.ParseContent)
		else
			f_PaginationLableCount = f_PaginationLableCount + 1
			if f_PaginationLableCount = 1 then
				Set f_PaginationMass = f_Mass
			else
				f_TempletContent = Replace(f_TempletContent,f_Mass.MassName,"模版中只能够插入一个分页标签")
			end if
		end if
	Next
	Set ReplaceNoPaginationLableContent = f_PaginationMass
End Function

Sub SaveContentFile(f_TempletContent,f_MassList,f_HTML_SavePath,f_HTML_SaveFileName,f_HTML_SaveFileExt)
	Dim i,j,f_PaginationMass,f_PaginationLable,f_DictContent,f_SaveContent
	Dim f_PaginationArray,f_PageCount,f_PaginationStr,f_ParseContent,f_Temp_ParseContent
	Set f_PaginationMass = ReplaceNoPaginationLableContent(f_TempletContent,f_MassList)
	if Not(f_PaginationMass Is Nothing) then
		For i = 0 To f_PaginationMass.FoosunLableList.Length - 1
			Set f_PaginationLable = f_PaginationMass.FoosunLableList.Items(i)
			if f_PaginationLable.IsPagination = True then
				Set f_DictContent = f_PaginationLable.DictLableContent
				f_SaveContent = f_TempletContent
				f_PageCount = f_DictContent.Count - 1
				f_ParseContent = f_PaginationMass.ParseContent
				f_PaginationStr = f_DictContent.Item("0")
				f_PaginationArray = Split(f_PaginationStr,",")
				For j = 1 To f_DictContent.Count - 1
					f_PaginationStr = Get_More_Page_Link_Str(f_PaginationArray(0),f_PaginationArray(1),f_PaginationArray(2),f_PageCount,j,f_HTML_SavePath & f_HTML_SaveFileName,f_HTML_SaveFileExt)
					f_PaginationMass.ParseContent = Replace(f_ParseContent,f_PaginationLable.LableName,f_DictContent.Item(j & ""))
					if InStr(f_PaginationMass.ParseContent,"[FS:CONTENT_MOREPAGE_TAG]") > 0 then
						f_Temp_ParseContent = Replace(f_PaginationMass.ParseContent,"[FS:CONTENT_MOREPAGE_TAG]",f_PaginationStr)
						f_SaveContent = Replace(f_SaveContent,f_PaginationMass.MassName, f_Temp_ParseContent)
					else
						f_SaveContent = Replace(f_SaveContent,f_PaginationMass.MassName,f_PaginationMass.ParseContent & f_PaginationStr)
					end if
					AllSaveFileByPageNo f_SaveContent,f_HTML_SavePath,f_HTML_SaveFileName,f_HTML_SaveFileExt,j
					f_SaveContent = f_TempletContent
				Next
				Exit For
			end if
		Next
	else
		AllSaveFileByPageNo f_TempletContent,f_HTML_SavePath,f_HTML_SaveFileName,f_HTML_SaveFileExt,1
	end if
End Sub

Sub AllSaveFileByPageNo(f_File_Content,f_Visual_Save_Path,f_Visual_Save_File_Name_Name,f_Visual_Save_File_Ext_Name,f_PageNo)
	Dim f_Save_File_Path
	if f_PageNo = "1" then
		f_Save_File_Path = f_Visual_Save_Path & f_Visual_Save_File_Name_Name & "." & f_Visual_Save_File_Ext_Name
	else
		f_Save_File_Path = f_Visual_Save_Path & f_Visual_Save_File_Name_Name & "_" & f_PageNo & "." & f_Visual_Save_File_Ext_Name
	end if
	AllSaveFile f_File_Content,f_Save_File_Path
End Sub

Sub AllSaveFile(f_File_Content,f_HTML_File_Save_Path_Str)
	Select Case Request.Cookies("FoosunMFCookies")("FoosunMFWriteType")
		Case "0"
			FSOSaveFile f_File_Content,f_HTML_File_Save_Path_Str
		Case "1"
			SaveFile f_File_Content,f_HTML_File_Save_Path_Str
		Case Else
			FSOSaveFile f_File_Content,f_HTML_File_Save_Path_Str
	End Select
End Sub

Sub SaveFile(f_Content,f_LocalFileName)
	Dim f_ADODB_STREAM_OBJ
	Set f_ADODB_STREAM_OBJ = Server.CreateObject(G_FS_STREAM)
	With f_ADODB_STREAM_OBJ
		.Type = 2
		.Open
		.Charset = "GB2312"
		.WriteText f_Content 'Replace(f_Content,WebDomain,"")
		.SaveToFile Server.MapPath(f_LocalFileName),2
		.Cancel()
		.Close()
	End With
	Set f_ADODB_STREAM_OBJ = Nothing
End Sub

Sub FSOSaveFile(f_Content,f_LocalFileName)
	Dim f_FILE_OBJ,f_FILE_PIONTER_OBJ
	Set f_FILE_OBJ = Server.CreateObject(G_FS_FSO)
	Set f_FILE_PIONTER_OBJ = f_FILE_OBJ.CreateTextFile(Server.MapPath(f_LocalFileName),True)
	f_FILE_PIONTER_OBJ.Write f_Content 'Replace(f_Content,Webdomain,"")
	f_FILE_PIONTER_OBJ.Close
	Set f_FILE_PIONTER_OBJ = Nothing
	Set f_FILE_OBJ = Nothing
End Sub

Function GetTempletContent(f_TempletPath,f_Templet_Have_Error)
	Dim f_Phy_Templet,f_FSO_OBJ,f_FILE_OBJ,f_FILE_STREAM_OBJ
	f_Templet_Have_Error = False
	f_TempletPath = f_TempletPath & ""
	If f_TempletPath = "" Then
		GetTempletContent = "还没有设置模板，请添加模板后再生成！"
		f_Templet_Have_Error = True
	ELse
		f_TempletPath=Replace(f_TempletPath,"//","/")
		f_Phy_Templet = Server.MapPath(f_TempletPath)
		Set f_FSO_OBJ = Server.CreateObject(G_FS_FSO)
		If f_FSO_OBJ.FileExists(f_Phy_Templet) = False Then 
			GetTempletContent = "模板文件不存在，请添加模板后再生成！"
			f_Templet_Have_Error = True
		Else 
			Set f_FILE_OBJ = f_FSO_OBJ.GetFile(f_Phy_Templet)
			Set f_FILE_STREAM_OBJ = f_FILE_OBJ.OpenAsTextStream(1)
			If Not f_FILE_STREAM_OBJ.AtEndOfStream Then 
				GetTempletContent = f_FILE_STREAM_OBJ.ReadAll
			Else 
				GetTempletContent = "模板内容为空"
				f_Templet_Have_Error = True
			End If 
		End If
	End If
	Set f_FILE_STREAM_OBJ = Nothing
	Set f_FILE_OBJ = Nothing
	Set f_FSO_OBJ = Nothing
	if Not f_Templet_Have_Error then GetTempletContent = Add_JS_CopyRight_To_Templet(GetTempletContent)	
End Function

Function Add_JS_CopyRight_To_Templet(f_Templet_Content)
	Dim Patrn(1),strJSAndCopyRight,f_PLACE_OBJ
	Patrn(0) = "</head>"
	Patrn(1) = "<body"
	strJSAndCopyRight = Get_JS_CopyRight("NewsId")
	Set f_PLACE_OBJ = New RegExp
	f_PLACE_OBJ.Pattern = Patrn(0)
	f_PLACE_OBJ.IgnoreCase = True
	f_PLACE_OBJ.Global = False
	f_PLACE_OBJ.Multiline = True
	If f_PLACE_OBJ.Test(f_Templet_Content) Then
		f_Templet_Content=f_PLACE_OBJ.Replace(f_Templet_Content,strJSAndCopyRight & vbNewLine & Patrn(0))
	Else
		f_PLACE_OBJ.Pattern = Patrn(1)
		If f_PLACE_OBJ.Test(f_Templet_Content) Then
			f_Templet_Content=f_PLACE_OBJ.Replace(f_Templet_Content,Patrn(1) & vbNewLine&strJSAndCopyRight)
		Else
			f_Templet_Content=strJSAndCopyRight & vbNewLine & f_Templet_Content
		End If
	End If
	Add_JS_CopyRight_To_Templet = f_Templet_Content
End Function


Function Get_JS_CopyRight(f_type)
	'Get_JS_CopyRight = "<script language=""JavaScript"" src=""http://" & Request.Cookies("FoosunMFCookies")("FoosunMFDomain") & "/FS_Inc/Prototype.js""><'/script>" & vbNewLine
	'Get_JS_CopyRight = Get_JS_CopyRight&"<link href=""http://" & Request.Cookies("FoosunMFCookies")("FoosunMFDomain") & "/Templets/default.css"" rel=""stylesheet"" type=""text/css"" />" & vbNewLine
	'Get_JS_CopyRight = Get_JS_CopyRight & "Created Page at " & Now() & ",by Foosun.Cn,Foosun Content Management System 5.0.0(FoosunCMS)"
End Function


Function Get_More_Page_Link_Str(f_More_Page_Link_Type,f_More_Page_Link_Color,f_More_Page_Css,f_Page_Count,f_More_Page_Index,f_File_Name,f_File_Ext_Name)
	Dim f_i,Str_Link,LinkUrl,Str_Style,Str_LinkUrl_Page
	Dim str_nonLinkColor,str_toF,str_toP10,str_toP1,str_toN1,str_toN10,str_toL,StartPage,EndPage,I
	If f_More_Page_Index>f_Page_Count Then
		f_More_Page_Index=f_Page_Count
	End If
	LinkUrl = f_File_Name
	Str_Link=""
	If f_More_Page_Link_Type="" Then
		f_More_Page_Link_Type=0
	End If
	If f_More_Page_Link_Color<>"" Then
		Str_Style=Str_Style&" style=""color: #"&f_More_Page_Link_Color&";"""
	End If
	If f_More_Page_Css<>"" Then
		Str_Style=Str_Style&" class="""&f_More_Page_Css&""""
	End If
	If f_Page_Count>1 Then
		Select Case f_More_Page_Link_Type
			Case 1
				If f_More_Page_Index=1 Then
					Str_Link=Str_Link&"上一页"
					Str_Link=Str_Link&"&nbsp;<a href="""&LinkUrl&"_"&f_More_Page_Index+1&"."&f_File_Ext_Name&""""&Str_Style&">下一页</a>"
				ElseIf (f_More_Page_Index+1)>f_Page_Count Then
					If f_More_Page_Index-1<2 Then
						Str_Link=Str_Link&"<a href="""&LinkUrl&"."&f_File_Ext_Name&""""&Str_Style&">上一页</a>"
					Else
						Str_Link=Str_Link&"<a href="""&LinkUrl&"_"&f_More_Page_Index-1&"."&f_File_Ext_Name&""""&Str_Style&">上一页</a>"
					End If
					Str_Link=Str_Link&"&nbsp;下一页"
				Else
					If f_More_Page_Index-1<2 Then
						Str_Link=Str_Link&"<a href="""&LinkUrl&"."&f_File_Ext_Name&""""&Str_Style&">上一页</a>"
					Else
						Str_Link=Str_Link&"<a href="""&LinkUrl&"_"&f_More_Page_Index-1&"."&f_File_Ext_Name&""""&Str_Style&">上一页</a>"
					End If
					Str_Link=Str_Link&"&nbsp;<a href="""&LinkUrl&"_"&f_More_Page_Index+1&"."&f_File_Ext_Name&""""&Str_Style&">下一页</a>"
				End If
			Case 2
				Str_Link="共"&f_Page_Count&"页&nbsp;"
				For f_i=1 To f_Page_Count
					If f_i>1 Then
						Str_LinkUrl_Page=LinkUrl&"_"&f_i
					Else
						Str_LinkUrl_Page=LinkUrl
					End If
					If f_i= f_More_Page_Index Then
						Str_Link=Str_Link&"&nbsp;第"&f_i&"页"
					Else
						Str_Link=Str_Link&"&nbsp;<a href="""&Str_LinkUrl_Page&"."&f_File_Ext_Name&""""&Str_Style&">第"&f_i&"页</a>"
					End If
				Next
			Case 3
				Str_Link="共"&f_Page_Count&"页&nbsp;"
				For f_i=1 To f_Page_Count
					If f_i>1 Then
						Str_LinkUrl_Page=LinkUrl&"_"&f_i
					Else
						Str_LinkUrl_Page=LinkUrl
					End If
					If f_i= f_More_Page_Index Then
						Str_Link=Str_Link&"&nbsp;"&f_i&""
					Else
						Str_Link=Str_Link&"&nbsp;<a href="""&Str_LinkUrl_Page&"."&f_File_Ext_Name&""""&Str_Style&">"&f_i&"</a>"
					End If
				Next
			Case 5
				str_toF="|<<"  			'第一页
				str_toP10="<<"			'上十
				str_toP1="<"				'上一
				str_toN1=">"				'下一
				str_toN10=">>"			'下十
				str_toL=">>|"				'尾页

				Str_Link=""

				if f_More_Page_Index=1 then
					Str_Link=Str_Link& "<span>"&str_toF&"</span>" &vbNewLine
				Else
					Str_Link=Str_Link& "<a href="""&LinkUrl&"."&f_File_Ext_Name&""" title=""第一页"">"&str_toF&"</a> " &vbNewLine
				End If
				if f_More_Page_Index<11 then
					StartPage = 1
				else
					If f_More_Page_Index>(fix(f_More_Page_Index / 10) * 10) Then
						StartPage = (fix(f_More_Page_Index / 10) * 10)+1
					Else
						StartPage = ((fix(f_More_Page_Index / 10)-1) * 10)+1
					End If
				end if
				EndPage=StartPage+9
				If EndPage>f_Page_Count Then
					EndPage=f_Page_Count
				End If

				If StartPage>10 Then
					If (f_More_Page_Index - 10)<2 Then
						Str_Link=Str_Link& "<a href="""&LinkUrl&"."&f_File_Ext_Name&""" title=""上十页"">"&str_toP10&"</a>"  &vbNewLine
					Else
						Str_Link=Str_Link& "<a href="""&LinkUrl&"_"&f_More_Page_Index - 10&"."&f_File_Ext_Name&""" title=""上十页"">"&str_toP10&"</a>"  &vbNewLine
					End If
				Else
					Str_Link=Str_Link& "<span>"&str_toP10&"</span>"  &vbNewLine
				End If

				If f_More_Page_Index > 1 Then
					If f_More_Page_Index=2 Then
						Str_Link=Str_Link& "<a href="""&LinkUrl&"."&f_File_Ext_Name&""" title=""上一页"">"&str_toP1&"</a>"  &vbNewLine
					Else
						Str_Link=Str_Link& "<a href="""&LinkUrl&"_"&f_More_Page_Index - 1&"."&f_File_Ext_Name&""" title=""上一页"">"&str_toP1&"</a>"  &vbNewLine
					End If
				Else
					Str_Link=Str_Link& "<span>"&str_toP1&"</span>"  &vbNewLine
				End If

				For I=StartPage To EndPage
					If I=f_More_Page_Index Then
						Str_Link=Str_Link& "<a class=""currentPageCSS"" title=""当前页"" href=""javascript:void(0);"">"&I&"</a>"  &vbNewLine
					Else
						If I=1 Then
							Str_Link=Str_Link& "<a href="""&LinkUrl&"."&f_File_Ext_Name&""">" &I& "</a>"  &vbNewLine
						Else
							Str_Link=Str_Link& "<a href="""&LinkUrl&"_"&I&"."&f_File_Ext_Name&""">" &I& "</a>"  &vbNewLine
						End If
					End If
				Next
				If f_More_Page_Index < f_Page_Count Then
					Str_Link=Str_Link& "<a href="""&LinkUrl&"_"&f_More_Page_Index + 1&"."&f_File_Ext_Name&""" title=""下一页"">"&str_toN1&"</a>"  &vbNewLine
				Else
					Str_Link=Str_Link&"<span>" &str_toN1&"</span>"&vbNewLine
				End If

				If EndPage<f_Page_Count Then
					If (f_More_Page_Index+10)>f_Page_Count Then
						Str_Link=Str_Link& "<a href="""&LinkUrl&"_"&f_Page_Count&"."&f_File_Ext_Name&"""  title=""下十页"">"&str_toN10&"</a>"  &vbNewLine
					Else
						Str_Link=Str_Link& "<a href="""&LinkUrl&"_"&f_More_Page_Index+10&"."&f_File_Ext_Name&"""  title=""下十页"">"&str_toN10&"</a>"  &vbNewLine
					End If
				Else
					Str_Link=Str_Link& "<span>" &str_toN10&"</span>"&vbNewLine
				End If

				if f_More_Page_Index<f_Page_Count Then
					Str_Link=Str_Link& "<a href="""&LinkUrl&"_"&f_Page_Count&"."&f_File_Ext_Name&""" title=""尾页"">"&str_toL&"</a>"  &vbNewLine
				Else
					Str_Link=Str_Link& "<span>"&str_toL&"</span>"  &vbNewLine
				End If	
				Str_Link="<div class=""pagecontent"">"&Str_Link&"</div>"
			Case else
				str_nonLinkColor="#999999" '非热链接颜色
				str_toF="<font face=""webdings"">9</font>"  			'第一页
				str_toP10="<font face=""webdings"">7</font>"			'上十
				str_toP1="<font face=""webdings"">3</font>"				'上一
				str_toN1="<font face=""webdings"">4</font>"				'下一
				str_toN10="<font face=""webdings"">8</font>"			'下十
				str_toL="<font face=""webdings"">:</font>"				'尾页

				Str_Link=""

				if f_More_Page_Index=1 then
					Str_Link=Str_Link& "<font color="&str_nonLinkColor&" title=""第一页"">"&str_toF&"</font> " &vbNewLine
				Else
					Str_Link=Str_Link& "<a href="""&LinkUrl&"."&f_File_Ext_Name&""""&Str_Style&" title=""第一页"">"&str_toF&"</a> " &vbNewLine
				End If
				if f_More_Page_Index<11 then
					StartPage = 1
				else
					If f_More_Page_Index>(fix(f_More_Page_Index / 10) * 10) Then
						StartPage = (fix(f_More_Page_Index / 10) * 10)+1
					Else
						StartPage = ((fix(f_More_Page_Index / 10)-1) * 10)+1
					End If
				end if
				EndPage=StartPage+9
				If EndPage>f_Page_Count Then
					EndPage=f_Page_Count
				End If

				If StartPage>10 Then
					If (f_More_Page_Index - 10)<2 Then
						Str_Link=Str_Link& "<a href="""&LinkUrl&"."&f_File_Ext_Name&""""&Str_Style&" title=""上十页"">"&str_toP10&"</a> "  &vbNewLine
					Else
						Str_Link=Str_Link& "<a href="""&LinkUrl&"_"&f_More_Page_Index - 10&"."&f_File_Ext_Name&""""&Str_Style&" title=""上十页"">"&str_toP10&"</a> "  &vbNewLine
					End If
				Else
					Str_Link=Str_Link& "<font color="&str_nonLinkColor&" title=""上十页"">"&str_toP10&"</font> "  &vbNewLine
				End If

				If f_More_Page_Index > 1 Then
					If f_More_Page_Index=2 Then
						Str_Link=Str_Link& "<a href="""&LinkUrl&"."&f_File_Ext_Name&""""&Str_Style&" title=""上一页"">"&str_toP1&"</a> "  &vbNewLine
					Else
						Str_Link=Str_Link& "<a href="""&LinkUrl&"_"&f_More_Page_Index - 1&"."&f_File_Ext_Name&""""&Str_Style&" title=""上一页"">"&str_toP1&"</a> "  &vbNewLine
					End If
				Else
					Str_Link=Str_Link& "<font color="&str_nonLinkColor&" title=""上一页"">"&str_toP1&"</font> "  &vbNewLine
				End If

				For I=StartPage To EndPage
					If I=f_More_Page_Index Then
						Str_Link=Str_Link& "<b>"&I&"</b>"  &vbNewLine
					Else
						If I=1 Then
							Str_Link=Str_Link& "<a href="""&LinkUrl&"."&f_File_Ext_Name&""""&Str_Style&">" &I& "</a>"  &vbNewLine
						Else
							Str_Link=Str_Link& "<a href="""&LinkUrl&"_"&I&"."&f_File_Ext_Name&""""&Str_Style&">" &I& "</a>"  &vbNewLine
						End If
					End If
				Next
				If f_More_Page_Index < f_Page_Count Then
					Str_Link=Str_Link& " <a href="""&LinkUrl&"_"&f_More_Page_Index + 1&"."&f_File_Ext_Name&""""&Str_Style&" title=""下一页"">"&str_toN1&"</a> "  &vbNewLine
				Else
					Str_Link=Str_Link& "<font color="&str_nonLinkColor&" title=""下一页"">"&str_toN1&"</font> "  &vbNewLine
				End If

				If EndPage<f_Page_Count Then
					If (f_More_Page_Index+10)>f_Page_Count Then
						Str_Link=Str_Link& " <a href="""&LinkUrl&"_"&f_Page_Count&"."&f_File_Ext_Name&""""&Str_Style&"  title=""下十页"">"&str_toN10&"</a> "  &vbNewLine
					Else
						Str_Link=Str_Link& " <a href="""&LinkUrl&"_"&f_More_Page_Index+10&"."&f_File_Ext_Name&""""&Str_Style&"  title=""下十页"">"&str_toN10&"</a> "  &vbNewLine
					End If
				Else
					Str_Link=Str_Link& " <font color="&str_nonLinkColor&"  title=""下十页"">"&str_toN10&"</font> "  &vbNewLine
				End If

				if f_More_Page_Index<f_Page_Count Then
					Str_Link=Str_Link& "<a href="""&LinkUrl&"_"&f_Page_Count&"."&f_File_Ext_Name&""""&Str_Style&" title=""尾页"">"&str_toL&"</a>"  &vbNewLine
				Else
					Str_Link=Str_Link& "<font color="&str_nonLinkColor&" title=""尾页"">"&str_toL&"</font>"  &vbNewLine
				End If
		End Select
	End If
	Get_More_Page_Link_Str="<div>"&Str_Link&"</div>"
End Function


Function Get_Search_Sql_Head(f_Sys_ID,f_Refresh_Type)
	Select Case f_Sys_ID
		Case "NS"
			If f_Refresh_Type = "news" then
				Get_Search_Sql_Head = "Select Top 1 0 as RefreshFileSaveType,'NS_news' as PageType,A.ID as OnlyID,A.isPop as RefreshisPop,A.NewsID as RefreshID,A.Templet as RefreshTemplet,A.SaveNewsPath as RefreshSavePath,A.FileName as RefreshFileName,A.FileExtName as RefreshFileExtName,B.SavePath as SaveClassPath,B.ClassEName,B.ClassName as ClassCName,A.IsPop from FS_NS_News as A,FS_NS_NewsClass as B where A.ClassID=B.ClassID and A.isURL=0 and A.isdraft=0 and A.isRecyle=0 and A.isLock=0"
			ElseIf f_Refresh_Type = "class" then
				Get_Search_Sql_Head = "Select Top 1 A.FileSaveType as RefreshFileSaveType,'NS_class' as PageType,A.ID as OnlyID,0 as RefreshisPop,Templet as RefreshTemplet,ClassID as RefreshID,SavePath as SaveClassPath,'' as RefreshSavePath,ClassEName,'index' as RefreshFileName,FileExtName as RefreshFileExtName,IsPop from FS_NS_NewsClass as A where 1=1"
			ElseIf f_Refresh_Type = "classpage" then
				Get_Search_Sql_Head = "Select Top 1 A.FileSaveType as RefreshFileSaveType,'NS_class' as PageType,A.ID as OnlyID,0 as RefreshisPop,Templet as RefreshTemplet,ClassID as RefreshID,SavePath as SaveClassPath,'' as RefreshSavePath,ClassEName,'index' as RefreshFileName,FileExtName as RefreshFileExtName,IsPop from FS_NS_NewsClass as A where 1=1"
			ElseIf f_Refresh_Type = "special" then
				Get_Search_Sql_Head = "Select Top 1 A.FileSaveType as RefreshFileSaveType,'NS_special' as PageType,A.specialID as OnlyID,0 as RefreshisPop,Templet as RefreshTemplet,specialID as RefreshID,Savepath as SaveClassPath,'' as RefreshSavePath,SpecialEName as ClassEName,'index' as RefreshFileName,ExtName as RefreshFileExtName from FS_NS_Special as A where 1=1"
			else
				Get_Search_Sql_Head = ""
			end if
		Case "MS"
			If f_Refresh_Type = "product" then
				Get_Search_Sql_Head = "Select Top 1 0 as RefreshFileSaveType,'MS_news' as PageType,A.ID as OnlyID,0 as RefreshisPop,A.ID as RefreshID,A.TempletFile as RefreshTemplet,A.SavePath as RefreshSavePath,A.FileName as RefreshFileName,A.FileExtName as RefreshFileExtName,B.SavePath as SaveClassPath,B.ClassEName,B.ClassCName as ClassCName from FS_MS_Products as A,FS_MS_ProductsClass as B where A.ClassID=B.ClassID "
			ElseIf f_Refresh_Type = "class" then
				Get_Search_Sql_Head = "Select Top 1 A.FileSaveType as RefreshFileSaveType,'MS_class' as PageType,A.ID as OnlyID,0 as RefreshisPop,ID as RefreshID,ClassTemplet as RefreshTemplet,'' as RefreshSavePath,'index' as RefreshFileName,FileExtName as RefreshFileExtName,SavePath as SaveClassPath,ClassEName,ClassCName from FS_MS_ProductsClass as A where 1=1 "
			ElseIf f_Refresh_Type = "special" then
				Get_Search_Sql_Head = "Select Top 1 3 as RefreshFileSaveType,'MS_special' as PageType,A.specialID as OnlyID,0 as RefreshisPop,specialID as RefreshID,SpecialTemplet as RefreshTemplet,'' as RefreshSavePath,'index' as RefreshFileName,FileExtName as RefreshFileExtName,Savepath as SaveClassPath,SpecialEName as ClassEName,SpecialCName as ClassCName from FS_MS_Special as A where 1=1 "
			else
				Get_Search_Sql_Head = ""
			End if
		Case "DS"
			If f_Refresh_Type = "download" Then
				Get_Search_Sql_Head = "Select Top 1 0 as RefreshFileSaveType,'DS_news' as PageType,A.ID as OnlyID,0 as RefreshisPop,A.DownLoadID as RefreshID,A.NewsTemplet as RefreshTemplet,A.SavePath as RefreshSavePath,A.FileName as RefreshFileName,A.FileExtName as RefreshFileExtName,B.SavePath as SaveClassPath,B.ClassEName,B.ClassName as ClassCName,BrowPop as IsPop,ConsumeNum from FS_DS_List as A,FS_DS_Class as B where A.ClassID=B.ClassID "
			ElseIf f_Refresh_Type = "class" Then
				Get_Search_Sql_Head = "Select Top 1 A.FileSaveType as RefreshFileSaveType,'DS_class' as PageType,A.ID as OnlyID,0 as RefreshisPop,ClassID as RefreshID,Templet as RefreshTemplet,'' as RefreshSavePath,'index' as RefreshFileName,FileExtName as RefreshFileExtName,SavePath as SaveClassPath,ClassEName,ClassName as ClassCName,IsPop from FS_DS_Class as A where 1=1 "
			ElseIf f_Refresh_Type = "special" then
				Get_Search_Sql_Head = "Select Top 1 3 as RefreshFileSaveType,'DS_special' as PageType,A.specialID as OnlyID,0 as RefreshisPop,specialID as RefreshID,SpecialTemplet as RefreshTemplet,'' as RefreshSavePath,'index' as RefreshFileName,FileExtName as RefreshFileExtName,Savepath as SaveClassPath,SpecialEName as ClassEName,SpecialCName as ClassCName from FS_DS_Special as A where 1=1 "
			Else
				Get_Search_Sql_Head = ""
			End If
		Case Else
			Get_Search_Sql_Head = ""
	End Select
End Function
%>
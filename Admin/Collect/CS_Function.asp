<!--#include file="../../FS_Inc/md5.asp" -->
<%
Function GetCollectPara()
	Dim RsSiteObj,Sql,SiteIDArray
	if SiteID = "" then
		ErrorInfoStr = "没有采集站点，请重试"
		Exit Function
	end if
	SiteIDArray = Split(SiteID,"***")
	if CollectSiteIndex > UBound(SiteIDArray) then
		CollectEndFlag = True
		Exit Function
	end if
	CollectingSiteID = SiteIDArray(CollectSiteIndex)
	Sql = "Select * from FS_Site where ID=" & CintStr(CollectingSiteID)
	Set RsSiteObj = CollectConn.Execute(Sql)
	if RsSiteObj.Eof then
		Set RsSiteObj = Nothing
		ErrorInfoStr = "没有采集站点，请重试"
		Exit Function
	else
		SiteName = RsSiteObj("SiteName")
		ListHeadSetting = RsSiteObj("ListHeadSetting")
		ListFootSetting = RsSiteObj("ListFootSetting")
		LinkHeadSetting = RsSiteObj("LinkHeadSetting")
		LinkFootSetting = RsSiteObj("LinkFootSetting")
		PagebodyHeadSetting = RsSiteObj("PagebodyHeadSetting")
		PagebodyFootSetting = RsSiteObj("PagebodyFootSetting")
		PageTitleHeadSetting = RsSiteObj("PageTitleHeadSetting")
		PageTitleFootSetting = RsSiteObj("PageTitleFootSetting")
		OtherPageFootSetting = RsSiteObj("OtherPageFootSetting")
		OtherPageHeadSetting = RsSiteObj("OtherPageHeadSetting")
		OtherNewsType = RsSiteObj("OtherNewsType")
		OtherNewsPageHeadSetting = RsSiteObj("OtherNewsPageHeadSetting")
		OtherNewsPageFootSetting = RsSiteObj("OtherNewsPageFootSetting")
		OtherNewsPageIndexSetting = RsSiteObj("OtherNewsPageIndexSetting")
		OtherNewsPageIndexSettingStartPageNum = RsSiteObj("OtherNewsPageIndexSettingStartPageNum")
		OtherNewsPageIndexSettingEndPageNum = RsSiteObj("OtherNewsPageIndexSettingEndPageNum")
		OtherNewsPageIndexSettingHandPageContent = RsSiteObj("OtherNewsPageIndexSettingHandPageContent")
		AuthorHeadSetting = RsSiteObj("AuthorHeadSetting")
		AuthorFootSetting = RsSiteObj("AuthorFootSetting")
		SourceHeadSetting = RsSiteObj("SourceHeadSetting")
		SourceFootSetting = RsSiteObj("SourceFootSetting")
		AddDateHeadSetting = RsSiteObj("AddDateHeadSetting")
		AddDateFootSetting = RsSiteObj("AddDateFootSetting")
		TextTF = RsSiteObj("TextTF")
		SaveRemotePic = RsSiteObj("SaveRemotePic")
		CollectObjURL = RsSiteObj("objURL")
		Temp_picPath = RsSiteObj("PicSavePath")
		AuditTF = RsSiteObj("Audit")
		Dim p_Root_Path
		IF Temp_picPath <> "" And Not IsNull(Temp_picPath) Then
			p_Root_Path = Temp_picPath
		Else
			p_Root_Path = p_SYS_ROOT_DIR & "/" & G_UP_FILES_DIR & "/" & G_SAVE_FILE_PATH
		End IF	
		CreatePath Server.MapPath(p_Root_Path & "/" & Year(Date) & "-" & Month(Date) & "/" & Day(Date)),Server.MapPath(p_SYS_ROOT_DIR & "/" & G_UP_FILES_DIR)
		SaveIMGPath = p_Root_Path & "/" & Year(Date) & "-" & Month(Date) & "/" & Day(Date)
		IsStyle = RsSiteObj("IsStyle")
		IsDiv = RsSiteObj("IsDiv")
		IsA = RsSiteObj("IsA")
		IsClass = RsSiteObj("IsClass")
		IsFont = RsSiteObj("IsFont")
		IsSpan = RsSiteObj("IsSpan")
		IsObjectTF = RsSiteObj("IsObject")
		IsIFrame = RsSiteObj("IsIFrame")
		IsScript = RsSiteObj("IsScript")
		IndexRule = RsSiteObj("IndexRule")
		StartPageNum = RsSiteObj("StartPageNum")
		EndPageNum = RsSiteObj("EndPageNum")
		HandPageContent = RsSiteObj("HandPageContent")
		OtherType = RsSiteObj("OtherType")
		HandSetAuthor = RsSiteObj("HandSetAuthor")
		HandSetSource = RsSiteObj("HandSetSource")
		HandSetAddDate = RsSiteObj("HandSetAddDate")
		ObjURL = GetOtherURL(CollectPageNumber,RsSiteObj)
		IsReverse=RsSiteObj("IsReverse")
		WebCharset = RsSiteObj("WebCharset")
		WaterPrintTF = RsSiteObj("WaterPrintTF")
		CS_SiteReKeyID = RsSiteObj("RulerID")
		if ObjURL = "" then
			CollectPageNumber = 0
			CollectStartLocation = 0
			CollectedPageURL = ""
			CollectSiteIndex = CollectSiteIndex + 1
			Set RsSiteObj = Nothing
			GetCollectPara
			Exit Function
		else
			if CollectPageNumber > G_NEWS_LIST_PAGES_NUMBER then
				CollectPageNumber = 0
				CollectStartLocation = 0
				CollectedPageURL = ""
				CollectSiteIndex = CollectSiteIndex + 1
				Set RsSiteObj = Nothing
				GetCollectPara
				Exit Function
			end if
		end if
	end if
	Set RsSiteObj = Nothing
End Function

Function GetOtherURL(PageNum,Obj) '取得其他新闻列表的URL
	Dim OtherObjURL,OtherResponseAllStr,OtherNewsListArray,i
	if PageNum = 0 then
		GetOtherURL = CollectObjURL
		CollectedPageURL = ""
	else
		Select Case OtherType
			Case 0 '不分页
				GetOtherURL = ""
			Case 1 '标记分页
				if IsNull(OtherPageHeadSetting) OR IsNull(OtherPageFootSetting) OR (OtherPageFootSetting = "") OR (OtherPageHeadSetting = "") then
					GetOtherURL = ""
				else
					if PageNum = 1 then
						CollectedPageURL = CollectObjURL
					end if
					OtherResponseAllStr = GetPageContent(FormatUrl(CollectedPageURL,CollectObjURL),WebCharset)
					OtherObjURL = GetOtherContent(OtherResponseAllStr,OtherPageHeadSetting,OtherPageFootSetting)
					if OtherObjURL <> "" then
						OtherObjURL = FormatUrl(OtherObjURL,CollectObjURL)
					else
						OtherObjURL = ""
					end if
					GetOtherURL = OtherObjURL
				end if
			Case 2 '索引分页
				if IsNull(IndexRule) OR (IndexRule = "") OR IsNull(StartPageNum) OR (StartPageNum = "") OR IsNull(EndPageNum) OR (EndPageNum = "") then
					GetOtherURL = ""
				else
					if Not IsNumeric(StartPageNum) OR Not IsNumeric(EndPageNum) then
						GetOtherURL = ""
					else
						if CInt(StartPageNum) < CInt(EndPageNum) Then '按从小到大的页数
							if PageNum >= CInt(EndPageNum) then
								GetOtherURL = ""
							else
								if PageNum = 1 then
									IndexRule = Replace(FormatUrl(IndexRule,CollectObjURL),"^$^",StartPageNum)
								else
									StartPageNum = CInt(StartPageNum) + PageNum - 1
									IndexRule = Replace(FormatUrl(IndexRule,CollectObjURL),"^$^",StartPageNum)
								end if
								GetOtherURL = IndexRule
							end if
						Else  '按从大到小的页数，从而实现倒序采集，比如从10到1
							if PageNum >= CInt(StartPageNum) then
								GetOtherURL = ""
							else
								if PageNum = 1 then
									IndexRule = Replace(FormatUrl(IndexRule,CollectObjURL),"^$^",StartPageNum)
								else
									EndPageNum = CInt(StartPageNum) - PageNum + 1
									IndexRule = Replace(FormatUrl(IndexRule,CollectObjURL),"^$^",EndPageNum)
								end if
								GetOtherURL = IndexRule
							end if
						end if
					end if
				end if
			Case 3 '手工分页
				if IsNull(HandPageContent) OR (HandPageContent = "") then
					GetOtherURL = ""
				ElseIf InStr(HandPageContent,Chr(10))=0 And PageNum<2 Then
					GetOtherURL = HandPageContent
				Else
					HandPageContent = Split(HandPageContent,Chr(10))
					if PageNum > UBound(HandPageContent) then
						GetOtherURL = ""
					else
						if HandPageContent(PageNum - 1) <> "" then
							GetOtherURL = HandPageContent(PageNum - 1)
						else
							GetOtherURL = ""
						end if
					end if
				end if
			Case Else
				GetOtherURL = ""
		End Select
	end if
End Function

Function GetNewsPageContent()
	Dim NewsPageStr,TitleStr,ContentStr,AuthorStr,SourceStr,AddDate,i
	Dim ResponseAllStr,NewsListStr,NewsLinkStr,RsCheckNewsObj
	Dim NewsListStrArray,TempArray
	ResponseAllStr = GetPageContent(FormatUrl(ObjURL,CollectObjURL),WebCharset)	
	if ResponseAllStr = False then
		CollectPageNumber = CollectPageNumber + 1
		ReturnValue = ReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>错误</strong>:读取新闻列表页面失败<br>"
		Exit Function
	end if

	Dim BLinkHeadSetting,BLinkFootSetting
	BLinkHeadSetting = False
	BLinkFootSetting = False
	
	If Instr(LinkHeadSetting,"[变量]")<=0 Then
		BLinkHeadSetting = True
	ElseIf Instr(LinkFootSetting,"[变量]")<=0 Then
		BLinkFootSetting = True
	End If
	If InStr(ResponseAllStr,ListHeadSetting)>0 And InStr(ResponseAllStr,ListFootSetting) <> 0 Then
		NewsListStr = GetOtherContent(ResponseAllStr,ListHeadSetting,ListFootSetting)
	Else 
		NewsListStr = ResponseAllStr
	End If

	If BLinkHeadSetting Then
		NewsListStr = Mid(NewsListStr,Instr(NewsListStr,LinkHeadSetting)+len(LinkHeadSetting))
		NewsListStrArray = Split(NewsListStr,LinkHeadSetting)
	elseif BLinkFootSetting Then 
		NewsListStr = Left(NewsListStr,InstrRev(NewsListStr,LinkFootSetting))
		NewsListStrArray = Split(NewsListStr,LinkFootSetting)
	else
		NewsListStrArray = Array("")
	End If


	'倒序采集
	
	If IsReverse="1" then 
		Dim TempArr,j
		TempArr=NewsListStrArray
		For j =0 to UBound(NewsListStrArray)
			NewsListStrArray(j)=TempArr(UBound(NewsListStrArray)-j)
		Next 
		If Num>0 And Num-1<=UBound(NewsListStrArray) Then
			TempArr=NewsListStrArray
			For j =0 to Num-1 'UBound(NewsListStrArray)
				NewsListStrArray(j)=TempArr(UBound(NewsListStrArray)-Num+j+1)
			Next 	
		End If 
	End If

	For i = CollectStartLocation to CollectStartLocation + CollectMaxOfOnePage - 1
		if i > UBound(NewsListStrArray) Or (i >= Num And Num<>0) then
			CollectPageNumber = CollectPageNumber + 1
			CollectStartLocation = 0
			CollectedPageURL = ObjURL
			Exit Function
		end If

		AllNewsNumber = AllNewsNumber + 1
		if NewsListStrArray(i) <> "" then
			If BLinkHeadSetting=True Then
				TempArray = GetOtherContent(LinkHeadSetting&NewsListStrArray(i),LinkHeadSetting,LinkFootSetting) 
			ElseIf BLinkFootSetting=True Then 
				TempArray = GetOtherContent(NewsListStrArray(i)&LinkFootSetting,LinkHeadSetting,LinkFootSetting) 
			End If 
			if TempArray <> "" Then
				NewsLinkStr = LoseHtml(FormatUrl(TempArray,CollectObjURL))
				NewsPageStr = GetPageContent(NewsLinkStr,WebCharset)
				if NewsPageStr <> False then		
					TitleStr = LoseHtml(GetOtherContent(NewsPageStr,PageTitleHeadSetting,PageTitleFootSetting))
					Set RsCheckNewsObj = CollectConn.Execute("Select * from FS_News where Links='" & NoSqlHack(NewsLinkStr) & "'")
					if Not RsCheckNewsObj.Eof then
						ReturnValue = GetOneNewsReturnValue(1,i + 1,TitleStr,"",NewsLinkStr) & ReturnValue
					else
						ContentStr = ReplaceKeyWords(GetOneNewsContent(NewsPageStr,NewsLinkStr))
						ContentStr = ReplaceContentStr(ContentStr)
						if SaveRemotePic then ContentStr = ReplaceIMGRemoteUrl(ContentStr,SaveIMGPath,p_DoMain_Str,p_SYS_ROOT_DIR,NewsLinkStr,SaveRemotePic,WaterPrintTF)
						if TitleStr = "" then
							ReturnValue = GetOneNewsReturnValue(2,i + 1,"","",NewsLinkStr) & ReturnValue
						elseif ContentStr = "" then
							ReturnValue = GetOneNewsReturnValue(3,i + 1,TitleStr,"",NewsLinkStr) & ReturnValue
						else
							ReturnValue = GetOneNewsReturnValue(4,i + 1,TitleStr,ContentStr,NewsLinkStr) & ReturnValue
							if IsNull(HandSetAuthor) OR (HandSetAuthor = "") then
								AuthorStr = LoseHtml(GetOtherContent(NewsPageStr,AuthorHeadSetting,AuthorFootSetting))
							else
								AuthorStr = HandSetAuthor
							end if
							if IsNull(HandSetSource) OR (HandSetSource = "") then
								SourceStr = LoseHtml(GetOtherContent(NewsPageStr,SourceHeadSetting,SourceFootSetting))
							else
								SourceStr = HandSetSource
							end if
							if IsNull(HandSetAddDate) OR Not IsDate(HandSetSource) then
								AddDate = LoseHtml(GetOtherContent(NewsPageStr,AddDateHeadSetting,AddDateFootSetting))
							else
								AddDate = HandSetSource
							end if
							if AddDate <> "" then
								if Not IsDate(AddDate) then	AddDate = Now
							else
								AddDate = Now
							end if
							SaveCollectContent TitleStr,NewsLinkStr,ContentStr,AuthorStr,SourceStr,AddDate
						end if
					end if
					Set RsCheckNewsObj = Nothing
				else
					ReturnValue = GetOneNewsReturnValue(5,i + 1,"","",NewsLinkStr) & ReturnValue
				end if
			else
				ReturnValue = GetOneNewsReturnValue(5,i + 1,"","",NewsLinkStr) & ReturnValue
			end if
		else
			ReturnValue = GetOneNewsReturnValue(5,i + 1,"","",NewsLinkStr) & ReturnValue
		end if
	Next
	CollectStartLocation = i
End Function

Function ResumeGetNewsPageContent()
	dim ResumeSql,RsResumeNewsObj,ResumeNewsURL,ResumeNewsURL1,ResumeNewsLocation
	ResumeSql = "Select top 1 Links from FS_News where SiteID='" & NoSqlHack(CollectingSiteID) &"' order by ID DESC"
	Set RsResumeNewsObj = CollectConn.Execute(ResumeSql)	
	If RsResumeNewsObj.EOF Then 
		set RsResumeNewsObj = nothing
		response.Write("<script>alert(""无法确定您以前的采集记录，\n续采失败！"");history.go(-2);</script>")	
	else
		ResumeNewsURL = RsResumeNewsObj("Links")
		set RsResumeNewsObj = nothing
	End If
	

	Dim NewsPageStr,TitleStr,ContentStr,AuthorStr,SourceStr,AddDate,i,n
	Dim ResponseAllStr,NewsListStr,NewsLinkStr,RsCheckNewsObj
	Dim NewsListStrArray,TempArray
	ResponseAllStr = GetPageContent(FormatUrl(ObjURL,CollectObjURL),WebCharset)	
	if ResponseAllStr = False then
		CollectPageNumber = CollectPageNumber + 1
		ReturnValue = ReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>错误</strong>:读取新闻列表页面失败<br>"
		Exit Function
	end if

	Dim BLinkHeadSetting,BLinkFootSetting
	BLinkHeadSetting = False
	BLinkFootSetting = False
	If Instr(LinkHeadSetting,"[变量]")<=0 Then
		BLinkHeadSetting = True
	elseif Instr(LinkFootSetting,"[变量]")<=0 Then
		BLinkFootSetting = True
	End If
	If InStr(ResponseAllStr,ListHeadSetting)>0 And InStr(ResponseAllStr,ListFootSetting) Then
		NewsListStr = GetOtherContent(ResponseAllStr,ListHeadSetting,ListFootSetting)
	Else 
		NewsListStr = ResponseAllStr
	End If

	If BLinkHeadSetting Then
		NewsListStr = Mid(NewsListStr,Instr(NewsListStr,LinkHeadSetting)+len(LinkHeadSetting))
		NewsListStrArray = Split(NewsListStr,LinkHeadSetting)
	elseif BLinkFootSetting Then 
		NewsListStr = Left(NewsListStr,InstrRev(NewsListStr,LinkFootSetting))
		NewsListStrArray = Split(NewsListStr,LinkFootSetting)
	End If
	
	For n = 0 to UBound(NewsListStrArray)					
		Dim tempURL
		tempURL=LoseHtml(FormatUrl(GetOtherContent(LinkHeadSetting&NewsListStrArray(n),LinkHeadSetting,LinkFootSetting),CollectObjURL))
		If ResumeNewsURL = tempURL Then
			Exit For
		ElseIf n>=UBound(NewsListStrArray) Then
			AllNewsNumber = AllNewsNumber+n
			CollectPageNumber = CollectPageNumber + 1
			CollectStartLocation = 0
			CollectedPageURL = ObjURL
			Exit Function 			
		End If
	Next 
	CollectStartLocation = n+1

	If IsReverse="1" then '倒序采集
		Dim TempArr,j
		TempArr=NewsListStrArray
		For j =0 to UBound(NewsListStrArray)
			NewsListStrArray(j)=TempArr(UBound(NewsListStrArray)-j)
		Next 
	End If

	For i = CollectStartLocation to CollectStartLocation + CollectMaxOfOnePage - 1
		if i > UBound(NewsListStrArray) Then
			CollectPageNumber = CollectPageNumber + 1
			CollectStartLocation = 0
			CollectedPageURL = ObjURL
			Exit Function
		end If

		AllNewsNumber = AllNewsNumber + 1
		If BLinkHeadSetting Then
			TempArray = GetOtherContent(LinkHeadSetting&NewsListStrArray(i),LinkHeadSetting,LinkFootSetting) 
		elseif BLinkFootSetting Then 
			TempArray = GetOtherContent(NewsListStrArray(i)&LinkFootSetting,LinkHeadSetting,LinkFootSetting) 
		End If  
		if TempArray <> "" Then
			NewsLinkStr = LoseHtml(FormatUrl(TempArray,CollectObjURL))
			Set RsCheckNewsObj = CollectConn.Execute("Select * from FS_News where Links='" & NoSqlHack(NewsLinkStr) & "'")
			if RsCheckNewsObj.Eof then
				NewsPageStr = GetPageContent(NewsLinkStr,WebCharset)
				if NewsPageStr <> False then
					TitleStr = LoseHtml(GetOtherContent(NewsPageStr,PageTitleHeadSetting,PageTitleFootSetting))
				Set RsCheckNewsObj = CollectConn.Execute("Select * from FS_News where Links='" & NoSqlHack(NewsLinkStr) & "'")
					ContentStr = ReplaceKeyWords(GetOneNewsContent(NewsPageStr,NewsLinkStr))
					ContentStr = ReplaceContentStr(ContentStr)
					if SaveRemotePic then ContentStr = ReplaceIMGRemoteUrl(ContentStr,SaveIMGPath,p_DoMain_Str,p_SYS_ROOT_DIR,NewsLinkStr,SaveRemotePic,WaterPrintTF)
					if TitleStr = "" then
						ReturnValue = GetOneNewsReturnValue(2,i + 1,"","",NewsLinkStr) & ReturnValue
					elseif ContentStr = "" then
						ReturnValue = GetOneNewsReturnValue(3,i + 1,TitleStr,"",NewsLinkStr) & ReturnValue
					else
						ReturnValue = GetOneNewsReturnValue(4,i + 1,TitleStr,ContentStr,NewsLinkStr) & ReturnValue
						if IsNull(HandSetAuthor) OR (HandSetAuthor = "") then
							AuthorStr = LoseHtml(GetOtherContent(NewsPageStr,AuthorHeadSetting,AuthorFootSetting))
						else
							AuthorStr = HandSetAuthor
						end if
						if IsNull(HandSetSource) OR (HandSetSource = "") then
							SourceStr = LoseHtml(GetOtherContent(NewsPageStr,SourceHeadSetting,SourceFootSetting))
						else
							SourceStr = HandSetSource
						end if
						if IsNull(HandSetAddDate) OR Not IsDate(HandSetAddDate) then
							AddDate = LoseHtml(GetOtherContent(NewsPageStr,AddDateHeadSetting,AddDateFootSetting))
						else
							AddDate = HandSetAddDate
						end if
						if AddDate <> "" then
							if Not IsDate(AddDate) then	AddDate = Now
						else
							AddDate = Now
						end if
						SaveCollectContent TitleStr,NewsLinkStr,ContentStr,AuthorStr,SourceStr,AddDate
					end if
					Set RsCheckNewsObj = Nothing
				else
					ReturnValue = GetOneNewsReturnValue(5,i + 1,"","",NewsLinkStr) & ReturnValue
				End If
			ElseIf session("ConfirmCollectRevert")<>"ConfirmCollectRevert" Then
				session("ConfirmCollectRevert") = "ConfirmCollectRevert"
				response.write("<script>if(confirm(""您改变过采集顺序吗？\n如果修改过，请单击确定改回原样再续采！\n没有修改过请单击取消继续！""))window.location=""site.asp""</script>")
			End If
		End If		
	Next
	CollectStartLocation = i
End Function

Function GetOneNewsContent(FirstPageContent,NewsLinkStr)
	Dim OtherPageNewsLink,OtherPageNewsContentStr,tempSplitArr1,tempSplitArr2
	Dim f_Collect_Index,f_Temp_Array,f_URL,f_Start,f_End,f_Int,f_I
	'On Error Resume Next
	f_Collect_Index = 0
	OtherPageNewsContentStr = FirstPageContent
	GetOneNewsContent = GetOtherContent(FirstPageContent,PagebodyHeadSetting,PagebodyFootSetting)
	Select Case OtherNewsType
		Case 0

		Case 1
			if IsNull(OtherNewsPageHeadSetting) OR IsNull(OtherNewsPageFootSetting) OR (OtherNewsPageHeadSetting = "") OR (OtherNewsPageFootSetting = "") Then
				OtherPageNewsLink = ""
			ElseIf InStr(OtherPageNewsContentStr,OtherNewsPageFootSetting)>0 And InStr(OtherPageNewsContentStr,OtherNewsPageHeadSetting)>0 Then
				tempSplitArr1 = Split(OtherPageNewsContentStr,OtherNewsPageFootSetting)
				tempSplitArr2 = Split(tempSplitArr1(0),OtherNewsPageHeadSetting)
				OtherPageNewsLink = tempSplitArr2(Ubound(tempSplitArr2))
			Else
				OtherPageNewsLink =  GetOtherContent(OtherPageNewsContentStr,PagebodyHeadSetting,PagebodyFootSetting)
			End If
			Do While (OtherPageNewsLink <> "")
				OtherPageNewsLink = FormatUrl(OtherPageNewsLink,NewsLinkStr)
				OtherPageNewsContentStr = GetPageContent(OtherPageNewsLink,WebCharset)
				If  InStr(OtherPageNewsContentStr,OtherNewsPageHeadSetting)>0 And InStr(OtherPageNewsContentStr,OtherNewsPageFootSetting)>0 Then
					tempSplitArr1 = Split(OtherPageNewsContentStr,OtherNewsPageFootSetting)
					tempSplitArr2 = Split(tempSplitArr1(0),OtherNewsPageHeadSetting)
					OtherPageNewsLink = tempSplitArr2(Ubound(tempSplitArr2))
				Else
					OtherPageNewsLink =  GetOtherContent(OtherPageNewsContentStr,PagebodyHeadSetting,PagebodyFootSetting)
				End If
				If OtherPageNewsContentStr<>False Then
					GetOneNewsContent = GetOneNewsContent & "[FS:PAGE]" & GetOtherContent(OtherPageNewsContentStr,PagebodyHeadSetting,PagebodyFootSetting)
				Else
					OtherPageNewsLink = ""
				End If
				If Err Then
					Err.clear
					OtherPageNewsLink = ""
				End If
			Loop
			If Right(GetOneNewsContent,9) = "[FS:PAGE]" Then
				GetOneNewsContent = Left(GetOneNewsContent,Len(GetOneNewsContent) - 9)
			End iF	
		Case 2
			Dim Temp_NewsPageStr,Temp_NewsFistStr,Temp_NewsArray1,Temp_NewsArray2
			If IsNull(OtherNewsPageIndexSetting) Or OtherNewsPageIndexSetting = "" Then
				OtherPageNewsLink = ""
			Else	
				If InStr(OtherNewsPageIndexSetting,"[分页新闻]")>0 And InStr(OtherNewsPageIndexSetting,"[变量]")>0 Then
					tempSplitArr1 = Split(OtherNewsPageIndexSetting,"[分页新闻]")
					tempSplitArr2 = Split(tempSplitArr1(1),"[变量]")
					Temp_NewsPageStr = tempSplitArr2(0)
					Temp_NewsFistStr = tempSplitArr1(0)
				End If
				If InStr(OtherPageNewsContentStr,Temp_NewsFistStr)>0 And InStr(OtherPageNewsContentStr,Temp_NewsPageStr)>0 Then
					Temp_NewsArray1 = Split(OtherPageNewsContentStr,Temp_NewsFistStr)
					Temp_NewsArray2	= Split(Temp_NewsArray1(1),Temp_NewsPageStr)
					OtherPageNewsLink = Temp_NewsArray2(0)
				Else
					OtherPageNewsLink =  GetOtherContent(OtherPageNewsContentStr,PagebodyHeadSetting,PagebodyFootSetting)
				End If
			End If
			Do While (OtherPageNewsLink <> "")
				OtherPageNewsLink = FormatUrl(OtherPageNewsLink,NewsLinkStr)
				OtherPageNewsContentStr = GetPageContent(OtherPageNewsLink,WebCharset)
				If InStr(OtherPageNewsContentStr,Temp_NewsFistStr)>0 And InStr(OtherPageNewsContentStr,Temp_NewsPageStr)>0 Then
					Temp_NewsArray1 = Split(OtherPageNewsContentStr,Temp_NewsFistStr)
					Temp_NewsArray2	= Split(Temp_NewsArray1(1),Temp_NewsPageStr)
					OtherPageNewsLink = Temp_NewsArray2(0)
				Else
					OtherPageNewsLink =  GetOtherContent(OtherPageNewsContentStr,PagebodyHeadSetting,PagebodyFootSetting)
				End If
				If OtherPageNewsContentStr <> False Then
					GetOneNewsContent = GetOneNewsContent & "[FS:PAGE]" & GetOtherContent(OtherPageNewsContentStr,PagebodyHeadSetting,PagebodyFootSetting)
				Else
					OtherPageNewsLink = ""
				End If
				If Err Then
					Err.clear
					OtherPageNewsLink = ""
				End If
			Loop
			If Right(GetOneNewsContent,9) = "[FS:PAGE]" Then
				GetOneNewsContent = Left(GetOneNewsContent,Len(GetOneNewsContent) - 9)
			End iF	
	End Select
End Function 

Function GetOneNewsReturnValue(CauseIndex,NewsIndex,Title,Content,LinkStr)
	Select Case CauseIndex
		Case 1  '不允许重名保存
			GetOneNewsReturnValue = "</br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>序号</strong>： " & NewsIndex
			GetOneNewsReturnValue = GetOneNewsReturnValue & "&nbsp;&nbsp;&nbsp;&nbsp;<strong>结果</strong>： <font color=red>已经采集，在等待审核或者在历史纪录里面</font>"
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>标题</strong>： " & Title
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>新闻链接</strong>： <a target=""_blank"" href=""" & LinkStr & """>" & LinkStr & "</a><br>"
		Case 2 '标题为空，没有保存
			GetOneNewsReturnValue = "</br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>序号</strong>： " & NewsIndex
			GetOneNewsReturnValue = GetOneNewsReturnValue & "&nbsp;&nbsp;&nbsp;&nbsp;<strong>结果</strong>： <font color=red>标题为空，没有保存</font>"
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>新闻链接</strong>： <a target=""_blank"" href=""" & LinkStr & """>" & LinkStr & "</a><br>"
		Case 3 '内容为空，没有保存
			GetOneNewsReturnValue = "</br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>序号</strong>： " & NewsIndex
			GetOneNewsReturnValue = GetOneNewsReturnValue & "&nbsp;&nbsp;&nbsp;&nbsp;<strong>结果</strong>： <font color=red>内容为空，没有保存</font>"
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>标题</strong>： " & Title
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>新闻链接</strong>： <a target=""_blank"" href=""" & LinkStr & """>" & LinkStr & "</a><br>"
		Case 4 '成功保存
			GetOneNewsReturnValue = "</br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>序号</strong>： " & NewsIndex
			GetOneNewsReturnValue = GetOneNewsReturnValue & "&nbsp;&nbsp;&nbsp;&nbsp;<strong>结果</strong>： 采集成功"
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>标题</strong>： " & Title
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>内容</strong>： " & Left(LoseHtml(Content),30) & "&nbsp;&nbsp;......"
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>新闻链接</strong>： <a target=""_blank"" href=""" & LinkStr & """>" & LinkStr & "</a><br>"
			CollectOKNumber = CollectOKNumber + 1
		Case 5 '不能够读取新闻目标页面
			GetOneNewsReturnValue = "</br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>序号</strong>： " & NewsIndex
			GetOneNewsReturnValue = GetOneNewsReturnValue & "&nbsp;&nbsp;&nbsp;&nbsp;<strong>结果</strong>： <font color=red>不能够读取新闻目标页面</font>"
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>新闻链接</strong>： <a target=""_blank"" href=""" & LinkStr & """>" & LinkStr & "</a><br>"
		Case else
	End Select
End Function

Function SaveCollectContent(Title,Links,Content,Author,SourceString,AddDate)
	Dim RsNewsObj,RsTempObj
	Set RsNewsObj = Server.CreateObject(G_FS_RS)
	RsNewsObj.Open "Select * from FS_News where 1=0",CollectConn,3,3
	RsNewsObj.AddNew
	RsNewsObj("Title") = NoSqlHack(LoseHtml(Title))
	RsNewsObj("Links") = NoSqlHack(Links)
	RsNewsObj("Content") = NoSqlHack(Content)
	RsNewsObj("ContentLength") = NoSqlHack(Len(Content))
	RsNewsObj("AddDate") = NoSqlHack(AddDate)
	RsNewsObj("ImagesCount") = 0
	RsNewsObj("SiteID") = NoSqlHack(CollectingSiteID)
	RsNewsObj("Author") = NoSqlHack(Left(Author,200))
	If AuditTF = False Then
		RsNewsObj("IsLock") = 1
	Else
		RsNewsObj("IsLock") = 0
	End If
	If AutoCollect Then
		RsNewsObj("History") = 0'信息在采集库中的状态0为没有入库1为入库
	End If
	RsNewsObj("Source") = Left(SourceString,200)
	RsNewsObj("ReviewTF") = 0
	RsNewsObj.UpDate
	RsNewsObj.Close
	Set RsNewsObj = Nothing
	'If AutoCollect Then
	'	NewsToSystem Title,Content,Author,SourceString,AddDate
	'End If
Rem #######################以上三行被注释掉的原因是新闻采集后自动被提交到新闻数据库中因为时间问题 暂时没有在这里加判断特此备注：Crazy
End Function

Function ReplaceKeyWords(Content)
	Dim RsRuleObj,HeadSeting,FootSeting,ReContent,regEx
	IF CS_SiteReKeyID = "" Or IsNull(CS_SiteReKeyID) Then 
		ReplaceKeyWords = Content
		Exit Function
	End IF	
	Set RsRuleObj = CollectConn.Execute("Select * from FS_Rule where ID In(" & FormatIntArr(CS_SiteReKeyID) & ")")
	do while Not RsRuleObj.Eof
		HeadSeting = RsRuleObj("HeadSeting")
		FootSeting = RsRuleObj("FootSeting")
		ReContent = RsRuleObj("ReContent")
		if IsNull(FootSeting) or FootSeting = "" then
			if HeadSeting <> "" then
				Content = Replace(Content,HeadSeting,ReContent)
			end if
		end if
		if Not IsNull(FootSeting) and FootSeting <> "" and Not IsNull(HeadSeting) and HeadSeting <> ""  then
			Set regEx = New RegExp
			regEx.Pattern = HeadSeting & "[^\0]*" & FootSeting
			regEx.IgnoreCase = False
			regEx.Global = True
			if IsNull(ReContent) then
				Content = regEx.Replace(Content,"")
			else
				Content = regEx.Replace(Content,ReContent)
			end if
			Set regEx = Nothing
		end if
		RsRuleObj.MoveNext
	loop
	Set RsRuleObj = Nothing
	ReplaceKeyWords = Content
End Function

Function NewsToSystem(Title,Content,Author,SourceString,AddDate)
	Dim f_Field_Array,f_Source_Sql,f_Object_Sql,f_Collect_RS,f_System_RS,f_i,TempNewsID,f_System_RS_Pop,OldID,Fs_news
	Dim Str_Temp_Flag,temp_j,StrSql,p_File_Ext_Name,p_Save_Path,sRootDir,str_CurrPath
	Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
	Temp_Admin_Name = Session("Admin_Name")
	Temp_Admin_Is_Super = Session("Admin_Is_Super")
	Temp_Admin_FilesTF = Session("Admin_FilesTF")
	

	if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/" & G_VIRTUAL_ROOT_DIR else sRootDir=""
	If Temp_Admin_Is_Super = 1 then
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	Else
		If Temp_Admin_FilesTF = 0 Then
			str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
		Else
			str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
		End If	
	End if

	Set Fs_news = New Cls_News
	Fs_News.GetSysParam()
	If Not Fs_news.IsSelfRefer Then
		p_File_Ext_Name = "html"
		p_Save_Path = "/" & Year(Now) & "-" & Month(Now) & "-" & Day(Now)
	else
		p_File_Ext_Name = Fs_News.fileExtName
		p_Save_Path = Fs_news.SaveNewsPath(Fs_news.fileDirRule)
	end if
	'-----2006-12-07 by ken 采集数据转移到主数据库时候，设置生成静态文件扩展名
	If p_File_Ext_Name <> "html" Then
		If CInt(p_File_Ext_Name) = 0 then
			p_File_Ext_Name = "html"
		ElseIf CInt(p_File_Ext_Name) = 1 then
			p_File_Ext_Name = "htm"
		ElseIf CInt(p_File_Ext_Name) = 2 then
			p_File_Ext_Name = "shtml"
		ElseIf CInt(p_File_Ext_Name) = 3 then
			p_File_Ext_Name = "shtm"
		ElseIf CInt(p_File_Ext_Name) = 4 then
			p_File_Ext_Name = "asp"
		Else
			p_File_Ext_Name = "html"
		End If				
	End If
	StrSql="INSERT INTO FS_NS_News([NewsID],[PopId],[ClassID],[NewsTitle],[isShowReview],[Content],[Templet],[Source],[Author],[SaveNewsPath],[FileName],[FileExtName],[NewsProperty],[isLock],[addtime],[isPicNews],[NewsPicFile],[NewsSmallPicFile]) VALUES ("
	TempNewsID=GetRamCode(15)
	StrSql=StrSql & "'" & TempNewsID & "'"
	StrSql=StrSql & ",0"
	StrSql=StrSql & ",'" & NoSqlHack(GetNewsInfoBySiteID(CollectingSiteID,"ClassID")) & "'"
	StrSql=StrSql & ",'"&NoSqlHack(Title)&"'"
	StrSql=StrSql & ",0"
	StrSql=StrSql & ",'"&Replace(Content,"'","''")&"'"
	StrSql=StrSql & ",'"&NoSqlHack(GetNewsInfoBySiteID(CollectingSiteID,"Temp"))&"'"
	StrSql=StrSql & ",'"&NoSqlHack(left(SourceString,50))&"'"
	StrSql=StrSql & ",'"&NoSqlHack(Left(Author,50))&"'"
	StrSql=StrSql & ",'"&NoSqlHack(Fs_news.SaveNewsPath(Fs_news.fileDirRule))&"'"
	'------
	OldID = Fs_News.strFileNameRule(Fs_News.fileNameRule,0,0)
	if instr(OldID,"自动编号ID") > 0 then OldID = Replace(OldID,"自动编号ID",TempNewsID)
	if instr(OldID,"唯一NewsID") > 0 then OldID = Replace(OldID,"唯一NewsID",TempNewsID)
	'------
	StrSql=StrSql & ",'"&OldID&"'"
	StrSql=StrSql & ",'"&p_File_Ext_Name&"'"
	StrSql=StrSql & ",'0,1,1,0,1,0,0,0,1,0,0'"
	If AuditTF = False Then
		StrSql=StrSql & ",1"
	Else
		StrSql=StrSql & ",0"
	End If
	StrSql=StrSql & ",'"&AddDate&"'"
	If GetCeSitePicTF(CollectingSiteID) = True Then
		If ContentInnerPicTF(Replace(Content,"'","''"),"TF") = True Then
			StrSql = StrSql & ",1"
			StrSql = StrSql & ",'" & ContentInnerPicTF(Replace(Content,"'","''"),"PicUrl") & "'"
			StrSql = StrSql & ",'" & ContentInnerPicTF(Replace(Content,"'","''"),"PicUrl") & "'"
		Else
			StrSql = StrSql & ",0"
			StrSql = StrSql & ",''"
			StrSql = StrSql & ",''"
		End If
	Else
		StrSql = StrSql & ",0"
		StrSql = StrSql & ",''"
		StrSql = StrSql & ",''"
	End If
	'====End=====================			
	StrSql=StrSql & ")"
	'On Error Resume Next
	If StrSql<>"" Then
		Conn.Execute(StrSql)
	End If
	Set Fs_news = Nothing
	NewsToSystem = True
End Function

'----
Function NUllToStr(num)
	If IsNull(num) Or num = "" Then
		NUllToStr = 0
	Else
		If Not IsNumeric(num) Then
			NUllToStr = 0
		Else
			NUllToStr = Cint(num)
		End If	
	End if
End Function

'===========================================================
'判断传入的字符传中是否包含本地图片并取得此图片地址
'===========================================================
Function ContentInnerPicTF(StrCon,ReturnTF)
	Dim ConStr,Re,InnerPicAll,FistPicUrl,PicUrlStr
	ConStr = StrCon & ""
	Set Re = New RegExp
	Re.IgnoreCase = True
	Re.Global = True
	Re.Pattern = "(src\S+\.{1}(gif|jpg|png)(""|\'|>|\s)?)"
	InnerPicAll = ""
	Set InnerPicAll = Re.Execute(ConStr)
	Set Re = Nothing
	FistPicUrl = ""
	For Each PicUrlStr in InnerPicAll
		FistPicUrl = Replace(Replace(Replace(PicUrlStr,"src=",""),"'",""),"""","")
		If LCase(Left(FistPicUrl,Len(sRootDir))) = LCase(sRootDir) Then
			FistPicUrl = Mid(FistPicUrl,Len(sRootDir)+1)
		End If
		Exit For
	Next
	If ReturnTF = "TF" Then
		If FistPicUrl <> "" And (Not IsNull(FistPicUrl)) then
			ContentInnerPicTF = True
		Else
			ContentInnerPicTF = False	
		End If
	ElseIf ReturnTF = "PicUrl" Then
		If FistPicUrl <> "" And (Not IsNull(FistPicUrl)) then
			ContentInnerPicTF = FistPicUrl
		End If
	End If					
End Function

'===========================================================
'判断传入的采集站点设置属性
'===========================================================
Function GetCeSitePicTF(SiteID)
	Dim GetSiteRs
	IF SiteID = "" Then : GetCeSitePicTF = False : Exit Function
	SiteID = Clng(SiteID)
	Set GetSiteRs = CollectConn.ExeCute("Select IsAutoPicNews From FS_Site Where ID = " & CintStr(SiteID) & " And IsLock = 0")
	If GetSiteRs.Eof Then
		GetCeSitePicTF = False
	Else
		If GetSiteRs(0) = 1 Then
			GetCeSitePicTF = True
		Else
			GetCeSitePicTF = False
		End If
	End If
	GetSiteRs.Close : Set GetSiteRs = NoThing			
End Function

Function GetNewsInfoBySiteID(SiteID,Act)
	Dim GetSiteRs
	IF SiteID = "" Or IsNull(SiteID) Or NOt IsNumeric(SiteID) Then
		If Act = "ClassID" Then
			GetNewsInfoBySiteID = 0
		Else
			GetNewsInfoBySiteID = "/" & G_TEMPLETS_DIR & "/NewsClass/new.htm"
		End IF	
	End If
	Set GetSiteRs = CollectConn.ExeCute("Select ToClassID,NewsTemplets From FS_Site Where ID = " & CintStr(SiteID) & " And IsLock = 0")
	If GetSiteRs.Eof Then
		If Act = "ClassID" Then
			GetNewsInfoBySiteID = 0
		Else
			GetNewsInfoBySiteID = "/" & G_TEMPLETS_DIR & "/NewsClass/new.htm"
		End IF
	Else
		If Act = "ClassID" Then
			GetNewsInfoBySiteID = GetSiteRs(0)
		Else
			GetNewsInfoBySiteID = GetSiteRs(1)
		End IF
	End If
	GetSiteRs.Close : Set GetSiteRs = NoThing
End Function

%>






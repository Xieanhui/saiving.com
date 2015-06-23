<%
Dim JSCodeStr,i,TempClassObj,AvailableDoMain
Dim CharIndexStr
GetFunctionstr
Dim ListSpaces,ListSpaceStrs,Temp_ii
ListSpaces =0   '左右两条新闻之间的空格字符个数 
ListSpaceStrs = ""
for Temp_ii = 0 to ListSpaces
	ListSpaceStrs = ListSpaceStrs & "&nbsp;"
next 

'从缓存中获取站点域名，便于完成跨站图片调用。   Fsj.08.09.26
AvailableDoMain="http://"&Session("DomainPath")

Function GetAllChildRenClassID(ClassID)
	If ClassID = "" Then Exit Function
	Dim GetClassIDRs,AllClassID
	Set GetClassIDRs = Conn.ExeCute("Select ClassID From FS_NS_NewsClass Where ParentID = '"&NoSqlHack(ClassID)&"' And ReycleTF = 0 Order By OrderID Desc,ID Desc") 
	If GetClassIDRs.Eof Then
		AllClassID = "'" & ClassID & "'"
	Else
		AllClassID = ""
		Do While Not GetClassIDRs.Eof
			AllClassID = AllClassID & "," & GetAllChildRenClassID(GetClassIDRs(0))
			GetClassIDRs.MoveNext
		Loop
		AllClassID = "'" & ClassID & "'" & AllClassID
	End If
	GetAllChildRenClassID = AllClassID
	GetClassIDRs.Close
	Set GetClassIDRs = Nothing				
End Function

Function CreateSysJS(FileName)'栏目JS新闻列表
	Dim FS_JsObj , RsSysJsObj,ClassIDStr,NewsNum,MarDirection,BrStr,MarSpeed,NaviPic,RsCreateSql,RsCreateObj,DateTF,RowNum,RowSpace,TitleNum,ShowClassTF
	Dim RightDate,ClassID,RsClassObj,PicHeight,MarWidth,OpenMode,MarHeight,PicWidth,ShowTitle,TitleCSS,SavePath,FileNameStr,DateCSS,DateType,LinkCSS,MoreContentStr,MoreContentTF,setHitsValue
	Dim SQLField
	Set FS_JsObj=New Cls_Js
	Set RsSysJsObj = Conn.Execute("Select * from FS_NS_Sysjs where FileName='"&NoSqlHack(FileName)&"'")
	If Not RsSysJsObj.eof then
		ClassID = RsSysJsObj("ClassID")
		If RsSysJsObj("NaviPic")<>"" then
			NaviPic = "<img src="""  & RsSysJsObj("NaviPic") & """ border=""0"">"
		Else
			NaviPic = ""
		End If
		setHitsValue = RsSysJsObj("MarSpeed")
		if isnull(setHitsValue) then setHitsValue= -1
		NewsNum = RsSysJsObj("NewsNum")
		RowNum = RsSysJsObj("RowNum")
		RowSpace = RsSysJsObj("RowSpace")
		TitleNum = RsSysJsObj("TitleNum")
		TitleCSS = RsSysJsObj("TitleCSS")
		SavePath = RsSysJsObj("FileSavePath")
		FileNameStr = RsSysJsObj("FileName")
		DateCSS = RsSysJsObj("DateCSS")
		DateType = RsSysJsObj("DateType")
		MarDirection = RsSysJsObj("MarDirection")
		MarSpeed = RsSysJsObj("MarSpeed")
		PicWidth = RsSysJsObj("PicWidth")
		PicHeight = RsSysJsObj("PicHeight")
		MarWidth = RsSysJsObj("MarWidth")
		MarHeight = RsSysJsObj("MarHeight")
		If RsSysJsObj("OpenMode")=1 then
			OpenMode = " target=""_blank"""
		Else
			OpenMode = " target=""_self"""
		End If
		If RsSysJsObj("ShowTitle")<>0 then
			ShowTitle = True
		Else
			ShowTitle = false
		End If
		If RsSysJsObj("MarDirection")="left" or RsSysJsObj("MarDirection")="right" then
			BrStr = ""
		Else
			BrStr = "<br>"
		End If
		If RsSysJsObj("MoreContent")<>0 then
			MoreContentTF = True
			MoreContentStr = RsSysJsObj("LinkWord")
			LinkCSS = RsSysJsObj("LinkCSS")
		Else
			MoreContentTF = False
		End If
		If RsSysJsObj("DateType")<>0 then
			DateTF = true
		Else
			DateTF = false
		End If
		If RsSysJsObj("ClassName")<>0 then
			ShowClassTF = true
		Else
			ShowClassTF = false
		End If
		If RsSysJsObj("RightDate")<>0 then
			RightDate = true
		Else
			RightDate = false
		End If
		ClassIDStr = ClassID
		If RsSysJsObj("SonClass")=1 then
			ClassIDStr = GetAllChildRenClassID(ClassIDStr)
			If Left(ClassIDStr,1) = "'" Then
				ClassIDStr = Right(ClassIDStr,Len(ClassIDStr) - 1)
			End If
			If Right(ClassIDStr,1) = "'" Then
				ClassIDStr = left(ClassIDStr,Len(ClassIDStr) - 1)
			End If		
		End If
		SQLField = "Class.ClassEName,Class.[Domain],Class.SavePath"
		'WriteInfo(RsSysJsObj("NewsType"))
		Select Case RsSysJsObj("NewsType")
			Case "RecNews"
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&CintStr(NewsNum)&" News.*," & SQLField & " From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.ClassID in ('"&ClassIDStr&"') and News.isRecyle=0 and News.isLock=0 and "& CharIndexStr &"(NewsProperty,1,1)='1' order by News.AddTime desc" '推荐新闻
				else
					RsCreateSql = "Select top "&CintStr(NewsNum)&" News.*," & SQLField & " From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.isRecyle=0 and News.isLock=0 and "& CharIndexStr &"(NewsProperty,1,1)='1' order by News.AddTime desc" '推荐新闻
				end if
				'WriteInfo(RsCreateSql&"-/-/-/"&RsSysJsObj("FileType"))
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&RowSpace&"""><tr>"
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassName,SavePath from FS_NS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href=""" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) &""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href=""" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) &""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						Else
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj) &""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href=""" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) &""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						End IF
					  Else
						If ShowClassTF = true then
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						Else
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						End If
					  End If
					  RsCreateObj.MoveNext
					  if i mod Cint(RowNum) = 0 or RsCreateObj.eof then
						if RightDate = true then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum*2&""" height=""1"" background="""  & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum&""" height=""1"" background="""  & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						end if
					  end if
					next 
					If RsSysJsObj("FileType")=1 then
						Dim RsTempClassObjs
					Set RsTempClassObjs = Conn.Execute("Select SavePath,ClassEName,FileExtName from FS_NS_NewsClass where ClassID='"&NoSqlHack(ClassID)&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在."
							Exit Function
						End If
					End If
					If RightDate = true then
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SavePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" height="&RowSpace&"></td></tr>"
						end if
					Else
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SavePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" height="&RowSpace&"></td></tr>"
						end if
					End If
					JSCodeStr = JSCodeStr & "</table>');"
					WriteFile SavePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_NS_Sysjs Set AddTime='"&Now()&"' where FileName='"&NoSqlHack(FileName)&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SavePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_NS_Sysjs Set AddTime='"&Now()&"' where FileName='"&NoSqlHack(FileName)&"'")
					CreateSysJS = "文件操作成功.但未找到符合条件的新闻。"
				End If
			Case "MarqueeNews"
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&CintStr(NewsNum)&" News.*," & SQLField & " From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.ClassID in ('"&FormatStrArr(ClassIDStr)&"') and News.isRecyle=0 and "& CharIndexStr &"(News.NewsProperty,3,1)='1' and News.isLock=0 order by News.AddTime desc" '滚动新闻
				else
					RsCreateSql = "Select top "&CintStr(NewsNum)&" News.*," & SQLField & " From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.isRecyle=0 and "& CharIndexStr &"(News.NewsProperty,3,1)='1' and News.isLock=0 order by News.AddTime desc" '滚动新闻
				end If
				'Response.write RsCreateSql
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<marquee onmouseout=start() onmouseover=stop() Width="&MarWidth&" Height="&MarHeight&" scrolldelay=80 direction="&MarDirection&" scrollamount="& CInt(MarSpeed) &">"
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassName,SavePath from FS_NS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							  JSCodeStr = JSCodeStr & NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&BrStr
						Else
							  JSCodeStr = JSCodeStr & NaviPic &"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&BrStr
						End IF
					  Else
						If ShowClassTF = true then
							  JSCodeStr = JSCodeStr & NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&BrStr
						Else
							  JSCodeStr = JSCodeStr & NaviPic &"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&BrStr
						End If
					  End If
					  RsCreateObj.MoveNext
					next 
					If RsSysJsObj("FileType")=1 then
					Set RsTempClassObjs = Conn.Execute("Select SavePath,ClassEName,FileExtName from FS_NS_NewsClass where ClassID='"&NoSqlHack(ClassID)&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在."
							Exit Function
						End If
					End If
					if RsSysJsObj("FileType")=1 and MoreContentTF=True then
						JSCodeStr = JSCodeStr &"<a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SavePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs
					end if
					JSCodeStr = JSCodeStr & "</marquee>');"
					WriteFile SavePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_NS_Sysjs Set AddTime='"&Now()&"' where FileName='"&NoSqlHack(FileName)&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SavePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_NS_Sysjs Set AddTime='"&Now()&"' where FileName='"&NoSqlHack(FileName)&"'")
					CreateSysJS = "文件操作成功.但未找到符合条件的新闻。"
				End If
			Case "PicNews"
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&CintStr(NewsNum)&" News.*," & SQLField & " From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.ClassID in ('"&FormatStrArr(ClassIDStr)&"') and News.isRecyle=0 and News.isPicNews=1 and News.isLock=0 order by News.AddTime desc" '图片新闻
				else
					RsCreateSql = "Select top "&CintStr(NewsNum)&" News.*," & SQLField & " From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.isRecyle=0 and News.isPicNews=1 and News.isLock=0 order by News.AddTime desc" '图片新闻
				end if
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&RowSpace&"""><tr>"
					Dim DoMainStr
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  'If LCase(Left(RsCreateObj("NewsSmallPicFile"),4))="http" Then
					  '	DoMainStr = ""
					  'Else
					  '	DoMainStr = AvailableDoMain
					  'End If
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassName,SavePath from FS_NS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							If RightDate = true then
								If ShowTitle = True then
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><table border=0 cellspacing=0 cellpadding=0><tr><td colspan=2 align=center valign=middle><a href="&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&OpenMode&"><img src="&DoMainStr&RsCreateObj("NewsSmallPicFile")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td></tr>"
								  JSCodeStr = JSCodeStr &"<tr><td align=center>"& NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href=""" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) &""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span></div></td></tr></table></td>"
								Else
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><a href=" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) & OpenMode&"><img src="&DoMainStr&RsCreateObj("NewsSmallPicFile")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td>"
								End If
							Else
								If ShowTitle = True then
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><table border=0 cellspacing=0 cellpadding=0><tr><td colspan=2 align=center valign=middle><a href="&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&OpenMode&"><img src="&DoMainStr&RsCreateObj("NewsSmallPicFile")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td></tr>"
								  JSCodeStr = JSCodeStr &"<tr><td align=center>"& NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href=""" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) &""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span></td></tr></table></td>"
								Else
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><a href=" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) & OpenMode&"><img src="&DoMainStr&RsCreateObj("NewsSmallPicFile")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td>"
								End If
							End If
						Else
							If RightDate = true then
								If ShowTitle = True then
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><table border=0 cellspacing=0 cellpadding=0><tr><td colspan=2 align=center valign=middle><a href="&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&OpenMode&"><img src="&DoMainStr&RsCreateObj("NewsSmallPicFile")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td></tr>"
								  JSCodeStr = JSCodeStr &"<tr><td align=center>"& NaviPic &"<a class="""&TitleCSS&""" href=""" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) &""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span></div></td></tr></table></td>"
								Else
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><a href=" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) & OpenMode&"><img src="&DoMainStr&RsCreateObj("NewsSmallPicFile")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td>"
								End If
							Else
								If ShowTitle = True then
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><table border=0 cellspacing=0 cellpadding=0><tr><td colspan=2 align=center valign=middle><a href="&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&OpenMode&"><img src="&DoMainStr&RsCreateObj("NewsSmallPicFile")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td></tr>"
								  JSCodeStr = JSCodeStr &"<tr><td align=center>"& NaviPic &"<a class="""&TitleCSS&""" href=""" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) &""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span></td></tr></table></td>"
								Else
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><a href=" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) & OpenMode & "><img src="&DoMainStr&RsCreateObj("NewsSmallPicFile")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td>"
								End If
							End If
						End IF
					  Else
						If ShowClassTF = true then
							If ShowTitle = True then
							  JSCodeStr = JSCodeStr &"<td align=center valign=middle><table border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=middle><a href=" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) & OpenMode & "><img src="&DoMainStr&RsCreateObj("NewsSmallPicFile")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td></tr>"
							  JSCodeStr = JSCodeStr &"<tr><td align=center>"& NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href=""" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) &""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a></td></tr></table></td>"
							Else
							  JSCodeStr = JSCodeStr &"<td align=center valign=middle><a href=" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) & OpenMode & "><img src="&DoMainStr&RsCreateObj("NewsSmallPicFile")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td>"
							End If
						Else
							If ShowTitle = True then
							  JSCodeStr = JSCodeStr &"<td align=center valign=middle><table border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=middle><a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&"><img src="&DoMainStr&RsCreateObj("NewsSmallPicFile")&" height="&PicHeight&" width="&PicWidth&" border=0 ></a></td></tr>"
							  JSCodeStr = JSCodeStr &"<tr><td align=center>"& NaviPic &"<a class="""&TitleCSS&""" href=""" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) & """" & OpenMode & ">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a></td></tr></table></td>"
							Else
							  JSCodeStr = JSCodeStr &"<td align=center valign=middle><a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&"><img src="&DoMainStr&RsCreateObj("NewsSmallPicFile")&" height="&PicHeight&" width="&PicWidth&" border=0 ></a></td>"
							End If
						End If
					  End If
					  RsCreateObj.MoveNext
					  if i mod Cint(RowNum) = 0 or RsCreateObj.eof then
						if RightDate = true then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum*2&""" height=""1"" background="""  & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum&""" height=""1"" background="""  & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						end if
					  end if
					next 
					If RsSysJsObj("FileType")=1 then
					Set RsTempClassObjs = Conn.Execute("Select SavePath,ClassEName,FileExtName from FS_NS_NewsClass where ClassID='"&NoSqlHack(ClassID)&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在."
							Exit Function
						End If
					End If
					If RightDate = true then
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SavePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" height="&RowSpace&"></td></tr>"
						end if
					Else
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SavePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" height="&RowSpace&"></td></tr>"
						end if
					End If
					JSCodeStr = JSCodeStr & "</table>');"
					WriteFile SavePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_NS_Sysjs Set AddTime='"&Now()&"' where FileName='"&NoSqlHack(FileName)&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻.')"
					WriteFile SavePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_NS_Sysjs Set AddTime='"&Now()&"' where FileName='"&NoSqlHack(FileName)&"'")
					CreateSysJS = "文件操作成功.但未找到符合条件的新闻。"
				End If
			Case "NewNews"
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&CintStr(NewsNum)&" News.*," & SQLField & " From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.ClassID=Class.ClassID And News.ClassID in ('"&ClassIDStr&"') and News.isRecyle=0 and News.isLock=0 order by News.AddTime desc" '最新新闻
				else
					RsCreateSql = "Select top "&CintStr(NewsNum)&" News.*," & SQLField & " From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.ClassID=Class.ClassID And News.isRecyle=0 and News.isLock=0 order by News.AddTime desc" '最新新闻
				end if
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&RowSpace&"""><tr>"
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassName,SavePath from FS_NS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						Else
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						End IF
					  Else
						If ShowClassTF = true then
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						Else
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						End If
					  End If
					  RsCreateObj.MoveNext
					  if i mod Cint(RowNum) = 0 or RsCreateObj.eof then
						if RightDate = true then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum*2&""" height=""1"" background="""  & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum&""" height=""1"" background="""  & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						end if
					  end if
					next 
					If RsSysJsObj("FileType")=1 then
					Set RsTempClassObjs = Conn.Execute("Select SavePath,ClassEName,FileExtName from FS_NS_NewsClass where ClassID='"&NoSqlHack(ClassID)&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在."
							Exit Function
						End If
					End If
					If RightDate = true then
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" align=Right><a class="""&LinkCSS&""" href="""&GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SavePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" height="&RowSpace&"></td></tr>"
						end if
					Else
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SavePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" height="&RowSpace&"></td></tr>"
						end if
					End If
					JSCodeStr = JSCodeStr & "</table>');"
					WriteFile SavePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_NS_Sysjs Set AddTime='"&Now()&"' where FileName='"&NoSqlHack(FileName)&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SavePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_NS_Sysjs Set AddTime='"&Now()&"' where FileName='"&NoSqlHack(FileName)&"'")
					CreateSysJS = "文件操作成功.但未找到符合条件的新闻。"
				End If
				
			Case "HotNews"
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&CintStr(NewsNum)&" News.*," & SQLField & " From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.ClassID in ('"&FormatStrArr(ClassIDStr)&"') and News.isRecyle=0 and News.Hits>"&setHitsValue&" and News.isLock=0 order by News.Hits desc,News.ID DESC" '热点新闻
				else
					RsCreateSql = "Select top "&CintStr(NewsNum)&" News.*," & SQLField & " From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.isRecyle=0 and News.isLock=0 and News.Hits>"&setHitsValue&" order by News.Hits desc,News.ID DESC" '热点新闻
				end If
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&RowSpace&"""><tr>"
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassName,SavePath from FS_NS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						Else
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						End IF
					  Else
						If ShowClassTF = true then
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						Else
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						End If
					  End If
					  RsCreateObj.MoveNext
					  if i mod Cint(RowNum) = 0 or RsCreateObj.eof then
						if RightDate = true then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum*2&""" height=""1"" background="""  & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum&""" height=""1"" background="""  & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						end if
					  end if
					next 
					If RsSysJsObj("FileType")=1 then
					Set RsTempClassObjs = Conn.Execute("Select SavePath,ClassEName,FileExtName from FS_NS_NewsClass where ClassID='"&ClassID&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在."
							Exit Function
						End If
					End If
					If RightDate = true then
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SavePath"),RsTempClassObjs("FileExtName")) &""">"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" height="&RowSpace&"></td></tr>"
						end if
					Else
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SavePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" height="&RowSpace&"></td></tr>"
						end if
					End If
					JSCodeStr = JSCodeStr & "</table>');"
					WriteFile SavePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_NS_Sysjs Set AddTime='"&Now()&"' where FileName='"&NoSqlHack(FileName)&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SavePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_NS_Sysjs Set AddTime='"&Now()&"' where FileName='"&NoSqlHack(FileName)&"'")
					CreateSysJS = "文件操作成功.但未找到符合条件的新闻。"
				End If
				
				
				
			Case "ProclaimNews"
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&CintStr(NewsNum)&" News.*," & SQLField & " From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.ClassID in ('"&FormatStrArr(ClassIDStr)&"') and News.isRecyle=0 and "& CharIndexStr &"(News.NewsProperty,19,1)='1' and News.isLock=0 order by News.AddTime desc" '公告新闻
				else
					RsCreateSql = "Select top "&CintStr(NewsNum)&" News.*," & SQLField & " From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.isRecyle=0 and "& CharIndexStr &"(News.NewsProperty,19,1)='1' and News.isLock=0 order by News.AddTime desc" '公告新闻
				end if
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<marquee onmouseout=start() onmouseover=stop() Width="&MarWidth&" Height="&MarHeight&"  scrolldelay=80 direction="&MarDirection&" scrollamount="& CInt(MarSpeed) &"><font color=red>【公告】</font>"&BrStr
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassName,SavePath from FS_NS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							If RsCreateObj("IsURL") <> 1 then
							  JSCodeStr = JSCodeStr & NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href="""& FS_JsObj.GetOneNewsLinkURL(RsCreateObj) &""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&BrStr
							  JSCodeStr = JSCodeStr & ListSpaceStrs & Left(Replace(Replace(Lose_Html(RsCreateObj("Content")),chr(13) & chr(10),""),"&nbsp;",""),100) & "..." & ListSpaceStrs&BrStr
							Else
							  JSCodeStr = JSCodeStr & NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href="""&RsCreateObj("HeadNewsPath")&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&BrStr
							End If
						Else
							If RsCreateObj("IsURL") <> 1 then
							  JSCodeStr = JSCodeStr & NaviPic &"<a class="""&TitleCSS&""" href=""" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) &""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&BrStr
							  JSCodeStr = JSCodeStr & ListSpaceStrs & Left(Replace(Replace(Lose_Html(RsCreateObj("Content")),chr(13) & chr(10),""),"&nbsp;",""),100) & "..." & ListSpaceStrs&BrStr
							Else
							  JSCodeStr = JSCodeStr & NaviPic &"<a class="""&TitleCSS&""" href="""&RsCreateObj("HeadNewsPath")&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&BrStr
							End If
						End IF
					  Else
						If ShowClassTF = true then
							If RsCreateObj("IsURL") <> 1 then
							  JSCodeStr = JSCodeStr & NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href=""" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) &""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&BrStr
							  JSCodeStr = JSCodeStr & ListSpaceStrs & Left(Replace(Replace(Lose_Html(RsCreateObj("Content")),chr(13) & chr(10),""),"&nbsp;",""),100) & "..." & ListSpaceStrs&BrStr
							Else
							  JSCodeStr = JSCodeStr & NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href="""&RsCreateObj("HeadNewsPath")&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&BrStr
							End If
						Else
							If RsCreateObj("IsURL") <> 1 then
							  JSCodeStr = JSCodeStr & NaviPic &"<a class="""&TitleCSS&""" href=""" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) &""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&BrStr
							  JSCodeStr = JSCodeStr & ListSpaceStrs & Left(Replace(Replace(Lose_Html(RsCreateObj("Content")),chr(13) & chr(10),""),"&nbsp;",""),100) & "..." & ListSpaceStrs&BrStr
							Else
							  JSCodeStr = JSCodeStr & NaviPic &"<a class="""&TitleCSS&""" href="""&RsCreateObj("HeadNewsPath")&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&BrStr
							End If
						End If
					  End If
					  RsCreateObj.MoveNext
					next 
					If RsSysJsObj("FileType")=1 then
					Set RsTempClassObjs = Conn.Execute("Select SavePath,ClassEName,FileExtName from FS_NS_NewsClass where ClassID='"&NoSqlHack(ClassID)&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在."
							Exit Function
						End If
					End If
					if RsSysJsObj("FileType")=1 and MoreContentTF=True then
						JSCodeStr = JSCodeStr &"<a class="""&LinkCSS&""" href="""&GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SavePath"),RsTempClassObjs("FileExtName"))&""">"& MoreContentStr&"</a>"&ListSpaceStrs
					end if
					JSCodeStr = JSCodeStr & "</marquee>');"
					WriteFile SavePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_NS_Sysjs Set AddTime='"&Now()&"' where FileName='"&NoSqlHack(FileName)&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SavePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_NS_Sysjs Set AddTime='"&Now()&"' where FileName='"&NoSqlHack(FileName)&"'")
					CreateSysJS = "文件操作成功.但未找到符合条件的新闻。"
				End If
			Case Else 
				
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&CintStr(NewsNum)&" News.*," & SQLField & " From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.ClassID in ('"&FormatStrArr(ClassIDStr)&"') and News.isRecyle=0 and News.isLock=0 order by News.AddTime desc" '最新新闻
				else
					RsCreateSql = "Select top "&CintStr(NewsNum)&" News.*," & SQLField & " From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.isRecyle=0 and News.isLock=0 order by News.AddTime desc" '最新新闻
				end if
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&RowSpace&"""><tr>"
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassName,SavePath from FS_NS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href=""" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) & """" & OpenMode & ">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href=""" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) &""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						Else
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href=""" & FS_JsObj.GetOneNewsLinkURL(RsCreateObj) &""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&FS_JsObj.DateFormat(RsCreateObj("AddTime"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						End IF
					  Else
						If ShowClassTF = true then
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassName")&"]"&"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						Else
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&FS_JsObj.GetOneNewsLinkURL(RsCreateObj)&""""&OpenMode&">"&FS_JsObj.GotTopic(Lose_Html(RsCreateObj("NewsTitle")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						End If
					  End If
					  RsCreateObj.MoveNext
					  if i mod Cint(RowNum) = 0 or RsCreateObj.eof then
						if RightDate = true then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum*2&""" height=""1"" background="""  & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum&""" height=""1"" background="""  & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						end if
					  end if
					next 
					If RsSysJsObj("FileType")=1 then
					Set RsTempClassObjs = Conn.Execute("Select SavePath,ClassEName,FileExtName from FS_NS_NewsClass where ClassID='"&NoSqlHack(ClassID)&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在."
							Exit Function
						End If
					End If
					If RightDate = true then
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" align=Right><a class="""&LinkCSS&""" href=""" & FS_JsObj.GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SavePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" height="&RowSpace&"></td></tr>"
						end if
					Else
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" align=Right><a class="""&LinkCSS&""" href=""" & FS_JsObj.GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SavePath"),RsTempClassObjs("FileExtName")) &""">"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" height="&RowSpace&"></td></tr>"
						end if
					End If
					JSCodeStr = JSCodeStr & "</table>');"
					WriteFile SavePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_NS_Sysjs Set AddTime='"&Now()&"' where FileName='"&NoSqlHack(FileName)&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SavePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_NS_Sysjs Set AddTime='"&Now()&"' where FileName='"&NoSqlHack(FileName)&"'")
					CreateSysJS = "文件操作成功.但未找到符合条件的新闻。"
				End If
		End Select
	Else
		CreateSysJS = "参数传递错误."
	End If
	set FS_JsObj=nothing
	RsSysJsObj.Close
	Set RsSysJsObj = Nothing
End Function

Function WriteFile(SaveFilePath,FileNameStr,JSCodeStr)
	'On Error ReSume Next
	Dim MyFile,CrHNJS
	Set MyFile=Server.CreateObject(G_FS_FSO)
	If MyFile.FolderExists(Server.MapPath(SaveFilePath))=false then
		MyFile.CreateFolder(Server.MapPath(SaveFilePath))
	End If
	If MyFile.FileExists(Server.MapPath(SaveFilePath)&"/"& FileNameStr &".js") then
		MyFile.DeleteFile(Server.MapPath(SaveFilePath)&"/"& FileNameStr &".js")
	End if
	Set CrHNJS=MyFile.CreateTextFile(Server.MapPath( SaveFilePath) &"/"& FileNameStr &".js")
		CrHNJS.write JSCodeStr
	Set MyFile=nothing
End Function
%>
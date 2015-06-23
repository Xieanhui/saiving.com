<%
Class Cls_Js
	private ns_id,ns_ename,ns_cname,ns_js_type,ns_manner,ns_picWidth,ns_picHeight,ns_newsNum,ns_newsTitleNum,ns_titleCSS,ns_contentCSS
	private ns_backCSS,ns_rowNum,ns_picPath,ns_addTime,ns_showTimeTF,ns_contentNum,ns_naviPic,ns_dateType,ns_dateCss,ns_info
	private ns_moreContent,ns_LinkWord,ns_LinkCSS,ns_rowSpace,ns_rowBettween,ns_openMode
	private m_NewsFields,m_NewsTable,m_NewsWhere
	Private TempSysRootDir
	Private ListSpace,ListSpaceStr,Temp_i,TableCellSpace,TitleSpace,TitleSpaceStr,MoreContentStr
	
	Private Sub Class_Initialize()
		m_NewsFields = "News.ID,NewsID,PopId,News.ClassID,SpecialEName,NewsTitle,CurtTitle,NewsNaviContent,isShowReview,TitleColor,titleBorder,TitleItalic,News.IsURL,News.URLAddress,Content,isPicNews,NewsPicFile,NewsSmallPicFile,PicborderCss,News.Templet,News.isPop,Source,Editor,Keywords,Author,SaveNewsPath,News.FileName,News.FileExtName,NewsProperty,TodayNewsPic,isLock,isRecyle,News.addtime,ClassEName,[Domain],SavePath"
		m_NewsTable = "FS_NS_News as News,FS_NS_NewsClass as Class"
		m_NewsWhere = "News.ClassID=Class.ClassID"
	End Sub
	'�������js�Ĳ���
	public Function getFreeJsParam(jsid)
		Dim F_FreeJsParam_Rs,sql_statement
		Set F_FreeJsParam_Rs=Server.CreateObject(G_FS_RS)
		sql_statement="select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode from FS_NS_FreeJS where id="&CintStr(jsid)
		F_FreeJsParam_Rs.open sql_statement,Conn,1,1
		if not F_FreeJsParam_Rs.eof and not F_FreeJsParam_Rs.bof then
			ns_id=F_FreeJsParam_Rs("ID")
			ns_ename=F_FreeJsParam_Rs("EName")
			ns_cname=F_FreeJsParam_Rs("CName")
			ns_js_type=F_FreeJsParam_Rs("Type")
			ns_manner=F_FreeJsParam_Rs("Manner")
			ns_picWidth=F_FreeJsParam_Rs("PicWidth")
			ns_picHeight=F_FreeJsParam_Rs("PicHeight")
			ns_newsNum=F_FreeJsParam_Rs("NewsNum")
			ns_newsTitleNum=F_FreeJsParam_Rs("NewsTitleNum")
			ns_titleCSS=F_FreeJsParam_Rs("TitleCSS")
			ns_contentCSS=F_FreeJsParam_Rs("ContentCSS")
			ns_backCSS=F_FreeJsParam_Rs("BackCSS")
			ns_rowNum=F_FreeJsParam_Rs("RowNum")
			ns_picPath=F_FreeJsParam_Rs("PicPath")
			ns_addTime=F_FreeJsParam_Rs("AddTime")
			ns_showTimeTF=F_FreeJsParam_Rs("ShowTimeTF")
			ns_contentNum=F_FreeJsParam_Rs("ContentNum")
			ns_naviPic=F_FreeJsParam_Rs("NaviPic")
			ns_dateType=F_FreeJsParam_Rs("DateType")
			ns_dateCss=F_FreeJsParam_Rs("DateCSS")
			ns_info=F_FreeJsParam_Rs("Info")
			ns_moreContent=F_FreeJsParam_Rs("MoreContent")
			ns_LinkWord=F_FreeJsParam_Rs("LinkWord")
			ns_LinkCSS=F_FreeJsParam_Rs("LinkCSS")
			ns_rowSpace=F_FreeJsParam_Rs("RowSpace")
			ns_rowBettween=F_FreeJsParam_Rs("RowBettween")
			ns_openMode=F_FreeJsParam_Rs("OpenMode")
		End if
	End Function
	'��ֵ
	public Property get id()'Free JS id
		id=ns_id
	End Property 
	
	public Property get ename()' Free Js Ӣ����
		ename=ns_ename
	End Property
	
	public Property get cname()' Free Js ������
		cname=ns_cname
	End Property
	
	public Property get js_type()' ����(0Ϊ����,1ΪͼƬ)
		js_type=ns_js_type
	End Property
	
	public Property get manner()' ��ʽ(1-5 Ϊ������ʽ,6-17ΪͼƬ��ʽ)(��)
		manner=ns_manner
	End Property
	
	public Property get picWidth()' ͼƬ���
		picWidth=ns_picWidth
	End Property
	
	public Property get picHeight()' ͼƬ�߶�
		picHeight=ns_picHeight
	End Property
	
	public Property get newsNum()' ���������������
		newsNum=ns_newsNum
	End Property
	
	public Property get newsTitleNum()'���ű�������
		newsTitleNum=ns_newsTitleNum
	End Property

	public Property get titleCSS()' ���ű�����ʽ
		titleCSS=ns_titleCSS
	End Property
	
	public Property get contentCSS()' ����������ʽ
		contentCSS=ns_contentCSS
	End Property
	
	public Property get backCSS() 'JS������ʽ
		backCSS=ns_backCSS
	End Property
	
	public Property get rowNum()' ÿ�в�������(����Ϊ��0��)
		rowNum=ns_rowNum
	End Property
	
	public Property get picPath()' Ϊĳ����ʽ����
		picPath=ns_picPath
	End Property
	
	public Property get addTime()' Free Js���ʱ��
		addTime=ns_addTime
	End Property
	
	public Property get showTimeTF()'�Ƿ������ű������ʾ����ʱ��(��0��Ϊ��,��1�� Ϊ��)
		showTimeTF=ns_showTimeTF
	End Property
	
	public Property get contentNum()' Free Js ������������
		contentNum=ns_contentNum
	End Property
	
	public Property get naviPic()' Free Js ���ű��⵼��ͼƬ
		naviPic=ns_naviPic
	End Property
	
	public Property get dateType()' Free Js ������ʽ(1-15)
		dateType=ns_dateType
	End Property
	
	public Property get dateCSS()' Free Js ����CSS��ʽ
		dateCSS=ns_dateCSS
	End Property
	
	public Property get info()' Free Js ��ע
		info=ns_info
	End Property
	
	public Property get moreContent()' Free Js ��������(��������)
		moreContent=ns_moreContent
	End Property
	
	public Property get linkWord()' Free Js �������ֻ���ͼƬ
		linkWord=ns_linkWord
	End Property
	
	public Property get linkCSS()' Free Js ������ʽ��
		linkCSS=ns_linkCSS
	End Property
		
	public Property get rowSpace()' Free Js �����о�
		rowSpace=ns_rowSpace
	End Property
	
	public Property get rowBettween()'�м�ͼƬ
		rowBettween=ns_rowBettween
	End Property
	
	public Property get openMode()'���ڴ򿪷�ʽ
		openMode=ns_openMode
	End Property
	'----------------------------------------------
	Public Property Let SysRootDir(ExteriorValue)
		TempSysRootDir = ExteriorValue
	End Property
	'----------------------------------------------
	Public Function GetOneNewsLinkURL(f_RS)
		Dim f_NewsLinkRecordSet,f_NewsLink
		Set f_NewsLinkRecordSet = New CLS_FoosunRecordSet
		Set f_NewsLinkRecordSet.Values("ClassEName,Domain,SavePath,IsURL,URLAddress,SaveNewsPath,FileName,FileExtName") = f_RS
		Set f_NewsLink = New CLS_FoosunLink
		GetOneNewsLinkURL = f_NewsLink.NewsLink(f_NewsLinkRecordSet)
		Set f_NewsLink = Nothing
		Set f_NewsLinkRecordSet = Nothing
	End Function
	'-----����JSʱ���ʽ��
	Public Function  DateFormat(DateStr,Types)
		Dim DateString
		if IsDate(DateStr) = False then
			DateString = ""
		end if
		Select Case Types
		  Case "1" 
			  DateString = Year(DateStr)&"-"&Month(DateStr)&"-"&Day(DateStr)
		  Case "2"
			  DateString = Year(DateStr)&"."&Month(DateStr)&"."&Day(DateStr)
		  Case "3"
			  DateString = Year(DateStr)&"/"&Month(DateStr)&"/"&Day(DateStr)
		  Case "4"
			  DateString = Month(DateStr)&"/"&Day(DateStr)&"/"&Year(DateStr)
		  Case "5"
			  DateString = Day(DateStr)&"/"&Month(DateStr)&"/"&Year(DateStr)
		  Case "6"
			  DateString = Month(DateStr)&"-"&Day(DateStr)&"-"&Year(DateStr)
		  Case "7"
			  DateString = Month(DateStr)&"."&Day(DateStr)&"."&Year(DateStr)
		  Case "8"
			  DateString = Month(DateStr)&"-"&Day(DateStr)
		  Case "9"
			  DateString = Month(DateStr)&"/"&Day(DateStr)
		  Case "10"
			  DateString = Month(DateStr)&"."&Day(DateStr)
		  Case "11"
			  DateString = Month(DateStr)&"��"&Day(DateStr)&"��"
		  Case "12"
			  DateString = Day(DateStr)&"��"&Hour(DateStr)&"ʱ"
		  case "13"
			  DateString = Day(DateStr)&"��"&Hour(DateStr)&"��"
		  Case "14"
			  DateString = Hour(DateStr)&"ʱ"&Minute(DateStr)&"��"
		  Case "15"
			  DateString = Hour(DateStr)&":"&Minute(DateStr)
		  Case "16"
			  DateString = Year(DateStr)&"��"&Month(DateStr)&"��"&Day(DateStr)&"��"
		  Case Else
			  DateString = DateStr
		 End Select
		 DateFormat = DateString
	 End Function
 '---------------------------------------------
	Public Function LoseHtml(ContentStr)
		Dim ClsTempLoseStr,regEx
		ClsTempLoseStr = ContentStr&""
		Set regEx = New RegExp
		regEx.Pattern = "<\/*[^<>]*>"
		regEx.IgnoreCase = True
		regEx.Global = True
		ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
		LoseHtml = ClsTempLoseStr
	End Function
'---------------------------------------------
	Function GotTopic(Str,StrLen)
		Dim l,t,c,i
		If StrLen=0 then
			GotTopic=""
			exit function
		End If
		if IsNull(Str) then 
			GotTopic = ""
			Exit Function
		end if
		if Str = "" then
			GotTopic=""
			Exit Function
		end if
		Str=Replace(Replace(Replace(Replace(Str,"&nbsp;"," "),"&quot;",Chr(34)),"&gt;",">"),"&lt;","<")
		l=len(str)
		t=0
		strlen=Clng(strLen)
		for i=1 to l
			c=Abs(Asc(Mid(str,i,1)))
			if c>255 then
				t=t+2
			else
				t=t+1
			end if
			if t>=strlen then
				GotTopic=left(str,i)
				exit for
			else
				GotTopic=str
			end if
		next
		GotTopic = Replace(Replace(Replace(Replace(GotTopic," ","&nbsp;"),Chr(34),"&quot;"),">","&gt;"),"<","&lt;")
	end Function
'-----------------------------------------------------
	Private Function ListTitle(TitleStr,TitleNum)
	   Dim ClsTitleStr,ClsTitleNum,i,j,ClsTempNum,k,ClsTitleStrResult,LeftStr,RightStr
		   ClsTitleNum = Cint(TitleNum)
		   ClsTempNum = Len(Cstr(TitleStr))
		   if ClsTitleNum > ClsTempNum then
			   ClsTitleNum = ClsTempNum
		   end if
		   ClsTitleStr = Left(Cstr(TitleStr),ClsTitleNum)
		   Dim TempStr
		   For i = 1 to ClsTitleNum - 1
			   TempStr = TempStr & Mid(ClsTitleStr,i,1) & "<br>"
		   Next
		   TempStr = TempStr & Right(ClsTitleStr,1)
		   ListTitle = TempStr
	End Function
 '���ɺ���
	Public Function WCssA(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&NoSqlHack(EName)&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then
				JSCodeStr = JSCodeStr & "<td>��JS����������</td>"
			  End If
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
				  if Not ClsNewsObj.Eof then
					  If ClsJSObj("ShowTimeTF")="1" then
						  JSCodeStr = JSCodeStr &"<td valign=middle ><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href=http://"&Replace(Conn.execute("SELECT MF_Domain FROM FS_MF_Config")(0),"/"&G_VIRTUAL_ROOT_DIR,"") & GetOneNewsLinkURL(ClsNewsObj) &" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a></td><td><Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddTime"),""&ClsJSObj("DateType")&"")&"</Span></td>"
					  Else
						  JSCodeStr = JSCodeStr &"<td valign=middle><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href=http://"&Replace(Conn.execute("SELECT MF_Domain FROM FS_MF_Config")(0),"/"&G_VIRTUAL_ROOT_DIR,"") & GetOneNewsLinkURL(ClsNewsObj) &" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a></td>"
					  End If
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					if ClsJSObj("ShowTimeTF")=1 then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))*2&""" height="""&ClsJSObj("RowSpace")&""" background="""& ClsJSObj("RowBettween")&"""></td></tr><tr>"
					else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""& ClsJSObj("RowBettween")&"""></td></tr><tr>"
					end if
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  Set MyFile=Server.CreateObject(G_FS_FSO)
			  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
			  CrHNJS.write JSCodeStr
			  Set MyFile=nothing
			  '---------
			  ClsJSObj.Close
			  Set ClsJSObj = Nothing
			Else
				WCssA = "����JS�ļ�ʱδ�ҵ�����"
			End If
	End Function 

	Public Function WCssB(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&NoSqlHack(EName)&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>��JS����������</td>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace"))
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
				 if Not ClsNewsObj.Eof then
					  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj)
					  If ClsJSObj("ShowTimeTF")=1 then
						  JSCodeStr = JSCodeStr &"<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href=" & NewsLinkStr &" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a></td><td><Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddTime"),""&ClsJSObj("DateType")&"")&"</Span></td><td rowspan=2>"&ListSpaceStr&"</td></tr>"
					  Else
						  JSCodeStr = JSCodeStr &"<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href=" & NewsLinkStr &" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a></td><td rowspan=2>"&ListSpaceStr&"</td></tr>"
					  End If
					  If ClsJSObj("ShowTimeTF")=1 then
							If  ClsJSObj("MoreContent")=1 then
								JSCodeStr = JSCodeStr & "<tr><td colspan=2><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
							Else
								JSCodeStr = JSCodeStr & "<tr><td colspan=2><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td></tr></table></td>"
							End If
					  Else
							If  ClsJSObj("MoreContent")=1 then
								JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
							Else
								JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td></tr></table></td>"
							End If
					  End If
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			  else
				  WCssB = JSCodeStr
			  end if
			Else
				WCssB = "����JS�ļ�ʱδ�ҵ�����"
			End If
	End Function 

	Public Function WCssC(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>��JS����������</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace"))
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
				 if Not ClsNewsObj.Eof then
					  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj)
					  If ClsJSObj("ShowTimeTF")="1" then
						  JSCodeStr = JSCodeStr &"<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top""><div align=""center""><img src="""&ClsJSObj("NaviPic")&""" /><br><Span class="""&ClsJSObj("TitleCSS")&"""><a href="""&NewsLinkStr&""">"&ListTitle(GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum")),ClsJSObj("NewsTitleNum"))&"</a></Span><br><Span class="""&ClsJSObj("DateCSS")&""">"&ListTitle(DateFormat(ClsNewsObj("AddTime"),""&ClsJSObj("DateType")&""),50)&"</Span></div></td>"
					  Else
						  JSCodeStr = JSCodeStr &"<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top""><div align=""center""><img src="""&ClsJSObj("NaviPic")&""" /><br><Span class="""&ClsJSObj("TitleCSS")&"""><a href="""&NewsLinkStr&""">"&ListTitle(GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum")),ClsJSObj("NewsTitleNum"))&"</a></Span></div></td>"
					  End If
					  If  ClsJSObj("MoreContent")=1 then
						  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href=" & NewsLinkStr&"."&ClsNewsObj("FileExtName")&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td><td>"&ListSpaceStr&"</td></tr></table></td>"
					  Else
						  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td><td>"&ListSpaceStr&"</td></tr></table></td>"
					  End If
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
				else
					WCssC = JSCodeStr
				end if
			Else
				WCssC = "����JS�ļ�ʱδ�ҵ�����"
			End If
	End Function 

	Public Function WCssD(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&NoSqlHack(EName)&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>��JS����������</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace"))
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
				  if Not ClsNewsObj.Eof then
					  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj)
					  If  ClsJSObj("MoreContent")=1 then
						  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top""><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td>"
					  Else
						  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top""><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td>"
					  End If
					  If ClsJSObj("ShowTimeTF")="1" then
						  JSCodeStr = JSCodeStr &"<td valign=""top""><div align=""center""><img src="""&ClsJSObj("NaviPic")&""" /><br><Span class="""&ClsJSObj("TitleCSS")&"""><a href="""&NewsLinkStr&""">"&ListTitle(GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum")),ClsJSObj("NewsTitleNum"))&"</a></Span><br><Span class="""&ClsJSObj("DateCSS")&""">"&ListTitle(DateFormat(ClsNewsObj("AddTime"),""&ClsJSObj("DateType")&""),50)&"</Span></div></td><td>"&ListSpaceStr&"</td></tr></table></td>"
					  Else
						  JSCodeStr = JSCodeStr &"<td valign=""top""><div align=""center""><img src="""&ClsJSObj("NaviPic")&""" /><br><Span class="""&ClsJSObj("TitleCSS")&"""><a href="""&NewsLinkStr&""">"&ListTitle(GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum")),ClsJSObj("NewsTitleNum"))&"</a></Span></div></td><td>"&ListSpaceStr&"</td></tr></table></td>"
					  End If
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
				else
					WCssD = JSCodeStr
				end if
			Else
				WCssD = "����JS�ļ�ʱδ�ҵ�����"
			End If
	End Function

	Public Function WCssE(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&NoSqlHack(EName)&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>��JS����������</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace"))
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
				  if Not ClsNewsObj.Eof then
					  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj)
					  If  ClsJSObj("MoreContent")=1 then
						  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td><td rowspan=2>"&ListSpaceStr&"</td></tr>"
					  Else
						  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td><td rowspan=2>"&ListSpaceStr&"</td></tr>"
					  End If
					  If ClsJSObj("ShowTimeTF")=1 then
						  JSCodeStr = JSCodeStr &"<tr><td><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a>&nbsp;<Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddTime"),""&ClsJSObj("DateType")&"")&"</Span></td></tr></table></td>"
					  Else
						  JSCodeStr = JSCodeStr &"<tr><td><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a></td></tr></table></td>"
					  End If
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			  else
				  WCssE = JSCodeStr
			  end if
			Else
				WCssE = "����JS�ļ�ʱδ�ҵ�����"
			End If
	End Function

	Public Function PCssA(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr,f_NewsPicFile
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&NoSqlHack(EName)&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>��JS����������</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
				  if Not ClsNewsObj.Eof then
					  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj)
					  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td align=""center""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td></tr>"
					 ' f_NewsPicFile = ClsJSObj("NaviPic")
					  f_NewsPicFile = ClsNewsObj("NewsSmallPicFile")
					  If ClsJSObj("ShowTimeTF")="1" then
						  JSCodeStr = JSCodeStr & "<tr><td align=""center""><img src="""&f_NewsPicFile&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a>&nbsp;<Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddTime"),""&ClsJSObj("DateType")&"")&"</Span></td></tr></table></td>"
					  Else
						  JSCodeStr = JSCodeStr & "<tr><td align=""center""><img src="""&f_NewsPicFile&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a></td></tr></table></td>"
					  End If
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			  else
				  PCssA = JSCodeStr
			  end if
			Else
				PCssA = "����JS�ļ�ʱδ�ҵ�����"
			End If
	End Function

	Public Function PCssB(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&NoSqlHack(EName)&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>��JS����������</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace")\2)
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
				  if Not ClsNewsObj.Eof then
					  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj)
					  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top"" align=""center"" rowspan=""2""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td><td rowspan=""2"">"&ListSpaceStr&"</td>"
					  If ClsJSObj("ShowTimeTF")="1" then
						  JSCodeStr = JSCodeStr & "<td align=""left""><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a>&nbsp;<Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddTime"),""&ClsJSObj("DateType")&"")&"</Span></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
					  Else
						'response.write ClsNewsObj("NewsTitle")&"---"&ClsJSObj("NewsTitleNum")
						  JSCodeStr = JSCodeStr & "<td align=""left""><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
					  End If
					  If  ClsJSObj("MoreContent")=1 Then
						  JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
					  Else
						  JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td></tr></table></td>"
					  End If
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(Replace(JSCodeStr,"<tr></tr>",""),"&nbsp;&nbsp;&nbsp;&nbsp;","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssB = JSCodeStr
			   end if
			Else
				PCssB = "����JS�ļ�ʱδ�ҵ�����"
			End If
	End Function

	Public Function PCssC(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>��JS����������</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
				  if Not ClsNewsObj.Eof then
					  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj)
					  JSCodeStr = JSCodeStr & "<td align=""center""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td>"
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssC = JSCodeStr
			   end if
			Else
				PCssC = "����JS�ļ�ʱδ�ҵ�����"
			End If
	End Function

	Public Function PCssD(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&NoSqlHack(EName)&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>��JS����������</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace"))
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
				  if Not ClsNewsObj.Eof then
					  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj)
					  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top"" colspan=""2"" align=""center""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
					  If ClsJSObj("ShowTimeTF")="1" then
						  JSCodeStr = JSCodeStr &"<tr><td valign=""top""><div align=""center""><img src="""&ClsJSObj("NaviPic")&""" /><br><Span class="""&ClsJSObj("TitleCSS")&"""><a href="""&NewsLinkStr&""">"&ListTitle(GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum")),ClsJSObj("NewsTitleNum"))&"</a></Span><br><Span class="""&ClsJSObj("DateCSS")&""">"&ListTitle(DateFormat(ClsNewsObj("AddTime"),""&ClsJSObj("DateType")&""),50)&"</Span></div></td>"
					  Else
						  JSCodeStr = JSCodeStr &"<tr><td valign=""top""><div align=""center""><img src="""& ClsJSObj("NaviPic") &""" border=""0""><br><Span class="""&ClsJSObj("TitleCSS")&"""><a href="""&NewsLinkStr&""">"&ListTitle(GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum")),ClsJSObj("NewsTitleNum"))&"</a></Span></div></td>"
					  End If
					  If  ClsJSObj("MoreContent")=1 then
						  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"......</Span><br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
					  Else
						  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"......</Span></td></tr></table></td>"
					  End If
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  JSCodeStr = Replace(Replace(JSCodeStr,"src='","src="),"' border"," border")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssD = JSCodeStr
			   end if
			Else
				PCssD = "����JS�ļ�ʱδ�ҵ�����"
			End If
	End Function

	Public Function PCssE(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>��JS����������</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace")\2)
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
				  if Not ClsNewsObj.Eof then
					  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj)
					  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td rowspan=""2"">"&ListSpaceStr&"</td><td valign=""top"" align=""center""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
					  If  ClsJSObj("MoreContent")=1 then
						  JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
					  Else
						  JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td></tr></table></td>"
					  End If
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssE = JSCodeStr
			   end if
			Else
				PCssE = "����JS�ļ�ʱδ�ҵ�����"
			End If
	End Function

	Public Function PCssF(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>��JS����������</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace"))
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
				  if Not ClsNewsObj.Eof then
					  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj)
					  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top"" align=""center""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td>"
					  If  ClsJSObj("MoreContent")=1 then
						  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td><td>"&ListSpaceStr&"</td></tr></table></td>"
					  Else
						  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td><td>"&ListSpaceStr&"</td></tr></table></td>"
					  End If
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</table>');"
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssF = JSCodeStr
			   end if
			Else
				PCssF = "����JS�ļ�ʱδ�ҵ�����"
			End If
	End Function

	Public Function PCssG(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>��JS����������</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
			  JSCodeStr = JSCodeStr & "<td valign=""top"" align=""center"" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&"><img src="""& ClsJSObj("PicPath") &""" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></td><td><table width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
				  if Not ClsNewsObj.Eof then
					  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj)
					  if i mod Cint(ClsJSObj("RowNum")) = 0 and not ClsJSFileObj.eof then
						  If ClsJSObj("ShowTimeTF")="1" then
							  JSCodeStr = JSCodeStr & "<td><div align=""left""><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a>&nbsp;<Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddTime"),""&ClsJSObj("DateType")&"")&"</Span></div></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
						  Else
							  JSCodeStr = JSCodeStr & "<td><div align=""left""><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a></div></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
						  End If
					  end if
					  if ClsJSFileObj.eof then
						  If ClsJSObj("ShowTimeTF")="1" then
							  JSCodeStr = JSCodeStr & "<td><div align=""left""><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a>&nbsp;<Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddTime"),""&ClsJSObj("DateType")&"")&"</Span></div></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
						  Else
							  JSCodeStr = JSCodeStr & "<td><div align=""left""><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a></div></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
						  End If
					  end if
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table></td></tr></table>');"
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssG = JSCodeStr
			   end if
			Else
				PCssG = "����JS�ļ�ʱδ�ҵ�����"
			End If
	End Function

	Public Function PCssH(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&NoSqlHack(EName)&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>��JS����������</td>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace"))
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
				  if Not ClsNewsObj.Eof then
					  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj)
					  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top"" align=""left""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td>"
					  If ClsJSObj("ShowTimeTF")="1" then
						  JSCodeStr = JSCodeStr & "<td><div align=""left""><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a>&nbsp;<Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddTime"),""&ClsJSObj("DateType")&"")&"</Span></div></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
					  Else
						  JSCodeStr = JSCodeStr & "<td><div align=""left""><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a></div></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
					  End If
					  If  ClsJSObj("MoreContent")=1 then
						  JSCodeStr = JSCodeStr & "<tr><td colspan=""2""><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
					  Else
						  JSCodeStr = JSCodeStr & "<tr><td colspan=""2""><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td></tr></table></td>"
					  End If
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssH = JSCodeStr
			   end if
			Else
				PCssH = "����JS�ļ�ʱδ�ҵ�����"
			End If
	End Function

	Public Function PCssI(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&NoSqlHack(EName)&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>��JS����������</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace")\2)
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
				  if Not ClsNewsObj.Eof then
					  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj)
					  If ClsJSObj("ShowTimeTF")="1" then
						  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td rowspan=""2"">"&ListSpaceStr&"</td><td colspan=""3""><div align=""center""><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a>&nbsp;<Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddTime"),""&ClsJSObj("DateType")&"")&"</Span></div></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
					  Else
						  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td rowspan=""2"">"&ListSpaceStr&"</td><td colspan=""3""><div align=""center""><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a></div></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
					  End If
					  JSCodeStr = JSCodeStr & "<tr><td valign=""top""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td><td>&nbsp;</td>"			  
					  If  ClsJSObj("MoreContent")=1 then
						  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
					  Else
						  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td></tr></table></td>"
					  End If
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssI = JSCodeStr
			   end if
			Else
				PCssI = "����JS�ļ�ʱδ�ҵ�����"
			End If
	End Function

	Public Function PCssJ(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&NoSqlHack(EName)&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>��JS����������</td>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
			      if Not ClsNewsObj.Eof then
					  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj)
					  If ClsJSObj("ShowTimeTF")="1" then
						  JSCodeStr = JSCodeStr &"<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table background="""& ClsJSFileObj("PicPath")&""" width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a>&nbsp;<Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddTime"),""&ClsJSObj("DateType")&"")&"</Span></td></tr>"
					  Else
						  JSCodeStr = JSCodeStr &"<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table background="""& ClsJSFileObj("PicPath")&""" width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a></td></tr>"
					  End If
					  If  ClsJSObj("MoreContent")=1 then
						  JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
					  Else
						  JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td></tr></table></td>"
					  End If
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssJ = JSCodeStr
			   end if
			Else
				PCssJ = "����JS�ļ�ʱδ�ҵ�����"
			End If
	End Function

	Public Function PCssK(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode From FS_NS_FreeJS Where EName='"&NoSqlHack(EName)&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table background="""& ClsJSObj("PicPath")&""" width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="Select ID,Title,JSName,NewsID,PicPath,ClassID,NewsTime,ToJsTime,DelFlag From FS_NS_FreeJSFile where JSName='"&NoSqlHack(EName)&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then
				JSCodeStr = JSCodeStr & "<td>��JS����������</td>"
			  End If
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select " & m_NewsFields & " From " & m_NewsTable & " where " & m_NewsWhere & " And NewsID='"&ClsJSFileObj("NewsID")&"'")
				  if Not ClsNewsObj.Eof then
					  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj)
					  If ClsJSObj("ShowTimeTF")="1" then
						  JSCodeStr = JSCodeStr &"<td valign=middle><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a></td><td><Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddTime"),""&ClsJSObj("DateType")&"")&"</Span></td>"
					  Else
						  JSCodeStr = JSCodeStr &"<td valign=middle><img src="""&ClsJSObj("NaviPic")&""" /><a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("NewsTitle"),ClsJSObj("NewsTitleNum"))&"</a></td>"
					  End If
				  end if
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					if ClsJSObj("ShowTimeTF")=1 then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))*2&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
					else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
					end if
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				  End If			  
				  CreatePath Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/")),Server.MapPath(replace(replace("../../../"&TempSysRootDir,"///","/"),"//","/")) 
				Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(replace(replace("../../../"&TempSysRootDir&"/JS/FreeJs","///","/"),"//","/"))&"\"& EName &".js")
				CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssK = JSCodeStr
			   end if
			Else
				PCssK = "����JS�ļ�ʱδ�ҵ�����"
			End If
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
End Class
%>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->
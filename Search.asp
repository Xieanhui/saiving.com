<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<%session.CodePage="936"%>
<!--#include file="FS_Inc/Const.asp" -->
<!--#include file="FS_InterFace/MF_Function.asp" -->
<!--#include file="FS_Inc/Function.asp" -->
<!--#include file="FS_Inc/Func_Page.asp" -->
<!--#include file="FS_InterFace/MS_Public.asp" -->
<!--#include file="FS_InterFace/DS_Public.asp" -->
<!--#include file="FS_InterFace/CLS_Foosun.asp" -->
<%'Copyright (c) 2006 Foosun Inc. 
Server.ScriptTimeOut=999
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache"
response.Charset = "gb2312"
Dim starttime,endtime
starttime=timer()
function morestr(str,length)
	if len(str)>length then 
		morestr = left(str,length)&"<strong>...</strong>"
	else
		morestr = str       
	end if	  
end function
Dim Conn,Old_News_Conn,User_Conn,Search_Sql,Search_RS,strShowErr,Cookie_Domain,Cookie_Copyright,Cookie_eMail,Cookie_Site_Name
Dim Server_Name,Server_V1,Server_V2
Dim TmpStr,TmpArr,SqlDateType,FileSize,FileEditDate,TmpStr1,TmpStr2
Dim Keyword,s_type,SubSys,ClassId,s_date,e_date  ,GetType,AreaID,PubType
Dim ChildDomain,ClassPath
Dim LocalUrl,RDSqlDateType
Dim MsMinPric,MsMaxPric
Dim StringType
IF G_IS_SQL_DB = 1 Then
	StringType = "SubString"
Else
	StringType = "Mid"
End If	

GetType = request.QueryString("GetType")''内部
if GetType = "" then response.Write("请指定必要的参数.") : response.End()
''日期
If G_IS_SQL_DB = 1 Then  
	SqlDateType = "'"
else
	SqlDateType = "#"
end If
Function Get_MF_Config()
	if request.Cookies("FoosunSearchCookie")("Cookie_Domain") = Get_MF_Domain() then exit Function
	set Search_RS=Conn.execute("select top 1 MF_Domain,MF_Site_Name,MF_eMail,MF_Copyright_Info  from FS_MF_Config")
	Response.Cookies("FoosunSearchCookie")("Cookie_Domain")=Search_RS("MF_Domain") 
	Response.Cookies("FoosunSearchCookie")("Cookie_Copyright")=Search_RS("MF_Copyright_Info") 
	Response.Cookies("FoosunSearchCookie")("Cookie_eMail")=Search_RS("MF_eMail") 
	Response.Cookies("FoosunSearchCookie")("Cookie_Site_Name")=Search_RS("MF_Site_Name") 
	Response.Cookies("FoosunSearchCookie").Expires=Date()+1
	Search_RS.close
End Function

Function get_NewsLink(f_NewsLinkRecordSet)
	Dim f_NewsLink
	Set f_NewsLink = New CLS_FoosunLink
	get_NewsLink = f_NewsLink.NewsLink(f_NewsLinkRecordSet)
	Set f_NewsLink = Nothing
End Function

'得到信息链接地址
Function GetTheInfoLink(InfoID,InfoType)
	Dim LinkObj
	Select Case InfoType
		Case "DS"
			Set LinkObj = New cls_DS
			GetTheInfoLink = LinkObj.get_DownLink(InfoID)
			Set LinkObj = Nothing
		Case "MS"
			Set LinkObj = New cls_MS
			GetTheInfoLink = LinkObj.get_productsLink(InfoID)
			Set LinkObj = Nothing
	End Select
End Function

''得到相关表的值。
Function Get_OtherTable_Value(This_Fun_Sql)
	Dim This_Fun_Rs
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
	else
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	end if
	if instr(lcase(This_Fun_Sql)," in ")>0 then 
		do while not This_Fun_Rs.eof
			Get_OtherTable_Value = Get_OtherTable_Value & This_Fun_Rs(0) &"&nbsp;"
			This_Fun_Rs.movenext
		loop
	else			
		if not This_Fun_Rs.eof then 
			Get_OtherTable_Value = This_Fun_Rs(0)
		else
			Get_OtherTable_Value = ""
		end if
	end if	
	set This_Fun_Rs=nothing 
End Function

''++++++++++++++++++++++++++++++++++++
'检查本地文件 返回大小和修改日期
Function CheckFile(PhFileName)
	dim fsv1,fsv2
	fsv1="":fsv2=""
	On Error Resume Next
	if isnull(PhFileName) or PhFileName="" then CheckFile="|":exit Function
	Dim Fso,MyFile
	Set Fso = CreateObject(G_FS_FSO)
	If Left(LCase(PhFileName),7) = "http://" Then
		CheckFile="|":exit Function
	Else
		IF Left(PhFileName,1) <> "/" Then
			CheckFile="|":exit Function
		End If	
	End If	
	If Fso.FileExists(server.MapPath(PhFileName)) Then
		set MyFile = Fso.GetFile(server.MapPath(PhFileName))
		fsv1 = formatnumber(MyFile.Size/1024,2,-1)&"K"
		fsv2 = MyFile.DateLastModified
		set MyFile = nothing 
	End if
	if Err<>0 then CheckFile="|":exit Function		
	Set Fso = Nothing
	CheckFile = fsv1&"|"&fsv2
End Function

MF_Default_Conn
MF_User_Conn
MF_Old_News_Conn

Get_MF_Config

SubSys_Cookies:MFConfig_Cookies:NSConfig_Cookies:DSConfig_Cookies

Cookie_Domain = request.Cookies("FoosunSearchCookie")("Cookie_Domain")
Cookie_Copyright = request.Cookies("FoosunSearchCookie")("Cookie_Copyright")
Cookie_eMail = request.Cookies("FoosunSearchCookie")("Cookie_eMail")
Cookie_Site_Name = request.Cookies("FoosunSearchCookie")("Cookie_Site_Name")

if Cookie_Domain="" then 
	Cookie_Domain = "http://localhost"
else
	if left(lcase(Cookie_Domain),len("http://"))<>"http://" then Cookie_Domain = "http://"&Cookie_Domain
	if right(Cookie_Domain,1)="/" then Cookie_Domain = mid(Cookie_Domain,1,len(Cookie_Domain) - 1)
end if	

''防盗连
Server_Name = Len(Request.ServerVariables("SERVER_NAME"))
Server_V1 = Left(Replace(Cstr(Request.ServerVariables("HTTP_REFERER")),"http://",""),Server_Name)
Server_V2 = Left(Cstr(Request.ServerVariables("SERVER_NAME")),Server_Name)
if Server_V1 <> Server_V2 and Server_V1 <> "" and Server_V2 <> "" then
	response.Write("没有权限，请访问<a href="""&Cookie_Domain&""">"&Cookie_Domain&"</a>.")
	response.End()
end if

''+++++++++++++++++++++++++++++++++++++++++++
select case GetType
case "LoginHtml"
%>
<%
	If session("FS_UserName") = "" Then
		Response.Write("北京思赛文站内搜索")
	Else
		Response.Write("<a href="&Cookie_Domain&"/User/Main.asp>["&session("FS_UserName")&"]欢迎您来到会员中心")
	End if
%>
<%case "FootHTML"%>

&nbsp;<BR>
        <BR>
        <FONT 
      size=-1><%=Cookie_Copyright%></FONT><BR><BR>
	  
<%case "CopyrightHTML"
	TmpStr = "<TABLE cellSpacing=0 cellPadding=2 width=""940"" border=0>"&vbNewLine _ 
	&"<TR>"&vbNewLine _ 
	&"<TD align=right height=25></TD>"&vbNewLine _
	&"</TR>"&vbNewLine _
	&"</TABLE>"&vbNewLine _ 
	&"</CENTER>"&vbNewLine
	response.Write(TmpStr)  
case "MainInfo"
SubSys = Ucase(NoHtmlHackInput(NoSqlHack(request.QueryString("SubSys"))))
Keyword = NoSqlHack(request.QueryString("Keyword"))
s_type = NoHtmlHackInput(NoSqlHack(request.QueryString("s_type")))
ClassId = NoHtmlHackInput(NoSqlHack(request.QueryString("ClassId")))
AreaID = NoHtmlHackInput(NoSqlHack(request.QueryString("s_area")))
PubType = NoHtmlHackInput(NoSqlHack(request.QueryString("PubType")))
s_date = NoSqlHack(trim(request.QueryString("s_date")))
e_date = NoSqlHack(trim(request.QueryString("e_date")))
If e_date <> "" And Len(e_date) <= 10 Then
	e_date = e_date & " " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now())
End If	
If SubSys <> "SD" And SubSys <> "HS" Then 
	If Keyword = "" then 
		strShowErr=strShowErr&"<li>关键字不能为空</li>"&vbnewLine
	End If
End If		
if s_date<>"" then if not isdate(s_date) then strShowErr=strShowErr&"<li>开始日期"&s_date&"非法</li>"&vbnewLine
if e_date<>"" then if not isdate(e_date) then strShowErr=strShowErr&"<li>结束日期"&e_date&"非法</li>"&vbnewLine
if strShowErr<>"" then strShowErr=strShowErr&"<li><a href="""&Cookie_Domain&""">"&Cookie_Domain&"</a>.</li>": response.Write(strShowErr):response.End()
if SubSys="" then SubSys = "NS"
'===========================================
''sql的处理
select case SubSys
	case "NS"
	TmpStr = "新闻"
	select case s_type
		case "title"
			s_type="NewsTitle"   
		case "stitle"
			s_type = "CurtTitle"
		case "content"
			s_type = "Content"
		case "NaviContent"
			s_type ="NewsNaviContent"
		case "author","source"
			s_type = s_type
		case "keyword"
			s_type = "Keywords"
		case else
			s_type = "NewsTitle,CurtTitle,NewsNaviContent,Content,author,source,Keywords"
	end select	
	Search_Sql = "select NewsID,NewsTitle,CurtTitle,NewsNaviContent,Content,A.addtime,PopId,ClassName,A.IsURL,isPicNews,NewsSmallPicFile,NewsPicFile," _
		&"Source,Author,Hits,TodayNewsPic,ClassEName,SaveNewsPath,FileName,A.FileExtName,B.[Domain],B.SavePath,A.URLAddress from FS_NS_News A,FS_NS_NewsClass B where A.ClassID=B.ClassID" _
		&" and isLock=0 and isRecyle=0 and isdraft=0"
		if Keyword<>"" then 
			if instr(s_type,",")=0 then 
				Search_Sql = and_where(Search_Sql) & Search_TextArr(Keyword,s_type,"") 
			else
				TmpArr = split(s_type,",")
				TmpStr2 = ""
				for each TmpStr1 in  TmpArr
					TmpStr2 = TmpStr2 & " or " & Search_TextArr(Keyword,TmpStr1,"")&"" 			
				next
				if left(TmpStr2,len(" or "))=" or " then 
					TmpStr2 = mid(TmpStr2,len(" or ")+1) : TmpStr2 = " ("&TmpStr2&") " :Search_Sql = and_where(Search_Sql) & TmpStr2				
				end if	
			end if
		end if	
		if ClassID<>"" then Search_Sql = and_where(Search_Sql) & Search_TextArr(ClassId,"A.ClassID","") 
		if s_date<>"" then Search_Sql = and_where(Search_Sql) & " A.addtime>="&SqlDateType&s_date&SqlDateType
		if e_date<>"" then Search_Sql = and_where(Search_Sql) & " A.addtime<="&SqlDateType&e_date&SqlDateType
	case "WS"
	TmpStr = "留言本"
	select case s_type
		case "title"
		s_type="Topic"   
		case "content"
		s_type = "Body"
		case "author"
		s_type = "User"
		case else
		s_type = "Topic,Body,User"
	end select	
	Search_Sql = "select A.ID,ClassName,ParentID,User,Topic,Body,A.AddDate,IsTop,State,Style,IsAdmin,Answer,Hit,LastUpdateDate,LastUpdateUser," _
		&"Face,IP,Isonline,Vistor,isUpoad " _
		&" from FS_WS_BBS A,FS_WS_Class B where A.ClassID=B.ClassID "
		if Keyword<>"" then 
			if instr(s_type,",")=0 then 
				Search_Sql = and_where(Search_Sql) & Search_TextArr(Keyword,s_type,"") 
			else
				TmpArr = split(s_type,",")
				TmpStr2 = ""
				for each TmpStr1 in  TmpArr
					TmpStr2 = TmpStr2 & " or " & Search_TextArr(Keyword,TmpStr1,"")&"" 			
				next
				if left(TmpStr2,len(" or "))=" or " then 
					TmpStr2 = mid(TmpStr2,len(" or ")+1) : TmpStr2 = " ("&TmpStr2&") " :Search_Sql = and_where(Search_Sql) & TmpStr2				
				end if	
			end if
		end if	
		if ClassID<>"" then Search_Sql = and_where(Search_Sql) & Search_TextArr(ClassId,"A.ClassID","") 
		if s_date<>"" then Search_Sql = and_where(Search_Sql) & " A.AddDate>="&SqlDateType&s_date&SqlDateType
		if e_date<>"" then Search_Sql = and_where(Search_Sql) & " A.AddDate<="&SqlDateType&e_date&SqlDateType
	case "DS"
	TmpStr = "下载"
	select case s_type
		case "title"
		s_type="Name"
		case "content"
		s_type = "Description"
		case else
		s_type = "Name,Description"
	end select	
	Dim DataStr
	IF G_IS_SQL_DB = 1 Then
		DataStr =  "datediff(d,A.AddTime,"&date()&")"
	Else
		DataStr =  "datediff('d',A.AddTime,'"&date()&"')"
	End If	
	Search_Sql = "select DownLoadID,ClassName,Name,Description,A.AddTime,ClickNum,ClassEName,A.SavePath,A.FileExtName,A.FileName," _
		&"FileSize,RecTF,Types,Hits,ConsumeNum,A.Pic from FS_DS_List A,FS_DS_Class B where A.ClassID=B.ClassID" _
		&" and AuditTF=1 and (OverDue=0 or (OverDue>0 and " & DataStr & " <= OverDue)) "
		if Keyword<>"" then 
			if instr(s_type,",")=0 then 
				Search_Sql = and_where(Search_Sql) & Search_TextArr(Keyword,s_type,"") 
			else
				TmpArr = split(s_type,",")
				TmpStr2 = ""
				for each TmpStr1 in  TmpArr
					TmpStr2 = TmpStr2 & " or " & Search_TextArr(Keyword,TmpStr1,"")&"" 			
				next
				if left(TmpStr2,len(" or "))=" or " then 
					TmpStr2 = mid(TmpStr2,len(" or ")+1) : TmpStr2 = " ("&TmpStr2&") " :Search_Sql = and_where(Search_Sql) & TmpStr2				
				end if	
			end if
		end if	
		if ClassID<>"" then Search_Sql = and_where(Search_Sql) & Search_TextArr(ClassId,"A.ClassID","") 
		if s_date<>"" then Search_Sql = and_where(Search_Sql) & " A.addtime>="&SqlDateType&s_date&SqlDateType
		if e_date<>"" then Search_Sql = and_where(Search_Sql) & " A.addtime<="&SqlDateType&e_date&SqlDateType
	case "MS"
	TmpStr = "商城"
	select case s_type
		case "title"
		s_type="ProductTitle"
		case "content"
		s_type = "ProductContent"
		case "keyword"
		s_type = "Keywords"
		case else
		s_type = "ProductTitle,ProductContent,Keywords"
	end select	
	'-----
	MsMinPric = Trim(NoHtmlHackInput(NoSqlHack(request.QueryString("MinPric"))))
	MsMaxPric = Trim(NoHtmlHackInput(NoSqlHack(request.QueryString("MaxPric"))))
	'------
	'---2007.9.25 Edit by arjun
	'双表查询，关键字段ID,如果不指定表，将造成查询时出错
	Search_Sql = "select A.ID as ID,ProductTitle,ClassCName,Stockpile,StockpileWarn,OldPrice,NewPrice,IsWholesale,ProductContent,MakeFactory," _
		&"ProductsAddress,Click,smallPic,BigPic,StyleFlagBit,SaleStyle,A.AddTime,Discount,DiscountStartDate,DiscountEndDate," _
		&" ClassEName,A.SavePath,A.FileExtName,A.FileName from FS_MS_Products A,FS_MS_ProductsClass B where A.ClassID=B.ClassID "
		if Keyword<>"" then 
			if instr(s_type,",")=0 then 
				Search_Sql = and_where(Search_Sql) & Search_TextArr(Keyword,s_type,"") 
			else
				TmpArr = split(s_type,",")
				TmpStr2 = ""
				for each TmpStr1 in  TmpArr
					TmpStr2 = TmpStr2 & " or " & Search_TextArr(Keyword,TmpStr1,"")&"" 			
				next
				if left(TmpStr2,len(" or "))=" or " then 
					TmpStr2 = mid(TmpStr2,len(" or ")+1) : TmpStr2 = " ("&TmpStr2&") " :Search_Sql = and_where(Search_Sql) & TmpStr2				
				end if	
			end if
		end if	
		if ClassID<>"" then Search_Sql = and_where(Search_Sql) & Search_TextArr(ClassId,"A.ClassID","") 
		if s_date<>"" then Search_Sql = and_where(Search_Sql) & " A.addtime>="&SqlDateType&s_date&SqlDateType
		if e_date<>"" then Search_Sql = and_where(Search_Sql) & " A.addtime<="&SqlDateType&e_date&SqlDateType
		If MsMinPric <> "" And IsNumeric(MsMinPric) Then Search_Sql = and_where(Search_Sql) & " A.NewPrice>="&MsMinPric
		If MsMaxPric <> "" And IsNumeric(MsMaxPric) Then Search_Sql = and_where(Search_Sql) & " A.NewPrice<="&MsMaxPric
		'Search_Sql=""
	case "RD"
	If G_IS_SQL_Old_News_DB = 1 Then  
		RDSqlDateType = "'"
	else
		RDSqlDateType = "#"
	end if
	TmpStr = "总站归档搜索"
	select case s_type
		case "title"
		s_type="NewsTitle"
		case "stitle"
		s_type = "CurtTitle"
		case "content"
		s_type = "Content"
		case "NaviContent"
		s_type ="NewsNaviContent"
		case "author","source"
		s_type = s_type
		case "keyword"
		s_type = "Keywords"
		case else
		s_type = "NewsTitle,CurtTitle,NewsNaviContent,Content,author,source,Keywords"
	end select	
	Search_Sql = "select ID,NewsID,NewsTitle,CurtTitle,NewsNaviContent,Content,addtime,PopId,IsURL,isPicNews,NewsSmallPicFile,NewsPicFile," _
		&"Source,Author,Hits,TodayNewsPic,SaveNewsPath,FileName,FileExtName,FileTime from FS_Old_News " _
		&" where isLock=0 and isRecyle=0 and isdraft=0 "
		if Keyword<>"" then 
			if instr(s_type,",")=0 then 
				Search_Sql = and_where(Search_Sql) & Search_TextArr(Keyword,s_type,"") 
			else
				TmpArr = split(s_type,",")
				TmpStr2 = ""
				for each TmpStr1 in  TmpArr
					TmpStr2 = TmpStr2 & " or " & Search_TextArr(Keyword,TmpStr1,"")&"" 			
				next
				if left(TmpStr2,len(" or "))=" or " then 
					TmpStr2 = mid(TmpStr2,len(" or ")+1) : TmpStr2 = " ("&TmpStr2&") " :Search_Sql = and_where(Search_Sql) & TmpStr2				
				end if	
			end if
		end If
		if s_date<>"" then Search_Sql = and_where(Search_Sql) & " addtime>="&RDSqlDateType&s_date&RDSqlDateType
		if e_date<>"" then Search_Sql = and_where(Search_Sql) & " addtime<="&RDSqlDateType&e_date&RDSqlDateType
	case "SD"
	Dim ValidTime
	TmpStr = "供求"
	if G_IS_SQL_DB=1 then 
		ValidTime = NoSqlHack(" and dateadd(d,ValidTime,EditTime)>=getdate() ")
	else
		ValidTime = NoSqlHack(" and dateadd('d',ValidTime,EditTime)>=date() ")
	end If
	Search_Sql = "select A.ID,A.UserNumber,PubTitle,PubType,PubContent,Keyword,CompType,PubNumber,PubPrice,PubPack,Pubgui,PubPic_1,PubPic_2,PubPic_3," _
		&"A.Addtime,EditTime,ValidTime,PubAddress,otherLink,hits,GQ_ClassName,C.ClassName,AreaID,A.ClassID from FS_SD_News A,FS_SD_Class B,FS_SD_Address C  " _
		&"where A.ClassID=B.ID and A.AreaID = C.ID and A.isLock=0 and A.HideTF=0 and isPass=1 "&ValidTime
		'---ken
		If Keyword <> "" Then Search_Sql = and_where(Search_Sql) & " A.PubTitle Like '%"&Keyword&"%' " 
		'---end
		if ClassID<>"" then Search_Sql = and_where(Search_Sql) & " A.ClassID="&ClassId&" " 
		if AreaID<>"" then Search_Sql = and_where(Search_Sql) & " A.AreaID="&AreaID&" " 
		if PubType<>"" then if cint(PubType)>0 then Search_Sql = and_where(Search_Sql) & " A.PubType="&cint(PubType)-1&" " 
		Search_Sql = Search_Sql & " order by C.ClassLevel desc,B.ClassOrder desc,A.hits desc,A.ID desc"	
'----------2007-01-18 Edit By Ken For Fs_House Search
	Case "HS"
		Dim HSType,QH_Type,THKWD_Type,THUse_Type,THInfo_Type,THHouse_Type,THZX_Type
		Dim SHKey_Type,SHUse_Type,SHHouse_Type,SHZX_Type
		Dim KWDsType,HSsql
		HSType =  NoHtmlHackInput(NoSqlHack(request.QueryString("HSType")))
		'If HSType = "" Then : Response.Write "参数传递错误" : Response.End : End If
		'----  匹配信息发布时间范围
		If s_date <> "" And e_date <> "" Then
			If G_IS_SQL_DB = 1 Then
				HSsql = HSsql & " And PubDate >= '" & s_date & "' And PubDate <= '" & e_date & "'"
			Else
				HSsql = HSsql & " And PubDate >= #" & s_date & "# And PubDate <= #" & e_date & "#"
			End If	
		Else
			HSsql = HSsql 		
		End If
		If HSType = "Quotation" Then
			TmpStr = "楼盘信息搜索"
			KWDsType = NoHtmlHackInput(NoSqlHack(request.QueryString("QHKey_Type")))
			QH_Type = NoHtmlHackInput(NoSqlHack(request.QueryString("QH_Type")))
			Select Case KWDsType   '匹配搜索关键字
				Case "QHTitle"
					If Keyword <> "" Then
						HSsql = HSsql & " And HouseName Like '%" & Keyword & "%'"
					Else
						HSsql = HSsql
					End If
				Case "QHAddress"
					If Keyword <> "" Then
						HSsql = HSsql & " And Position Like '%" & Keyword & "%'"
					Else
						HSsql = HSsql
					End If
				Case "QHSell"
					If Keyword <> "" Then
						HSsql = HSsql & " And PreSaleRange Like '%" & Keyword & "%'"
					Else
						HSsql = HSsql
					End If
				Case "QHSellNum"
					If Keyword <> "" Then
						HSsql = HSsql & " And PreSaleNumber Like '%" & Keyword & "%'"
					Else
						HSsql = HSsql
					End If
				Case "QHInfo"
					If Keyword <> "" Then
						HSsql = HSsql & " And introduction Like '%" & Keyword & "%'"
					Else
						HSsql = HSsql
					End If
				Case "QHTelNum"
					If Keyword <> "" Then
						HSsql = HSsql & " And Tel Like '%" & Keyword & "%'"
					Else
						HSsql = HSsql
					End If
				Case Else
					HSsql = HSsql					 
			End Select
			'----  匹配房屋状况
			If QH_Type <> "" Then
				QH_Type = Cint(QH_Type)
				HSsql = HSsql & " And Status = " & QH_Type
			Else
				HSsql = HSsql
			End If
			'---匹配查询sql语句
			Search_Sql = "Select ID,HouseName,Position,Direction,Class,OpenDate,PreSaleNumber,IssueDate,PreSaleRange,Status,Price,PubDate,Tel,Click,UserNumber,introduction From FS_HS_Quotation Where isRecyle = 0 And Audited = 1" & NoSqlHack(HSsql) & " Order By ID Desc,PubDate Desc"
		ElseIf HSType = "Tenancy" Then
			TmpStr = "租赁信息搜索"
			THKWD_Type = NoHtmlHackInput(NoSqlHack(request.QueryString("THKey_Type")))
			THUse_Type = NoHtmlHackInput(NoSqlHack(request.QueryString("THUse_Type")))
			THInfo_Type = NoHtmlHackInput(NoSqlHack(request.QueryString("THInfo_Type")))
			THHouse_Type = NoHtmlHackInput(NoSqlHack(request.QueryString("THHouse_Type")))
			THZX_Type = NoHtmlHackInput(NoSqlHack(request.QueryString("THZX_Type")))
			'-----匹配查询关键字
			Select Case THKWD_Type
				Case "THTitle"
					If Keyword <> "" Then
						HSsql = HSsql & " And Position Like '%" & Keyword & "%'"
					Else
						HSsql = HSsql
					End If		
				Case "THAdd"
					If Keyword <> "" Then
						HSsql = HSsql & " And CityArea Like '%" & Keyword & "%'"
					Else
						HSsql = HSsql
					End If
				Case "THAouth"
					If Keyword <> "" Then
						HSsql = HSsql & " And LinkMan Like '%" & Keyword & "%'"
					Else
						HSsql = HSsql
					End If
				Case "THInfo"
					If Keyword <> "" Then
						HSsql = HSsql & " And Remark Like '%" & Keyword & "%'"
					Else
						HSsql = HSsql
					End If
				Case Else
					HSsql = HSsql
			End Select
			'---匹配房屋用途
			If THUse_Type <> "" Then
				THUse_Type = Cint(THUse_Type)
				HSsql = HSsql & " And UseFor = " & THUse_Type
			Else
				HSsql = HSsql
			End If
			'---匹配信息类型
			If THInfo_Type <> "" Then
				THInfo_Type = Cint(THInfo_Type)
				HSsql = HSsql & " And Class = " & THInfo_Type
			Else
				HSsql = HSsql
			End If
			'---匹配户型
			If THHouse_Type <> "" Then
				If THHouse_Type = "Other" Then
					HSsql = HSsql & " And " & StringType & "(HouseStyle,1,1) > 3 And " & StringType & "(HouseStyle,3,1) > 2"
				Else
					HSsql = HSsql & " And " & StringType & "(HouseStyle,1,1) = '" & Split(THHouse_Type,",")(0) & "' And " & StringType & "(HouseStyle,3,1) = '" & Split(THHouse_Type,",")(1) & "'"
				End If
			Else
				HSsql = HSsql
			End If
			'---匹配装修类型
			If THZX_Type <> "" Then
				THZX_Type = Cint(THZX_Type)
				HSsql = HSsql & " And Decoration = " & THZX_Type
			Else
				HSsql = HSsql
			End If
			'---匹配查询sql语句
			Search_Sql = "Select TID,UseFor,Class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate From FS_HS_Tenancy Where Audited = 1 And isRecyle = 0" & NoSqlHack(HSsql) & " Order By TID Desc,PubDate Desc"
		Else
			TmpStr = "二手房信息搜索"
			SHKey_Type = NoHtmlHackInput(NoSqlHack(request.QueryString("SHKey_Type")))
			SHUse_Type = NoHtmlHackInput(NoSqlHack(request.QueryString("SHUse_Type")))
			SHHouse_Type = NoHtmlHackInput(NoSqlHack(request.QueryString("SHHouse_Type")))
			SHZX_Type = NoHtmlHackInput(NoSqlHack(request.QueryString("SHZX_Type")))
			'---匹配搜索关键字
			Select Case SHKey_Type
				Case "SHNum"
					If Keyword <> "" Then
						HSsql = HSsql & " And Label = '%" & Keyword & "%'"
					Else
						HSsql = HSsql
					End If
				Case "SHAdd"
					If Keyword <> "" Then
						HSsql = HSsql & " And Address Like '%" & Keyword & "%'"
					Else
						HSsql = HSsql
					End If
				Case "SHCent"
					If Keyword <> "" Then
						HSsql = HSsql & " And CityArea Like '%" & Keyword & "%'"
					Else
						HSsql = HSsql
					End If
				Case "SHAouth"
					If Keyword <> "" Then
						HSsql = HSsql & " And LinkMan Like '%" & Keyword & "%'"
					Else
						HSsql = HSsql
					End If
				Case "SHInfo"
					If Keyword <> "" Then
						HSsql = HSsql & " And Remark Like '%" & Keyword & "%'"
					Else
						HSsql = HSsql
					End If
				Case Else
					HSsql = HSsql
			End Select
			'---匹配房屋用途
			If SHUse_Type <> "" Then
				SHUse_Type = Cint(SHUse_Type)
				HSsql = HSsql & " And UseFor = " & SHUse_Type
			Else
				HSsql = HSsql
			End If
			'---匹配户型
			If SHHouse_Type <> "" Then
				If SHHouse_Type = "Other" Then
					HSsql = HSsql & " And " & StringType & "(HouseStyle,1,1) > 3 And " & StringType & "(HouseStyle,3,1) > 2"
				Else
					HSsql = HSsql & " And " & StringType & "(HouseStyle,1,1) = '" & Split(SHHouse_Type,",")(0) & "' And " & StringType & "(HouseStyle,3,1) = '" & Split(SHHouse_Type,",")(1) & "'"
				End If
			Else
				HSsql = HSsql
			End If
			'---匹配装修类型
			If SHZX_Type <> "" Then
				SHZX_Type = Cint(SHZX_Type)
				HSsql = HSsql & " And Decoration = " & SHZX_Type
			Else
				HSsql = HSsql
			End If
			'HSsql=""
			'---匹配搜索sql语句
			Search_Sql = "Select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,Contact,Remark,PubDate From FS_HS_Second Where Audited = 1 And isRecyle = 0" & NoSqlHack(HSsql) & " Order By SID Desc,PubDate Desc"
		End If
'------------------
case else
	strShowErr="<li>错误的参数传递.SubSys</li><li><a href="""&Cookie_Domain&""">"&Cookie_Domain&"</a>.</li>": response.Write(strShowErr):response.End()			
end select
On Error Resume Next
Set Search_RS = CreateObject(G_FS_RS)
if SubSys="RD" then 
	Search_RS.Open Search_Sql,Old_News_Conn,1,1
Else
	Search_RS.Open Search_Sql,Conn,1,1
end if	
if Err<>0 then 
	response.Write("<li>查询条件不匹配.无法继续.<br />"&"<br/>"&Err.Number&":"&Err.description&"</li>")
	response.End()
end if
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=10 '设置每页显示数目
int_showNumberLink_=10 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings>9</font>"   			'首页 
toP10_=" <font face=webdings>7</font>"			'上十 
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"				'尾页

IF Search_RS.eof THEN%>
<TABLE class="t bt" cellSpacing=0 cellPadding=0 width="940" border=0 align="center">
  <TBODY>
  <TR>
    <TD noWrap><%=morestr(Keyword,30)%> <FONT size=+1>&nbsp;<B><FONT size=+1>&nbsp;<B>搜索结果</B></FONT>&nbsp;</B></FONT>&nbsp;</TD>
    <TD noWrap align=right>
	<FONT size=-1>共有<B>0</B>项符合<B><%=Keyword%></B>的查询结果
	（搜索用时 <B><%=FormatNumber((timer()-starttime),2,-1)%></B>毫秒）&nbsp;</FONT>
	</TD>
   </TR>
  </TBODY>
</TABLE>
<p><font size=-1 color=#666666>未查询到符合条件的记录。</font></p>
<%else
Dim UrlAndTitle,SaveNewsPath,Content,NewsSmallPicFile,NewsPicFile,addtime,NaviContent ,SysRs_Tmp,ChildPath,picShuXing,picShuXingB
Dim rndpic,DBUserName,sdnextdoman
Search_RS.PageSize=int_RPP
cPageNo=CintStr(Request.QueryString("Page"))
If cPageNo="" or not isnumeric(cPageNo) Then cPageNo = 1
cPageNo = Clng(cPageNo)
If cPageNo<1 Then cPageNo=1
If cPageNo>Search_RS.PageCount Then cPageNo=Search_RS.PageCount 
Search_RS.AbsolutePage=cPageNo
  FOR int_Start=1 TO int_RPP 
 
select case  SubSys
case "NS"
	Dim f_NewsLinkRecordSet
	Set f_NewsLinkRecordSet = New CLS_FoosunRecordSet
	Set f_NewsLinkRecordSet.Values("ClassEName,Domain,SavePath,IsURL,URLAddress,SaveNewsPath,FileName,FileExtName") = Search_RS
	SaveNewsPath = get_NewsLink(f_NewsLinkRecordSet)
	Set f_NewsLinkRecordSet = Nothing
	UrlAndTitle = "<A class=l href="""&SaveNewsPath&""" target=_blank>"&Search_RS("NewsTitle")&"</A>"
	addtime = Search_RS("addtime")
	if isnull(addtime) then addtime=""
	if isdate(addtime) then addtime = FormatDateTime(addtime,0)
	NewsSmallPicFile = Search_RS("NewsSmallPicFile")
	NewsPicFile = Search_RS("NewsPicFile")
	if NewsSmallPicFile = "" then 
		NewsSmallPicFile = "sys_images/NoPic.jpg"
	end if		
		
	NaviContent = Search_RS("Content")
	if isnull(NaviContent) or NaviContent="" then 
		NaviContent = "暂无"
	else
		NaviContent = morestr(Lose_Html(NaviContent),255)
		NaviContent = morestr(Lose_Html(NaviContent),255)
		If NaviContent <> "" Then
			NaviContent = Search_TextArr(Keyword,"Content",NaviContent)
		Else
			NaviContent = "详细内容请点击标题浏览"
		End If
	end if
	FileSize = split(CheckFile(LocalUrl),"|")(0)
	FileEditDate = split(CheckFile(LocalUrl),"|")(1)
case "RD"	
	SaveNewsPath = ""&Cookie_Domain&"/historynews.asp?id="&Search_RS("ID")&""
	UrlAndTitle = "<A class=l href="""&SaveNewsPath&""" target=_blank>"&Search_RS("NewsTitle")&"</A>"
	addtime = Search_RS("addtime")
	if isnull(addtime) then addtime=""
	if isdate(addtime) then addtime = FormatDateTime(addtime,0)
	NewsSmallPicFile = Search_RS("NewsSmallPicFile")
	NewsPicFile = Search_RS("NewsPicFile")
	if NewsSmallPicFile = "" then 
		NewsSmallPicFile = "sys_images/NoPic.jpg"
	end if		
			
	NaviContent = Search_RS("Content")
	if isnull(NaviContent) or NaviContent="" then 
		NaviContent = "暂无"
	else
		NaviContent = morestr(Lose_Html(NaviContent),255)
		If NaviContent <> "" Then
			NaviContent = Search_TextArr(Keyword,"Content",NaviContent)
		Else
			NaviContent = "详细内容请点击标题浏览"
		End If		
	end if
	FileSize = ""
	FileEditDate = addtime
case "WS"	
	SaveNewsPath = ""&Cookie_Domain&"/historynews.asp?id="&Search_RS("ID")&""
	UrlAndTitle = "<A class=l href="""&SaveNewsPath&""" target=_blank>"&Search_RS("Topic")&"</A>"
	addtime = Search_RS("AddDate")
	if isnull(addtime) then addtime=""
	if isdate(addtime) then addtime = FormatDateTime(addtime,0)
	NewsSmallPicFile = Search_RS("Face")
	NewsPicFile = Search_RS("Face")
	if NewsSmallPicFile = "" then 
		NewsSmallPicFile = "sys_images/NoPic.jpg"
	else
		NewsSmallPicFile = Cookie_Domain&"/sys_images/emot/face"&NewsSmallPicFile&".gif"
	end if		
			
	NaviContent = Search_RS("Body")
	if isnull(NaviContent) or NaviContent="" then 
		NaviContent = "暂无"
	else
		NaviContent = morestr(Lose_Html(NaviContent),255)
		NaviContent = Search_TextArr(Keyword,"Body",NaviContent)
	end if
	FileSize = ""
	FileEditDate = addtime
case "DS"
	set SysRs_Tmp = Conn.execute("select DownDir,IsDomain From FS_DS_SysPara")
	if not SysRs_Tmp.eof then 
		if not isnull(SysRs_Tmp("IsDomain")) and SysRs_Tmp("IsDomain") <> "" then 
			ChildDomain  =  SysRs_Tmp("IsDomain")
		else
			ChildDomain =Cookie_Domain&"/"&SysRs_Tmp("DownDir")
		end if
		LocalUrl = "/"&SysRs_Tmp("DownDir")
	end If
	SysRs_Tmp.close
	if isnull(ChildDomain) then ChildDomain = ""
	ClassPath = Search_RS("ClassEName")
	if isnull(ClassPath) then ClassPath = ""
	if right(ChildDomain,1)="/" then ChildDomain = mid(ChildDomain,1,len(ChildDomain) - 1)
	if ChildDomain<>"" then 
		if left(lcase(ChildDomain),len("http://"))<>"http://" then ChildDomain = "http://"&ChildDomain
	else
		ChildDomain = Cookie_Domain
	end if		
	if ClassPath<>"" then ClassPath = replace(ClassPath,"/","")   
	if len(ChildDomain&"/"&ClassPath)>1 then ChildPath = ChildDomain&"/"&ClassPath
	SaveNewsPath = ChildPath & Search_RS("SavePath")
	if right(SaveNewsPath,1)<>"/" then SaveNewsPath = SaveNewsPath&"/"
	SaveNewsPath = SaveNewsPath & Search_RS("FileName") &"."& Search_RS("FileExtName")
	
	SaveNewsPath = GetTheInfoLink(Search_RS("DownLoadID"),"DS")
	UrlAndTitle = "<A class=l href="""&SaveNewsPath&""" target=_blank>"&Search_RS("Name")&"</A>"
	addtime = Search_RS("addtime")
	if isnull(addtime) then addtime=""
	if isdate(addtime) then addtime = FormatDateTime(addtime,0)
	NewsSmallPicFile = Search_RS("Pic")
	NewsPicFile = Search_RS("Pic")
	if NewsSmallPicFile = "" then 
		NewsSmallPicFile = "sys_images/NoPic.jpg"
	end if		
		
	NaviContent = Search_RS("Description")
	if isnull(NaviContent) or NaviContent="" then 
		NaviContent = "暂无"
	else
		NaviContent = morestr(Lose_Html(NaviContent),255)
		NaviContent = Search_TextArr(Keyword,"Description",NaviContent)
	end if
	LocalUrl = LocalUrl&"/"&ClassPath&Search_RS("SavePath")&"/"&Search_RS("FileName") &"."& Search_RS("FileExtName")
	FileSize = split(CheckFile(LocalUrl),"|")(0)
	FileEditDate = split(CheckFile(LocalUrl),"|")(1)
case "MS"
	set SysRs_Tmp = Conn.execute("select SavePath,IsDomain From FS_MS_SysPara")
	if not SysRs_Tmp.eof then 
		if not isnull(SysRs_Tmp("IsDomain")) and SysRs_Tmp("IsDomain") <> "" then 
			ChildDomain  =  SysRs_Tmp("IsDomain")
		else
			ChildDomain = Cookie_Domain & "/" & SysRs_Tmp("SavePath")
		end if
		LocalUrl = "/"&SysRs_Tmp("SavePath")
	end if
	SysRs_Tmp.close
	if isnull(ChildDomain) then ChildDomain = ""
	ClassPath = Search_RS("ClassEName")
	if isnull(ClassPath) then ClassPath = ""
	if right(ChildDomain,1)="/" then ChildDomain = mid(ChildDomain,1,len(ChildDomain) - 1)
	if ChildDomain<>"" then 
		if left(lcase(ChildDomain),len("http://"))<>"http://" then ChildDomain = "http://"&ChildDomain
	else
		ChildDomain =Cookie_Domain&"/"&SysRs_Tmp("NewsDir")     
	end if		
	if ClassPath<>"" then ClassPath = replace(ClassPath,"/","")
	if len(ChildDomain&"/"&ClassPath)>1 then ChildPath = ChildDomain&"/"&ClassPath
	SaveNewsPath = ChildPath & Search_RS("SavePath")
	if right(SaveNewsPath,1)<>"/" then SaveNewsPath = SaveNewsPath&"/"

	SaveNewsPath = SaveNewsPath & Search_RS("FileName") &"."& Search_RS("FileExtName")
	
	SaveNewsPath = GetTheInfoLink(Search_RS("ID"),"MS")
	UrlAndTitle = "<A class=l href="""&SaveNewsPath&""" target=_blank>"&Search_RS("ProductTitle")&"</A>"
	addtime = Search_RS("addtime")
	if isnull(addtime) then addtime=""
	if isdate(addtime) then addtime = FormatDateTime(addtime,0)
	NewsSmallPicFile = Search_RS("smallPic")
	NewsPicFile = Search_RS("BigPic")
	if NewsSmallPicFile = "" then 
		NewsSmallPicFile = "sys_images/NoPic.jpg"
	end if		
			
	NaviContent = Search_RS("ProductContent")
	if isnull(NaviContent) or NaviContent="" then 
		NaviContent = "暂无"
	else
		NaviContent = morestr(Lose_Html(NaviContent),255)
		NaviContent = Search_TextArr(Keyword,"ProductContent",NaviContent)
	end if
	LocalUrl = LocalUrl&"/"&ClassPath&Search_RS("SavePath")&"/"&Search_RS("FileName") &"."& Search_RS("FileExtName")
	FileSize = split(CheckFile(LocalUrl),"|")(0)
	FileEditDate = split(CheckFile(LocalUrl),"|")(1)
case "SD"	
	sdnextdoman =  Get_OtherTable_Value("select top 1 [Domain] from FS_SD_Config")
	if sdnextdoman<>"" then 
		SaveNewsPath = "http://"&sdnextdoman&"/Supply.asp?id="&Search_RS("ID")
	else
		SaveNewsPath = ""&Cookie_Domain&"/Supply/Supply.asp?id="&Search_RS("ID")
	end if	
	UrlAndTitle = "["&Replacestr(Search_RS("PubType"),"0:供应,1:求购,2:合作,3:代理,4:其它")&"] " _
		&"<A class=l href="""&SaveNewsPath&""" target=_blank>"&Search_RS("PubTitle")&"</A>"
	addtime = Search_RS("Addtime")
	if isnull(addtime) then addtime=""
	if isdate(addtime) then addtime = FormatDateTime(addtime,0)
	randomize
	rndpic = Int(3 * Rnd + 1)
	NewsSmallPicFile = Search_RS("PubPic_"&rndpic)
	NewsPicFile = NewsSmallPicFile
	if NewsSmallPicFile = "" then NewsSmallPicFile = "sys_images/NoPic.jpg"
			
	NaviContent = Search_RS("PubContent")
	if isnull(NaviContent) or NaviContent="" then 
		NaviContent = "暂无"
	else
		NaviContent = morestr(Lose_Html(NaviContent),255)
		NaviContent = Search_TextArr(Keyword,"PubContent",NaviContent)
	end if
	FileSize = ""
	FileEditDate = Search_RS("EditTime")
Case "HS"
	Dim HsPicSql,HsPicRs
	sdnextdoman =  Get_OtherTable_Value("select top 1 isDomain from FS_HS_SysPara")
	If LCase(HSType) = "quotation" Then
		if sdnextdoman<>"" then 
			SaveNewsPath = "http://"&sdnextdoman&"/BuildingRead.asp?id="&Search_RS("ID")
		else
			SaveNewsPath = ""&Cookie_Domain&"/House/BuildingRead.asp?id="&Search_RS("ID")
		end if
		UrlAndTitle = "<A class=l href="""&SaveNewsPath&""" target=_blank>"&Search_RS("HouseName")&"</A>"
		HsPicSql = "Select * From FS_HS_Picture Where HS_Type = 1 And ID = " & Search_RS("ID")
	ElseIf LCase(HSType) = "second" Then
		if sdnextdoman<>"" then 
			SaveNewsPath = "http://"&sdnextdoman&"/SecondRead.asp?id="&Search_RS("SID")
		else
			SaveNewsPath = ""&Cookie_Domain&"/House/SecondRead.asp?id="&Search_RS("SID")
		end if
		UrlAndTitle = "<A class=l href="""&SaveNewsPath&""" target=_blank>"&Search_RS("Address")&"</A>"
		HsPicSql = "Select * From FS_HS_Picture Where HS_Type = 3 And ID = "&Search_RS("SID")
	Else
		if sdnextdoman<>"" then 
			SaveNewsPath = "http://"&sdnextdoman&"/HouseRead.asp?id="&Search_RS("TID")
		else
			SaveNewsPath = ""&Cookie_Domain&"/House/HouseRead.asp?id="&Search_RS("TID")
		end if
		UrlAndTitle = "<A class=l href="""&SaveNewsPath&""" target=_blank>"&Search_RS("Position")&"</A>"
		HsPicSql = "Select * From FS_HS_Picture Where HS_Type = 2 And ID = "&Search_RS("TID")
	End If
	Set HsPicRs = Conn.ExeCute(HsPicSql)
	If HsPicRs.Eof Then
		NewsSmallPicFile = ""
	Else
		NewsSmallPicFile = HsPicRs("Pic")
	End If
	HsPicRs.Close : Set HsPicRs = Nothing		
	NewsPicFile = NewsSmallPicFile
	if NewsSmallPicFile = "" then NewsSmallPicFile = "sys_images/NoPic.jpg"		
	addtime = Search_RS("PubDate")
	if isnull(addtime) then addtime=""
	if isdate(addtime) then addtime = FormatDateTime(addtime,0)
	If LCase(HSType) = "quotation" Then
		NaviContent = Search_RS("introduction")
	Else
		NaviContent = Search_RS("Remark")
	End if	
	if isnull(NaviContent) or NaviContent="" then 
		NaviContent = "暂无"
	else
		NaviContent = morestr(Lose_Html(NaviContent),255)
		NaviContent = Search_TextArr(Keyword,"Remark",NaviContent)
	end if
	FileSize = ""
	FileEditDate = Search_RS("PubDate")	
end select
	Content="<TABLE cellSpacing=1 cellPadding=1 border=0 width=""80%"">"&vbNewLine _
		&"<TBODY>"&vbNewLine _
		  &"<TR>"&vbNewLine _  
			&"<TD class=pic rowspan=2 align=center>"&vbNewLine 			
			if NewsPicFile<>"" then 
				picShuXingB = CheckFile(NewsPicFile)
				if len(picShuXingB)>5 then picShuXingB = "["&picShuXingB&"]" else picShuXingB="" end if 
				Content=Content&"<a href="""&NewsPicFile&""" target=""_blank""><img border=0 src="""&NewsSmallPicFile&""" alt=""点击看大图"&picShuXingB&""" onload=""if(this.offsetWidth>120)this.width=120;""></a></TD>"&vbNewLine
			else		
				Content=Content&"<img border=0 src=""sys_images/NoPic.jpg"" onload=""if(this.offsetWidth>120)this.width=120;""></TD>"&vbNewLine
			end if
			picShuXing="":picShuXingB=""		
	Content=Content	&"<TD class=content valign=top>"&vbNewLine _
			&"<font size=-1>"&NaviContent&"</font>"&vbNewLine _
			&"</TD>"&vbNewLine _
		  &"</TR>"&vbNewLine _
		  &"<TR>"&vbNewLine _ 
			&"<TD height=21><font size=-1>"&vbNewLine _
			&"<font color=#008000>"&SaveNewsPath&" - "&FileSize&" "&FileEditDate&" </font>"
	Content=Content	&"</font></TD>"&vbNewLine _
		  &"</TR>"&vbNewLine _
		&"</TBODY>"&vbNewLine _
	   &"</TABLE>"&vbNewLine
if int_Start = 1 then%>      
<TABLE class="t bt" cellSpacing=0 cellPadding=0 width="940" border=0 align="center">
  <TBODY>
  <TR>
    <TD noWrap><%=morestr(Keyword,30)%> <FONT size=+1>&nbsp;<B><FONT size=+1>&nbsp;<B>搜索结果</B></FONT>&nbsp;</B></FONT>&nbsp;</TD>
    <TD noWrap align=right>
	<FONT size=-1>共有<B><%=Search_RS.recordcount%></B>项符合 <B><%=morestr(Keyword,30)%></B> 的查询结果，
	以下是第 <B>1</B> - <B>10</B> 项。（搜索用时 <B><%=FormatNumber((timer()-starttime),2,-1)%></B> 秒）&nbsp;</FONT>
	</TD></TR></TBODY></TABLE>
<%end if%>
<DIV>
	<div>	
  <P class=g style="text-align:center">
  <%
  ''标题
  response.Write(UrlAndTitle)
  response.Write("<font size=-2 color=#666666>")
select case  SubSys
case "NS","RD"
  response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&addtime&vbNewLine)
  response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
  if SubSys="NS" then response.Write("<img border=0 src=""sys_images/alert.gif"">"&Replacestr(Search_Rs("ClassName"),":,else:"&Search_Rs("ClassName")))
  response.Write(Replacestr(Search_Rs("PopId"),"5:总置顶,4:栏目置顶,3:总推荐文章,else:普通")&vbNewLine)
  response.Write(Replacestr(Search_Rs("IsURL"),"0:,else: | 标题新闻"))
  response.Write(Replacestr(Search_Rs("isPicNews"),"0:,else: | <img border=0 title=""图"" src=""sys_images/img.jpg"">"))
  response.Write(Replacestr(Search_Rs("TodayNewsPic"),":,else: | 图片头条"))
  response.Write(Replacestr(Search_Rs("Source"),":,else: | "&Search_Rs("Source")))
  response.Write(Replacestr(Search_Rs("Author"),":,else: | "&Search_Rs("Author")))
  response.Write(Replacestr(Search_Rs("Hits"),":,else: | ["&Search_Rs("Hits")&"]"))
case "DS" 
  response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&addtime&vbNewLine)
  response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&Replacestr(Search_Rs("ClassName"),":,else:"&Search_Rs("ClassName")))
  response.Write(Replacestr(Search_Rs("RecTF"),"1: | 推荐,else:"))
  response.Write(Replacestr(Search_Rs("Types"),"1: | 图片,2: | 文件,3: | 程序,4: | Flash,5: | 音乐,6: | 影视,7: | 其它"))
  response.Write(Replacestr(Search_Rs("ClickNum"),":,else: | 热度["&Search_Rs("ClickNum")))
  response.Write(Replacestr(Search_Rs("Hits"),":,else: | "&Search_Rs("Hits")&"]"))
  response.Write(Replacestr(Search_Rs("ConsumeNum"),"0:,else: | 需要点数["&Search_Rs("ConsumeNum")&"]"))
case "WS" 
  response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&addtime&vbNewLine)
  response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&Replacestr(Search_Rs("ClassName"),":,else:"&Search_Rs("ClassName")))
  response.Write(Replacestr(Search_Rs("ParentID"),"0: | <strong>话题</strong>,else:回复"))
  response.Write(Replacestr(Search_Rs("User"),":未知用户,else:"&Search_Rs("User")))
  response.Write(Replacestr(Search_Rs("IsTop"),"1: | 推荐,else:"))
  response.Write(Replacestr(Search_Rs("Answer"),":,else: | 回复["&Search_Rs("Answer")&"]"))
  response.Write(Replacestr(Search_Rs("Hit"),":,else: | Hit["&Search_Rs("Hit")&"]"))
  response.Write(Replacestr(Search_Rs("LastUpdateDate"),":,else: | LastTime["&Search_Rs("LastUpdateDate")&""))
  response.Write(Replacestr(Search_Rs("LastUpdateUser"),":,else: | "&Search_Rs("LastUpdateUser")&"]"))
  response.Write(Replacestr(Search_Rs("IP"),":,else:"&Search_Rs("IP")))
  response.Write(Replacestr(Search_Rs("Vistor"),"0: | 游客可访,else:"))
case "MS"
  response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&addtime&vbNewLine)
  response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&Replacestr(Search_Rs("ClassCName"),":,else:"&Search_Rs("ClassCName")))
  response.Write(Replacestr(Search_Rs("StyleFlagBit"),":,else:"&Replacestr(split(Search_Rs("StyleFlagBit"),",")(0),"1: | 推荐,else:")))
  response.Write(Replacestr(Search_Rs("StyleFlagBit"),":,else:"&Replacestr(split(Search_Rs("StyleFlagBit"),",")(1),"1: | 热点,else:")))
  response.Write(Replacestr(Search_Rs("StyleFlagBit"),":,else:"&Replacestr(split(Search_Rs("StyleFlagBit"),",")(2),"1: | 置顶,else:")))
  response.Write(Replacestr(Search_Rs("StyleFlagBit"),":,else:"&Replacestr(split(Search_Rs("StyleFlagBit"),",")(3),"1: | 特价,else:")))
  response.Write(Replacestr(Search_Rs("StyleFlagBit"),":,else:"&Replacestr(split(Search_Rs("StyleFlagBit"),",")(4),"1: | 锁定,else:")))
  response.Write(Replacestr(Search_Rs("StyleFlagBit"),":,else:"&Replacestr(split(Search_Rs("StyleFlagBit"),",")(5),"1: | 促销,else:")))
  response.Write(Replacestr(Search_Rs("Click"),":,else: | ["&Search_Rs("Click")&"]"))
  response.Write(Replacestr(Search_Rs("Stockpile"),":,else: | 库存"&Search_Rs("Stockpile")))
  response.Write(Replacestr(Search_Rs("StockpileWarn"),"1:[<font color=red>库存不够</font>],else:"))
  response.Write(Replacestr(Search_Rs("OldPrice"),"0:,else: | [市场价"&Search_Rs("OldPrice")))
  response.Write(Replacestr(Search_Rs("NewPrice"),"0:0],else: | 商城价"&Search_Rs("NewPrice")&"]"))
  response.Write(Replacestr(left(Search_Rs("IsWholesale"),1),"1: | 可批发,else:"))
  response.Write(Replacestr(Search_Rs("MakeFactory"),":,else: | "&Search_Rs("MakeFactory")))
  response.Write(Replacestr(Search_Rs("ProductsAddress"),":,else: | "&replace(Search_Rs("ProductsAddress"),","," ")))
  response.Write(" | "&Replacestr(Search_Rs("SaleStyle"),"0:正常,1:竞拍,2:一口价,3:降价,4:特价"))
  response.Write(Replacestr(Search_Rs("Discount"),":,else: | 折率["&Search_Rs("Discount")&"]"))
  response.Write(Replacestr(Search_Rs("DiscountStartDate"),":,else:["&Search_Rs("DiscountStartDate")&"-"))
  response.Write(Replacestr(Search_Rs("DiscountEndDate"),":,else:"&Search_Rs("DiscountEndDate")&"]"))
case "SD" 
  response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&addtime&vbNewLine)
  response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&Replacestr(Search_Rs("ClassName"),":,else:"&Search_Rs("ClassName")))
  response.Write("|"&Replacestr(Search_Rs("GQ_ClassName"),":,else:"&Search_Rs("GQ_ClassName")))
  response.Write("|"&Replacestr(Search_Rs("CompType"),"0:批发,1:零售,2:批零兼营"))
  response.Write("|价格:"&Replacestr(Search_Rs("PubPrice"),"0:面议,else:"&Search_Rs("PubPrice")&"元"))
  response.Write("|数量:"&Replacestr(Search_Rs("PubNumber"),"0:充足,else:"&Search_Rs("PubNumber")))
  response.Write("|有效期限:"&Replacestr(Search_Rs("ValidTime"),":长期有效,0:长期有效,else:"&Search_Rs("ValidTime")&"天"))
  response.Write("|产地:"&Replacestr(Search_Rs("PubAddress"),":电话告之,else:"&Search_Rs("PubAddress")))
  response.Write("|人气:"&Replacestr(Search_Rs("Hits"),":,else:["&Search_Rs("Hits")&"]"))
  DBUserName = Get_OtherTable_Value("select C_Name from FS_ME_CorpUser where UserNumber='"&Search_Rs("UserNumber")&"'")	
  if DBUserName = "" then DBUserName = Get_OtherTable_Value("select UserName from FS_ME_Users where UserNumber='"&Search_Rs("UserNumber")&"'")
  if DBUserName = "" then 
  	DBUserName = "系统管理员"
  else
  	DBUserName = "<a href=""/User/ShowUser.asp?UserNumber="&Search_Rs("UserNumber")&""" title=""点击查看该用户详细信息"" target=_blank>"&DBUserName&"</a>"
  end if		
  response.Write("|发布者:"&DBUserName)
end select
  response.Write("</font>")
  response.Write(Content)
%>
</div>
<%
	
	''+++++++++++++++++++++++++++++++++++++++		
	Search_RS.MoveNext
	if Search_RS.eof or Search_RS.bof then exit for
  NEXT
%>
<BR clear=all>
<DIV class=n id=navbar style="text-align:center"> 
  <TABLE cellSpacing=0 cellPadding=0 width="1%" align=center border=0>
    <TBODY>
      <TR style="TEXT-ALIGN: center" vAlign=top align=middle> 
        <TD vAlign=bottom noWrap class=i><FONT size=-1>结果页码:&nbsp;</FONT> 
        <TD noWrap class="i"><font size=-1>&nbsp; 
		<%response.Write( fPageCount(Search_RS,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf )%>
		</font></TR>
    </TBODY>
  </TABLE>
</DIV> 
<%
END IF
RsClose
end select
Sub RsClose()
	Search_RS.Close
	Set Search_RS = Nothing
end Sub

Set Old_News_Conn = Nothing
Set User_Conn = Nothing
Set Conn = Nothing
response.End()
%>
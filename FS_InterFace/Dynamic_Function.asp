<!--#include file="NS_Public.asp" -->
<!--#include file="MS_Public.asp" -->
<!--#include file="DS_Public.asp" -->
<!--#include file="ME_Public.asp" -->
<!--#include file="MF_Public.asp" -->
<!--#include file="SD_Public.asp" -->
<!--#include file="HS_Public.asp" -->
<!--#include file="AP_Public.asp" -->
<!--#include file="Other_Public.asp" -->
<!--#include file="Refresh_Function.asp" -->
<%
Function Get_Dynamic_Refresh_Content(f_Templet,f_ID,f_sysFlag,f_Page,f_PageType)
	if Request.Cookies("FoosunSUBCookie")="" then
		MF_Default_Conn
		SubSys_Cookies:MFConfig_Cookies:NSConfig_Cookies:DSConfig_Cookies
		if Request.Cookies("FoosunSUBCookie")("FoosunSUBMS")=1 Then:MSConfig_Cookies:end if
	end if
	if f_Page < 1 then f_Page = 1
	Dim f_File_Content,f_Templet_Have_Error,f_LabelMassList,f_PaginationMass,f_RefreshContent
	f_File_Content = GetTempletContent(f_Templet,f_Templet_Have_Error)
	if Not f_Templet_Have_Error then
		if f_PageType <> "" then f_PageType = f_sysFlag & "_" & f_PageType else f_PageType = f_sysFlag
		Set f_LabelMassList = Replace_All_Flag(f_File_Content,f_ID,f_PageType)
		Set f_PaginationMass = ReplaceNoPaginationLableContent(f_File_Content,f_LabelMassList)
		f_RefreshContent = GetSaveContentForPage(f_File_Content,f_PaginationMass,f_Page)
	else
		f_RefreshContent = f_File_Content
	end if
	Get_Dynamic_Refresh_Content = f_RefreshContent
End Function

Function GetSaveContentForPage(f_SaveContent,f_PaginationMass,f_PageNO)
	Dim f_PaginationLable,i,j,f_PageCount,f_ParseContent,f_DictContent,f_PaginationStr,f_PaginationArray,f_MassParseContent
	if f_PaginationMass IS Nothing then
		if f_PageNO = 1 then GetSaveContentForPage = f_SaveContent else GetSaveContentForPage = ""
	else
		f_MassParseContent = f_PaginationMass.ParseContent
		For i = 0 To f_PaginationMass.FoosunLableList.Length - 1
			Set f_PaginationLable = f_PaginationMass.FoosunLableList.Items(i)
			if f_PaginationLable.IsPagination = True then
				Set f_DictContent = f_PaginationLable.DictLableContent
				f_PageCount = f_DictContent.Count - 1
				if f_DictContent.Exists(f_PageNO & "") then
					f_PaginationStr = f_DictContent.Item("0")
					f_PaginationArray = Split(f_PaginationStr,",")
					f_PaginationStr = Get_More_Page_Link_Str_Dynamic(f_PaginationArray(0),f_PaginationArray(1),f_PaginationArray(2),f_PageCount,f_PageNO)
					f_MassParseContent = Replace(f_MassParseContent,f_PaginationLable.LableName,f_DictContent.Item(f_PageNO & ""))
					f_MassParseContent = Replace(f_MassParseContent,"[FS:CONTENT_MOREPAGE_TAG]","")
				end if
				Exit For
			else
				f_PaginationMass.ParseContent = Replace(f_PaginationMass.ParseContent,f_PaginationLable.LableName,f_PaginationLable.DictLableContent.Item("1"))
				f_PaginationStr = ""
			end if
		Next
		GetSaveContentForPage = Replace(f_SaveContent,f_PaginationMass.MassName,f_MassParseContent & f_PaginationStr)
	end if
End Function

Function Get_More_Page_Link_Str_Dynamic(f_More_Page_Link_Type,f_More_Page_Link_Color,f_More_Page_Css,f_Page_Count,f_More_Page_Index)
	Dim f_i,Str_Link,LinkUrl,Str_Style
	Dim str_nonLinkColor,str_toF,str_toP10,str_toP1,str_toN1,str_toN10,str_toL,StartPage,EndPage,I
	If f_More_Page_Index>f_Page_Count Then
		f_More_Page_Index=f_Page_Count
	End If
	If f_More_Page_Link_Type="" Then
		f_More_Page_Link_Type=0
	End If
	Str_Link=""
	LinkUrl = ThisPageUrl("page","submit")
	Str_Style=""
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
					Str_Link=Str_Link&"&nbsp;<a href="""&LinkUrl&f_More_Page_Index+1&""""&Str_Style&">下一页</a>"
				ElseIf (f_More_Page_Index+1)>f_Page_Count Then
					Str_Link=Str_Link&"<a href="""&LinkUrl&f_More_Page_Index-1&""""&Str_Style&">上一页</a>"
					Str_Link=Str_Link&"&nbsp;下一页"
				Else
					Str_Link=Str_Link&"<a href="""&LinkUrl&f_More_Page_Index-1&""""&Str_Style&">上一页</a>"
					Str_Link=Str_Link&"&nbsp;<a href="""&LinkUrl&f_More_Page_Index+1&""""&Str_Style&">下一页</a>"
				End If
			Case 2
				Str_Link="共"&f_Page_Count&"页，"
				For f_i=1 To f_Page_Count
					If f_i= f_More_Page_Index Then
						Str_Link=Str_Link&",第"&f_i&"页"
					Else
						Str_Link=Str_Link&",<a href="""&LinkUrl&f_i&""""&Str_Style&">第"&f_i&"页</a>"
					End If
				Next
			Case 3
				Str_Link="共"&f_Page_Count&"页。"
				For f_i=1 To f_Page_Count
					If f_i= f_More_Page_Index Then
						Str_Link=Str_Link&"&nbsp;"&f_i&""
					Else
						Str_Link=Str_Link&"&nbsp;<a href="""&LinkUrl&f_i&""""&Str_Style&">"&f_i&"</a>"
					End If
				Next
				
			Case 5
				str_nonLinkColor="#999999" '非热链接颜色
				str_toF="|<<"  			'第一页
				str_toP10="<<"			'上十
				str_toP1="<"				'上一
				str_toN1=">"				'下一
				str_toN10=">>"			'下十
				str_toL=">>|"				'尾页


				Str_Link=""

				if f_More_Page_Index=1 then
					Str_Link=Str_Link& "<span>"&str_toF&"</span> " &vbNewLine
				Else
					Str_Link=Str_Link& "<a href="""&LinkUrl&"1"""&Str_Style&" title=""首页"">"&str_toF&"</a> " &vbNewLine
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
					Str_Link=Str_Link& "<a href="""&LinkUrl&f_More_Page_Index-10&""""&Str_Style&" title=""上10页"">"&str_toP10&"</a> "  &vbNewLine
				Else
					Str_Link=Str_Link& "<span>"&str_toP10&"</span> "  &vbNewLine
				End If

				If f_More_Page_Index > 1 Then
					Str_Link=Str_Link& "<a href="""&LinkUrl&f_More_Page_Index-1&""""&Str_Style&" title=""上一页"">"&str_toP1&"</a> "  &vbNewLine
				Else
					Str_Link=Str_Link& "<span>"&str_toP1&"</span>"  &vbNewLine
				End If

				For I=StartPage To EndPage
					If I=f_More_Page_Index Then
						Str_Link=Str_Link& "<a class=""currentPageCSS"" title=""当前页"" href=""javascript:void(0);"">"&I&"</a>"  &vbNewLine
					Else
						Str_Link=Str_Link& "<a href="""&LinkUrl&I&""""&Str_Style&">" &I& "</a>"  &vbNewLine
					End If
				Next
				If f_More_Page_Index < f_Page_Count Then
					Str_Link=Str_Link& " <a href="""&LinkUrl&f_More_Page_Index+1&""""&Str_Style&" title=""下一页"">"&str_toN1&"</a> "  &vbNewLine
				Else
					Str_Link=Str_Link& "<span>"&str_toN1&"</span>"  &vbNewLine
				End If

				If EndPage<f_Page_Count Then
					If (f_More_Page_Index+10)>f_Page_Count Then
						Str_Link=Str_Link& " <a href="""&LinkUrl&f_Page_Count&""""&Str_Style&"  title=""下10页"">"&str_toN10&"</a> "  &vbNewLine
					Else
						Str_Link=Str_Link& " <a href="""&LinkUrl&f_More_Page_Index+10&""""&Str_Style&" title=""下10页"">"&str_toN10&"</a> "  &vbNewLine
					End If
				Else
					Str_Link=Str_Link& " <span>"&str_toN10&"</span>"  &vbNewLine
				End If

				if f_More_Page_Index<f_Page_Count Then
					Str_Link=Str_Link& "<a href="""&LinkUrl&f_Page_Count&""""&Str_Style&" title=""尾页"">"&str_toL&"</a>"  &vbNewLine
				Else
					Str_Link=Str_Link& "<span>"&str_toL&"</span>"  &vbNewLine
				End If		
				Str_Link="<div class=""pagecontent"">"&Str_Link&"</div>"
			Case Else
				str_nonLinkColor="#999999" '非热链接颜色
				str_toF="<font face=""webdings"">9</font>"  			'首页
				str_toP10="<font face=""webdings"">7</font>"			'上十
				str_toP1="<font face=""webdings"">3</font>"				'上一
				str_toN1="<font face=""webdings"">4</font>"				'下一
				str_toN10="<font face=""webdings"">8</font>"			'下十
				str_toL="<font face=""webdings"">:</font>"				'尾页

				Str_Link=""

				if f_More_Page_Index=1 then
					Str_Link=Str_Link& "<font color="&str_nonLinkColor&" title=""首页"">"&str_toF&"</font> " &vbNewLine
				Else
					Str_Link=Str_Link& "<a href="""&LinkUrl&"1"""&Str_Style&" title=""首页"">"&str_toF&"</a> " &vbNewLine
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
					Str_Link=Str_Link& "<a href="""&LinkUrl&f_More_Page_Index-10&""""&Str_Style&" title=""上10页"">"&str_toP10&"</a> "  &vbNewLine
				Else
					Str_Link=Str_Link& "<font color="&str_nonLinkColor&" title=""上10页"">"&str_toP10&"</font> "  &vbNewLine
				End If

				If f_More_Page_Index > 1 Then
					Str_Link=Str_Link& "<a href="""&LinkUrl&f_More_Page_Index-1&""""&Str_Style&" title=""上一页"">"&str_toP1&"</a> "  &vbNewLine
				Else
					Str_Link=Str_Link& "<font color="&str_nonLinkColor&" title=""上一页"">"&str_toP1&"</font> "  &vbNewLine
				End If

				For I=StartPage To EndPage
					If I=f_More_Page_Index Then
						Str_Link=Str_Link& "<b>"&I&"</b>"  &vbNewLine
					Else
						Str_Link=Str_Link& "<a href="""&LinkUrl&I&""""&Str_Style&">" &I& "</a>"  &vbNewLine
					End If
				Next
				If f_More_Page_Index < f_Page_Count Then
					Str_Link=Str_Link& " <a href="""&LinkUrl&f_More_Page_Index+1&""""&Str_Style&" title=""下一页"">"&str_toN1&"</a> "  &vbNewLine
				Else
					Str_Link=Str_Link& "<font color="&str_nonLinkColor&" title=""下一页"">"&str_toN1&"</font> "  &vbNewLine
				End If

				If EndPage<f_Page_Count Then
					If (f_More_Page_Index+10)>f_Page_Count Then
						Str_Link=Str_Link& " <a href="""&LinkUrl&f_Page_Count&""""&Str_Style&"  title=""下10页"">"&str_toN10&"</a> "  &vbNewLine
					Else
						Str_Link=Str_Link& " <a href="""&LinkUrl&f_More_Page_Index+10&""""&Str_Style&" title=""下10页"">"&str_toN10&"</a> "  &vbNewLine
					End If
				Else
					Str_Link=Str_Link& " <font color="&str_nonLinkColor&"  title=""下10页"">"&str_toN10&"</font> "  &vbNewLine
				End If

				if f_More_Page_Index<f_Page_Count Then
					Str_Link=Str_Link& "<a href="""&LinkUrl&f_Page_Count&""""&Str_Style&" title=""尾页"">"&str_toL&"</a>"  &vbNewLine
				Else
					Str_Link=Str_Link& "<font color="&str_nonLinkColor&" title=""尾页"">"&str_toL&"</font>"  &vbNewLine
				End If
		End Select
	End If
	Get_More_Page_Link_Str_Dynamic="<div>"&Str_Link&"</div>"
End Function

Function ThisPageUrl(moveParam,removeList)
	dim strName
	dim KeepUrl,KeepForm,KeepMove
	removeList=removeList&","&moveParam
	KeepForm=""
	For Each strName in Request.Form
		'判断form参数中的submit、空值
		if not InstrRev(","&removeList&",",","&strName&",", -1, 1)>0 and Request.Form(strName)<>"" then
			KeepForm=KeepForm&"&"&strName&"="&Server.URLencode(Request.Form(strName))
		end if
		removeList=removeList&","&strName
	Next

	KeepUrl=""
	For Each strName In Request.QueryString
		If not (InstrRev(","&removeList&",",","&strName&",", -1, 1)>0) Then
			KeepUrl = KeepUrl & "&" & strName & "=" & Server.URLencode(Request.QueryString(strName))
		End If
	Next

	KeepMove=KeepForm&KeepUrl

	If (KeepMove <> "") Then
	  KeepMove = Right(KeepMove, Len(KeepMove) - 1)
	  KeepMove = Server.HTMLEncode(KeepMove) & "&"
	End If

	ThisPageUrl = Request.ServerVariables("URL") & "?" & KeepMove & moveParam & "="
End Function
%>
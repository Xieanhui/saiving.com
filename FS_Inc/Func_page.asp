<%''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
''��������
'Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
'int_RPP=2 '����ÿҳ��ʾ��Ŀ
'int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
'showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
'str_nonLinkColor_="#999999" '����������ɫ
'toF_="<font face=webdings>9</font>"  			'��ҳ 
'toP10_=" <font face=webdings>7</font>"			'��ʮ
'toP1_=" <font face=webdings>3</font>"			'��һ
'toN1_=" <font face=webdings>4</font>"			'��һ
'toN10_=" <font face=webdings>8</font>"			'��ʮ
'toL_="<font face=webdings>:</font>"				'βҳ

'============================================
'��δ���һ��Ҫ��VClass_Rs.Open �� forѭ��֮��
'	Set VClass_Rs = CreateObject(G_FS_RS)
'	VClass_Rs.Open This_Fun_Sql,User_Conn,1,1
'	IF not VClass_Rs.eof THEN 
'	VClass_Rs.PageSize=int_RPP
'	cPageNo=NoSqlHack(Request.QueryString("Page"))
'	If cPageNo="" Then cPageNo = 1
'	If not isnumeric(cPageNo) Then cPageNo = 1
'	cPageNo = Clng(cPageNo)
'	If cPageNo<=0 Then cPageNo=1
'	If cPageNo>VClass_Rs.PageCount Then cPageNo=VClass_Rs.PageCount 
'	VClass_Rs.AbsolutePage=cPageNo
'	  FOR int_Start=1 TO int_RPP 
	  ''++++++++++
	  '��ѭ������ʾ����
	  ''++++++++++
'		VClass_Rs.MoveNext
'		if VClass_Rs.eof or VClass_Rs.bof then exit for
'      NEXT
'	END IF	  
'============================================
'response.Write "<p>"&  fPageCount(VClass_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)

''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'*********************************************************
' Ŀ�ģ���ҳ��ҳ���������
'          �ύ��ѯ��һ����
' ���룺moveParam����ҳ����
'         removeList��Ҫ�Ƴ��Ĳ���
' ���أ���ҳUrl
'*********************************************************
Function PageUrl(moveParam,removeList)
	dim strName
	dim KeepUrl,KeepForm,KeepMove
	removeList=removeList&","&moveParam
	KeepForm=""
	For Each strName in Request.Form 
		'�ж�form�����е�submit����ֵ
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
	
	'PageUrl = replace(Request.ServerVariables("URL"),"/Search.asp","/Search.html") & "?" & KeepMove & moveParam & "="
	PageUrl =  "?" & KeepMove & moveParam & "="
End Function 


Function fPageCount(Page_Rs,showNumberLink_,nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,Page)
	Dim This_Func_Get_Html_,toPage_,p_,sp2_,I,tpagecount
	Dim NaviLength,StartPage,EndPage
	This_Func_Get_Html_ = ""
	I = 1   
	NaviLength=showNumberLink_ 
	if IsEmpty(showMorePageGo_Type_) then showMorePageGo_Type_ = 1
	tpagecount=Page_Rs.pagecount
	If tPageCount<1 Then tPageCount=1 
	if not Page_Rs.eof or not Page_Rs.bof then
		toPage_ = PageUrl("Page","submit,GetType,no-cache,_")
		if Page=1 then 
			This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""��ҳ"">"&toF_&"</font> " &vbNewLine
		Else 
			This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&"1 title=""��ҳ"">"&toF_&"</a> " &vbNewLine
		End If 
		if Page<NaviLength then
			StartPage = 1
		else
			StartPage = fix(Page / NaviLength) * NaviLength	
		end if	
		EndPage=StartPage+NaviLength-1 
		If EndPage>tPageCount Then EndPage=tPageCount 
		If StartPage>1 Then 
			This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_& Page - NaviLength &" title=""��"&int_showNumberLink_&"ҳ"">"&toP10_&"</a> "  &vbNewLine
		Else 
			This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""��"&int_showNumberLink_&"ҳ"">"&toP10_&"</font> "  &vbNewLine
		End If 
		If Page <> 1 and Page <>0 Then 
			This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&(Page-1)&"  title=""��һҳ"">"&toP1_&"</a> "  &vbNewLine
		Else 
			This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""��һҳ"">"&toP1_&"</font> "  &vbNewLine
		End If 
		For I=StartPage To EndPage 
			If I=Page Then 
				This_Func_Get_Html_=This_Func_Get_Html_& "<b>"&I&"</b>"  &vbNewLine
			Else 
				This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&I&">" &I& "</a>"  &vbNewLine
			End If 
			If I<>tPageCount Then This_Func_Get_Html_=This_Func_Get_Html_& vbNewLine
		Next 
		If Page <> Page_Rs.PageCount and Page <>0 Then 
			This_Func_Get_Html_=This_Func_Get_Html_& " <a href="&toPage_&(Page+1)&" title=""��һҳ"">"&toN1_&"</a> "  &vbNewLine
		Else 
			This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""��һҳ"">"&toN1_&"</font> "  &vbNewLine
		End If 
		If EndPage<tpagecount Then  
			This_Func_Get_Html_=This_Func_Get_Html_& " <a href="&toPage_& Page + NaviLength &"  title=""��"&int_showNumberLink_&"ҳ"">"&toN10_&"</a> "  &vbNewLine
		Else 
			This_Func_Get_Html_=This_Func_Get_Html_& " <font color="&nonLinkColor_&"  title=""��"&int_showNumberLink_&"ҳ"">"&toN10_&"</font> "  &vbNewLine
		End If 
		if Page_Rs.PageCount<>Page then  
			This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&Page_Rs.PageCount&" title=""βҳ"">"&toL_&"</a>"  &vbNewLine
		Else 
			This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""βҳ"">"&toL_&"</font>"  &vbNewLine
		End If 
		If showMorePageGo_Type_ = 1 then 
			Dim Show_Page_i
			Show_Page_i = Page + 1
			if Show_Page_i > tPageCount then Show_Page_i = 1
			This_Func_Get_Html_=This_Func_Get_Html_& "<input type=""text"" size=""4"" maxlength=""10"" name=""Func_Input_Page"" onmouseover=""this.focus();"" onfocus=""this.value='"&Show_Page_i&"';"" onKeyUp=""value=value.replace(/[^1-9]/g,'')"" onbeforepaste=""clipboardData.setData('text',clipboardData.getData('text').replace(/[^1-9]/g,''))"">" &vbNewLine _
				&"<input type=""button"" value=""Go"" onmouseover=""Func_Input_Page.focus();"" onclick=""javascript:var Js_JumpValue;Js_JumpValue=document.all.Func_Input_Page.value;if(Js_JumpValue=='' || !isNaN(Js_JumpValue)) location='"&topage_&"'+Js_JumpValue; else location='"&topage_&"1';"">"  &vbNewLine
		Else 
			This_Func_Get_Html_=This_Func_Get_Html_& " ��ת:<select NAME=menu1 onChange=""var Js_JumpValue;Js_JumpValue=this.options[this.selectedIndex].value;if(Js_JumpValue!='') location=Js_JumpValue;"">"
			for i=1 to tPageCount
				This_Func_Get_Html_=This_Func_Get_Html_& "<option value="&topage_&i
				if Page=i then This_Func_Get_Html_=This_Func_Get_Html_& " selected style='color:#0000FF'"
				This_Func_Get_Html_=This_Func_Get_Html_& ">��"&cstr(i)&"ҳ</option>" &vbNewLine
			next
			This_Func_Get_Html_=This_Func_Get_Html_& "</select>" &vbNewLine
		End if
		This_Func_Get_Html_=This_Func_Get_Html_& p_&sp2_&" &nbsp;ÿҳ<b>"&Page_Rs.PageSize&"</b>����¼��������:<b><span class=""tx"">"&sp2_&Page&"</span>/"&tPageCount&"</b>ҳ����<b><span id='recordcount'>"&sp2_&Page_Rs.recordCount&"</span></b>����¼��"
	else
		This_Func_Get_Html_ = ""
	end if
	fPageCount = This_Func_Get_Html_
End Function

Function fPageCountNews(Page_Rs_Count,perCount,showNumberLink_,nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,Page)
	Dim This_Func_Get_Html_,toPage_,p_,sp2_,I,tpagecount
	Dim NaviLength,StartPage,EndPage,pageCount
	if perCount*(Page_Rs_Count\perCount)<Page_Rs_Count then 
		pageCount=(Page_Rs_Count\perCount)+1
	else
		pageCount=(Page_Rs_Count\perCount)
	end if
	This_Func_Get_Html_ = ""  : I = 1   
	NaviLength=showNumberLink_ 
	if IsEmpty(showMorePageGo_Type_) then showMorePageGo_Type_ = 1
	tpagecount=Page_Rs_Count
	If tPageCount<1 Then tPageCount=1 
	if tPageCount>0 then
		toPage_ = PageUrl("Page","submit,GetType,no-cache,_")
		if Page=1 then 
			This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""��ҳ"">"&toF_&"</font> " &vbNewLine
		Else 
			This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&"1 title=""��ҳ"">"&toF_&"</a> " &vbNewLine
		End If 
		if Page<NaviLength then
			StartPage = 1
		else
			StartPage = fix(Page / NaviLength) * NaviLength	
		end if	
		EndPage=StartPage+NaviLength-1 
		If EndPage>pageCount Then EndPage=pageCount 
		If StartPage>1 Then 
			This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_& Page - NaviLength &" title=""��"&int_showNumberLink_&"ҳ"">"&toP10_&"</a> "  &vbNewLine
		Else 
			This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""��"&int_showNumberLink_&"ҳ"">"&toP10_&"</font> "  &vbNewLine
		End If 
		If Page <> 1 and Page <>0 Then 
			This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&(Page-1)&"  title=""��һҳ"">"&toP1_&"</a> "  &vbNewLine
		Else 
			This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""��һҳ"">"&toP1_&"</font> "  &vbNewLine
		End If 
		For I=StartPage To EndPage 
			If I=Page Then 
				This_Func_Get_Html_=This_Func_Get_Html_& "<b>"&I&"</b>"  &vbNewLine
			Else 
				This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&I&">" &I& "</a>"  &vbNewLine
			End If 
			If I<>tPageCount Then This_Func_Get_Html_=This_Func_Get_Html_& vbNewLine
		Next 
		If Page <> pageCount and Page <>0 Then 
			This_Func_Get_Html_=This_Func_Get_Html_& " <a href="&toPage_&(Page+1)&" title=""��һҳ"">"&toN1_&"</a> "  &vbNewLine
		Else 
			This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""��һҳ"">"&toN1_&"</font> "  &vbNewLine
		End If 
		If EndPage<pageCount Then  
			This_Func_Get_Html_=This_Func_Get_Html_& " <a href="&toPage_& Page + NaviLength &"  title=""��"&int_showNumberLink_&"ҳ"">"&toN10_&"</a> "  &vbNewLine
		Else 
			This_Func_Get_Html_=This_Func_Get_Html_& " <font color="&nonLinkColor_&"  title=""��"&int_showNumberLink_&"ҳ"">"&toN10_&"</font> "  &vbNewLine
		End If 
		if pageCount<>Page then  
			This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&pageCount&" title=""βҳ"">"&toL_&"</a>"  &vbNewLine
		Else 
			This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""βҳ"">"&toL_&"</font>"  &vbNewLine
		End If 
		If showMorePageGo_Type_ = 1 then 
			Dim Show_Page_i
			Show_Page_i = Page + 1
			if Show_Page_i > tPageCount then Show_Page_i = 1
			This_Func_Get_Html_=This_Func_Get_Html_& "<input type=""text"" size=""4"" maxlength=""10"" name=""Func_Input_Page"" onmouseover=""this.focus();"" onfocus=""this.value='"&Show_Page_i&"';"" onKeyUp=""value=value.replace(/[^1-9]/g,'')"" onbeforepaste=""clipboardData.setData('text',clipboardData.getData('text').replace(/[^1-9]/g,''))"">" &vbNewLine _
				&"<input type=""button"" value=""Go"" onmouseover=""Func_Input_Page.focus();"" onclick=""javascript:var Js_JumpValue;Js_JumpValue=document.all.Func_Input_Page.value;if(Js_JumpValue=='' || !isNaN(Js_JumpValue)) location='"&topage_&"'+Js_JumpValue; else location='"&topage_&"1';"">"  &vbNewLine
		Else 
			This_Func_Get_Html_=This_Func_Get_Html_& " ��ת:<select NAME=menu1 onChange=""var Js_JumpValue;Js_JumpValue=this.options[this.selectedIndex].value;if(Js_JumpValue!='') location=Js_JumpValue;"">"
			for i=1 to pageCount
				This_Func_Get_Html_=This_Func_Get_Html_& "<option value="&topage_&i
				if Page=i then This_Func_Get_Html_=This_Func_Get_Html_& " selected style='color:#0000FF'"
				This_Func_Get_Html_=This_Func_Get_Html_& ">��"&cstr(i)&"ҳ</option>" &vbNewLine
			next
			This_Func_Get_Html_=This_Func_Get_Html_& "</select>" &vbNewLine
		End if
		This_Func_Get_Html_=This_Func_Get_Html_& p_&sp2_&" &nbsp;ÿҳ<b>"&perCount&"</b>����¼��������:<b><span class=""tx"">"&sp2_&Page&"</span>/"&pageCount&"</b>ҳ����<b><span id='recordcount'>"&sp2_&Page_Rs_Count&"</span></b>����¼��"
	else
		This_Func_Get_Html_ = ""
	end if
	fPageCountNews = This_Func_Get_Html_
End Function
%>
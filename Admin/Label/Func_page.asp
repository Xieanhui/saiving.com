<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
''调用例子
'Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
'int_RPP=2 '设置每页显示数目
'int_showNumberLink_=8 '数字导航显示数目
'showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
'str_nonLinkColor_="#999999" '非热链接颜色
'toF_="<font face=webdings>9</font>"  			'首页 
'toP10_=" <font face=webdings>7</font>"			'上十
'toP1_=" <font face=webdings>3</font>"			'上一
'toN1_=" <font face=webdings>4</font>"			'下一
'toN10_=" <font face=webdings>8</font>"			'下十
'toL_="<font face=webdings>:</font>"				'尾页
 
'============================================
'这段代码一定要在VClass_Rs.Open 与 for循环之间
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
	  '加循环体显示数据
	  ''++++++++++
'		VClass_Rs.MoveNext
'		if VClass_Rs.eof or VClass_Rs.bof then exit for
'      NEXT
'	END IF	  
'============================================
'response.Write "<p>"&  fPageCount(VClass_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)

''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function fPageCount(Page_Rs,showNumberLink_,nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,Page)

Dim This_Func_Get_Html_,toPage_,p_,sp2_,I,tpagecount
Dim NaviLength,StartPage,EndPage

This_Func_Get_Html_ = ""  : I = 1
NaviLength=showNumberLink_ 

if IsEmpty(showMorePageGo_Type_) then showMorePageGo_Type_ = 1
tpagecount=Page_Rs.pagecount
If tPageCount<1 Then tPageCount=1 

if not Page_Rs.eof or not Page_Rs.bof then

if request.ServerVariables("QUERY_STRING")<>"" then 
	toPage_ = request.ServerVariables("SCRIPT_NAME")&"?"&request.ServerVariables("QUERY_STRING")
	if instr(toPage_,"Page=")>0 then 
		toPage_ = mid(toPage_,1,instrrev( toPage_,"Page=" ) + len("Page=") - 1 )
	else
		toPage_ = toPage_ & "&Page="	
	end if
else
	toPage_ = request.ServerVariables("SCRIPT_NAME")&"?Page="
end if

'This_Func_Get_Html_=This_Func_Get_Html_& "<form NAME=pageform ID=pageform>"

if Page=1 then 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""首页"">"&toF_&"</font> " &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&"1 title=""首页"">"&toF_&"</a> " &vbNewLine
End If 
if Page<NaviLength then
	StartPage = 1
else
	StartPage = fix(Page / NaviLength) * NaviLength	
end if	
EndPage=StartPage+NaviLength-1 
If EndPage>tPageCount Then EndPage=tPageCount 

If StartPage>1 Then 
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_& Page - NaviLength &" title=""上"&int_showNumberLink_&"页"">"&toP10_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""上"&int_showNumberLink_&"页"">"&toP10_&"</font> "  &vbNewLine
End If 

If Page <> 1 and Page <>0 Then 
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&(Page-1)&"  title=""上一页"">"&toP1_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""上一页"">"&toP1_&"</font> "  &vbNewLine
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
	This_Func_Get_Html_=This_Func_Get_Html_& " <a href="&toPage_&(Page+1)&" title=""下一页"">"&toN1_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""下一页"">"&toN1_&"</font> "  &vbNewLine
End If 

If EndPage<tpagecount Then  
	This_Func_Get_Html_=This_Func_Get_Html_& " <a href="&toPage_& Page + NaviLength &"  title=""下"&int_showNumberLink_&"页"">"&toN10_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& " <font color="&nonLinkColor_&"  title=""下"&int_showNumberLink_&"页"">"&toN10_&"</font> "  &vbNewLine
End If 

if Page_Rs.PageCount<>Page then  
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&Page_Rs.PageCount&" title=""尾页"">"&toL_&"</a>"  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""尾页"">"&toL_&"</font>"  &vbNewLine
End If 

If showMorePageGo_Type_ = 1 then 

	This_Func_Get_Html_=This_Func_Get_Html_& "<input type=""text"" size=""4"" name=""Func_Input_Page"" value="""&Page+1&""">" &vbNewLine _
		&"<input type=""button"" value=""Go"" onclick=""javascript:var Js_JumpValue;Js_JumpValue=document.all.Func_Input_Page.value;if(! isNaN(Js_JumpValue)) location='"&topage_&"'+Js_JumpValue; else alert('1、输入的页面必须是数字“Js_JumpValue='+Js_JumpValue+'”\n2、当程序多次调用分页函数时不能用该模式');"">"  &vbNewLine

Else

	This_Func_Get_Html_=This_Func_Get_Html_& " 跳转:<select NAME=menu1 onChange=""var Js_JumpValue;Js_JumpValue=this.options[this.selectedIndex].value;if(Js_JumpValue!='') location=Js_JumpValue;"">"
	for i=1 to tPageCount
		This_Func_Get_Html_=This_Func_Get_Html_& "<option value="&topage_&i
		if Page=i then This_Func_Get_Html_=This_Func_Get_Html_& " selected style='color:#0000FF'"
		This_Func_Get_Html_=This_Func_Get_Html_& ">第"&cstr(i)&"页</option>" &vbNewLine
	next
	This_Func_Get_Html_=This_Func_Get_Html_& "</select>" &vbNewLine

End if

This_Func_Get_Html_=This_Func_Get_Html_& p_&sp2_&" "&Page_Rs.PageSize&"项 "&sp2_&Page&"/"&tPageCount&"页 "&sp2_&Page_Rs.recordCount&"项"

'This_Func_Get_Html_=This_Func_Get_Html_& "</form>" 
else
	'没有记录
end if
fPageCount = This_Func_Get_Html_
End Function
%>








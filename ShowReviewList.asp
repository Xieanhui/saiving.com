<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<%session.CodePage="936"%>
<!--#include file="FS_Inc/Const.asp" -->
<!--#include file="FS_InterFace/MF_Function.asp" -->
<!--#include file="FS_Inc/Function.asp" -->
<!--#include file="FS_Inc/Func_Page.asp" -->
<!--#include file="FS_Inc/Md5.asp" -->
<%dim cssnum
cssnum = Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")
if isnull(cssnum) or cssnum="" then cssnum = "2"
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>查看评论---网站内容管理系统</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="<%=G_USER_DIR%>/images/skin/Css_<%=cssnum%>/<%=cssnum%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<iframe name="TempFrame" id="TempFrame" frameborder="0" src="" width="0" height="0"></iframe>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
  <tr>
    <td><%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	response.Charset = "gb2312"
	Dim User_Conn,review_Sql,review_RS,review_RS1,strShowErr,Cookie_Domain
	Dim Server_Name,Server_V1,Server_V2,FormReview
	Dim TmpStr,TmpArr,ReviewTypes
	Dim stype,Id,SpanId
	TmpStr = "" 
	Dim Conn 
	  
	MF_Default_Conn

	Cookie_Domain = Get_MF_Domain()
	if Cookie_Domain="" then 
		Cookie_Domain = "localhost"
	else
		if left(lcase(Cookie_Domain),len("http://"))="http://" then Cookie_Domain = mid(Cookie_Domain,len("http://")+1)
		if right(Cookie_Domain,1)="/" then Cookie_Domain = mid(Cookie_Domain,1,len(Cookie_Domain) - 1)
	end if	
	''防盗连
	Dim Main_Name,Name_Str1,V_MainName,V_Str
	Server_Name = NoHtmlHackInput(NoSqlHack(LCase(Request.ServerVariables("SERVER_NAME"))))
'	IF Server_Name <> LCase(Split(Cookie_Domain,"/")(0)) Then
'		Response.Write ("没有权限访问")
'		Response.End
'	End If
	Server_V1 = NoHtmlHackInput(NoSqlHack(Replace(Lcase(Cstr(Request.ServerVariables("HTTP_REFERER"))),"http://","")))
	Server_V1 = Replace(Replace(Server_V1,"//","/"),"///","/")
	IF Server_V1 = "" Then
		Response.Write ("没有权限访问")
		Response.End
	End If
	IF Instr(Server_V1,"/") = 0 Then
		Server_V2 = Server_V1
	Else
		Server_V2 = Split(Server_V1,"/")(0)
	End If	
	If Instr(Server_Name,".") = 0 Then
		Main_Name = Server_Name
	Else
		Name_Str1 = Split(Server_Name,".")(0)
		Main_Name = Trim(Replace(Server_Name,Name_Str1 & ".",""))
	End If
	If Instr(Server_V2,".") = 0 Then
		V_MainName = Server_V2
	Else
		V_Str = Split(Server_V2,".")(0)
		V_MainName = Trim(Replace(Server_V2,V_Str & ".",""))
	End If
	If Main_Name <> V_MainName And (Main_Name = "" OR V_MainName = "") Then
		Response.Write ("没有权限访问")
		Response.End
	End If
	
	stype = NoSqlHack(request.QueryString("type")) 'NS
	Id = CintStr(request.QueryString("Id")) 'Id
	if stype="" then response.Write("Error:type is null!</body></html>")  :  response.End()
	if Id="" or not isnumeric(Id) then response.Write("Error:Id is no number!</body></html>")  :  response.End()
	select case stype
		case "NS"
		ReviewTypes=0
		case "DS"
		ReviewTypes=1
		case "MS"
		ReviewTypes=2
		case "HS"
		ReviewTypes=3
		case "SD"
		ReviewTypes=4
		case "LOG"
		ReviewTypes=5	
		case else
			response.Write("Error:type("&stype&") is not found!</body></html>")  :  response.End()
	end select
	
	MF_User_Conn
'//------2007-01-23 By Ken 所有评论页面增加发表评论表单
	Dim ReName,RePwds,ReNoNameTF,ReTitle,ReContent,ReTypes,ReID,ReUserNum
	Dim CheckUserRs,AuditTF,AuditTFRs,ReIP
	Dim ReRs
	'----评论入库
	If Request.Form("ReAction") = "AddNew" Then
		ReName = NoSqlHack(Request.Form("UserNumber"))
		RePwds = NoSqlHack(Request.Form("password"))
		ReNoNameTF = NoSqlHack(Request.Form("noname"))
		ReTitle = NoHtmlHackInput(ReplaceKeys(NoSqlHack(Trim(Request.Form("title")))))
		ReContent = NoHtmlHackInput(ReplaceKeys(NoSqlHack(Trim(Request.Form("content")))))
		ReTypes = ReviewTypes
		ReID = Id
		If ReNoNameTF = "" Then
			ReNoNameTF = "0"
		Else
			ReNoNameTF = ReNoNameTF
		End If		
		If ReTitle = "" Or ReContent = "" Then
			Response.Write "<script>alert('标题或评论内容不能为空');</script>"
			Response.End
		End If
		If Len(ReTitle)	> 100 Then
			Response.Write "<script>alert('标题不要超过100字');</script>"
			Response.End
		End If
		If Len(ReContent) > 1000 Then
			Response.Write "<script>alert('内容不要超过1000字');</script>"
			Response.End	
		End If		
		If ReNoNameTF <> "1" Then  '--非匿名
			If ReName = "" Or RePwds = "" Then
				IF Session("FS_UserNumber") <> "" And Session("FS_UserPassword") <> "" Then
					ReUserNum = Session("FS_UserNumber")
				Else
					Response.Write "<script>alert('用户名或密码不能为空');</script>"
					Response.End
				End If
			Else
				Set CheckUserRs = User_Conn.ExeCute("Select UserNumber From FS_ME_Users Where UserName = '"& ReName &"' And UserPassword = '"& Md5(RePwds,16) &"'")
				If Not CheckUserRs.Eof Then
					Session("FS_UserNumber") = CheckUserRs(0)
					Session("FS_UserPassword") = Md5(RePwds,16)
					ReUserNum = CheckUserRs(0)
				Else
					Response.Write "<script>alert('用户名或密码错误');</script>"
					Response.End
				End If		
				CheckUserRs.Close : Set CheckUserRs = Nothing
			End If
		Else'匿名
			ReUserNum = 0	
		End If
		'----是否需要审核
		Set AuditTFRs = User_Conn.ExeCute("Select Top 1 ReviewTF From FS_ME_SysPara Order By SysID")
		IF Not AuditTFRs.Eof Then
			AuditTF = AuditTFRs(0)	
		Else
			AuditTF = 1
		End If
		AuditTFRs.Close : Set AuditTFRs = Nothing
		'---获取ip
		ReIP = NoSqlHack(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))
		If ReIP = "" Then ReIP = NoSqlHack(Request.ServerVariables("REMOTE_ADDR"))
		'---写数据库
		Set ReRs = Server.CreateObject(G_FS_RS)
		ReRs.Open "Select UserNumber,InfoID,ReviewTypes,Title,Content,AddTime,ReviewIP,isLock,QuoteID,AdminLock From FS_ME_Review Where 1 =2",User_Conn,1,3 
		ReRs.AddNew
		ReRs(0) = ReUserNum
		ReRs(1) = ReID
		ReRs(2) = ReTypes
		ReRs(3) = ReTitle
		ReRs(4) = ReContent
		ReRs(5) = Now()
		ReRs(6) = ReIP
		ReRs(7) = 0
		ReRs(8) = 0
		ReRs(9) = AuditTF
		ReRs.Update
		ReRs.Close
		Set ReRs = NoThing
		Response.Write "<script>alert('评论成功');parent.location.reload();</script>"
	End If
	'---评论表单
	FormReview = "<table border=0 width=""100%"" cellspacing=""1"" cellpadding=""2"" align=center class=""table"">"
	FormReview = FormReview & "<tr><td class=""xingmu"">发表评论：</td></tr>"
	FormReview = FormReview & "<tr>"
	FormReview = FormReview & "<td class=""xingmu"">"
	FormReview= FormReview & "<form action="""" name=""reviewform"" method=""post"" style=""margin:0px;"" target=""TempFrame"">"
	if session("FS_UserNumber")= "" Or session("FS_UserPassword") = "" then 
		FormReview = FormReview & "<tr><td class=""hback_1"" height=""26"">"
		FormReview=FormReview&"用户名<input name=""UserNumber"" type=""text"" id=""UserNumber"" size=""15"" />"
		FormReview=FormReview&"密码<input name=""password"" type=""password"" id=""password"" size=""12""/>"
		FormReview=FormReview&"匿名<input name=""noname"" type=""checkbox"" id=""noname"" value=""1"">"
		FormReview = FormReview & "</td></tr>"
		FormReview = FormReview & "<tr><td class=""hback_1"" height=""26"">"
		FormReview=FormReview&"标　题<input name=""title"" type=""text"" id=""title"" size=""40""/>"
		FormReview = FormReview & "</td></tr>"
	else
		FormReview = FormReview & "<tr><td class=""hback_1"" height=""26"">"
		FormReview=FormReview&"标　题<input name=""title"" type=""text"" id=""title"" size=""36""/>"
		FormReview=FormReview&"&nbsp;匿名<input name=""noname"" type=""checkbox"" id=""noname"" value=""1""/>"
		FormReview = FormReview & "</td></tr>"
	end if
	FormReview = FormReview & "<tr><td class=""xingmu"">"
	FormReview=FormReview&"<textarea name=""content"" style=""width:98%;"" rows=""5"" id=""content""/></textarea><input type=""hidden"" name=""ReAction"" value=""AddNew""/>"
	FormReview = FormReview & "</td></tr>"
	FormReview = FormReview & "<tr><td class=""xingmu"">"
	FormReview=FormReview&"<input type=""submit"" name=""Submit"" value=""发表评论""/>&nbsp;&nbsp;<input type=""reset"" name=""Submit2"" value=""重新填写""/>"
	FormReview = FormReview & "</td></tr>"
	FormReview=FormReview&"</form>"
	FormReview = FormReview & "</td></tr></table><br>"
	
	Response.Write FormReview 

	Function ReplaceKeys(Content)
		ReplaceKeys=Content
		Dim KeyRs,KWDs,KArray,k
		If Content = "" Or IsNull(Content) Then 
			ReplaceKeys = ""
			Exit Function
		End If
		Set KeyRs = User_Conn.ExeCute("Select Top 1 LimitReviewChar From FS_ME_SysPara Where SysID > 0 Order By SysID")
		If KeyRs.Eof Then
			ReplaceKeys = Content
		Else
			KWDs = KeyRs(0)
			If KWDs = "" Or IsNull(KWDs) Then
				ReplaceKeys = Content
			Else
				If Instr(KWDs,",") > 0 Then
					KArray = Split(KWDs,",")
					For k = Lbound(KArray) To Ubound(KArray)
						ReplaceKeys = Replace(ReplaceKeys,KArray(k),"**")
					Next
				Else
					ReplaceKeys = Replace(ReplaceKeys,KWDs,"**")
				End If
			End If
		End If
		KeyRs.Close : Set KeyRs = Nothing					
	End Function

'//------------------发表评论表单结束	
	
	
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
	int_RPP=15 '设置每页显示数目
	int_showNumberLink_=10 '数字导航显示数目
	showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
	str_nonLinkColor_="#999999" '非热链接颜色
	toF_="<font face=webdings>9</font>"   			'首页 
	toP10_=" <font face=webdings>7</font>"			'上十 
	toP1_=" <font face=webdings>3</font>"			'上一
	toN1_=" <font face=webdings>4</font>"			'下一
	toN10_=" <font face=webdings>8</font>"			'下十
	toL_="<font face=webdings>:</font>"				'尾页
	
	Function review_Data()	
		On Error Resume Next
		Dim UserName
		UserName=""
		review_Data = "<table border=0 width=""100%"" cellspacing=""1"" cellpadding=""2"" align=center class=""table"">"&vbNewLine
		review_Data = review_Data & "<tr><td class=""xingmu"">所有评论：</td></tr>"&vbNewLine
		review_Sql = "select ReviewID,UserNumber,Title,Content,AddTime,ReviewIP from FS_ME_Review where isLock=0 and AdminLock=0 and ReviewTypes="&ReviewTypes&" and InfoID="&CintStr(ID)&" order by ReviewID desc"
		set review_RS = CreateObject(G_FS_RS)
		review_RS.Open review_Sql,User_Conn,1,1
		if Err<>0 then 
			Err.Clear
			review_Data = "<tr><td colspan=2>抱歉，系统错误。</td></tr></table>"&vbNewline:exit function
		end if	
		if review_RS.eof then
			review_Data = "<tr><td colspan=2>抱歉，暂无评论。</td></tr>"&vbNewline
		else
			review_RS.PageSize=int_RPP
			cPageNo=CintStr(Request.QueryString("Page"))
			If cPageNo="" Then cPageNo = 1
			If not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo<=0 Then cPageNo=1
			If cPageNo>review_RS.PageCount Then cPageNo=review_RS.PageCount 
			review_RS.AbsolutePage=cPageNo
			
			  FOR int_Start=1 TO int_RPP 
			  
				set review_RS1=User_Conn.execute("select UserName from FS_ME_Users where UserNumber='"&review_RS("UserNumber")&"'")
				if not review_RS1.eof then 
					UserName = review_RS1("UserName")
					'if session("FS_UserNumber")<>"" then 
						UserName = "·<a href=""http://"&Cookie_Domain&"/"&G_USER_DIR&"/ShowUser.asp?UserNumber="&review_RS("UserNumber")&""" title=""点击查看该用户信息"" target=""_blank"">"&UserName&"</a>"
					'end if
				else
					UserName = "·匿名"	
				end if
				review_Data = review_Data &"<tr><td class=""hback_1"" height=""26"">"&UserName&"&nbsp;&nbsp;"&review_RS("Title")&"&nbsp;&nbsp;日期:"&review_RS("AddTime")&"&nbsp;&nbsp;IP:"&showip(review_RS("ReviewIP"))&"</td></tr>"&vbNewLine
				review_Data = review_Data &"<tr><td height=""30"" class=""hback""><font style=""font-size:14px"">"&Replace(Replace(Replace(review_RS("Content"),CHR(13),""),CHR(10)&CHR(10),"</P><P>"),CHR(10),"<BR>")&"</font></td></tr>"&vbNewLine
				review_RS1.close
				review_RS.movenext
				if review_RS.eof or review_RS.bof then exit for
			  NEXT
			  review_Data = review_Data &"<tr><td>"& fPageCount(review_RS,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo) &"</td></tr>"&vbNewLine
		end if
		review_Data = review_Data &"<tr><td align=center><button onClick=""window.close();"">关 · 闭</button></td></tr>"&vbNewLine
		review_Data = review_Data &"</table>"&vbNewLine	
	End Function

	function showip(ip)
		dim tmp_1,arr_1,ii_1
		tmp_1 = ""
		if ip="" or isnull(ip) then showip="":exit function
		arr_1 = split(ip,".")
		for ii_1=0 to ubound(arr_1)
			if ii_1<2 then 
				tmp_1 = tmp_1 &"."&arr_1(ii_1)
			else
				tmp_1 = tmp_1 & ".*"
			end if		
		next
		showip = mid(tmp_1,2)
	end function

	
	response.Write(review_Data())
	User_ConnClose()
	
	Sub RsClose()
		review_RS.Close
		Set review_RS = Nothing
	end Sub
	
	Sub User_ConnClose()
		Set Conn = Nothing
		Set User_Conn = Nothing
		response.Write(vbNewline&"</body></html>")
		response.End()
	End Sub
%></td>
  </tr>
</table>

</body>
</html>






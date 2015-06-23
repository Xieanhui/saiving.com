<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<%session.CodePage="936"%>
<!--#include file="FS_Inc/Const.asp" -->
<!--#include file="FS_InterFace/MF_Function.asp" -->
<!--#include file="FS_Inc/Function.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	response.Charset = "gb2312"
	Dim User_Conn,review_Sql,review_RS,review_RS1,strShowErr,Cookie_Domain
	Dim Server_Name,Server_V1,Server_V2
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
	Server_Name = NoHtmlHackInput(NoSqlHack(LCase(Trim(Request.ServerVariables("SERVER_NAME")))))
	if Request.ServerVariables("Server_Port")<>80 then
		Server_Name = Server_Name&":"&Request.ServerVariables("Server_Port")
	end if
	IF Server_Name <> LCase(Split(Cookie_Domain,"/")(0)) Then
		Response.Write ("没有权限访问")
		Response.End
	End If
	Server_V1 = NoHtmlHackInput(NoSqlHack(Trim(Replace(Lcase(Cstr(Request.ServerVariables("HTTP_REFERER"))),"http://",""))))
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
	SpanId = NoSqlHack(request.QueryString("SpanId"))
	if stype="" then response.Write("Error:type is null!")  :  response.End()
	if Id="" or not isnumeric(Id) then response.Write("Error:Id is no Validata!")  :  response.End()
	if SpanId="" then response.Write("Error:SpanId is null!")  :  response.End()
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
			response.Write("Error:type("&stype&") is not found!")  :  response.End()
	end select
	
	MF_User_Conn
	
	Function review_Data()	
		Dim UserName
		UserName=""
		review_Sql = "select top 10 ReviewID,UserNumber,Title,Content,AddTime,ReviewIP from FS_ME_Review where isLock=0 and AdminLock=0 and ReviewTypes="&CintStr(ReviewTypes)&" and InfoID="&CintStr(ID)&" order by ReviewID desc"
		set review_RS = CreateObject(G_FS_RS)
		review_RS.Open review_Sql,User_Conn,1,1
		if not review_RS.eof then
			review_Data = "<table border=0 width=""100%"" align=center>"&vbNewLine
			review_Data = review_Data &"<tr><td align=right bgcolor=""#efefef""><a href=""http://"&Cookie_Domain&"/ShowReviewList.asp?Type="&stype&"&Id="&Id&""" target=""_blank"">更多评论</td></tr>"&vbNewLine
			do while not review_RS.eof 
				set review_RS1=User_Conn.execute("select UserName from FS_ME_Users where UserNumber='"&review_RS("UserNumber")&"'")
				if not review_RS1.eof then 
					UserName = review_RS1("UserName")
					'if session("FS_UserNumber")<>"" then 
						UserName = "<a href=""http://"&Cookie_Domain&"/"&G_USER_DIR&"/ShowUser.asp?UserNumber="&review_RS("UserNumber")&""" title=""点击查看该用户信息"" target=""_blank"">"&UserName&"</a>"
					'end if
				else
					UserName = "匿名"	
				end if
				review_Data = review_Data &"<tr><td>"&UserName&"&nbsp;&nbsp;"&GetCStrLen(review_RS("Title"),20)&"&nbsp;&nbsp;"&review_RS("AddTime")&"&nbsp;&nbsp;"&showip(review_RS("ReviewIP"))&"</td></tr>"&vbNewLine
				review_Data = review_Data &"<tr><td>"&GetCStrLen(review_RS("Content"),60)&"..."&"</td></tr>"&vbNewLine
				review_RS1.close
				if review_RS.eof then exit do : exit Function 	
				review_RS.movenext
			loop	
			review_Data = review_Data &"</table>"&vbNewLine
		end if	
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
		response.End()
	End Sub
%>







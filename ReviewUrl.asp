<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<%session.CodePage="936"%>
<!--#include file="FS_Inc/Const.asp" -->
<!--#include file="FS_InterFace/MF_Function.asp" -->
<!--#include file="FS_Inc/Function.asp" -->
<!--#include file="FS_Inc/Md5.asp" -->
<!--#include file="FS_Inc/FS_Users_conformity.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
response.Charset = "gb2312"
Dim Conn,User_Conn,review_Sql,review_RS,Cookie_Domain
Dim Server_Name,Server_V1,Server_V2
Dim TmpStr,TmpArr,ReviewTypes,needAudited,ReviewIP
Dim stype,Id,UserNumber,noname,password,title,Action,content,LimitReviewChar
TmpStr = "":needAudited = True

MF_Default_Conn
MF_User_Conn

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
'response.Write(Server_Name&"--"&LCase(Split(Cookie_Domain,"/")(0)))
'response.End()
'IF Server_Name <> LCase(Split(Cookie_Domain,"/")(0)) Then
'	Response.Write ("没有权限访问")
'	Response.End
'End If
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
	Main_Name = NoSqlHack(Replace(Server_Name,Name_Str1 & ".",""))
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





stype = NoSqlHack(request.Form("type")) 'NS
Id = CintStr(request.Form("Id")) 'Id
UserNumber = NoSqlHack(request.Form("UserNumber"))
password = md5(request.Form("password"),16)
noname = NoSqlHack(request.Form("noname")) ''匿名 UserNumber=0
title = NoHtmlHackInput(ReplaceKeys(NoSqlHack(request.Form("title"))))
content = NoHtmlHackInput(ReplaceKeys(NoSqlHack(request.Form("content"))))
title = title
content = content
Action = NoHtmlHackInput(NoSqlHack(request.Form("Action")))
if Action="" then Call HTMLEnd("Action不能为空","back")
if stype="" then Call HTMLEnd("type不能为空","back")
if title="" then Call HTMLEnd("评论标题不能为空","back")
if content="" then Call HTMLEnd("评论内容不能为空","back")
if len(content)>1000 then Call HTMLEnd("评论内容超过1000字符。中文算两个字符。","back")
if not isnumeric(Id) then Call HTMLEnd("Id必须是数字","back")

if noname="" and noname<>"1" Then
	If UserNumber="" Or password="" Then
		If Session("FS_UserNumber")<>"" And session("FS_UserPassword") <> "" Then
			UserNumber=session("FS_UserNumber")
			password = session("FS_UserPassword")
		Else
			if UserNumber="" then Call HTMLEnd("用户名不能为空","back")
			if password="" then Call HTMLEnd("用户密码不能为空","back")
		End If
	Else
		If Session("FS_UserNumber")<>"" And session("FS_UserPassword") <> "" Then
			UserNumber=session("FS_UserNumber")
			password = session("FS_UserPassword")
		Else
			UserNumber = UserNumber
		End If
	End If
else
	UserNumber = "0"
end if
'------
set review_RS=User_Conn.execute("select top 1 ReviewTF,LimitReviewChar from FS_ME_SysPara")
if not review_RS.eof then : if not isnull(review_RS(0)) then : needAudited=cbool(review_RS(0)):LimitReviewChar=review_RS(1)
RsClose

If UserNumber<>"0" then
	Dim t_return,t_returnStr,CheckUserObj,t_returnPas
	If HaveDvbbs=1 or HaveOblog=1 Then
		t_return = Login(UserNumber,password,0)
	else
		Set CheckUserObj =  User_Conn.ExeCute("select UserNumber,UserPassWord from FS_ME_Users where UserName='"&NoSqlHack(UserNumber)&"' and UserPassWord='"&NoSqlHack(password)&"' or ( UserNumber='"&NoSqlHack(UserNumber)&"' and UserPassWord='"&NoSqlHack(password)&"')")
		If Not CheckUserObj.Eof Then
			t_returnStr = CheckUserObj(0)
			t_returnPas = CheckUserObj(1)
			t_return = True
		Else
			t_return = False
		End If
		CheckUserObj.Close : Set CheckUserObj = Nothing
	end if
	if t_return then
		If Session("FS_UserNumber") = "" Or session("FS_UserPassword") = "" THen
			Session("FS_UserNumber") = t_returnStr
			session("FS_UserPassword") = t_returnPas
		End if
		UserNumber = t_returnStr
	else
		Call HTMLEnd("用户名或密码错误，请重新输入。","back")
	end if
end if
select case ucase(stype)
	case "NS"
		ReviewTypes=0
		If CheckReviewTF(Id)=False Then Call HTMLEnd("该信息不允许评论","back")
	case "DS"
		ReviewTypes=1
		if not needAudited then
		'如果不需要审核则看该条下载是否需要
			set review_RS=Conn.execute("select ShowReviewTF,ReviewTF from FS_DS_List where ID = "&CintStr(Id))
			if not review_RS.eof then
				If review_RS("ReviewTF")<>1 Then
					Call HTMLEnd("该信息不允许评论","back")
				Else
					if not isnull(review_RS("ShowReviewTF")) then
						needAudited=cbool(review_RS("ShowReviewTF"))
					End If
				End If
			end if
			RsClose
		end if
	case "MS"
		ReviewTypes=2
	case "HS"
		ReviewTypes=3
	case "SD"
		ReviewTypes=4
	case "LOG"
		ReviewTypes=5
	case else
		Call HTMLEnd("Error:type("&stype&") is not found!","back")
end select
ReviewIP = NoSqlHack(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))
if ReviewIP="" then ReviewIP = NoSqlHack(Request.ServerVariables("REMOTE_ADDR"))
Call review_Data()
Fs_User.close
set Fs_User=Nothing
User_Conn.close
Conn.close



'/* Functions */
Sub HTMLEnd(Info,URL)
	URL=Request.ServerVariables("HTTP_REFERER")
	if URL="back" then
		response.Write("<script>alert('"&Info&"\n ');history.back();</script>")
		response.End()
	elseif URL<>"" then
		response.Write("<script>alert('"&Info&"\n ');location='"&URL&"';</script>")
		response.End()
	else
		response.Write(""&Info&"<br /> "&vbNewLine)
		response.End()
	end if
End Sub

Sub review_Data()
	Dim UserName
	review_Sql = "select UserNumber,InfoID,ReviewTypes,Title,Content,AddTime,ReviewIP,isLock,AdminLock,QuoteID from FS_ME_Review where ReviewID=0"
	set review_RS = CreateObject(G_FS_RS)
	review_RS.Open review_Sql,User_Conn,1,3
	if review_RS.eof then
		review_RS.addnew
		review_RS("ReviewTypes") = NoSqlHack(ReviewTypes)
		review_RS("InfoID") = NoSqlHack(Id)
		review_RS("UserNumber") = NoSqlHack(UserNumber)
		review_RS("Title") = NoSqlHack(title)
		review_RS("content") = NoSqlHack(content)
		review_RS("QuoteID") = 0
		review_RS("isLock") = 0
		''需要审核
		if needAudited then
			review_RS("AdminLock") = 1
		else
			review_RS("AdminLock") = 0
		end if
		review_RS("AddTime") = now
		review_RS("ReviewIP") = NoSqlHack(ReviewIP)
		review_RS.update
		RsClose:Set User_Conn = Nothing
		if needAudited then TmpStr = "我们审核通过后即可显示。"
		Call HTMLEnd("感谢您的评论。"&TmpStr,"back")
		Response.Write("<script language=""javascript"">")
		Response.Write("window.location.href='"&Request.ServerVariables("HTTP_REFERER")&"'")
		Response.Write("</script>")
	else

	end if
End Sub

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

Function CheckReviewTF(NewsID)
	Dim CheckRs,CheckSql,TempArr
	CheckSql="Select NewsProperty From FS_NS_News Where ID="&CintStr(NewsID)
	Set CheckRs=Server.CreateObject(G_FS_RS)
	CheckRs.Open CheckSql,Conn,1,1
	If CheckRs.Eof Then
		CheckReviewTF=False
	Else
		TempArr=Split(CheckRs("NewsProperty"),",")
		If TempArr(2)="1" Then
			CheckReviewTF=True
		Else
			CheckReviewTF=False
		End If
	End If
	CheckRs.Close
	Set CheckRs=Nothing
End Function

''过滤关键字
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

Sub RsClose()
	review_RS.Close
	Set review_RS = Nothing
end Sub
%>







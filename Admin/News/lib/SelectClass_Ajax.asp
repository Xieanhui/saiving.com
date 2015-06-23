<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp"-->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.Charset="GB2312"
Server.ScriptTimeOut=999999999
Dim Conn
Dim ParentId,Str_Sql,Rs_Class,Str_ClassInfo,allowMulitSelect
MF_Default_Conn
ParentId = NoSqlHack(Trim(Request.QueryString("ParentId")))
allowMulitSelect = NoSqlHack(Trim(Request.QueryString("Mulit")))
If ParentId = "" Then
	ParentId = "0"
End If
'On Error Resume Next
	Dim AndSQL,HavePopClassID
	HavePopClassID = GetClassIDOfPop("NS001")
	AndSQL = GetAndSQLOfSearchClass("NS001")
Str_Sql = "Select ClassID,ClassName,NewsTemplet,(Select Count(id) from FS_NS_NewsClass where ParentID=a.ClassID and ReycleTF=0 and isUrl=0) as HasSub from FS_NS_NewsClass a where ParentID='"&NoSqlHack(ParentId)&"' and ReycleTF=0 and isUrl=0 " & AndSQL & " order by OrderID desc,id desc"
Set Rs_Class = Conn.Execute(Str_Sql)
Str_ClassInfo=""
While Not Rs_Class.Eof
	Dim Str_NewsTemplet
	If Rs_Class("NewsTemplet") = "" Or IsNull(Rs_Class("NewsTemplet")) Then
		Str_NewsTemplet = "/Templets/NewsClass/news.htm"
	Else
		Str_NewsTemplet = Trim(Rs_Class("NewsTemplet"))
	End If	
	Dim imageHtml,titleHtml,checkboxHtml,childHtml
	If Rs_Class("HasSub")>0 Then
		imageHtml="<img src=""../images/+.gif"" alt=""点击展开子栏目"" width=""15"" height=""15"" border=""0"" class=""LableItem"" onclick=""javascript:SwitchImg(this,'"&Rs_Class("ClassID")&"');"" />"
		childHtml="<div id=""Parent"&Rs_Class("ClassID")&""" class=""SubItem"" HasSub=""True"" style=""display:none;""></div>"
	Else
		imageHtml="<img src=""../images/-.gif"" alt=""没有子栏目"" width=""15"" height=""15"" border=""0"" class=""LableItem"" />"
		childHtml=""
	End If
	If allowMulitSelect<>"" Then
		checkboxHtml="<input type=""checkbox"" id="""&Rs_Class("ClassID")&""" name=""chkNewsClasses"" value="""&Rs_Class("ClassName")&""" />"
	Else
		checkboxHtml=""
	End If
	If InStr(HavePopClassID,Rs_Class("Classid")) > 0 or Session("Admin_Is_Super")=1  Then
		titleHtml="<span id=""class_"&Rs_Class("ClassID")&""" ondblclick=""SubmitLable(this);"" class=""LableItem"" onclick=""SelectLable(this)"" value=""" & Str_NewsTemplet & """>"&Rs_Class("ClassName")&"</span>"
	Else
		titleHtml=Rs_Class("ClassName")
	End If
	Str_ClassInfo = Str_ClassInfo&"<div>"&imageHtml&checkboxHtml&titleHtml&childHtml&"</div>"&VbNewLine
	Rs_Class.MoveNext
Wend
If Err Then
	Str_ClassInfo="Fail|||"&ParentId&"|||"
Else
	Str_ClassInfo="Succee|||"&ParentId&"|||"&Str_ClassInfo
End If
Response.Write(Str_ClassInfo)
%>






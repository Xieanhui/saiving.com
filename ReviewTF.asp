<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="FS_Inc/Const.asp" -->
<!--#include file="FS_InterFace/MF_Function.asp" -->
<!--#include file="FS_Inc/Function.asp" -->
<!--#include file="FS_Inc/FS_Users_conformity.asp" -->
<%session.CodePage="936"%>
<%
response.Charset="gb2312"
''前台页面,由JS调用得到 调用该文件必须给定一些参数.
Dim Conn
MF_Default_Conn
Dim stype,Id,SpanId
Dim Server_Name,Server_V1,Server_V2,Cookie_Domain,TmpArr,FormReview,Str_Js,getType
Cookie_Domain = Get_MF_Domain()
Conn.close
Set Conn=Nothing
if Cookie_Domain="" then    
	Cookie_Domain = "localhost"
else
	if left(lcase(Cookie_Domain),len("http://"))="http://" then Cookie_Domain = mid(Cookie_Domain,len("http://")+1)
	if right(Cookie_Domain,1)="/" then Cookie_Domain = mid(Cookie_Domain,1,len(Cookie_Domain) - 1)
end if	
''防盗连
Dim Main_Name,Name_Str1,V_MainName,V_Str,Port
Server_Name = NoHtmlHackInput(NoSqlHack(LCase(Request.ServerVariables("SERVER_NAME"))))
Port=NoHtmlHackInput(NoSqlHack(LCase(Request.ServerVariables("SERVER_PORT"))))
If Port<>80 Then
	Server_Name = Server_Name&":"&Port
End If
IF Server_Name <> LCase(Split(Cookie_Domain,"/")(0)) Then
	Response.Write ("没有权限访问")
	Response.End
End If
Server_V1 = NoHtmlHackInput(NoSqlHack(Replace(Lcase(Cstr(Request.ServerVariables("HTTP_REFERER"))),"http://","")))
Server_V1 = Replace(Replace(Server_V1,"//","/"),"///","/")
IF Server_V1 = "" Then
	'Response.Write ("没有权限访问")
	'Response.End
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
	'Response.Write ("没有权限访问")
	'Response.End
End If

stype = NoHtmlHackInput(NoSqlHack(request.QueryString("type"))) 'NS
Id = CintStr(request.QueryString("Id")) 'Id
getType = NoHtmlHackInput(NoSqlHack(request.QueryString("getType"))) 'getType

SpanId = "Review_TF_"&Id
if stype="" then stype="NS"
if not isnumeric(Id) then  
	FormReview = "ID必须是数字。"
else
	FormReview = "<span style=""width:100px; height:25px; line-height:25px; text-align:left;""><a href=""http://"&Cookie_Domain&"/ShowReviewList.asp?type="&stype&"&Id="&Id&""" target=""_blank"">点击查看所有评论</a></span><br>"
	FormReview=FormReview & "<form action=""http://"&Cookie_Domain&"/ReviewUrl.asp"" name=""reviewform"" method=""post""  style=""margin:0px;"">"
	if session("FS_UserNumber")= "" Or session("FS_UserPassword") = "" And request.Cookies(Forum_sn)("username")="" And request.Cookies(ObCookies_name)("username")="" then 
		FormReview=FormReview&"用户名<input name=""UserNumber"" type=""text"" class=""f-text"" id=""UserNumber"" size=""15"" />"
		FormReview=FormReview&"密码<input name=""password"" type=""password"" class=""f-text"" id=""password"" size=""12""/>"
		FormReview=FormReview&"匿名<input name=""noname"" type=""checkbox"" id=""noname"" value=""1"" onClick=""if(this.checked==true){UserNumber.disabled=true;password.disabled=true;}else{UserNumber.disabled=false;password.disabled=false;};""/><br />"
		FormReview=FormReview&"标　题<input name=""title"" type=""text"" class=""f-text"" id=""title"" size=""40""/><br />"
	else
		FormReview=FormReview&"标　题<input name=""title"" type=""text"" class=""f-text"" id=""title"" size=""36""/>"
		FormReview=FormReview&"&nbsp;匿名<input name=""noname"" type=""checkbox"" class=""f-text"" id=""noname"" value=""1""/><br />"
	end if
	FormReview=FormReview&"<textarea name=""content"" class=""f-text"" cols=""50"" rows=""5""/></textarea><input type=""hidden"" name=""Id"" value="""&Id&"""/><input type=""hidden"" name=""type"" value="""&ucase(stype)&"""/><input type=""hidden"" name=""Action"" value=""add_save""/><br />"
	FormReview=FormReview&"<input type=""submit"" name=""Submit"" class=""f-button"" value=""发表评论""/>&nbsp;&nbsp;<input type=""reset""  class=""f-reset"" name=""Submit2"" value=""重新填写""/>"
	FormReview=FormReview&"</form>"
end if
if getType = "2" then 
	Response.write FormReview	
else
	Response.write "if($('"&SpanId&"')!=null) {$('"&SpanId&"').innerHTML='"&FormReview&"';} else {alert('"&SpanId&" Is Not Fonud');}"	
end if	
%>






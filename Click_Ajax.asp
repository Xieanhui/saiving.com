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
Dim Conn,User_Conn,Click_Sql,Click_RS,strShowErr,Cookie_Domain
Dim Server_Name,Server_V1,Server_V2
Dim TmpStr,TmpArr
Dim stype,SubSys,spanid,WriteID
TmpStr = ""

MF_Default_Conn   
MF_User_Conn  

stype = NoHtmlHackInput(NoSqlHack(request.QueryString("type")))
SubSys = NoHtmlHackInput(NoSqlHack(request.QueryString("SubSys")))
spanid = NoHtmlHackInput(NoSqlHack(request.QueryString("spanid")))
WriteID = spanid
if stype="" then stype="js"
if SubSys="" then TmpStr = "Error:SubSys is null!"
if spanid="" then TmpStr = "Error:spanid is null!"
TmpArr = split(spanid,"_")
if ubound(TmpArr)<3 then TmpStr = "Error:spanid's _ is Err!"
spanid = TmpArr(3)

If TmpStr="" Then
	select case SubSys
		case "NS"
			Conn.execute("Update FS_NS_News set Hits=1 where Hits is null")
			Conn.execute("Update FS_NS_News set Hits=Hits+1 where NewsID='"&NoSqlHack(spanid)&"'")
			set Click_RS=Conn.execute("select Hits from FS_NS_News where NewsID='"&NoSqlHack(spanid)&"'")
			if not Click_RS.eof then TmpStr = cstr(Click_RS(0))
			RsClose()
		case "DS"
			if NoSqlHack(request.QueryString("Get"))="ClickNum" then 
				set Click_RS=Conn.execute("select ClickNum from FS_DS_List where ID="&CintStr(spanid)&"")
				if not Click_RS.eof then TmpStr = cstr(Click_RS(0))
				RsClose()
			else			
				Conn.execute("Update FS_DS_List set Hits=1 where Hits is null")
				Conn.execute("Update FS_DS_List set Hits=Hits+1 where ID="&CintStr(spanid)&"")
				set Click_RS=Conn.execute("select Hits from FS_DS_List where ID="&CintStr(spanid)&"")
				if not Click_RS.eof then TmpStr = cstr(Click_RS(0))
				RsClose()
			end if
		case "MS"
			Conn.execute("Update FS_MS_Products set Click=1 where Click is null")
			Conn.execute("Update FS_MS_Products set Click=Click+1 where ID="&CintStr(spanid)&"")
			set Click_RS=Conn.execute("select Click from FS_MS_Products where ID="&CintStr(spanid)&"")
			if not Click_RS.eof then TmpStr = cstr(Click_RS(0))
			RsClose()
		case "SD"
			Conn.execute("Update FS_SD_News set Hits=1 where Hits is null")
			Conn.execute("Update FS_SD_News set Hits=Hits+1 where ID="&CintStr(spanid)&"")
			set Click_RS=Conn.execute("select Hits from FS_SD_News where ID="&CintStr(spanid)&"")
			if not Click_RS.eof then TmpStr = cstr(Click_RS(0))
			RsClose()
		case "HS" ''房产有三个表，还需要一个参数
			spanid=Split(spanid,"$")
			Select Case spanid(0)
				Case "QU"
					Conn.execute("Update FS_HS_Quotation set Click=1 where Click is null")
					Conn.execute("Update FS_HS_Quotation set Click=Click+1 where ID="&CintStr(spanid(1))&"")
					set Click_RS=Conn.execute("select Click from FS_HS_Quotation where ID="&CintStr(spanid(1))&"")
					if not Click_RS.eof then TmpStr = cstr(Click_RS(0))
				Case "SE"
					Conn.execute("Update FS_HS_Second set Click=1 where Click is null")
					Conn.execute("Update FS_HS_Second set Click=Click+1 where SID="&CintStr(spanid(1))&"")
					set Click_RS=Conn.execute("select Click from FS_HS_Second where SID="&CintStr(spanid(1))&"")
					if not Click_RS.eof then TmpStr = cstr(Click_RS(0))
				Case "TE"
					Conn.execute("Update FS_HS_Tenancy set Click=1 where Click is null")
					Conn.execute("Update FS_HS_Tenancy set Click=Click+1 where TID="&CintStr(spanid(1))&"")
					set Click_RS=Conn.execute("select Click from FS_HS_Tenancy where TID="&CintStr(spanid(1))&"")
					if not Click_RS.eof then TmpStr = cstr(Click_RS(0))
			End Select
			RsClose()
		case "Log" '日志
			User_Conn.execute("Update FS_ME_Infoilog set Hits=1 where Hits is null")
			User_Conn.execute("Update FS_ME_Infoilog set Hits=Hits+1 where iLogID="&CintStr(spanid)&"")
			set Click_RS=User_Conn.execute("select Hits from FS_ME_Infoilog where iLogID="&CintStr(spanid)&"")
			if not Click_RS.eof then TmpStr = cstr(Click_RS(0))
			RsClose()
		case "PHOTO" '相册
			User_Conn.execute("Update FS_ME_Photo set Hits=1 where Hits is null")
			User_Conn.execute("Update FS_ME_Photo set Hits=Hits+1 where ID="&CintStr(spanid)&"")
			set Click_RS=User_Conn.execute("select Hits from FS_ME_Photo where  ID="&CintStr(spanid)&"")
			if not Click_RS.eof then TmpStr = cstr(Click_RS(0))
			RsClose()
		case "AP" '求职者
			Conn.execute("Update FS_AP_Resume_BaseInfo set click=1 where click is null")
			Conn.execute("Update FS_AP_Resume_BaseInfo set click=click+1 where BID="&CintStr(spanid)&"")
			set Click_RS=Conn.execute("select click from FS_AP_Resume_BaseInfo where  BID="&CintStr(spanid)&"")
			if not Click_RS.eof then TmpStr = cstr(Click_RS(0))
			RsClose()
		case else
			response.Write("Error:"&SubSys&" is not found.")
	End Select
End If
If stype="js" Then
	TmpStr ="try"&VbNewLine&_
"{"&VbNewLine&_
	"document.getElementById("""&WriteID&""").innerHTML="""&TmpStr&""";"&VbNewLine&_
"}"&VbNewLine&_
"catch (e)"&VbNewLine&_
"{"&VbNewLine&_
	"document.writeln(""Load Error"");"&VbNewLine&_
"}"&VbNewLine
End If
Response.Write(TmpStr)
ConnClose()
Sub RsClose()
	Click_RS.Close
	Set Click_RS = Nothing
end Sub

Sub ConnClose()
	Set User_Conn = Nothing
	Set Conn = Nothing
	response.End()
End Sub
%>
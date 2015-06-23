<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Dim User_Conn,AwardID,PrizeID,AwardsUserRs,UserInfoRs,AwardsUserArray,ArrayIndex,Result
MF_User_Conn
MF_Session_TF
AwardID=NoSqlHack(Request.QueryString("AwardID"))
PrizeID=NoSqlHack(Request.QueryString("PrizeID"))
Result="PrizeUsers_"&AwardID&"*"&chr(10)&chr(13)
Result=Result&"<select name='AwardUsers_"&AwardID&"'>"&chr(10)&chr(13)
if isnumeric(PrizeID) then
Set AwardsUserRs=User_Conn.execute("Select UserNumber,winner From FS_ME_User_Prize where PrizeID="&CintStr(PrizeID)&" And awardID="&CintStr(AwardID)&" and winner=1")
if not AwardsUserRs.eof then
	while not AwardsUserRs.eof  
		Set UserInfoRs=User_Conn.execute("Select UserName from FS_ME_Users where UserNumber='"&AwardsUserRs("UserNumber")&"'")
		Result=Result&"<option value='"&AwardsUserRs("UserNumber")&"'>"&UserInfoRs("UserName")&"</option>"&Chr(10)&Chr(13)
		AwardsUserRs.movenext
	Wend
ELse
	Result=Result&"<option value='-1'>ÔÝÎÞÖÐ½±</option>"&Chr(10)&Chr(13)
End if
AwardsUserRs.close
Set AwardsUserRs=nothing
Set UserInfoRs=nothing
end if
Result=Result&"</select>"
Response.Charset="GB2312"
Response.Write(Result)
User_Conn.close
Set User_Conn=nothing
%>






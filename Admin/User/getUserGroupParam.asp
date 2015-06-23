<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Dim Conn,result,User_Conn,GroupID,GroupParamRs
MF_Default_Conn
MF_User_Conn
MF_Session_TF
GroupID=Request.QueryString("id")
if isNumeric(GroupID) then 
	Set GroupParamRs=Server.CreateObject(G_FS_RS)
	GroupParamRs.open "select GroupName,GroupName,UpfileNum,UpfileSize,GroupDate,GroupPoint,GroupMoney,ProductDiscount,GroupType,LimitInfoNum,CorpTemplet,GroupDebateNum,JuniorDomain,KeywordsNumber,isHtml,BcardNumber,Templetwatermark from FS_ME_Group where GroupID="&CintStr(GroupID),User_Conn,1,3
	if not GroupParamRs.eof then
		result=GroupParamRs("GroupName")&"|"&GroupParamRs("UpfileNum")&"|"&GroupParamRs("UpfileSize")&"|"&GroupParamRs("GroupDate")&"|"&GroupParamRs("GroupPoint")&"|"&GroupParamRs("GroupMoney")&"|"&GroupParamRs("GroupType")&"|"&GroupParamRs("LimitInfoNum")&"|"&GroupParamRs("CorpTemplet")&"|"&GroupParamRs("GroupDebateNum")&"|"&GroupParamRs("JuniorDomain")&"|"&GroupParamRs("KeywordsNumber")&"|"&GroupParamRs("isHtml")&"|"&GroupParamRs("BcardNumber")&"|"&GroupParamRs("Templetwatermark")&"|"&GroupParamRs("ProductDiscount")
	end if
	Response.Charset="GB2312"
	Response.Write(result)
	GroupParamRs.close
	Set GroupParamRs=nothing
	User_Conn.close
	Set User_Conn=nothing
end if

%>






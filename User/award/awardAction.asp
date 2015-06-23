<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="lib/cls_award.asp"-->
<%
Response.Charset="GB2312"
Dim prizeID,awardID,action,Rs,joinNumber,Integral,PrizeRs,answerID,questionID,rightAnswerID
action=request.QueryString("action")
awardID=CintStr(request.QueryString("awardID"))
prizeID=CintStr(request.QueryString("prizeID"))
Integral=NoSqlHack(request.QueryString("Integral"))
if action="join" then
	User_Conn.execute("Insert into FS_ME_User_Prize (prizeid,usernumber,awardID) values("&CintStr(prizeID)&",'"&session("FS_UserNumber")&"',"&CintStr(awardID)&")")
	'获得当前参加人数--------------------------------
	User_Conn.execute("Update FS_ME_Users set Integral=(Integral-"&Integral&") where usernumber='"&session("FS_UserNumber")&"'")
	Response.Write("成功，请关注抽奖结果！")
	Call Fs_User.AddLog("积分抽奖",Fs_User.UserNumber,Integral,"0","消耗积分",1)		
	Set Rs=nothing
elseif action="change" then
	Set Rs=User_Conn.execute("Select count(ID) From FS_ME_User_Prize where prizeid="&CintStr(prizeID)&" And usernumber='"&session("FS_UserNumber")&"'")
	Set PrizeRs=User_Conn.execute("Select perUserNum from FS_ME_Prize where prizeid="&CintStr(prizeID))
	if not Rs.eof then
		if Clng(Rs(0))>Clng(PrizeRs("perUserNum")) or  Clng(Rs(0))=Clng(PrizeRs("perUserNum")) then
			Response.Write("每人只能兑换"&PrizeRs("perUserNum")&"次！")
		Else
			User_Conn.execute("Insert into FS_ME_User_Prize (prizeid,usernumber,winner) values("&CintStr(prizeID)&",'"&session("FS_UserNumber")&"',1)")
			User_Conn.execute("Update FS_ME_Users set Integral=(Integral-"&Integral&") where usernumber='"&session("FS_UserNumber")&"'")
			Response.Write("兑换成功")
			Call Fs_User.AddLog("积分兑换",Fs_User.UserNumber,Integral,"0","消耗积分",1)		
		End if
	Else
		User_Conn.execute("Insert into FS_ME_User_Prize (prizeid,usernumber,winner) values("&CintStr(prizeID)&",'"&session("FS_UserNumber")&"',1)")
		User_Conn.execute("Update FS_ME_Users set Integral=(Integral-"&Integral&") where usernumber='"&session("FS_UserNumber")&"'")
		Call Fs_User.AddLog("积分兑换",Fs_User.UserNumber,Integral,"0","消耗积分",1)		
		Response.Write("兑换成功")
	End if
	Rs.close
	PrizeRs.close
	Set Rs=nothing
	Set PrizeRs=nothing
Elseif action="answer" then
	questionID=NoSqlHack(request.QueryString("questionID"))
	answerID=CintStr(request.QueryString("answerID"))
	Set Rs=User_Conn.execute("Select NeedPoint,PrizePoint,RightAnswerID From FS_ME_AnswerForPoint where AID="&answerID)
	if not Rs.eof then
		if Clng(questionID)=Clng(Rs("RightAnswerID")) then
			User_Conn.execute("Insert into FS_ME_Answer_User (questionID,usernumber) values("&CintStr(answerID)&",'"&session("FS_UserNumber")&"')")
			User_Conn.execute("Update FS_ME_Users set Integral=(Integral+("&Rs("PrizePoint")&"-"&Rs("NeedPoint")&")) where usernumber='"&session("FS_UserNumber")&"'")
			Response.Write("恭喜！回答正确，你将获得积分："&Rs("PrizePoint"))
			Call Fs_User.AddLog("积分问答",Fs_User.UserNumber,Rs("NeedPoint"),"0","消耗积分",1)		
			Call Fs_User.AddLog("积分问答",Fs_User.UserNumber,Rs("PrizePoint"),"0","获得积分",0)		
		Else
			User_Conn.execute("Update FS_ME_Users set Integral=(Integral-"&Rs("NeedPoint")&") where usernumber='"&session("FS_UserNumber")&"'")
			Response.Write("回答错误！")
			Call Fs_User.AddLog("积分问答",Fs_User.UserNumber,Rs("NeedPoint"),"0","消耗积分",1)		
		End if
	Else
		Response.Write("系统问题，请与管理员联系，积分将不被扣除！")
	End if
Elseif action="menu" then
		Dim menuAwardObj
		Set menuAwardObj=new Cls_Award
		if menuAwardObj.activeAward then
			Response.write("<img src="""&s_savepath&"/images/active.gif"" border=""0""/>")
		End if
End if
Set Conn=nothing
Set User_Conn=nothing
Set Fs_User = Nothing
%>






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
	'��õ�ǰ�μ�����--------------------------------
	User_Conn.execute("Update FS_ME_Users set Integral=(Integral-"&Integral&") where usernumber='"&session("FS_UserNumber")&"'")
	Response.Write("�ɹ������ע�齱�����")
	Call Fs_User.AddLog("���ֳ齱",Fs_User.UserNumber,Integral,"0","���Ļ���",1)		
	Set Rs=nothing
elseif action="change" then
	Set Rs=User_Conn.execute("Select count(ID) From FS_ME_User_Prize where prizeid="&CintStr(prizeID)&" And usernumber='"&session("FS_UserNumber")&"'")
	Set PrizeRs=User_Conn.execute("Select perUserNum from FS_ME_Prize where prizeid="&CintStr(prizeID))
	if not Rs.eof then
		if Clng(Rs(0))>Clng(PrizeRs("perUserNum")) or  Clng(Rs(0))=Clng(PrizeRs("perUserNum")) then
			Response.Write("ÿ��ֻ�ܶһ�"&PrizeRs("perUserNum")&"�Σ�")
		Else
			User_Conn.execute("Insert into FS_ME_User_Prize (prizeid,usernumber,winner) values("&CintStr(prizeID)&",'"&session("FS_UserNumber")&"',1)")
			User_Conn.execute("Update FS_ME_Users set Integral=(Integral-"&Integral&") where usernumber='"&session("FS_UserNumber")&"'")
			Response.Write("�һ��ɹ�")
			Call Fs_User.AddLog("���ֶһ�",Fs_User.UserNumber,Integral,"0","���Ļ���",1)		
		End if
	Else
		User_Conn.execute("Insert into FS_ME_User_Prize (prizeid,usernumber,winner) values("&CintStr(prizeID)&",'"&session("FS_UserNumber")&"',1)")
		User_Conn.execute("Update FS_ME_Users set Integral=(Integral-"&Integral&") where usernumber='"&session("FS_UserNumber")&"'")
		Call Fs_User.AddLog("���ֶһ�",Fs_User.UserNumber,Integral,"0","���Ļ���",1)		
		Response.Write("�һ��ɹ�")
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
			Response.Write("��ϲ���ش���ȷ���㽫��û��֣�"&Rs("PrizePoint"))
			Call Fs_User.AddLog("�����ʴ�",Fs_User.UserNumber,Rs("NeedPoint"),"0","���Ļ���",1)		
			Call Fs_User.AddLog("�����ʴ�",Fs_User.UserNumber,Rs("PrizePoint"),"0","��û���",0)		
		Else
			User_Conn.execute("Update FS_ME_Users set Integral=(Integral-"&Rs("NeedPoint")&") where usernumber='"&session("FS_UserNumber")&"'")
			Response.Write("�ش����")
			Call Fs_User.AddLog("�����ʴ�",Fs_User.UserNumber,Rs("NeedPoint"),"0","���Ļ���",1)		
		End if
	Else
		Response.Write("ϵͳ���⣬�������Ա��ϵ�����ֽ������۳���")
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






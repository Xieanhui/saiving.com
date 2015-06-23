<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Dim Conn,User_Conn,awardRs,prizeRs,AwardName,AwardPic,StartDate,EndDate,PrizeGradeNum,PrizeIDS,PrizeNames,ForIndex,PrizeName,PrizeGrade,PrizePic,PrizeNum,CurrentPrizeIDS,answerNum,AnswerDesc,RightAnswer,CurrentAnswerID,CurrentAnswerIDS,prizeNeedPoint,usernumber
Dim prizeID
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("ME_award") then Err_Show 
if request.QueryString("act")="addaction" then
	Set awardRs=Server.CreateObject(G_FS_RS)
	Set prizeRs=Server.CreateObject(G_FS_RS)
	AwardName=NoSqlHack(Request.Form("AwardName"))
	AwardPic=NoSqlHack(Request.Form("AwardPic"))
	StartDate=NoSqlHack(Request.Form("StartDate"))
	EndDate=NoSqlHack(Request.Form("EndDate"))
	PrizeGradeNum=NoSqlHack(Request.Form("PrizeGradeNum"))
	'添加奖品
	for ForIndex=0 to PrizeGradeNum-1
		PrizeName=NoSqlHack(Request.Form("Prize_"&(ForIndex+1)&"_name"))
		PrizeGrade=ForIndex+1
		prizeNeedPoint=NoSqlHack(Request.Form("NeedPoint_"&(ForIndex+1)))
		PrizePic=NoSqlHack(Request.Form("prize_"&(ForIndex+1)&"_pic"))
		PrizeNum=NoSqlHack(Request.Form("prize_"&(ForIndex+1)&"_number"))
		if PrizeName="" or PrizeNum="" or not isnumeric(PrizeNum) then
			Response.Redirect("../error.asp?ErrCodes=<li>奖品名称可能为空</li><li>奖品数量可能为空</li><li>奖品数量中可能包含有字符</li>")
			Response.End()
		end if
		Set prizeRs=User_Conn.Execute("insert into FS_ME_Prize (PrizeName,NeedPoint,PrizeGrade,PrizePic,PrizeNum) values('"&PrizeName&"','"&prizeNeedPoint&"','"&NoSqlHack(PrizeGrade)&"','"&PrizePic&"','"&PrizeNum&"')")
		CurrentPrizeIDS=CurrentPrizeIDS&","&User_Conn.Execute("Select Max(PrizeID) from FS_ME_Prize")(0)
	next
	CurrentPrizeIDS=DelHeadAndEndDot(CurrentPrizeIDS)
	awardRs.open "select AwardName,AwardPic,StartDate,EndDate,PrizeIDS,opened from FS_ME_award",User_Conn,1,3
	awardRs.addNew
	awardRs("AwardName")=AwardName
	awardRs("AwardPic")=AwardPic
	awardRs("StartDate")=StartDate
	awardRs("EndDate")=EndDate
	awardRs("PrizeIDS")=CurrentPrizeIDS
	awardRs("opened")=0
	awardRs.update
	awardRs.close
	if err.number=0 then
		Response.Redirect("../success.asp?ErrCodes=<li>操作成功</li>&ErrorURL=user/award.asp")
		Response.End()
	end if
elseif request.QueryString("act")="delete" then
	if not MF_Check_Pop_TF("ME029") then Err_Show 
	User_Conn.Execute("Delete From FS_ME_Award where awardid in ("&FormatIntArr(request.Form("DeleteAwards"))&")")
	if err.number=0 then
		Response.Redirect("../success.asp?ErrCodes=<li>操作成功</li>&ErrorURL=user/award.asp")
		Response.End()
	end if
elseif Request.QueryString("act")="editaction" then
	Set awardRs=Server.CreateObject(G_FS_RS)
	Set prizeRs=Server.CreateObject(G_FS_RS)
	AwardName=NoSqlHack(Request.Form("AwardName"))
	AwardPic=NoSqlHack(Request.Form("AwardPic"))
	StartDate=NoSqlHack(Request.Form("StartDate"))
	EndDate=NoSqlHack(Request.Form("EndDate"))
	PrizeGradeNum=NoSqlHack(Request.Form("PrizeGradeNum"))
	'添加奖品
	for ForIndex=0 to PrizeGradeNum-1
		PrizeName=NoSqlHack(Request.Form("Prize_"&(ForIndex+1)&"_name"))
		PrizeGrade=ForIndex+1
		prizeNeedPoint=NoSqlHack(Request.Form("NeedPoint_"&(ForIndex+1)))
		PrizePic=NoSqlHack(Request.Form("prize_"&(ForIndex+1)&"_pic"))
		PrizeNum=NoSqlHack(Request.Form("prize_"&(ForIndex+1)&"_number"))
		if PrizeName="" or PrizeNum="" or not isnumeric(PrizeNum) then
			Response.Redirect("../error.asp?ErrCodes=<li>奖品名称可能为空</li><li>奖品数量可能为空</li><li>奖品数量中可能包含有字符</li>")
			Response.End()
		end if
		Set prizeRs=User_Conn.Execute("insert into FS_ME_Prize (PrizeName,NeedPoint,PrizeGrade,PrizePic,PrizeNum) values('"&PrizeName&"',"&prizeNeedPoint&","&PrizeGrade&",'"&PrizePic&"',"&PrizeNum&")")
		CurrentPrizeIDS=CurrentPrizeIDS&","&User_Conn.Execute("Select Max(PrizeID) from FS_ME_Prize")(0)
	next
	CurrentPrizeIDS=DelHeadAndEndDot(CurrentPrizeIDS)
	Response.Write("select AwardName,AwardPic,StartDate,EndDate,PrizeIDS,opened from FS_ME_award where awardid="&CintStr(Request.QueryString("awardid")))
	awardRs.open "select AwardName,AwardPic,StartDate,EndDate,PrizeIDS,opened from FS_ME_award where awardid="&CintStr(Request.QueryString("awardid")),User_Conn,1,3
	awardRs("AwardName")=AwardName
	awardRs("AwardPic")=AwardPic
	awardRs("StartDate")=StartDate
	awardRs("EndDate")=EndDate
	Response.Write("Delete From FS_ME_Prize where prizeid in("&awardRs("PrizeIDS")&")")
	User_Conn.execute("Delete From FS_ME_Prize where prizeid in("&awardRs("PrizeIDS")&")")
	awardRs("PrizeIDS")=CurrentPrizeIDS
	awardRs("opened")=0
	awardRs.update
	awardRs.close
	if err.number=0 then
		Response.Redirect("../success.asp?ErrCodes=<li>操作成功</li>&ErrorURL=user/award.asp")
		Response.End()
	end if
elseif Request.QueryString("Act")="editPrizeaction" then
	Set prizeRs=Server.CreateObject(G_FS_RS)
	prizeRS.open  "select prizeID,PrizeName,prizeDesc,PrizePic,NeedPoint,storage,StartDate,EndDate,provider,perUserNum from FS_ME_Prize where prizeID="&NoSqlHack(request.QueryString("prizeid")),User_Conn,1,3
	prizeRS("PrizeName")=NoSqlHack(Request.Form("PrizeName"))
	prizeRS("prizeDesc")=NoSqlHack(Request.Form("prizeDesc"))
	prizeRS("NeedPoint")=NoSqlHack(Request.Form("NeedPoint"))
	prizeRS("storage")=NoSqlHack(Request.Form("storage"))
	prizeRS("StartDate")=NoSqlHack(Request.Form("StartDate"))
	prizeRS("EndDate")=NoSqlHack(Request.Form("EndDate"))	
	prizeRS("provider")=NoSqlHack(Request.Form("provider"))	
	prizeRS("perUserNum")=NoSqlHack(Request.Form("perUserNum"))
	prizeRs("PrizePic")=NoSqlHack(request.Form("PrizePic"))
	prizeRs.update
	prizeRs.close
	if err.number=0 then
		Response.Redirect("../success.asp?ErrCodes=<li>操作成功</li>&ErrorURL=user/ChangePrize.asp")
		Response.End()
	end if
elseif Request.QueryString("Act")="addPrizeaction" then
	Set prizeRs=Server.CreateObject(G_FS_RS)
	prizeRS.open  "select prizeID,PrizeName,prizeDesc,PrizePic,NeedPoint,storage,StartDate,EndDate,provider,perUserNum,isChange from FS_ME_Prize ",User_Conn,1,3
	prizeRs.addnew
	prizeRS("PrizeName")=NoSqlHack(Request.Form("PrizeName"))
	prizeRS("prizeDesc")=NoSqlHack(Request.Form("prizeDesc"))
	prizeRS("NeedPoint")=NoSqlHack(Request.Form("NeedPoint"))
	prizeRS("storage")=NoSqlHack(Request.Form("storage"))
	prizeRS("StartDate")=NoSqlHack(Request.Form("StartDate"))
	prizeRS("EndDate")=NoSqlHack(Request.Form("EndDate"))	
	prizeRS("provider")=NoSqlHack(Request.Form("provider"))	
	prizeRS("perUserNum")=NoSqlHack(Request.Form("perUserNum"))
	prizeRS("isChange")=1
	prizeRs("PrizePic")=NoSqlHack(request.Form("PrizePic"))	
	prizeRs.update
	prizeRs.close
	if err.number=0 then
		Response.Redirect("../success.asp?ErrCodes=<li>操作成功</li>&ErrorURL=user/ChangePrize.asp")
		Response.End()
	end if
elseif Request.QueryString("Act")="deletePrizeaction" then
	User_Conn.Execute("Update FS_ME_Prize set isChange=0 where PrizeID in ("&FormatIntArr(Request.Form("DeleteChangePrize"))&")")
	if err.number=0 then
		Response.Redirect("../success.asp?ErrCodes=<li>操作成功</li>&ErrorURL=user/ChangePrize.asp")
		Response.End()
	end if
elseif Request("Act")="editAFPointaction" then
	Set awardRs=Server.CreateObject(G_FS_RS)
	awardRs.open "Select ATopic,needPoint,PrizePoint,APic,ADesc,AStartDate,AEndDate,AnswerIDS,RightAnswerID from FS_ME_AnswerForPoint where Aid="&NoSqlHack(Request.QueryString("AID")),User_Conn,1,3
	awardRs("ATopic")=NoSqlHack(Request.Form("ATopic"))
	awardRs("needPoint")=NoSqlHack(Request.Form("needPoint"))
	awardRs("PrizePoint")=NoSqlHack(Request.Form("PrizePoint"))
	awardRs("APic")=NoSqlHack(Request.Form("APic"))
	awardRs("ADesc")=NoSqlHack(Request.Form("ADesc"))
	awardRs("AStartDate")=NoSqlHack(Request.Form("StartDate"))
	awardRs("AEndDate")=NoSqlHack(Request.Form("EndDate"))
	answerNum=NoSqlHack(Request.Form("AnswerNum"))
	RightAnswer=NoSqlHack(Request.Form("rightAnswer"))
	for ForIndex=0 to answerNum-1
		AnswerDesc=NoSqlHack(Request.Form("Answer_"&(ForIndex+1)))
		if AnswerDesc="" then
			Response.Redirect("../error.asp?ErrCodes=<li>答案内容为空</li>")
			Response.End()
		end if
		User_Conn.Execute("Insert into FS_ME_Answer (AnswerDesc) values('"&NoSqlHack(AnswerDesc)&"')")
		CurrentAnswerID=User_Conn.execute("select Max(answerid) from FS_ME_Answer")(0)
		if (ForIndex+1)=Cint(RightAnswer) then
			awardRs("RightAnswerID")=CurrentAnswerID
		end if
		CurrentAnswerIDS=CurrentAnswerIDS&","&CurrentAnswerID
	next
	User_Conn.execute("Delete From FS_ME_Answer where AnswerID in ("&awardRs("AnswerIDS")&")")
	awardRs("AnswerIDS")=DelHeadAndEndDot(CurrentAnswerIDS)
	awardRs.update
	awardRs.close
	if err.number=0 then
		Response.Redirect("../success.asp?ErrCodes=<li>操作成功!</li>&ErrorURL=user/AnswerForPoint.asp")
		Response.End()
	end if
elseif Request.QueryString("Act")="addAFPointaction" then
	Set awardRs=Server.CreateObject(G_FS_RS)
	awardRs.open "Select ATopic,needPoint,PrizePoint,APic,ADesc,AStartDate,AEndDate,AnswerIDS,RightAnswerID from FS_ME_AnswerForPoint",User_Conn,1,3
	awardRs.addNew
	awardRs("ATopic")=NoSqlHack(Request.Form("ATopic"))
	awardRs("needPoint")=NoSqlHack(Request.Form("needPoint"))
	awardRs("PrizePoint")=NoSqlHack(Request.Form("PrizePoint"))
	awardRs("APic")=NoSqlHack(Request.Form("APic"))
	awardRs("ADesc")=NoSqlHack(Request.Form("ADesc"))
	awardRs("AStartDate")=NoSqlHack(Request.Form("StartDate"))
	awardRs("AEndDate")=NoSqlHack(Request.Form("EndDate"))
	answerNum=NoSqlHack(Request.Form("AnswerNum"))
	RightAnswer=NoSqlHack(Request.Form("rightAnswer"))
	for ForIndex=0 to answerNum-1
		AnswerDesc=NoSqlHack(Request.Form("Answer_"&(ForIndex+1)))
		if AnswerDesc="" then
			Response.Redirect("../error.asp?ErrCodes=<li>答案内容为空</li>")
			Response.End()
		end if
		User_Conn.Execute("Insert into FS_ME_Answer (AnswerDesc) values('"&AnswerDesc&"')")
		CurrentAnswerID=User_Conn.execute("select Max(answerid) from FS_ME_Answer")(0)
		if (ForIndex+1)=Cint(RightAnswer) then
			awardRs("RightAnswerID")=CurrentAnswerID
		end if
		CurrentAnswerIDS=CurrentAnswerIDS&","&CurrentAnswerID
	next
	awardRs("AnswerIDS")=DelHeadAndEndDot(CurrentAnswerIDS)
	Response.Write(DelHeadAndEndDot(CurrentAnswerIDS))
	awardRs.update
	awardRs.close
	if err.number=0 then
		Response.Redirect("../success.asp?ErrCodes=<li>操作成功</li>&ErrorURL=user/AnswerForPoint.asp")
		Response.End()
	end if
elseif Request.QueryString("Act")="deleteAFPointaction" then
	CurrentAnswerIDS=User_Conn.execute("Select AnswerIDS From FS_ME_AnswerForPoint where AID in ("&FormatIntArr(Request("DeleteAFPoint"))&")")(0)
	User_Conn.execute("Delete From FS_ME_Answer where answerid in ("&FormatIntArr(CurrentAnswerIDS)&")")
	User_Conn.execute("Delete From FS_ME_AnswerForPoint where AID in ("&FormatIntArr(Request("DeleteAFPoint"))&")")
	if err.number=0 then
		Response.Redirect("../success.asp?ErrCodes=<li>操作成功</li>&ErrorURL=user/AnswerForPoint.asp")
		Response.End()
	end if
elseif request.QueryString("Act")="open" then
	if not MF_Check_Pop_TF("ME030") then Err_Show 
	Dim awardID,resultRs,prizeNumber
	awardID=NoSqlHack(request.QueryString("awardiD"))
	Set awardRs=User_Conn.execute("Select prizeID from FS_ME_User_Prize where awardID="&NoSqlHack(awardID))
	while not awardRs.eof
		Set prizeRs=User_Conn.execute("Select PrizeNum from FS_ME_Prize where prizeID="&NoSqlHack(awardRs("prizeID")))
		if not PrizeRs.eof then
			prizeNumber=prizeRs("PrizeNum")
			if G_IS_SQL_DB=0 then
				Randomize
				Response.Write(awardRs("prizeID"))
				Set resultRs=User_Conn.execute("Select top "&prizeNumber&" id from FS_ME_User_Prize order by Rnd(-(ID+"&Rnd()&"))")
			Else
				Set resultRs=User_Conn.execute("Select top "&prizeNumber&" id from FS_ME_User_Prize order BY NEWID()")
			End if
			while not resultRs.eof
				User_Conn.execute("Update FS_ME_User_Prize set winner=1 where id="&CintStr(resultRs("id")))
				resultRs.movenext
			wend
		ENd if
		awardRs.movenext
	wend
	awardRs.close
	User_Conn.execute("Update FS_ME_Award set opened=1 where awardid="&NoSqlHack(awardid))
	Set awardRs=nothing
	Set resultRs=nothing
	if err.number=0 then
		Response.Redirect("../success.asp?ErrCodes=<li>操作成功</li>&ErrorURL=user/award.asp")
		Response.End()
	end if
elseif request.QueryString("Act")="deleteresult" then
	prizeID=NoSqlHack(request("prizeid"))
	usernumber=NoSqlHack(request("usernumber"))
	if trim(prizeID)<>"" And  trim(usernumber)<>"" then
		User_Conn.execute("Delete from FS_ME_User_Prize where prizeid="&CintStr(prizeid)&" and usernumber='"&NoSqlHack(usernumber)&"'")
	End if
	response.Write("ok")
	Response.End()
end if
if err.number>0 then
	Response.Redirect("../error.asp?ErrCodes=<li>"&err.description&"</li>")
	Response.End()
end if
Set awardRs=nothing
Set prizeRs=nothing
Conn.close
User_Conn.close
Set Conn=nothing
Set User_Conn=nothing
%>







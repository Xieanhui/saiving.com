<%
Class Cls_Award
	private cls_AwardID,cls_AwardName,cls_AwardPic,cls_award_StartDate,cls_award_EndDate,cls_PrizeIDS
	private cls_PrizeID,cls_PrizeName,cls_prize_NeedPoint,cls_PrizeGrade,cls_PrizePic,cls_PrizeNum,cls_isChange,cls_storage,cls_prize_StartDate,cls_prize_EndDate,cls_PrizeDesc,cls_perUserNum,cls_provider
	private cls_AnswerID,cls_AnswerDesc
	private cls_AID,cls_ATopic,cls_answer_NeedPoint,cls_PrizePoint,cls_ADesc,cls_APic,cls_AStartDate,cls_AEndDate,cls_AnswerIDS,cls_RightAnswerID
	'������������������������������������������������������������������������������������������������������������������������������������������������������
	'��û��ֳ齱�Ļ�����Ϣ
	public function getAwardInfo(id)
		Dim awardRs,sql_cmd
		Set awardRs=server.CreateObject(G_FS_RS)
		sql_cmd="select AwardName,AwardPic,StartDate,EndDate,PrizeIDS from FS_ME_award where AwardID="&CintStr(id)
		awardRs.open sql_cmd,User_Conn,1,1
		cls_AwardID=id
		cls_AwardName=awardRs("AwardName")
		cls_AwardPic=awardRs("AwardPic")
		cls_award_StartDate=awardRs("StartDate")
		cls_award_EndDate=awardRs("EndDate")
		cls_PrizeIDS=awardRs("PrizeIDS")
		awardRs.close
		set awardRs=nothing
	End function
	'������������������������������������������������������������������������������������������������������������������������������������������������������
	'��ý�Ʒ�Ļ�����Ϣ
	public function getPrizeInfo(id)
		Dim prizeRs,sql_cmd
		Set prizeRs=Server.CreateObject(G_FS_RS)
		sql_cmd="Select PrizeName,NeedPoint,PrizeGrade,PrizePic,PrizeNum,isChange,storage,StartDate,EndDate,PrizeDesc,perUserNum,provider from FS_ME_Prize where PrizeID="&CintStr(id)
		prizeRs.open sql_cmd,User_Conn,1,1
		cls_PrizeID=id
		cls_PrizeName=prizeRs("PrizeName")
		cls_prize_NeedPoint=prizeRs("NeedPoint")
		cls_PrizeGrade=prizeRs("PrizeGrade")
		cls_PrizePic=prizeRs("PrizePic")
		cls_PrizeNum=prizeRs("PrizeNum")
		cls_isChange=prizeRs("isChange")
		cls_storage=prizeRs("storage")
		cls_Prize_StartDate=prizeRs("StartDate")
		cls_Prize_EndDate=prizeRs("EndDate")
		cls_PrizeDesc=prizeRs("PrizeDesc")
		cls_perUserNum=prizeRs("perUserNum")
		cls_provider=prizeRs("provider")
	End function
	'������������������������������������������������������������������������������������������������������������������������������������������������������
	'��û��־�����Ϣ
	public function getAnswerForPoint(id)
		Dim AnswerRs,sql_cmd
		Set AnswerRs=Server.CreateObject(G_FS_RS)
		sql_cmd="Select ATopic,NeedPoint,PrizePoint,ADesc,APic,AStartDate,AEndDate,AnswerIDS,RightAnswerID from FS_ME_AnswerForPoint where AID="&CintStr(id)
		answerRs.open sql_cmd,User_Conn,1,1
		cls_AID=id
		cls_ATopic=AnswerRs("ATopic")
		cls_answer_NeedPoint=AnswerRs("NeedPoint")
		cls_PrizePoint=AnswerRs("PrizePoint")
		cls_ADesc=AnswerRs("ADesc")
		cls_APic=AnswerRs("APic")
		cls_AStartDate=AnswerRs("AStartDate")
		cls_AEndDate=AnswerRs("AEndDate")
		cls_AnswerIDS=AnswerRs("AnswerIDS")
		cls_RightAnswerID=AnswerRs("RightAnswerID")
	End function
	'������������������������������������������������������������������������������������������������������������������������������������������������������
	'����ʴ��
	public function getAnswer(id)
		Dim answerRs,sql_cmd
		Set answerRs=Server.CreateObject(G_FS_RS)
		sql_cmd="Select AnswerID,AnswerDesc from FS_ME_Answer where AnswerID="&CintStr(Id)
		answerRs.open sql_cmd,User_Conn,1,1
		cls_AnswerID=id
		cls_AnswerDesc=answerRs("AnswerDesc")
	End function 
	'������������������������������������������������������������������������������������������������������������������������������������������������������
	'�Ƿ��н����еĳ齱
	public function activeAward()
		Dim active_TF_Rs1,active_TF_Rs2,sql_cmd,sql_cmd2,activeTF1,activeTF2
		activeTF1=false
		activeTF2=false
		if  G_IS_SQL_DB=0 then
			sql_cmd="select AID from FS_ME_AnswerForPoint where AEndDate>#"&DateValue(Now)&"#"
			sql_cmd2="select AwardID from FS_ME_award where EndDate>#"&DateValue(Now)&"#"
		Else
			sql_cmd="select AID from FS_ME_AnswerForPoint where AEndDate>'#"&DateValue(Now)&"#'"
			sql_cmd2="select AwardID from FS_ME_award where EndDate>'#"&DateValue(Now)&"#'"
		End if
		Set active_TF_Rs1=User_Conn.execute(sql_cmd)
		Set active_TF_Rs2=User_Conn.execute(sql_cmd2)
		if not active_TF_Rs1.eof or not active_TF_Rs1.bof then
			activeTF1=true
		End if
		if not active_TF_Rs2.eof or not active_TF_Rs2.bof then
			activeTF2=true
		End if
		activeAward=activeTF1 and activeTF2
		active_TF_Rs1.close
		active_TF_Rs2.close
		set active_TF_Rs1=nothing
		set active_TF_Rs2=nothing
	End function
	'������������������������������������������������������������������������������������������������������������������������������������������������������
	'���ֳ齱���ֶ�[cls_AwardID,cls_AwardName,cls_NeedPoint,cls_AwardPic,cls_StartDate,cls_EndDate,cls_PrizeIDS]
	public property get awardid
		awardid=cls_AwardID
	end property
	
	public property get AwardName
		AwardName=cls_AwardName
	end property
		
	public property get AwardPic
		AwardPic=cls_AwardPic
	end property
	
	public property get award_StartDate
		award_StartDate=cls_award_StartDate
	end property
	
	public property get award_EndDate
		award_EndDate=cls_award_EndDate
	end property
	
	public property get PrizeIDS'��Ʒid����
		PrizeIDS=cls_PrizeIDS
	end property
	'������������������������������������������������������������������������������������������������������������������������������������������������������
	'��Ʒ���ֶ�[cls_PrizeID,cls_PrizeName,cls_NeedPoint,cls_PrizeGrade,cls_PrizePic,cls_PrizeNum,cls_PrizeUserID,cls_isChange,cls_storage,cls_StartDate,cls_EndDate,cls_PrizeDesc,cls_perUserNum,cls_provider]

	public property get PrizeID
		PrizeID=cls_PrizeID
	End property
	
	public property get PrizeName
		PrizeName=cls_PrizeName
	End property
	
	public property get prize_NeedPoint'�μӳ齱��Ҫ�Ļ���
		prize_NeedPoint=cls_prize_NeedPoint
	End property
	
	public property get PrizeGrade'���Ƚ�
		PrizeGrade=cls_PrizeGrade
	End property
	public property get PrizePic
		PrizePic=cls_PrizePic
	End property	
	
	public property get PrizeNum'��Ʒ����
		PrizeNum=cls_PrizeNum
	End property
		
	public property get isChange
		isChange=cls_isChange
	End property
	
	public property get storage
		storage=cls_storage
	End property
	
	public property get Prize_StartDate
		Prize_StartDate=cls_Prize_StartDate
	End property
	
	public property get Prize_EndDate
		Prize_EndDate=cls_Prize_EndDate
	End property
	
	public property get PrizeDesc
		PrizeDesc=cls_PrizeDesc
	End property
	
	public property get perUserNum
		perUserNum=cls_perUserNum
	End property
	
	public property get provider
		provider=cls_provider
	End property
	'������������������������������������������������������������������������������������������������������������������������������������������������������
	'�����ʴ��ֶ�[cls_AID,cls_ATopic,cls_NeedPoint,cls_PrizePoint,cls_ADesc,cls_APic,cls_AStartDate,cls_AEndDate,cls_AnswerIDS,cls_RightAnswerID]
	public property get aid
		aid=cls_AID
	end property
	
	public property get ATopic
		ATopic=cls_ATopic
	end property
	
	public property get answer_NeedPoint
		answer_NeedPoint=cls_answer_NeedPoint
	end property
	
	public property get PrizePoint
		PrizePoint=cls_PrizePoint
	end property
	
	public property get ADesc
		ADesc=cls_ADesc
	end property
	
	public property get APic
		APic=cls_APic
	end property
	
	public property get AStartDate
		AStartDate=cls_AStartDate
	end property
	
	public property get AEndDate
		AEndDate=cls_AEndDate
	end property
	
	public property get AnswerIDS
		AnswerIDS=cls_AnswerIDS
	end property
	
	public property get RightAnswerID
		RightAnswerID=cls_RightAnswerID
	end property
	'������������������������������������������������������������������������������������������������������������������������������������������������������
	'��cls_AnswerID,cls_AnswerDesc
	public property get AnswerID
		AnswerID=cls_AnswerID
	end property

	public property get AnswerDesc
		AnswerDesc=cls_AnswerDesc
	end property
End Class
%>






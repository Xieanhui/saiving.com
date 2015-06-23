<%
Class cls_Contr
	Private C_ContID,C_ContSytle,C_ContTitle,C_SubTitle,C_ContContent,C_AddTime,C_PassTime,C_ClassID,C_MainID,C_KeyWords,C_IsPublic,C_InfoType,C_UserNumber,C_OtherContent,C_IsLock,C_isTF,C_Hits,C_AdminLock,C_PicFile,C_TempletID,C_FileName,C_FileExeName,C_AuditTF,C_Untread,C_type
	
	Public function getContrInfo(id)
		dim contrRs
		Set contrRs=Server.CreateObject(G_FS_RS)
		contrRs.open "select ContID,ContSytle,ContTitle,SubTitle,ContContent,AddTime,PassTime,ClassID,MainID,KeyWords,IsPublic,InfoType,UserNumber,OtherContent,IsLock,isTF,Hits,AdminLock,PicFile,TempletID,FileName,FileExeName,AuditTF,Untread,type from FS_ME_InfoContribution where ContID="&CintStr(ID),User_Conn,1,1
		if not contrRs.eof then
			C_ContID=contrRs("ContID")
			C_ContSytle=contrRs("ContSytle")
			C_ContTitle=contrRs("ContTitle")
			C_SubTitle=contrRs("SubTitle")
			C_ContContent=contrRs("ContContent")
			C_AddTime=contrRs("AddTime")
			C_PassTime=contrRs("PassTime")
			C_ClassID=contrRs("ClassID")
			C_MainID=contrRs("MainID")
			C_KeyWords=contrRs("KeyWords")
			C_IsPublic=contrRs("IsPublic")
			C_InfoType=contrRs("InfoType")
			C_UserNumber=contrRs("UserNumber")
			C_OtherContent=contrRs("OtherContent")
			C_IsLock=contrRs("IsLock")
			C_isTF=contrRs("isTF")
			C_Hits=contrRs("Hits")
			C_AdminLock=contrRs("AdminLock")
			C_PicFile=contrRs("PicFile")
			C_TempletID=contrRs("TempletID")
			C_FileName=contrRs("FileName")
			C_FileExeName=contrRs("FileExeName")
			C_AuditTF=contrRs("AuditTF")
			C_Untread=contrRs("Untread")
			C_type=contrRs("type")
		End if
		contrRs.close
		Set contrRs=nothing
	End Function
	'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	public property get id
		id=C_ContID
	End property
	
	public property get ContSytle'0原创，1转载，3代理
		ContSytle=C_ContSytle
	End property
	
	public property get ContTitle'主标题
		ContTitle=C_ContTitle
	End property
	
	public property get SubTitle'副标题
		SubTitle=C_SubTitle
	End property
		
	public property get ContContent'正文
		ContContent=C_ContContent
	End property

	public property get AddTime'添加时间
		AddTime=C_AddTime
	End property

	public property get PassTime'审核通过时间
		PassTime=C_PassTime
	End property

	public property get ClassID'专栏ID
		ClassID=C_ClassID
	End property

	public property get MainID'主站分类ID
		MainID=C_MainID
	End property

	public property get KeyWords'关键字
		KeyWords=C_KeyWords
	End property

	public property get IsPublic'是否发布到总站，1为是发布到总站显示，0为发布到自己的空间。空间地址：/用户目录/用户编号
		IsPublic=C_IsPublic
	End property

	public property get InfoType'信息级；普通：0，优先：1，加急：2
		InfoType=C_InfoType
	End property

	public property get UserNumber'发布者编号
		UserNumber=C_UserNumber
	End property

	public property get OtherContent'备注，比如文学类：此评语供编辑审核或推荐参考，内容可为作品导读、评析或创作感言等，较好的评语在审核通过后会显示到文章页面。填写此项有助于您的作品快速审核或加深读者理解。
		OtherContent=C_OtherContent
	End property

	public property get IsLock'是否锁定
		IsLock=C_IsLock
	End property

	public property get isTF'是否推荐
		isTF=C_isTF
	End property
	
	public property get Hits'点击数
		Hits=C_Hits
	End property

	public property get AdminLock'管理员锁定
		AdminLock=C_AdminLock
	End property

	public property get PicFile'图片地址
		PicFile=C_PicFile
	End property

	public property get TempletID'信息模板ID（暂时保留）
		TempletID=C_TempletID
	End property

	public property get FileName'静态文件文件名
		FileName=C_FileName
	End property
	
	public property get FileExeName'扩展名
		FileExeName=C_FileExeName
	End property
	
	public property get AuditTF'是否已审核(1：已审核；0:未审核）
		AuditTF=C_AuditTF
	End property
	
	public property get Untread'是否退稿
		Untread=C_Untread
	End property

	public property get ctype'0为新闻，1为下载，2为商品
		ctype=C_type
	End property

End Class
%>






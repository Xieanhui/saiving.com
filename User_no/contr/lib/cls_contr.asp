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
	'����������������������������������������������������������������������������
	public property get id
		id=C_ContID
	End property
	
	public property get ContSytle'0ԭ����1ת�أ�3����
		ContSytle=C_ContSytle
	End property
	
	public property get ContTitle'������
		ContTitle=C_ContTitle
	End property
	
	public property get SubTitle'������
		SubTitle=C_SubTitle
	End property
		
	public property get ContContent'����
		ContContent=C_ContContent
	End property

	public property get AddTime'���ʱ��
		AddTime=C_AddTime
	End property

	public property get PassTime'���ͨ��ʱ��
		PassTime=C_PassTime
	End property

	public property get ClassID'ר��ID
		ClassID=C_ClassID
	End property

	public property get MainID'��վ����ID
		MainID=C_MainID
	End property

	public property get KeyWords'�ؼ���
		KeyWords=C_KeyWords
	End property

	public property get IsPublic'�Ƿ񷢲�����վ��1Ϊ�Ƿ�������վ��ʾ��0Ϊ�������Լ��Ŀռ䡣�ռ��ַ��/�û�Ŀ¼/�û����
		IsPublic=C_IsPublic
	End property

	public property get InfoType'��Ϣ������ͨ��0�����ȣ�1���Ӽ���2
		InfoType=C_InfoType
	End property

	public property get UserNumber'�����߱��
		UserNumber=C_UserNumber
	End property

	public property get OtherContent'��ע��������ѧ�ࣺ�����﹩�༭��˻��Ƽ��ο������ݿ�Ϊ��Ʒ�����������������Եȣ��Ϻõ����������ͨ�������ʾ������ҳ�档��д����������������Ʒ������˻���������⡣
		OtherContent=C_OtherContent
	End property

	public property get IsLock'�Ƿ�����
		IsLock=C_IsLock
	End property

	public property get isTF'�Ƿ��Ƽ�
		isTF=C_isTF
	End property
	
	public property get Hits'�����
		Hits=C_Hits
	End property

	public property get AdminLock'����Ա����
		AdminLock=C_AdminLock
	End property

	public property get PicFile'ͼƬ��ַ
		PicFile=C_PicFile
	End property

	public property get TempletID'��Ϣģ��ID����ʱ������
		TempletID=C_TempletID
	End property

	public property get FileName'��̬�ļ��ļ���
		FileName=C_FileName
	End property
	
	public property get FileExeName'��չ��
		FileExeName=C_FileExeName
	End property
	
	public property get AuditTF'�Ƿ������(1������ˣ�0:δ��ˣ�
		AuditTF=C_AuditTF
	End property
	
	public property get Untread'�Ƿ��˸�
		Untread=C_Untread
	End property

	public property get ctype'0Ϊ���ţ�1Ϊ���أ�2Ϊ��Ʒ
		ctype=C_type
	End property

End Class
%>






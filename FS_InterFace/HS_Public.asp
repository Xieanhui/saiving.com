<%
Class cls_HS
	'��Ѷ����ˢ�´������࣬���ڻ�ô���
	'������Fs_ns.Get_LableChar("��ǩ����","")
	'�Ǿ���Ѷ��˾������ɣ�����������ҵ��;
	'�ٷ�վ��Foosun.cn
	'����֧����̳��bbs.foosun.net

	'�������____________________________________________________________________
	Private m_Rs,m_FSO,m_Dict
	Private m_PathDir,m_Path_UserDir,m_Path_User,m_Path_adminDir,m_Path_UserPageDir,m_Path_Templet
	Private m_Err_Info,m_Err_NO

	Public Property Get Err_Info()
		Err_Info = m_Err_Info
	End Property

	Public Property Get Err_NO()
		Err_NO = m_Err_NO
	End Property

	Private Sub Class_initialize()
		Set m_Rs = Server.CreateObject(G_FS_RS)
		Set m_FSO = Server.CreateObject(G_FS_FSO)
		Set m_Dict = Server.CreateObject(G_FS_DICT)
		m_PathDir = replace("/"&G_VIRTUAL_ROOT_DIR&"/","//","/")
		m_Path_UserDir = replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_USER_DIR&"/","//","/")
		m_Path_UserPageDir = replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_USERFILES_DIR&"/","//","/")
		m_Path_Templet  = replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_TEMPLETS_DIR&"/","//","/")
	End Sub

	Private Sub Class_Terminate()
		Set m_Rs = Nothing
		Set m_FSO = Nothing
		Set m_Dict = Nothing
	End Sub

	'��ò���____________________________________________________________________
	Public Function get_LableChar(f_Lable,f_Id,f_Type)
		get_LableChar = ""
	End Function
End Class
%>
<%
Class Cls_SysConfig
	private MF_ID,MF_Domain,MF_SiteName,MF_eMail,MF_Soft_Version,MF_Copyright_Info,MF_Site_lock,MF_UpFile_Type,MF_UpFile_Size,MF_Index_templet,MF_Index_File_Name,MF_WriteType
	private MF_Mail_Server,MF_Mail_Name,MF_Mail_Pass_word,MF_Index_Refresh,MF_Class_Refresh,MF_Write_Type,MF_Date_Path,MF_Copy_Right,MF_login_style,f_MF_Encript_SN
	private p_Save_Rs,strShowErr
	'---2007-02-12 By Ken
	private MF_MaxNum_Style,MF_Define_MaxNum,LabelContent_MaxNum
	'-----
	private MF_picClassid,MF_markType,MF_markText,MF_markFontSize,MF_markFontColor,MF_markFontName,MF_markFontBond,MF_markPicture,MF_markOpacity,MF_MarkTranspColor,MF_markWidth,MF_markHeight,MF_markPosition
'����ͼ���ܲ���
	private MF_thumbnailComponent,MF_rateTF,MF_thumbnailWidth,MF_thumbnailHeight,MF_thumbnailRate
	
	public function getSysParam()
		Dim sysRs,sqlStatement
		Set sysRs=Server.CreateObject(G_FS_RS)
		sqlStatement="select ID,MF_Domain,MF_Site_Name,MF_eMail,MF_Soft_Version,MF_Encript_SN,MF_Copyright_Info,MF_Site_lock,MF_UpFile_Type,MF_UpFile_Size,MF_Index_Templet,MF_Index_File_Name,MF_Mail_Server,MF_Mail_Name,MF_Mail_Pass_Word,MF_Index_Refresh,MF_Class_Refresh,MF_Login_style,MF_writeType,MarkType,MarkText,MarkFontSize,MarkFontColor,MarkFontName,MarkFontBond,MarkPicture,MarkOpacity,MarkWidth,MarkHeight,MarkTranspColor,MarkPosition,ThumbnailComponent,RateTF,ThumbnailWidth,ThumbnailHeight,ThumbnailRate,PicClassid,Style_MaxNum,Define_MaxNum,LabelContent_MaxNum from FS_MF_Config"
		sysRs.open sqlStatement,Conn,1,3
		MF_ID=sysRs("ID")
		f_MF_Encript_SN = sysRs("MF_Encript_SN")
		MF_Domain=sysRs("MF_Domain")
		MF_SiteName=sysRs("MF_Site_Name")
		MF_eMail=sysRs("MF_eMail")
		MF_Soft_Version=sysRs("MF_Soft_Version")
		MF_Copyright_Info=sysRs("MF_Copyright_Info")
		MF_Site_lock=sysRs("MF_Site_lock")
		MF_UpFile_Type=sysRs("MF_UpFile_Type")
		MF_UpFile_Size=sysRs("MF_UpFile_Size")
		MF_Index_Templet=sysRs("MF_Index_Templet")
		MF_Index_File_Name=sysRs("MF_Index_File_Name")
		MF_Mail_Server=sysRs("MF_Mail_Server")
		MF_Mail_Name=sysRs("MF_Mail_Name")
		MF_Mail_Pass_Word=sysRs("MF_Mail_Pass_Word")
		MF_Index_Refresh=sysRs("MF_Index_Refresh")
		MF_Class_Refresh=sysRs("MF_Class_Refresh")
		MF_Login_style=sysRs("MF_Login_style")
		MF_WriteType=sysRs("MF_WriteType")
		MF_MarkType=sysRs("MarkType")
		MF_MarkText=sysRs("MarkText")
		MF_MarkFontSize=sysRs("MarkFontSize")
		MF_MarkFontColor=sysRs("MarkFontColor")
		MF_MarkFontName=sysRs("MarkFontName")
		MF_MarkFontBond=sysRs("MarkFontBond")
		MF_MarkPicture=sysRs("MarkPicture")
		MF_MarkOpacity=sysRs("MarkOpacity")
		MF_MarkWidth=sysRs("MarkWidth")
		MF_MarkHeight=sysRs("MarkHeight")
		MF_MarkTranspColor=sysRs("MarkTranspColor")
		MF_MarkPosition=sysRs("MarkPosition")
		MF_ThumbnailComponent=sysRs("ThumbnailComponent")
		MF_RateTF=sysRs("RateTF")
		MF_ThumbnailWidth=sysRs("ThumbnailWidth")
		MF_ThumbnailHeight=sysRs("ThumbnailHeight")
		MF_ThumbnailRate=sysRs("ThumbnailRate")
		MF_PicClassid=sysRs("PicClassid")
		'---2007-02-12 By Ken
		MF_MaxNum_Style = sysRs("Style_MaxNum")
		MF_Define_MaxNum = sysRs("Define_MaxNum")
		LabelContent_MaxNum = sysRs("LabelContent_MaxNum")
		'---End
	End function
	
	public property get id()'ϵͳ����id���������ã��ü�¼ֻ��һ����
		id=MF_ID
	end property
	
	public property get MF_Encript_SN()
		MF_Encript_SN=f_MF_Encript_SN
	end property
	
	public property get domain()'ϵͳ������
		domain=MF_Domain
	end property
	
	public property get sitename()'ϵͳ������
		sitename=MF_SiteName
	end property
	
	public property get email()'
		email=MF_eMail
	end property
	
	public property get soft_version()'ϵͳ�汾��
		Soft_Version=MF_Soft_Version
	end property
	
	public property get copyright_info()'ϵͳ��Ȩ��Ϣ
		copyright_info=MF_Copyright_Info
	end property
	
	public property get site_lock()'վ������
		site_lock=MF_Site_lock
	end property
	
	public property get upFile_type()
		upFile_type=MF_upFile_type
	end property
	
	public property get upFile_Size()
		upFile_Size=MF_UpFile_Size
	end property
	
	public property get index_Templet()'ϵͳ��ҳģ��
		index_Templet=MF_Index_Templet
	end property
	
	public property get index_File_Name()'ϵͳ��ҳģ��
		index_File_Name=MF_Index_File_Name
	end property
	
	public property get mail_Server()'ϵͳ�ʼ�������
		mail_Server=MF_Mail_Server
	end property
	public property get mail_Name()
		mail_Name=MF_Mail_Name
	end property
	
	public property get mail_Pass_Word()'�ʼ���������½����
		mail_Pass_Word=MF_Mail_Pass_Word
	end property
	
	public property get index_Refresh()'��ҳ�Զ�ˢ�¼��ʱ�䣬��λ���ӣ�-1���Զ�ˢ�£�0Ϊ����ˢ�£�����ֵΪˢ��ʱ��ļ�������磬5��Ϊ5���ӣ�
		index_Refresh=MF_Index_Refresh
	end property

	public property get class_Refresh()'��Ŀ�Զ�ˢ�¼��ʱ�䣬��λ���ӣ�-1���Զ�ˢ�£�0Ϊ����ˢ�£�����ֵΪˢ��ʱ��ļ�������磬5��Ϊ5���ӣ�
		class_Refresh=MF_Class_Refresh
	end property

	public property get login_style()'��½���
		login_style=MF_Login_style
	end property
	
	public property get writeType()
		writeType=MF_WriteType
	end property
	'--------�������---------------------------------
	public property get picClassid()'ˮӡ���
		picClassid=MF_picClassid
	end property

	public property get markType()'ˮӡ����(���֣�ͼƬ)
		markType=MF_MarkType
	end property
	
	public property get markText()'ˮӡ����
		markText=MF_MarkText
	end property

	public property get markFontSize()'ˮӡ���ִ�С
		markFontSize=MF_MarkFontSize
	end property
	
	public property get markFontColor()'ˮӡ������ɫ
		markFontColor=MF_MarkFontColor
	end property

	public property get markFontName()'ˮӡ��������
		markFontName=MF_MarkFontName
	end property
	
	public property get markFontBond()'ˮӡ�����Ƿ����
		markFontBond=MF_MarkFontBond
	end property
	
	public property get markPicture()'ˮӡͼƬ
		markPicture=MF_MarkPicture
	end property

	public property get markOpacity()'ˮӡͼƬ͸����
		markOpacity=MF_MarkOpacity
	end property
	
	public property get markWidth()'ˮӡͼƬ͸����
		markWidth=MF_MarkWidth
	end property

	public property get markHeight()'ˮӡͼƬ͸����
		markHeight=MF_MarkHeight
	end property
	
	public property get markTranspColor()'ˮӡͼƬȥ����ɫ
		markTranspColor=MF_MarkTranspColor
	end property

	public property get markPosition()'ˮӡλ��
		markPosition=MF_MarkPosition
	end property
	
	'����ͼ���-----------------------
	public property get thumbnailComponent()'����ͼ���
		thumbnailComponent=MF_ThumbnailComponent
	end property

	public property get rateTF()'�Ƿ񰴱�����������ͼ
		rateTF=MF_RateTF
	end property
	
	public property get thumbnailWidth()'����ͼ���
		thumbnailWidth=MF_ThumbnailWidth
	end property
	
	public property get thumbnailHeight()'����ͼ�߶�
		thumbnailHeight=MF_ThumbnailHeight
	end property
	
	public property get thumbnailRate()'����ͼ���ű���
		thumbnailRate=MF_ThumbnailRate
	end property
	'---2007-02-12 By Ken
	public property get Style_MaxNum()
		Style_MaxNum=MF_MaxNum_Style
	end property
	
	public property get Define_MaxNum()
		Define_MaxNum=MF_Define_MaxNum
	end property
	
	public property get Label_MaxNum()
		Label_MaxNum=LabelContent_MaxNum
	end property
	'--------End 
End class
%>
<%
Class Cls_House
	private hs_id,hs_HouseName,hs_Position,hs_Direction,hs_Class,hs_OpenDate,hs_PreSaleNumber,hs_IssueDate,hs_PreSaleRange,hs_Status,hs_Price,hs_PubDate,hs_tel,hs_Click,hs_UserNumber,hs_Audited,hs_editor,hs_picNumber,hs_introduction,hs_KaiFaShang
	Private house_rs
	public Function getQuotationInfo(id)
		Set house_rs=Server.CreateObject(G_FS_RS)
		house_rs.open "select ID,HouseName,Position,Direction,Class,OpenDate,PreSaleNumber,IssueDate,PreSaleRange,Status,Price,PubDate,Tel,Click,UserNumber,Audited,editor,picNumber,introduction,KaiFaShang from FS_HS_Quotation where id="&CintStr(id),Conn,1,1
		hs_id=house_rs("id")
		hs_HouseName=house_rs("HouseName")
		hs_Position=house_rs("Position")
		hs_KaiFaShang=house_rs("KaiFaShang")'--------2/13--chen-----
		hs_Direction=house_rs("Direction")
		hs_Class=house_rs("Class")
		hs_OpenDate=house_rs("OpenDate")
		hs_PreSaleNumber=house_rs("PreSaleNumber")
		hs_IssueDate=house_rs("IssueDate")
		hs_PreSaleRange=house_rs("PreSaleRange")
		hs_Status=house_rs("Status")
		hs_Price=house_rs("Price")
		hs_PubDate=house_rs("PubDate")
		hs_tel=house_rs("tel")
		hs_Click=house_rs("Click")
		hs_UserNumber=house_rs("UserNumber")
		hs_Audited=house_rs("Audited")
		hs_editor=house_rs("editor")
		hs_picNumber=house_rs("picNumber")
		hs_introduction=house_rs("introduction")
	End Function
	
	public property get id()
		id=hs_id
	end property
	public property get houseName()'1.¥������
		houseName=hs_HouseName
	end property
	'---------------------2/13-----chen---------
	public property get KaiFaShang()'������
		KaiFaShang=hs_KaiFaShang
	end property
	'-------------------------------------------
	public property get position()'2.λ��
		position=hs_Position
	end property
	public property get direction()'3.��λ
		direction=hs_Direction
	end property
	public property get hclass()'4.����
		hclass=hs_class
	end property
	public property get openDate()'5.��������
		openDate=hs_OpenDate
	end property
	public property get preSaleNumber()'6.Ԥ�����֤����
		preSaleNumber=hs_PreSaleNumber
	end property
	public property get issueDate()'7.��֤����
		issueDate=hs_IssueDate
	end property
	public property get preSaleRange()'8.Ԥ�۷�Χ
		preSaleRange=hs_PreSaleRange
	end property
	public property get status()'9.¥��״̬[1.չʾ��2.�ڷ���3.�ַ�]
		status=hs_Status
	end property
	public property get price()'10.�۸�
		price=hs_Price
	end property
	public property get pubDate()'11.��������
		pubDate=hs_PubDate
	end property
	public property get click()'12.�����
		click=hs_Click
	end property
	public property get userNumber()'13.������Ϣ��
		userNumber=hs_UserNumber
	end property
	public property get audited()'14.���״̬
		audited=hs_Audited
	end property
	public property get editor()'15.�༭��
		editor=hs_editor
	end property
	public property get picNumber()'16.ͼƬ����
		picNumber=hs_picNumber
	end property
	public property get tel()'17.��ϵ�绰
		tel=hs_Tel
	end property
	public property get introduction()'18.����
		introduction=hs_introduction
	end property
End Class
%>








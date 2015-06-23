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
	public property get houseName()'1.楼盘名称
		houseName=hs_HouseName
	end property
	'---------------------2/13-----chen---------
	public property get KaiFaShang()'开发商
		KaiFaShang=hs_KaiFaShang
	end property
	'-------------------------------------------
	public property get position()'2.位置
		position=hs_Position
	end property
	public property get direction()'3.方位
		direction=hs_Direction
	end property
	public property get hclass()'4.类型
		hclass=hs_class
	end property
	public property get openDate()'5.开盘日期
		openDate=hs_OpenDate
	end property
	public property get preSaleNumber()'6.预售许可证号码
		preSaleNumber=hs_PreSaleNumber
	end property
	public property get issueDate()'7.发证日期
		issueDate=hs_IssueDate
	end property
	public property get preSaleRange()'8.预售范围
		preSaleRange=hs_PreSaleRange
	end property
	public property get status()'9.楼盘状态[1.展示；2.期房；3.现房]
		status=hs_Status
	end property
	public property get price()'10.价格
		price=hs_Price
	end property
	public property get pubDate()'11.发布日期
		pubDate=hs_PubDate
	end property
	public property get click()'12.点击数
		click=hs_Click
	end property
	public property get userNumber()'13.发布信息者
		userNumber=hs_UserNumber
	end property
	public property get audited()'14.审核状态
		audited=hs_Audited
	end property
	public property get editor()'15.编辑者
		editor=hs_editor
	end property
	public property get picNumber()'16.图片数量
		picNumber=hs_picNumber
	end property
	public property get tel()'17.联系电话
		tel=hs_Tel
	end property
	public property get introduction()'18.介绍
		introduction=hs_introduction
	end property
End Class
%>








<%
Class cls_Second
	Private s_SID,s_Class,s_UserNumber,s_Label,s_UseFor,s_FloorType,s_BelongType,s_HouseStyle,s_Structure,s_Area,s_BuildDate,s_Price,s_CityArea,s_Address,s_Floor,s_Position,s_Decoration,s_LinkMan,s_Contact,s_equip,s_Remark,s_PubDate,s_Audited,s_PicNumber
	
	Public Function getSecondInfo(id)
		Dim sqlstatement,secondRS
		sqlstatement="select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark,PubDate,Audited,PicNumber from FS_HS_Second where sid="&CintStr(id)
		Set secondRs=server.CreateObject(G_FS_RS)
		secondRs.open sqlstatement,Conn,1,1
		if not secondRs.eof then
			s_SID=secondRs("SID")
			s_Class=secondRs("Class")
			s_UserNumber=secondRs("UserNumber")
			s_Label=secondRs("Label")
			s_UseFor=secondRs("UseFor")
			s_FloorType=secondRs("FloorType")
			s_BelongType=secondRs("BelongType")
			s_HouseStyle=secondRs("HouseStyle")
			s_Structure=secondRs("Structure")
			s_Area=secondRs("Area")
			s_BuildDate=secondRs("BuildDate")
			s_Price=secondRs("Price")
			s_CityArea=secondRs("CityArea")
			s_Address=secondRs("Address")
			s_Floor=secondRs("Floor")
			s_Position=secondRs("Position")
			s_Decoration=secondRs("Decoration")
			s_LinkMan=secondRs("LinkMan")
			s_Contact=secondRs("Contact")
			s_equip=secondRs("equip")
			s_Remark=secondRs("Remark")
			s_PubDate=secondRs("PubDate")
			s_Audited=secondRs("Audited")
			s_PicNumber=secondRs("PicNumber")
		End if
	End function 
	
	public property get sid()
		sid=S_Sid
	End property
	
	public property get sClass()
		sClass=S_Class
	End property
	
	public property get UserNumber()
		UserNumber=S_UserNumber
	End property

	public property get Label()
		Label=S_Label
	End property

	public property get UseFor()
		UseFor=S_UseFor
	End property

	public property get FloorType()
		FloorType=S_FloorType
	End property

	public property get BelongType()
		BelongType=S_BelongType
	End property

	public property get HouseStyle()
		HouseStyle=s_HouseStyle
	End property

	public property get Structure()
		Structure=S_Structure
	End property

	public property get Area()
		Area=S_Area
	End property

	public property get BuildDate()
		BuildDate=s_BuildDate
	End property

	public property get Price()
		Price=s_Price
	End property

	public property get CityArea()
		CityArea=s_CityArea
	End property

	public property get Address()
		Address=s_Address
	End property

	public property get Floor()
		Floor=s_Floor
	End property

	public property get Position()
		Position=s_Position
	End property

	public property get Decoration()
		Decoration=s_Decoration
	End property

	public property get LinkMan()
		LinkMan=s_LinkMan
	End property

	public property get Contact()
		Contact=s_Contact
	End property

	public property get equip()
		equip=s_equip
	End property

	public property get Remark()
		Remark=s_Remark
	End property

	public property get PubDate()
		PubDate=s_PubDate
	End property

	public property get Audited()
		Audited=s_Audited
	End property

	public property get PicNumber()
		PicNumber=s_PicNumber
	End property


End Class
%>






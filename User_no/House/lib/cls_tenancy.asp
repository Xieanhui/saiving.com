<%
Class Cls_Tenancy
	private tenancy_tid,tenancy_UseFor,tenancy_class,tenancy_Position,tenancy_CityArea,tenancy_Price,tenancy_HouseStyle,tenancy_Area,tenancy_Floor,tenancy_BuildDate,tenancy_equip,tenancy_Decoration,tenancy_LinkMan,tenancy_Contact,tenancy_Period,tenancy_Remark,tenancy_PubDate,tenancy_Audited,tenancy_picNumber,tenancy_XiaoQuName,tenancy_XingZhi,tenancy_ZaWuJian,tenancy_JiaoTong
	Private tenancy_rs
	public Function getTenancyInfo(id)
		Set tenancy_rs=Server.CreateObject(G_FS_RS)
		tenancy_rs.open "select tid,UseFor,class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate,Audited,picNumber,XiaoQuName,XingZhi,ZaWuJian,JiaoTong from FS_HS_Tenancy where tid="&CintStr(id),Conn,1,1
	tenancy_tid=tenancy_rs("tid")
	tenancy_UseFor=tenancy_rs("UseFor")
	tenancy_XingZhi=tenancy_rs("XingZhi")
	tenancy_class=tenancy_rs("class")
	tenancy_Position=tenancy_rs("Position")
	tenancy_CityArea=tenancy_rs("CityArea")
	tenancy_Price=tenancy_rs("Price")
	tenancy_HouseStyle=tenancy_rs("HouseStyle")
	tenancy_Area=tenancy_rs("Area")
	tenancy_Floor=tenancy_rs("Floor")
	tenancy_BuildDate=tenancy_rs("BuildDate")
	tenancy_equip=tenancy_rs("equip")
	tenancy_Decoration=tenancy_rs("Decoration")
	tenancy_LinkMan=tenancy_rs("LinkMan")
	tenancy_Contact=tenancy_rs("Contact")
	tenancy_Period=tenancy_rs("Period")
	tenancy_Remark=tenancy_rs("Remark")
	tenancy_PubDate=tenancy_rs("PubDate")
	tenancy_Audited=tenancy_rs("Audited")
	tenancy_picNumber=tenancy_rs("picNumber")
	tenancy_XiaoQuName=tenancy_rs("XiaoQuName")
	tenancy_XingZhi=tenancy_rs("XingZhi")
	tenancy_ZaWuJian=tenancy_rs("ZaWuJian")
	tenancy_JiaoTong=tenancy_rs("JiaoTong")
	End Function
	
	public property get tid()
		id=tenancy_tid
	end property
	public property get UseFor()'1.用途:1住房,2写字间
		UseFor=tenancy_UseFor
	end property
	public property get tclass()'2.类型:1:出租2:求租3:出售4:求购5:合租6:转让
		tclass=tenancy_class
	end property
	public property get Position()'3.房源地址
		Position=tenancy_Position
	end property
	'-------------------------------------------
		public property get XiaoQuName()'小区名称'---------------2/13-chen-----
		XiaoQuName=tenancy_XiaoQuName
	end property
	public property get XingZhi()'房屋性质 1商品房,2集资房，3其他
		XingZhi=tenancy_XingZhi
	end property
	public property get ZaWuJian()'杂物间0 无 1 有
		ZaWuJian=tenancy_ZaWuJian
	end property
	public property get JiaoTong()'交通状况描述
		JiaoTong=tenancy_JiaoTong
	end property
	'-------------------------------------------
	public property get cityArea()'4.区县
		cityArea=tenancy_cityArea
	end property
	public property get Price()'5.租金(单位:元/月)
		Price=tenancy_Price
	end property
	public property get HouseStyle()'6.户型,存储格式:l,m,nl:室m:厅n:卫
		HouseStyle=tenancy_HouseStyle
	end property
	public property get Area()'7.建筑面积
		Area=tenancy_Area
	end property
	public property get Floor()'8.楼层,存储格式:m,nm:总层n:第几层
		Floor=tenancy_Floor
	end property
	public property get BuildDate()'9.建筑年代
		BuildDate=tenancy_BuildDate
	end property
	public property get equip()'10.配套设施,保存格式:l,m,n,x,y,zl:通水m:电n:气x:电话y:光纤z:表示宽带1表示有,0表示无 
		equip=tenancy_equip
	end property
	public property get Decoration()'11.装修情况:1: 简单装修2. 中档装修3. 高档装修
		Decoration=tenancy_Decoration
	end property
	public property get LinkMan()'12.联系人
		LinkMan=tenancy_LinkMan
	end property
	public property get Contact()'13.联系方式
		Contact=tenancy_Contact
	end property
	public property get Period()'14.有效期:一周,两周,三周,一月,两月),所有只保留三月
		Period=tenancy_Period
	end property
	public property get Remark()'15.备注
		Remark=tenancy_Remark
	end property
	public property get PubDate()'16.发布时间
		PubDate=tenancy_PubDate
	end property
	public property get Audited()'17.是否通过审核1:是,0未
		Audited=tenancy_Audited
	end property

	public property get picNumber()'18图片数量
		picNumber=tenancy_picNumber
	end property
End Class
%>







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
	public property get UseFor()'1.��;:1ס��,2д�ּ�
		UseFor=tenancy_UseFor
	end property
	public property get tclass()'2.����:1:����2:����3:����4:��5:����6:ת��
		tclass=tenancy_class
	end property
	public property get Position()'3.��Դ��ַ
		Position=tenancy_Position
	end property
	'-------------------------------------------
		public property get XiaoQuName()'С������'---------------2/13-chen-----
		XiaoQuName=tenancy_XiaoQuName
	end property
	public property get XingZhi()'�������� 1��Ʒ��,2���ʷ���3����
		XingZhi=tenancy_XingZhi
	end property
	public property get ZaWuJian()'�����0 �� 1 ��
		ZaWuJian=tenancy_ZaWuJian
	end property
	public property get JiaoTong()'��ͨ״������
		JiaoTong=tenancy_JiaoTong
	end property
	'-------------------------------------------
	public property get cityArea()'4.����
		cityArea=tenancy_cityArea
	end property
	public property get Price()'5.���(��λ:Ԫ/��)
		Price=tenancy_Price
	end property
	public property get HouseStyle()'6.����,�洢��ʽ:l,m,nl:��m:��n:��
		HouseStyle=tenancy_HouseStyle
	end property
	public property get Area()'7.�������
		Area=tenancy_Area
	end property
	public property get Floor()'8.¥��,�洢��ʽ:m,nm:�ܲ�n:�ڼ���
		Floor=tenancy_Floor
	end property
	public property get BuildDate()'9.�������
		BuildDate=tenancy_BuildDate
	end property
	public property get equip()'10.������ʩ,�����ʽ:l,m,n,x,y,zl:ͨˮm:��n:��x:�绰y:����z:��ʾ���1��ʾ��,0��ʾ�� 
		equip=tenancy_equip
	end property
	public property get Decoration()'11.װ�����:1: ��װ��2. �е�װ��3. �ߵ�װ��
		Decoration=tenancy_Decoration
	end property
	public property get LinkMan()'12.��ϵ��
		LinkMan=tenancy_LinkMan
	end property
	public property get Contact()'13.��ϵ��ʽ
		Contact=tenancy_Contact
	end property
	public property get Period()'14.��Ч��:һ��,����,����,һ��,����),����ֻ��������
		Period=tenancy_Period
	end property
	public property get Remark()'15.��ע
		Remark=tenancy_Remark
	end property
	public property get PubDate()'16.����ʱ��
		PubDate=tenancy_PubDate
	end property
	public property get Audited()'17.�Ƿ�ͨ�����1:��,0δ
		Audited=tenancy_Audited
	end property

	public property get picNumber()'18ͼƬ����
		picNumber=tenancy_picNumber
	end property
End Class
%>







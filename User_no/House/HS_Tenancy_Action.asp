<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
Dim tid,action,tenancyRs,sqlstatement,i
Dim floorArray(2), HouseStyleArray(3)
DIm t_TID,t_UseFor,t_Class,t_Position,t_CityArea,t_Price,t_HouseStyle,t_Area,t_Floor,t_BuildDate,t_equip,t_Decoration,t_LinkMan,t_Contact,t_period,t_Remark,t_PubDate,t_Audited,t_PicNumber,t_XiaoQuName,t_XingZhi,t_ZaWuJian,t_JiaoTong
MF_Default_Conn
action=request.QueryString("action")
tid=FormatIntArr(request("id"))
response.Charset="GB2312"
response.buffer=true	
Response.CacheControl = "no-cache"

if action="delete" then
	Conn.execute("Delete from FS_HS_Tenancy where tid in ("&tid&")")
	Response.Write("ok")
	response.End()
elseif action="Audit" then
	Conn.execute("Update FS_HS_Tenancy Set audited=1 where tid in ("&tid&")")
	Response.Write("ok")
	response.End()
elseif action="UnAudit" then
	Conn.execute("Update FS_HS_Tenancy Set audited=0 where tid in ("&tid&")")
	Response.Write("ok")
	response.End()
elseif action="add" then'------------------------------------------------------------------------------------------
	Set tenancyRs=Server.CreateObject(G_FS_RS)
	sqlstatement="select TID,UseFor,Class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate,Audited,PicNumber,UserNumber,XiaoQuName,XingZhi,ZaWuJian,JiaoTong from FS_HS_Tenancy"
	t_usefor=NoSqlHack(request.Form("sel_usefor"))
	t_class=NoSqlHack(request.Form("sel_class"))
	t_Position=NoSqlHack(request.Form("txt_Position"))
	t_XiaoQuName=NoSqlHack(request.Form("txt_XiaoQuName"))'---------------2/13-chen-----
	t_XingZhi=NoSqlHack(request.Form("sel_XingZhi"))
	t_ZaWuJian=NoSqlHack(request.Form("ZaWuJian"))
	t_JiaoTong=NoSqlHack(request.Form("txt_JiaoTong"))
	t_cityarea=NoSqlHack(request.Form("txt_cityarea"))
	HouseStyleArray(0)="1"
	HouseStyleArray(1)="1"
	HouseStyleArray(2)="1"
	if trim(request.Form("txt_HouseStyle_1"))<>"" then HouseStyleArray(0)=NoSqlHack(request.Form("txt_HouseStyle_1"))
	if trim(request.Form("txt_HouseStyle_2"))<>"" then HouseStyleArray(1)=NoSqlHack(request.Form("txt_HouseStyle_2"))
	if trim(request.Form("txt_HouseStyle_3"))<>"" then HouseStyleArray(2)=NoSqlHack(request.Form("txt_HouseStyle_3"))
	t_HouseStyle=HouseStyleArray(0)&","&HouseStyleArray(1)&","&HouseStyleArray(2)
	t_BuildDate=NoSqlHack(request.Form("txt_BuildDate"))
	t_Area=NoSqlHack(request.Form("txt_Area"))
	floorArray(0)="0"
	floorArray(1)="0"
	if trim(request.Form("txt_Floor_1"))<>"" then  floorArray(0)=NoSqlHack(request.Form("txt_Floor_1"))
	if trim(request.Form("txt_Floor_2"))<>"" then  floorArray(1)=NoSqlHack(request.Form("txt_Floor_2"))
	t_floor=floorArray(0)&","&floorArray(1)
	t_equip=NoSqlHack(request.Form("chk_equip"))
	t_Decoration=NoSqlHack(request.Form("sel_Decoration"))
	t_price=NoSqlHack(request.Form("txt_Price"))
	t_LinkMan=NoSqlHack(request.Form("txt_LinkMan"))
	t_contact=NoSqlHack(request.Form("txt_Contact"))
	t_Period=NoSqlHack(request.Form("txt_Period"))
	t_Remark=NoSqlHack(request.Form("txt_Remark"))
	t_picNumber=NoSqlHack(request.Form("PicNum"))
	t_pubDate=DateValue(Now)
	t_Audited=0
	tenancyRs.open sqlstatement,Conn,1,3
	tenancyRs.addnew
	tenancyRs("UseFor")=t_usefor
	tenancyRs("Class")=t_class
	tenancyRs("Position")=t_Position
	tenancyRs("XiaoQuName")=t_XiaoQuName '---------------2/13-chen-----
	tenancyRs("XingZhi")=t_XingZhi
	tenancyRs("ZaWuJian")=t_ZaWuJian
	tenancyRs("JiaoTong")=t_JiaoTong
	tenancyRs("CityArea")=t_cityarea
	tenancyRs("HouseStyle")=t_HouseStyle
	tenancyRs("Area")=t_Area
	if t_price="" then t_price=0
	tenancyRs("Price")=t_Price
	tenancyRs("Floor")=t_Floor
	if tenancyRs("BuildDate")="" then t_BuildDate=0
	tenancyRs("BuildDate")=t_BuildDate
	tenancyRs("equip")=t_equip
	tenancyRs("Decoration")=t_Decoration
	tenancyRs("Period")=t_Period
	tenancyRs("LinkMan")=t_LinkMan
	tenancyRs("contact")=t_contact
	tenancyRs("Remark")=right(t_Remark,250)
	tenancyRs("PubDate")=t_PubDate
	tenancyRs("Audited")=t_Audited
	tenancyRs("UserNumber")=session("FS_UserNumber")
	t_picNumber=trim(request.Form("txt_PicNum"))
	if t_picNumber="" then t_picNumber=0
	tenancyRs("picNumber")=t_picNumber
	tenancyRs.update
	if Cint(t_picNumber)>0 then
		for i=0 to Cint(t_picNumber)
			if trim(request.Form("txt_PicNum_"&(i+1)))<>"" then
				Conn.execute("Insert into FS_HS_Picture (HS_Type,ID,PIC) values(2,"&tenancyRs("tID")&",'"&NoSqlHack(request.Form("txt_PicNum_"&(i+1)))&"')")
			End if
		next
	End if
	tenancyRs.close
elseif action="edit" then
	Set tenancyRs=Server.CreateObject(G_FS_RS)
	sqlstatement="select TID,UseFor,Class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate,Audited,PicNumber,UserNumber,XiaoQuName,XingZhi,ZaWuJian,JiaoTong from FS_HS_Tenancy where tid="&tid
	t_usefor=NoSqlHack(request.Form("sel_usefor"))
	t_class=NoSqlHack(request.Form("sel_class"))
	t_Position=NoSqlHack(request.Form("txt_Position"))
	t_XiaoQuName=NoSqlHack(request.Form("txt_XiaoQuName"))'---------------2/13-chen-----
	t_XingZhi=NoSqlHack(request.Form("sel_XingZhi"))
	t_ZaWuJian=NoSqlHack(request.Form("ZaWuJian"))
	t_JiaoTong=NoSqlHack(request.Form("txt_JiaoTong"))
	t_cityarea=NoSqlHack(request.Form("txt_cityarea"))
	HouseStyleArray(0)="1"
	HouseStyleArray(1)="1"
	HouseStyleArray(2)="1"
	if trim(request.Form("txt_HouseStyle_1"))<>"" then HouseStyleArray(0)=NoSqlHack(request.Form("txt_HouseStyle_1"))
	if trim(request.Form("txt_HouseStyle_2"))<>"" then HouseStyleArray(1)=NoSqlHack(request.Form("txt_HouseStyle_2"))
	if trim(request.Form("txt_HouseStyle_3"))<>"" then HouseStyleArray(2)=NoSqlHack(request.Form("txt_HouseStyle_3"))
	t_HouseStyle=HouseStyleArray(0)&","&HouseStyleArray(1)&","&HouseStyleArray(2)
	t_BuildDate=NoSqlHack(request.Form("txt_BuildDate"))
	t_Area=NoSqlHack(request.Form("txt_Area"))
	if trim(request.Form("txt_Floor_1"))<>"" then  floorArray(0)=NoSqlHack(request.Form("txt_Floor_1"))
	if trim(request.Form("txt_Floor_2"))<>"" then  floorArray(1)=NoSqlHack(request.Form("txt_Floor_2"))
	t_floor=floorArray(0)&","&floorArray(1)
	t_equip=NoSqlHack(request.Form("chk_equip"))
	t_Decoration=NoSqlHack(request.Form("sel_Decoration"))
	t_price=NoSqlHack(request.Form("txt_Price"))
	t_LinkMan=NoSqlHack(request.Form("txt_LinkMan"))
	t_contact=NoSqlHack(request.Form("txt_Contact"))
	t_Period=NoSqlHack(request.Form("txt_Period"))
	t_Remark=NoSqlHack(request.Form("txt_Remark"))
	t_picNumber=NoSqlHack(request.Form("PicNum"))
	t_pubDate=DateValue(Now)
	t_Audited=0
	tenancyRs.open sqlstatement,Conn,1,3
	tenancyRs("UseFor")=t_usefor
	tenancyRs("Class")=t_class
	tenancyRs("Position")=t_Position
	tenancyRs("HouseStyle")=t_HouseStyle
	tenancyRs("Area")=t_Area
	tenancyRs("CityArea")=t_cityarea
	if t_price="" then t_price=0
	tenancyRs("Price")=t_Price
	tenancyRs("Floor")=t_Floor
	if tenancyRs("BuildDate")="" then t_BuildDate=0
	tenancyRs("BuildDate")=t_BuildDate
	tenancyRs("equip")=t_equip
	tenancyRs("Decoration")=t_Decoration
	tenancyRs("Period")=t_Period
	tenancyRs("Remark")=t_Remark
	tenancyRs("LinkMan")=t_LinkMan
	tenancyRs("contact")=t_contact
	tenancyRs("PubDate")=t_PubDate
	tenancyRs("Audited")=t_Audited
	tenancyRs("XiaoQuName")=t_XiaoQuName'---------------2/13-chen-----
	tenancyRs("XingZhi")=t_XingZhi
	tenancyRs("ZaWuJian")=t_ZaWuJian
	tenancyRs("JiaoTong")=t_JiaoTong
	tenancyRs("UserNumber")=session("FS_UserNumber")
	t_picNumber=trim(request.Form("txt_PicNum"))
	if t_picNumber="" then t_picNumber=0
	tenancyRs("picNumber")=t_picNumber
	tenancyRs.update
	if Cint(t_picNumber)>0 then
		for i=0 to Cint(t_picNumber)
			if trim(request.Form("txt_PicNum_"&(i+1)))<>"" then
				Conn.execute("Insert into FS_HS_Picture (HS_Type,ID,PIC) values(2,"&tenancyRs("tid")&",'"&NoSqlHack(request.Form("txt_PicNum_"&(i+1)))&"')")
			End if
		next
	End if
	tenancyRs.close
End if
Conn.close
User_Conn.close
Set User_Conn=nothing
Set Conn=nothing
Set tenancyRs=nothing
if err.number=0 then
	Response.Redirect("../lib/success.asp?ErrCodes=<li>操作成功</li>&ErrorURL=../House/HS_Tenancy.asp")
	Response.End()
Else
	Response.Redirect("../lib/error.asp?ErrCodes=<li>请检查输入是否合法</li>")
	Response.End()
End if
%>







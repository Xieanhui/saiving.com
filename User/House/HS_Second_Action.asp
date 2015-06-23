<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
Dim sid,action,secondRs,sqlstatement,i
Dim floorArray(2), HouseStyleArray(3)
DIm s_SID,s_Class,s_UserNumber,s_Label,s_UseFor,s_FloorType,s_BelongType,s_HouseStyle,s_Structure,s_Area,s_BuildDate,s_Price,s_CityArea,s_Address,s_Floor,s_Position,s_Decoration,s_LinkMan,s_Contact,s_equip,s_Remark,s_PubDate,s_Audited,s_PicNumber
action=request.QueryString("action")
sid=FormatIntArr(request("id"))
response.Charset="GB2312"
if action="delete" then
	Conn.execute("Delete from FS_HS_Second where sid in ("&sid&")")
	Response.Write("ok")
	response.End()
elseif action="add" then'------------------------------------------------------------------------------------------
	Set secondRs=Server.CreateObject(G_FS_RS)
	sqlstatement="select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark,PubDate,Audited,PicNumber from FS_HS_Second"
	s_Class = NoSqlHack(request.Form("sel_class"))
	s_UserNumber=session("FS_UserNumber")
	s_Label = NoSqlHack(request.Form("txt_Label"))
	s_UseFor = NoSqlHack(request.Form("sel_usefor"))
	s_FloorType = NoSqlHack(request.Form("sel_FloorType"))
	s_BelongType = NoSqlHack(request.Form("sel_BelongType"))
	HouseStyleArray(0)="1"
	HouseStyleArray(1)="1"
	HouseStyleArray(2)="1"
	if trim(request.Form("txt_HouseStyle_1"))<>"" then HouseStyleArray(0) = NoSqlHack(request.Form("txt_HouseStyle_1"))
	if trim(request.Form("txt_HouseStyle_2"))<>"" then HouseStyleArray(1) = NoSqlHack(request.Form("txt_HouseStyle_2"))
	if trim(request.Form("txt_HouseStyle_3"))<>"" then HouseStyleArray(2) = NoSqlHack(request.Form("txt_HouseStyle_3"))
	s_HouseStyle=HouseStyleArray(0)&","&HouseStyleArray(1)&","&HouseStyleArray(2)
	s_Structure = NoSqlHack(request.Form("sel_Structure"))
	s_Area = NoSqlHack(request.Form("txt_Area"))
	s_BuildDate = NoSqlHack(request.Form("txt_BuildDate"))
	s_Price = NoSqlHack(request.Form("txt_price"))
	s_CityArea = NoSqlHack(request.Form("txt_cityarea"))
	s_Address = NoSqlHack(request.Form("txt_Address"))
	floorArray(0)="0"
	floorArray(1)="0"
	if trim(request.Form("txt_Floor_1"))<>"" then  floorArray(0) = NoSqlHack(request.Form("txt_Floor_1"))
	if trim(request.Form("txt_Floor_2"))<>"" then  floorArray(1) = NoSqlHack(request.Form("txt_Floor_2"))
	s_Floor=floorArray(0)&","&floorArray(1)
	s_Position = NoSqlHack(request.Form("txt_Position"))
	s_Decoration = NoSqlHack(request.Form("sel_Decoration"))
	s_LinkMan = NoSqlHack(request.Form("txt_LinkMan"))
	s_Contact = NoSqlHack(request.Form("txt_Contact"))
	s_equip = NoSqlHack(request.Form("chk_equip"))
	s_Remark = NoSqlHack(request.Form("txt_Remark"))
	s_PubDate = DateValue(Now)
	s_Audited = 0
	s_PicNumber =  NoSqlHack(request.Form("PicNum"))
	'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	secondRs.open sqlstatement,Conn,1,3
	secondRs.addnew
	secondRs("Class")=s_Class
	secondRs("UserNumber")=s_UserNumber
	secondRs("label")=s_label
	secondRs("UseFor")=s_UseFor
	secondRs("FloorType")=s_FloorType
	secondRs("BelongType")=s_BelongType
	secondRs("HouseStyle")=s_HouseStyle
	secondRs("Structure")=s_Structure
	secondRs("Area")=s_Area
	if s_BuildDate="" then s_BuildDate=0	
	secondRs("BuildDate")=s_BuildDate
	if s_price="" then s_price=0
	secondRs("Price")=s_Price
	secondRs("CityArea")=s_CityArea
	secondRs("Address")=s_Address
	secondRs("Floor")=s_Floor
	secondRs("Position")=s_Position
	secondRs("LinkMan")=s_LinkMan
	secondRs("Contact")=s_Contact
	secondRs("equip")=s_equip
	secondRs("Decoration")=s_Decoration
	secondRs("Remark")=Right(s_Remark,250)
	secondRs("PubDate")=s_PubDate
	secondRs("Audited")=s_Audited
	s_picNumber = NoSqlHack(request.Form("txt_PicNum"))
	if s_picNumber="" then s_picNumber=0
	secondRs("picNumber")=s_picNumber
	secondRs.update
	if Cint(s_picNumber)>0 then
		for i=0 to Cint(s_picNumber)
			if trim(request.Form("txt_PicNum_"&(i+1)))<>"" then
				Conn.execute("Insert into FS_HS_Picture (ID,HS_Type,PIC) values("&secondRs("sID")&",3,'"& NoSqlHack(request.Form("txt_PicNum_"&(i+1)))&"')")
			End if
		next
	End if
	secondRs.close
elseif action="edit" then
	Set secondRs=Server.CreateObject(G_FS_RS)
	sqlstatement="select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark,PubDate,Audited,PicNumber from FS_HS_Second where sid="&CintStr(sid)
	s_Class = NoSqlHack(request.Form("sel_class"))
	s_UserNumber=session("FS_UserNumber")
	s_Label = NoSqlHack(request.Form("txt_Label"))
	s_UseFor = NoSqlHack(request.Form("sel_usefor"))
	s_FloorType = NoSqlHack(request.Form("sel_FloorType"))
	s_BelongType = NoSqlHack(request.Form("sel_BelongType"))
	HouseStyleArray(0)="1"
	HouseStyleArray(1)="1"
	HouseStyleArray(2)="1"
	if trim(request.Form("txt_HouseStyle_1"))<>"" then HouseStyleArray(0) = NoSqlHack(request.Form("txt_HouseStyle_1"))
	if trim(request.Form("txt_HouseStyle_2"))<>"" then HouseStyleArray(1) = NoSqlHack(request.Form("txt_HouseStyle_2"))
	if trim(request.Form("txt_HouseStyle_3"))<>"" then HouseStyleArray(2) = NoSqlHack(request.Form("txt_HouseStyle_3"))
	s_HouseStyle=HouseStyleArray(0)&","&HouseStyleArray(1)&","&HouseStyleArray(2)
	s_Structure = NoSqlHack(request.Form("sel_Structure"))
	s_Area = NoSqlHack(request.Form("txt_Area"))
	s_BuildDate = NoSqlHack(request.Form("txt_BuildDate"))
	s_Price = NoSqlHack(request.Form("txt_price"))
	s_CityArea = NoSqlHack(request.Form("txt_cityarea"))
	s_Address = NoSqlHack(request.Form("txt_Address"))
	floorArray(0)="0"
	floorArray(1)="0"
	if NoSqlHack(request.Form("txt_Floor_1"))<>"" then  floorArray(0) = NoSqlHack(request.Form("txt_Floor_1"))
	if NoSqlHack(request.Form("txt_Floor_2"))<>"" then  floorArray(1) = NoSqlHack(request.Form("txt_Floor_2"))
	s_Floor=floorArray(0)&","&floorArray(1)
	s_Position = NoSqlHack(request.Form("txt_Position"))
	s_Decoration=NoSqlHack(request.Form("sel_Decoration"))
	s_LinkMan=NoSqlHack(request.Form("txt_LinkMan"))
	s_Contact=NoSqlHack(request.Form("txt_Contact"))
	s_equip=NoSqlHack(request.Form("chk_equip"))
	s_Remark=NoSqlHack(request.Form("txt_Remark"))
	s_PubDate=DateValue(Now)
	s_Audited=0
	s_PicNumber=NoSqlHack(request.Form("PicNum"))
	'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	secondRs.open sqlstatement,Conn,1,3
	secondRs("Class")=s_Class
	secondRs("UserNumber")=s_UserNumber
	secondRs("label")=s_label
	secondRs("UseFor")=s_UseFor
	secondRs("FloorType")=s_FloorType
	secondRs("BelongType")=s_BelongType
	secondRs("HouseStyle")=s_HouseStyle
	secondRs("Structure")=s_Structure
	secondRs("Area")=s_Area
	if s_BuildDate="" then s_BuildDate=0	
	secondRs("BuildDate")=s_BuildDate
	if s_price="" then s_price=0
	secondRs("Price")=s_Price
	secondRs("CityArea")=s_CityArea
	secondRs("Address")=s_Address
	secondRs("Floor")=s_Floor
	secondRs("Position")=s_Position
	secondRs("LinkMan")=s_LinkMan
	secondRs("Contact")=s_Contact
	secondRs("equip")=s_equip
	secondRs("Decoration")=s_Decoration
	secondRs("Remark")=s_Remark
	secondRs("PubDate")=s_PubDate
	secondRs("Audited")=s_Audited
	s_picNumber = NoSqlHack(request.Form("txt_PicNum"))
	if s_picNumber="" then s_picNumber=0
	secondRs("picNumber")=s_picNumber
	secondRs.update
	if Cint(s_picNumber)>0 then
		for i=0 to Cint(s_picNumber)
			if trim(request.Form("txt_PicNum_"&(i+1)))<>"" then
				Conn.execute("Insert into FS_HS_Picture (ID,HS_Type,PIC) values("&secondRs("sID")&",3,'"&NoSqlHack(request.Form("txt_PicNum_"&(i+1)))&"')")
			End if
		next
	End if
	secondRs.close
End if
Conn.close
User_Conn.close
Set Conn=nothing
Set User_Conn=nothing
set secondRs=nothing
if err.number=0 then
	Response.Redirect("../lib/success.asp?ErrCodes=<li>操作成功</li>&ErrorURL=../House/HS_Second.asp")
	Response.End()
Else
	Response.Redirect("../lib/error.asp?ErrCodes=<li>请检查输入是否合法</li>")
	Response.End()
End if
%>







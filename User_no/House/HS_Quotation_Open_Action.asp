<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
  ' Powered by: FoosunCMS5.0系列,Company:Foosun Inc	
Dim house_rs,action,id,res,sqlstatement,i,pic_rs
Dim hs_HouseName,hs_Position,hs_Direction,hs_Class,hs_OpenDate,hs_PreSaleNumber,hs_IssueDate,hs_PreSaleRange,hs_Status,hs_Price,hs_PubDate,hs_tel,hs_UserNumber,hs_Audited,hs_editor,hs_picNumber,hs_introduction,hs_KaiFaShang
action=request.QueryString("action")
id=NoSqlHack(request("id"))
response.Charset="GB2312"
if action="delete" then
	Conn.execute("Delete from FS_HS_Quotation where id in("& FormatIntArr(id) &")")
	Conn.execute("Delete from FS_HS_Picture where id in("& FormatIntArr(id) &")")
	Response.Write("ok")
	response.End()
elseif action="add" then
	Set house_rs=Server.CreateObject(G_FS_RS)
	sqlstatement="select ID,HouseName,Position,Direction,Class,OpenDate,PreSaleNumber,IssueDate,PreSaleRange,Status,Price,PubDate,Tel,Click,UserNumber,Audited,editor,picNumber,introduction,KaiFaShang from FS_HS_Quotation"
	hs_HouseName = NoSqlHack(request.Form("txt_HouseName"))
	hs_Position = NoSqlHack(request.Form("txt_Position"))
	hs_KaiFaShang = NoSqlHack(request.Form("txt_KaiFaShang"))'---------2/13--chen
	hs_Direction = NoSqlHack(request.Form("txt_Direction"))
	hs_Class = NoSqlHack(request.Form("txt_Class"))
	hs_OpenDate = NoSqlHack(request.Form("txt_OpenDate"))
	hs_PreSaleNumber = NoSqlHack(request.Form("txt_PreSaleNumber"))
	hs_IssueDate = NoSqlHack(request.Form("txt_IssueDate"))
	hs_PreSaleRange = NoSqlHack(request.Form("txt_PreSaleRange"))
	hs_Status = NoSqlHack(request.Form("sel_Status"))
	hs_Price = NoSqlHack(request.Form("txt_Price"))
	hs_PubDate = NoSqlHack(request.Form("txt_PubDate"))
	hs_tel = NoSqlHack(request.Form("txt_tel"))
	hs_UserNumber=session("FS_UserNumber")
	hs_Audited=0
	hs_editor=Session("Admin_Name")
	hs_introduction=NoSqlHack(request.Form("txt_introduction"))
	house_rs.open sqlstatement,Conn,1,3
	house_rs.addnew
	house_rs("HouseName")=hs_HouseName
	house_rs("Position")=hs_Position
	house_rs("KaiFaShang")=hs_KaiFaShang'-------------2/13----------
	house_rs("Direction")=hs_Direction
	house_rs("Class")=hs_Class
	if hs_OpenDate<>"" then house_rs("OpenDate")=hs_OpenDate
	house_rs("PreSaleNumber")=hs_PreSaleNumber
	if hs_IssueDate<>"" then house_rs("IssueDate")=hs_IssueDate
	house_rs("PreSaleRange")=hs_PreSaleRange
	if hs_Status="" then hs_Status=1
	house_rs("Status")=hs_Status
	if hs_price="" then hs_price=0
	house_rs("Price")=hs_Price
	if hs_PubDate<>"" then house_rs("PubDate")=hs_PubDate
	house_rs("tel")=hs_tel
	house_rs("Click")=0
	house_rs("UserNumber")=hs_UserNumber
	house_rs("Audited")=hs_Audited
	house_rs("editor")=hs_editor
	house_rs("introduction")=hs_introduction
	hs_picNumber=NoSqlHack(request.Form("txt_PicNum"))
	if hs_picNumber="" then hs_picNumber=0
	house_rs("picNumber")=hs_picNumber
	house_rs.update
	if Cint(hs_picNumber)>0 then
		for i=0 to Cint(hs_picNumber)
			if trim(request.Form("txt_PicNum_"&(i+1)))<>"" then
				Conn.execute("Insert into FS_HS_Picture (ID,HS_Type,PIC) values("&house_rs("ID")&",1,'"&NoSqlHack(request.Form("txt_PicNum_"&(i+1)))&"')")
			End if
		next
	End if
	house_rs.close
elseif action="edit" then
	Set house_rs=Server.CreateObject(G_FS_RS)
	sqlstatement="select ID,HouseName,Position,Direction,Class,OpenDate,PreSaleNumber,IssueDate,PreSaleRange,Status,Price,PubDate,Tel,Click,UserNumber,Audited,editor,picNumber,introduction,KaiFaShang from FS_HS_Quotation where id="&CintStr(id)
	hs_HouseName = NoSqlHack(request.Form("txt_HouseName"))
	hs_Position = NoSqlHack(request.Form("txt_Position"))
	hs_KaiFaShang = NoSqlHack(request.Form("txt_KaiFaShang"))'----------------2/13----
	hs_Direction = NoSqlHack(request.Form("txt_Direction"))
	hs_Class = NoSqlHack(request.Form("txt_Class"))
	hs_OpenDate = NoSqlHack(request.Form("txt_OpenDate"))
	hs_PreSaleNumber = NoSqlHack(request.Form("txt_PreSaleNumber"))
	hs_IssueDate = NoSqlHack(request.Form("txt_IssueDate"))
	hs_PreSaleRange = NoSqlHack(request.Form("txt_PreSaleRange"))
	hs_Status = NoSqlHack(request.Form("sel_Status"))
	hs_Price = NoSqlHack(request.Form("txt_Price"))
	hs_PubDate = NoSqlHack(request.Form("txt_PubDate"))
	hs_tel = NoSqlHack(request.Form("txt_tel"))
	hs_UserNumber =session("FS_UserNumber")
	hs_Audited =0
	hs_editor =Session("Admin_Name")
	hs_introduction = NoSqlHack(request.Form("txt_introduction"))
	house_rs.open sqlstatement,Conn,1,3
	
	house_rs("HouseName")=hs_HouseName
	house_rs("Position")=hs_Position
	house_rs("KaiFaShang")=hs_KaiFaShang'----------------------2/13----
	house_rs("Direction")=hs_Direction
	house_rs("Class")=hs_Class
	if hs_OpenDate<>"" then house_rs("OpenDate")=hs_OpenDate
	house_rs("PreSaleNumber")=hs_PreSaleNumber
	if hs_IssueDate<>"" then house_rs("IssueDate")=hs_IssueDate
	house_rs("PreSaleRange")=hs_PreSaleRange
	if hs_Status="" then hs_Status=1
	house_rs("Status")=hs_Status
	if hs_price="" then hs_price=0
	house_rs("Price")=hs_Price
	if hs_PubDate<>"" then house_rs("PubDate")=hs_PubDate
	
	house_rs("tel")=hs_tel
	house_rs("Click")=0
	house_rs("UserNumber")=hs_UserNumber
	house_rs("Audited")=hs_Audited
	house_rs("editor")=hs_editor
	house_rs("introduction")=hs_introduction
	hs_picNumber = NoSqlHack(request.Form("txt_PicNum"))
	if hs_picNumber="" then hs_picNumber=0
	house_rs("picNumber")=hs_picNumber
	house_rs.update
	house_rs.close
	Conn.execute("Delete from FS_HS_Picture where  HS_type=1 and id="&id)
	if Cint(hs_picNumber)>0 then
		for i=0 to Cint(hs_picNumber)
			if NoSqlHack(request.Form("txt_PicNum_"&(i+1)))<>"" then
				Conn.execute("Insert into FS_HS_Picture (ID,HS_Type,PIC) values("&CintStr(id)&",1,'"&NoSqlHack(request.Form("txt_PicNum_"&(i+1)))&"')")
			End if
		next
	End if
End if
User_Conn.close
Conn.close
Set pic_rs=nothing
Set house_rs=nothing
Set Conn=nothing
Set User_Conn=nothing
if err.number=0 then
	Response.Write("ok")
	Response.Redirect("lib/success.asp?ErrCodes=<li>操作成功</li>&ErrorURL=../houseManage.asp")
	Response.End()
Else
	Response.Redirect("lib/error.asp?ErrCodes=<li>请检查输入是否合法</li>")
	Response.End()
End if

%>







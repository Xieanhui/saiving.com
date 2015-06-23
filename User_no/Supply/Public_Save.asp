<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
	dim p_s_id,p_s_ClassID,p_s_MyClassID,p_s_PubTitle,p_s_PubType,p_s_PubContent,p_s_Keyword,p_s_ValidTime,p_s_CompType,p_s_PubNumber,p_s_PubPrice,p_s_PubPack,p_s_Pubgui,p_s_PubPic_1,p_s_PubPic_2,p_s_PubPic_3,p_s_PubAddress,obj_save_rs
	dim obj_sys_rs,sys_islock,p_s_AreaID
	Dim Info_Temp
	p_s_ClassID = NoSqlHack(Request.Form("ClassID"))
	p_s_MyClassID = NoSqlHack(Request.Form("MyClassID"))
	p_s_PubTitle = NoSqlHack(Request.Form("PubTitle"))
	p_s_PubType = NoSqlHack(Request.Form("PubType"))
	p_s_PubContent = NoSqlHack(Request.Form("PubContent"))
	p_s_Keyword = NoSqlHack(Request.Form("keyword1"))&","&NoSqlHack(Request.Form("keyword2"))&","&NoSqlHack(Request.Form("keyword3"))
	p_s_CompType = NoSqlHack(Request.Form("CompType"))
	p_s_PubNumber = CintStr(Request.Form("PubNumber"))
	p_s_PubPrice = CintStr(Request.Form("PubPrice"))
	p_s_PubPack = NoSqlHack(Request.Form("PubPack"))
	p_s_Pubgui = NoSqlHack(Request.Form("Pubgui"))
	p_s_PubPic_1 = NoSqlHack(Request.Form("pic_1"))
	p_s_PubPic_2 = NoSqlHack(Request.Form("pic_2"))
	p_s_PubPic_3 = NoSqlHack(Request.Form("pic_3"))
	p_s_ValidTime = NoSqlHack(Request.Form("ValidTime")) 
	p_s_PubAddress = NoSqlHack(Request.Form("PubAddress"))
	p_s_AreaID = NoSqlHack(Request.Form("AreaID"))
	p_s_id =  NoSqlHack(Request.Form("id"))
	set obj_sys_rs = Conn.execute("select top 1 islock,m_islock,s_Templet from FS_SD_Config")
	if obj_sys_rs.eof then
			Info_Temp = replace(G_VIRTUAL_ROOT_DIR&"/"&G_TEMPLETS_DIR&"/Supply/list.htm","//","/")
			strShowErr = "<li>找不到系统配置，请与管理员联系</li>"
			Response.Redirect(""& s_savepath &"/lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	Else
		Info_Temp = obj_sys_rs(2)
	end if
	if trim(p_s_PubNumber)<>"" or trim(p_s_PubNumber)<>empty then
		if Not isnumeric(p_s_PubNumber) then
			strShowErr = "<li>产品数量应该为数字</li>"
			Response.Redirect(""& s_savepath &"/lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
	end if
	if clng(p_s_PubNumber)>=10000 then
			strShowErr = "<li>产品数量应该在10000以下</li>"
			Response.Redirect(""& s_savepath &"/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	end if
	if Not isnumeric(p_s_PubPrice)  then
		strShowErr = "<li>产品价格应该为数字</li>"
		Response.Redirect(""& s_savepath &"/lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	if len(p_s_PubContent)>5000 then
		strShowErr = "<li>信息描述不能超过5000个字符</li>"
		Response.Redirect(""& s_savepath &"/lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	if not isnumeric(p_s_ValidTime) then
		strShowErr = "<li>有效期不为数字</li>"
		Response.Redirect(""& s_savepath &"/lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	if int(p_s_ValidTime)<1 or int(p_s_ValidTime)>360 then
		strShowErr = "<li>有效期为1~360</li>"
		Response.Redirect(""& s_savepath &"/lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	set obj_save_rs = Server.CreateObject(G_FS_RS)
	if Request.Form("Action")="add" or Request.Form("Action")="Copy" then
		obj_save_rs.open "select * from FS_SD_News where 1=0",Conn,1,3
		obj_save_rs.addnew
		obj_save_rs("Addtime")=now
	elseif  Request.Form("Action")="edit" then
		obj_save_rs.open "select * from FS_SD_News where id="&CintStr(p_s_id),Conn,1,3
	end if
	obj_save_rs("EditTime")=now
	obj_save_rs("ClassID")=p_s_ClassID
	if trim(p_s_MyClassID)<>"" then:obj_save_rs("MyClassID")=p_s_MyClassID:else:obj_save_rs("MyClassID")=0:end if
	obj_save_rs("UserNumber")=session("FS_UserNumber")
	obj_save_rs("PubTitle")=p_s_PubTitle
	if p_s_PubType = "0" then
		obj_save_rs("PubType")=0
	elseif  p_s_PubType ="1" then
		obj_save_rs("PubType")=1
	elseif  p_s_PubType ="2" then
		obj_save_rs("PubType")=2
	elseif  p_s_PubType ="3" then
		obj_save_rs("PubType")=3
	else
		obj_save_rs("PubType")=4
	end if
	obj_save_rs("PubContent")=p_s_PubContent
	obj_save_rs("Keyword")=p_s_Keyword
	if p_s_CompType="0" then
		obj_save_rs("CompType")=0
	elseif  p_s_CompType="1" then
		obj_save_rs("CompType")=1
	else
		obj_save_rs("CompType")=2
	end if
	if p_s_PubNumber<>"" then:obj_save_rs("PubNumber")=p_s_PubNumber:else:obj_save_rs("PubNumber")=0:end if
	obj_save_rs("PubPrice")=p_s_PubPrice
	obj_save_rs("PubPack")=p_s_PubPack
	obj_save_rs("Pubgui")=p_s_Pubgui
	if p_s_PubPic_1<>"" then  obj_save_rs("PubPic_1")=p_s_PubPic_1
	if p_s_PubPic_2<>"" then  obj_save_rs("PubPic_2")=p_s_PubPic_2
	if p_s_PubPic_3<>"" then  obj_save_rs("PubPic_3")=p_s_PubPic_3
	obj_save_rs("ValidTime")=p_s_ValidTime
	if Request.Form("Action")="edit" then
		if obj_sys_rs("m_islock")=1 then
			obj_save_rs("isPass")=0
		end if
	elseif Request.Form("Action")="add" then
		if obj_sys_rs("islock")=1 then
			obj_save_rs("isPass")=0
		else
			obj_save_rs("isPass")=1
		end if
	end if
	obj_save_rs("AdminUserName")="0"
	obj_save_rs("s_Templet")=Info_Temp
	obj_save_rs("FileExtName")="html"
	obj_save_rs("HideTF")=0
	obj_save_rs("OrderID")=0
	obj_save_rs("Fax")=NoSqlHack(Request.Form("Fax"))
	obj_save_rs("tel")=NoSqlHack(Request.Form("tel"))
	obj_save_rs("Mobile")=NoSqlHack(Request.Form("Mobile"))
	obj_save_rs("PubAddress")=p_s_PubAddress
	obj_save_rs("AreaID")=p_s_AreaID
	obj_save_rs.update
	obj_save_rs.close:set obj_save_rs = nothing
	obj_sys_rs.close:set obj_sys_rs = nothing
	if Request.Form("Action")="add" then
		strShowErr = "<li>添加供求信息成功</li><br><li><a href="&s_savepath&"/supply/PublicSupply.asp>继续添加</a>&nbsp;&nbsp;<a href="&s_savepath&"/supply/PublicManage.asp>返回</a></li>"
		'--------------2006-12-29 发布信息成功后扣除点数和金币数  by ken
		Dim MustPoint,MustMoney,Get_RulerRs,MustAuditTF
		Set Get_RulerRs = Conn.ExeCute("Select Top 1 PublicPoint,PublicMoney,isLock From FS_SD_Config Where ID > 0 Order By ID")
		If Get_RulerRs.Eof Then
			MustPoint = 0
			MustMoney = 0
			MustAuditTF = 0
		Else
			MustPoint = Clng(Get_RulerRs(0))	
			MustMoney = Clng(Get_RulerRs(1))
			MustAuditTF = Cint(Get_RulerRs(2))
		End if
		Get_RulerRs.Close : Set Get_RulerRs = Nothing
		If MustAuditTF = 0 Then
			User_Conn.ExeCute("Update Fs_Me_Users Set Integral = Integral - "&MustPoint&",FS_Money = FS_Money - "&MustMoney&" Where UserID = " & Fs_User.UserID)
		End If	
		'--------------------------------------------------------------
	elseif Request.Form("Action")="edit" then
		strShowErr = "<li>修改供求信息成功</li><br><li><a href="&s_savepath&"/supply/PublicSupplyEdit.asp?Id="& p_s_id &">继续修改</a>&nbsp;&nbsp;<a href="&s_savepath&"/supply/PublicManage.asp>返回</a></li>"
	else
		strShowErr = "<li>复制信息成功</li><br><li>&nbsp;&nbsp;<a href="&s_savepath&"/supply/PublicManage.asp>返回</a></li>"
	end if
	Response.Redirect(""& s_savepath &"/lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="& s_savepath &"/Supply/PublicManage.asp")
	Response.end
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






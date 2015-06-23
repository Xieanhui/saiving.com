<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%
	Dim Conn,User_Conn
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF 
	Dim Fs_news
	set Fs_news = new cls_news
	if not Get_SubPop_TF(NoSqlHack(Request("ClassID")),"NS016","NS","class") then Err_Show
	If Not Fs_news.IsSelfRefer Then response.write "非法提交数据":Response.end
	'获取参数
	Dim obj_Save_Rs1,ClassID1,str_ClassKeywords1,str_Classdescription1,str_action1,strShowErr
	Dim str_ClassID1,lng_OrderID1,str_ClassName1,str_ClassEName1,str_ParentID1,str_Templet1,str_NewsTemplet1,str_Domain1,lng_AdminID1,lng_RefreshNumber
	Dim  lng_GroupID1,lng_PointNumber1,flt_Money1,str_FileExtName1,dtm_Addtime1,int_isConstr1,int_IsURL1,str_UrlAddress1,lng_Oldtime1,int_isShow1
	Dim str_ClassNaviContent1,str_ClassNaviPic1,lng_DefineID1,int_NewsCheck1,int_AddNewsType1,str_SavePath1,str_FileSaveType1,int_isConstrDel1,str_GetParentID,IsAdPic,AdPicWH,AdPicLink,AdPicAdress
	
	str_action1 = NoSqlHack(Request.Form("str_add"))
	str_ClassID1 = NoSqlHack(Request.Form("ClassID"))
	lng_OrderID1 = NoSqlHack(Request.Form("OrderID"))
	str_ClassName1 = NoSqlHack(Request.Form("ClassName"))
	str_ClassEName1 = NoSqlHack(Trim(Request.Form("ClassEName")))
	str_ParentID1 = NoSqlHack(Request.Form("ParentID"))
	str_Templet1 = NoSqlHack(Request.Form("Templet"))
	str_NewsTemplet1 = NoSqlHack(Request.Form("NewsTemplet"))
	str_Domain1 = NoSqlHack(Request.Form("Domain"))
	lng_AdminID1 = NoSqlHack(Request.Form("ClassAdmin"))
	lng_RefreshNumber = NoSqlHack(Request.Form("RefreshNumber"))
	lng_GroupID1 = NoSqlHack(Request.Form("BrowPop"))
	lng_PointNumber1 = NoSqlHack(Request.Form("PointNumber")) 
	flt_Money1 = NoSqlHack(Request.Form("Money"))
	str_FileExtName1 = NoSqlHack(Request.Form("FileExtName"))
	dtm_Addtime1 = NoSqlHack(Request.Form("Addtime"))
	int_isConstr1 = NoSqlHack(Request.Form("isConstr"))
	int_IsURL1 = NoSqlHack(Request.Form("IsURL"))
	str_UrlAddress1 = NoSqlHack(Request.Form("UrlAddress"))
	lng_Oldtime1 = NoSqlHack(Request.Form("Oldtime"))
	int_isShow1 = NoSqlHack(Request.Form("isShow"))
	
	str_ClassNaviContent1 = NoSqlHack(Request.Form("ClassNaviContent"))
	str_ClassNaviPic1 = NoSqlHack(Request.Form("ClassNaviPic"))
	lng_DefineID1 = NoSqlHack(Request.Form("DefineID"))
	int_NewsCheck1 = NoSqlHack(Request.Form("NewsCheck"))
	int_AddNewsType1 = NoSqlHack(Request.Form("AddNewsType"))
	if  Trim(Request.Form("SavePath")) = "" then
		str_SavePath1 = "/"
	Else
		str_SavePath1 = NoSqlHack(Request.Form("SavePath"))
	End if
	str_FileSaveType1 = NoSqlHack(Request.Form("FileSaveType"))
	int_isConstrDel1 = NoSqlHack(Request.Form("isConstrDel"))
	str_ClassKeywords1  = NoSqlHack(Request.Form("ClassKeywords"))
	str_Classdescription1  = NoSqlHack(Request.Form("Classdescription"))
	
	if int_IsURL1="2" then
		str_ClassEName1 = NoSqlHack(Request.Form("ClassEName"))
		str_Templet1 = NoSqlHack(Request.Form("Templet"))
		str_Domain1 = NoSqlHack(Request.Form("Domain"))
		lng_AdminID1 = NoSqlHack(Request.Form("ClassAdmin"))
		str_FileExtName1 = NoSqlHack(Request.Form("FileExtName"))
		str_FileSaveType1 = NoSqlHack(Request.Form("FileSaveType"))
		if  Trim(Request.Form("SavePath")) = "" then
			str_SavePath1 = "/"
		Else
			str_SavePath1 = NoSqlHack(Request.Form("SavePath"))
		End if
		str_ClassKeywords1 = NoSqlHack(Request.Form("ClassKeywords"))
		str_Classdescription1 = NoSqlHack(Request.Form("Classdescription"))
	end if
		
	IsAdPic = CintStr(Request.Form("IsAdPic"))
	AdPicWH = NoSqlHack(Request.Form("AdPicWH"))
	AdPicLink = NoSqlHack(Request.Form("AdPicLink"))
	AdPicAdress =  NoSqlHack(Request.Form("AdPicAdress"))
	
	If IsAdPic=1 and  Cintstr(Request.Form("Checkbox2"))=1 Then       	 
		If not IsNumeric(Request.Form("AdPicWHw")) Then
	   		strShowErr = "<li>插入位置格式有误，应为正整数</li>"
		    Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		   	Response.end			
		end if
		if Request.Form("IsApicArea")="" then 
		    strShowErr = "<li>文字画中画代码不能为空！</li>"
		    Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		    Response.end
		End If
		if len(Request.Form("IsApicArea"))>250 then
		    strShowErr = "<li>文字画中画代码不能超过250个字符！</li>"
		    Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		    Response.end
		end if
		IsAdPic = 2
		AdPicWH = NoSqlHack(Request.Form("AdPicWHw"))
		AdPicLink = NoSqlHack(Request.Form("IsApicArea"))
	end if
	'判断数据是否正确
	if str_Domain1 <>"" then
		if len(Trim(str_Domain1))<6  then
			strShowErr = "<li>请正确填写您的二级域名</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
	End if
	if trim(int_IsURL1) = "1" then 
		if isnull(Trim(str_UrlAddress1))  then
			strShowErr = "<li>请填写外部地址</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		if Trim(str_UrlAddress1)="http://" then
			strShowErr = "<li>请填写外部地址</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		if trim(str_ClassName1) = ""  then
			strShowErr = "<li>请填写外部栏目名称</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
	Else
			if trim(str_SavePath1) = ""  then
				strShowErr = "<li>请填写栏目保存路径</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			if trim(str_ClassName1) = "" or trim(str_ClassEName1) = ""  or trim(str_Templet1) = ""  or trim(str_NewsTemplet1) = ""  or trim(str_SavePath1) = ""  then
				strShowErr = "<li>带*的是必须填写的</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			if isdate(dtm_Addtime1) =false then
				strShowErr = "<li>请填写正确的日期格式</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			if isnumeric(lng_Oldtime1) =false or isnumeric(lng_RefreshNumber) = false then
				strShowErr = "<li>归档日期不是正确的数字</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			if isnumeric(lng_OrderID1) = false then
				strShowErr = "<li>排列权重(序号)不是正确的数字</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			if fs_news.chkinputchar(str_ClassEName1) = false then
				strShowErr = "<li>英文名称只能为英文、数字及下划线</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			if trim(lng_GroupID1)<>"" or lng_PointNumber1 <>"" or flt_Money1<>"" then 
				if trim(str_FileExtName1)<>"asp" then
						strShowErr = "<li>您设置了浏览权限，扩展名必须为.asp</li>"
						Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
						Response.end
				End if
			End if
			If IsAdPic=1 Then
				If AdPicWH="" or IsNull(AdPicWH) Then
					strShowErr = "<li>请填写图片高度与宽度</li>"
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
				End If
				if instr(AdPicWH,",")=0 or instr(AdPicWH,"，")>0 then
					strShowErr = "<li>图片高度与宽度填写错误</li>"
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
				end if
		 	    If Ubound(split(AdPicWH,","))<>3 Then
					strShowErr = "<li>图片高度与宽度,显示布局格式有误,插入位置有误，格式为100,200,1,400</li>"
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
				End If
				If Not IsNumeric(split(AdPicWH,",")(3)) Then
				   strShowErr = "<li>插入位置格式有误，应为正整数</li>"
				   Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				   Response.end
				End If
			    If Cint(split(AdPicWH,",")(3))<0 Then 
				   strShowErr = "<li>插入位置格式有误，应为正整数</li>"
				   Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				   Response.end
			    End If
				If AdPicAdress="" Or IsNull(AdPicAdress) Then
					strShowErr = "<li>图片地址不能为空</li>"
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
				End If
			End If
	End if
	Dim GetClassReturnValue1,obj_SaveTF_Rs1,obj_TF_Rs1,Newsadd_SQL1,NewsaddTF_SQL1
	GetClassReturnValue1 = Fs_news.GetRamCode(15)
	Set obj_Save_Rs1 = server.CreateObject(G_FS_RS)
	If str_action1 = "add" then
		Set obj_SaveTF_Rs1 = server.CreateObject(G_FS_RS)
		obj_SaveTF_Rs1.Open "Select ID from FS_NS_NewsClass where ClassID='"& GetClassReturnValue1 &"' order by id desc",Conn,1,3
		if  Not obj_SaveTF_Rs1.eof then
				strShowErr = "<li>栏目ClassID意外出现重复，请重新输入</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
		End if
		set obj_SaveTF_Rs1 = nothing
		Set obj_TF_Rs1 = server.CreateObject(G_FS_RS)
		If str_ParentID1="0" Then 
			NewsaddTF_SQL1 ="Select ID from FS_NS_NewsClass where ClassEName='"& NoSqlHack(trim(str_ClassEName1)) &"' And ParentID='0'"
		Else
			NewsaddTF_SQL1 ="Select ID from FS_NS_NewsClass where ParentID='"&NoSqlHack(str_ParentID1)&"' and ClassEName='"& NoSqlHack(trim(str_ClassEName1)) &"'"
		End If
		obj_TF_Rs1.Open NewsaddTF_SQL1,Conn,1,3
		if Not (obj_TF_Rs1.eof and obj_TF_Rs1.bof)  then
			if trim(int_IsURL1)="" then
				strShowErr = "<li>栏目英文名称重复，请重新输入.或者您的回收站中存在此目录</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			end if
		End if
		set obj_TF_Rs1 = nothing
		Newsadd_SQL1 = "Select * from FS_NS_NewsClass where 1=2"
		obj_Save_Rs1.Open Newsadd_SQL1,Conn,1,3
		obj_Save_Rs1.AddNew
		obj_Save_Rs1("ClassEName") = str_ClassEName1
		obj_Save_Rs1("ClassID") = GetClassReturnValue1
		obj_Save_Rs1("ParentID") = str_ParentID1
	ElseIf str_action1 = "edit" then
		Newsadd_SQL1 = "Select * from FS_NS_NewsClass where ClassID='"& NoSqlHack(str_ClassID1) &"'"
		obj_Save_Rs1.Open Newsadd_SQL1,Conn,1,3
	End if
		if int_isShow1 <>"" then:obj_Save_Rs1("isShow") = 1:Else:obj_Save_Rs1("isShow") = 0:End if
		obj_Save_Rs1("Addtime") = dtm_Addtime1
		obj_Save_Rs1("RefreshNumber") = lng_RefreshNumber
		obj_Save_Rs1("ClassNaviContent") = str_ClassNaviContent1
		obj_Save_Rs1("ClassNaviPic") = str_ClassNaviPic1
		if lng_OrderID1<>"" then:obj_Save_Rs1("OrderID") = clng(lng_OrderID1):Else:obj_Save_Rs1("OrderID") = 10:End if
		if int_IsURL1 = "1" then
			obj_Save_Rs1("IsURL") = int_IsURL1
			obj_Save_Rs1("UrlAddress") = str_UrlAddress1
			obj_Save_Rs1("ClassName") = str_ClassName1
		Else
			obj_Save_Rs1("ClassName") = str_ClassName1
			obj_Save_Rs1("Templet") = str_Templet1
			obj_Save_Rs1("NewsTemplet") = str_NewsTemplet1
			obj_Save_Rs1("Domain") = str_Domain1
			obj_Save_Rs1("ClassAdmin") = lng_AdminID1
			obj_Save_Rs1("FileExtName") = str_FileExtName1
			if int_isConstr1 <>"" then:obj_Save_Rs1("isConstr") = 1:Else:obj_Save_Rs1("isConstr") = 0:End if
			obj_Save_Rs1("IsURL") = int_IsURL1
			obj_Save_Rs1("UrlAddress") = ""
			obj_Save_Rs1("Oldtime") = clng(lng_Oldtime1)
			obj_Save_Rs1("DefineID") = lng_DefineID1
			if int_NewsCheck1 <> "" then:obj_Save_Rs1("NewsCheck") = 1:Else:obj_Save_Rs1("NewsCheck") = 0:End if
			if int_AddNewsType1 <>"" then:obj_Save_Rs1("AddNewsType") = 0:Else:obj_Save_Rs1("AddNewsType") = 1:End if
			obj_Save_Rs1("SavePath") = str_SavePath1
			obj_Save_Rs1("FileSaveType") = str_FileSaveType1
			if int_isConstrDel1 <>"" then:obj_Save_Rs1("isConstrDel") = 1:Else:obj_Save_Rs1("isConstrDel") = 0:End if
			if Trim(lng_GroupID1) <>"" or lng_PointNumber1 <> "" or flt_Money1<>"" then:obj_Save_Rs1("isPop") = 1:Else:obj_Save_Rs1("isPop") = 0:End if
			obj_Save_Rs1("ClassKeywords") = str_ClassKeywords1
			obj_Save_Rs1("Classdescription") = str_Classdescription1
			If IsAdPic=1 or IsAdPic=2 Then
				obj_Save_Rs1("IsAdPic")=IsAdPic
				obj_Save_Rs1("AdPicWH")=AdPicWH
				obj_Save_Rs1("AdPicLink")=AdPicLink
				obj_Save_Rs1("AdPicAdress")=AdPicAdress
			Else
				obj_Save_Rs1("IsAdPic")=0
			End If
		End if
		'插入权限数据表
	'	lng_GroupID1,lng_PointNumber1,flt_Money1
		if Trim(lng_GroupID1) <>"" or lng_PointNumber1 <> "" or flt_Money1<>"" then 
			Dim obj_insert_rs
			set obj_insert_rs = Server.CreateObject(G_FS_RS)
			If str_action1 = "add" then
				obj_insert_rs.Open "select  GroupName,PointNumber,FS_Money,InfoID,PopType,isClass From FS_MF_POP",Conn,1,3
				obj_insert_rs.addnew
				obj_insert_rs("InfoID")=GetClassReturnValue1
			elseIf str_action1 = "edit" then
				obj_insert_rs.Open "select  GroupName,PointNumber,FS_Money,InfoID,PopType,isClass From FS_MF_POP  where InfoID='"& NoSqlHack(str_ClassID1) &"' and PopType='NS' and isClass=1",Conn,1,3
				If obj_insert_rs.eof Then
					obj_insert_rs.addnew
				End If
				obj_insert_rs("InfoID")=str_ClassID1
			End if
			obj_insert_rs("GroupName")=lng_GroupID1
			if lng_PointNumber1 <>""  then:obj_insert_rs("PointNumber")=lng_PointNumber1:Else:obj_insert_rs("PointNumber")=0:End if
			if flt_Money1 <>"" then:obj_insert_rs("FS_Money")=flt_Money1:Else:obj_insert_rs("FS_Money")=0:End if
			obj_insert_rs("PopType")="NS"
			obj_insert_rs("isClass")=1
			obj_insert_rs.update
			obj_insert_rs.close:set obj_insert_rs = nothing
		End if
	obj_Save_Rs1.update
	obj_Save_Rs1.close
	set obj_Save_Rs1 = nothing
	Call Makexml(str_ParentID1)
	Call Makexml("0")
	strShowErr = "<li>恭喜，栏目保存成功</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Class_Manage.asp")
	Response.end
	set Fs_news = nothing 
%>







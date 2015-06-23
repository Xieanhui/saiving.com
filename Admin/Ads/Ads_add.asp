<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="Cls_Ads.asp"-->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("AS_site") then Err_Show
if not MF_Check_Pop_TF("AS001") then Err_Show
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_
Dim Page,cPageNo,str_CurrPath,sRootDir,str_AdOpType,str_ShowType,str_AdsType
str_AdOpType=Request.QueryString("OpType")
Dim o_Ad_Rs,strShowErr,o_Crs,strLock,lng_TempAdID,lng_TempLoopAdID,G_Ads_FILES_DIR

Dim Str_SysDir
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
Str_SysDir=""
if G_VIRTUAL_ROOT_DIR<>"" then
	Str_SysDir="/"&G_VIRTUAL_ROOT_DIR&"/Ads"
Else
	Str_SysDir="/Ads"
end if
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if

If str_AdOpType="Add" Then
	str_AdOpType="Add"
	str_ShowType="<a href=""javascript:if (Ad_Flag()==true){Ad_Save();}"">保存</a>"
Else
	str_AdOpType="Update"
	str_ShowType="<a href=""javascript:if (Ad_Flag()==true){Ad_Update();}"">修改</a>"
	Dim AdId
	AdId=Request.QueryString("ID")

	If AdId="" or IsNull(AdId) Then
		If IsNumeric(AdId)=False Then
			strShowErr = "<li>参数错误!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
	Else
		AdId=CintStr(AdId)
		Set o_Ad_Rs=Conn.execute("Select * from FS_AD_Info Where AdID="&CintStr(AdID)&"")
		If Not o_Ad_Rs.Eof Then
			Dim temp_Adname,temp_adType,temp_adloopadname,temp_adloop,temp_adloopfollow,temp_adloopspeed,temp_adloopic,temp_adloopRpic,temp_picH,temp_picW,temp_adlink,temp_adcaptiontxt,temp_loopfactor,temp_loopenddate,temp_maxclicknum,temp_maxshownum,temp_adremarks,Temp_AdID,temp_Txtcontentstr,temp_Txt_Rs,temp_Txt_i,temp_AdTxtColNum
			Temp_AdID=o_Ad_Rs("AdID")
			temp_Adname=o_Ad_Rs("AdName")
			temp_adType=o_Ad_Rs("AdType")
			temp_adloop=o_Ad_Rs("AdLoop")
			temp_adloopadname=o_Ad_Rs("AdLoopAdID")
			temp_adloopfollow=o_Ad_Rs("AdLoopFollow")
			temp_adloopspeed=o_Ad_Rs("AdLoopSpeed")
			temp_adloopic=o_Ad_Rs("AdPicPath")
			temp_adloopRpic=o_Ad_Rs("AdRightPicPath")
			temp_picH=o_Ad_Rs("AdPicHeight")
			temp_picW=o_Ad_Rs("AdPicWidth")
			temp_adlink=o_Ad_Rs("AdLinkAdress")
			temp_adcaptiontxt=o_Ad_Rs("AdCaptionTxt")
			temp_loopfactor=o_Ad_Rs("AdLoopFactor")
			temp_loopenddate=o_Ad_Rs("AdEndDate")
			temp_maxclicknum=o_Ad_Rs("AdMaxClickNum")
			temp_maxshownum=o_Ad_Rs("AdMaxShowNum")
			temp_adremarks=o_Ad_Rs("AdRemarks")
			temp_AdTxtColNum=o_Ad_Rs("AdTxtColNum")
		Else
			strShowErr = "<li>参数错误!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end		
		End If
		Set o_Ad_Rs=Nothing
	End If
End If

Dim str_AdName,str_AdType,str_AdLoopAdName,str_AdLoopFollow,str_AdLoopSpeed,str_AdLoopPicAdress,str_AdLoopRPicAdress,str_AdPicHeight,str_AdPicWidth,str_AdLinkUrl,str_AdCaptionTxt,str_LoopFactor,str_LoopEndDate,str_AdClickNum,str_AdShowNum,str_AdRemarks,str_IsLoopvalue,str_AdClassID,int_Txt_i,int_AdTxtColNum
	
str_AdName=NoSqlHack(Request.Form("AdName"))
str_AdType=Clng(Request.Form("AdType"))
str_IsLoopvalue=NoSqlHack(Request.Form("IsLoopvalue"))
str_AdLoopAdName=NoSqlHack(Request.Form("LoopAdName"))
str_AdLoopFollow=Clng(Request.Form("LoopFollow"))
str_AdLoopSpeed=NoSqlHack(Request.Form("LoopSpeed"))
str_AdLoopPicAdress=NoSqlHack(Request.Form("LoopPicAdress"))
str_AdLoopRPicAdress=NoSqlHack(Request.Form("LoopRPicAdress"))
str_AdPicHeight=NoSqlHack(Request.Form("AdPicHeight"))
str_AdPicWidth=NoSqlHack(Request.Form("AdPicWidth"))
str_AdLinkUrl=NoSqlHack(Request.Form("AdLinkUrl"))
str_AdCaptionTxt=NoSqlHack(Request.Form("AdCaptionTxt"))
str_LoopFactor=Clng(Request.Form("LoopFactor"))
str_LoopEndDate=NoSqlHack(Request.Form("LoopEndDate"))
str_AdClickNum=NoSqlHack(Request.Form("AdMaxClickNum"))
str_AdShowNum=NoSqlHack(Request.Form("AdMaxShowNum"))
str_AdRemarks=NoSqlHack(Request.Form("AdRemarks"))
str_AdClassID=NoSqlHack(Request.Form("AdClassID"))
int_AdTxtColNum=NoSqlHack(Request.Form("AdTxtColNum"))
If str_AdClickNum = "" Or Not IsNumeric(str_AdClickNum) Then str_AdClickNum = 0
If str_AdShowNum = "" Or Not IsNumeric(str_AdShowNum) Then str_AdShowNum = 0
If int_AdTxtColNum="" or IsNull(int_AdTxtColNum) Then
	int_AdTxtColNum=0
Else
	int_AdTxtColNum=Cint(int_AdTxtColNum)
End If
Dim SubUp,UpRs,ID
SubUp=Request.QueryString("Submit")

If SubUp="SubUp" Then
	ID=Request.QueryString("ID")
	Set UpRs=CreateObject(G_FS_RS)
	UpRs.open "Select * from FS_AD_Info Where AdID="&Clng(ID)&"",Conn,3,3
		'Set o_Crs=Conn.execute("Select Lock From FS_AD_Class Where AdClassID="&str_AdClassID&"")
		'If Not o_Crs.Eof Then
			'strLock=o_Crs("Lock")
		'Else
			'strLock=0
		'End If

		UpRs("AdName")=NoSqlHack(str_AdName)
		UpRs("AdType")=CintStr(str_AdType)
		UpRs("AdLoop")=CintStr(str_IsLoopvalue)
		UpRs("AdLoopAdID")=CintStr(str_AdLoopAdName)
		UpRs("AdLoopFollow")=CintStr(str_AdLoopFollow)
		UpRs("AdLoopSpeed")=CintStr(str_AdLoopSpeed)
		UpRs("AdPicPath")=NoSqlHack(str_AdLoopPicAdress)
		UpRs("AdRightPicPath")=str_AdLoopRPicAdress
		UpRs("AdPicHeight")=CintStr(str_AdPicHeight)
		UpRs("AdPicWidth")=CintStr(str_AdPicWidth)
		UpRs("AdLinkAdress")=NoSqlHack(str_AdLinkUrl)
		UpRs("AdCaptionTxt")=NoSqlHack(str_AdCaptionTxt)
		UpRs("AdLoopFactor")=CintStr(str_LoopFactor)
		UpRs("AdEndDate")=NoSqlHack(str_LoopEndDate)
		UpRs("AdMaxClickNum")=CintStr(str_AdClickNum)
		UpRs("AdMaxShowNum")=CintStr(str_AdShowNum)
		UpRs("AdRemarks")=NoSqlHack(str_AdRemarks)
		UpRs("AdClassID")=NoSqlHack(str_AdClassID)
		UpRs("AdTxtColNum")=NoSqlHack(int_AdTxtColNum)
		'If strLock=0 Then
			'UpRs("AdLock")=0
		'Else
			'UpRs("AdLock")=1
		'End If
	UpRs.Update
	UpRs.Close
	Set UpRs=Nothing
	
	Conn.execute("delete  From FS_AD_TxtInfo Where AdID="&CintStr(ID)&"")
	For int_Txt_i=1 to Request.Form("AdTxtContent").Count
		Conn.execute("insert into FS_AD_TxtInfo(AdID,AdTxtContent,Css,LinkUrl) values("&CintStr(ID)&",'"&NoSqlHack(Request.Form("AdTxtContent")(int_Txt_i))&"','"&NoSqlHack(Request.Form("AdTxtCss")(int_Txt_i))&"','"&NoSqlHack(Request.Form("AdTxtLink")(int_Txt_i))&"')")
	Next
	
	Set o_Ad_Rs=Conn.execute("Select AdLoopAdID From FS_AD_Info Where AdID="&CintStr(ID)&"")
		If Not o_Ad_Rs.Eof Then 
			lng_TempLoopAdID=o_Ad_Rs("AdLoopAdID")
		Else
			lng_TempLoopAdID=0
		End If
	Set o_Ad_Rs=Nothing

	Select Case CintStr(str_AdType)
		Case 0 call ShowAds(ID)
		Case 1 call NewWindow(ID)
		Case 2 call OpenWindow(ID)
		Case 3 call FilterAway(ID)
		Case 4 call DialogBox(ID)
		Case 5 call ClarityBox(ID)
		Case 6 call DriftBox(ID)
		Case 7 call LeftBottom(ID)
		Case 8 call RightBottom(ID)
		Case 9 call Couplet(ID)
		Case 10 call Cycle(ID,lng_TempLoopAdID)
		Case 11 call AdTxt(ID)
	End Select
	strShowErr = "<li>修改成功!</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Manage.asp&Page="&Request.QueryString("OpPage")&"")
	Response.end

End If

if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if

str_AdsType=Request.Form("Type")
If str_AdsType<>"" Then	
	Set o_Ad_Rs= CreateObject(G_FS_RS)
	o_Ad_Rs.open"select * from FS_AD_Info where 1=0",Conn,3,3
	If str_AdsType="Add" Then
		Set o_Crs=Conn.execute("Select Lock From FS_AD_Class Where AdClassID="&CintStr(str_AdClassID&""))
		If Not o_Crs.Eof Then
			strLock=o_Crs("Lock")
		Else
			strLock=0
		End If
		Set o_Crs=Nothing

		o_Ad_Rs.addnew
		o_Ad_Rs("AdName")=str_AdName
		o_Ad_Rs("AdType")=CintStr(str_AdType)
		o_Ad_Rs("AdLoop")=CintStr(str_IsLoopvalue)
		o_Ad_Rs("AdLoopAdID")=CintStr(str_AdLoopAdName)
		o_Ad_Rs("AdLoopFollow")=CintStr(str_AdLoopFollow)
		o_Ad_Rs("AdLoopSpeed")=CintStr(str_AdLoopSpeed)
		o_Ad_Rs("AdPicPath")=str_AdLoopPicAdress
		o_Ad_Rs("AdRightPicPath")=str_AdLoopRPicAdress
		If str_AdPicHeight="" or IsNull(str_AdPicHeight) Then
			o_Ad_Rs("AdPicHeight")=0
		Else
			o_Ad_Rs("AdPicHeight")=str_AdPicHeight
		End If
		If str_AdPicWidth="" Or IsNull(str_AdPicWidth) Then
			o_Ad_Rs("AdPicWidth")=0
		Else
			o_Ad_Rs("AdPicWidth")=str_AdPicWidth
		End If
		'if left(Cstr(str_AdLinkUrl),len("Http://"))=0 then str_AdLinkUrl = "Http://"&str_AdLinkUrl
		o_Ad_Rs("AdLinkAdress")=str_AdLinkUrl
		o_Ad_Rs("AdCaptionTxt")=str_AdCaptionTxt
		o_Ad_Rs("AdLoopFactor")=CintStr(str_LoopFactor)
		o_Ad_Rs("AdEndDate")=str_LoopEndDate
		o_Ad_Rs("AdMaxClickNum")=CintStr(str_AdClickNum)
		o_Ad_Rs("AdMaxShowNum")=CintStr(str_AdShowNum)
		o_Ad_Rs("AdRemarks")=str_AdRemarks
		o_Ad_Rs("AdAddDate")=Now()
		o_Ad_Rs("AdClickNum")=0
		o_Ad_Rs("AdClassID")=str_AdClassID
		If strLock=0 Then
			o_Ad_Rs("AdLock")=0
		Else
			o_Ad_Rs("AdLock")=1
		End If
		o_Ad_Rs("AdTxtColNum")=int_AdTxtColNum
		o_Ad_Rs.Update
		o_Ad_Rs.Close
		Set o_Ad_Rs=Nothing
		
		Set o_Ad_Rs=Conn.execute("Select Top 1 AdID,AdLoopAdID From FS_AD_Info Order By AdID Desc")
			lng_TempAdID=o_Ad_Rs("AdID")
			lng_TempLoopAdID=o_Ad_Rs("AdLoopAdID")
		Set o_Ad_Rs=Nothing
		
		For int_Txt_i=1 to Request.Form("AdTxtContent").Count
			Conn.execute("insert into FS_AD_TxtInfo(AdID,AdTxtContent,Css,LinkUrl) values("&CintStr(lng_TempAdID)&",'"&NoSqlHack(Request.Form("AdTxtContent")(int_Txt_i))&"','"&NoSqlHack(Request.Form("AdTxtCss")(int_Txt_i))&"','"&NoSqlHack(Request.Form("AdTxtLink")(int_Txt_i))&"')")			
		Next
		
		Select Case Clng(str_AdType)
			Case 0 call ShowAds(lng_TempAdID)
			Case 1 call NewWindow(lng_TempAdID)
			Case 2 call OpenWindow(lng_TempAdID)
			Case 3 call FilterAway(lng_TempAdID)
			Case 4 call DialogBox(lng_TempAdID)
			Case 5 call ClarityBox(lng_TempAdID)
			Case 6 call DriftBox(lng_TempAdID)
			Case 7 call LeftBottom(lng_TempAdID)
			Case 8 call RightBottom(lng_TempAdID)
			Case 9 call Couplet(lng_TempAdID)
			Case 10 call Cycle(lng_TempAdID,lng_TempLoopAdID)
			Case 11 call AdTxt(lng_TempAdID)
		End Select
		If Clng(str_IsLoopvalue) = 1 then
			Call Cycle(AdID,TempLocation)
		End if
		strShowErr = "<li>添加成功!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Manage.asp")
		Response.end
	End If
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>广告管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<script src="../../FS_Inc/PublicJS.js" language="JavaScript"></script>
<body>
<form action="" name="AddAds" method="post">
<%
	Dim str_AdClass_Sql,o_AdClass_Rs,str_AdClass_str,str_Selected,str_ClassType
	str_ClassType=CintStr(Request.QueryString("AdClassID"))
	str_AdClass_Sql="Select AdClassID,AdClassName,Lock from FS_AD_Class"
	Set o_AdClass_Rs=Conn.execute(str_AdClass_Sql)
	If Not o_AdClass_Rs.Eof Then
		While Not o_AdClass_Rs.Eof
			If CintStr(str_ClassType)=CintStr(o_AdClass_Rs("AdClassID")) Then
				str_Selected=" selected"
			Else
			    str_Selected = ""
			End If
			str_AdClass_str=str_AdClass_str&"<option value="&o_AdClass_Rs("AdClassID")&str_Selected&">"&o_AdClass_Rs("AdClassName")&"</option>"
		o_AdClass_Rs.MoveNext
		Wend
	End If
	Set o_AdClass_Rs=Nothing
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td colspan="4" class="xingmu">添加广告</td>
  </tr>
  <tr> 
      <td height="8" colspan="4" class="hback"><%=str_ShowType%> | <a href="javascript:history.go(-1);">返回上级</a></td>
  </tr>
  <tr>
    <td height="4" colspan="4" class="hback"><input type="hidden" name="Type" value="<%=str_AdOpType%>"></td>
  </tr>
  <tr>
    <td height="4" align="right" class="hback">广告分类</td>
    <td height="4" colspan="3" class="hback"><select name="AdClassID" size="1" style="width:600px" o>
      <option value="-1">请选择广告栏目,如果不选择(默认不把此广告加入栏目)</option>
      <%=str_AdClass_str%>
    </select></td>
    </tr>
  <tr>
    <td width="12%" height="20" align="right" class="hback">广告名称</td>
    <td width="38%" height="20" align="left" class="hback"><input name="AdName" type="text" id="AdName" size="32" maxlength="20" title="广告名称，必填" value="<%=temp_Adname%>">
      <font color="#FF0000">*必须填写项目</font></td>
    <td width="12%" height="20" align="right" class="hback">广告类型</td>
    <td width="38%" align="left" class="hback"><select name="AdType" id="AdType" style="width:200" onChange="javascript:ChooseType(this.value);">
      <option value="0" <%if Clng(temp_adType)=0 Then Response.write "selected"%>>普通显示广告</option>
      <option value="1" <%if Clng(temp_adType)=1 Then Response.write "selected"%>>弹出新窗口</option>
      <option value="2" <%if Clng(temp_adType)=2 Then Response.write "selected"%>>打开新窗口</option>
      <option value="3" <%if Clng(temp_adType)=3 Then Response.write "selected"%>>渐隐消失</option>
      <option value="4" <%if Clng(temp_adType)=4 Then Response.write "selected"%>>网页对话框</option>
      <option value="5" <%if Clng(temp_adType)=5 Then Response.write "selected"%>>透明对话框</option>
      <option value="6" <%if Clng(temp_adType)=6 Then Response.write "selected"%>>满屏浮动</option>
      <option value="7" <%if Clng(temp_adType)=7 Then Response.write "selected"%>>左下底端</option>
      <option value="8" <%if Clng(temp_adType)=8 Then Response.write "selected"%>>右下底端</option>
      <option value="9" <%if Clng(temp_adType)=9 Then Response.write "selected"%>>对联广告</option>
      <option value="10" <%if Clng(temp_adType)=10 Then Response.write "selected"%>>循环广告</option>      
	  <option value="11" <%if Clng(temp_adType)=11 Then Response.write "selected"%>>文字广告</option>
    </select></td>
  </tr>
  <tr id="tr1">
    <td height="32" align="right" class="hback">循环广告</td>
    <td height="32" align="left" class="hback">
      <input name="IsLoop" type="checkbox" id="IsLoop" value="0" title="将非循环类广告添加到循环广告中循环显示" onClick="javascript:ChooseCycleDis();" <%If temp_adloop=1 Then Response.write "checkend"%>><input type="hidden" name="IsLoopvalue" value="0">
      循环广告位  
    <select name="LoopAdName" title="将非循环广告设置为循环广告后必选" disabled style="width:110">
		<%
			Dim o_Temp_Loopname
			Set o_Temp_Loopname=Conn.execute("select AdID,AdName From FS_AD_Info order by AdID")
			If Not o_Temp_Loopname.Eof Then
				Response.Write("<option value=""-1""></option>")
				While Not o_Temp_Loopname.Eof
					Response.Write("<option value="&o_Temp_Loopname("AdID")&">"&o_Temp_Loopname("AdName")&"</option>")
					o_Temp_Loopname.MoveNext
				Wend
			Else
				Response.Write("<option value=""-1""></option>")
				Response.Write("<option value=""-1"">当前无广告,请添加</option>")
			End If
			Set o_Temp_Loopname=Nothing
		%>
    </select>    </td>
    <td height="32" align="right" class="hback">循环方向</td>
    <td height="32" align="left" class="hback"><select name="LoopFollow" id="select" style="width:70" title="将广告设置为循环广告后必选" disabled>
      <option value="0" <%If temp_adloopfollow=0 Then Response.write "selected"%>>向   上</option>
      <option value="1" <%If temp_adloopfollow=1 Then Response.write "selected"%>>向   下</option>
      <option value="2" <%If temp_adloopfollow=2 Then Response.write "selected"%>>向   左</option>
      <option value="3" <%If temp_adloopfollow=3 Then Response.write "selected"%>>向   右</option>
                        </select>
    循环速度    
    <input name="LoopSpeed" type="text" id="LoopSpeed" size="17" maxlength="20" style="width:70"title="将广告设置为循环广告后必填" disabled value="<%=temp_adloopspeed%>" onKeyUp="if(isNaN(value))execCommand('undo')"  onafterpaste="if(isNaN(value))execCommand('undo')"></td>
  </tr>
  <tr id="tr2">
    <td height="18" align="right" class="hback">图片/动画地址</td>
    <td height="18" align="left" class="hback"><input name="LoopPicAdress" type="text" id="LoopPicAdress" size="19" title="广告图片地址：必选项" value="<%=temp_adloopic%>">
    <input name="SelectPic" type="button" id="SelectPic" value="选择图片" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',400,320,window,document.AddAds.LoopPicAdress);">
    <font color="#FF0000">*必须填写项目</font></td>
    <td height="18" align="right" class="hback">图片/动画地址</td>
    <td height="18" align="left" class="hback"><input name="LoopRPicAdress" type="text" id="LoopRPicAdress" size="19" title="如果广告类型为对联广告，请选择此项，其它类型不用选择" disabled value="<%=temp_adloopRpic%>">
      <input name="SelectRPic" type="button" id="SelectPic2" value="选择图片" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',400,320,window,document.AddAds.LoopRPicAdress);" disabled></td>
  </tr>
  <tr id="tr3">
    <td height="8" align="right" class="hback">图片/动画高度</td>
    <td height="8" align="left" class="hback"><input name="AdPicHeight" type="text" id="AdPicHeight" size="32" maxlength="20" title="图片高度：必选项" value="<%=temp_picH%>" onKeyUp="if(isNaN(value))execCommand('undo')"  onafterpaste="if(isNaN(value))execCommand('undo')">
      <font color="#FF0000">*必须填写项目</font></td>
    <td height="8" align="right" class="hback">图片/动画宽度</td>
    <td height="8" align="left" class="hback"><input name="AdPicWidth" type="text" id="AdPicWidth" size="32" maxlength="20" title="图片宽度：必选项" value="<%=temp_picW%>" onKeyUp="if(isNaN(value))execCommand('undo')"  onafterpaste="if(isNaN(value))execCommand('undo')">
      <font color="#FF0000">*必须填写项目</font></td>
  </tr>
  <tr id="tr4" style="display:none">
	<%
	If str_AdOpType<>"Add" Then
		Dim temp_Ad_TxtID,temp_Content_txt,temp_Css_txt,temp_link_txt
		Dim temp_Ad_TxtID1,temp_Content_txt1,temp_Css_txt1,temp_link_txt1
		Set temp_Txt_Rs=Conn.execute("select Ad_TxtID,AdID,AdTxtContent,Css,LinkUrl from FS_AD_TxtInfo where AdID="&CintStr(Temp_AdID)&" order by Ad_TxtID")
		If Not temp_Txt_Rs.Eof Then	
			temp_Ad_TxtID1=temp_Txt_Rs("Ad_TxtID")
			temp_Content_txt1=temp_Txt_Rs("AdTxtContent")
			temp_Css_txt1=temp_Txt_Rs("Css")
			temp_link_txt1=temp_Txt_Rs("LinkUrl")
		temp_Txt_i=1
		temp_Txt_Rs.Movenext
		Do While Not temp_Txt_Rs.Eof
			temp_Ad_TxtID=temp_Txt_Rs("Ad_TxtID")
			temp_Content_txt=temp_Txt_Rs("AdTxtContent")
			temp_Css_txt=temp_Txt_Rs("Css")
			temp_link_txt=temp_Txt_Rs("LinkUrl")

			temp_Txtcontentstr=temp_Txtcontentstr&"<tr><td> 显示文本 <input name=""AdTxtContent"" type=""text"" size=""30"" maxlength=""200"" value="""&temp_Content_txt&"""> 样式 <input name=""AdTxtCss"" type=""text"" size=""7"" maxlength=""40"" value="""&temp_Css_txt&"""> 链接地址 <input name=""AdTxtLink"" type=""text"" value="""&temp_Txt_Rs("LinkUrl")&""" size=""16"" maxlength=""100""><input name=""TxtID"" type=""hidden"" value="""&temp_Ad_TxtID&"""> <a href=""#"" onclick='f_delete(this.parentElement.parentElement)'>删除</a> </td></tr>"
			temp_Txt_Rs.Movenext
			temp_Txt_i=temp_Txt_i+1
		Loop
		End If
	Set temp_Txt_Rs=Nothing
	End If
%>
    <td height="4" align="right" valign="top" class="hback">广告内容<br>
      (<a href=javascript:f_add()><img src="/admin/images/add.gif" border="0" alt="点击添加更多文字广告"></a>)
      </td>
    <td height="4" colspan="3" align="left" class="hback">
	<table width="641" id="tb1">
	  <tr><td>显示文本
 <input name="AdTxtContent" type="text" size="30" maxlength="200" value="<%=temp_Content_txt1%>">
 样式 
      <input name="AdTxtCss" type="text" size="7" maxlength="40" value="<%=temp_Css_txt1%>">
      链接地址
      <input name="AdTxtLink" type="text" id="AdTxtLink" value="<%=temp_link_txt1%>" size="16" maxlength="100">
      列数
      <input name="AdTxtColNum" type="text" id="AdTxtColNum" value="<%=temp_AdTxtColNum%>" size="4" maxlength="2" onKeyUp="if(isNaN(value))execCommand('undo')"  onafterpaste="if(isNaN(value))execCommand('undo')"></td>
	  </tr><%=temp_Txtcontentstr%></table>
      <input name="TxtID" type="hidden" value="<%=temp_Ad_TxtID1%>"></td></tr>
  <tr id="tr6">
    <td height="18" align="right" class="hback">链接地址</td>
    <td height="18" align="left" class="hback"><input name="AdLinkUrl" type="text" id="AdLinkUrl" size="32" maxlength="200" title="广告链接地址，必填项" value="<%=temp_adlink%>">
      <font color="#FF0000">*必须填写项目</font></td>
    <td height="18" align="right" class="hback">说明文字</td>
    <td height="18" align="left" class="hback"><input name="AdCaptionTxt" type="text" size="32" maxlength="100" title="广告说明文字，可选项" value="<%=temp_adcaptiontxt%>"></td>
  </tr>
  <tr>
    <td height="18" align="right" class="hback">显示条件</td>
    <td height="18" align="left" class="hback"><input name="LoopFactor" type="radio" title="将此广告设置为不受任何条件限制而永不过期。选择此项后，最大点击次数、最大显示次数和截止日期不用填写" value="0" onClick="javascript:LoopFactorClick('1');" <%If temp_loopfactor=0 Then Response.write "checked"%>>    
    无条件显示
      <input type="radio" name="LoopFactor" value="1" title="将广告设置为有条件显示后,此广告将在满足最大点击次数、最大显示次数和截止日期中的任何一项后失效" onClick="javascript:LoopFactorClick('2');" <%If temp_loopfactor=1 Then Response.write "checked"%>>
有条件显示</td>
    <td height="18" align="right" class="hback">截止日期</td>
    <td height="18" align="left" class="hback"><input name="LoopEndDate" type="text" id="LoopEndDate" size="19" disabled value="<%=temp_loopenddate%>" readonly>
      <input name="SelectDate" type="button" id="SelectDate" value="选择时间" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.AddAds.LoopEndDate);" disabled></td>
  </tr>
  <tr>
    <td height="8" align="right" class="hback">点击次数</td>
    <td height="8" align="left" class="hback"><input name="AdMaxClickNum" type="text" id="AdClickNum" size="32" maxlength="30" title="设置广告的最大点击数量,广告将在点击次数达到此数量后失效。如果不设置此项，请置空" disabled value="<%=temp_maxclicknum%>" onKeyUp="if(isNaN(value))execCommand('undo')"  onafterpaste="if(isNaN(value))execCommand('undo')"></td>
    <td height="8" align="right" class="hback">显示次数</td>
    <td height="8" align="left" class="hback"><input name="AdMaxShowNum" type="text" id="AdShowNum" size="32" maxlength="30" title="设置广告的最大显示数量,广告将在显示次数达到此数量后失效。如果不设置此项，请置空" disabled value="<%=temp_maxshownum%>" onKeyUp="if(isNaN(value))execCommand('undo')"  onafterpaste="if(isNaN(value))execCommand('undo')"></td>
  </tr>
  <tr>
    <td height="9" align="right" valign="top" class="hback">广告备注</td>
    <td height="9" colspan="3" align="left" class="hback"><textarea name="AdRemarks" cols="98" rows="5" title="广告备注,仅供后台查阅,不做前台调用"><%=temp_adremarks%></textarea></td>
  </tr>
</table>
</form>
</body>
</html>
<script language="javascript">
function Ad_Flag()
{
	if (document.AddAds.AdName.value=="")
	{
			alert("请输入广告名称");
			document.AddAds.AdName.focus();
			return false;
	}
	if (parseInt(document.AddAds.AdType.value)==11)
	{
		var _arr=document.all.AdTxtContent
	   if (typeof(_arr.length)=="undefined")
	   {
				if (_arr.value=="")
				{
					alert("请输入文字显示内容");
					_arr.focus();
					return false;
				}
	   }
	   else
	   {
			for (var j=0;j<_arr.length;j++)
			{	
				if (_arr[j].value=="")
				{
					alert("请输入文字显示内容");
					_arr[j].focus();
					return false;
				}
			}
		}
	}
	else
	{
		if (document.AddAds.LoopPicAdress.value=="")
		{
			alert("请输入图片地址");
			document.AddAds.LoopPicAdress.focus();
			return false;
		}
		if (parseInt(document.AddAds.AdType.value)==9)
		{
			if (document.AddAds.LoopRPicAdress.value=="")
			{
				alert("请输入图片地址");
				document.AddAds.LoopRPicAdress.focus();
				return false;
			}
		}
		if (parseInt(document.AddAds.IsLoopvalue.value)==1)
		{
			if (document.AddAds.LoopSpeed.value=="")
			{
				alert("请输入循环速度");
				document.AddAds.LoopSpeed.focus();
				return false;
			}
			else
			{
				if (isNaN(document.AddAds.LoopSpeed.value))
				{
					alert("循环速度必须为整数");
					document.AddAds.LoopSpeed.value="";
					document.AddAds.LoopSpeed.focus();
					return false;
				}
				else
				{
					if (parseInt(document.AddAds.LoopSpeed.value)<0)
					{
						alert("循环速度必须为整数");
						document.AddAds.LoopSpeed.focus();
						document.AddAds.LoopSpeed.value="";
						return false;
					}
				}
			}
		}
		if (document.AddAds.AdPicHeight.value=="")
		{
			alert("请输入图片高度");
			document.AddAds.AdPicHeight.focus();
			return false;
		}
		else
		{
			if (isNaN(document.AddAds.AdPicHeight.value)==true)
			{
				alert("图片高度必须为整数");
				document.AddAds.AdPicHeight.value="";
				document.AddAds.AdPicHeight.focus();
				return false;
			}
			else
			{
				if (parseInt(document.AddAds.AdPicHeight.value)<0)
				{
					alert("图片高度必须为整数");
					document.AddAds.AdPicHeight.value="";
					document.AddAds.AdPicHeight.focus();
					return false;
				}
			}
		}
		if (document.AddAds.AdPicWidth.value=="")
		{
			alert("请输入图片宽度");
			document.AddAds.AdPicWidth.focus();
			return false;
		}
		else
		{
			if (isNaN(document.AddAds.AdPicWidth.value)==true)
			{
				alert("图片宽度必须为整数");
				document.AddAds.AdPicWidth.value="";
				document.AddAds.AdPicWidth.focus();
				return false;
			}
			else
			{
				if (parseInt(document.AddAds.AdPicWidth.value)<0)
				{
					alert("图片宽度必须为整数");
					document.AddAds.AdPicWidth.value="";
					document.AddAds.AdPicWidth.focus();
					return false;
				}
			}
		}
		if (document.AddAds.AdLinkUrl.value=="")
		{
			alert("请输入链接地址");
			document.AddAds.AdLinkUrl.focus();
			return false;
		}
	}
	if (document.AddAds.AdClickNum.value!="")
	{
		if (isNaN(document.AddAds.AdClickNum.value)==true)
		{
			alert("点击次数必须为整数");
			document.AddAds.AdClickNum.value="";
			document.AddAds.AdClickNum.focus();
			return false;
		}
		else
		{
			if (parseInt(document.AddAds.AdClickNum.value)<0)
			{
				alert("点击次数必须为整数");
				document.AddAds.AdClickNum.value="";
				document.AddAds.AdClickNum.focus();
				return false;
			}
		}
	}
	else
	{
		document.AddAds.AdClickNum.value = 0;
	}
	if (document.AddAds.AdShowNum.value!="")
	{
		if (isNaN(document.AddAds.AdShowNum.value)==true)
		{
			alert("显示次数必须为整数");
			document.AddAds.AdShowNum.value="";
			document.AddAds.AdShowNum.focus();
			return false;
		}
		else
		{
			if (parseInt(document.AddAds.AdShowNum.value)<0)
			{
				alert("显示次数必须为整数");
				document.AddAds.AdShowNum.value="";
				document.AddAds.AdShowNum.focus();
				return false;
			}
		}
	}
	return true;
}
function LoopFactorClick(type)
{
	if(type==1)
	{
		document.AddAds.LoopEndDate.disabled=true;
		document.AddAds.SelectDate.disabled=true;
		document.AddAds.AdClickNum.disabled=true;
		document.AddAds.AdShowNum.disabled=true;
		document.AddAds.LoopFactor.value=1;
	}
	else
	{
		document.AddAds.LoopEndDate.disabled=false;
		document.AddAds.SelectDate.disabled=false;
		document.AddAds.AdClickNum.disabled=false;
		document.AddAds.LoopFactor.value=0;
		document.AddAds.AdShowNum.disabled=false;
	}
}
function ChooseType(type)
{
	switch (parseInt(type))
		{
		case 9:
			document.AddAds.LoopRPicAdress.disabled=false;
			document.AddAds.SelectRPic.disabled=false;
			show();
			break;
		case 0:
			document.AddAds.LoopRPicAdress.disabled=true;
			document.AddAds.SelectRPic.disabled=true;
			document.AddAds.IsLoop.disabled=true;
			document.AddAds.LoopAdName.disabled=true;
			document.AddAds.LoopFollow.disabled=false;
			document.AddAds.LoopSpeed.disabled=false;
			show();
			break;
		case 11:
			document.all.tr1.style.display="none";
			document.all.tr2.style.display="none";
			document.all.tr3.style.display="none";
			document.all.tr4.style.display="";
			document.all.tr6.style.display="none";			
			break;
		default:
			document.AddAds.IsLoop.disabled=false;
			document.AddAds.LoopAdName.disabled=true;
			document.AddAds.LoopFollow.disabled=true;
			document.AddAds.LoopSpeed.disabled=true;
			document.AddAds.LoopRPicAdress.disabled=true;
			document.AddAds.SelectRPic.disabled=true;
			document.AddAds.IsLoop.checked=false;
			show();
			break;
		}
}
function ChooseCycleDis()
{
	if (document.AddAds.IsLoop.checked==true)
	{
		document.AddAds.LoopFollow.disabled=false;
		document.AddAds.LoopSpeed.disabled=false;
		document.AddAds.LoopAdName.disabled=false;
		document.AddAds.IsLoopvalue.value=1;
	}
	else
	{
		document.AddAds.LoopFollow.disabled=true;
		document.AddAds.LoopSpeed.disabled=true;
		document.AddAds.LoopAdName.disabled=true;
		document.AddAds.IsLoopvalue.value=0;
	}
}
function Ad_Save()
{
	document.AddAds.submit();
}
function Ad_Update()
{
	document.AddAds.action="?Submit=SubUp&ID=<%=Temp_AdID%>&OpPage=<%=Replace(Request.QueryString("OpPage"),"'","''")%>";
	document.AddAds.submit();
}
function show()
{
	document.all.tr1.style.display="";
	document.all.tr2.style.display="";
	document.all.tr3.style.display="";
	document.all.tr4.style.display="none";
	document.all.tr6.style.display="";			
}
function f_add()
{
    var _tbobj=document.all.tb1
	var _trobj=_tbobj.rows
	if (_trobj.length>9)
	{
		return;
	} 
	var _newRow=_tbobj.insertRow(_tbobj.rows.length)
	var _newCell=_newRow.insertCell(0)
	_newCell.innerHTML="<tr><td> 显示文本 <input name=AdTxtContent type=text id=AdTxtContent size=30 maxlength=200> 样式 <input name=AdTxtCss type=text size=7 maxlength=40> 链接地址 <input name=AdTxtLink type=text id=AdTxtLink size=16 maxlength=100><a href='#'onclick='f_delete(this.parentElement.parentElement)'> 删除 </a></td></tr>"; 
	_newCell=_newRow.insertCell(1) 
}
function f_delete(_aobj)
{
	var _tbobj=document.all.tb1
	var _trobj=_tbobj.rows
	var _deltr=_aobj
   for(var i=0;i<_trobj.length;i++)
   {
   		if (_deltr==_trobj[i])
		{
			_tbobj.deleteRow(i)
			break;
		}
   }
}
</script>
<%
If temp_loopfactor=1 Then Response.write "<script language=""javascript"">LoopFactorClick('2');</script>"
If Clng(temp_adType)=11 Then Response.write "<script language=""javascript"">ChooseType(11);</script>"
If Clng(temp_adType)=9 Then Response.Write("<script language=""javascript"">document.AddAds.LoopRPicAdress.disabled=false;document.AddAds.SelectRPic.disabled=false;</script>")
Sub Alert(Msg)
	Response.write "<script language=""javascript"">alert('"&Msg&"');history.go(-1);</script>"
End Sub
%>
<%
Set Conn=nothing
%><!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






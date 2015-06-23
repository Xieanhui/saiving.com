<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_InterFace/NS_Public.asp" -->
<!--#include file="../../FS_InterFace/MS_Public.asp" -->
<!--#include file="../../FS_InterFace/DS_Public.asp" -->
<!--#include file="../../FS_InterFace/ME_Public.asp" -->
<!--#include file="../../FS_InterFace/MF_Public.asp" -->
<!--#include file="../../FS_InterFace/SD_Public.asp" -->
<!--#include file="../../FS_InterFace/HS_Public.asp" -->
<!--#include file="../../FS_InterFace/AP_Public.asp" -->
<!--#include file="../../FS_InterFace/Other_Public.asp" -->
<!--#include file="../../FS_InterFace/Refresh_Function.asp" -->
<!--#include file="../../FS_Inc/WaterPrint_Function.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.CacheControl = "no-cache"
	Dim Conn,User_Conn,obj_Save_Rs,strShowErr,str_checkTF
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF 
	set Fs_news = new Cls_News
	Fs_News.GetSysParam()
	If Not Fs_news.IsSelfRefer Then response.write "非法提交数据":Response.end
	Dim Fs_news,str_News_Action,str_NewsType,str_isdraft,str_ClassID,str_SpecialEName,str_SpecialID_EName,str_NewsTitle,str_TitleColor,str_titleBorder,str_TitleItalic,str_isShowReview,str_PopID
	Dim str_URLAddress,str_CurtTitle,str_KeyWords,str_KeywordSaveTF,str_Templet,str_NewsSmallPicFile,str_NewsPicFile,str_PicborderCss,str_Author,str_AuthorSaveTF,str_Source,str_SourceSaveTF
	Dim str_NewsNaviContent,str_Content,str_PointNumber,str_Money,str_GroupID,str_FileName,str_FileExtName,str_addtime,str_Hits,str_NewsID,str_TodayNewsPicTF,str_isDraftTF,str_ReturnUrl,obj_isdraft_rs
	Dim str_NewsProperty_Rec,str_NewsProperty_mar,str_NewsProperty_rev,str_NewsProperty_constr,str_NewsProperty_tt,str_NewsProperty_hots,str_NewsProperty_jc,str_NewsProperty_unr,str_NewsProperty_ann,str_NewsProperty_filt,str_NewsProperty_Remote
	Dim str_NewsProperty_Rec_1,str_NewsProperty_mar_1,str_NewsProperty_rev_1,str_NewsProperty_constr_1,str_NewsProperty_tt_1 ,str_NewsProperty_hots_1,str_NewsProperty_jc_1,str_NewsProperty_unr_1,str_NewsProperty_ann_1,str_NewsProperty_filt_1,str_NewsProperty_Remote_1,IsAdPic,AdPicWH,AdPicLink,AdPicAdress,DefaultFileExtName
    Dim AdPicWHw,IsApicArea
	If not trim(Request.Form("d_Id"))="" Then
		If trim(Request.Form("d_Id"))>0 Then
			Dim CustColumnRs,CustSql,CustColumnArr
			CustSql="select DefineID,ClassID,D_Name,D_Coul,D_Type,D_isNull,D_Value,D_Content,D_SubType from [FS_MF_DefineTable] Where D_SubType='NS' and  Classid="& CintStr(trim(Request.Form("d_Id"))) &""
			Set CustColumnRs=CreateObject(G_FS_RS)
			CustColumnRs.Open CustSql,Conn,1,3
			If Not CustColumnRs.Eof Then
				CustColumnArr=CustColumnRs.GetRows()
			End If
			CustColumnRs.close:Set CustColumnRs = Nothing
		End If
	end if
	'=====================================
	str_News_Action = Request.Form("News_Action")
	if str_News_Action = "add_Save" then
		Dim HaveNewsIDTF,ChedkNewsIDObj,Temp_NewsID_Str
		HaveNewsIDTF = False
		Do While Not HaveNewsIDTF
			Temp_NewsID_Str = Fs_News.GetRamCode(15)
			Set ChedkNewsIDObj = Conn.ExeCute("Select NewsID From FS_NS_News Where NewsID = '" & NoSqlHack(Temp_NewsID_Str) & "'")
			If ChedkNewsIDObj.Eof Then
				str_NewsID = Temp_NewsID_Str
				HaveNewsIDTF = True
				Exit Do
			End IF
			ChedkNewsIDObj.Close : Set ChedkNewsIDObj = NOthing	
		Loop
	Else
		str_NewsID = NoSqlHack(Trim(Request.Form("NewsID")))
	End if
	str_NewsType = NoSqlHack(Request.Form("NewsType"))
	str_isdraft = NoSqlHack(Request.Form("isdraft"))
	if str_isdraft<>"" then
		Set obj_isdraft_rs = server.CreateObject(G_FS_RS)
		obj_isdraft_rs.Open "Select ID from FS_NS_News where isdraft=1  Order by ID desc",Conn,1,3
		if obj_isdraft_rs.recordcount>20 then
			strShowErr = "<li>您的草稿箱中的"& Fs_news.allInfotitle&"已经超过20条信息，操作失败</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
	End if
	str_ClassID = NoSqlHack(Request.Form("ClassID"))
	str_SpecialID_EName =  NoSqlHack(Request.Form("SpecialID_EName"))
	str_SpecialEName = NoSqlHack(Request.Form("SpecialID"))
	str_NewsTitle= NoSqlHack(Request.Form("NewsTitle"))
	str_TitleColor = NoSqlHack(Request.Form("TitleColor"))
	str_titleBorder = NoSqlHack(Request.Form("titleBorder"))
	str_TitleItalic =  NoSqlHack(Request.Form("TitleItalic"))
	str_isShowReview = NoSqlHack(Request.Form("isShowReview"))
	str_PopID = NoSqlHack(Request.Form("PopID"))
	str_URLAddress = NoSqlHack(Request.Form("URLAddress"))
	str_CurtTitle = NoSqlHack(Request.Form("CurtTitle"))
	str_KeyWords =  NoSqlHack(Request.Form("KeywordText"))
	str_KeywordSaveTF =  NoSqlHack(Trim(Request.Form("KeywordSaveTF")))
	str_Templet =  NoSqlHack(Request.Form("Templet"))
	str_NewsSmallPicFile =  NoSqlHack(Request.Form("NewsSmallPicFile"))
	str_NewsPicFile =  NoSqlHack(Request.Form("NewsPicFile"))
	str_PicborderCss = Request.Form("PicborderCss")
	str_Author = NoSqlHack(Trim(Request.Form("Author")))
	str_AuthorSaveTF= NoSqlHack(Request.Form("AuthorSaveTF"))
	str_Source = NoSqlHack(Request.Form("Source"))
	str_SourceSaveTF = NoSqlHack(Request.Form("SourceSaveTF"))
	str_NewsNaviContent = NoSqlHack(Request.Form("NewsNaviContent"))

	str_Content = NoSqlHack(Request.Form("Content"))

	str_PointNumber = NoSqlHack(Request.Form("PointNumber"))
	str_Money = NoSqlHack(Request.Form("Money"))
	str_GroupID = NoSqlHack(Request.Form("BrowPop"))
	str_FileName = NoSqlHack(Request.Form("FileName"))
	str_FileExtName = NoSqlHack(Request.Form("FileExtName"))
	str_addtime = Request.Form("addtime")
	str_Hits = NoSqlHack(Request.Form("Hits") )
	str_NewsProperty_Rec = NoSqlHack(Trim(Request.Form("NewsProperty_Rec"))) 
	str_NewsProperty_mar = NoSqlHack(Trim(Request.Form("NewsProperty_mar"))) 
	str_NewsProperty_rev = NoSqlHack(Trim(Request.Form("NewsProperty_rev"))) 
	str_NewsProperty_constr =  NoSqlHack(Trim(Request.Form("NewsProperty_constr"))) 
	str_NewsProperty_tt =   NoSqlHack(Trim(Request.Form("NewsProperty_tt"))) 
	str_NewsProperty_hots=   NoSqlHack(Trim(Request.Form("NewsProperty_hots"))) 
	str_NewsProperty_jc=   NoSqlHack(Trim(Request.Form("NewsProperty_jc")) )
	str_NewsProperty_unr = NoSqlHack(Trim(Request.Form("NewsProperty_unr")) )
	str_NewsProperty_ann = NoSqlHack(Trim(Request.Form("NewsProperty_ann")) )
	str_NewsProperty_filt = NoSqlHack(Trim(Request.Form("NewsProperty_filt"))) 
	str_NewsProperty_Remote = NoSqlHack(Trim(Request.Form("NewsProperty_Remote")) )
	str_TodayNewsPicTF = NoSqlHack(Trim(Request.Form("TodayNewsPicTF"))) 
	IsAdPic = CintStr(Request.Form("IsAdPic"))
	AdPicWH = NoSqlHack(Request.Form("AdPicWH"))
	
	'处理文字画中画。Fsj 08.12.2
	if ""<>Request.Form("AdPicWHw") then 
	    AdPicWHw = CintStr(Request.Form("AdPicWHw"))
	end if
	IsApicArea= NoSqlHack(Request.Form("IsApicArea"))
	
	AdPicLink = NoSqlHack(Request.Form("AdPicLink"))
	AdPicAdress =  NoSqlHack(Request.Form("AdPicAdress"))
	
	Set obj_Save_Rs = server.CreateObject(G_FS_RS)
	'判断合法性
	If Trim(str_NewsTitle)="" then
		strShowErr = "<li>请填写标题</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	'自定义字段检查（默认值，必填值）
	If IsArray(CustColumnArr) Then
		For i = 0 to UBound(CustColumnArr,2)
			If CustColumnArr(5,i)="0" Then
				If Request.Form("FS_NS_Define_"&CustColumnArr(3,i))="" Then
					strShowErr = "<li>["&CustColumnArr(2,i)&"-自定义字段]不可以为空，请填写或设置默认值</li>"
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
				End If
				If CustColumnArr(4,i)="6" Then
					If Not Isdate(Request.Form("FS_NS_Define_"&CustColumnArr(3,i))) Then
						strShowErr="<li>["&CustColumnArr(2,i)&"-自定义字段]必须为为日期，请重新填写！</li>"
						Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
						Response.end
					End If
				End If
			End If
			If CustColumnArr(4,i)="5" Then
				If Not IsNumeric(Request.Form("FS_NS_Define_"&CustColumnArr(3,i))) Then
					strShowErr="<li>["&CustColumnArr(2,i)&"-自定义字段]必须为数字，请重新填写！</li>"
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
				End If
			End If
		Next
	End If
	If Trim(str_ClassID)="" then
		strShowErr = "<li>请选择栏目或您填写的栏目不正确</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	if str_NewsType="TitleNews" then
		If Trim(str_URLAddress)="" then
			strShowErr = "<li>选择标题新闻,请填写新闻连接地址</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
	Else
		If isNull(Trim(str_Templet))  or len(Trim(str_Templet))<5 then
			strShowErr = "<li>请正确填写模板地址</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		If isNull(Trim(str_Content))  or len(Trim(str_Content))<3 then
			strShowErr = "<li>请填写内容</li><li>您填写的内容少于3个字符</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		If isNull(Trim(str_FileName))  then
			strShowErr = "<li>文件名不正确，请填写正确</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		If isNull(Trim(str_addtime))  then
			strShowErr = "<li>请填写日期</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		If isdate(Trim(str_addtime)) =false then
			strShowErr = "<li>请正确填写添加日期</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		If isnull(Trim(str_FileName)) then
			strShowErr = "<li>请填写文件名</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		If isnumeric(Trim(str_hits)) =false then
			strShowErr = "<li>请正确填写点击率</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		if str_NewsType="PicNews" then
			If isnull(Trim(str_NewsSmallPicFile))  or len(Trim(str_NewsSmallPicFile))<5 then
				strShowErr = "<li>请填写图片小图</li><li>请正确填写图片小图</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
		End if
		if Instr(str_FileName,"自动编号ID") =false and Instr(str_FileName,"唯一NewsID")=false then
			if Not fs_news.chkinputchar(str_FileName) then
					strShowErr = "<li>文件名格式有错误</li><li>允许字符为：<br>&nbsp;&nbsp;&nbsp;&nbsp;大小写字母及@,.0123456789|-_</li>"
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
			End if
	   End if
	   If str_Author <> "" And Not IsNull(str_Author) Then
			If InStr(str_Author,"'") > 0 Then
				strShowErr = "<li>作者请不要包含特殊符号</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End If	
	   End If
	   If str_Source <> "" And Not IsNull(str_Source) Then
			If InStr(str_Source,"'") > 0 Then
				strShowErr = "<li>来源请不要包含特殊符号</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End If	
	   End If
	   If str_KeyWords <> "" And Not IsNull(str_KeyWords) Then
			If InStr(str_KeyWords,"'") > 0 Then
				strShowErr = "<li>关键字请不要包含特殊符号</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End If	
	   End If
	   If IsAdPic=1 and  Cintstr(Request.Form("Checkbox1"))=1 Then
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
		     strShowErr = "<li>图片高度与宽度,显示布局格式,插入位置有误，格式为100,200,1,400</li>"
			 Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			 Response.end
		  End If
		  If Not IsNumeric(split(AdPicWH,",")(3)) Then
			strShowErr = "<li>插入位置格式有误，应为正整数</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		  End If
		  If split(AdPicWH,",")(3)<0 Then 
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
       	If IsAdPic=1 and  Cintstr(Request.Form("Checkbox2"))=1 Then       	 
	        If not IsNumeric(Request.Form("AdPicWHw")) Then
	   			strShowErr = "<li>插入位置格式有误，应为正整数</li>"
			    Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		    	Response.end			
			end if
			if IsApicArea="" then 
			    strShowErr = "<li>文字画中画代码不能为空！</li>"
			    Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			    Response.end
		    End If
		    
		    
		    '处理文字画中画。
		    if AdPicWHw<=0 then 
		        AdPicWHw=0
		        str_Content="<!---文字画中画star----><table width=0 border=0 align="&G_CodeContentAlign&"><tr><td>"&IsApicArea&"</td></tr></table><!---文字画中画end--->"&str_Content		        
		    end if
		     if AdPicWHw>=len(str_Content) and  AdPicWHw>0 then 
		        AdPicWHw=len(str_Content)
		        str_Content=str_Content&"<!---文字画中画star----><table width=0 border=0 align="&G_CodeContentAlign&"><tr><td>"&IsApicArea&"</td></tr></table><!---文字画中画end--->"
		     end if
		      if AdPicWHw<len(str_Content) and  AdPicWHw>0 then     		        
		         str_Content=left(str_Content,AdPicWHw)&"<!---文字画中画star----><table width=0 border=0 align="&G_CodeContentAlign&"><tr><td>"&IsApicArea&"</td></tr></table><!---文字画中画end--->"&Right(str_Content,len(str_Content)-AdPicWHw)
		     end if
	   End if
	if str_News_Action = "add_Save" then
		'判断是否有添加权限
		if not Get_SubPop_TF(str_ClassID,"NS001","NS","news") then Err_Show
		Set obj_Save_Rs=Server.CreateObject(G_FS_RS)
		obj_Save_Rs.Open "Select * from FS_NS_News where NewsID='"& str_NewsID &"'",Conn,1,3
		if Not obj_Save_Rs.eof then
				strShowErr = "<li>NewsID意外重复，请重新添加</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
		End if
		'判断文件名是否重复
		Dim tmp_filename_rs
		set tmp_filename_rs = Conn.execute("select ID From FS_NS_News where ClassID ='"& str_ClassID &"' and FileName='"& str_FileName &"' and FileExtName='"& str_FileExtName &"' order by id desc")
		if Not tmp_filename_rs.eof then
				strShowErr = "<li>文件名重复</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
		End if
		tmp_filename_rs.close:set tmp_filename_rs = nothing
		obj_Save_Rs.addnew
		obj_Save_Rs("NewsID") = str_NewsID
		obj_Save_Rs("SaveNewsPath") = Fs_news.SaveNewsPath(Fs_news.fileDirRule)
	    Elseif str_News_Action = "Edit_Save" then
		'判断是否有修改权限
		if not Get_SubPop_TF(str_ClassID,"NS002","NS","news") then Err_Show
		DefaultFileExtName = Trim(Request.Form("DefaultFileExtName"))
		IF DefaultFileExtName <> "" And DefaultFileExtName <> str_FileExtName Then
			Dim DelDefaultObj,DefaultSavePath
			Set DelDefaultObj =  Conn.execute("select FS_NS_News.SaveNewsPath,FS_NS_News.FileName,FS_NS_News.FileExtName,FS_NS_NewsClass.SavePath,FS_NS_NewsClass.ClassEName from FS_NS_News,FS_NS_NewsClass where NewsID='"& NoSqlHack(str_NewsID) &"' and FS_NS_News.IsURL=0 and FS_NS_NewsClass.ClassID=FS_NS_News.ClassID")
			If G_VIRTUAL_ROOT_DIR = "" Then
				DefaultSavePath = ""
			Else
				DefaultSavePath = "/" & G_VIRTUAL_ROOT_DIR
			End If
			DefaultSavePath = DefaultSavePath & DelDefaultObj("SavePath")&"/"&DelDefaultObj("ClassEName")&DelDefaultObj("SaveNewsPath")&"/"&DelDefaultObj("FileName")&"."&DelDefaultObj("FileExtName")
			fso_DeleteFile(DefaultSavePath)
			DelDefaultObj.CLose : Set DelDefaultObj = NOthing
		End If
		obj_Save_Rs.Open "Select * from FS_NS_News where NewsID='"& str_NewsID &"'",Conn,1,3
	End if
	
	'生成图片头条,有待修正
	Dim str_FontSize,str_FontSpace,str_FontColor,str_FontBgColor,FontFace,PicTitle,str_PicTitle,str_Picwidth
	str_FontSize = NoSqlHack(Request.Form("FontSize"))
	str_FontSpace = NoSqlHack(Request.Form("FontSpace"))
	str_FontColor = NoSqlHack(Request.Form("FontColor"))
	str_FontBgColor = NoSqlHack(Request.Form("FontBgColor"))
	PicTitle = NoSqlHack(Request.Form("PicTitle"))
	if PicTitle = "" then
		str_PicTitle = str_NewsTitle
	else
		str_PicTitle = PicTitle
	end if
	str_Picwidth = NoSqlHack(Request.Form("PicTitlewidth"))
	FontFace =  NoSqlHack(Request.Form("FontFace"))
	If str_TodayNewsPicTF = "FoosunCMS" Then
		'图片地址：TempSysRootDir&TmpClassInfo(7) &"/"& TmpClassInfo(2)&NewsPath
		Dim NumCanvasWidth,NumCanvasHeight,StrSavePath,TodayClassEName,TempSysRootDir
		if str_picwidth = "" then
			NumCanvasWidth	=CintStr(GetStrLengthE(str_PicTitle))*CintStr(str_FontSize)
		else
			NumCanvasWidth	=CintStr(str_Picwidth)
		end if
		NumCanvasHeight	= str_FontSize
		if trim(G_VIRTUAL_ROOT_DIR) = "" then
			TempSysRootDir = ""
		else
			TempSysRootDir = "/"&G_VIRTUAL_ROOT_DIR
		end if
		TodayClassEName=Conn.Execute("select ClassEName From FS_NS_NewsClass Where ClassID='" & NoSqlHack(str_ClassID) & "'")(0)
		StrSavePath=Server.MapPath(TempSysRootDir & "/" & G_UP_FILES_DIR & "/TodayPicFiles/"& str_NewsID & ".jpg")
	   '得到水印组件类型
		Dim tmp_returnvalue
		Select Case request.Cookies("FoosunMFCookies")("FoosunMFPicClassid")
			Case "0"
				tmp_returnvalue = AspJpegCreateTextPic(NumCanvasWidth,NumCanvasHeight,"&H"&str_FontBgColor,0,str_FontColor,FontFace,0,str_FontSize,str_PicTitle,0,StrSavePath)
			Case "1"
				tmp_returnvalue = WsImgWatermarkText(NumCanvasWidth,NumCanvasHeight,"&H"&str_FontColor&"&",FontFace,str_FontSize,0,0,0,str_NewsTitle,StrSavePath)
			Case "2"
				NumCanvasWidth	= CInt(GetStrLengthE(str_NewsTitle))*(Cint(str_FontSize))
				tmp_returnvalue = ImageGenCreateTextPic(NumCanvasWidth,NumCanvasHeight,rgb(HexToDec(left(str_FontBgColor,2)),HexToDec(mid(str_FontBgColor,3,2)),HexToDec(right(str_FontBgColor,2))),rgb(HexToDec(left(str_FontColor,2)),HexToDec(mid(str_FontColor,3,2)),HexToDec(right(str_FontColor,2))),FontFace,str_FontSize,str_NewsTitle,0,0,StrSavePath)			
			Case Else
				tmp_returnvalue = True
		End Select
		if tmp_returnvalue then
			Call saveTody()
		else
			str_TodayNewsPicTF = ""
		end if
	End If		
	
	'开始插入或者更新数据
	if trim(str_GroupID)<>"" or str_PointNumber <>"" or str_Money<>"" then 
		if trim(str_FileExtName)<>"asp" then
				strShowErr = "<li>您设置了浏览权限，扩展名必须为.asp</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
		End if
	End if
	obj_Save_Rs("PopId") = str_PopId
	obj_Save_Rs("ClassID") = str_ClassID
	obj_Save_Rs("SpecialEName") = str_SpecialID_EName
	obj_Save_Rs("NewsTitle") = str_NewsTitle
	obj_Save_Rs("NewsNaviContent") = str_NewsNaviContent
	obj_Save_Rs("CurtTitle") = str_CurtTitle
	if str_isShowReview<>"" then:obj_Save_Rs("isShowReview") = 1:Else:obj_Save_Rs("isShowReview") = 0:End if
	obj_Save_Rs("TitleColor") = str_TitleColor
	if str_titleBorder <>"" then:obj_Save_Rs("titleBorder") = 1:Else:obj_Save_Rs("titleBorder") = 0:End if	
	if str_TitleItalic <>"" then:obj_Save_Rs("TitleItalic") = 1:Else:obj_Save_Rs("TitleItalic") = 0:End if	
	if str_NewsType="TitleNews" then
		obj_Save_Rs("URLAddress") = str_URLAddress
		obj_Save_Rs("IsURL") = 1 
		obj_Save_Rs("NewsPicFile") = str_NewsPicFile
		obj_Save_Rs("NewsSmallPicFile") = str_NewsSmallPicFile
		If str_NewsSmallPicFile <> "" And (Right(LCase(str_NewsSmallPicFile),4) = ".jpg" Or Right(LCase(str_NewsSmallPicFile),4) = ".gif" Or Right(LCase(str_NewsSmallPicFile),4) = ".bmp" Or Right(LCase(str_NewsSmallPicFile),4) = ".png") Then
			obj_Save_Rs("isPicNews") = 1
		Else
			obj_Save_Rs("isPicNews") = 0
		End If	
	Else 
		obj_Save_Rs("IsURL") = 0 
		dim sys_Config,Save_Content,inside_rs ,Remote_Content,config_domain,config_insideLink
		set sys_Config = Conn.execute("select top 1 IsDomain,InsideLink From FS_NS_SysParam")
		config_domain = sys_Config("isDomain")
		config_insideLink = sys_Config("insideLink")
		sys_Config.close:set sys_Config=nothing
		if config_domain="" then
			config_domain = request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		else
			config_domain = config_domain
		end if
		Dim TempForVar
		'分批保存新闻内容入数据库
		Save_Content = ""
		'For TempForVar = 1 To NoSqlHack(Request.Form("Content").Count)
		'	Save_Content = Save_Content & NoSqlHack(Request.Form("Content")(TempForVar))
		'Next
		'处理文字画中画功能 Fsj 08.12.2
		Save_Content=str_Content
		'如果开启远程保存，此处需要引入远程保存函数
		if Request.Form("NewsProperty_Remote")<>"" then   
			if instr(Lcase(Save_Content),"<img")>0 then
				CreateDateDir(Server.MapPath(Replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_UP_FILES_DIR&"/Remoteupfile","//","/")))
				Remote_Content = ReplaceRemoteUrl(Save_Content,replace("/" & G_VIRTUAL_ROOT_DIR &"/"&G_UP_FILES_DIR & "/Remoteupfile/"&year(Now())&"-"&month(now())&"/"&day(Now()),"//","/"),"http://"&config_domain,replace("/"&G_VIRTUAL_ROOT_DIR,"//","/"))
			else
				Remote_Content =Save_Content
			end if
		 else
		 	Remote_Content =Save_Content
		 end if
		''--------------------------------------自动分页的处理 
		 '' G_FS_Page_Txtlength  [FS:PAGE]
		if request.Form("ClearAllPage")="1" then Remote_Content = replace(Remote_Content,"[FS:PAGE]","")
		Remote_Content = AutoSplitPages(Remote_Content,"[FS:PAGE]",G_FS_Page_Txtlength)
		''--------------------------------------
		'2009-4-2 姚修改 RSS内容显示错误
		Dim Replace_Special_Str,Arr_Replace_Special_Str,i_Replace_Special_Str,Temp_Arr_Replace_Special_Str
		Replace_Special_Str = "&ldquo;:“|&rdquo;:”|&lsquo;:‘"
		Arr_Replace_Special_Str = Split(Replace_Special_Str,"|")
		For i_Replace_Special_Str = LBound(Arr_Replace_Special_Str) to UBound(Arr_Replace_Special_Str)
			Temp_Arr_Replace_Special_Str = Split(Arr_Replace_Special_Str(i_Replace_Special_Str),":")
			if UBound(Temp_Arr_Replace_Special_Str) >= 1 then
				Remote_Content = Replace(Remote_Content,Temp_Arr_Replace_Special_Str(0),Temp_Arr_Replace_Special_Str(1))
			end if
		Next
		'2009-4-2 姚修改 RSS内容显示错误

		obj_Save_Rs("Content") = Remote_Content
		'远程保存结束
		if str_NewsType="PicNews" then:obj_Save_Rs("isPicNews") = 1:Else:obj_Save_Rs("isPicNews") = 0:End if
		obj_Save_Rs("NewsPicFile") = str_NewsPicFile
		obj_Save_Rs("NewsSmallPicFile") = str_NewsSmallPicFile
		obj_Save_Rs("PicborderCss") = str_PicborderCss
		obj_Save_Rs("Templet") = str_Templet
		if Trim(str_GroupID)<>"" or Trim(str_PointNumber)<>"" or Trim(str_Money)<>"" then:obj_Save_Rs("isPop") =1:else:obj_Save_Rs("isPop")=0:End if
		obj_Save_Rs("Source") = str_Source
		obj_Save_Rs("Keywords") = str_Keywords
		obj_Save_Rs("Author") = str_Author
		obj_Save_Rs("Hits") = clng(str_Hits)
		obj_Save_Rs("FileName") = str_FileName
		obj_Save_Rs("FileExtName") = str_FileExtName
	End if
	Dim Temp_Admin_Name
	Temp_Admin_Name = Session("Admin_Name")
	if str_TodayNewsPicTF <> "" then:obj_Save_Rs("TodayNewsPic") = 1:Else:obj_Save_Rs("TodayNewsPic") = 0:End if
	obj_Save_Rs("Editor") = Temp_Admin_Name
	if Fs_News.isCheck = 1 then
		obj_Save_Rs("isLock") = 1
		str_checkTF=true
	Else
		Dim obj_lockTF,tmp_definedid_class
		Set obj_lockTF=conn.execute("select  NewsCheck,DefineID From FS_NS_NewsClass Where ClassID='"& NoSqlHack(str_ClassID) &"'")
		if obj_lockTF("NewsCheck") = 1 then
			obj_Save_Rs("isLock") = 1
			str_checkTF=true
		Else
			obj_Save_Rs("isLock") = 0
			str_checkTF=false
		End if
		tmp_definedid_class = obj_lockTF(1)
		obj_lockTF.close:set obj_lockTF = nothing
	End if
	'得到新闻类型参数 
	if str_NewsProperty_Rec <>"" then:str_NewsProperty_Rec_1 = 1:else:str_NewsProperty_Rec_1 = 0:End if
	if str_NewsProperty_mar <>"" then:str_NewsProperty_mar_1 = 1:else:str_NewsProperty_mar_1 = 0:End if
	if str_NewsProperty_rev <>"" then:str_NewsProperty_rev_1 = 1:else:str_NewsProperty_rev_1 = 0:End if
	if str_NewsProperty_constr <>"" then:str_NewsProperty_constr_1 = 1:else:str_NewsProperty_constr_1 = 0:End if
	if str_NewsProperty_tt <>"" then:str_NewsProperty_tt_1 = 1:else:str_NewsProperty_tt_1 = 0:End if
	if str_NewsProperty_hots <>"" then:str_NewsProperty_hots_1 = 1:else:str_NewsProperty_hots_1 = 0:End if
	if str_NewsProperty_jc <>"" then:str_NewsProperty_jc_1 = 1:else:str_NewsProperty_jc_1 = 0:End if
	if str_NewsProperty_unr <>"" then:str_NewsProperty_unr_1 = 1:else:str_NewsProperty_unr_1 = 0:End if
	if str_NewsProperty_ann <>"" then:str_NewsProperty_ann_1 = 1:else:str_NewsProperty_ann_1 = 0:End if
	if str_NewsProperty_filt <>"" then:str_NewsProperty_filt_1 = 1:else:str_NewsProperty_filt_1 = 0:End if
	if str_NewsProperty_Remote <>"" then:str_NewsProperty_Remote_1 = 1:else:str_NewsProperty_Remote_1 = 0:End if
	obj_Save_Rs("NewsProperty") = str_NewsProperty_Rec_1&","&str_NewsProperty_mar_1&","&str_NewsProperty_rev_1&","&str_NewsProperty_constr_1&","&str_NewsProperty_Remote_1&","&str_NewsProperty_tt_1&","&str_NewsProperty_hots_1&","&str_NewsProperty_jc_1&","&str_NewsProperty_unr_1&","&str_NewsProperty_ann_1&","&str_NewsProperty_filt_1
	obj_Save_Rs("isRecyle") = 0
	obj_Save_Rs("addtime") = str_addtime
	'辅助字段信息保存
	if str_News_Action = "add_Save" then
		dim obj_c_rs,i 
		If IsArray(CustColumnArr) Then
			set obj_c_rs = Server.CreateObject(G_FS_RS)
			obj_c_rs.open "select * From [FS_MF_DefineData] where 1=2",Conn,1,3
			For i = 0 to UBound(CustColumnArr,2)
				If Request.Form("FS_NS_Define_"&CustColumnArr(3,i))<>"" Then
					obj_c_rs.Addnew
					obj_c_rs("InfoType") = "NS"
					obj_c_rs("InfoID") = NoSqlHack(str_NewsID)
					obj_c_rs("TableEName")	=CustColumnArr(3,i)
					obj_c_rs("ColumnValue")=NoSqlHack(Request.Form("FS_NS_Define_"&CustColumnArr(3,i)))
					obj_c_rs.Update
				End If
			Next
			obj_c_rs.Close:Set obj_c_rs = nothing
		End If
	else
		'CustSql="select DefineID,ClassID,D_Name,D_Coul,D_Type,D_isNull,D_Value,D_Content,D_SubType from [FS_MF_DefineTable] Where D_SubType='NS' and  Classid="& trim(Request.Form("d_Id")) &""
		If IsArray(CustColumnArr) Then
			Dim SaveAuxiSql
			For i = 0 to Ubound(CustColumnArr,2)
				set obj_c_rs = Server.CreateObject(G_FS_RS)
				SaveAuxiSql="select * From [FS_MF_DefineData] where InfoID='"&NoSqlHack(str_NewsID)&"' and TableEName='" & NoSqlHack(CustColumnArr(3,i)) & "' and InfoType='NS'"
				obj_c_rs.Open SaveAuxiSql,Conn,1,3
				If obj_c_rs.Eof Then
					If Request.Form("FS_NS_Define_"&CustColumnArr(3,i))<>"" Then
						obj_c_rs.Addnew
						obj_c_rs("InfoType") = "NS"
						obj_c_rs("TableEName")=NoSqlHack(CustColumnArr(3,i))
						obj_c_rs("ColumnValue")=NoSqlHack(Request.Form("FS_NS_Define_"&CustColumnArr(3,i)))
						obj_c_rs("InfoID") = NoSqlHack(str_NewsID)
						obj_c_rs.Update
					End If
				Else
					obj_c_rs("ColumnValue")=NoSqlHack(Request.Form("FS_NS_Define_"&CustColumnArr(3,i)))
					obj_c_rs.Update
				End If
				obj_c_rs.Close:set obj_c_rs=nothing
			Next
		End If
	end if
	if Trim(str_isdraft)<>"" then:obj_Save_Rs("isdraft")=1:else:obj_Save_Rs("isdraft")=0:end if
	
	If IsAdPic=1 Then
	    if Cintstr(Request.Form("Checkbox1"))=1 then
		    obj_Save_Rs("IsAdPic")=1
		    obj_Save_Rs("AdPicWH")=NoSqlHack(AdPicWH)
		    obj_Save_Rs("AdPicLink")=NoSqlHack(AdPicLink)
		    obj_Save_Rs("AdPicAdress")=NoSqlHack(AdPicAdress)
		 end if
		 if Cintstr(Request.Form("Checkbox2"))=1 then
		    obj_Save_Rs("IsAdPic")=2
		    obj_Save_Rs("AdPicWH")=NoSqlHack(AdPicWHw)
		 end if		
	Else
		obj_Save_Rs("IsAdPic")=0
	End If
	obj_Save_Rs.update
	Dim Get_News_ID,rssql '取自动编号ID
	if G_IS_SQL_DB = 0 then
		Get_News_ID = obj_Save_Rs("ID")
	Else
		if str_News_Action = "add_Save" then
			set rssql = Conn.execute("SELECT ident_current('FS_NS_News')")
			Get_News_ID = rssql(0)
			rssql.close:set rssql = nothing
		else
			Get_News_ID = obj_Save_Rs("ID")
		end if
	End if
	obj_Save_Rs.close:set obj_Save_Rs = nothing 
	if str_NewsType<>"TitleNews" then
		'开始保存关键字，作者，来源等  
		Dim obj_save_Gener_Rs,obj_save_Gener_Rs1,obj_save_Gener_Rs2,obj_TF_Rs,tmp
		if str_KeywordSaveTF <>"" then
			for each tmp in split(NoSqlHack(Request.Form("KeywordText")),",")
			  if tmp <> "" then 
					Set obj_save_Gener_Rs=server.CreateObject(G_FS_RS)
					  obj_save_Gener_Rs.open "select Gid,G_Type,G_Name From [FS_NS_General] where G_Type=1 and G_Name='"&NoSqlHack(tmp)&"'",Conn,1,3
					  If obj_save_Gener_Rs.eof Then
						obj_save_Gener_Rs.Addnew
						obj_save_Gener_Rs("G_Name") = NoSqlHack(tmp)
						obj_save_Gener_Rs("G_Type") = 1
						obj_save_Gener_Rs.Update
					  End If
					  obj_save_Gener_Rs.Close:set obj_save_Gener_Rs = nothing
			  end if
			next
		End if
		'保存作者
		if str_AuthorSaveTF <>"" then
			Set obj_save_Gener_Rs1=server.CreateObject(G_FS_RS)
			obj_save_Gener_Rs1.Open "select Gid,G_Type,G_Name from FS_NS_General where G_Type=3 and G_Name='"& NoSqlHack(Request.Form("Author")) &"'",Conn,3,3
			if obj_save_Gener_Rs1.eof then:obj_save_Gener_Rs1.addnew:obj_save_Gener_Rs1("G_Type") =3:obj_save_Gener_Rs1("G_Name") = NoSqlHack(Request.Form("Author")):obj_save_Gener_Rs1.update:end if:set obj_save_Gener_Rs1 = nothing
		End if
		'保存来源
		if str_SourceSaveTF <>"" then
			Set obj_save_Gener_Rs2=server.CreateObject(G_FS_RS)
			obj_save_Gener_Rs2.Open "select Gid,G_Type,G_Name from FS_NS_General where G_Type=2 and G_Name='"& NoSqlHack(Request.Form("Source")) &"'",Conn,3,3
			if obj_save_Gener_Rs2.eof then:obj_save_Gener_Rs2.addnew:obj_save_Gener_Rs2("G_Type") =2:obj_save_Gener_Rs2("G_Name") = NoSqlHack(Request.Form("Source")):obj_save_Gener_Rs2.update:end if:set obj_save_Gener_Rs2 = nothing
		End if
		'如果生成文件名规则中有：新闻自动编号ID 
		Dim TempRsObj
		If Instr(str_FileName,"自动编号ID") Then
			str_FileName = Replace(str_FileName,"自动编号ID",Get_News_ID)
			Set TempRsObj=server.CreateObject(G_FS_RS)
			TempRsObj.open "select FileName From [Fs_NS_News] where NewsID='"&NoSqlHack(str_NewsID)&"' and ID="&CintStr(Get_News_ID)&"",Conn,1,3
			if not TempRsObj.eof Then
				TempRsObj("FileName") = Replace(TempRsObj("FileName"),"自动编号ID",Get_News_ID)
				TempRsObj.update
			End If
			TempRsObj.Close
		End IF
		Dim TempRsObj_1
		If Instr(str_FileName,"唯一NewsID") Then
			str_FileName = Replace(str_FileName,"唯一NewsID",str_NewsID)
			Set TempRsObj_1=server.CreateObject(G_FS_RS)
			TempRsObj_1.open "select FileName From [Fs_NS_News] where NewsID='"&NoSqlHack(str_NewsID)&"'",Conn,1,3
			if not TempRsObj_1.eof Then
				TempRsObj_1("FileName") = Replace(TempRsObj_1("FileName"),"唯一NewsID",str_NewsID)
				TempRsObj_1.update
			End If
			TempRsObj_1.Close
		End IF
	End if
	if Trim(str_GroupID) <>"" or str_PointNumber <> "" or str_Money<>"" then 
		Dim obj_insert_rs
		set obj_insert_rs = Server.CreateObject(G_FS_RS)
		if str_News_Action = "add_Save" then
			obj_insert_rs.Open "select  GroupName,PointNumber,FS_Money,InfoID,PopType,isClass From FS_MF_POP",Conn,1,3
			obj_insert_rs.addnew
		Elseif str_News_Action = "Edit_Save" then
			obj_insert_rs.Open "select  GroupName,PointNumber,FS_Money,InfoID,PopType,isClass From FS_MF_POP where InfoID='"& NoSqlHack(str_NewsID) &"' and PopType='NS' and isClass=0",Conn,1,3
			if obj_insert_rs.eof then
				obj_insert_rs.addnew
			end if
		End if
		obj_insert_rs("GroupName")=str_GroupID
		if str_PointNumber <>""  then:obj_insert_rs("PointNumber")=str_PointNumber:Else:obj_insert_rs("PointNumber")=0:End if
		if str_Money <>"" then:obj_insert_rs("FS_Money")=str_Money:Else:obj_insert_rs("FS_Money")=0:End if
		obj_insert_rs("InfoID")=str_NewsID
		obj_insert_rs("PopType")="NS"
		obj_insert_rs("isClass")=0
		obj_insert_rs.update
		obj_insert_rs.close:set obj_insert_rs = nothing
	End if
	'保存图片头条,插入数据库
	sub saveTody()
		if str_TodayNewsPicTF<>"" then
			dim insertTodayPicobj
			set insertTodayPicobj = Server.CreateObject(G_FS_RS)
			insertTodayPicobj.open "select * From FS_NS_TodayPic where NewsId='"& NoSqlHack(str_NewsID) &"'",Conn,1,3
			if  insertTodayPicobj.eof then
				insertTodayPicobj.addnew
				insertTodayPicobj("NewsID")= str_NewsID
				insertTodayPicobj("TodayPic_font")= NoSqlHack(Request.Form("FontFace"))
				insertTodayPicobj("TodayPic_size")= NoSqlHack(Request.Form("FontSize"))
				insertTodayPicobj("TodayPic_color")= NoSqlHack(Request.Form("FontColor"))
				insertTodayPicobj("TodayPic_space")= NoSqlHack(Request.Form("FontSpace"))
				insertTodayPicobj("TodayPic_PicColor")= NoSqlHack(Request.Form("FontBgColor"))
				insertTodayPicobj("ClassID")= str_ClassID
				insertTodayPicobj("TodayPic_SavePath")= str_NewsID
				insertTodayPicobj("TodayTitle")= NoSqlHack(Request.Form("PicTitle"))
				if not isnumeric(Request.Form("PicTitlewidth")) then
					insertTodayPicobj("Todaywidth")= 300
				else
					insertTodayPicobj("Todaywidth")= CintStr(Request.Form("PicTitlewidth"))
				end if
				insertTodayPicobj("addtime")= now
				insertTodayPicobj.update
				insertTodayPicobj.close:set insertTodayPicobj = nothing
			Else
				insertTodayPicobj("TodayPic_font")= NoSqlHack(Request.Form("FontFace"))
				insertTodayPicobj("TodayPic_size")= NoSqlHack(Request.Form("FontSize"))
				insertTodayPicobj("TodayPic_color")= NoSqlHack(Request.Form("FontColor"))
				insertTodayPicobj("TodayPic_space")= NoSqlHack(Request.Form("FontSpace"))
				insertTodayPicobj("TodayPic_PicColor")= NoSqlHack(Request.Form("FontBgColor"))
				insertTodayPicobj("TodayPic_SavePath")= str_NewsID 
				insertTodayPicobj("addtime")= now
				insertTodayPicobj("TodayTitle")= NoSqlHack(Request.Form("PicTitle"))
				if not isnumeric(Request.Form("PicTitlewidth")) then
					insertTodayPicobj("Todaywidth")= 300
				else
					insertTodayPicobj("Todaywidth")= CintStr(Request.Form("PicTitlewidth"))
				end if
				insertTodayPicobj.update
				insertTodayPicobj.close:set insertTodayPicobj = nothing
			End if
		End if
	end sub
	'生成静态文件开始
	If Not str_NewsType="TitleNews" Then
		Call Refresh("NS_news",Get_News_ID)
	End If
	'*****保留
	if Trim(str_isdraft)<>"" then
		str_isDraftTF="<li>"& Fs_news.allInfotitle &"已经保存到草稿箱中</li><li>"& Fs_news.allInfotitle &"未发布</li>"
		str_ReturnUrl ="../News_MyFolder.asp?Action=draft"
	else
		if str_News_Action = "Edit_Save" then
			str_isDraftTF=""
		Else
			str_isDraftTF="<li>"& Fs_news.allInfotitle &"添加成功!</li>"
		End if
		str_ReturnUrl = "../news_Manage.asp?ClassID="&str_ClassID&""
	end if
	if str_News_Action = "add_Save" then
		if str_checkTF = true then
			strShowErr = ""& str_isDraftTF &"<li>系统设置(或者栏目)为需要审核后才能发布到前台</li><br><li><a href=""../News_Add.asp?ClassID="& str_ClassID &"""><b>继续添加</b></a>&nbsp;&nbsp;<b><a href=../News_Edit.asp?ClassID="&NoSqlHack(str_Classid)&"&NewsID="&NoSqlHack(str_NewsID)&"><font color=red>返回编辑 </font></a></b>&nbsp;&nbsp;<a href=""../News_Manage.asp?ClassID="& NoSqlHack(str_ClassID) &"""><b>返回管理</b></a></li>"
		Else
			strShowErr = ""& str_isDraftTF &"<br><li><a href=""../News_Add.asp?ClassID="& str_ClassID &"""><b>继续添加</b></a>&nbsp;&nbsp;<b><a href=../News_Edit.asp?ClassID="&NoSqlHack(str_ClassID)&"&NewsID="&NoSqlHack(str_NewsID)&"><font color=red>返回编辑 </font></a></b>&nbsp;&nbsp;<a href=""../News_Manage.asp?ClassID="& NoSqlHack(str_ClassID) &"""><b>返回管理</b></a></li>"
		End if
		if str_TodayNewsPicTF<>"" then
			If tmp_returnvalue = False Then
				strShowErr = ""& str_isDraftTF &"<li>新闻添加成功</li><li><font color=red>注意：头条图片生成失败,没有开启水印组件或者水印组件异常!!</font></li><br><li><a href=""../News_Add.asp?ClassID="& str_ClassID &"""><b>继续添加</b></a>&nbsp;&nbsp;<b><a href=../News_Edit.asp?ClassID="&str_ClassID&"&NewsID="&str_NewsID&"><font color=red>返回编辑 </font></a></b>&nbsp;&nbsp;<a href=""../News_Manage.asp?ClassID="& str_ClassID &"""><b>返回管理</b></a></li>"
			End If	
		end if
	Elseif str_News_Action = "Edit_Save" then
		strShowErr = ""& str_isDraftTF &"<li>修改成功!!</li><li><a href=""../News_Edit.asp?NewsID="& str_NewsID &"&ClassID="&str_ClassID&"""><b>继续修改新闻?</b></a>&nbsp;&nbsp;<a href=""../News_Manage.asp?ClassID="& str_ClassID &"""><b>返回新闻管理?</b></a></li>"
		if str_TodayNewsPicTF<>"" then
			If tmp_returnvalue = False Then
				strShowErr = ""& str_isDraftTF &"<li>修改成功</li><li><font color=red>注意：头条图片生成失败,没有开启水印组件或者水印组件异常!!</font></li><li><a href=""../News_Edit.asp?NewsID="& str_NewsID &"&ClassID="&str_ClassID&"""><b>继续修改新闻??</b></a>&nbsp;&nbsp;<a href=""../News_Manage.asp?ClassID="& str_ClassID &"""><b>返回新闻管理??</b></a></li>"
			End If	
		end if
	End if
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="& str_ReturnUrl &"")
	Response.end
%>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->
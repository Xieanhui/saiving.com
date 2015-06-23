<%
'脚本超时
Server.ScriptTimeout=600
'=========================================
'Writes函数仅仅用来断点测试不做其它用途
'=========================================
Function Writes(Str,types)
    Select case types
        Case "1"
            Response.Write(Str)
            Response.End
        Case "2"
            Response.Write(ubound(Str))
            Response.End
   End Select
End Function
Function Add_Root_Dir(f_Path)
	Dim f_All_Path
	If Left(f_Path,1)="/" Then
		f_All_Path = G_VIRTUAL_ROOT_DIR & f_Path
	Else
		f_All_Path = G_VIRTUAL_ROOT_DIR & "/" & f_Path
	End If
	If Trim(G_VIRTUAL_ROOT_DIR) <> "" Then
		f_All_Path = "/" & f_All_Path
	End If
	Add_Root_Dir = f_All_Path
End Function

Function FiltBad(Str)
    dim badStr
	badStr=Str
    if Str="" then
        FiltBad=""
   else
        dim badConst,i,badtm
		badConst = Split(G_Badwords,",")
	    For i = 0 To Ubound(badConst)
		   badtm=Split(badConst(i),"|")
		   badStr  = Replace(badStr,trim(badtm(0)),badtm(1))
	    Next
        FiltBad=badStr
   end if
end Function

Function FiltBad1(Str)
    dim badStr
    FiltBad1=false
    if Str<>"" then
        dim badConst,i,badtm
		badConst = Split(G_Badwords,",")
	    For i = 0 To Ubound(badConst)
		    badtm=Split(badConst(i),"|")
	        if instrRev(Str,badtm(0))>0 then
		      FiltBad1  = true
		   end if
	    Next
   end if
end Function


Function Lose_Html(f_Str)
	Dim regEx
	if Not IsNull(f_Str) Then
		f_Str=f_Str&""
		Set regEx = New RegExp
		regEx.Pattern = "<\/*[^<>]*>"
		regEx.IgnoreCase = True
		regEx.Global = True
		f_Str = regEx.Replace(f_Str,"")
		Lose_Html = f_Str
	Else
		Lose_Html=""
	End If
End Function

Function Intercept_Char(f_Str,f_Length,f_Flag)
	'f_Flag为1，一个中文字符的长度算1；f_Flag为2，一个中文字符的长度算2
	Dim f_Str_Total_Len,f_i,f_Str_Curr_Len,f_One_Char
	If f_Length = 0  Or f_Str = "" Or IsNull(f_Str) Then
		Intercept_Char = ""
		Exit Function
	End If
	f_Str=Replace(Replace(Replace(Replace(f_Str,"&nbsp;"," "),"&quot;",Chr(34)),"&gt;",">"),"&lt;","<")
	f_Str_Total_Len = Len(f_Str)
	If f_Flag = 1 Then
		If f_Length>=f_Str_Total_Len Then
			Intercept_Char = f_Str
		Else
			Intercept_Char = Left(f_Str,f_Length)
		End If
	Else
		For f_i = 1 To f_Str_Total_Len
			f_One_Char = Mid(f_Str,f_i,1)
			If Abs(Asc(f_One_Char)) > 255 then
				f_Str_Curr_Len=f_Str_Curr_Len+2
			Else
				f_Str_Curr_Len=f_Str_Curr_Len+1
			End If
			If f_Str_Curr_Len >= f_Length Then
				Intercept_Char = Left(f_Str,f_i)
				Exit For
			End If
		Next
		If f_Str_Curr_Len < f_Length Then
			Intercept_Char = f_Str
		End If
	End If
	Intercept_Char = Replace(Replace(Replace(Replace(Intercept_Char," ","&nbsp;"),Chr(34),"&quot;"),">","&gt;"),"<","&lt;")
End Function

Function Mod_IS_Installed_Bool(f_Mod_Str)
	On Error Resume Next
	Mod_IS_Installed_Bool = False
	Err = 0
	Dim f_TestObj
	Set f_TestObj = Server.CreateObject(f_Mod_Str)
	If Err = 0 Then
		Mod_IS_Installed_Bool = True
	End If
	Set f_TestObj = Nothing
	Err = 0
End Function

Function SendMail(f_Mailto_Address,f_Mailto_Name,f_Subject,f_Mail_Body,f_From_Name,f_Mail_From,f_Priority)
	On Error Resume Next
	Dim f_JMail,f_True_Mail_From,f_Mail_Server,f_Server_Domain
	Set f_JMail=Server.CreateObject(G_JMAIL_MESSAGE)
	If Err Then
		SendMail= "<br><li>没有安装JMail组件</li>"
		Err.Clear
		Exit Function
	End If
	f_Mail_Server = Get_Cache_Value("MF","MF_Mail_Server")
	f_True_Mail_From = Get_Cache_Value("MF","MF_Mail_Name")
	f_JMail.Silent = True
	f_JMail.Logging = True
	f_JMail.Charset = "gb2312"
	f_JMail.MailServerUserName = f_True_Mail_From
	f_JMail.MailServerPassword = Get_Cache_Value("MF","MF_Mail_Pass_Word")
	f_JMail.ContentType = "text/html"
	f_True_Mail_From =f_True_Mail_From & "@"
	f_Server_Domain = Left(f_Mail_Server,InStrRev(f_Mail_Server,".")-1)
	f_Server_Domain = Left(f_Server_Domain,InStrRev(f_Server_Domain,"."))
	f_True_Mail_From =f_True_Mail_From & Right(f_Mail_Server,Len(f_Mail_Server)-Len(f_Server_Domain))
	f_JMail.From = f_True_Mail_From
	f_JMail.FromName = f_From_Name & "(" & f_Mail_From & ")"
	f_JMail.Subject = f_Subject
	f_JMail.AddRecipient f_Mailto_Address
	f_JMail.Body = f_Mail_Body
	f_JMail.Priority = 3
	f_JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
	f_JMail = ObjJmail.Send(f_Mail_Server)
	f_JMail.Close
	Set f_JMail=nothing
End Function

'======================================
'在SQL语句组合前运行
'======================================
Function NoSqlHack(FS_inputStr)
	FS_inputStr = Trim(FS_inputStr)
	If FS_inputStr = "" Or Isnull(FS_inputStr) Then
	    FS_inputStr = ""
	End if
	FS_inputStr = Replace(FS_inputStr,Chr(39),"&#39;")			'单引号
    'FS_inputStr = Replace(FS_inputStr,";","")
    'FS_inputStr = Replace(FS_inputStr," ","")
    'FS_inputStr = Replace(FS_inputStr,"%","")
    'FS_inputStr = Replace(FS_inputStr,"&nbsp;","")
	NoSqlHack = FS_inputStr
End Function

'======================================
'格式化字符逗号分隔
'参数：str－需要处理的字符串
'======================================
Function FormatStrArr(Oclassid)
	If Oclassid <> "" and not Isnull(Oclassid) Then
		Oclassid = Split(Oclassid,",")
		Dim Oclassids,i
	    For i = 0 To Ubound(Oclassid)
		    Oclassids = Oclassids&NoSqlHack(Oclassid(i))&"','"
	    Next
		If trim(Oclassids)<>"" then  Oclassids=mid(Oclassids,1,len(Oclassids)-3)
	End if
	FormatStrArr = Oclassids
End Function

'======================================
'格式化整数逗号分隔
'参数：str－需要处理的字符串
'======================================
Function FormatIntArr(str)
    dim arr,i,re
    if trim(str)="" Then
        FormatIntArr = ""   '如果为空返回原值按业务逻辑处理
        Exit Function
    End IF
    re = "0"
    if instr(1,str,",")=0 then
        re = CintStr(str)
    else
        arr = split(str,",")
        for i=0 to ubound(arr)
            if IsNumeric(arr(i)) Then re = re & "," & CintStr(arr(i))
        next
        if re<>"0" then re = mid(re,3)
    end if
    FormatIntArr = re
End Function

'/////////////////////////////////////////////////////
'整型转换
'////////////////////////////////////////////////////
Function CintStr(Intstr)
	On Error Resume Next
	If clng(Intstr) <= 0 Then
		CintStr = 0
	Else
		CintStr = clng(Intstr)
	End if
	If Err Then
		Err.clear()
		CintStr = 0
	End if
End Function
'/////////////////////////////////////////////////////
'参数长度检测
'Parastr被检测的参数
'参数对比长度
'////////////////////////////////////////////////////
Function Lenstr(Parastr,Lengthstr)
	Parastr = NoSqlHack(Parastr)
	If Trim(Parastr) = "" Then Exit Function
		If Len(Parastr)>Lengthstr Then
			Lenstr = false Rem 如果Len(Parastr)长度超出规定长度返回false
		Else
			Lenstr = Parastr
		End if
End Function

Function CheckIpSafe(ip)
	Dim test,test_i,test_j,ascnum,safe,iplen
	test=Split(ip,".")
	safe=True
	For test_i=LBound(test) To UBound(test)
		iplen=Len(test(test_i))
		For test_j=1 To iplen
			ascnum=Asc(Mid(test(test_i),test_j,1))
			If Not (ascnum>=48 And ascnum<=57) Then
				Response.Write "<html><title>警告</title><body bgcolor=""EEEEEE"" leftmargin=""60"" topmargin=""30""><font style=""font-size:16px;font-weight:bolder;color:blue;""><li>您提交的数据有恶意字符</li></font><font style=""font-size:14px;font-weight:bolder;color:red;""><br><li>您的数据已经被记录!</li><br><li>您的IP："&Request.ServerVariables("Remote_Addr")&"</li><br><li>操作日期："&Now&"</li></font></body></html><!--Powered by Foosun Inc.,AddTime:"&now&"-->"
				Response.End
			End If
		Next
	Next
	CheckIpSafe=ip
End Function
Function NoHtmlHackInput(Str) '过滤跨站脚本和HTML标签
	Dim regEx
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Pattern = "<|>|(script)|on(mouseover|mouseon|mouseout|click|dblclick|blur|focus|change)|eval|\t"
	If regEx.Test(LCase(Str)) Then
			Response.Write "<html><title>警告</title><body bgcolor=""EEEEEE"" leftmargin=""60"" topmargin=""30""><font style=""font-size:16px;font-weight:bolder;color:blue;""><li>您提交的数据有恶意字符</li></font><font style=""font-size:14px;font-weight:bolder;color:red;""><br><li>提交的内容不能包括[<|>|(script)|on(mouseover|mouseon|mouseout|click|dblclick|blur|focus|change)|eval]</li><li>您的数据已经被记录!</li><br><li>您的IP："&Request.ServerVariables("Remote_Addr")&"</li><br><li>操作日期："&Now&"</li></font></body></html><!--Powered by Foosun Inc.,AddTime:"&now&"-->"
			Response.End
	End If
	Set regEx = Nothing
	NoHtmlHackInput = Str
End Function

'Function GotTopic(Str,StrLen)
'	Dim l,t,c, i,LableStr,regEx,Match,Matches
'	If StrLen=0 then
'		GotTopic=""
'		exit function
'	End If
'	if IsNull(Str) then
'		GotTopic = ""
'		Exit Function
'	end if
'	if Str = "" then
'		GotTopic=""
'		Exit Function
'	end If
'	'Str=Replace(Replace(Replace(Replace(Str,"&nbsp;"," "),"&quot;",Chr(34)),"&gt;",">"),"&lt;","<")
'    'Str=Replace(Replace(Replace(Replace(Str, "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
'	l=len(str)
'	t=0
'	strlen=Clng(strLen)
'	for i=1 to l
'		c=Abs(Asc(Mid(str,i,1)))
'		if c>255 then
'			t=t+2
'		else
'			t=t+1
'		end if
'		if t>=strlen then
'			GotTopic=left(str,i)
'			exit for
'		else
'			GotTopic=str
'		end if
'	Next
'	GotTopic = GotTopic'Replace(Replace(Replace(Replace(GotTopic," ","&nbsp;"),Chr(34),"&quot;"),">","&gt;"),"<","&lt;")
'End Function

'**************************************************
'函数名：GotTopic
'作  用：截字符串，汉字一个算两个字符，英文算一个字符
'参  数：str   ----原字符串
'        strlen ----截取长度
'        bShowPoint ---- 是否显示省略号
'返回值：截取后的字符串
'**************************************************
Function GotTopic(ByVal str, ByVal strlen)
	dim bShowPoint
	bShowPoint=false
    If IsNull(str) Or str = ""  Then
        GotTopic = ""
        Exit Function
    End If
    Dim l, t, c, i, strTemp

	str=replace(str,"&lt;","<")
	str=replace(str,"&gt;",">")
    str=replace(str,"&nbsp;"," ")
	str=replace(str,"&quot;",Chr(34))
	str=replace(str,"&#39;",Chr(39))
	str=replace(str,"&mdash;","—")
	str=replace(str,"&ldquo;","“")
	str=replace(str,"&rdquo;","”")

    l = Len(str)
    t = 0
    strTemp = str
    strlen = Clng(strlen)
    For i = 1 To l
        c = Abs(Asc(Mid(str, i, 1)))
        If c > 255 Then
            t = t + 2
        Else
            t = t + 1
        End If
        If t >= strlen Then
            strTemp = Left(str, i)
            Exit For
        End If
    Next

'	str=replace(str,"<","&lt;")
'	str=replace(str,">","&gt;")
'   str=replace(str," ","&nbsp;")
'	str=replace(str,Chr(34),"&quot;")
'	str=replace(str,Chr(39),"&#39;")
'	str=replace(str,"—","&mdash;")
'	str=replace(str,"“","&ldquo;")
'	str=replace(str,"”","&rdquo;")
'
'	strTemp=replace(strTemp,"<","&lt;")
'	strTemp=replace(strTemp,">","&gt;")
'    strTemp=replace(strTemp," ","&nbsp;")
'	strTemp=replace(strTemp,Chr(34),"&quot;")
'	strTemp=replace(strTemp,Chr(39),"&#39;")
'	strTemp=replace(strTemp,"—","&mdash;")
'	strTemp=replace(strTemp,"“","&ldquo;")
'	strTemp=replace(strTemp,"”","&rdquo;")

    If strTemp <> str And bShowPoint = True Then
        strTemp = strTemp & "…"
    End If
    GotTopic = strTemp
End Function



'返回中文字符的前StrLen位字符 By Wen Yongzhong
Function GetCStrLen(Str,StrLen)
	Dim l,t,c, i,LableStr,regEx,Match,Matches
	If StrLen=0 Then
		GetCStrLen=""
		Exit Function
	End If
	If IsNull(Str) Then
		GetCStrLen = ""
		Exit Function
	End If
	If Str = "" Then
		GetCStrLen=""
		Exit Function
	End If
	l=len(str)
	t=0
	strlen=Clng(strLen)
	For i=1 To l
		c=Abs(Asc(Mid(str,i,1)))
		If c>255 Then
			t=t+2
		Else
			t=t+1
		End If
		If t>=strlen Then
			GetCStrLen=left(str,i)
			Exit For
		Else
			GetCStrLen=str
		End If
	Next
End Function
'远程存图
Function ReplaceRemoteUrl(NewsContent,SaveFilePath,FunDoMain,DummyPath)
	Dim re,RemoteFile,RemoteFileurl,SaveFileName,FileName,FileExtName,SaveImagePath,tNewsContent
	Set re = New RegExp
	re.IgnoreCase = True
	re.Global=True
	re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}((\w)+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(gif|jpg|png|bmp|swf|flv|mp3|wma)))"
	tNewsContent = NewsContent
	Set RemoteFile = re.Execute(tNewsContent)
	Set re = Nothing
	For Each RemoteFileurl in RemoteFile
		SaveFileName = Mid(RemoteFileurl,InstrRev(RemoteFileurl,"/")+1)
		Call SaveRemoteFile(SaveFilePath & "/" & SaveFileName,RemoteFileurl,1)
		'Call SaveRemoteFile(DummyPath & SaveFilePath & "/" & SaveFileName,RemoteFileurl)
		'tNewsContent = Replace(tNewsContent,RemoteFileurl,FunDoMain & SaveFilePath & "/" & SaveFileName)
		tNewsContent = Replace(tNewsContent,RemoteFileurl,SaveFilePath & "/" & SaveFileName)
	Next
	ReplaceRemoteUrl = tNewsContent
End Function

Sub SaveRemoteFile(LocalFileName,RemoteFileUrl,WTF)
	LocalFileName=Server.MapPath(replace(LocalFileName,"//","/"))
	'PathExistCheck LocalFileName
	On Error Resume Next
	Dim StreamObj,Retrieval,GetRemoteData
	Set Retrieval = Server.CreateObject(G_FS_XMLHTTP)
	If Err Then
		Response.Write "<script language='JavaScript'>alert('你的系统不支持"&G_FS_XMLHTTP&"\n，无法保存远程文件！');</script>"
		Err.clear
		Set Retrieval = Nothing
		Exit Sub
	End If
	With Retrieval
		.Open "Get", RemoteFileUrl, False, "", ""
		.Send
		if Err then
			Response.Write "<script language='JavaScript'>alert('目标服务器不支持"&G_FS_XMLHTTP&"\n，无法保存远程文件！');</script>"
			Err.Clear
			Set Retrieval = Nothing
			Exit Sub
		end if
		GetRemoteData = .ResponseBody
	End With
	Set Retrieval = Nothing
	If Err Then Err.clear
	Set StreamObj = Server.CreateObject(G_FS_STREAM)
	If Err Then
		Response.Write "<script language='JavaScript'>alert('你的系统不支持"&G_FS_STREAM&"\n，无法保存远程文件！');</script>"
		Err.clear
		Set StreamObj = Nothing
		Exit Sub
	End If
	With StreamObj
		.Type = 1
		.Open
		.Write GetRemoteData
		.SaveToFile LocalFileName,2
		.Cancel()
		.Close()
	End With
	IF WTF = 1 Then
		AddWaterMark LocalFileName
	End IF
	Set StreamObj = Nothing
End Sub
'创建
Function CreateDateDir(Path)
	Dim sBuild,FSO
	sBuild=path&"\"&year(Now())&"-"&month(now())
	Set FSO = Server.CreateObject(G_FS_FSO)
	If FSO.FolderExists(sBuild)=false then
		FSO.CreateFolder(sBuild)
	End IF
	sBuild=sBuild&"\"&day(Now())
	If FSO.FolderExists(sBuild)=false then
		FSO.CreateFolder(sBuild)
	End IF
	set FSO=Nothing
End Function

'创建目录
Sub savePathdirectory(Path)
	Dim FSO
	Set FSO = Server.CreateObject(G_FS_FSO)
	if Trim(G_VIRTUAL_ROOT_DIR) ="" then
		FSO.CreateFolder(Path)
	Else
		FSO.CreateFolder(G_VIRTUAL_ROOT_DIR)
		FSO.CreateFolder(Path)
	End if
End Sub

' 传入：字符串、位置、长度
' 返回：在字符串指定位置取出指定长度的字符串，如果位置大于等于字符串长度，返回空值
Function getStrLoc(FS_Str,FS_StrLoc,FS_StrLen)
	Dim FS_CharFind
	If Len(FS_Str)>=FS_StrLoc Then
		FS_CharFind = Mid(FS_Str,FS_StrLoc,FS_StrLen)
		getStrLoc = FS_CharFind
	Else
		getStrLoc = ""
	End If
End Function

'======================================================================
' 用AspJpeg组件建立带有新闻标题的图片
' 参数说明
' NumCanvasWidth画布宽度，NumCanvasHeight画布高度，bgColor背景颜色，borderColor图片边框颜色(为空或者0不显示边框)
' TextColor文字颜色,TextFamily文字字体，BoldTF是否粗体（1为加粗），TextSize文字大小，StrTitle文字内容
' NumTopMargin文字垂直距离画布的顶边距(横向默认是居中的)，StrSavePath图片保存路径（需要绝对路径）
' 测试代码如下：
'	AspJpegCreateTextPic 400,60,&Hcccccc,&H0000ff,&H000000,"宋体",1,40,"文字转换图片AspJpeg",8,server.mappath("frontpage.jpg")
'	response.write "<img src='frontpage.jpg'><br>"
'======================================================================
Function AspJpegCreateTextPic(NumCanvasWidth,NumCanvasHeight,bgColor,borderColor,TextColor,TextFamily,BoldTF,TextSize,StrTitle,NumTopMargin,StrSavePath)
	AspJpegCreateTextPic = true
	If GetIsOpenWater=True Then Exit Function
	If Not IsObjInstalled("Persits.Jpeg") Then
		AspJpegCreateTextPic = false
	else
		If IsExpired("Persits.Jpeg")=true Then
			AspJpegCreateTextPic = false
		else
			Dim Title,objJpeg,TitleWidth
			Title = StrTitle
			Set objJpeg = Server.CreateObject(G_PERSITS_JPEG)
			objJpeg.New NumCanvasWidth, NumCanvasHeight, bgColor
			If borderColor<>"" And borderColor<>0 Then
				objJpeg.Canvas.Pen.Color = borderColor
				objJpeg.Canvas.Brush.Solid = False
				objJpeg.Canvas.DrawBar 1, 1, objJpeg.Width, objJpeg.Height
			End If
			objJpeg.Canvas.Font.Color = "&H"&TextColor'&HFF0000
			objJpeg.Canvas.Font.Family = TextFamily
			If BoldTF=1 Then objJpeg.Canvas.Font.Bold = True
			objJpeg.Canvas.Font.Size = TextSize
			objJpeg.Canvas.Font.Quality = 4

			TitleWidth = objJpeg.Canvas.GetTextExtent( Title )
			objJpeg.Canvas.Print (objJpeg.Width-TitleWidth)/2, NumTopMargin, Title
			objJpeg.Save StrSavePath
			Set objJpeg = Nothing
		end if
	end if
End Function

'======================================================================
' 用WsImage组件建立带有新闻标题的图片
' 参数说明：
' NumCanvasWidth画布宽度，NumCanvasHeight画布高度，，TextColor文字颜色,TextFamily文字字体，TextSize文字大小
' NumRotation旋转角度（文字保持水平请填0），StrTitle文字内容
' NumLeft，文字水平与画布的左边距，NumTop文字垂直距离画布的顶边距，StrSavePath图片保存路径（需要绝对路径）
' 返回值：
' 如果发生错误，返回错误代码
' 测试代码如下：
'	x = WsImgWatermarkText(440,300,&H0000FF&,"宋体",20,0,110,300,"测试水印WsImage",server.MapPath("apple111.jpg"))
'	response.write x&server.mappath("../admin/Images/wsimg.jpg")&"<br><img src='../admin/Images/wsimg.jpg'><img src='apple111.jpg'>"
'======================================================================

Function WsImgWatermarkTextToPic(NumCanvasWidth,NumCanvasHeight,TextColor,TextFamily,TextSize,NumRotation,NumLeft,NumTop,StrTitle,StrSavePath)
	WsImgWatermarkTextToPic = true
	If GetIsOpenWater=True Then Exit Function
	On Error Resume Next
	Dim StrPicPath
	If Not IsObjInstalled("wsImage.Resize") then
		WsImgWatermarkTextToPic = false
	else
		If IsExpired("wsImage.Resize")=true  then
			WsImgWatermarkTextToPic = false
		else
			StrPicPath = server.mappath("../Images/wsimg.jpg")
			WsImgIndentPicSize1 StrPicPath,NumCanvasWidth,NumCanvasHeight
			Dim objWsImg,strError
			set objWsImg=Server.CreateObject(G_WSIMAGE_RESIZE)
			objWsImg.LoadSoucePic StrPicPath
			objWsImg.Quality=75
			objWsImg.TxtMarkFont = TextFamily
			objWsImg.TxtMarkBond = false
			objWsImg.MarkRotate = NumRotation
			objWsImg.TxtMarkHeight = TextSize
			objWsImg.AddTxtMark CStr(StrSavePath), StrTitle, TextColor, NumTop, NumLeft
			strError=objWsImg.errorinfo
			If strError<>"" Then WsImgIndentPicScale = strError
			objWsImg.free:Set objWsImg=Nothing
			IF Err Then
				WsImgWatermarkTextToPic=False
			End If
		end if
	end if
End Function
Function WsImgIndentPicSize1(StrPicPath,NumWidth,NumHeight)
	On Error Resume Next
	Dim objWsImg,strError,NumType
	NumType = 0
	If NumHeight<=0 Then NumHeight=0:NumType=1
	If NumWidth<=0 Then NumWidth=0:NumType=2
	set objWsImg=Server.CreateObject(G_WSIMAGE_RESIZE)
	objWsImg.LoadSoucePic CStr(StrPicPath)
	objWsImg.Quality=75
	objWsImg.OutputSpic CStr(StrPicPath),NumWidth,NumHeight,NumType
	strError=objWsImg.errorinfo
	If strError<>"" Then WsImgIndentPicSize1 = strError
	objWsImg.free:Set objWsImg=Nothing
End Function


'======================================================================
' 用SA-ImgWriter组件建立带有新闻标题的图片
' 参数说明
' NumCanvasWidth画布宽度，NumCanvasHeight画布高度，bgColor背景颜色
' TextColor文字颜色,TextFamily文字字体，TextSize文字大小，StrTitle文字内容
' NumleftMargin文字水平与画布的左边距，NumTopMargin文字垂直距离画布的顶边距，StrSavePath图片保存路径（需要绝对路径）
' 测试代码如下：
'	ImageGenCreateTextPic 420,60,rgb(255,255,255),rgb(0,0,0),"宋体",40,"文字转换图片ImageGen",8,8,server.mappath("frontpage.jpg")
'	response.write "<img src='frontpage.jpg'><br>"
'======================================================================
Function ImageGenCreateTextPic(NumCanvasWidth,NumCanvasHeight,bgColor,TextColor,TextFamily,TextSize,StrTitle,NumleftMargin,NumTopMargin,StrSavePath)
	ImageGenCreateTextPic = true
	If GetIsOpenWater=True Then Exit Function
	If Not IsObjInstalled("softartisans.ImageGen") Then
		ImageGenCreateTextPic=false
	else
		If IsExpired("softartisans.ImageGen")=true Then
			ImageGenCreateTextPic=false
		else
			Dim objImageGen,objFont
			Set objImageGen = Server.CreateObject(G_SOFTARTISANS_IMAGEGEN)
			'Response.Write "<br>"&NumCanvasWidth &"<br>"& NumCanvasHeight&"<br>"& bgColor
			'Response.end
			objImageGen.CreateImage NumCanvasWidth, NumCanvasHeight, bgColor	'rgb(255,255,255)注意格式
			Set objFont = objImagegen.Font
			objFont.name = TextFamily
			objFont.Color = TextColor	'rgb(0,0,0)	'注意格式
			objFont.height = TextSize
			objImageGen.DrawTextOnImage NumleftMargin, NumTopMargin, objImageGen.Width-NumleftMargin, objImageGen.Height-NumTopMargin, StrTitle
			'Response.Write "<br>" &StrSavePath
			objImageGen.SaveImage 0, 3, StrSavePath
			Set objFont = Nothing
			Set objImageGen = Nothing
		end if
	end if
End Function

Function GetStrLengthE(Str)
'按英文计算字符串的长度，计算头条新闻图片大小用
	Dim i,StrLenth
	For i = 1 to len(Str)
		If Abs(Asc(Mid(Str,i,1)))>255 Then
			StrLenth=StrLenth+1
		Else
			StrLenth=StrLenth+0.5
		End If
	Next
	GetStrLengthE=StrLenth
End Function


'判断组件是否可用
Function IsObjInstalled(strClassString)
	IsObjInstalled = False
	Dim xTestObj
	On Error Resume Next
	Set xTestObj = Server.CreateObject(strClassString)
	If Err Then
		IsObjInstalled = False
		Err.Clear
	Else
		IsObjInstalled = True
	End If
	Set xTestObj = Nothing
End Function
'组件是否过期
Function IsExpired(strClassString)
	IsExpired = True
	Dim xTestObj
	On Error Resume Next
	Err.Clear
	Set xTestObj = Server.CreateObject(strClassString)
	Select Case LCase(strClassString)
		Case "Persits.Jpeg"
			If DateDiff("s",xTestObj.Expires,now)<0 Then
				IsExpired = False
			End if
		Case "wsImage.Resize"
			If instr(xTestObj.errorinfo,"已经过期") = 0 Then
				IsExpired = False
			End if
		Case "softartisans.ImageGen"
			xTestObj.CreateImage 500, 500, rgb(255,255,255)
			If Err Then
				Err.Clear
				IsExpired = False
			End if
		Case Else
			IsExpired = False
	End Select

	Set xTestObj = Nothing
End Function

'去掉首尾,号
Function DelHeadAndEndDot(Str)
	Dim StrLen
	StrLen=Len(Str)
	if StrLen>0 then
		if instr(str,",")=1 then
			Str=right(str,StrLen-1)
		end if
		StrLen=Len(Str)
		if instrrev(str,",")=StrLen then
			Str=left(str,StrLen-1)
		end if
	end if
	DelHeadAndEndDot=Str
End Function

'验证字符串是否合法,匹配到即为合法
Function IsValidStr(Str,FilterStr)
	IsValidStr=False
	If Str<>"" Then
		Dim regEx
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Pattern = FilterStr
		If regEx.Test(LCase(Str)) Then
			IsValidStr=True
		End If
		Set regEx = Nothing
	End If
End Function
'检查是否外部输入
Function IsSelfRefer()
	Dim sHttp_Referer, sServer_Name
	sHttp_Referer = CStr(Request.ServerVariables("HTTP_REFERER"))
	sServer_Name = CStr(Request.ServerVariables("SERVER_NAME"))
	If Mid(sHttp_Referer, 8, Len(sServer_Name)) = sServer_Name Then
		IsSelfRefer = True
	Else
		IsSelfRefer = False
	End If
End Function

'得到多少位数的随机函数
Function GetRamCode(f_number)
	Randomize
	Dim f_Randchar,f_Randchararr,f_RandLen,f_Randomizecode,f_iR
	f_Randchar="0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
	f_Randchararr=split(f_Randchar,",")
	f_RandLen=f_number '定义密码的长度或者是位数
	for f_iR=1 to f_RandLen
		f_Randomizecode=f_Randomizecode&f_Randchararr(Int((21*Rnd)))
	next
	GetRamCode = f_Randomizecode
End Function

'检查英文名称是否合法
Function chkinputchar(f_char)
	Dim f_name, i, c
	f_name = f_char
	chkinputchar = True
	If Len(f_name) <= 0 Then
		chkinputchar = False
		Exit Function
	End If
	For i = 1 To Len(f_name)
	   c = Mid(f_name, i, 1)
		If InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ@,.0123456789|-_", c) <= 0  Then
		   chkinputchar = False
		Exit Function
	   End If
   Next
End Function

''替换成自己想显示的信息
''格式:Replacestr(Hs_Rs("FloorType"),"1:多层,2:单层")
''格式:Replacestr(Rs("Audited"),"1:已通过审核,0:<span class=""tx"">未通过审核</span>")
Function Replacestr(dbvalue,strlist)
	Dim f_oldstr,f_tmpstr,f_tmparr,f_tmparr1
	f_oldstr = strlist
	if isnull(dbvalue) then dbvalue=""
	f_tmparr = split(f_oldstr,",")
	for each f_tmpstr in f_tmparr
		f_tmparr1 = split(f_tmpstr,":")
		if ubound(f_tmparr1) = 1 then
			if trim(dbvalue) = trim(f_tmparr1(0)) then
				f_oldstr = trim(f_tmparr1(1)) : exit for
			elseif trim(f_tmparr1(0)) = "else" then
				f_oldstr = trim(f_tmparr1(1))
			else
				f_oldstr = dbvalue
			end if
		else
		end if
	next
	Replacestr = f_oldstr
End Function

''显示下拉
''格式PrintOption(rs("language"),":<font color=#999999>请选择</font>,英语:英语,日语:日语,法语:法语")
Function PrintOption(Equvalue,valuelist)
	Dim f_oldstr,f_tmpstr,f_tmparr,f_tmparr1,isselected
	isselected=false:f_oldstr=""
	if isnull(Equvalue) then Equvalue=""
	f_tmparr = split(valuelist,",")
	for each f_tmpstr in f_tmparr
		f_tmparr1 = split(f_tmpstr,":")
		if ubound(f_tmparr1) = 1 then
			if trim(Equvalue) = trim(f_tmparr1(0)) and isselected=false then
				f_oldstr = f_oldstr & "<option value="""&f_tmparr1(0)&""" selected>"&f_tmparr1(1)&"</option>"
				isselected=true
			elseif trim(f_tmparr1(0))+trim(f_tmparr1(1))<>"" then
				f_oldstr = f_oldstr & "<option value="""&f_tmparr1(0)&""">"&f_tmparr1(1)&"</option>"
			end if
		else
		end if
	next
	PrintOption = f_oldstr
End Function

''文本框查询处理，方式可 “A B*”“A *B*”“A B”
''查询的时候 FildValue为空，显示的时候的 FildValue 不为空，则会将关键字颜色替换
Function Search_TextArr(StrKey,FildName,FildValue)
	Dim StrTmp,ArrTmp,New_StrTmp,Bol_Xin
	StrTmp = "" : New_StrTmp = FildValue
	Bol_Xin = False
	ArrTmp = split(StrKey,chr(32))
	for each StrTmp in ArrTmp
	  if Trim(StrTmp)<>"" then
		if New_StrTmp <> "" then
			StrTmp = replace(StrTmp,"*","")
			New_StrTmp = replace(New_StrTmp,StrTmp,"<font color=""red"">"&StrTmp&"</font>")
		else
			if left(StrTmp,1) = "*" then StrTmp = "%"&mid(StrTmp,2) : Bol_Xin = True
			if right(StrTmp,1) = "*" then StrTmp = mid(StrTmp,1,len(StrTmp) - 1)&"%" : Bol_Xin = True
			if not Bol_Xin then StrTmp = "%"&StrTmp&"%"
			New_StrTmp = New_StrTmp & " And "&FildName&" like '"&StrTmp&"' "
		end if
	  end if
	  Bol_Xin = False
	next
	''去掉得sql模式时的第一个and
	if FildValue="" and New_StrTmp<>"" then New_StrTmp = " ("&mid(New_StrTmp,len(" And ")+1)&") "
	Search_TextArr = New_StrTmp
End Function

''暂不支持中文
'可配合server.URLEncode如：server.URLEncode(Encrypt(防止被转化成'报错。
Function Encrypt(ecode)
''加密
dim texts
dim i
for i=1 to len(ecode)
texts=texts & chr(asc(mid(ecode,i,1))+3)
next
Encrypt = texts
End Function
''暂不支持中文
Function Decrypt(dcode)
''解密
dim texts
dim i
for i=1 to len(dcode)
texts=texts & chr(asc(mid(dcode,i,1))-3)
next
Decrypt=texts
End Function

Function and_where(sql)
	if instr(lcase(sql)," where ")>0 then
		and_where = sql & " and "
	else
		and_where = sql & " where "
	end if
End Function

Function Get_Date(f_getDate,f_datestyle)
	dim tmp_f_datestyle
	tmp_f_datestyle = f_datestyle
	if instr(1,f_datestyle,"YY02",1)>0 then
		tmp_f_datestyle= replace(tmp_f_datestyle,"YY02",right(year(f_getDate),2))
	end if
	if instr(f_datestyle,"YY04")>0 then
		tmp_f_datestyle= replace(tmp_f_datestyle,"YY04",year(f_getDate))
	end if
	if instr(f_datestyle,"MM")>0 then
		if month(f_getDate)<10 then
			tmp_f_datestyle= replace(tmp_f_datestyle,"MM","0"&month(f_getDate))
		else
			tmp_f_datestyle= replace(tmp_f_datestyle,"MM",month(f_getDate))
		end if
	end if
	if instr(f_datestyle,"DD")>0 then
		if day(f_getDate)<10 then
			tmp_f_datestyle= replace(tmp_f_datestyle,"DD","0"&day(f_getDate))
		else

			tmp_f_datestyle= replace(tmp_f_datestyle,"DD",day(f_getDate))
		end if
	end if
	if instr(f_datestyle,"HH")>0 then
		if hour(f_getDate)<10 then
			tmp_f_datestyle= replace(tmp_f_datestyle,"HH","0"&hour(f_getDate))
		else
			tmp_f_datestyle= replace(tmp_f_datestyle,"HH",hour(f_getDate))
		end if
	end if
	if instr(f_datestyle,"MI")>0 then
		if minute(f_getDate)<10 then
			tmp_f_datestyle= replace(tmp_f_datestyle,"MI","0"&minute(f_getDate))
		else
			tmp_f_datestyle= replace(tmp_f_datestyle,"MI",minute(f_getDate))
		end if
	end if
	if instr(f_datestyle,"SS")>0 then
		if second(f_getDate)<10 then
			tmp_f_datestyle= replace(tmp_f_datestyle,"SS","0"&second(f_getDate))
		else
			tmp_f_datestyle= replace(tmp_f_datestyle,"SS",second(f_getDate))
		end if
	end if
	Get_Date = tmp_f_datestyle
End Function
	'html转换函数
Function Encode(str)
	str=Replace(str,"&","&amp;")
	str=Replace(str,"'","''")
	str=Replace(str,"""","&quot;")
	str=Replace(str," ","&nbsp;")
	str=Replace(str,"<","&lt;")
	str=Replace(str,">","&gt;")
	str=Replace(str,"\n","<br>")
	Encode=str
End Function

''删除相关文件.
Function fso_DeleteFile(PhFileName)
	On Error Resume Next
	if isnull(PhFileName) or PhFileName = "" or instr(lcase(PhFileName),"http://")>0 then fso_DeleteFile=true:exit function
	Dim Fso,isTrue
	isTrue = False
	Set Fso = CreateObject(G_FS_FSO)
	Fso.DeleteFile server.MapPath(PhFileName),True
	Set Fso = Nothing
	if Err then
		isTrue = False
	else
		isTrue = True
	end if
	fso_DeleteFile = isTrue
End Function

''删除相关文件.
Function fso_DeleteFolder(PhFileName)
	On Error Resume Next
	if isnull(PhFileName) or PhFileName = "" or instr(lcase(PhFileName),"http://")>0 then fso_DeleteFile=true:exit function
	Dim Fso,isTrue
	isTrue = False
	Set Fso = CreateObject(G_FS_FSO)
	Fso.DeleteFolder server.MapPath(PhFileName),True
	Set Fso = Nothing
	If Err Then
		isTrue = False
	Else
		isTrue = True
	End If
	fso_DeleteFile = isTrue
End Function

' ============================================
' 传入：新闻内容
' 返回：插入分页标记后的新闻内容
' 作用：自动为新闻分页
' ============================================
Function AutoSplitPages(StrNewsContent,Page_Split_page,AutoPagesNum)
	Dim i,IsCount,OneChar,StrCount,FoundStr,Pages_i_Str,Pages_i_Arr
	AutoPagesNum = Clng(AutoPagesNum)
	Page_Split_page = Cstr(Page_Split_page)
	If StrNewsContent<>"" and AutoPagesNum<>0 and InStr(1,StrNewsContent,Page_Split_page)=0 then
		IsCount=True
		Pages_i_Str=""
		For i= 1 To Len(StrNewsContent)
			OneChar=Mid(StrNewsContent,i,1)
			If OneChar="<" Then
				IsCount=False
			ElseIf OneChar=">" Then
				IsCount=True
			Else
				If IsCount=True Then
					If Abs(Asc(OneChar))>255 Then
						StrCount=StrCount+2
					Else
						StrCount=StrCount+1
					End If
					If StrCount>=AutoPagesNum And i<Len(StrNewsContent) Then
						FoundStr=Left(StrNewsContent,i)
						If AllowSplitPages(FoundStr,"table|a|b>|i>|strong|div|span")=true then
							Pages_i_Str=Pages_i_Str & Trim(CStr(i)) & ","
							StrCount=0
						End If
					End If
				End If
			End If
		Next
		If Len(Pages_i_Str)>1 Then Pages_i_Str=Left(Pages_i_Str,Len(Pages_i_Str)-1)
		Pages_i_Arr=Split(Pages_i_Str,",")
		For i = UBound(Pages_i_Arr) To LBound(Pages_i_Arr) Step -1
			StrNewsContent=Left(StrNewsContent,Pages_i_Arr(i)) & Page_Split_page & Mid(StrNewsContent,Pages_i_Arr(i)+1)
		Next
	End If
	AutoSplitPages=StrNewsContent
End Function

' ============================================
' 传入：被查找字符串，查找字符串
' 返回：True,False
' 作用：判断是否允许字符串加入分页标记
' ============================================

Function AllowSplitPages(TempStr,FindStr)
	Dim Inti,BeginStr,EndStr,BeginStrNum,EndStrNum,ArrStrFind,i
	TempStr=LCase(TempStr)
	FindStr=LCase(FindStr)
	If TempStr<>"" and FindStr<>"" then
		ArrStrFind=split(FindStr,"|")
		For i = 0 to Ubound(ArrStrFind)
			BeginStr="<"&ArrStrFind(i)
			EndStr  ="</"&ArrStrFind(i)
			Inti=0
			do while instr(Inti+1,TempStr,BeginStr)<>0
				Inti=instr(Inti+1,TempStr,BeginStr)
				BeginStrNum=BeginStrNum+1
			Loop
			Inti=0
			do while instr(Inti+1,TempStr,EndStr)<>0
				Inti=instr(Inti+1,TempStr,EndStr)
				EndStrNum=EndStrNum+1
			Loop
			If EndStrNum=BeginStrNum then
				AllowSplitPages=true
			Else
				AllowSplitPages=False
				Exit Function
			End If
		Next
	Else
		AllowSplitPages=False
	End If
End Function

Function Recv(Str_Number)
	Dim Arr_Number,Str_Return,Temp_i
	Arr_Number = Split(Str_Number,chr(108))
	Str_Return = ""
	For Temp_i = LBound(Arr_Number) To UBound(Arr_Number)
		Str_Return = Str_Return & Chr(Arr_Number(Temp_i)+31)
	Next
	Recv = Str_Return
End Function
Function IStrLen(TempStr)
	Dim iLen,i,StrAsc
	iLen=0
	for i=1 to len(TempStr)
			StrAsc=Abs(Asc(Mid(TempStr,i,1)))
			if StrAsc>255 then
				iLen=iLen+2
			else
				iLen=iLen+1
			end if
	next
	IStrLen=iLen
End Function
Function GetInfo(GetPath)
	Dim http,ErrContentLength,Report,ContentLength,ErrContent
	ErrContent = ""
	On Error Resume Next
	Response.Clear
	Set http=Server.CreateObject(G_FS_HTTP)
	If Err Then
		Err.Clear
		Set http = Server.CreateObject(G_FS_XMLHTTP)
		If Err Then
			ErrContent = "服务器不支持XML对象"
			Err.Clear
		End If
	End If
	If ErrContent<>"" Then
		GetInfo = "False||"&ErrContent
	Else
		http.setTimeouts 1000,1000,1000,1000
		http.Open "GET",GetPath,False
		http.Send
		If http.readyState<>4 Or http.status<>200 Then
			GetInfo = "False||可能是防火墙阻止或服务器不可用。"
		Else
			GetInfo = "True||"&http.ResponseText
		End If
	End If
End Function

Function GetIsOpenWater()'判断是否开启水印组件
	Dim IsOpenRs,IsOpenSql
	IsOpenSql="Select PicClassid From FS_MF_Config"
	Set IsOpenRs=Conn.execute(IsOpenSql)
	If Not IsOpenRs.Eof Then
		If Cint(IsOpenRs("PicClassid"))=9 Then
			IsOpenRs.Close
			Set IsOpenRs=Nothing
			GetIsOpenWater=True
			Exit Function
		Else
			IsOpenRs.Close
			Set IsOpenRs=Nothing
			GetIsOpenWater=False
		End If
		Exit Function
	Else
		IsOpenRs.Close
		Set IsOpenRs=Nothing
		GetIsOpenWater=False
	Exit Function
	End If
End Function


Function GetGuestBookTitle()'获得留言板参数设置中的留言板Title值
	Dim GuestBookRs,GuestBookSql,TempTitle
	GuestBookSql="Select Title From FS_WS_Config"
	Set GuestBookRs=Conn.execute(GuestBookSql)
	If Not GuestBookRs.Eof Then
		TempTitle=GuestBookRs("Title")
	Else
		TempTitle="留言系统"
	End If
	GuestBookRs.Close
	Set GuestBookRs=Nothing
	GetGuestBookTitle=TempTitle
End Function


Function GetUserSystemTitle()'获得会员系统参数设置中的Title值
	Dim UserSystemRs,UserSystemSql,UserSystemTitle
	UserSystemSql="Select UserSystemName From FS_ME_SysPara"
	Set UserSystemRs=User_Conn.execute(UserSystemSql)
	If Not UserSystemRs.Eof Then
		UserSystemTitle=UserSystemRs("UserSystemName")
	Else
		UserSystemTitle="会员系统"
	End If
	UserSystemRs.Close
	Set UserSystemRs=Nothing
	GetUserSystemTitle=UserSystemTitle
End Function

Function CheckBlogOpen()'检查是否开启日志
	Dim CheckRs,CheckSql
	CheckSql="Select isOpen From FS_ME_iLogSysParam"
	Set CheckRs=User_Conn.Execute(CheckSql)
	If Not CheckRs.Eof Then
		If CheckRs("isOpen")=1 Then
			CheckBlogOpen=True
		Else
			CheckBlogOpen=False
		End If
	Else
		CheckBlogOpen=False
	End If
	CheckRs.Close
	Set CheckRs=Nothing
End Function

'===============================================================
'函数名称：随机函数
'参数：
'-------Num-----产生几位随机数
'返回：一串数字
'函数添加者：Samjun

Function GetRand(Num)
	On Error Resume Next
	Dim R
    Randomize
    For R=1 To Num
        Getrand=Getrand & Int(10*Rnd)
    Next
End Function

'===============================================================
'函数名称：获取IP地址
Function GetIP()
    If Request.ServerVariables("Http_X_Forwarded_For") = "" Then
	GetIP = Request.ServerVariables("Remote_Addr")
    Else
	GetIP = Request.ServerVariables("Http_X_Forwarded_For")
    End If
    GetIP = Replace(GetIP, "'", "")
End Function

'================================================================
'函数名称：将字符串格式的字符转换为HTML格式输入
'参数：
'-------str-----字符串变量
'返回：格式化后的字符串
function Invert(str)
	On Error Resume Next
	str=replace(str,"&lt;","<")
	str=replace(str,"&gt;",">")
	'str=replace(str,"<br>",chr(13))
	'str=replace(str,"&nbsp;"," ")
	str=replace(str,"&quot;","""")
	str=replace(str,"&#39;","'")
	Invert=str
end function
'-------
''长新闻自动分页
'Function AutoSplitPages(StrNewsContent,Page_Split_page,AutoPagesNum)
'	Dim Inti,StrTrueContent,iPageLen,DLocation,XLocation,FoundStr
'	 If StrNewsContent <> "" and AutoPagesNum <> 0 and instr(1,StrNewsContent,Page_Split_page)=0 then
'	  Inti=instr(1,StrNewsContent,"<")
'	  If inti>=1 then '新闻中存在Html标记
'	   StrTrueContent=left(StrNewsContent,Inti-1)
'	   iPageLen=IStrLen(StrTrueContent)
'	   inti=inti+1
'	  Else   '新闻中不存在Html标记，对内容直接分页即可
'	   dim i,c,t
'	   do while i< len(StrNewsContent)
'	   i=i+1
'		c=Abs(Asc(Mid(StrNewsContent,i,1)))
'		if c>255 then '判断为汉字则为两个字符，英文为一个字符
'		 t=t+2
'		else
'		 t=t+1
'		end if
'		Response.Write AutoPagesNum
'		if t>=AutoPagesNum then  '如果字数达到了分页的数量则插入分页符号
'		 StrNewsContent=left(StrNewsContent,i)&Page_Split_page&mid(StrNewsContent,i+1)
'		 i=i+6
'		 t=0
'		end if
'	   loop
'	   AutoSplitPages=StrNewsContent '返回插入分页符号的内容
'	   Response.End
'	   Exit Function
'	  End If
'	  iPageLen=0
'
'	''新闻中存在Html标记时，则用下面的语句来处理
'
'	do while instr(Inti,StrNewsContent,">")<>0
'	   DLocation=instr(Inti,StrNewsContent,">")  '只计算Html标记之外的字符数量
'	   XLocation=instr(DLocation,StrNewsContent,"<")
'	   If XLocation>DLocation+1 then
'		Inti=XLocation
'		StrTrueContent=mid(StrNewsContent,DLocation+1,XLocation-DLocation-1)
'		iPageLen=iPageLen+IStrLen(StrTrueContent) '统计Html之外的字符的数量
'		If iPageLen>AutoPagesNum then    '如果达到了分页的数量则插入分页字符
'		 FoundStr=Lcase(left(StrNewsContent,XLocation-1))
'		 If AllowSplitPages(FoundStr,"table|a|b>|i>|strong|div")=true then
'		  StrNewsContent=left(StrNewsContent,XLocation-1)&Page_Split_page&mid(StrNewsContent,XLocation)
'		  iPageLen=0        '重新统计Html之外的字符
'		 End If
'		End If
'	   ElseIf XLocation=0 then       '在后面再也找不到<，即后面没有Html标记了
'		Exit Do
'	   ElseIf XLocation=DLocation+1 then    '找到的Html标记之间的内容为空，则继续向后找
'		Inti=XLocation
'	   End If
'	  loop
'	 End If
'	AutoSplitPages=StrNewsContent
'End Function
'Function AllowSplitPages(TempStr,FindStr)
'	Dim Inti,BeginStr,EndStr,BeginStrNum,EndStrNum,ArrStrFind,i
'	If TempStr<>"" and FindStr<>"" then
'		ArrStrFind=split(FindStr,"|")
'		For i = 0 to Ubound(ArrStrFind)
'			BeginStr="<"&ArrStrFind(i)
'			EndStr  ="</"&ArrStrFind(i)
'			Inti=0
'			do while instr(Inti+1,TempStr,BeginStr)<>0
'				Inti=instr(Inti+1,TempStr,BeginStr)
'				BeginStrNum=BeginStrNum+1
'			Loop
'			Inti=0
'			do while instr(Inti+1,TempStr,EndStr)<>0
'				Inti=instr(Inti+1,TempStr,EndStr)
'				EndStrNum=EndStrNum+1
'			Loop
'			If EndStrNum=BeginStrNum then
'				AllowSplitPages=true
'			Else
'				AllowSplitPages=False
'				Exit Function
'			End If
'		Next
'	Else
'		AllowSplitPages=False
'	End If
'End Function
'获取汉字的拼音，Str参数为中文字符串，LetterCount参数为每个汉字的拼音截取数
Function GetLetterByChinese(Str,LetterCount)
	Set d = CreateObject(G_FS_DICT)
	d.add "a",-20319
	d.add "ai",-20317
	d.add "an",-20304
	d.add "ang",-20295
	d.add "ao",-20292
	d.add "ba",-20283
	d.add "bai",-20265
	d.add "ban",-20257
	d.add "bang",-20242
	d.add "bao",-20230
	d.add "bei",-20051
	d.add "ben",-20036
	d.add "beng",-20032
	d.add "bi",-20026
	d.add "bian",-20002
	d.add "biao",-19990
	d.add "bie",-19986
	d.add "bin",-19982
	d.add "bing",-19976
	d.add "bo",-19805
	d.add "bu",-19784
	d.add "ca",-19775
	d.add "cai",-19774
	d.add "can",-19763
	d.add "cang",-19756
	d.add "cao",-19751
	d.add "ce",-19746
	d.add "ceng",-19741
	d.add "cha",-19739
	d.add "chai",-19728
	d.add "chan",-19725
	d.add "chang",-19715
	d.add "chao",-19540
	d.add "che",-19531
	d.add "chen",-19525
	d.add "cheng",-19515
	d.add "chi",-19500
	d.add "chong",-19484
	d.add "chou",-19479
	d.add "chu",-19467
	d.add "chuai",-19289
	d.add "chuan",-19288
	d.add "chuang",-19281
	d.add "chui",-19275
	d.add "chun",-19270
	d.add "chuo",-19263
	d.add "ci",-19261
	d.add "cong",-19249
	d.add "cou",-19243
	d.add "cu",-19242
	d.add "cuan",-19238
	d.add "cui",-19235
	d.add "cun",-19227
	d.add "cuo",-19224
	d.add "da",-19218
	d.add "dai",-19212
	d.add "dan",-19038
	d.add "dang",-19023
	d.add "dao",-19018
	d.add "de",-19006
	d.add "deng",-19003
	d.add "di",-18996
	d.add "dian",-18977
	d.add "diao",-18961
	d.add "die",-18952
	d.add "ding",-18783
	d.add "diu",-18774
	d.add "dong",-18773
	d.add "dou",-18763
	d.add "du",-18756
	d.add "duan",-18741
	d.add "dui",-18735
	d.add "dun",-18731
	d.add "duo",-18722
	d.add "e",-18710
	d.add "en",-18697
	d.add "er",-18696
	d.add "fa",-18526
	d.add "fan",-18518
	d.add "fang",-18501
	d.add "fei",-18490
	d.add "fen",-18478
	d.add "feng",-18463
	d.add "fo",-18448
	d.add "fou",-18447
	d.add "fu",-18446
	d.add "ga",-18239
	d.add "gai",-18237
	d.add "gan",-18231
	d.add "gang",-18220
	d.add "gao",-18211
	d.add "ge",-18201
	d.add "gei",-18184
	d.add "gen",-18183
	d.add "geng",-18181
	d.add "gong",-18012
	d.add "gou",-17997
	d.add "gu",-17988
	d.add "gua",-17970
	d.add "guai",-17964
	d.add "guan",-17961
	d.add "guang",-17950
	d.add "gui",-17947
	d.add "gun",-17931
	d.add "guo",-17928
	d.add "ha",-17922
	d.add "hai",-17759
	d.add "han",-17752
	d.add "hang",-17733
	d.add "hao",-17730
	d.add "he",-17721
	d.add "hei",-17703
	d.add "hen",-17701
	d.add "heng",-17697
	d.add "hong",-17692
	d.add "hou",-17683
	d.add "hu",-17676
	d.add "hua",-17496
	d.add "huai",-17487
	d.add "huan",-17482
	d.add "huang",-17468
	d.add "hui",-17454
	d.add "hun",-17433
	d.add "huo",-17427
	d.add "ji",-17417
	d.add "jia",-17202
	d.add "jian",-17185
	d.add "jiang",-16983
	d.add "jiao",-16970
	d.add "jie",-16942
	d.add "jin",-16915
	d.add "jing",-16733
	d.add "jiong",-16708
	d.add "jiu",-16706
	d.add "ju",-16689
	d.add "juan",-16664
	d.add "jue",-16657
	d.add "jun",-16647
	d.add "ka",-16474
	d.add "kai",-16470
	d.add "kan",-16465
	d.add "kang",-16459
	d.add "kao",-16452
	d.add "ke",-16448
	d.add "ken",-16433
	d.add "keng",-16429
	d.add "kong",-16427
	d.add "kou",-16423
	d.add "ku",-16419
	d.add "kua",-16412
	d.add "kuai",-16407
	d.add "kuan",-16403
	d.add "kuang",-16401
	d.add "kui",-16393
	d.add "kun",-16220
	d.add "kuo",-16216
	d.add "la",-16212
	d.add "lai",-16205
	d.add "lan",-16202
	d.add "lang",-16187
	d.add "lao",-16180
	d.add "le",-16171
	d.add "lei",-16169
	d.add "leng",-16158
	d.add "li",-16155
	d.add "lia",-15959
	d.add "lian",-15958
	d.add "liang",-15944
	d.add "liao",-15933
	d.add "lie",-15920
	d.add "lin",-15915
	d.add "ling",-15903
	d.add "liu",-15889
	d.add "long",-15878
	d.add "lou",-15707
	d.add "lu",-15701
	d.add "lv",-15681
	d.add "luan",-15667
	d.add "lue",-15661
	d.add "lun",-15659
	d.add "luo",-15652
	d.add "ma",-15640
	d.add "mai",-15631
	d.add "man",-15625
	d.add "mang",-15454
	d.add "mao",-15448
	d.add "me",-15436
	d.add "mei",-15435
	d.add "men",-15419
	d.add "meng",-15416
	d.add "mi",-15408
	d.add "mian",-15394
	d.add "miao",-15385
	d.add "mie",-15377
	d.add "min",-15375
	d.add "ming",-15369
	d.add "miu",-15363
	d.add "mo",-15362
	d.add "mou",-15183
	d.add "mu",-15180
	d.add "na",-15165
	d.add "nai",-15158
	d.add "nan",-15153
	d.add "nang",-15150
	d.add "nao",-15149
	d.add "ne",-15144
	d.add "nei",-15143
	d.add "nen",-15141
	d.add "neng",-15140
	d.add "ni",-15139
	d.add "nian",-15128
	d.add "niang",-15121
	d.add "niao",-15119
	d.add "nie",-15117
	d.add "nin",-15110
	d.add "ning",-15109
	d.add "niu",-14941
	d.add "nong",-14937
	d.add "nu",-14933
	d.add "nv",-14930
	d.add "nuan",-14929
	d.add "nue",-14928
	d.add "nuo",-14926
	d.add "o",-14922
	d.add "ou",-14921
	d.add "pa",-14914
	d.add "pai",-14908
	d.add "pan",-14902
	d.add "pang",-14894
	d.add "pao",-14889
	d.add "pei",-14882
	d.add "pen",-14873
	d.add "peng",-14871
	d.add "pi",-14857
	d.add "pian",-14678
	d.add "piao",-14674
	d.add "pie",-14670
	d.add "pin",-14668
	d.add "ping",-14663
	d.add "po",-14654
	d.add "pu",-14645
	d.add "qi",-14630
	d.add "qia",-14594
	d.add "qian",-14429
	d.add "qiang",-14407
	d.add "qiao",-14399
	d.add "qie",-14384
	d.add "qin",-14379
	d.add "qing",-14368
	d.add "qiong",-14355
	d.add "qiu",-14353
	d.add "qu",-14345
	d.add "quan",-14170
	d.add "que",-14159
	d.add "qun",-14151
	d.add "ran",-14149
	d.add "rang",-14145
	d.add "rao",-14140
	d.add "re",-14137
	d.add "ren",-14135
	d.add "reng",-14125
	d.add "ri",-14123
	d.add "rong",-14122
	d.add "rou",-14112
	d.add "ru",-14109
	d.add "ruan",-14099
	d.add "rui",-14097
	d.add "run",-14094
	d.add "ruo",-14092
	d.add "sa",-14090
	d.add "sai",-14087
	d.add "san",-14083
	d.add "sang",-13917
	d.add "sao",-13914
	d.add "se",-13910
	d.add "sen",-13907
	d.add "seng",-13906
	d.add "sha",-13905
	d.add "shai",-13896
	d.add "shan",-13894
	d.add "shang",-13878
	d.add "shao",-13870
	d.add "she",-13859
	d.add "shen",-13847
	d.add "sheng",-13831
	d.add "shi",-13658
	d.add "shou",-13611
	d.add "shu",-13601
	d.add "shua",-13406
	d.add "shuai",-13404
	d.add "shuan",-13400
	d.add "shuang",-13398
	d.add "shui",-13395
	d.add "shun",-13391
	d.add "shuo",-13387
	d.add "si",-13383
	d.add "song",-13367
	d.add "sou",-13359
	d.add "su",-13356
	d.add "suan",-13343
	d.add "sui",-13340
	d.add "sun",-13329
	d.add "suo",-13326
	d.add "ta",-13318
	d.add "tai",-13147
	d.add "tan",-13138
	d.add "tang",-13120
	d.add "tao",-13107
	d.add "te",-13096
	d.add "teng",-13095
	d.add "ti",-13091
	d.add "tian",-13076
	d.add "tiao",-13068
	d.add "tie",-13063
	d.add "ting",-13060
	d.add "tong",-12888
	d.add "tou",-12875
	d.add "tu",-12871
	d.add "tuan",-12860
	d.add "tui",-12858
	d.add "tun",-12852
	d.add "tuo",-12849
	d.add "wa",-12838
	d.add "wai",-12831
	d.add "wan",-12829
	d.add "wang",-12812
	d.add "wei",-12802
	d.add "wen",-12607
	d.add "weng",-12597
	d.add "wo",-12594
	d.add "wu",-12585
	d.add "xi",-12556
	d.add "xia",-12359
	d.add "xian",-12346
	d.add "xiang",-12320
	d.add "xiao",-12300
	d.add "xie",-12120
	d.add "xin",-12099
	d.add "xing",-12089
	d.add "xiong",-12074
	d.add "xiu",-12067
	d.add "xu",-12058
	d.add "xuan",-12039
	d.add "xue",-11867
	d.add "xun",-11861
	d.add "ya",-11847
	d.add "yan",-11831
	d.add "yang",-11798
	d.add "yao",-11781
	d.add "ye",-11604
	d.add "yi",-11589
	d.add "yin",-11536
	d.add "ying",-11358
	d.add "yo",-11340
	d.add "yong",-11339
	d.add "you",-11324
	d.add "yu",-11303
	d.add "yuan",-11097
	d.add "yue",-11077
	d.add "yun",-11067
	d.add "za",-11055
	d.add "zai",-11052
	d.add "zan",-11045
	d.add "zang",-11041
	d.add "zao",-11038
	d.add "ze",-11024
	d.add "zei",-11020
	d.add "zen",-11019
	d.add "zeng",-11018
	d.add "zha",-11014
	d.add "zhai",-10838
	d.add "zhan",-10832
	d.add "zhang",-10815
	d.add "zhao",-10800
	d.add "zhe",-10790
	d.add "zhen",-10780
	d.add "zheng",-10764
	d.add "zhi",-10587
	d.add "zhong",-10544
	d.add "zhou",-10533
	d.add "zhu",-10519
	d.add "zhua",-10331
	d.add "zhuai",-10329
	d.add "zhuan",-10328
	d.add "zhuang",-10322
	d.add "zhui",-10315
	d.add "zhun",-10309
	d.add "zhuo",-10307
	d.add "zi",-10296
	d.add "zong",-10281
	d.add "zou",-10274
	d.add "zu",-10270
	d.add "zuan",-10262
	d.add "zui",-10260
	d.add "zun",-10256
	d.add "zuo",-10254
	Dim i,j,ASC_Code,DItems,DKeys,IsCut
	if LetterCount = "" then
		IsCut = False
	else
		if Not IsNumeric(LetterCount) then
			IsCut = False
		else
			LetterCount = CInt(LetterCount)
			IsCut = True
		end if
	end if
	For i = 1 To Len(Str)
		ASC_Code = ASC(MID(Str,i,1))
		If ASC_Code > 0 And ASC_Code < 160 Then
			GetLetterByChinese = GetLetterByChinese & CHR(ASC_Code)
		Else
			If ASC_Code >= -20319 AND ASC_Code <= -10247 Then
				DItems = d.Items
				DKeys = d.Keys
				For j = d.Count - 1 To 0 Step -1
					If DItems(j) <= ASC_Code Then Exit For
				Next
				if IsCut then
					GetLetterByChinese = GetLetterByChinese & UCase(Left(DKeys(j),LetterCount))
				else
					GetLetterByChinese = GetLetterByChinese & UCase(DKeys(j))
				end if
			End If
		End If
	Next
End Function

Function HandleEditorContent(f_Content)
	'f_Content = Replace(f_Content & "",Chr(13) & Chr(10),"")
	if f_Content<>"" then
		f_Content = Server.HTMLEncode(f_Content)
	end if
	HandleEditorContent = f_Content
End Function
%>
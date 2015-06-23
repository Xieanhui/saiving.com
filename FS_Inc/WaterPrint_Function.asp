<%
'为文件添加水印
Function AddWaterMark(FileName)
	Dim strMarkSettingSql,MarkSettingRs,objFileSystem,strFileExtName,objImage
	If InStr(FileName,":") = 0 Then												'把文件名转换为实际路径
		FileName = Server.Mappath(FileName)
	End if
	If FileName <> "" and not IsNull(FileName) Then								'文件名是否不为空,否则退出
		strFileExtName = ""
		If InStr(FileName,".") <> 0 Then
			strFileExtName = Lcase(Trim(Mid(FileName,InStrRev(FileName,".")+1)))
		End if
		If strFileExtName <> "jpg" and strFileExtName <> "gif" and strFileExtName <> "bmp" and strFileExtName <> "png" Then'文件不是可用图片则退出
			Exit Function
		End if
		Set objFileSystem = Server.CreateObject(G_FS_FSO)
		If objFileSystem.FileExists(FileName) Then				'文件存在,否则退出
			strMarkSettingSql = "select top 1 * from FS_MF_config"
			Set MarkSettingRs = conn.Execute(strMarkSettingSql)
			If MarkSettingRs("PicClassid") <> "9" Then						'选择了某个水印组件,否则退出
				Select Case MarkSettingRs("PicClassid")
					Case "0"													'使用AspJpeg组件												
						If IsObjInstalled("Persits.Jpeg") Then					'AspJpeg组件已安装,否则退出
							If IsExpired("Persits.Jpeg") Then
								Response.Write("Persits.Jpeg组件已过期，请选择其他组件或关闭水印功能。")
								Response.End
							End if
							If MarkSettingRs("MarkType") = "1" Then				'添加文字水印
								AddTextMark 1,MarkSettingRs("MarkText"),MarkSettingRs("MarkFontColor"),MarkSettingRs("MarkFontName"),MarkSettingRs("MarkFontBond"),MarkSettingRs("MarkFontSize"),MarkSettingRs("MarkPosition"),FileName
							Else												'添加图片水印
								AddPictureMark 1,MarkSettingRs("MarkWidth"),MarkSettingRs("MarkHeight"),MarkSettingRs("MarkPicture"),MarkSettingRs("MarkOpacity"),MarkSettingRs("MarkTranspColor"),MarkSettingRs("MarkPosition"),FileName
							End if
						End if
					Case "1"													'使用wsImage组件
						If strFileExtName = "png" Then							'wsImage组件不支持PNG文件,是则退出
							Exit Function
						End if
						If IsObjInstalled("wsImage.Resize") Then				'wsImage组件已安装,否则退出
							If IsExpired("wsImage.Resize") Then
								Response.Write("wsImage.Resize组件已过期，请选择其他组件或关闭水印功能。")
								Response.End
							End if
							If MarkSettingRs("MarkType") = "1" Then				'添加文字水印
								AddTextMark 2,MarkSettingRs("MarkText"),MarkSettingRs("MarkFontColor"),MarkSettingRs("MarkFontName"),MarkSettingRs("MarkFontBond"),MarkSettingRs("MarkFontSize"),MarkSettingRs("MarkPosition"),FileName
							Else												'添加图片水印
								AddPictureMark 2,MarkSettingRs("MarkWidth"),MarkSettingRs("MarkHeight"),MarkSettingRs("MarkPicture"),MarkSettingRs("MarkOpacity"),MarkSettingRs("MarkTranspColor"),MarkSettingRs("MarkPosition"),FileName
							End if
						End if
					Case "2"													'使用SA-ImgWriter组件
						If IsObjInstalled("SoftArtisans.ImageGen") Then			'SA-ImgWriter组件已安装,否则退出
							If IsExpired("SoftArtisans.ImageGen") Then
								Response.Write("SoftArtisans.ImageGen组件已过期，请选择其他组件或关闭水印功能。")
								Response.End
							End if
							If MarkSettingRs("MarkType") = "1" Then				'添加文字水印
								AddTextMark 3,MarkSettingRs("MarkText"),MarkSettingRs("MarkFontColor"),MarkSettingRs("MarkFontName"),MarkSettingRs("MarkFontBond"),MarkSettingRs("MarkFontSize"),MarkSettingRs("MarkPosition"),FileName
							Else												'添加图片水印
								AddPictureMark 3,MarkSettingRs("MarkWidth"),MarkSettingRs("MarkHeight"),MarkSettingRs("MarkPicture"),MarkSettingRs("MarkOpacity"),MarkSettingRs("MarkTranspColor"),MarkSettingRs("MarkPosition"),FileName
							End if
						End if
				End Select
			End if
			Set MarkSettingRs = nothing
		End if
		Set objFileSystem = nothing
	End if
End Function
'为图片添加文字水印
Function AddTextMark(MarkComponentID,MarkText,MarkFontColor,MarkFontName,MarkFontBond,MarkFontSize,MarkPosition,FileName)
	Dim objImage,X,Y,Text,TextWidth,FontColor,FontName,FondBond,FontSize,OriginalWidth,OriginalHeight
	If InStr(FileName,":") = 0 Then																'把文件名转换为实际路径
		FileName = Server.Mappath(FileName)
	End if
	Text = Trim(MarkText)
	If Text = "" Then
		Exit Function
	End if
	'FontColor = Replace(MarkFontColor,"#","&H")
	FontColor="&H"&MarkFontColor
	FontName = MarkFontName
	If MarkFontBond = "1" Then
		FondBond = True
	Else
		FondBond = False
	End if
	FontSize = Cint(MarkFontSize)
	Select Case MarkComponentID
		Case "1"
			If Not IsObjInstalled("Persits.Jpeg") Then
				Exit Function
			End if
			Set objImage = Server.CreateObject(G_PERSITS_JPEG)
			objImage.Open FileName
			objImage.Canvas.Font.Color =FontColor
			objImage.Canvas.Font.Family = FontName
			objImage.Canvas.Font.Bold = FondBond
			objImage.Canvas.Font.Size = FontSize
			TextWidth = objImage.Canvas.GetTextExtent(Text)										'计算GB2313编码的字符串所占宽度
			
			If objImage.OriginalWidth < TextWidth Or objImage.OriginalHeight < FontSize Then	'如果图片高度小于字体大小或宽度小于字符串宽度则退出
				Exit Function
			End if
			GetPostion Cint(MarkPosition),X,Y,objImage.OriginalWidth,objImage.OriginalHeight,TextWidth,FontSize '计算坐标
			objImage.Canvas.Print X, Y, Text
			objImage.Save FileName
		Case "2"
			If Not IsObjInstalled("wsImage.Resize") Then
				Exit Function
			End if
			Set objImage = Server.CreateObject(G_WSIMAGE_RESIZE)
			objImage.LoadSoucePic Cstr(FileName)
			objImage.TxtMarkFont = CStr(FontName)
			objImage.TxtMarkBond = FondBond
			objImage.TxtMarkHeight = FontSize
			'objImage.GetSourceInfo OriginalWidth,OriginalHeight
			'GetPostion Cint(MarkPosition),X,Y,OriginalWidth,OriginalHeight,Len(Text)*FontSize*3/4,FontSize '计算坐标
			FontColor = "&H"&Mid(FontColor,7)&Mid(FontColor,5,2)&Mid(FontColor,3,2)				'颜色代码转换&HBBGGRR
			objImage.AddTxtMark Cstr(FileName),CStr(Text),Clng(FontColor),1,1
		Case "3"
			If Not IsObjInstalled("SoftArtisans.ImageGen") Then
				Exit Function
			End if
			Set objImage = Server.CreateObject(G_SOFTARTISANS_IMAGEGEN)
			objImage.LoadImage FileName
			objImage.Font.height = FontSize
			objImage.Font.name	= FontName
			FontColor = "&H"&Mid(FontColor,7)&Mid(FontColor,5,2)&Mid(FontColor,3,2)				'颜色代码转换&HBBGGRR
			objImage.Font.Color	= Clng(FontColor)
			objImage.Text = Text
			GetPostion Cint(MarkSettingRs("MarkPosition")),X,Y,objImage.Width,objImage.Height,objImage.TextWidth,objImage.TextHeight '计算坐标
			objImage.DrawTextOnImage X, Y,objImage.TextWidth,objImage.TextHeight
			objImage.SaveImage 0, objImage.ImageFormat, FileName 
	End Select
	Set objImage = nothing
End Function
'为图片添加图片水印
Function AddPictureMark(MarkComponentID,MarkWidth,MarkHeight,MarkPicture,MarkOpacity,MarkTranspColor,MarkPosition,FileName)
	Dim objImage,objMark,X,Y,OriginalWidth,OriginalHeight,Position
	If InStr(FileName,":") = 0 Then																'把文件名转换为实际路径
		FileName = Server.Mappath(FileName)
	End if
	If IsNull(MarkWidth) Or MarkWidth = "" Then
		MarkWidth = 40
	Else
		MarkWidth = Cint(MarkWidth)
	End if
	If IsNull(MarkHeight) Or MarkHeight = "" Then
		MarkHeight = 20
	Else
		MarkHeight = Cint(MarkHeight)
	End if
	If MarkPicture = "" Then
		Exit Function
	End if
	If IsNull(MarkOpacity) Or MarkOpacity = "" Then
		MarkOpacity = 1
	Else
		MarkOpacity = Csng(MarkOpacity)
	End if
	If MarkTranspColor <> "" Then																'转换颜色代码
		MarkTranspColor = "&H"&MarkTranspColor
	End if
	Select Case MarkComponentID
		Case 1
			If Not IsObjInstalled("Persits.Jpeg") Then
				Exit Function
			End if
			Set objImage = Server.CreateObject(G_PERSITS_JPEG)
			Set objMark = Server.CreateObject(G_PERSITS_JPEG)
			objImage.Open FileName
			If objImage.OriginalWidth < MarkWidth Or objImage.OriginalHeight < MarkHeight Then	'如果图片高度小于水印高度或宽度小于字水印宽度则退出
				Exit Function
			End if
			objMark.Open Server.Mappath(MarkPicture)
			GetPostion Cint(MarkPosition),X,Y,objImage.OriginalWidth,objImage.OriginalHeight,MarkWidth,MarkHeight '计算坐标
			If MarkTranspColor <> "" Then
				objImage.DrawImage X,Y,objMark,MarkOpacity,MarkTranspColor
			else
				objImage.DrawImage X,Y,objMark,MarkOpacity
			End if
			objImage.Save FileName
		Case 2
			If Not IsObjInstalled("wsImage.Resize") Then
				Exit Function
			End if
			Set objImage = Server.CreateObject(G_WSIMAGE_RESIZE)
			objImage.LoadSoucePic Cstr(FileName)
			objImage.LoadImgMarkPic Server.Mappath(MarkPicture)
			objImage.GetSourceInfo OriginalWidth,OriginalHeight
			GetPostion Cint(MarkPosition),X,Y,OriginalWidth,OriginalHeight,MarkWidth,MarkHeight '计算坐标
			If MarkTranspColor = "" Then
				MarkTranspColor = 0
			Else
				MarkTranspColor = "&H"&Mid(MarkTranspColor,7)&Mid(MarkTranspColor,5,2)&Mid(MarkTranspColor,3,2)				'颜色代码转换&HBBGGRR
			End if
			objImage.AddImgMark Cstr(FileName),int(X),int(Y),Clng(MarkTranspColor),Int(CSng(MarkOpacity)*100)
		Case 3
			If Not IsObjInstalled("SoftArtisans.ImageGen") Then
				Exit Function
			End if
			Set objImage = Server.CreateObject(G_SOFTARTISANS_IMAGEGEN)
			objImage.LoadImage FileName
			Select Case Cint(MarkSettingRs("MarkPosition"))
				Case 1
					Position = 3
				Case 2
					Position = 5
				Case 3
					Position = 1
				Case 4
					Position = 6
				Case 5
					Position = 8
			End Select
			If MarkTranspColor <> "" Then
				MarkTranspColor = "&H"&Mid(MarkTranspColor,7)&Mid(MarkTranspColor,5,2)&Mid(MarkTranspColor,3,2)				'颜色代码转换&HBBGGRR
				objImage.AddWatermark Server.MapPath(MarkPicture), Position,CSng(MarkOpacity),Clng(MarkTranspColor)
			else
				objImage.AddWatermark Server.MapPath(MarkPicture), Position,CSng(MarkOpacity)
			End if
			'Position:saiTopMiddle 0 saiCenterMiddle 1 saiBottomMiddle 2 saiTopLeft 3 saiCenterLeft 4 saiBottomLeft 5 saiTopRight 6 saiCenterRight 7 saiBottomRight 8 
			objImage.SaveImage 0, objImage.ImageFormat,FileName 
	End Select
	Set objImage = nothing
	Set objMark = nothing
End Function
'计算水印相对图片的坐标
Function GetPostion(MarkPosition,X,Y,ImageWidth,ImageHeight,MarkWidth,MarkHeight)
	Rem 随机水印5格 左上、左下、居中、右上、右下
	If int(MarkPosition) = 6 Then
		MarkPosition = 5 * Rnd + 1
	End if
	Select Case Cint(MarkPosition)
		Case 1
			X = 1
			Y = 1
		Case 2
			X = 1
			Y = Int(ImageHeight - MarkHeight - 1)
		Case 3
			X = Int((ImageWidth - MarkWidth)/2)
			Y = Int((ImageHeight - MarkHeight)/2)
		Case 4
			X = Int(ImageWidth - MarkWidth - 1)
			Y = 1
		Case 5
			X = Int(ImageWidth - MarkWidth - 1)
			Y = Int(ImageHeight - MarkHeight - 1)
	End Select						
End Function
'由原图片根据数据里保存的设置生成缩略图
Function CreateThumbnailEx(FileName,ThumbnailFileName)
	Dim strSql,RsThumbnailSetting
	strSql = "Select ThumbnailComponent,RateTF,ThumbnailWidth,ThumbnailHeight,ThumbnailRate From FS_MF_Config"
	Set RsThumbnailSetting = Conn.Execute(strSql)
	If RsThumbnailSetting("ThumbnailComponent") <> "9" and (not IsNull(RsThumbnailSetting("ThumbnailComponent")))Then
		If RsThumbnailSetting("RateTF") = "0" Then
			CreateThumbnailEx = CreateThumbnail(FileName,Cint(RsThumbnailSetting("ThumbnailWidth")),Cint(RsThumbnailSetting("ThumbnailHeight")),0,ThumbnailFileName)
		Else
			CreateThumbnailEx = CreateThumbnail(FileName,0,0,Csng(RsThumbnailSetting("ThumbnailRate")),ThumbnailFileName)
		End if
	End if
	Set RsThumbnailSetting = nothing
End Function
'由原图片生成指定宽度和高度的缩略图
Function CreateThumbnail(FileName,Width,Height,Rate,ThumbnailFileName)
	Dim strSql,RsSetting,objImage,iWidth,iHeight,strFileExtName
	CreateThumbnail = False
	If IsNull(FileName) Then									'如果原图片未指定直接退出
		Exit Function
	Elseif FileName="" Then
		Exit Function
	End if
	If InStr(FileName,".") <> 0 Then
		strFileExtName = Lcase(Trim(Mid(FileName,InStrRev(FileName,".")+1)))
	End if
	If strFileExtName <> "jpg" and strFileExtName <> "gif" and strFileExtName <> "bmp" and strFileExtName <> "png" Then'文件不是可用图片则退出
		Exit Function
	End if
	If IsNull(ThumbnailFileName) Then							'如果缩略图未指定保存路径直接退出
		Exit Function
	Elseif ThumbnailFileName="" Then
		Exit Function
	End if
	If IsNull(Width) Then										'如果缩略图宽度未指定则将其指定为0
		Width = 120
	Elseif Width="" Then
		Width = 120
	End if
	If IsNull(Rate) Then										'如果缩略图缩放比例未指定则将其指定为0
		Rate = 0
	Elseif Rate="" Then
		Rate = 0
	End if
	If IsNull(Height) Then										'如果缩略图高度未指定则将其指定为0
		Height = 200
	Elseif Height="" Then
		Height = 200
	End if
	If InStr(FileName,":") = 0 Then								'原图片路径转换化物理路径
		FileName = Server.Mappath(FileName)
	End if
	If InStr(ThumbnailFileName,":") = 0 Then					'缩略图路径转换化物理路径
		ThumbnailFileName = Server.Mappath(ThumbnailFileName)
	End if
	Width = Cint(Width)
	Height = Cint(Height)
	Rate = CSng(Rate)
	
	strSql = "Select ThumbnailComponent From FS_MF_Config"
	Set RsSetting = Conn.Execute(strSql)
	Select Case Cint(RsSetting("ThumbnailComponent"))
		Case 9													'缩略图功能关闭,退出
			Exit Function
		Case 0
			If Not IsObjInstalled("Persits.Jpeg") Then			'Persits.Jpeg未安装,退出
				Exit Function
			End if
			If IsExpired("Persits.Jpeg") Then
				Response.Write("Persits.Jpeg组件已过期，请选择其他组件或关闭缩略图功能。")
				Response.End
			End if
			Set objImage = Server.CreateObject(G_PERSITS_JPEG)
			objImage.Open FileName
			If Rate = 0 and (Width <> 0 Or Height<> 0) Then
				If Width < objImage.OriginalWidth And Height < objImage.OriginalHeight Then
					If Width = 0 and Height <> 0 Then
						objImage.Width = objImage.OriginalWidth/objImage.OriginalHeight*Height
						objImage.Height = Height
					Elseif Width <> 0 and Height = 0 Then
						objImage.Width = Width
						objImage.Height = objImage.OriginalHeight/objImage.OriginalWidth*Width
					ElseIf Width <> 0 and Height <> 0 Then
						objImage.Width = Width
						objImage.Height = Height
					End if
				End if
			Elseif  Rate <> 0 Then
				objImage.Width = objImage.OriginalWidth*Rate
				objImage.Height = objImage.OriginalHeight*Rate
			End if
			
			objImage.Save ThumbnailFileName
		Case 1
			If Not IsObjInstalled("wsImage.Resize") Then			'wsImage.Resize未安装,退出
				Exit Function
			End if
			If IsExpired("wsImage.Resize") Then
				Response.Write("wsImage.Resize组件已过期，请选择其他组件或关闭缩略图功能。")
				Response.End
			End if
			If strFileExtName = "png" Then							'wsImage.Resize不支持PNG图片,是则退出
				Exit Function
			End if
			Set objImage = Server.CreateObject(G_WSIMAGE_RESIZE)
			objImage.LoadSoucePic CStr(FileName)
			If Rate = 0 and (Width <> 0 Or Height<> 0) Then
				objImage.GetSourceInfo iWidth,iHeight
				If Width < iWidth And Height < iHeight Then
					If Width = 0 and Height <> 0 Then
						objImage.OutputSpic CStr(ThumbnailFileName),0,Height,2
					Elseif Width <> 0 and Height = 0 Then
						objImage.OutputSpic CStr(ThumbnailFileName),Width,0,1
					ElseIf Width <> 0 and Height <> 0 Then
						objImage.OutputSpic CStr(ThumbnailFileName),Width,Height,0
					Else
						objImage.OutputSpic CStr(ThumbnailFileName),1,1,3
					End if
				Else
					objImage.OutputSpic CStr(ThumbnailFileName),1,1,3
				End if
			Elseif  Rate <> 0 Then
				objImage.OutputSpic CStr(ThumbnailFileName),Rate,Rate,3
			Else
				objImage.OutputSpic CStr(ThumbnailFileName),1,1,3
			End if
		Case 2
			If Not IsObjInstalled("SoftArtisans.ImageGen") Then		'SoftArtisans.ImageGen未安装,退出
				Exit Function
			End if
			If IsExpired("SoftArtisans.ImageGen") Then
				Response.Write("SoftArtisans.ImageGen组件已过期，请选择其他组件或关闭缩略图功能。")
				Response.End
			End if
			Set objImage = Server.CreateObject(G_SOFTARTISANS_IMAGEGEN)
			objImage.LoadImage FileName
			If Rate = 0 and (Width <> 0 Or Height<> 0) Then
				If Width < objImage.Width And Height < objImage.Height Then
					If Width = 0 and Height <> 0 Then
						objImage.CreateThumbnail  ,Clng(Height),0,true
					Elseif Width <> 0 and Height = 0 Then
						objImage.CreateThumbnail  Clng(Width),objImage.Height/objImage.Width*Width,0,false
					ElseIf Width <> 0 and Height <> 0 Then
						objImage.CreateThumbnail  Clng(Width),Clng(Height),0,false
					End if
				End if
			Elseif  Rate <> 0 Then
				objImage.CreateThumbnail Clng(objImage.Width*Rate),Clng(objImage.Height*Rate),0,false
			End if
			objImage.SaveImage 0,objImage.ImageFormat,ThumbnailFileName
		Case 3
			If Not IsObjInstalled("CreatePreviewImage.cGvbox") Then		'CreatePreviewImage.cGvbox未安装,退出
				Exit Function
			End if
			set objImage = Server.CreateObject(G_CREATEPREVIEW_CGVBOX)
			objImage.SetImageFile = FileName							'imagename原始文件的物理路径
			If Rate = 0 and (Width <> 0 Or Height<> 0) Then
				objImage.SetPreviewImageSize = Width					'预览图宽度
			Elseif  Rate <> 0 Then
				objImage.SetPreviewImageSize = objImage.SetPreviewImageSize*Rate				'预览图宽度
			End if
			objImage.SetSavePreviewImagePath = ThumbnailFileName		'预览图存放路径
			If objImage.DoImageProcess = False Then						'创建预览图的文件
				Exit Function
			End if
	End Select
	CreateThumbnail = True	
End Function
%>






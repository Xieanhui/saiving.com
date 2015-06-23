<%
	Dim MyFile,CrHNJS,TempStateFlag,DnsPath,GetUrl,LDnsPath,AdsJSStr,AdsTempStr,JsFileName,objStream,AdsTempStrRight
	function AdsTempPicStr(Location)
	    dim FunLocation,FunAdsObj
		
		FunLocation = clng(Location)
		Set FunAdsObj = Conn.Execute("select * from FS_AD_Info where AdID="&CintStr(FunLocation)&"")
		if InStr(1,LCase(FunAdsObj("AdPicPath")),".swf",1)<>0 Then
			If InStr(1,LCase(FunAdsObj("AdPicPath")),"http://")<>0 then
				AdsTempStr="<EMBED src="""& FunAdsObj("AdPicPath") &""" quality=high WIDTH="""& FunAdsObj("AdPicWidth") &""" HEIGHT="""& FunAdsObj("AdPicHeight") &""" TYPE=""application/x-shockwave-flash"" PLUGINSPAGE=""http://www.macromedia.com/go/getflashplayer""></EMBED>"
			Else
				AdsTempStr="<EMBED src=""http://"& request.Cookies("FoosunMFCookies")("FoosunMFDomain")& FunAdsObj("AdPicPath") &""" quality=high WIDTH="""& FunAdsObj("AdPicWidth") &""" HEIGHT="""& FunAdsObj("AdPicHeight") &""" TYPE=""application/x-shockwave-flash"" PLUGINSPAGE=""http://www.macromedia.com/go/getflashplayer""></EMBED>"
			End If
		Else
			If InStr(1,LCase(FunAdsObj("AdPicPath")),"http://")<>0 then
				AdsTempStr="<a href=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/AdsClick.asp?Location="& FunLocation &""" target=_blank><img src="""& FunAdsObj("AdPicPath") &""" border=0 width="""& FunAdsObj("AdPicWidth") &""" height="""& FunAdsObj("AdPicHeight") &""" alt="""& FunAdsObj("AdCaptionTxt") &""" align=top></a>"
			Else
				AdsTempStr="<a href=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/AdsClick.asp?Location="& FunLocation &""" target=_blank><img src="""& FunAdsObj("AdPicPath") &""" border=0 width="""& FunAdsObj("AdPicWidth") &""" height="""& FunAdsObj("AdPicHeight") &""" alt="""& FunAdsObj("AdCaptionTxt") &""" align=top></a>"
			End If
		End If
		if InStr(1,LCase(FunAdsObj("AdRightPicPath")),".swf",1)<>0 Then
			If InStr(1,LCase(FunAdsObj("AdRightPicPath")),"http://")<>0 then
				AdsTempStrRight="<EMBED src="""& FunAdsObj("AdRightPicPath") &""" quality=high WIDTH="""& FunAdsObj("AdPicWidth") &""" HEIGHT="""& FunAdsObj("AdPicHeight") &""" TYPE=""application/x-shockwave-flash"" PLUGINSPAGE=""http://www.macromedia.com/go/getflashplayer""></EMBED>"
			Else
				AdsTempStrRight="<EMBED src="""& request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"& FunAdsObj("AdRightPicPath") &""" quality=high WIDTH="""& FunAdsObj("AdPicWidth") &""" HEIGHT="""& FunAdsObj("AdPicHeight") &""" TYPE=""application/x-shockwave-flash"" PLUGINSPAGE=""http://www.macromedia.com/go/getflashplayer""></EMBED>"
			End If
		Else
			If InStr(1,LCase(FunAdsObj("AdRightPicPath")),"http://")<>0 then
				AdsTempStrRight="<a href=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/AdsClick.asp?Location="& FunLocation &""" target=_blank><img src="""& FunAdsObj("AdRightPicPath") &""" border=0 width="""& FunAdsObj("AdPicWidth") &""" height="""& FunAdsObj("AdPicHeight") &""" alt="""& FunAdsObj("AdCaptionTxt") &""" align=top></a>"
			Else
				AdsTempStrRight="<a href=""http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/AdsClick.asp?Location="& FunLocation &""" target=_blank><img src=""" & FunAdsObj("AdRightPicPath") &""" border=0 width="""& FunAdsObj("AdPicWidth") &""" height="""& FunAdsObj("AdPicHeight") &""" alt="""& FunAdsObj("AdCaptionTxt") &""" align=top></a>"
			End If
		End If
		FunAdsObj.close
		set FunAdsObj = nothing
	end function
 
        Sub ShowAds(TempLocation)
		    dim ShowAdsStr,ShowAdsLocation,ShowAdsObj
			ShowAdsLocation = clng(TempLocation)
			AdsTempPicStr(ShowAdsLocation)
			ShowAdsStr = AdsTempStr
			Set ShowAdsObj = Conn.Execute("select * from FS_AD_Info where AdID="&CintStr(ShowAdsLocation)&"")
			if CheckAd(ShowAdsObj)=False  then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				AdsJSStr = "document.write('"& ShowAdsStr &"');" & vbCrLf & _
				           "document.write('<script src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/show.asp?Location="&ShowAdsLocation&"></script>');"
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(Str_SysDir)) = false then
				MyFile.CreateFolder(Server.MapPath(Str_SysDir))
			End If
			 if MyFile.FileExists(Server.MapPath(Str_SysDir)&"/"& ShowAdsLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(Str_SysDir)&"/"& ShowAdsLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(Str_SysDir)&"/"& ShowAdsLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			ShowAdsObj.close
			Set ShowAdsObj = Nothing
		 End Sub
		 
		 Sub NewWindow(TempLocation)
		    dim NewWindowObj,NewWindowLocation,dialogConent,dialogConent1 ,sUrl
			NewWindowLocation = clng(TempLocation)
			Set NewWindowObj = Conn.Execute("select * from FS_AD_Info where AdID="&CintStr(NewWindowLocation)&"")
			If CheckAd(NewWindowObj)=True Then
				If Instr(1,LCase(NewWindowObj("AdPicPath")),"http://") <> 0 then
					AdsJSStr = "window.open('http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/pic.asp?pic="&NewWindowLocation&"','','width="& NewWindowObj("AdPicWidth") &",height="& NewWindowObj("AdPicHeight") &",scrollbars=1');"
				Else
					AdsJSStr = "window.open('http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/pic.asp?pic="&NewWindowLocation&"','','width="& NewWindowObj("AdPicWidth") &",height="& NewWindowObj("AdPicHeight") &",scrollbars=1');" & vbCrLf & _
					"document.write('<script src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/show.asp?Location="&NewWindowLocation&"></script>');"
				End If
			Else
				AdsJSStr="document.write('此广告已经暂停或是失效');"
			End If
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(Str_SysDir)) = false then
				MyFile.CreateFolder(Server.MapPath(Str_SysDir))
			End If
			 if MyFile.FileExists(Server.MapPath(Str_SysDir)&"/"& NewWindowLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(Str_SysDir)&"/"& NewWindowLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(Str_SysDir)&"/"& NewWindowLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=Nothing
			NewWindowObj.close
			Set NewWindowObj = Nothing
		 End Sub
		 
		 Sub OpenWindow(TempLocation)
		    dim OpenWindowObj,OpenWindowLocation

			OpenWindowLocation = clng(TempLocation)
			Set OpenWindowObj = Conn.Execute("select * from FS_AD_Info where AdID="&CintStr(OpenWindowLocation)&"")
			If 	CheckAd(OpenWindowObj)=True Then
				AdsJSStr = "window.open('http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/pic.asp?pic="&OpenWindowLocation&"','_blank');"
			Else
				AdsJSStr="document.write('此广告已经暂停或是失效');"
			End If
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(Str_SysDir)) = false then
				MyFile.CreateFolder(Server.MapPath(Str_SysDir))
			End If
			 if MyFile.FileExists(Server.MapPath(Str_SysDir)&"/"& OpenWindowLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(Str_SysDir)&"/"& OpenWindowLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(Str_SysDir)&"/"& OpenWindowLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			OpenWindowObj.close
			Set OpenWindowObj = Nothing
		 End Sub
		 
		 Sub FilterAway(TempLocation)
		    dim FilterAwayStr,FilterAwayLocation,FilterAwayObj
			FilterAwayLocation = clng(TempLocation)
			AdsTempPicStr(FilterAwayLocation)
			FilterAwayStr = AdsTempStr
			Set FilterAwayObj = Conn.Execute("select * from FS_AD_Info where AdID="&CintStr(FilterAwayLocation)&"")
			if CheckAd(FilterAwayObj)=False then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				AdsJSStr = "FilterAwayStr=(document.layers)?true:false;" & vbCrLf & _
						   "if(FilterAwayStr){document.write('<layer id=FilterAwayT onLoad=""moveToAbsolute(layer1.pageX-160,layer1.pageY);clip.height="& FilterAwayObj("AdPicHeight") &";clip.width="& FilterAwayObj("AdPicWidth") &"; visibility=show;""><layer id=FilterAwayF position:absolute; bottom:20; center:1>"& FilterAwayStr &"</layer></layer>');}" & vbCrLf & _
						   "else{document.write('<div style=""position:absolute;bottom:"& cint(FilterAwayObj("AdPicHeight")+20) &"; center:1;""><div id=FilterAwayT style=""position:absolute; width:"& FilterAwayObj("AdPicWidth") &"; height:"& FilterAwayObj("AdPicHeight") &";clip:rect(0,"& FilterAwayObj("AdPicWidth") &","& FilterAwayObj("AdPicHeight") &",0)""><div id=FilterAwayF style=""position:absolute;bottom:20; center:1"">"& FilterAwayStr &"</div></div></div>');} " & vbCrLf & _
						   "document.write('<script language=javascript src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/CreateJs/FilterAway.js></script>');" & vbCrLf & _
				           "document.write('<script src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/show.asp?Location="&FilterAwayLocation&"></script>');"
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(Str_SysDir)) = false then
				MyFile.CreateFolder(Server.MapPath(Str_SysDir))
			End If
			 if MyFile.FileExists(Server.MapPath(Str_SysDir)&"/"& FilterAwayLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(Str_SysDir)&"/"& FilterAwayLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(Str_SysDir)&"/"& FilterAwayLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			FilterAwayObj.close
			Set FilterAwayObj = Nothing
		 End Sub
		 
		 Sub DialogBox(TempLocation)
		    dim DialogBoxObj,DialogBoxLocation
			DialogBoxLocation = clng(TempLocation)
			Set DialogBoxObj = Conn.Execute("select * from FS_AD_Info where AdID="&CintStr(DialogBoxLocation)&"")
			if CheckAd(DialogBoxObj)=False then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				AdsJSStr = "window.showModalDialog('http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/pic.asp?pic="&DialogBoxLocation&"','','dialogWidth:"& DialogBoxObj("AdPicWidth")+10 &"px;dialogHeight:"& DialogBoxObj("AdPicHeight")+30 &"px;center:0;status:no');"
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(Str_SysDir)) = false then
				MyFile.CreateFolder(Server.MapPath(Str_SysDir))
			End If
			 if MyFile.FileExists(Server.MapPath(Str_SysDir)&"/"& DialogBoxLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(Str_SysDir)&"/"& DialogBoxLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(Str_SysDir)&"/"& DialogBoxLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
		 	DialogBoxObj.Close
			Set DialogBoxObj = Nothing
		 End Sub
		 
		 Sub ClarityBox(TempLocation)
		    dim ClarityBoxObj,ClarityBoxLocation,ClarityBoxStr
			ClarityBoxLocation = clng(TempLocation)
			AdsTempPicStr(ClarityBoxLocation)
			ClarityBoxStr = AdsTempStr
			Set ClarityBoxObj = Conn.Execute("select * from FS_AD_Info where AdID="&CintStr(ClarityBoxLocation)&"")
			if CheckAd(ClarityBoxObj)=False then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				AdsJSStr = "document.write('<script language=javascript src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/CreateJs/ClarityBox.js></script>'); " & vbCrLf & _
						   "document.write('<div style=""position:absolute;left:300px;top:150px;width:"& ClarityBoxObj("AdPicWidth") &"; height:"& ClarityBoxObj("AdPicHeight") &";z-index:1;solid;filter:alpha(opacity=90)"" id=ClarityBoxID onmousedown=""ClarityBox(this)"" onmousemove=""ClarityBoxMove(this)"" onMouseOut=""down=false"" onmouseup=""down=false""><table cellpadding=0 border=0 cellspacing=1 width="& ClarityBoxObj("AdPicWidth") &" height="& cint(ClarityBoxObj("AdPicHeight")+20) &" bgcolor=#000000><tr><td height=20 align=right style=""cursor:move;""><a href=# style=""font-size: 9pt; color: white; text-decoration: none"" onClick=ClarityBoxclose(""ClarityBoxID"") >>>关闭>></a></td></tr><tr><td>"&ClarityBoxStr&"</td></tr></table></div>');" & vbCrLf & _
				           "document.write('<script src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/show.asp?Location="&ClarityBoxLocation&"></script>');"
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(Str_SysDir)) = false then
				MyFile.CreateFolder(Server.MapPath(Str_SysDir))
			End If
			 if MyFile.FileExists(Server.MapPath(Str_SysDir)&"/"& ClarityBoxLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(Str_SysDir)&"/"& ClarityBoxLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(Str_SysDir)&"/"& ClarityBoxLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			ClarityBoxObj.close
			Set ClarityBoxObj = Nothing
		 End Sub
		 
		 Sub RightBottom(TempLocation)
		    dim RightBottomStr,RightBottomLocation,RightBottomObj
			RightBottomLocation = clng(TempLocation)
			AdsTempPicStr(RightBottomLocation)
			RightBottomStr = AdsTempStr
			Set RightBottomObj = Conn.Execute("select * from FS_AD_Info where AdID="&CintStr(RightBottomLocation)&"")
			if CheckAd(RightBottomObj)=False then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				AdsJSStr = "if (navigator.appName == 'Netscape')" & vbCrLf & _
						   "{document.write('<layer id=RightBottom top=150 width="& RightBottomObj("AdPicWidth") &" height="& RightBottomObj("AdPicHeight") &">"& RightBottomStr &"</layer>');}" & vbCrLf & _
						   "else{document.write('<div id=RightBottom style=""position: absolute;width:"& RightBottomObj("AdPicWidth") &";height:"& RightBottomObj("AdPicHeight") &";visibility: visible;z-index: 1"">"& RightBottomStr &"</div>');}" & vbCrLf & _
						   "document.write('<script language=javascript src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/CreateJs/RightBottom.js></script>');" & vbCrLf & _
				           "document.write('<script src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/show.asp?Location="&RightBottomLocation&"></script>');"
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(Str_SysDir)) = false then
				MyFile.CreateFolder(Server.MapPath(Str_SysDir))
			End If
			 if MyFile.FileExists(Server.MapPath(Str_SysDir)&"/"& RightBottomLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(Str_SysDir)&"/"& RightBottomLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(Str_SysDir)&"/"& RightBottomLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			RightBottomObj.close
			Set RightBottomObj = Nothing
		 End Sub
		 
		 Sub DriftBox(TempLocation)
		    dim DriftBoxStr,DriftBoxLocation,DriftBoxObj
			DriftBoxLocation = clng(TempLocation)
			AdsTempPicStr(DriftBoxLocation)
			DriftBoxStr = AdsTempStr
			Set DriftBoxObj = Conn.Execute("select * from FS_AD_Info where AdID="&CintStr(DriftBoxLocation)&"")
			if CheckAd(DriftBoxObj)=False then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				AdsJSStr = "DriftBoxStr=(document.layers)?true:false;" & vbCrLf & _
						   "if(DriftBoxStr){document.write('<layer id=DriftBox width="& DriftBoxObj("AdPicWidth") &" height="& DriftBoxObj("AdPicHeight") &" onmouseover=DriftBoxSM(""DriftBox"") onmouseout=movechip(""DriftBox"")>"& DriftBoxStr &"</layer>');}" & vbCrLf & _
						   "else{document.write('<div id=DriftBox style=""position:absolute; width:"& DriftBoxObj("AdPicWidth") &"px; height:"& DriftBoxObj("AdPicHeight") &"px; z-index:9; filter: Alpha(Opacity=90)"" onmouseover=DriftBoxSM(""DriftBox"") onmouseout=movechip(""DriftBox"")>"& DriftBoxStr &"</div>');}" & vbCrLf & _
						   "document.write('<script language=javascript src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/CreateJs/DriftBox.js></script>');" & vbCrLf & _
				           "document.write('<script src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/show.asp?Location="&DriftBoxLocation&"></script>');"
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(Str_SysDir)) = false then
				MyFile.CreateFolder(Server.MapPath(Str_SysDir))
			End If
			 if MyFile.FileExists(Server.MapPath(Str_SysDir)&"/"& DriftBoxLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(Str_SysDir)&"/"& DriftBoxLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(Str_SysDir)&"/"& DriftBoxLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			DriftBoxObj.close
			Set DriftBoxObj = Nothing
		 End Sub
		 
		 Sub LeftBottom(TempLocation)
		    dim LeftBottomStr,LeftBottomLocation,LeftBottomObj
			LeftBottomLocation = clng(TempLocation)
			AdsTempPicStr(LeftBottomLocation)
			LeftBottomStr = AdsTempStr
			Set LeftBottomObj = Conn.Execute("select * from FS_AD_Info where AdID="&CintStr(LeftBottomLocation)&"")
			if CheckAd(LeftBottomObj)=False then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				AdsJSStr = "if (navigator.appName == 'Netscape')" & vbCrLf & _
						   "{document.write('<layer id=LeftBottom top=150 width="& LeftBottomObj("AdPicWidth") &" height="& LeftBottomObj("AdPicHeight") &">"& LeftBottomStr &"</layer>');}" & vbCrLf & _
						   "else{document.write('<div id=LeftBottom style=""position: absolute;width:"& LeftBottomObj("AdPicWidth") &";height:"& LeftBottomObj("AdPicHeight") &";visibility: visible;z-index: 1"">"& LeftBottomStr &"</div>');}" & vbCrLf & _
						   "document.write('<script language=javascript src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/CreateJs/LeftBottom.js></script>');" & vbCrLf & _
				           "document.write('<script src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/show.asp?Location="&LeftBottomLocation&"></script>');"
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(Str_SysDir)) = false then
				MyFile.CreateFolder(Server.MapPath(Str_SysDir))
			End If
			 if MyFile.FileExists(Server.MapPath(Str_SysDir)&"/"& LeftBottomLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(Str_SysDir)&"/"& LeftBottomLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(Str_SysDir)&"/"& LeftBottomLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			LeftBottomObj.close
			Set LeftBottomObj = Nothing
		 End Sub
		 
		 Sub Couplet(TempLocation)
		    dim CoupletLeftStr,CoupletLocation,CoupletRightStr,CoupletObj
			CoupletLocation = clng(TempLocation)
			AdsTempPicStr(CoupletLocation)
			CoupletLeftStr = AdsTempStr
			CoupletRightStr = AdsTempStrRight
			Set CoupletObj = Conn.Execute("select * from FS_Ad_Info where AdID="&CintStr(CoupletLocation)&"")
			if CheckAd(CoupletObj)=False then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				AdsJSStr =  "function winload()" & vbCrLf & _
							"{" & vbCrLf & _
							"AdsLayerLeft.style.top=20;" & vbCrLf & _
							"AdsLayerLeft.style.left=5;" & vbCrLf & _
							"AdsLayerRight.style.top=20;" & vbCrLf & _
							"AdsLayerRight.style.right=5;" & vbCrLf & _
							"}" & vbCrLf & _
							"if(screen.availWidth>800){" & vbCrLf & _
							"{" & vbCrLf & _
							"document.write('<div id=AdsLayerLeft style=""position: absolute;visibility:visible;z-index:1""><table  border=0 cellspacing=0 cellpadding=0><tr><td>" & CoupletLeftStr & "</td></tr><tr><td height=""20"" align=""left"" valign=""middle""><span onclick=""Javascript:$(\'AdsLayerLeft\').style.display=\'none\';$(\'AdsLayerRight\').style.display=\'none\';"" style=""cursor:pointer;"">×关闭</span></td></tr></table></div>'" & vbCrLf & _
							"+'<div id=AdsLayerRight style=""position: absolute;visibility:visible;z-index:1""><table border=0 cellspacing=0 cellpadding=0><tr><td>" & CoupletRightStr & "</td></tr><tr><td height=""20"" align=""right"" valign=""middle""><span onclick=""Javascript:$(\'AdsLayerLeft\').style.display=\'none\';$(\'AdsLayerRight\').style.display=\'none\';"" style=""cursor:pointer;"">×关闭</span></td></tr></table></div>');" & vbCrLf & _
							"}" & vbCrLf & _
							"document.write('<script language=javascript src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/CreateJs/Couplet.js></script>');" & vbCrLf & _
							"winload()" & vbCrLf & _
							"}" & vbCrLf & _
				           "document.write('<script src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/show.asp?Location="&CoupletLocation&"></script>');"
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(Str_SysDir)) = false then
				MyFile.CreateFolder(Server.MapPath(Str_SysDir))
			End If
			 if MyFile.FileExists(Server.MapPath(Str_SysDir)&"/"& CoupletLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(Str_SysDir)&"/"& CoupletLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(Str_SysDir)&"/"& CoupletLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			CoupletObj.close
			Set CoupletObj = Nothing
		 End Sub
		 
		 Sub AdTxt(AdID)
		 	Dim str_AdTxt,o_AdTxtRs,int_AdID,int_TxtAdID_i,o_AdTxtAdRs,o_Crtxt
			int_AdID=Clng(AdID)
			Set o_AdTxtAdRs=Conn.execute("Select * from FS_Ad_Info where AdID="&CintStr(int_AdID)&"")
			If CheckAd(o_AdTxtAdRs)=False Then
				str_AdTxt="document.write('此广告已经暂停或是失效');"
			Else
				Set o_AdTxtRs=Conn.execute("select * from FS_AD_TxtInfo where AdID="&CintStr(int_AdID)&"")
				str_AdTxt="document.write('<script language=""javascript"" src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/show.asp?Location="&int_AdID&"></script>');"
				If o_AdTxtAdRs("AdTxtColNum")<>"" and o_AdTxtAdRs("AdTxtColNum")>0 Then
					int_TxtAdID_i=0
					str_AdTxt=str_AdTxt&"document.write('<table border=0>"					
					str_AdTxt=str_AdTxt&"<tr>"
					Do While Not o_AdTxtRs.Eof And int_TxtAdID_i<Cint(o_AdTxtAdRs("AdTxtColNum")) 
						str_AdTxt=str_AdTxt&"<td><a href=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/AdsClick.asp?AdTxtID="&o_AdTxtRs("Ad_TxtID")&" class="&o_AdTxtRs("Css")&">"&o_AdTxtRs("AdTxtContent")&"</a></td>"
						o_AdTxtRs.MoveNext
						int_TxtAdID_i=int_TxtAdID_i+1    
						If int_TxtAdID_i = CintStr(o_AdTxtAdRs("AdTxtColNum")) then
							int_TxtAdID_i = 0
							str_AdTxt=str_AdTxt&"</tr><tr>"
						End If
					Loop
					str_AdTxt=str_AdTxt&"</table>');"
				Else
					str_AdTxt=str_AdTxt&"document.write('<table border=0>"					
					While Not o_AdTxtRs.Eof 
						str_AdTxt=str_AdTxt&"<tr>"
						str_AdTxt=str_AdTxt&"<td><a href=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/AdsClick.asp?AdTxtID="&o_AdTxtRs("Ad_TxtID")&" class="&o_AdTxtRs("Css")&">"&o_AdTxtRs("AdTxtContent")&"</a></td>"
						str_AdTxt=str_AdTxt&"</tr>"
						o_AdTxtRs.MoveNext
					Wend 
					str_AdTxt=str_AdTxt&"</table>');"
				End If
				Set MyFile=Server.CreateObject(G_FS_FSO)
				If MyFile.FolderExists(Server.MapPath(Str_SysDir)) = false then
					MyFile.CreateFolder(Server.MapPath(Str_SysDir))
				End If
			 	If MyFile.FileExists(Server.MapPath(Str_SysDir)&"/"& RightBottomLocation &".js") Then
					MyFile.DeleteFile(Server.MapPath(Str_SysDir)&"/"& RightBottomLocation &".js")
			 	End if
				Set o_Crtxt=MyFile.CreateTextFile(Server.MapPath(Str_SysDir)&"/"& int_AdID &".js")
				o_Crtxt.write str_AdTxt
				Set MyFile=Nothing
				Set o_AdTxtRs=Nothing
				Set o_AdTxtAdRs=Nothing
			End If
		 End Sub
		 
	Sub Cycle(ALocation,TempLocation)
		dim CycleSelfObj,CycleSelfLocation,CycleLocation,CycleObj,JsFileName,CycleStr
		CycleSelfLocation = clng(ALocation)
		CycleLocation = clng(TempLocation)
		Set CycleSelfObj = Conn.Execute("select * from FS_Ad_Info where AdID="&CintStr(CycleSelfLocation)&"")'自身查询
		If Not CycleSelfObj.Eof Then
			If Cint(CycleSelfObj("AdLoop")) = 1 then '所有循环广告	
				If CycleSelfObj("AdLoopAdID")<>0 then '所有被添加到循环广告的非循环广告 
					Set CycleObj = Conn.Execute("select * from FS_Ad_Info where AdLock=0 and AdID="&CycleLocation&"")
					If Not CycleObj.Eof THen
						CycleStr="<a href=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/AdsClick.asp?Location="& CycleObj("AdID") &" title="""&CycleObj("AdCaptionTxt")&""" target=_blank><img src="& CycleObj("AdPicPath")&" width="""&CycleSelfObj("AdPicWidth")&""" height="""&CycleSelfObj("AdPicHeight")&""" border=""0""></a>"
					End If
					Set CycleObj=Nothing
				End if			  
			End if
			Dim str_LoopFollow
			if isnull(CycleSelfObj("AdLoopFollow")) or not isnumeric(CycleSelfObj("AdLoopFollow")) then 
				str_LoopFollow="up"
			else	
				Select Case Cint(CycleSelfObj("AdLoopFollow"))
					Case 0
						str_LoopFollow="up"
					Case 1
						str_LoopFollow="down"
					Case 2
						str_LoopFollow="left"
					Case 3
						str_LoopFollow="right"
				End Select
			end if
			
			AdsJSStr = "document.write('<marquee onmouseout=start() onmouseover=stop() width="&CycleSelfObj("AdPicWidth")&" height="&CycleSelfObj("AdPicHeight")&" direction="&str_LoopFollow&" scrollamount="&CycleSelfObj("AdLoopSpeed")&">"
			If Instr(1,LCase(CycleSelfObj("AdPicPath")),"http://") <> 0 then
				AdsJSStr = AdsJSStr & " <a href=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/AdsClick.asp?Location="& CycleSelfObj("AdID") &" title="""&CycleSelfObj("AdCaptionTxt")&""" target=_blank><img src="& CycleSelfObj("AdPicPath")&" width="""&CycleSelfObj("AdPicWidth")&""" height="""&CycleSelfObj("AdPicHeight")&""" border=""0""></a>"&CycleStr
			Else
				AdsJSStr = AdsJSStr & " <a href=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/Ads/AdsClick.asp?Location="& CycleSelfObj("AdID") &" title="""&CycleSelfObj("AdCaptionTxt")&""" target=_blank><img src=http://"&request.Cookies("FoosunMFCookies")("FoosunMFDomain")& CycleSelfObj("AdPicPath")&" width="""&CycleSelfObj("AdPicWidth")&""" height="""&CycleSelfObj("AdPicHeight")&""" border=""0""></a>"&CycleStr
			End If
			CycleSelfObj.movenext
			if not CycleSelfObj.eof then
				if CintStr(CycleSelfObj("AdLoopFollow")) = 0 or CintStr(CycleSelfObj("AdLoopFollow")) = 1 then 
					AdsJSStr = AdsJSStr & "<br><br>"
				else
					AdsJSStr = AdsJSStr & "&nbsp;&nbsp;"
				end if
			end if
			AdsJSStr = AdsJSStr & "</marquee>');"
			if CheckAd(CycleSelfObj)=False  then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			end if
			
			JsFileName = clng(CycleSelfLocation)
			
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(Str_SysDir)) = false then
				MyFile.CreateFolder(Server.MapPath(Str_SysDir))
			End If
			if MyFile.FileExists(Server.MapPath(Str_SysDir)&"/"& JsFileName &".js") then
				MyFile.DeleteFile(Server.MapPath(Str_SysDir)&"/"& JsFileName &".js")
			end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(Str_SysDir)&"/"& JsFileName &".js")
			CrHNJS.write AdsJSStr
			set MyFile=nothing
			CycleSelfObj.close
		End If
		Set CycleSelfObj = Nothing			
	End Sub

	Function CheckAd(obj)
		If Not obj.Eof Then
			If CintStr(obj("AdLock"))=1 Then
				CheckAd=False
				Exit Function
			End If
			If CintStr(obj("AdLoopFactor"))=1 Then
				If CLng(obj("AdClickNum")) > CLng(obj("AdMaxClickNum")) or CLng(obj("AdShowNum")) > CLng(obj("AdMaxShowNum")) Then
					CheckAd=False
					Exit Function
				End If
				If Trim(obj("AdEndDate"))<>"" And not IsNull(obj("AdEndDate")) Then
					If Cdate(obj("AdEndDate"))<Now() Then
						CheckAd=False
						Exit Function
					End If		
				End If
			Else
				CheckAd=True
				Exit Function
			End If
		Else
			CheckAd=""
			Exit Function
		End If
		CheckAd=True
	End Function
'-------------生成广告JS结束----------------  

'-----------------------------------------------------
	Function UpdateAdsJsContent(AdsID)
		Dim AdsObj,Ads_ID,Ads_Type,AdLoopAdID
		Set AdsObj = Conn.ExeCute("Select AdID,AdType,AdLoopAdID From FS_AD_Info Where AdID = " & CintStr(AdsID))
		If Not(AdsObj.Bof And AdsObj.Eof) Then
			Ads_ID = Clng(AdsObj(0))
			Ads_Type = Cint(AdsObj(1))
			AdLoopAdID = Clng(AdsObj(2))
			Select Case Ads_Type
				Case 0 call ShowAds(Ads_ID)
				Case 1 call NewWindow(Ads_ID)
				Case 2 call OpenWindow(Ads_ID)
				Case 3 call FilterAway(Ads_ID)
				Case 4 call DialogBox(Ads_ID)
				Case 5 call ClarityBox(Ads_ID)
				Case 6 call DriftBox(Ads_ID)
				Case 7 call LeftBottom(Ads_ID)
				Case 8 call RightBottom(Ads_ID)
				Case 9 call Couplet(Ads_ID)
				Case 10 call Cycle(Ads_ID,AdLoopAdID)
				Case 11 call AdTxt(Ads_ID)
			End Select	
		End If
		AdsObj.CLose : Set AdsObj = Nothing
	End Function
%>








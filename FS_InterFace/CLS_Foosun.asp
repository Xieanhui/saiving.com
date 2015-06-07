<%
Class CLS_FoosunRecordSet
	Private m_DictPara
	Private Sub Class_Initialize()
		Set m_DictPara = Server.CreateObject(G_FS_DICT)
	End Sub
	Private Sub Class_Terminate()
		m_DictPara.RemoveAll
		Set m_DictPara = Nothing
	End Sub
	'Values
	Public Property Set Values(f_Fields,f_RS)
		Dim i,f_Array,f_Field
		m_DictPara.RemoveAll
		f_Array = Split(f_Fields,",")
		For i = LBound(f_Array) To UBound(f_Array)
			f_Field = LCase(f_Array(i))
			on error resume next
			m_DictPara.Add f_Field,f_RS(f_Field)
			if err.number <> 0 then
				Response.Write(f_Field)
				Response.End
			end if
		Next
	End Property
	'Value
	Public Property Let Value(f_Field,f_Value)
		f_Field = LCase(f_Field)
		if Not m_DictPara.Exists(f_Field) then
			m_DictPara.Add f_Field,f_Value
		else
			m_DictPara.Item(f_Field) = f_Value
		end if
	End Property
	Public Property Get Value(f_Field)
		f_Field = LCase(f_Field)
		Value = m_DictPara.Item(f_Field)
	End Property
End Class

Class CLS_FoosunStyle
	Public StyleID,StyleContent,StyleLoopContent,StyleType,IsParsed
	Private Sub Class_Initialize()
		IsParsed = False
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CLS_FoosunLable
	Private m_Next,m_LableName,m_FoosunStyle
	Public DictLableContent,IsParsed,SysID,LableFun,IsSave,IsPagination
	Public f_LableChar '测试使用
	'Initial
	Private Sub Class_Initialize()
		Set m_Next = Nothing
		Set DictLableContent = Server.CreateObject(G_FS_DICT)
		DictLableContent.Add "0",""
		Set m_FoosunStyle = New CLS_FoosunStyle
		IsParsed = False
		IsSave = False
		IsPagination = False
	End Sub
	Private Sub Class_Terminate()
		Set m_Next = Nothing
		DictLableContent.RemoveAll
		Set DictLableContent = Nothing
		Set m_FoosunStyle = Nothing
	End Sub
	'Next_
	Public Property Set Next_(f_FoosunLable)
		Set m_Next = f_FoosunLable
	End Property
	Public Property Get Next_()
		Set Next_ = m_Next
	End Property
	'FoosunStyle
	Public Property Set FoosunStyle(f_FoosunStyle)
		Set m_FoosunStyle = f_FoosunStyle
	End Property
	Public Property Get FoosunStyle()
		Set FoosunStyle = m_FoosunStyle
	End Property
	'LableName
	Public Property Let LableName(f_LableName)
		Dim f_REG,f_Matchs,f_MatchValue,f_Array
		m_LableName = f_LableName
		'测试使用
		f_LableChar = Mid(f_LableName,8,Len(f_LableName) - 8)
		'测试使用
		if m_LableName <> "" then
			Set f_REG = New RegExp
			f_REG.IgnoreCase = True
			f_REG.Global = True
			f_REG.Pattern = "(\{FS\:)\w*\=[^\$┆\}]*"
			Set f_Matchs = f_REG.Execute(m_LableName)
			if f_Matchs.Count >= 1 then
				f_MatchValue = f_Matchs(0).Value
				f_MatchValue = Replace(f_MatchValue,"=",":")
				f_Array = Split(f_MatchValue,":")
				if UBound(f_Array) >= 2 then
					SysID = Trim(f_Array(1))
					LableFun = Trim(f_Array(2))
				else
					SysID = ""
					LableFun = ""
				end if
			else
				SysID = ""
				LableFun = ""
			end if
			Set f_Matchs = Nothing
		else
			SysID = ""
			LableFun = ""
		end if
		m_FoosunStyle.StyleID = ME.LablePara("引用样式")
	End Property
	Public Property Get LableName()
		LableName = m_LableName
	End Property
	'LablePara
	Public Property Get LablePara(f_ParaName)
		Dim f_REG,f_Matchs,f_MatchValue,f_Array
		if m_LableName <> "" then
			Set f_REG = New RegExp
			f_REG.IgnoreCase = True
			f_REG.Global = True
			f_REG.Pattern = "" & f_ParaName & "\s*\$\s*[^\$┆\}]*"
			Set f_Matchs = f_REG.Execute(m_LableName)
			if f_Matchs.Count >= 1 then
				f_MatchValue = f_Matchs(0).Value
				f_Array = Split(f_MatchValue,"$")
				if UBound(f_Array) >= 1 then
					LablePara = Trim(f_Array(1))
				else
					LablePara = ""
				end if
			else
				LablePara = ""
			end if
			Set f_Matchs = Nothing
		else
			LablePara = ""
		end if
	End Property
End Class

Class CLS_FoosunLableMass
	Private m_Next
	Public MassName,m_MassContent,ParseContent,IsParsed,IsSave,IsPagination,FoosunLableList
	'Initial
	Private Sub Class_Initialize()
		Set m_Next = Nothing
		Set FoosunLableList = New CLS_FoosunList
		IsParsed = False
		IsSave = False
		IsPagination = False
	End Sub
	Private Sub Class_Terminate()
		Set m_Next = Nothing
		Set FoosunLableList = Nothing
	End Sub
	'Next_
	Public Property Set Next_(f_FoosunLableMass)
		Set m_Next = f_FoosunLableMass
	End Property
	Public Property Get Next_()
		Set Next_ = m_Next
	End  Property
	'Next_
	Public Property Let MassContent(f_MassContent)
		m_MassContent = f_MassContent
		ParseContent = f_MassContent
	End Property
	Public Property Get MassContent()
		MassContent = m_MassContent
	End  Property
End Class

Class CLS_FoosunList
	Private m_FirstListItem
	'Initial
	Private Sub Class_Initialize()
		Set m_FirstListItem = Nothing
	End Sub
	'Add
	Public Sub Add(f_ListItem)
		Dim f_TempListItem
		If m_FirstListItem IS Nothing Then
			Set m_FirstListItem = f_ListItem
		Else
			Set f_TempListItem = m_FirstListItem
			While Not f_TempListItem.Next_ IS Nothing
				Set f_TempListItem=f_TempListItem.Next_
			Wend  
			Set f_TempListItem.Next_ = f_ListItem
		End If
	End Sub  
	'RemoveAt
	Public Sub RemoveAt(f_Index)
		Dim f_TempListItem,i,f_Length
		Set f_TempListItem = m_FirstListItem
		f_Length = ME.Length
		If f_Index < 0 OR f_Index > f_Length - 1 Then Err.Raise vbObjectError + 1,"","索引越界"
		For i = 1 TO f_Index - 1
			Set f_TempListItem = f_TempListItem.Next_
		Next  
		If f_Index = 0 Then
			Set m_FirstListItem = m_FirstListItem.Next_
		ElseIf f_Index = f_Length - 1 Then
			Set f_TempListItem.Next_ = Nothing
		Else
			Set f_TempListItem.Next_ = f_TempListItem.Next_.Next_
		End If
	End Sub
	'Items
	Public Function Items(f_Index)
		Dim f_TempListItem,i,f_Length
		Set f_TempListItem = m_FirstListItem
		f_Length = ME.Length
		if f_Index < 0 OR f_Index > f_Length - 1 Then Err.Raise vbObjectError + 1,"","索引越界"
		For i = 0 TO f_Index - 1
			Set f_TempListItem = f_TempListItem.Next_
		Next
		Set Items = f_TempListItem
	End Function
	'Length
	Public Property Get Length()
		Dim f_TempListItem,f_Num
		f_Num = 0
		Set f_TempListItem = m_FirstListItem
		While Not f_TempListItem IS Nothing
			Set f_TempListItem = f_TempListItem.Next_
			f_Num = f_Num + 1
		Wend
		Length = f_Num
	End Property
	'Clear
	Public Sub Clear()
		Set m_FirstListItem = Nothing
	End Sub
End Class

Class CLS_RefreshSession
	Private m_LableMass,m_Lable,m_StyleContent,m_StyleLoopContent,m_StyleType
	Private m_ItemSeperator,m_AttributeSeperator
	'Initial
	Private Sub Class_Initialize()
		m_ItemSeperator = "$$$"
		m_AttributeSeperator = "###"
		Set m_LableMass = Server.CreateObject(G_FS_DICT)
		Set m_Lable = Server.CreateObject(G_FS_DICT)
		Set m_StyleContent = Server.CreateObject(G_FS_DICT)
		Set m_StyleLoopContent = Server.CreateObject(G_FS_DICT)
		Set m_StyleType = Server.CreateObject(G_FS_DICT)
		Dict_Initialize Session("FOOSUN_LABLE_MASS"),m_LableMass
		Dict_Initialize Session("FOOSUN_LABLE"),m_Lable
		Dict_Initialize Session("FOOSUN_STYLECONTENT"),m_StyleContent
		Dict_Initialize Session("FOOSUN_STYLELOOPCONTENT"),m_StyleLoopContent
		Dict_Initialize Session("FOOSUN_STYLETYPE"),m_StyleType
	End Sub
	Private Sub Dict_Initialize(f_DictStr,f_DictObj)
		Dim i,f_ItemDictArray,f_AttributDictArray
		f_ItemDictArray = Split(f_DictStr,m_ItemSeperator)
		For i = LBound(f_ItemDictArray) To UBound(f_ItemDictArray)
			f_AttributDictArray = Split(f_ItemDictArray(i),m_AttributeSeperator)
			if UBound(f_AttributDictArray) >= 1 then
				if Not f_DictObj.Exists(f_AttributDictArray(0)) Then f_DictObj.Add f_AttributDictArray(0),f_AttributDictArray(1)
			End If
		Next
	End Sub
	Private Sub Dict_Terminate(SessionName,f_Dict)
		Dim i,f_Items,f_Keys
		f_Items = f_Dict.Items
		f_Keys = f_Dict.Keys
		for i = 0 to f_Dict.Count - 1
			if Session(SessionName) = "" then
				Session(SessionName) = f_Keys(i) & m_AttributeSeperator & f_Items(i)
			else
				Session(SessionName) = Session(SessionName) & m_ItemSeperator & f_Keys(i) & m_AttributeSeperator & f_Items(i)
			end if
		Next
	End Sub
	Private Sub Class_Terminate()
		Session("FOOSUN_LABLE_MASS") = ""
		Session("FOOSUN_LABLE") = ""
		Session("FOOSUN_STYLECONTENT") = ""
		Session("FOOSUN_STYLELOOPCONTENT") = ""
		Session("FOOSUN_STYLETYPE") = ""
		Dict_Terminate "FOOSUN_LABLE_MASS",m_LableMass
		Dict_Terminate "FOOSUN_LABLE",m_Lable
		Dict_Terminate "FOOSUN_STYLECONTENT",m_StyleContent
		Dict_Terminate "FOOSUN_STYLELOOPCONTENT",m_StyleLoopContent
		Dict_Terminate "FOOSUN_STYLETYPE",m_StyleType
		m_LableMass.RemoveAll
		m_Lable.RemoveAll
		m_StyleContent.RemoveAll
		m_StyleLoopContent.RemoveAll
		m_StyleType.RemoveAll
		Set m_LableMass = Nothing
		Set m_Lable = Nothing
		Set m_StyleContent = Nothing
		Set m_StyleLoopContent = Nothing
		Set m_StyleType = Nothing
	End Sub
	'Content
	Public Property Get Content(f_Type,f_Key)
		Select Case f_Type
			Case 1
				Content = m_LableMass.Item(f_Key)
			Case 2
				Content = m_Lable.Item(f_Key)
			Case 31
				Content = m_StyleContent.Item(f_Key)
			Case 32
				Content = m_StyleLoopContent.Item(f_Key)
			Case 33
				Content = m_StyleType.Item(f_Key)
			Case Else
				Content = ""
		End Select
	End Property
	'Add
	Public Sub Add(f_Type,f_Key,f_Item)
		Select Case f_Type
			Case 1
				if Not m_LableMass.Exists(f_Key) then m_LableMass.Add f_Key,f_Item
			Case 2
				if Not m_Lable.Exists(f_Key) then m_Lable.Add f_Key,f_Item
			Case 31
				if Not m_StyleContent.Exists(f_Key) then m_StyleContent.Add f_Key,f_Item
			Case 32
				if Not m_StyleLoopContent.Exists(f_Key) then m_StyleLoopContent.Add f_Key,f_Item
			Case 33
				if Not m_StyleType.Exists(f_Key) then m_StyleType.Add f_Key,f_Item
			Case Else
				
		End Select
	End Sub
	'Exists
	Public Property Get Exists(f_Type,f_Key)
		Select Case f_Type
			Case 1
				Exists = m_LableMass.Exists(f_Key)
			Case 2
				Exists = m_Lable.Exists(f_Key)
			Case 31
				Exists = m_StyleContent.Exists(f_Key)
			Case 32
				Exists = m_StyleLoopContent.Exists(f_Key)
			Case 33
				Exists = m_StyleType.Exists(f_Key)
			Case Else
				Exists = False
		End Select
	End Property
End Class

Class CLS_FoosunLink
	Private m_NS_LinkType,m_NS_IsDomain,m_MF_Domain,m_NS_NewsDir
	Private Sub Class_Initialize()
		m_NS_LinkType = Request.Cookies("FoosunNSCookies")("FoosunNSLinkType")
		m_NS_IsDomain = Request.Cookies("FoosunNSCookies")("FoosunNSDomain")
		m_MF_Domain = Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
		m_NS_NewsDir = Request.Cookies("FoosunNSCookies")("FoosunNSNewsDir")
	End Sub
	Private Sub Class_Terminate()
	End Sub
	Public Property Get SpecialLink(f_SpecialLinkRecordSet)
		dim SpecialEName,ExtName,c_SavePath,Url_Domain,Special_SavePath,Special_SaveType
		SpecialEName = f_SpecialLinkRecordSet.Value("SpecialEName")
		ExtName = f_SpecialLinkRecordSet.Value("ExtName")
		c_SavePath = f_SpecialLinkRecordSet.Value("SavePath")
		Special_SaveType = Cint(f_SpecialLinkRecordSet.Value("FileSaveType"))
		if m_NS_LinkType = 1 then
			if trim(m_NS_IsDomain)<>"" then
				Url_Domain = "http://"&m_NS_IsDomain
				If m_NS_NewsDir <> "" And m_NS_NewsDir <> "/" Then
					If InStr(LCase(c_SavePath),m_NS_NewsDir) > 0 Then
						c_SavePath = Replace(Replace(LCase(c_SavePath),m_NS_NewsDir,""),"//","/")
					End IF	
				End If
			else
				Url_Domain = "http://"&m_MF_Domain
			end if
		else
			if trim(m_NS_IsDomain)<>"" then
				Url_Domain = "http://"&IsDomain
				If m_NS_NewsDir <> "" And m_NS_NewsDir <> "/" Then
					If InStr(LCase(c_SavePath),m_NS_NewsDir) > 0 Then
						c_SavePath = Replace(Replace(LCase(c_SavePath),m_NS_NewsDir,""),"//","/")
					End IF	
				End If
			else
				if G_VIRTUAL_ROOT_DIR<>"" then
					Url_Domain = "/"& G_VIRTUAL_ROOT_DIR
				else
					Url_Domain = ""
				end if
			end if
		end if
		Select Case Special_SaveType
			Case 0
				Special_SavePath = c_SavePath & "/" & SpecialEName & "/index." & ExtName
			Case 1
				Special_SavePath = c_SavePath & "/" & SpecialEName & "/" & SpecialEName & "." & ExtName	
			Case 2
				Special_SavePath = c_SavePath & "/" & SpecialEName & "." & ExtName
			Case 3
				Special_SavePath = c_SavePath & "/Special_" & SpecialEName & "." & ExtName			
		End Select
		SpecialLink = Url_Domain & Replace(Special_SavePath,"//","/")
	End Property
	Public Property Get ClassLink(f_ClassLinkRecordSet)
		Dim ClassEName,c_Domain,Url_Domain,ClassSaveType,class_savepath,FileExtName,c_SavePath
		if f_ClassLinkRecordSet.Value("IsURL") = 0 OR f_ClassLinkRecordSet.Value("IsURL") = 2 then
			ClassEName = f_ClassLinkRecordSet.Value("ClassEName")
			c_Domain = f_ClassLinkRecordSet.Value("Domain")
			FileExtName = f_ClassLinkRecordSet.Value("FileExtName")
			ClassSaveType = f_ClassLinkRecordSet.Value("FileSaveType")
			c_SavePath= f_ClassLinkRecordSet.Value("SavePath")
			if trim(c_Domain)<>"" then
				if ClassSaveType=0 then
					class_savepath = "index."&FileExtName
				elseif ClassSaveType=1 then
					class_savepath =ClassEName &"."&FileExtName
				else
					class_savepath = ClassEName &"."&FileExtName
				end if
			else
				if ClassSaveType=0 then
					class_savepath = ClassEName&"/index."&FileExtName
				elseif ClassSaveType=1 then
					class_savepath = ClassEName&"/"& ClassEName &"."&FileExtName
				else
					class_savepath = ClassEName &"."&FileExtName
				end if
			end if
			if m_NS_LinkType = 1 then
				if trim(c_Domain)<>"" then
					Url_Domain = "http://"&c_Domain
				else
					if trim(m_NS_IsDomain)<>"" then
						Url_Domain = "http://"&m_NS_IsDomain
						If m_NS_NewsDir <> "" And m_NS_NewsDir <> "/" Then
							If InStr(LCase(c_SavePath),m_NS_NewsDir) > 0 Then
								c_SavePath = Replace(Replace(LCase(c_SavePath),m_NS_NewsDir,""),"//","/")
							End IF	
						End If	
					else
						Url_Domain = "http://"&m_MF_Domain
						c_SavePath = c_SavePath
					end if
				end if
			else
				if trim(c_Domain)<>"" then
					Url_Domain = "http://"&c_Domain
				else
					if trim(m_NS_IsDomain)<>"" then
						Url_Domain = "http://"&m_NS_IsDomain
						If m_NS_NewsDir <> "" And m_NS_NewsDir <> "/" Then
							If InStr(LCase(c_SavePath),m_NS_NewsDir) > 0 Then
								c_SavePath = Replace(Replace(LCase(c_SavePath),m_NS_NewsDir,""),"//","/")
							End IF	
						End If
					else
						if G_VIRTUAL_ROOT_DIR<>"" then
							Url_Domain = "/"& G_VIRTUAL_ROOT_DIR
						else
							Url_Domain = ""
						end if
						c_SavePath = c_SavePath
					end if
				end if
			end if
			if trim(c_Domain)<>"" then
				ClassLink = Url_Domain&Replace("/"&class_savepath,"//","/")
			else
				ClassLink = Url_Domain&Replace(c_SavePath&"/"&class_savepath,"//","/")
			end if
		else
			ClassLink = f_ClassLinkRecordSet.Value("UrlAddress")
		end if
	End Property
	Public Property Get NewsLink(f_NewsLinkRecordSet)
		Dim SaveNewsPath,FileName,FileExtName,Url_Domain,ClassEName,c_Domain,c_SavePath
		if f_NewsLinkRecordSet.Value("IsURL") = 0 then
			SaveNewsPath = f_NewsLinkRecordSet.Value("SaveNewsPath")
			FileName = f_NewsLinkRecordSet.Value("FileName")
			FileExtName = f_NewsLinkRecordSet.Value("FileExtName")
			ClassEName = f_NewsLinkRecordSet.Value("ClassEName")
			c_Domain = f_NewsLinkRecordSet.Value("Domain")
			c_SavePath = f_NewsLinkRecordSet.Value("SavePath")
			if m_NS_LinkType = 1 then
				if trim(c_Domain)<>"" then
					Url_Domain = "http://"&c_Domain
				else
					if trim(m_NS_IsDomain)<>"" then
						Url_Domain = "http://"&m_NS_IsDomain
						c_SavePath = ""
					else
						Url_Domain = "http://"&m_MF_Domain
						c_SavePath = c_SavePath
					end if
				end if
			else
				if trim(c_Domain)<>"" then
					Url_Domain = "http://"&c_Domain
				else
					if trim(m_NS_IsDomain)<>"" then
						Url_Domain = "http://"&m_NS_IsDomain
						c_SavePath = ""
					else
						if G_VIRTUAL_ROOT_DIR<>"" then
							Url_Domain = "/"& G_VIRTUAL_ROOT_DIR
						else
							Url_Domain = ""
						end if
						c_SavePath = c_SavePath
					end if
				end if
			end if
			if trim(c_Domain)<>"" then
				NewsLink = Url_Domain & replace(SaveNewsPath &"/"&FileName&"."&FileExtName,"//","/")
			else
				NewsLink = Url_Domain & replace(c_SavePath& "/" & ClassEName &SaveNewsPath &"/"&FileName&"."&FileExtName,"//","/")
			end if
		else
			NewsLink = f_NewsLinkRecordSet.Value("URLAddress")
		end if
	End Property
End Class
%>
<!--#include file="Api_Config.asp"-->
<%
'=========================================================
' File: Api_Config.asp
' Version:1.0
' Date: 2007-3-8
' Code by Terry
'=========================================================

Class PassportApi
	Public AppID,Status,GetData,GetAppid
	Private XmlDoc,XmlHttp
	Private MessageCode,ArrUrls,SysKey,XmlPath

	Private Sub Class_Initialize()
		GetAppid = ""
		AppID = "FoosunCMS"
		ArrUrls = Split(Trim(API_Urls),"|")
		Status = "1"
		SysKey = API_SysKey
		MessageCode = ""
		XmlPath = "API/api_user.xml"
		XmlPath = Server.MapPath(XmlPath)
		Set XmlDoc = Server.CreateObject(G_MSXML2_DOCUMENT & MsxmlVersion)
		Set GetData = Server.Createobject(G_FS_DICT)
		XmlDoc.ASYNC = False
		LoadXmlData()
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(XmlDoc) Then Set XmlDoc = Nothing
		If IsObject(GetData) Then Set GetData = Nothing
	End Sub

	Public Sub LoadXmlData()
		If Not XmlDoc.Load(XmlPath) Then
			XmlDoc.LoadXml "<?xml version=""1.0"" encoding=""gb2312""?><root/>"
		End If
		NodeValue "appID",AppID,1,False
	End Sub
	
	'--------------------------------------------------
	'参数 ：
	'NodeName 节点名
	'NodeText 值
	'NodeType 保存类型 [text=0,cdata=1] 
	'blnEncode 是否编码 [true,false]
	'--------------------------------------------------
	Public Sub NodeValue(Byval NodeName,Byval NodeText,Byval NodeType ,Byval blnEncode)
		Dim ChildNode,CreateCDATASection
		NodeName = Lcase(NodeName)
		If XmlDoc.documentElement.selectSingleNode(NodeName) is nothing Then
			Set ChildNode = XmlDoc.documentElement.appendChild(XmlDoc.createNode(1,NodeName,""))
		Else
			Set ChildNode = XmlDoc.documentElement.selectSingleNode(NodeName)
		End If
		If blnEncode = True Then
			NodeText = AnsiToUnicode(NodeText)
		End If
		If NodeType = 1 Then
			ChildNode.Text = ""
			Set CreateCDATASection = XmlDoc.createCDATASection(Replace(NodeText,"]]>","]]&gt;"))
			ChildNode.appendChild(createCDATASection)
		Else
			ChildNode.Text = NodeText
		End If
	End Sub

	'--------------------------------------------------
	'获取发送包XML中节点的值
	'参数 ：
	'Str 节点名
	'--------------------------------------------------
	Public Property Get XmlNode(Byval Str)
		If XmlDoc.documentElement.selectSingleNode(Str) is Nothing Then
			XmlNode = "Null"
		Else
			XmlNode = XmlDoc.documentElement.selectSingleNode(Str).text
		End If
	End Property

	'--------------------------------------------------
	'获取返回XML数据对象
	'例：
	'API_Obj.GetAppid = "dvbbs"
	'If API_Obj.GetXmlData<>Null Then Response.Write API_Obj.GetXmlData.xml
	'当GetXmlData不为NULL时，GetXmlData为XML对象
	'--------------------------------------------------
	Public Property Get GetXmlData()
		Dim GetXmlDoc
		GetXmlData = Null
		If GetAppid <> "" Then
			GetAppid = Lcase(GetAppid)
			If GetData.Exists(GetAppid) Then
				Set GetXmlData = GetData(GetAppid)
			End If
		End If
	End Property

	Public Sub SendHttpData()
		Dim i,GetXmlDoc,LoadAppid
		Set Xmlhttp = Server.CreateObject(G_MSXML2_SERVERXMLHTTP & MsxmlVersion)
		Set GetXmlDoc = Server.CreateObject(G_MSXML2_DOCUMENT & MsxmlVersion)
		For i = 0 to Ubound(ArrUrls)
			XmlHttp.Open "POST", Trim(ArrUrls(i)), false
			XmlHttp.SetRequestHeader "content-type", "text/xml"
			XmlHttp.Send XmlDoc
			'Response.Write strAnsi2Unicode(xmlhttp.responseBody)
			'Response.End()
			If GetXmlDoc.load(XmlHttp.responseXML) Then
				LoadAppid = Lcase(GetXmlDoc.documentElement.selectSingleNode("appid").Text)
				GetData.add LoadAppid,GetXmlDoc
				Status = GetXmlDoc.documentElement.selectSingleNode("status").Text
				MessageCode = MessageCode & LoadAppid & "(" & Status &")：" & GetXmlDoc.documentElement.selectSingleNode("body/message").Text
				If Status = "1" Then '当发生错误时退出
					Exit For
				End If
			Else
				Status = "1"
				MessageCode = "请求数据错误！"
				Exit For
			End If
		Next
		Set GetXmlDoc = Nothing
		Set XmlHttp = Nothing
	End Sub

	Public Property Get Message()
		Message = MessageCode
	End Property
	
	'--------------------------------------------------
	'写COOKIE调用 
	'参数
	'C_Syskey 密钥，C_UserName 用户名，C_PassWord 加密的用户密码 ，C_SetType 保存COOKIE时间
	'--------------------------------------------------
	Public Function SetCookie(Byval C_Syskey,Byval C_UserName,Byval C_PassWord,Byval C_SetType)
		Dim i,TempStr
		TempStr = ""
		For i = 0 to Ubound(ArrUrls)
			TempStr = TempStr & vbNewLine & "<script language=""JavaScript"" src="""&Trim(ArrUrls(i))&"?syskey="&Server.URLEncode(C_Syskey)&"&username="&Server.URLEncode(C_UserName)&"&password="&Server.URLEncode(C_PassWord)&"&savecookie="&Server.URLEncode(C_SetType)&"""></script>"
		Next
		SetCookie = TempStr
	End Function
	Public Function SetCookie1(Byval C_Syskey,Byval C_UserName,Byval C_PassWord,Byval C_SetType)
		
		Dim i,GetXmlDoc,LoadAppid,TempStr
		TempStr = "?syskey="&Server.URLEncode(C_Syskey)&"&username="&Server.URLEncode(C_UserName)&"&password="&Server.URLEncode(C_PassWord)&"&savecookie="&Server.URLEncode(C_SetType)
		Set Xmlhttp = Server.CreateObject(G_MSXML2_SERVERXMLHTTP & MsxmlVersion)
		For i = 0 to Ubound(ArrUrls)
			XmlHttp.Open "GET", Trim(ArrUrls(i))&TempStr, false
			XmlHttp.SetRequestHeader "content-type", "text/xml"
			XmlHttp.Send
		Next
		Set XmlHttp = Nothing
	End Function
	'--------------------------------------------------
	'打印发送请求XML数据
	'--------------------------------------------------
	Public Sub PrintXmlData()
		Response.Clear
		Response.ContentType = "text/xml"
		Response.CharSet = "gb2312"
		Response.Expires = 0
		Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"&vbNewLine
		Response.Write XmlDoc.documentElement.XML
	End Sub

	'--------------------------------------------------
	'打印返回XML数据
	'API_Obj.GetAppid = "dvbbs"
	'API_Obj.PrintGetXmlData
	'--------------------------------------------------
	Public Sub PrintGetXmlData()
		Response.Clear
		Response.ContentType = "text/xml"
		Response.CharSet = "gb2312"
		Response.Expires = 0
		Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"&vbNewLine
		Response.Write GetXmlData.documentElement.XML
	End Sub

	Private Function AnsiToUnicode(ByVal str)
		Dim i, j, c, i1, i2, u, fs, f, p
		AnsiToUnicode = ""
		p = ""
		For i = 1 To Len(str)
			c = Mid(str, i, 1)
			j = AscW(c)
			If j < 0 Then
				j = j + 65536
			End If
			If j >= 0 And j <= 128 Then
				If p = "c" Then
					AnsiToUnicode = " " & AnsiToUnicode
					p = "e"
				End If
				AnsiToUnicode = AnsiToUnicode & c
			Else
				If p = "e" Then
					AnsiToUnicode = AnsiToUnicode & " "
					p = "c"
				End If
				AnsiToUnicode = AnsiToUnicode & ("&#" & j & ";")
			End If
		Next
	End Function

	Private Function strAnsi2Unicode(asContents)
		Dim len1,i,varchar,varasc
		strAnsi2Unicode = ""
		len1=LenB(asContents)
		If len1=0 Then Exit Function
		  For i=1 to len1
			varchar=MidB(asContents,i,1)
			varasc=AscB(varchar)
			If varasc > 127  Then
				If MidB(asContents,i+1,1)<>"" Then
					strAnsi2Unicode = strAnsi2Unicode & chr(ascw(midb(asContents,i+1,1) & varchar))
				End If
				i=i+1
			 Else
				strAnsi2Unicode = strAnsi2Unicode & Chr(varasc)
			 End If	
		Next
	End Function
End Class
%>






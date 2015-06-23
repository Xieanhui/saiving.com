<% Option Explicit %>
<!--#include file="FS_Inc/Const.asp" -->
<!--#include file="FS_Inc/Function.asp" -->
<!--#include file="FS_InterFace/MF_Function.asp" -->
<%
Dim Conn
MF_Default_Conn
Dim DownCacheNameStr
DownCacheNameStr = "Http://"&Get_MF_Domain()
Dim ResponseBodyStr,ResponseStr,ErrorStr,RsAddressObj,FileURL
Dim Server_Name,Server_V1,Server_V2
Dim OnlyFileUrlTF '只有文件地址
OnlyFileUrlTF = False     
ResponseBodyStr = "<title>下载</title>" & Chr(13)
ResponseBodyStr = ResponseBodyStr & "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & Chr(13)
ResponseBodyStr = ResponseBodyStr & "<style>body{font-size:9pt;line-height:140%}</style>" & Chr(13)
ResponseBodyStr = ResponseBodyStr & "<body>" & Chr(13)
ErrorStr = "<meta http-equiv='Refresh' content='5; URL="&DownCacheNameStr&"'>" & Chr(13)
ErrorStr = ErrorStr & ResponseBodyStr & Chr(13)
ErrorStr = ErrorStr & "<b>错误!&nbsp;</b>读取地址时出错&nbsp;5秒后自动<a href="&DownCacheNameStr&">返回首页</a>..." & Chr(13)
FileURL = Request("FileUrl")

if (Request("DownLoadID")="" or Request("ID")="" or not isnumeric(Request("ID"))) And FileURL = "" then
	Response.Write ErrorStr
	Set Conn = Nothing
	Response.End
end if

if FileUrl = "" then
	set RsAddressObj=Conn.execute("Select Url from FS_DS_Address where ID=" & CintStr(request("ID")))
	if Not RsAddressObj.Eof then
		FileURL = RsAddressObj("Url")
	else
		RsAddressObj.Close
		Set RsAddressObj = Nothing
		Set Conn = Nothing
		Response.Write ErrorStr
		Response.End
	end if
	RsAddressObj.Close
	OnlyFileUrlTF = False
else
	OnlyFileUrlTF = True
end if
'防盗链
Dim DownLoadConfigObj
Set DownLoadConfigObj = Conn.execute("select top 1 Lock,IPType,IPList,OverDueMode from FS_DS_SysPara")
if DownLoadConfigObj.eof then response.Write("系统错误"): Set DownLoadConfigObj = Nothing:Set Conn = Nothing: response.End()
if DownLoadConfigObj("Lock") = 1 then
	Server_Name = Len(Request.ServerVariables("SERVER_NAME"))
	Server_V1 = Left(Replace(Cstr(Request.ServerVariables("HTTP_REFERER")),"http://",""),Server_Name)
	Server_V2 = Left(Cstr(Request.ServerVariables("SERVER_NAME")),Server_Name)
	if Server_V1 <> Server_V2 and Server_V1 <> "" and Server_V2 <> "" then
		Set DownLoadConfigObj = Nothing
		Set Conn = Nothing
		response.Redirect(DownCacheNameStr)
		Response.End
	end if
end if
'判断过期
if DownLoadConfigObj("OverDueMode") = 2 then
	Dim tmpsql_ 
	if G_IS_SQL_DB = "1" then
		tmpsql_ = "datediff(day,AddTime,'"&date()&"') > OverDue"
	else	
		tmpsql_ = "datediff('d',AddTime,'"&date()&"') > OverDue"
	end if
	set RsAddressObj=Conn.execute("Select ID from FS_DS_List where OverDue>0 and "&tmpsql_&"  and DownLoadID='" & NoSqlHack(Request("DownLoadID")) & "'")
	if not RsAddressObj.eof then 
		RsAddressObj.close
		Response.write("<script>alert('该下载已经过期!');history.back();</script>")
		Response.End		
	else
		Set RsAddressObj = Nothing		
	end if	
end if
'判断IP限制
Dim RequestIPAddress,IPList,IPType,Flag,DownLoadTF
RequestIPAddress = NoSqlHack(Request.ServerVariables("REMOTE_ADDR"))
IPList = DownLoadConfigObj("IPList")
IPType = DownLoadConfigObj("IPType")
Flag = CheckIPAddress(IPList,RequestIPAddress)
'Response.Write(Flag)
'Response.End
if Not IsNull(IPList) And IPList <> "" then
	if Flag = True then
		if IPType = 1 then 
			DownLoadTF = False
		else
			DownLoadTF = True
		end if
	else
		if IPType = 1 then 
			DownLoadTF = True
		else
			DownLoadTF = False
		end if
	end if
else
	DownLoadTF = True
end if

if DownLoadTF then
	if OnlyFileUrlTF = False Then
		Set RsAddressObj = Server.CreateObject(G_FS_RS)
		RsAddressObj.Open "Select ClickNum from FS_DS_List where DownLoadID='" & NoSqlHack(Request("DownLoadID")) & "'",Conn,1,2
		if Not RsAddressObj.eof then
			RsAddressObj("ClickNum") = CLng(RsAddressObj("ClickNum")) + 1
			RsAddressObj.UpDate
		else
			RsAddressObj.Close
			Set RsAddressObj = Nothing
			Set Conn = Nothing
			Response.Write ErrorStr
			Response.End
		end if
	end if
	Set RsAddressObj = Nothing
	if InStr(LCase(FileURl),"://") = 0 then
	'	Response.Redirect FileURL
	'else
	'	downloadFile FileURL
		FileURl = DownCacheNameStr & FileUrl
	end if
	Response.Write "<script>location.href='" & FileURl & "';</script>"
	Response.End
	'Response.Redirect FileURL
else
	Response.write("<script>alert('没有权限,或者IP被锁定');history.back();</script>")
end if
Response.End

Set DownLoadConfigObj = Nothing
Set Conn = Nothing
Function CheckIPAddress(IPList,IPAddress)
	Dim TempArray,i,j,AddressArray,BeginAddressArray,EndAddressArray,IPAddressArray
	IPAddressArray = Split(IPAddress,".")
	if UBound(IPAddressArray) <> 3 then
		CheckIPAddress = False
		Exit Function
	end if
	if IsNull(IPList) then
		CheckIPAddress = False
	else
		if IPList <> "" then
			TempArray = Split(IPList,"$")
			for i = LBound(TempArray) to UBound(TempArray)
				AddressArray = Split(TempArray(i),"-")
				if UBound(AddressArray) = 1 then
					BeginAddressArray = Split(AddressArray(0),".")
					EndAddressArray = Split(AddressArray(1),".")
					if (UBound(BeginAddressArray) = 3) and (UBound(EndAddressArray) = 3) then
						for j = LBound(BeginAddressArray) to UBound(BeginAddressArray)
								'Response.Write(EndAddressArray(j) = BeginAddressArray(j))
							if (EndAddressArray(j) = BeginAddressArray(j)) then
								if EndAddressArray(j) <> IPAddressArray(j) then
									if (CInt(IPAddressArray(j)) >= CInt(BeginAddressArray(j))) And (CInt(IPAddressArray(j)) <= CInt(EndAddressArray(j))) then
										CheckIPAddress = True
										Exit Function
									end if
								end if
							else
								if (CInt(IPAddressArray(j)) >= CInt(BeginAddressArray(j))) And (CInt(IPAddressArray(j)) <= CInt(EndAddressArray(j))) then
									CheckIPAddress = True
									Exit Function
								end if
							end if
						Next
					end if
				end if
				'Response.End
			Next
			CheckIPAddress = False
		else
			CheckIPAddress = False
		end if
	end if
End Function

Function downloadFile(strFile) 
	Dim strFilename,s,fso,intFilelength,f
	strFilename = server.MapPath(strFile) 
	Response.Buffer = True 
	Response.Clear 
	Set s = Server.CreateObject(G_FS_STREAM) 
	s.Open 
	s.Type = 1 
	'on error resume next 
	Set fso = Server.CreateObject(G_FS_FSO) 
	if not fso.FileExists(strFilename) then 
	Response.Write("<h1>Error:</h1>"&strFilename&" does not exists!<p>") 
	Response.End 
	end if 
	
	' get length of file 
	Set f = fso.GetFile(strFilename) 
	intFilelength = f.size 
	
	
	s.LoadFromFile(strFilename) 
	if err then 
		Response.Write("<h1>Error: </h1>Unknown Error!<p>") 
		Response.End 
	end if 
	
	' send the headers to the users Browse 
	Response.AddHeader "Content-Disposition","attachment; filename="& f.name
	Response.AddHeader "Content-Length",intFilelength 
	Response.CharSet = "gb2312" 
	Response.ContentType = "application/octet-stream" 
	
	' output the file to the browser 
	Response.BinaryWrite s.Read 
	Response.Flush 
	
	' tidy up 
	s.Close 
	Set s = Nothing 

End Function 
%>






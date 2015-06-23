<%
Function NoCSSHackAdmin(Str,StrTittle) '过滤跨站脚本和HTML标签
	Dim regEx
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Pattern = "<|>|\t"
	If regEx.Test(LCase(Str)) Then
		Response.Write "<script>alert('"& StrTittle &"含有非法字符(<,>,tab)');history.back();</script>"
		Response.End
	End If
	Set regEx = Nothing
	NoCSSHackAdmin = Str
End Function
Function Str(n,St)
Dim i
for i=1 to n
Str=Str&St
next
End Function
'获得栏目的子栏目
Function GetChildNewsList(ID,f_k)
	Dim ChildRs
	Set ChildRs=server.CreateObject(G_FS_RS)	
	ChildRs.open "select * from FS_NS_NewsClass where ParentID='"&NoSqlHack(ID)&"' Order by Orderid desc",Conn,1,1
	if not ChildRs.eof then
	do while not ChildRs.eof
		GetChildNewsList=GetChildNewsList& ("<option value="&ChildRs("ID")&">"& Str(f_k,"—") &ChildRs("ClassName")&"</option>"& vbcrlf)
		f_k=f_k+1
		GetChildNewsList=GetChildNewsList&GetChildNewsList(ChildRs("ID"),f_k)
		f_k=f_k-1
	ChildRs.movenext
	loop
	end if
	Set ChildRs=nothing
End Function	
Private Function GetOneNewsLinkURL(NewsID)
			Dim DoMain,TempParentID,RsParentObj,RsDoMainObj,ReturnValue
			Dim CheckRootClassIndex,CheckRootClassNumber,TempClassSaveFilePath
			Dim NewsSql,RsNewsObj
			'-----------------------/l
			dim DatePathStr
			CheckRootClassNumber = 30
			ReturnValue = ""
			NewsSql = "Select *,FS_NS_NewsClass.FileExtName as ClassFileExtName,FS_NS_News.FileExtName as NewsFileExtName from FS_NS_News,FS_NS_NewsClass where FS_NS_News.ClassID=FS_NS_NewsClass.ClassID and FS_NS_News.isLock<>1 and FS_NS_News.NewsID='" & NoSqlHack(NewsID) & "'"
			Set RsNewsObj = Conn.Execute(NewsSql)
			if RsNewsObj.Eof then
				Set RsNewsObj = Nothing
				GetOneNewsLinkURL = ""
				Exit Function
			else
				if RsNewsObj("IsURL") = 1 then
					ReturnValue = RsNewsObj("URLAddress")
				else
					if RsNewsObj("ParentID") <> "0" then
						Set RsParentObj = Conn.Execute("Select ParentID,[Domain] from FS_NS_NewsClass where ClassID='" & NoSqlHack(RsNewsObj("ParentID")) & "'")
						if Not RsParentObj.Eof then
							CheckRootClassIndex = 1
							TempParentID = RsParentObj("ParentID")
							do while Not (TempParentID = "0")
								CheckRootClassIndex = CheckRootClassIndex + 1
								RsParentObj.Close
								Set RsParentObj = Nothing
								Set RsParentObj = Conn.Execute("Select ParentID,[Domain] from FS_NS_NewsClass where ClassID='" & NoSqlHack(TempParentID) & "'")
								if RsParentObj.Eof then
									Set RsParentObj = Nothing
									Set RsNewsObj = Nothing
									GetOneNewsLinkURL = ""
									Exit Function
								end if
								TempParentID = RsParentObj("ParentID")
								if CheckRootClassIndex > CheckRootClassNumber then TempParentID = "0" '防止死循环
							Loop
							DoMain = RsParentObj("DoMain")
							Set RsParentObj = Nothing
						else
							Set RsParentObj = Nothing
							Set RsNewsObj = Nothing
							GetOneNewsLinkURL = ""
							Exit Function
						end if
					else
						DoMain = RsNewsObj("DoMain")
					end if
					'---------------/l
					If Application(LoginCacheNameStr)(21)="1" Then DatePathStr=RsNewsObj("Path") else DatePathStr=""
					if (Not IsNull(DoMain)) And (DoMain <> "") then
						ReturnValue = "http://" & DoMain & "/" & RsNewsObj("ClassEName") & DatePathStr &"/" & RsNewsObj("FileName") & "." & RsNewsObj("NewsFileExtName")
					else
						if RsNewsObj("SaveFilePath") = "/" then
							TempClassSaveFilePath = RsNewsObj("SaveFilePath")
						else
							TempClassSaveFilePath = RsNewsObj("SaveFilePath") & "/"
						end if
						ReturnValue = AvailableDoMain & TempClassSaveFilePath & RsNewsObj("ClassEName") & DatePathStr & "/" & RsNewsObj("FileName") & "." & RsNewsObj("NewsFileExtName")
					end if
					'------------------/l
				end if
			end if
			Set RsNewsObj = Nothing
			GetOneNewsLinkURL = ReturnValue
End Function
%>






<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
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
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.Charset="GB2312"
Server.ScriptTimeOut=999999999
Dim Conn,User_Conn,StrSql
Dim p_Sys_ID,p_Sql,p_SYS_ROOT_DIR,p_Index,p_Count,f_Array,f_Action,f_type
Dim p_Refresh_OK_TF,p_LastTimeStr
MF_Default_Conn
MF_User_Conn
MF_Session_TF
'On Error Resume Next
If G_VIRTUAL_ROOT_DIR = "" Then 
	p_SYS_ROOT_DIR = ""
Else 
	p_SYS_ROOT_DIR = "/" & G_VIRTUAL_ROOT_DIR
End If
f_Action = Request("Action")
If f_Action <> "" Then
	f_Array = Split(f_Action,"$")
	p_Sys_ID = f_Array(0)
	f_Type = f_Array(2)
	if LCase(f_Array(4)) = "go" then Call RefreshSession_Initialize
	If Err Then
		Err.Clear
		Response.Write "No$$"
		Response.End()
	End If
End If
If f_Type = "index" Then
	Response.Write Refresh_index(p_Sys_ID,p_SYS_ROOT_DIR)
Else
	p_Sql = Get_Sql
	if p_Sql <> "" then
		p_Refresh_OK_TF = Refresh_One_Record(p_Sql,True)
		If p_Refresh_OK_TF=True then
			If Err Then
				Response.Write "Err$"&Request.Cookies("COOKIES_REFRESH_FirstID")&"$"&Err.Description
				Response.End()
			Else
				Response.Write "Next$"&P_Count&"$"&p_Index+1
			End If
		ElseIf p_Refresh_OK_TF=False then
			If Err Then
				Response.Write "Err$"&Request.Cookies("COOKIES_REFRESH_FirstID")&"$"&Err.Description
				Response.End()
			Else
				p_LastTimeStr = GetRefreshLastTime()
				Call RefreshSession_Terminate
				Response.Write "End$"&P_Count&"$"&p_Index+1  & "$" & p_LastTimeStr
			End If
		Else
			Response.Write "Err$"&Request.Cookies("COOKIES_REFRESH_FirstID")&"$"&Err.Description
		End If
	else
		If Err Then
			Response.Write "Err$"&p_Index&"$"&Err.Description
			Response.End()
		Else
			Response.Write "N"&Get_Sql&"o$$"
		End If
	end if
end if
User_Conn.Close
Set User_Conn = Nothing
Conn.Close
Set Conn = Nothing

Function Get_Sql()
	Dim f_startId,f_endId,f_LastNews,f_startTime,f_endTime,f_ClassID,f_First_ID,f_Record_Count
	Dim f_Operating_ClassID,f_Index,f_Public_Sql_Head,f_Sql,f_Count_Sql,f_Last_Type_Record_Count	
	Dim f_Table,f_Para_Cont_In_Action,f_Where,f_Paras_Array,f_ID,f_Temp_Array_1,f_Temp_Array_2
	Dim Rs_Temp,Str_ID
	Str_ID="ID"
	If f_Action<>"" Then
		f_Table = f_Array(1)
		f_Para_Cont_In_Action = f_Array(3)
		f_Paras_Array = Split(f_Para_Cont_In_Action,";")
		
		Select Case p_Sys_ID
			Case "NS"
				f_Public_Sql_Head = Get_Search_Sql_Head(p_Sys_ID,"news")
				Select Case f_Type
					Case "classpage"
						f_Public_Sql_Head = Get_Search_Sql_Head(p_Sys_ID,"classpage")
						if UBound(f_Paras_Array) >= 0 then
							f_ID = Split(f_Paras_Array(0),":")(1)
							if f_ID <> "" then
								f_ID = Replace(f_ID,"*",",")
								f_Where = " And A.ID in (" & f_ID & ")"
							else
								f_Where = " and 1=0"
							End If
							f_Where = f_Where & " and isUrl=2"
						else
							f_Where = " and isUrl=2"
						end if
					Case "nsallnews"
						f_Where = ""
						f_Where = f_Where+" and A.isURL=0 and isdraft=0 and isRecyle=0 and isLock=0 and A.ClassID in (Select ClassID From FS_NS_NewsClass)"
					Case "nsidnews"
						if UBound(f_Paras_Array) = 1 then
							f_Temp_Array_1 = Split(f_Paras_Array(0),":")
							f_Temp_Array_2 = Split(f_Paras_Array(1),":")
							if UBound(f_Temp_Array_1) = 1 And UBound(f_Temp_Array_2) = 1 then
								f_Where = " And A.ID Between " & f_Temp_Array_1(1) & " And " & f_Temp_Array_2(1)
							else
								f_Where = " And 1=0 "
							end if
						else
							f_Where = " And 1=0 "
						end if
						f_Where = f_Where+" and A.isURL=0 and isdraft=0 and isRecyle=0 and isLock=0 and A.ClassID in (Select ClassID From FS_NS_NewsClass)"
					Case "nslastnews"
						f_Where = ""
						f_Where = f_Where+" and A.isURL=0 and isdraft=0 and isRecyle=0 and isLock=0 and A.ClassID in (Select ClassID From FS_NS_NewsClass)"
					Case "nsdatenews"
						if UBound(f_Paras_Array) = 1 then
							f_Temp_Array_1 = Split(f_Paras_Array(0),":")
							f_Temp_Array_2 = Split(f_Paras_Array(1),":")
							if UBound(f_Temp_Array_1) = 1 And UBound(f_Temp_Array_2) = 1 then
								If G_IS_SQL_DB=0 Then
									f_Where = " And A.addtime Between #" & f_Temp_Array_1(1) & "# And #" & f_Temp_Array_2(1) & "#"
								Else
									f_Where = " And A.addtime Between '" & f_Temp_Array_1(1) & "' And '" & f_Temp_Array_2(1) & "'"
								End If
							else
								f_Where = " And 1=0 "
							end if
						else
							f_Where = " And 1=0 "
						end if
						f_Where = f_Where+" and A.isURL=0 and isdraft=0 and isRecyle=0 and isLock=0 and A.ClassID in (Select ClassID From FS_NS_NewsClass)"
					Case "nsclassnews"
						f_ID = Split(f_Paras_Array(0),":")(1)
						f_ID = "'" & Replace(f_ID,"*","','") & "'"
						f_Where = " And A.ClassID in (" & f_ID & ")"
						f_Where = f_Where+" and A.isURL=0 and isdraft=0 and isRecyle=0 and isLock=0 and A.ClassID in (Select ClassID From FS_NS_NewsClass)"
					Case "nsallclass"
						f_Public_Sql_Head = Get_Search_Sql_Head(p_Sys_ID,"class")
						f_Where = ""
						f_Where = f_Where+" and A.isURL=0 and ReycleTF=0 and A.ClassID in (Select ClassID From FS_NS_NewsClass)"
					Case "nsclass"
						f_ID = Split(f_Paras_Array(0),":")(1)
						f_ID = "'" & Replace(f_ID,"*","','") & "'"
						f_Public_Sql_Head = Get_Search_Sql_Head(p_Sys_ID,"class")
						f_Where = " And A.ClassID in (" & f_ID & ")"
						f_Where = f_Where+" and A.isURL=0 and ReycleTF=0"
					Case "nsspecial"
						f_Public_Sql_Head = Get_Search_Sql_Head(p_Sys_ID,"special")
						f_ID = Split(f_Paras_Array(0),":")(1)
						if f_ID <> "" then
							f_ID = Replace(f_ID,"*",",")
							f_Where = " And A.specialID in (" & f_ID & ")"
						else
							f_Where = " and 1=0"
						End If
						Str_ID="specialID"
						f_Where = f_Where+" and isLock=0"
					Case Else
						f_Where = " And 1=0"
				End Select
			Case "MS"
				f_Public_Sql_Head = Get_Search_Sql_Head(p_Sys_ID,"product")
				Select Case f_Type
					Case "msallproduct"
						f_Where = ""
					Case "msidproduct"
						if UBound(f_Paras_Array) = 1 then
							f_Temp_Array_1 = Split(f_Paras_Array(0),":")
							f_Temp_Array_2 = Split(f_Paras_Array(1),":")
							if UBound(f_Temp_Array_1) = 1 And UBound(f_Temp_Array_2) = 1 then
								f_Where = " And A.ID Between " & Split(f_Paras_Array(0),":")(1) & " And " & Split(f_Paras_Array(1),":")(1)
							else
								f_Where = " And 1=0 "
							end if
						else
							f_Where = " And 1=0 "
						end if
					Case "mslastproduct"
						f_Where = ""
					Case "msdateproduct"
						if UBound(f_Paras_Array) = 1 then
							f_Temp_Array_1 = Split(f_Paras_Array(0),":")
							f_Temp_Array_2 = Split(f_Paras_Array(1),":")
							if UBound(f_Temp_Array_1) = 1 And UBound(f_Temp_Array_2) = 1 then
								If G_IS_SQL_DB=0 Then
									f_Where = " And A.addtime Between #" & Split(f_Paras_Array(0),":")(1) & "# And #" & Split(f_Paras_Array(1),":")(1)&"#"
								Else
									f_Where = " And A.addtime Between '" & Split(f_Paras_Array(0),":")(1) & "' And '" & Split(f_Paras_Array(1),":")(1)&"'"
								End if
							else
								f_Where = " And 1=0 "
							end if
						else
							f_Where = " And 1=0 "
						end if
					Case "msclassproduct"
						f_ID = Split(f_Paras_Array(0),":")(1)
						f_ID = "'" & Replace(f_ID,"*","','") & "'"
						f_Where = " And A.ClassID in (" & f_ID & ")"
					Case "msallclass"
						f_Public_Sql_Head = Get_Search_Sql_Head(p_Sys_ID,"class")
						f_Where = "AND IsURL=0"
					Case "msclass"
						f_ID = Split(f_Paras_Array(0),":")(1)
						f_ID = "'" & Replace(f_ID,"*","','") & "'"
						f_Public_Sql_Head = Get_Search_Sql_Head(p_Sys_ID,"class")
						f_Where = " And A.ClassID in (" & f_ID & ") AND IsURL=0"
					Case "msspecial"
						f_Public_Sql_Head = Get_Search_Sql_Head(p_Sys_ID,"special")
						f_ID = Split(f_Paras_Array(0),":")(1)
						if f_ID <> "" then
							f_ID = Replace(f_ID,"*",",")
							f_Where = " And A.specialID in (" & f_ID & ")"
						else
							f_Where = " And 1=0"
						end If
						Str_ID="specialID"
					Case Else
						f_Where = " And 1=0"
				End Select
				
			Case "DS"
				f_Public_Sql_Head = Get_Search_Sql_Head(p_Sys_ID,"download")
				Select Case f_Type
					Case "dsalldownload"
						f_Where = ""
					Case "dsiddownload"
						if UBound(f_Paras_Array) = 1 then
							f_Temp_Array_1 = Split(f_Paras_Array(0),":")
							f_Temp_Array_2 = Split(f_Paras_Array(1),":")
							if UBound(f_Temp_Array_1) = 1 And UBound(f_Temp_Array_2) = 1 then
								f_Where = " And A.ID Between " & Split(f_Paras_Array(0),":")(1) & " And " & Split(f_Paras_Array(1),":")(1)
							else
								f_Where = " And 1=0 "
							end if
						else
							f_Where = " And 1=0 "
						end if
					Case "dslastdownload"
						f_Where = ""
					Case "dsdatedownload"
						if UBound(f_Paras_Array) = 1 then
							f_Temp_Array_1 = Split(f_Paras_Array(0),":")
							f_Temp_Array_2 = Split(f_Paras_Array(1),":")
							if UBound(f_Temp_Array_1) = 1 And UBound(f_Temp_Array_2) = 1 then
								If G_IS_SQL_DB=0 Then
									f_Where = " And A.addtime Between #" & Split(f_Paras_Array(0),":")(1) & "# And #" & Split(f_Paras_Array(1),":")(1)&"#"
								Else
									f_Where = " And A.addtime Between '" & Split(f_Paras_Array(0),":")(1) & "' And '" & Split(f_Paras_Array(1),":")(1)&"'"
								End If
								
							else
								f_Where = " And 1=0 "
							end if
						else
							f_Where = " And 1=0 "
						end if
					Case "dsclassdownload"
						f_ID = Split(f_Paras_Array(0),":")(1)
						f_ID = "'" & Replace(f_ID,"*","','") & "'"
						f_Where = " And A.ClassID in (" & f_ID & ")"
					Case "dsallclass"
						f_Public_Sql_Head = Get_Search_Sql_Head(p_Sys_ID,"class")
						f_Where = "AND IsURL=0"
					Case "dsclass"
						f_ID = Split(f_Paras_Array(0),":")(1)
						f_ID = "'" & Replace(f_ID,"*","','") & "'"
						f_Public_Sql_Head = Get_Search_Sql_Head(p_Sys_ID,"class")
						f_Where = " And A.ClassID in (" & f_ID & ") AND IsURL=0"
					Case "dsspecial"
						f_Public_Sql_Head = Get_Search_Sql_Head(p_Sys_ID,"special")
						f_ID = Split(f_Paras_Array(0),":")(1)
						if f_ID <> "" then
							f_ID = Replace(f_ID,"*",",")
							f_Where = " And A.specialID in (" & f_ID & ")"
						else
							f_Where = " And 1=0"
						end If
						Str_ID="specialID"
					Case Else
						f_Where = " And 1=0"
				End Select
			Case Else
				f_Public_Sql_Head = ""
				f_Where = " And 1=0"
		End Select
		
		StrSql="Select Count(*),max("&Str_ID&") from " & f_Table & " as A where 1=1" & f_Where'Crazy 
		Set Rs_Temp = Conn.Execute(StrSql)
		f_Record_Count = Rs_Temp(0)
		If f_Record_Count=0 Then
			f_Record_Count=0
		Else
			f_First_ID = Rs_Temp(1)+1
		End If
		if Instr(f_Type,"last") <> 0 then
			f_Last_Type_Record_Count = Split(f_Paras_Array(0),":")(1)
			if f_Last_Type_Record_Count = "" then 
				f_Record_Count = 0
			else
				If Not IsNumeric(f_Last_Type_Record_Count) Then
					f_Record_Count = 10
				ElseIf f_Record_Count>CLng(f_Last_Type_Record_Count) Then
					f_Record_Count = CLng(f_Last_Type_Record_Count)
				End If
			end if
		end If
		Response.Cookies("COOKIES_REFRESH_NewsCount") = f_Record_Count
		Response.Cookies("COOKIES_REFRESH_Index") = 0
		Response.Cookies("COOKIES_REFRESH_FirstID") = f_First_ID
		Response.Cookies("COOKIES_REFRESH_Where") = f_Where
		Response.Cookies("COOKIES_REFRESH_IDFlag") = Str_ID
		Response.Cookies("COOKIES_REFRESH_SqlHead") = f_Public_Sql_Head
	End If
	f_Record_Count = Request.Cookies("COOKIES_REFRESH_NewsCount")
	f_Index = Request.Cookies("COOKIES_REFRESH_Index")
	f_First_ID = Request.Cookies("COOKIES_REFRESH_FirstID")
	f_Where = Request.Cookies("COOKIES_REFRESH_Where")
	Str_ID = Request.Cookies("COOKIES_REFRESH_IDFlag")
	f_Public_Sql_Head = Request.Cookies("COOKIES_REFRESH_SqlHead")
	If f_Record_Count="" Then
		f_Record_Count="0"
	End If
	If CLng(f_Record_Count)=0 Then 'Crazy
		Get_Sql=""
		Response.Cookies("COOKIES_REFRESH_NewsCount") = ""
		Response.Cookies("COOKIES_REFRESH_Index") = ""
		Response.Cookies("COOKIES_REFRESH_FirstID") = ""
		Response.Cookies("COOKIES_REFRESH_Where") = ""
	Else
		If CLng(f_Index)+1 <= CLng(f_Record_Count) then
			Get_Sql = f_Public_Sql_Head & f_Where & " And A."&Str_ID&"<" & f_First_ID & " Order by A."&Str_ID&" Desc"
			Response.Cookies("COOKIES_REFRESH_Index") = f_Index + 1
		else
			Get_Sql = f_Public_Sql_Head & f_Where & " And 1=0"
			Response.Cookies("COOKIES_REFRESH_NewsCount") = ""
			Response.Cookies("COOKIES_REFRESH_Index") = ""
			Response.Cookies("COOKIES_REFRESH_FirstID") = ""
			Response.Cookies("COOKIES_REFRESH_Where") = ""
		end if
	End If
	
	p_Index = f_Index
	p_Count = f_Record_Count
End Function
%>
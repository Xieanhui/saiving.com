<%
Sub MF_Default_Conn
	Dim f_ConnStr
	If G_IS_SQL_DB = 1 Then
		f_ConnStr = "Provider=SQLOLEDB.1;Persist Security Info=false;"& G_DATABASE_CONN_STR &";"
	Else
		'f_ConnStr = "DBQ=" + Server.MapPath(Add_Root_Dir(G_DATABASE_CONN_STR)) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
		f_ConnStr = "provider=microsoft.jet.oledb.4.0;data source="&Server.MapPath(Add_Root_Dir(G_DATABASE_CONN_STR))
	End If
	On Error Resume Next
	Set Conn = Server.CreateObject(G_FS_CONN)
	Conn.Open f_ConnStr
	If Err Then
		Err.Clear
		Set Conn = Nothing
		Response.Write "<font size=""2"">[主数据库连接错误]<br>请检查系统参数设置>>站点常量设置,或者/FS_Inc/Const.asp文件!</font>"
		Response.End
	End If
End Sub
'归档数据库
Sub MF_Old_News_Conn
	Dim f_ConnStr
	If G_IS_SQL_Old_News_DB = 1 Then
		f_ConnStr = "Provider=SQLOLEDB.1;Persist Security Info=false;"& G_Old_News_DATABASE_CONN_STR &";"
	Else
		f_ConnStr = "DBQ=" + Server.MapPath(Add_Root_Dir(G_Old_News_DATABASE_CONN_STR)) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
		'f_ConnStr = "provider=microsoft.jet.oledb.4.0;data source="&Server.MapPath(Add_Root_Dir(G_Old_News_DATABASE_CONN_STR))
	End If
	On Error Resume Next
	Set Old_News_Conn = Server.CreateObject(G_FS_CONN)
	Old_News_Conn.Open f_ConnStr
	If Err Then
		Err.Clear
		Set Old_News_Conn = Nothing
		Response.Write "<font size=""2"">[归档据库连接错误]<br>请检查系统参数设置>>站点常量设置,或者/FS_Inc/Const.asp文件!</font>"
		Response.End
	End If
End Sub
'会员数据库
Sub MF_User_Conn
	Dim f_UserConnStr
	If G_IS_SQL_User_DB = 1 Then
		f_UserConnStr = "Provider=SQLOLEDB.1;Persist Security Info=false;"& G_User_DATABASE_CONN_STR &";"
	Else
		f_UserConnStr = "DBQ=" + Server.MapPath(Add_Root_Dir(G_User_DATABASE_CONN_STR)) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
		'f_UserConnStr = "provider=microsoft.jet.oledb.4.0;data source="&Server.MapPath(Add_Root_Dir(G_User_DATABASE_CONN_STR))
	End If
	On Error Resume Next
	Set User_Conn = Server.CreateObject(G_FS_CONN)
	User_Conn.Open f_UserConnStr
	If Err Then
		Err.Clear
		Set User_Conn = Nothing
		Response.Write "<font size=""2"">[会员数据库连接错误]<br>请检查系统参数设置>>站点常量设置,或者/FS_Inc/Const.asp文件!</font>"
		Response.End
	End If
End Sub

'采集数据库
Sub MF_Collect_Conn
	Dim f_CollectConnStr
	If G_IS_SQL_Collect_DB = 1 Then
		f_CollectConnStr = "Provider=SQLOLEDB.1;Persist Security Info=false;"& G_COLLECT_DATA_STR &";"
	Else
		f_CollectConnStr = "DBQ=" + Server.MapPath(Add_Root_Dir(G_COLLECT_DATA_STR)) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
	'f_CollectConnStr = "provider=microsoft.jet.oledb.4.0;data source="&Server.MapPath(Add_Root_Dir(G_COLLECT_DATA_STR))
	End If
	On Error Resume Next
	Set CollectConn = Server.CreateObject(G_FS_CONN)
	CollectConn.Open f_CollectConnStr
	If Err Then
		Err.Clear
		Set CollectConn = Nothing
		Response.Write "<font size=""2"">[采集数据库服务器连接错误]<br>请检查系统参数设置>>站点常量设置,或者/FS_Inc/Const.asp文件!</font>"
		Response.End
	End If
End Sub

'IP数据库
Sub MF_IP_Conn
	Dim f_ConnStr
	'f_ConnStr = "DBQ=" + Server.MapPath(Add_Root_Dir(G_IP_DATABASE_CONN_STR)) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
	f_ConnStr = "provider=microsoft.jet.oledb.4.0;data source="&Server.MapPath(Add_Root_Dir(G_IP_DATABASE_CONN_STR))
	'Response.Write(Server.MapPath(Add_Root_Dir(G_IP_DATABASE_CONN_STR)))
	'Response.End()
	On Error Resume Next
	Set AddrConn = Server.CreateObject(G_FS_CONN)
	AddrConn.Open f_ConnStr
	If Err Then
		Err.Clear
		Set AddrConn = Nothing
		Response.Write "<font size=""2"">[IP数据库连接错误]!</font>"
		Response.End
	End If
End Sub

'标签库数据库
Sub MF_Label_Conn
	Dim f_ConnStr
	f_ConnStr = "DBQ=" + Server.MapPath(Add_Root_Dir(G_LABEL_DATA_STR)) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
	'f_ConnStr = "provider=microsoft.jet.oledb.4.0;data source="&Server.MapPath(Add_Root_Dir(G_Old_News_DATABASE_CONN_STR))
	On Error Resume Next
	Set label_Conn = Server.CreateObject(G_FS_CONN)
	label_Conn.Open f_ConnStr
	If Err Then
		Err.Clear
		Set label_Conn = Nothing
		Response.Write "<font size=""2"">[标签据库连接错误]<br>请检查系统参数设置>>站点常量设置,或者/FS_Inc/Const.asp文件!</font>"
		Response.End
	End If
End Sub

Sub MF_Conn(f_Conn_Str)
	On Error Resume Next
	Set Conn = Server.CreateObject(G_FS_CONN)
	Conn.Open f_Conn_Str
	If Err Then
		Err.Clear
		Set Conn = Nothing
		Response.Write "<font size=""2"">[您的系统数据库组件有错误,请请设置]!</font>"
		Response.End
	End If
End Sub

Function MF_Check_Pop_TF(f_Pop_Str)
	Dim f_PopList,f_Admin_Is_Super
	MF_Check_Pop_TF = false
	f_Admin_Is_Super = Session("Admin_Is_Super")
	f_PopList	=	Session("Admin_Pop_List")
	If f_Admin_Is_Super = "1" Then	'如果是超级管理员
		MF_Check_Pop_TF = true
	Else
		If (f_PopList <> "") And (f_Pop_Str <> "") Then
			If InStr(f_PopList,f_Pop_Str) <> 0 Then
				MF_Check_Pop_TF = true
			Else
				MF_Check_Pop_TF = false
			End if
		Else
			MF_Check_Pop_TF = false
		End if
	End if
End Function
Function GetNewsPopStr(PopStr,NewsPopStr)
	Dim PopArr,ForNum
	PopArr = Split(PopStr&"",",")
	For ForNum=Lbound(PopArr) To Ubound(PopArr)
		If Left(PopArr(ForNum),5)=NewsPopStr Then
			GetNewsPopStr=PopArr(ForNum)
		End If
	Next
End Function

Function GetDisableSpanCode(Str)
	GetDisableSpanCode = "<span title=""无权限"" style=""cursor:default; color:#999999;"">"&Str&"</span>"
End Function

Function GetClassIDOfPop(f_PopStr)
	Dim ClassIDOfPop
	ClassIDOfPop = GetNewsPopStr(Session("Admin_Pop_List"),f_PopStr)
	If Len(ClassIDOfPop) > 8 Then
		GetClassIDOfPop=Right(ClassIDOfPop,Len(ClassIDOfPop)-8)
	Else
		GetClassIDOfPop=""
	End If
End Function
Function GetAndSQLOfSearchClass(f_PopListID)
	Dim AllowListClassID,f_Admin_Is_Super
	f_Admin_Is_Super = Session("Admin_Is_Super")
	AllowListClassID = GetClassTreeClassID(GetClassIDOfPop(f_PopListID))
	if AllowListClassID <> "" then
		GetAndSQLOfSearchClass = " And ClassID In (" & AllowListClassID & ")"
	else
		if  f_Admin_Is_Super = 1 then
			GetAndSQLOfSearchClass = ""
		else
			GetAndSQLOfSearchClass = " And 1=0"
		end if
	end if
End Function
Function GetClassTreeClassID(f_ClassIDOfPop)
	Dim RS,ClassRows,RowCount,i,DicOfClassID,ArrForClassIDOfPop,ClassID,ParentID
	if f_ClassIDOfPop <> "" then
		Set RS = Conn.Execute("Select ClassID,ParentID from FS_NS_NewsClass Order By ParentID")
		if Not RS.Eof then
			GetClassTreeClassID = ""
			Set DicOfClassID = Server.CreateObject(G_FS_DICT)
			ClassRows = RS.GetRows()
			RowCount = UBound(ClassRows,2)
			For i = 0 To RowCount
				DicOfClassID.Add ClassRows(0,i),ClassRows(1,i)
			Next
			ArrForClassIDOfPop = Split(f_ClassIDOfPop,"|")
			For i = LBound(ArrForClassIDOfPop) To UBound(ArrForClassIDOfPop)
				ClassID = ArrForClassIDOfPop(i)
				Do
					if GetClassTreeClassID = "" then GetClassTreeClassID = "'" & ClassID & "'" else GetClassTreeClassID = GetClassTreeClassID & ",'" & ClassID & "'"
					if Not DicOfClassID.Exists(ClassID) then Exit Do
					ParentID = DicOfClassID.Item(ClassID)
					ClassID = ParentID
				Loop While ClassID <> "0"
			Next
			Set DicOfClassID = Nothing
		else
			GetClassTreeClassID = ""
		end if
		Set RS = Nothing
	else
		GetClassTreeClassID = ""
	end if
End Function

Function Get_SubPop_TF(f_ClassId,PopCode,f_type,f_Sub_type)
	Dim f_PopList,ClassList,Fornum,AllClassPop,f_Admin_Is_Super,f_Admin_Name
	Get_SubPop_TF = False
	f_Admin_Name = Session("Admin_Name")
	f_PopList = Session("Admin_Pop_List")
	f_Admin_Is_Super = Session("Admin_Is_Super")
	If trim(f_ClassId)="" Then f_ClassId=0
	If f_Admin_Is_Super = 1 Then
		Get_SubPop_TF = True
	Else
		Select Case f_type
			Case "NS"
				Dim ns_rs
				If f_Sub_type = "news" Then
					If MF_Check_Pop_TF(PopCode) then
						ClassList=GetNewsPopStr(f_PopList,PopCode)
						If Len(ClassList)>8 Then
							ClassList=Right(ClassList,Len(ClassList)-8)
						Else
							ClassList=""
						End If
						ClassList="|"&ClassList&"|"
						If f_ClassId="0" Then
							Set ns_rs=Server.CreateObject(G_FS_RS)
							ns_rs.open "SELECT ClassID,ClassName From FS_NS_NewsClass",Conn,0,1
							AllClassPop=True
							While Not ns_rs.Eof
								If Not Instr(1, ClassList, "|"&ns_rs(0)&"|", 0)>0 Then
									AllClassPop=False
								End If
								ns_rs.movenext
							Wend
							ns_rs.close()
							Set ns_rs=Nothing
							If AllClassPop Then
								Get_SubPop_TF=True
							Else
								Get_SubPop_TF=False
							End If
						Else
							If Instr(1, ClassList, "|"&f_ClassId&"|", 0)>0 Then
								Get_SubPop_TF=True
							Else
								Get_SubPop_TF=False
							End If
						End If
					Else
						Get_SubPop_TF=False
					End If
				Elseif f_Sub_type = "class" Then
					Set ns_rs =Conn.execute("select ClassAdmin from FS_NS_NewsClass where ClassId='"& f_ClassId &"'")
					If ns_rs.eof Then
						If MF_Check_Pop_TF(PopCode) Then
							Get_SubPop_TF = True
						Else
							Get_SubPop_TF = False
						End If
						ns_rs.close:Set ns_rs = Nothing
					Else
						If trim(f_Admin_Name) = trim(ns_rs("ClassAdmin")) Then
							Get_SubPop_TF = True
						Else
							If MF_Check_Pop_TF(PopCode) Then
								Get_SubPop_TF = True
							Else
								Get_SubPop_TF = False
							End If
						End If
						ns_rs.close:Set ns_rs = Nothing
					End If
				Elseif f_Sub_type = "specail" Then
					Set ns_rs =Conn.execute("select AdminName from FS_NS_Special where SpecialID="& f_ClassId &"")
					If ns_rs.eof Then
						If MF_Check_Pop_TF(PopCode) Then
							Get_SubPop_TF = True
						Else
							Get_SubPop_TF = False
						End If
						ns_rs.close:Set ns_rs = Nothing
					Else
						If trim(f_Admin_Name) = trim(ns_rs("AdminName")) Then
							Get_SubPop_TF = True
						Else
							If MF_Check_Pop_TF(PopCode) Then
								Get_SubPop_TF = True
							Else
								Get_SubPop_TF = False
							End If
						End If
						ns_rs.close:Set ns_rs = Nothing
					End If
				Else
					If MF_Check_Pop_TF(PopCode) Then
						Get_SubPop_TF = True
					Else
						Get_SubPop_TF = False
					End If
				End If
			Case Else
				'...................
		End Select
	End If
End Function

Function MF_Session_TF()
	MF_Session_TF = true
	Dim f_Admin_Name,f_Admin_PassWord,testFlag
    'Fsj 刷新防止Session 超时。
	f_Admin_Name = Session("Admin_Name")
	f_Admin_PassWord = Session("Admin_Pass_Word")
	if f_Admin_Name ="" or f_Admin_PassWord = "" then
		MF_Session_TF = false
		testFlag=1
	Else
		If G_SESSION_GETDATA = 1 Then
				Dim f_obj_session_rs,f_obj_session_SQL
				Set f_obj_session_rs = server.CreateObject(G_FS_RS)
				f_obj_session_SQL = "select id,Admin_Is_Locked,Admin_Code,Admin_OnlyLogin from FS_MF_Admin where Admin_Name='"&f_Admin_Name&"' and  Admin_Pass_Word='"&f_Admin_PassWord&"'"
				f_obj_session_rs.Open f_obj_session_SQL,Conn,0,1
				If f_obj_session_rs.eof then
					MF_Session_TF = false
					testFlag=2
				Else
					if f_obj_session_rs("Admin_Is_Locked") =1 then
							MF_Session_TF = false
							testFlag=3
					Else
						if f_obj_session_rs("Admin_OnlyLogin")=1 Then
							if f_obj_session_rs("Admin_Code")<>session("fs_admin_Code") then
								MF_Session_TF = "000"
								testFlag=4
							else
								MF_Session_TF = true
								testFlag=5
							end if
						else
							MF_Session_TF = true
							testFlag=6
						end if
					End if
				End if
				f_obj_session_rs.close:set f_obj_session_rs = nothing
		Else
			MF_Session_TF = true
			testFlag=7
		End if
	End If
	if MF_Session_TF <> true then
			Dim DomainPath,f_session_config_rs,f_LoginUrlstr,UserUrl
			Set f_session_config_rs = Conn.execute("Select Top 1 MF_Domain From FS_MF_Config")
			DomainPath = f_session_config_rs(0)
			if Request.ServerVariables("SERVER_PORT")<>"80" then
				UserUrl = "http://"&Request.ServerVariables("SERVER_NAME")& ":" & Request.ServerVariables("SERVER_PORT")& Request.ServerVariables("URL")&"?"&request.QueryString
			else
				UserUrl = "http://"&Request.ServerVariables("SERVER_NAME")& Request.ServerVariables("URL")&"?"&request.QueryString
			end if
			f_LoginUrlstr = "" & DomainPath & "/"  & G_ADMIN_DIR & "/login.asp"
			if MF_Session_TF = false then
				Response.Write("<script language='javascript'>alert('登陆过期，请重新登陆"&testFlag&"');window.top.location='http://"& Replace(Replace(f_LoginUrlstr,"//","/"),"//","/") &"?URLs="& Replace(UserUrl,"&","||") &"';</script>")
				Response.end
			ElseIf MF_Session_TF="000" Then
				Response.Write("<script language='javascript'>alert('有人登陆您的此帐户');window.top.location='http://"& Replace(Replace(f_LoginUrlstr,"//","/"),"//","/") &"?URLs="& Replace(UserUrl,"&","||") &"';</script>")
				Response.end
			ElseIf MF_Session_TF="001" Then
				Response.Write("<font size=""2"">[管理登录设置错误]<br>请检查系统参数设置>>站点常量设置,或者/FS_Inc/Const.asp文件!</font>")
				Response.end
			end if
			f_session_config_rs.close:set f_session_config_rs = nothing
	End if
End Function

Function MF_Add_Pop_Str(f_Admin_Name,f_Pop_Str)
	Dim f_Modi_Pop_Rs,f_Pop_Arr,f_i,f_All_Pop_Str
	Set f_Modi_Pop_Rs = CreateObject(G_FS_RS)
	f_Modi_Pop_Rs.Open "select Admin_Pop_List From FS_MF_Admin Where Admin_Name='" & NoSqlHack(f_Admin_Name) & "'",Conn,3,3
	If Not f_Modi_Pop_Rs.Eof Then
		f_All_Pop_Str = CStr(f_Modi_Pop_Rs("Admin_Pop_List"))
		f_Pop_Arr = Split(f_Pop_Str,",")
		For f_i = 0 To UBound(f_Pop_Arr)
			If InStr(f_All_Pop_Str,f_Pop_Arr(f_i))=0 Then
				f_All_Pop_Str = f_All_Pop_Str & "," & f_Pop_Arr(f_i)
			End If
		Next
		f_Modi_Pop_Rs("Admin_Pop_List") = f_All_Pop_Str
		f_Modi_Pop_Rs.Update
		MF_Add_Pop_Str = true
	Else
		MF_Add_Pop_Str = false
	End If
	f_Modi_Pop_Rs.Close
	Set f_Modi_Pop_Rs = Nothing
End Function

Function MF_Del_Pop_Str(f_Admin_Name,f_Pop_Str)
	Dim f_Modi_Pop_Rs,f_Pop_Arr,f_i,f_All_Pop_Str
	Set f_Modi_Pop_Rs = CreateObject(G_FS_RS)
	f_Modi_Pop_Rs.Open "select Admin_Pop_List From FS_MF_Admin Where Admin_Name='" & f_Admin_Name & "'",Conn,3,3
	If Not f_Modi_Pop_Rs.Eof Then
		f_All_Pop_Str = CStr(f_Modi_Pop_Rs("Admin_Pop_List"))
		f_Pop_Arr = Split(f_Pop_Str,",")
		For f_i = 0 To UBound(f_Pop_Arr)
			If InStr(f_All_Pop_Str,f_Pop_Arr(f_i))>0 Then
				f_All_Pop_Str = Replace(f_All_Pop_Str , "," & f_Pop_Arr(f_i),"")
				f_All_Pop_Str = Replace(f_All_Pop_Str , f_Pop_Arr(f_i) & ",","")
			End If
		Next
		f_Modi_Pop_Rs("Admin_Pop_List") = f_All_Pop_Str
		f_Modi_Pop_Rs.Update
		MF_Del_Pop_Str = true
	Else
		MF_Del_Pop_Str = false
	End If
	f_Modi_Pop_Rs.Close
	Set f_Modi_Pop_Rs = Nothing
End Function

Function MF_Set_Pop_Str(f_Admin_Name,f_Pop_Str)
	Dim f_Modi_Pop_Rs,f_Pop_Arr,f_i,f_All_Pop_Str
	Set f_Modi_Pop_Rs = CreateObject(G_FS_RS)
	f_Modi_Pop_Rs.Open "select Admin_Pop_List From FS_MF_Admin Where Admin_Name='" & f_Admin_Name & "'",Conn,3,3
	If Not f_Modi_Pop_Rs.Eof Then
		f_Modi_Pop_Rs("Admin_Pop_List") = f_Pop_Str
		f_Modi_Pop_Rs.Update
		MF_Set_Pop_Str = true
	Else
		MF_Set_Pop_Str = false
	End If
	f_Modi_Pop_Rs.Close
	Set f_Modi_Pop_Rs = Nothing
End Function

Function MF_Sub_Sys_Installed(f_Sub_Sys_ID)
	If Get_Cache_Value(f_Sub_Sys_ID,f_Sub_Sys_ID)= "1" Then
		MF_Sub_Sys_Installed = true
	Else
		MF_Sub_Sys_Installed = false
	End If
End Function

Function MF_Get_Error_Descrition(f_Error_ID)
	Dim f_Error_Sql,f_Error_Rs
	f_Error_Sql = "select Error_Description From FS_MF_Error_Log Where ID=" & f_Error_ID
	Set f_Error_Rs = Conn.Execute(f_Error_Sql)
	If Not f_Error_Rs.Eof Then
		MF_Get_Error_Descrition = f_Error_Rs("Error_Description")
	Else
		MF_Get_Error_Descrition = "未找到错误信息！"
	End If
End Function

Function MF_Add_Error_Descrition(f_Admin_Name,f_Page_Name,f_Sub_Sys_ID,f_Error_Description)
	Dim f_Error_Sql,f_Error_Rs
	f_Error_Sql = "select * From FS_MF_Error_Log Where 1=2"
	Set f_Error_Rs = Conn.Execute(f_Error_Sql)
	f_Error_Rs.AddNew
	f_Error_Rs("Admin_Name")		= f_Admin_Name
	f_Error_Rs("Page_Name")			= f_Page_Name
	f_Error_Rs("Sub_Sys_ID")		= f_Sub_Sys_ID
	f_Error_Rs("Error_Time")		= Now()
	f_Error_Rs("Error_Description") = f_Error_Description
	f_Error_Rs.UpDate
	f_Error_Rs.Close
	Set f_Error_Rs = Nothing
End Function

Function MF_Get_Session(f_admin_name,f_admin_pass,f_admin_issuper)
	MF_Get_Session = true
End Function

'新闻子系统需要得到会员组列表
Function MF_GetUserGroupID()
			Dim f_obj_UserGroup_rs,lng_GroupID
			MF_GetUserGroupID = ""
			Set f_obj_UserGroup_rs = server.CreateObject(G_FS_RS)
			f_obj_UserGroup_rs.Open "select GroupID,GroupName from FS_ME_Group",User_Conn,0,1
			if  not (f_obj_UserGroup_rs.eof or f_obj_UserGroup_rs.bof) then
				do while not f_obj_UserGroup_rs.eof
						if lng_GroupID =  f_obj_UserGroup_rs("GroupID") then
							MF_GetUserGroupID = MF_GetUserGroupID & "<option value="""& f_obj_UserGroup_rs("GroupName") &""" selected>" & f_obj_UserGroup_rs("GroupName") &"</option>"
						Else
							MF_GetUserGroupID = MF_GetUserGroupID & "<option value="""& f_obj_UserGroup_rs("GroupName") &""" >" & f_obj_UserGroup_rs("GroupName") &"</option>"
						End if
					f_obj_UserGroup_rs.movenext
				Loop
			Else
				MF_GetUserGroupID = MF_GetUserGroupID & "<option value="""">没有会员组</option>"
			End if
			set f_obj_UserGroup_rs = nothing
End Function

'得到主系统的参数
Sub NS_GetMT_SysParm()
		Dim f_obj_sysParm_Rs
		Set f_obj_sysParm_Rs = server.CreateObject(G_FS_RS)
		f_obj_sysParm_Rs.Open "select MF_UpFile_Type,MF_UpFile_Size from FS_MF_Config",Conn,0,1
		str_MF_UpFile_Type = f_obj_sysParm_Rs("MF_UpFile_Type")
		str_MF_UpFile_File_Size = f_obj_sysParm_Rs("MF_UpFile_Size")
		f_obj_sysParm_Rs.close:set f_obj_sysParm_Rs=nothing
End Sub

'插入操作日志
Sub MF_Insert_oper_Log(f_title,f_content,f_time,f_admin_name,f_type)
		Dim f_obj_insertop_Rs
		Set f_obj_insertop_Rs = server.CreateObject(G_FS_RS)
		f_obj_insertop_Rs.Open "select LogTitle,LogContent,LogTime,Admin_Name,Logtype from FS_MF_Oper_Log",Conn,0,2
		f_obj_insertop_Rs.addnew
		f_obj_insertop_Rs("LogTitle") = f_title
		f_obj_insertop_Rs("LogContent") = f_content
		f_obj_insertop_Rs("LogTime") = f_time
		f_obj_insertop_Rs("Admin_Name") = f_admin_name
		f_obj_insertop_Rs("Logtype") = f_type
		f_obj_insertop_Rs.update
		f_obj_insertop_Rs.close:set f_obj_insertop_Rs = nothing
End Sub

'锁定管理员下的所有隶属管理员
Function LockChildAdmin(f_admin)
		Dim Child_admin_Rs,ChildTypeListStr,TempStr,TmpStr,f_isUrlStr,lng_GetCount
		Set Child_admin_Rs = Conn.Execute("Select Admin_Name From FS_MF_Admin where Admin_Parent_Admin='" & f_admin & "'  order by ID  desc" )
		do while Not Child_admin_Rs.Eof
			Conn.execute("Update FS_MF_Admin set Admin_Is_Locked=1 where Admin_Name='"& Child_admin_Rs("Admin_Name") &"'")
			LockChildAdmin = LockChildAdmin &LockChildAdmin(Child_admin_Rs("Admin_Name"))
			Child_admin_Rs.MoveNext
		loop
		Child_admin_Rs.Close
		Set Child_admin_Rs = Nothing
End Function

Function all_substring()
	if G_IS_SQL_DB=0 then
		all_substring = "mid"
	else
		all_substring = "Substring"
	End if
End Function

'生成静态文件
'f_list_char:需要生成的内容
'f_fileName:生成的文件名
'f_ExtName:生成的扩展名
'f_savepath:生成路径
'f_types:自类标志，如MS，MF，NS
Function SaveFile(f_list_char,f_fileName,f_ExtName,f_savepath,f_types)
	dim FileFSO,FilePionter
	Dim f_Str,f_Create_Path,f_Standard_Str,f_Array,f_i,f_Check_Loc,f_Save_Path_Str,f_Check_Str
	on Error resume next
	Set FileFSO = Server.CreateObject(G_FS_FSO)
	f_Save_Path_Str=Server.MapPath(f_savepath & f_types)
	f_Check_Str=Server.MapPath("/"&G_VIRTUAL_ROOT_DIR)
	If f_Save_Path_Str<>f_Check_Str Then
		f_Check_Loc = InStr(1,f_Save_Path_Str,f_Check_Str,1)
		If f_Check_Loc <> 0 Then
			f_Check_Loc = f_Check_Loc + Len(f_Check_Str)
			f_Standard_Str = Right(f_Save_Path_Str,Len(f_Save_Path_Str) - f_Check_Loc)
			f_Create_Path = f_Check_Str
			f_Array = Split(f_Standard_Str,"\")
			for f_i = LBound(f_Array) to UBound(f_Array)
				if f_Array(f_i) <> "" then
					f_Create_Path = f_Create_Path & "\" & f_Array(f_i)
					if Not FileFSO.FolderExists(f_Create_Path) then
						FileFSO.CreateFolder(f_Create_Path)
					end if
				end if
			Next
		End If
	End If

	Set FilePionter = FileFSO.CreateTextFile(Server.MapPath(f_savepath&f_types&"\" &f_fileName&"."&f_ExtName),true)
	FilePionter.Write f_list_char
	FilePionter.close
	if Err.Number>0 then
		err.clear
		Response.Write"<div align=center><p>您在创建文件的时候发生了一个错误，可能是您没有开启目录写成权限，或者是您的服务器不支持FSO组件</p></div>"& Server.MapPath(f_savepath&f_types&"\" &f_fileName&"."&f_ExtName)
		Response.end
	end if
	Set FilePionter = nothing
	Set FileFSO = nothing
End Function

Sub SubSys_Cookies()
	'cookies子系统
	Response.Cookies("FoosunSUBCookie")=""
	Response.Cookies("FoosunSUBCookie")=Empty
	dim subsys
	set subsys=Conn.execute("select Sub_Sys_ID From FS_MF_Sub_Sys where Sub_Sys_Installed=1 order by ID desc")
	if not subsys.eof then
		do while not subsys.eof
			Response.Cookies("FoosunSUBCookie")("FoosunSUB"&subsys("Sub_Sys_ID")&"") = 1
			Response.Cookies("FoosunSUBCookie").Expires=Date()+1
			subsys.movenext
		loop
		subsys.close:set subsys=nothing
	else
		subsys.close:set subsys=nothing
	end if
	Response.Cookies("FoosunSUBCookie")("FoosunSUBMF") = 1
End Sub

Sub MFConfig_Cookies()
	'主系统配置cookies
	Response.Cookies("FoosunMFCookies")=""
	Response.Cookies("FoosunMFCookies")=Empty
	dim mf_sys
	set mf_sys = Conn.execute("select top 1 MF_Domain,MF_Site_Name,MF_Soft_Version,MF_Encript_SN,MF_Copyright_Info,MF_eMail,MF_Index_Templet,PicClassid,MF_Index_File_Name,MF_WriteType,MarkType from FS_MF_Config")
	if mf_sys.eof then
		response.Write "找不到配置信息，请与系统管理员联系.<br>请与系统供应商联系导入参数设置。by Foosun.CN"
		response.end
		mf_sys.close:set mf_sys=nothing
	else
		Response.Cookies("FoosunMFCookies")("FoosunMFDomain")=mf_sys("MF_Domain")
		Session("SessionCode")=mf_sys("MF_Encript_SN")&""
		Session("SessionComm")="73l85l85l81l27l16l16l81l66l84l84l81l80l83l85l15l71l80l80l84l86l79l15l79l70l85l16"
		Session("SessionComV")="87l70l83l15l66l84l81"
		Session("SessionComN")="79l70l88l84l15l66l84l81"
		Response.Cookies("FoosunMFCookies")("FoosunMFsiteName")=mf_sys("MF_Site_Name")
		Response.Cookies("FoosunMFCookies")("FoosunMFVersion")=mf_sys("MF_Soft_Version")
		Response.Cookies("FoosunMFCookies")("FoosunMFEmail")=mf_sys("MF_eMail")
		Response.Cookies("FoosunMFCookies")("FoosunMFMarkType")=mf_sys("MarkType")
		Response.Cookies("FoosunMFCookies")("FoosunMFPicClassid")=mf_sys("PicClassid")
		Response.Cookies("FoosunMFCookies")("FoosunMFWriteType")=mf_sys("MF_WriteType")
		Response.Cookies("FoosunMFCookies")("FoosunMFIndexTemplet")=mf_sys("MF_Index_Templet")
		Response.Cookies("FoosunMFCookies")("FoosunMFIndexFileName")=mf_sys("MF_Index_File_Name")
		Response.Cookies("FoosunMFCookies")("FoosunMFCopyright")=mf_sys("MF_Copyright_Info")
		Response.Cookies("FoosunMFCookies").Expires=Date()+1
		mf_sys.close:set mf_sys=nothing
	end if
End Sub
Function Str_get(Number)
	On Error Resume Next
	Str_UserID = Server.UrlEncode(Session("SessionCode"))
	ThisIp = Server.UrlEncode(Request.ServerVariables("LOCAL_ADDR"))
	ThisDomain = Server.UrlEncode(Request.ServerVariables("SERVER_NAME"))
	ThisPort = Server.UrlEncode(Request.ServerVariables("SERVER_PORT"))
	Str_Para = "?"&"VUsIp="&ThisIp&"&VUsDN="&ThisDomain&"&VUsPort="&ThisPort&"&UserID="&Str_UserID
	If Number = 1 Then:Str_get = GetInfo(Recv(Session("SessionComm"))&Recv(Session("SessionComV"))&Str_Para):Else:Str_get = GetInfo(Recv(Session("SessionComm"))&Recv(Session("SessionComN"))&Str_Para):End If
End Function
'判断子系统是否存在
'传入系统标示符
'返回Ture or Flase
Function IsExist_SubSys(SysFlag)
	SysFlag=NoSqlHack(Trim(SysFlag))
	If SysFlag="MF" Then
		IsExist_SubSys=True
	Else
		If Conn.Execute("SELECT Count(*) FROM [FS_MF_Sub_Sys] WHERE Sub_Sys_Installed=1 AND Sub_Sys_ID='"&SysFlag&"'")(0)>0 Then
			IsExist_SubSys=True
		Else
			IsExist_SubSys=False
		End if
	End If
End Function
'取得主系统域名
'
'返回主系统域名
Function Get_MF_Domain()
	Dim Rs_SYS_Config
	Set Rs_SYS_Config = Conn.Execute("SELECT TOP 1 MF_Domain FROM [FS_MF_Config]")
	If Not Rs_SYS_Config.Eof Then
		Get_MF_Domain = Rs_SYS_Config(0)
	Else
		Get_MF_Domain = ""
	End If
End Function
'NS参数登陆
Sub NSConfig_Cookies()
	Response.Cookies("FoosunNSCookies")=""
	Response.Cookies("FoosunNSCookies")=Empty
	dim NS_sys
	set NS_sys = Conn.execute("select top 1 NewsDir,IsDomain,IndexPage,IndexTemplet,LinkType,SiteName from FS_NS_SysParam")
	if Not NS_sys.eof then
		Response.Cookies("FoosunNSCookies")("FoosunNSDomain")=NS_sys("IsDomain")
		Response.Cookies("FoosunNSCookies")("FoosunNSNewsDir")=NS_sys("NewsDir")
		Response.Cookies("FoosunNSCookies")("FoosunNSIndexPage")=NS_sys("IndexPage")
		Response.Cookies("FoosunNSCookies")("FoosunNSIndexTemplet")=NS_sys("IndexTemplet")
		Response.Cookies("FoosunNSCookies")("FoosunNSLinkType")=NS_sys("LinkType")
		Response.Cookies("FoosunNSCookies")("FoosunNSSiteName")=NS_sys("SiteName")
		Response.Cookies("FoosunNSCookies").Expires=Date()+1
		NS_sys.close:set NS_sys=nothing
	end if
End Sub
'MS参数登陆
Sub MSConfig_Cookies()
	if Conn.Execute("Select * from FS_MF_Sub_Sys Where Sub_Sys_ID='MS'").Eof then Exit Sub
	Response.Cookies("FoosunMSCookies")=""
	Response.Cookies("FoosunMSCookies")=Empty
	dim MS_sys,MallIndexExtNameNum,MallIndexExtName
	set MS_sys = Conn.execute("select top 1 FileExtName,SavePath,isDomain,IndexTemplt from FS_MS_SysPara")
	if Not MS_sys.eof then
		If IsNull(MS_sys("isDomain")) Then
			Response.Cookies("FoosunMSCookies")("FoosunMSDomain")=""
		Else
			Response.Cookies("FoosunMSCookies")("FoosunMSDomain")=MS_sys("isDomain")
		End If
		Response.Cookies("FoosunMSCookies")("FoosunMSDir")=MS_sys("SavePath")
		If Isnull(MS_sys("IndexTemplt")) Then
			Response.Cookies("FoosunMSCookies")("FoosunMSIndexTemplet")=""
		Else
			Response.Cookies("FoosunMSCookies")("FoosunMSIndexTemplet")=MS_sys("IndexTemplt")
		End If
		Response.Cookies("FoosunMSCookies").Expires=Date()+1
		MallIndexExtNameNum = MS_sys("FileExtName")
	Else
		MallIndexExtNameNum = 0
	End If
	MS_sys.close:set MS_sys=nothing
	If MallIndexExtNameNum = 0 Then
		MallIndexExtName = "html"
	ElseIf MallIndexExtNameNum = 1 Then
		MallIndexExtName = "htm"
	ElseIf MallIndexExtNameNum = 2 Then
		MallIndexExtName = "shtml"
	ElseIf MallIndexExtNameNum = 3 Then
		MallIndexExtName = "shtm"
	ElseIf MallIndexExtNameNum = 4 Then
		MallIndexExtName = "asp"
	Else
		MallIndexExtName = "html"
	End if
	Response.Cookies("FoosunMSCookies")("FoosunMSIndexHtml") = MallIndexExtName
End Sub
'DS参数登陆
Sub DSConfig_Cookies()
	if Conn.Execute("Select * from FS_MF_Sub_Sys Where Sub_Sys_ID='DS'").Eof then Exit Sub
	Response.Cookies("FoosunDSCookies")=""
	Response.Cookies("FoosunDSCookies")=Empty
	dim DS_sys
	set DS_sys = Conn.execute("select top 1 Lock,IPType,IPList,OverDueMode,DownDir,IsDomain,IndexPage,IndexTemplet,LinkType from FS_DS_SysPara")
	if Not DS_sys.eof then
		Response.Cookies("FoosunDSCookies")("FoosunDSLock")=DS_sys("Lock")
		Response.Cookies("FoosunDSCookies")("FoosunDSIPType")=DS_sys("IPType")
		Response.Cookies("FoosunDSCookies")("FoosunDSIPList")=""&DS_sys("IPList")
		Response.Cookies("FoosunDSCookies")("FoosunDSOverDueMode")=""&DS_sys("OverDueMode")

		Response.Cookies("FoosunDSCookies")("FoosunDSDomain")=""&DS_sys("IsDomain")
		Response.Cookies("FoosunDSCookies")("FoosunDSDownDir")=""&DS_sys("DownDir")
		Response.Cookies("FoosunDSCookies")("FoosunDSIndexPage")=""&DS_sys("IndexPage")
		Response.Cookies("FoosunDSCookies")("FoosunDSIndexTemplet")=""&DS_sys("IndexTemplet")
		Response.Cookies("FoosunDSCookies")("FoosunDSLinkType")=""&DS_sys("LinkType")
		Response.Cookies("FoosunDSCookies").Expires=Date()+1
		DS_sys.close:set DS_sys=nothing
	end if
End Sub

Sub Err_Show()
	dim Err_ShowChar
	Err_ShowChar = "<html xmlns=""http://www.w3.org/1999/xhtml"">"&chr(10)
	Err_ShowChar = Err_ShowChar & "<head>"&chr(10)
	Err_ShowChar = Err_ShowChar & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />"&chr(10)
	Err_ShowChar = Err_ShowChar & "<title>无标题文档</title>"&chr(10)
	Err_ShowChar = Err_ShowChar & "</head>"&chr(10)
	Err_ShowChar = Err_ShowChar & "<link href=""http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"/"&G_ADMIN_DIR&"/images/skin/Css_"&Session("Admin_Style_Num")&"/"&Session("Admin_Style_Num")&".css"" rel=""stylesheet"" type=""text/css"">"&chr(10)
	Err_ShowChar = Err_ShowChar & "<body  topmargin=""80""><p>"&chr(10)
	Err_ShowChar = Err_ShowChar & "<table width=""80%"" align=""center"" cellpadding=""10""><tr><td>"&chr(10)
	Err_ShowChar = Err_ShowChar & "<li>您没有此页操作权限没有权限!!</li>"&chr(10)
	Err_ShowChar = Err_ShowChar & "<li>请与系统管理员联系</li>"&chr(10)
	Err_ShowChar = Err_ShowChar & "<li>Powered by Foosun.Cn</li>"&chr(10)
	Err_ShowChar = Err_ShowChar & "</td><tr></table>"&chr(10)
	Err_ShowChar = Err_ShowChar & "</body>"&chr(10)
	Err_ShowChar = Err_ShowChar & "</html>"&chr(10)
	Response.Write Err_ShowChar
	Response.end
End Sub

Function isCorp()
	Dim userRs
	Set userRs=User_Conn.execute("Select UserNumber from FS_ME_CorpUser where usernumber='"&session("FS_usernumber")&"'")
	if not userRs.eof then
		isCorp=true
	else
		isCorp=false
	End if
End Function
Function Get_Label_Content(Str_Label_Name)
	'Get_Label_Content=Str_Label_Name
	'exit function
	Dim Rs_Label,StrSql,Label_ID,Str_Label_Names
	'处理标签缓存。Fsj
	'
	'
	'新标签部分，缓存
	
	If ""<>Application("Temp_Label_"&Str_Label_Name) then			
			Get_Label_Content=Application("Temp_Label_"&Str_Label_Name)					
		Else				
			StrSql = "SELECT LableContent FROM FS_MF_Lable WHERE isDel=0 AND LableName='"&Str_Label_Name&"'"
			Set Rs_Label = Conn.Execute(StrSql)
			If Not Rs_Label.Eof Then
				Application.Lock()
				Application("Temp_Label_"&Str_Label_Name)=Rs_Label("LableContent")
				Application.UnLock()
				Get_Label_Content=Application("Temp_Label_"&Str_Label_Name)		
			Else
				Get_Label_Content=""
			End If		
			Rs_Label.Close
			set Rs_Label=nothing
	End if
End Function

Function AddInCludeFileForPopNews()
	AddInCludeFileForPopNews = chr(10)&"<!--#include virtual=""/" & G_VIRTUAL_ROOT_DIR & "/FS_Inc/Const.asp"" -->" & Chr(13) & Chr(10)
	AddInCludeFileForPopNews = AddInCludeFileForPopNews&"<!--#include virtual=""/"&G_VIRTUAL_ROOT_DIR&"/FS_Inc/Function.asp""-->"&chr(10)
	AddInCludeFileForPopNews = AddInCludeFileForPopNews&"<!--#include virtual=""/"&G_VIRTUAL_ROOT_DIR&"/Fs_InterFace/MF_Function.asp""-->"&chr(10)
	AddInCludeFileForPopNews = AddInCludeFileForPopNews&"<!--#include virtual=""/"&G_VIRTUAL_ROOT_DIR&"/"&G_USER_DIR&"/lib/strlib.asp""-->"&chr(10)
	AddInCludeFileForPopNews = AddInCludeFileForPopNews&"<!--#include virtual=""/"&G_VIRTUAL_ROOT_DIR&"/"&G_USER_DIR&"/lib/Usercheck.asp""-->"&chr(10)
	AddInCludeFileForPopNews = AddInCludeFileForPopNews&"<!--#include virtual=""/"&G_VIRTUAL_ROOT_DIR&"/"&G_USER_DIR&"/lib/ReadPop.asp""-->"&chr(10)
End Function

Function GetPaginationSQL(f_TableName,f_SelectFieldNames,f_PageSize,f_PageIndex,f_OrderBY,f_OrderDESC,f_Where)
	Dim f_strTemp,f_strSQL,f_strOrder
	if f_OrderBY <> "" And f_OrderDESC <> "" then f_strOrder = f_OrderBY & " " & f_OrderDESC
	if f_OrderBY <> "ID" then
		if f_strOrder <> "" then
			f_strOrder = " order by " & "ID asc," & f_strOrder
		else
			f_strOrder = " order by ID asc"
		end if
		f_strTemp = ">(select max(ID)"
	else
		if LCase(f_OrderDESC) = "asc" then
			f_strTemp = ">(select max(ID)"
		else
			f_strTemp = "<(select min(ID)"
		end if
	end if
	if f_PageIndex = 1 then
		f_strTemp = ""
		if f_Where <> "" then f_strTemp = " where " + f_Where
		f_strSQL = "select top " & f_PageSize & " " + f_SelectFieldNames + " from " & f_TableName & "" & f_strTemp & f_strOrder
	else
		f_strSQL = "select top " & f_PageSize & " " + f_SelectFieldNames + " from " & f_TableName & " where ID" & f_strTemp & " from (select top " & (f_PageIndex-1)*f_PageSize & " ID from " & f_TableName & "" 
		if f_Where <> "" then f_strSQL = f_strSQL & " where " & f_Where
		f_strSQL = f_strSQL & f_strOrder & ") as tblTemp)"
		if f_Where <> "" then f_strSQL = f_strSQL & " And " & f_Where
		f_strSQL = f_strSQL & f_strOrder
				Response.Write(f_strSQL)
				Response.End
	end if
	GetPaginationSQL = f_strSQL
End Function

Function CheckTextNumber(f_Text,f_DefaultValue)
	if f_Text = "" then f_Text = f_DefaultValue
	if Not IsNumeric(f_Text) then f_Text = f_DefaultValue
	CheckTextNumber = CInt(f_Text)
End Function
%>
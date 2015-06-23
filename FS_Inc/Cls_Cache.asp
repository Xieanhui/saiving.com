<%
	Function Pre_Cache_Name
		Dim f_Pre_Cache_Name
		f_Pre_Cache_Name = Server.MapPath("/"&G_VIRTUAL_ROOT_DIR)
		Pre_Cache_Name = LCase(Replace(Replace(f_Pre_Cache_Name,":",""),"\",""))
	End Function

	Function Judge_Cache_Data(f_Sys_Name)
		Judge_Cache_Data = Application(Pre_Cache_Name & "_" & LCase(f_Sys_Name) & "_" & "judge")
	End Function
	
	Sub Load_Cache(f_Sys_Name,f_Select_Column_Str)
		If Judge_Cache_Data(f_Sys_Name) <> "1" Then
			Load_Cache_Config f_Sys_Name,f_Select_Column_Str
		End If
	End Sub

	Sub ReLoad_Cache(f_Sys_Name,f_Select_Column_Str)
		Load_Cache_Config f_Sys_Name,f_Select_Column_Str
	End Sub

	Sub Load_Cache_Config(f_Sys_Name,f_Select_Column_Str)
		Dim f_Rs,f_Sql,f_App_Name_Str,f_App_Value_Str,Field_Obj
		f_Sys_Name = LCase(f_Sys_Name)
		f_Sql = "select Top 1 " & f_Select_Column_Str & " From FS_" & f_Sys_Name & "_Config"
		Set f_Rs = Server.CreateObject(G_FS_RS)
		f_Rs.Open f_Sql,Conn,1,1
		If Not f_Rs.Eof Then
			If f_Select_Column_Str = "*" Then
				Application.Lock
				Application(Pre_Cache_Name & "_" & f_Sys_Name & "_" & "judge")="1"
				Application.UnLock
				Load_Install_Flag("*")
			End If
			For Each Field_Obj in f_Rs.Fields
				f_App_Name_Str = Pre_Cache_Name & "_" & f_Sys_Name & "_" & LCase(Field_Obj.Name)
				f_App_Value_Str= f_Rs(Field_Obj.Name)
				Application.Lock
				Application(f_App_Name_Str)=f_App_Value_Str
				Application.UnLock
			Next
		End If
		f_Rs.Close
		Set f_Rs = Nothing
	End Sub
	
	Sub Load_Install_Flag(f_Sys_Name)
		Dim f_Rs,f_Sql
		f_Sys_Name = LCase(f_Sys_Name)
		If f_Sys_Name = "*" Then
			f_Sql = "Select Sub_Sys_ID,Sub_Sys_Installed From FS_MF_Sub_Sys"
		Else
			f_Sql = "Select Sub_Sys_ID,Sub_Sys_Installed From FS_MF_Sub_Sys Where Sub_Sys_ID ='" & NoSqlHack(f_Sys_Name) & "'"
		End If
		Set f_Rs = Conn.Execute(f_Sql)
		Do While Not f_Rs.Eof
			If Not f_Rs.Eof Then
				Application.Lock
				Application(Pre_Cache_Name & "_" & f_Rs("Sub_Sys_ID") & "_installed")=f_Rs("Sub_Sys_Installed")
				Application.UnLock
			Else
				Application.Lock
				Application(Pre_Cache_Name & "_" & f_Rs("Sub_Sys_ID") & "_installed")="0"
				Application.UnLock
			End If
			f_Rs.MoveNext
		Loop
		f_Rs.Close
		Set f_Rs = Nothing
	End Sub

	Function Get_Cache_Value(f_Sys_Name,f_Column_Str)
		If Judge_Cache_Data(f_Sys_Name) <> "1" Then
			Call Load_Cache_Config(f_Sys_Name,"*")
		End If
		Get_Cache_Value = Application(Pre_Cache_Name & "_" & LCase(f_Sys_Name) & "_" & LCase(f_Column_Str))
	End Function

%>






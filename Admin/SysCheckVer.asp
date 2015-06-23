<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<%
'Copyright (c) 2006 Foosun Inc. Code by Terry Wen Time:2006.6
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.Charset="GB2312"
Dim Str_Act,StrInfo,ThisDomain,Str_UserID,ThisIp,ThisPort,Str_IsBiz,Str_Para,Str_SqlScript,Conn,Rs_Config
MF_Default_Conn
Function Replace_MF_default_Flag(f_File_Cont,f_NewsID,PageType)
	Dim f_REG_EX,f_REG_MATCH,f_REG_MATCHS,f_REG_PLACE_OBJ,f_TEST_LABLE_CONT_MATCHS,f_TEST_LABLE_CONT_MATCH,f_MORE_PAGES_DICT_OBJ
	Dim f_REG_Head_Str,f_REG_Tailor_Str,f_Match_Str,f_Lable_Cont,f_RERESH_OBJ,f_Sys_ID,f_Lable_Para_Value,f_Raw_Data
	Dim f_LABLE_DICT_OBJ,f_Lable,f_DICT_ITEMS_OBJ,f_DICT_KEYS_OBJ,f_i,f_j,f_ARRAY_OBJ,f_More_Pages_Flag_TF,f_NEWS_CONT_REPLACE_FLAG_OBJ
	Dim f_DICT_MATCHS,f_Replace_Flag_TF,f_NEWS_CONT_REG_OBJ,f_NEWS_CONT_MATCHS,f_Lable_News_Cont,f_Lable_One_Page_News_Cont,Str_MFFlag
	f_More_Pages_Flag_TF = False
	f_REG_Head_Str = "{FS400_"
	f_REG_Tailor_Str = "}"
	Set f_MORE_PAGES_DICT_OBJ = Server.CreateObject(G_FS_DICT)
	Set f_LABLE_DICT_OBJ = Server.CreateObject(G_FS_DICT)
	Set f_REG_EX = New RegExp
	f_REG_EX.Pattern = f_REG_Head_Str & ".*?" & f_REG_Tailor_Str
	f_REG_EX.IgnoreCase = True
	f_REG_EX.Global = True
	Set f_REG_MATCHS = f_REG_EX.Execute(f_File_Cont)
	f_Raw_Data = f_File_Cont
	For Each f_REG_MATCH in f_REG_MATCHS
		f_Match_Str = f_REG_MATCH.Value
		f_Lable = f_Match_Str
		f_Match_Str = Replace(f_Match_Str,Chr(13) & Chr(10),"")
		Set f_REG_PLACE_OBJ = New RegExp
		f_REG_PLACE_OBJ.IgnoreCase = True
		f_REG_PLACE_OBJ.Global = True
		f_Match_Str=Get_Label_Content(f_Match_Str)
		f_REG_PLACE_OBJ.Pattern = "{FS:.*}"
		Set f_TEST_LABLE_CONT_MATCHS = f_REG_PLACE_OBJ.Execute(f_Match_Str)
		if (f_TEST_LABLE_CONT_MATCHS.Count>0) then
			for Each f_TEST_LABLE_CONT_MATCH in f_TEST_LABLE_CONT_MATCHS
				f_Lable_Para_Value = f_TEST_LABLE_CONT_MATCH.Value
				f_Sys_ID = Mid(f_Lable_Para_Value,5,2)
				f_Lable_Para_Value = Mid(f_Lable_Para_Value,8,Len(f_Lable_Para_Value) - 8)
				Str_MFFlag=False
				if Request.Cookies("FoosunSUBCookie")("FoosunSUB" & f_Sys_ID) = "1" then
					Select Case f_Sys_ID
						Case "NS"
							Set f_RERESH_OBJ = New cls_NS
						Case "DS"
							Set f_RERESH_OBJ = New cls_DS
						Case "ME"
							Set f_RERESH_OBJ = New cls_ME
						Case "MF"
							Set f_RERESH_OBJ = New cls_MF
							Str_MFFlag=True
						Case Else
							Set f_RERESH_OBJ = New cls_Other
					End Select
					If Str_MFFlag Then
						f_Lable_Cont = f_RERESH_OBJ.get_LableChar(f_Lable_Para_Value,f_NewsID,PageType)
					Else
						f_Lable_Cont = f_RERESH_OBJ.get_LableChar(f_Lable_Para_Value,f_NewsID)
					End If
					f_Match_Str = Replace(f_Match_Str,f_TEST_LABLE_CONT_MATCH.Value,"{|}"&f_Lable_Cont&"{|}")
				End If
				Set f_RERESH_OBJ = Nothing
			Next
		End If
		If Not f_LABLE_DICT_OBJ.Exists(f_Lable) Then
			f_LABLE_DICT_OBJ.Add f_Lable,f_Match_Str
		End If
		Set f_REG_PLACE_OBJ = Nothing
	Next
	f_MORE_PAGES_DICT_OBJ.Add "-3",""
	f_DICT_ITEMS_OBJ = f_LABLE_DICT_OBJ.Items
	f_DICT_KEYS_OBJ = f_LABLE_DICT_OBJ.Keys
	Dim Str_Other,PageNum_Flag
	For f_i = 0 To f_LABLE_DICT_OBJ.Count - 1
		if Not f_More_Pages_Flag_TF then
			f_Replace_Flag_TF = False
			Set f_REG_PLACE_OBJ = New RegExp
			f_REG_PLACE_OBJ.Pattern = "{foosun_page_news}.*?{/foosun_page_news}"
			f_REG_PLACE_OBJ.IgnoreCase = True
			f_REG_PLACE_OBJ.Global = True

			Set f_DICT_MATCHS = f_REG_PLACE_OBJ.Execute(f_DICT_ITEMS_OBJ(f_i))

			if (f_DICT_MATCHS.Count>=1) And (f_More_Pages_Flag_TF=False) Then
				f_More_Pages_Flag_TF = True
				Str_Other=Split(f_DICT_ITEMS_OBJ(f_i),"{|}")
				f_ARRAY_OBJ = Split(Str_Other(1),f_DICT_MATCHS(0).Value)
				f_MORE_PAGES_DICT_OBJ.Add "-2",f_DICT_KEYS_OBJ(f_i)
				f_MORE_PAGES_DICT_OBJ.Add "-1",f_DICT_MATCHS(0).Value
				PageNum_Flag=False
				for f_j = LBound(f_ARRAY_OBJ) To UBound(f_ARRAY_OBJ)
					If Len(f_ARRAY_OBJ(f_j)) > 1 then
						If Not f_MORE_PAGES_DICT_OBJ.Exists(f_j) Then
							f_MORE_PAGES_DICT_OBJ.Add f_j,f_ARRAY_OBJ(f_j)
							PageNum_Flag=True
						End If 
					end if
				Next
				If PageNum_Flag Then
					f_Raw_Data = Replace(f_Raw_Data,f_DICT_KEYS_OBJ(f_i),Str_Other(0)&f_DICT_KEYS_OBJ(f_i)&Str_Other(2))
				End If 
				Set Str_Other=Nothing 
				f_Replace_Flag_TF = True
			Else
				f_DICT_ITEMS_OBJ(f_i)=Replace(f_DICT_ITEMS_OBJ(f_i),"{|}","")
			End If
			f_REG_PLACE_OBJ.Pattern = "\[fs\:page\]"
			Set f_DICT_MATCHS = f_REG_PLACE_OBJ.Execute(f_DICT_ITEMS_OBJ(f_i))
			Set f_NEWS_CONT_REG_OBJ = New RegExp
			f_NEWS_CONT_REG_OBJ.IgnoreCase = True
			f_NEWS_CONT_REG_OBJ.Global = True
			f_NEWS_CONT_REG_OBJ.Pattern = "\[FS\:CONTENT_START\][^\0]*\[FS\:CONTENT_END\]"
			Set f_NEWS_CONT_MATCHS = f_NEWS_CONT_REG_OBJ.Execute(f_DICT_ITEMS_OBJ(f_i))
			if (f_DICT_MATCHS.Count>=1) And (f_More_Pages_Flag_TF=False) And (f_NEWS_CONT_MATCHS.Count=1) then
				f_More_Pages_Flag_TF = True
				f_MORE_PAGES_DICT_OBJ.Add "-2",f_DICT_KEYS_OBJ(f_i)
				f_MORE_PAGES_DICT_OBJ.Add "-1","newsmorepage"
				f_Lable_News_Cont = f_NEWS_CONT_MATCHS(0).Value
				Set f_NEWS_CONT_REPLACE_FLAG_OBJ = New RegExp
				f_NEWS_CONT_REPLACE_FLAG_OBJ.IgnoreCase = True
				f_NEWS_CONT_REPLACE_FLAG_OBJ.Global = True
				f_NEWS_CONT_REPLACE_FLAG_OBJ.Pattern = "\[FS\:CONTENT_START\]"
				f_Lable_News_Cont = f_NEWS_CONT_REPLACE_FLAG_OBJ.Replace(f_Lable_News_Cont,"")
				f_NEWS_CONT_REPLACE_FLAG_OBJ.Pattern = "\[FS\:CONTENT_END\]"
				f_Lable_News_Cont = f_NEWS_CONT_REPLACE_FLAG_OBJ.Replace(f_Lable_News_Cont,"")
				Set f_NEWS_CONT_REPLACE_FLAG_OBJ = Nothing
				f_ARRAY_OBJ = Split(f_Lable_News_Cont,f_DICT_MATCHS(0).Value)
				PageNum_Flag=False
				Str_Other=Split(f_DICT_ITEMS_OBJ(f_i),f_NEWS_CONT_MATCHS(0).Value)
				for f_j = LBound(f_ARRAY_OBJ) To UBound(f_ARRAY_OBJ)
					if Len(f_ARRAY_OBJ(f_j)) > 1 then
						'f_Lable_One_Page_News_Cont = f_DICT_ITEMS_OBJ(f_i)
						'f_Lable_One_Page_News_Cont = Replace(f_Lable_One_Page_News_Cont,f_NEWS_CONT_MATCHS(0).Value,f_ARRAY_OBJ(f_j))
						if Not f_MORE_PAGES_DICT_OBJ.Exists(f_j) Then
							'f_MORE_PAGES_DICT_OBJ.Add f_j,f_Lable_One_Page_News_Cont
							f_MORE_PAGES_DICT_OBJ.Add f_j,f_ARRAY_OBJ(f_j)
						End If 
						PageNum_Flag=True
					end if
				Next
				If PageNum_Flag Then
					f_Raw_Data = Replace(f_Raw_Data,f_DICT_KEYS_OBJ(f_i),Str_Other(0)&f_DICT_KEYS_OBJ(f_i)&Str_Other(1))
				End If 
				Set Str_Other=Nothing 
				f_Replace_Flag_TF = True
			end if
			Set f_NEWS_CONT_REG_OBJ = Nothing
			Set f_REG_PLACE_OBJ = Nothing
			if Not f_Replace_Flag_TF Then
				f_DICT_ITEMS_OBJ(f_i)=Replace(f_DICT_ITEMS_OBJ(f_i),"{|}","")
				f_Raw_Data = Replace(f_Raw_Data,f_DICT_KEYS_OBJ(f_i),f_DICT_ITEMS_OBJ(f_i))
			End If
		Else
			f_DICT_ITEMS_OBJ(f_i)=Replace(f_DICT_ITEMS_OBJ(f_i),"{|}","")
			f_Raw_Data = Replace(f_Raw_Data,f_DICT_KEYS_OBJ(f_i),f_DICT_ITEMS_OBJ(f_i))
		End If
	Next
	f_MORE_PAGES_DICT_OBJ.Item("-3") = Replace(f_Raw_Data,"{|}","")
	f_LABLE_DICT_OBJ.RemoveAll
	Set f_LABLE_DICT_OBJ = Nothing
	''Dic_Test(f_DICT_ITEMS_OBJ)
	Set Replace_All_Flag = f_MORE_PAGES_DICT_OBJ
	Set f_MORE_PAGES_DICT_OBJ = Nothing
End Function
Str_Act = Trim(Request.Form("Act"))
If Str_Act = "" Then Str_Act = "Ver"
Select Case Str_Act
	Case "Ver"
		StrInfo = Str_get(1)
	Case "News"
		StrInfo = Str_get(2)
	Case Else
		StrInfo = "||"
End Select
StrInfo = Split(StrInfo,"||")
If StrInfo(0)="True" Then
	Response.Write "True||"&Str_Act&"||"&StrInfo(1)
ElseIf StrInfo(0)="False" Then
	Response.Write "False||"&Str_Act&"||"
End If
%>
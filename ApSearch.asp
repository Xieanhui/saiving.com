<% Option Explicit %>
<!--#include file="FS_Inc/Const.asp" -->
<!--#include file="FS_InterFace/MF_Function.asp" -->
<!--#include file="FS_Inc/Function.asp" -->
<!--#include file="FS_Inc/Func_Page.asp" -->
<!--#include file="FS_InterFace/AP_Public.asp" -->
<%
Server.ScriptTimeOut=999
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache"
response.Charset = "gb2312"
'-----�����ݿ�����
Dim Conn,User_Conn,StarTimeStr,EndDisTimeStr,New_Cls
MF_Default_Conn
MF_User_Conn
Get_MF_Config
StarTimeStr = Timer()
Set New_Cls = New cls_AP
'---�����������
Function Get_MF_Config()
	If Request.Cookies("FoosunSearchCookie")("Cookie_Domain") = Get_MF_Domain() then exit Function
	set CookRs = Conn.execute("select top 1 MF_Domain,MF_Site_Name,MF_eMail,MF_Copyright_Info  from FS_MF_Config")
	Response.Cookies("FoosunSearchCookie")("Cookie_Domain")=CookRs("MF_Domain") 
	Response.Cookies("FoosunSearchCookie")("Cookie_Copyright")=CookRs("MF_Copyright_Info") 
	Response.Cookies("FoosunSearchCookie")("Cookie_eMail")=CookRs("MF_eMail") 
	Response.Cookies("FoosunSearchCookie")("Cookie_Site_Name")=CookRs("MF_Site_Name") 
	Response.Cookies("FoosunSearchCookie").Expires=Date()+1
	CookRs.close : Set CookRs = Nothing 
End Function

'----------------------------------------------------------
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=10 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=10 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings>9</font>"   			'��ҳ 
toP10_=" <font face=webdings>7</font>"			'��ʮ 
toP1_=" <font face=webdings>3</font>"			'��һ
toN1_=" <font face=webdings>4</font>"			'��һ
toN10_=" <font face=webdings>8</font>"			'��ʮ
toL_="<font face=webdings>:</font>"				'βҳ
'----------------------------------------------------------

'----------------------
'��ȡ�������ݲ���
Dim GetTypeStr,Se_Type,Se_KeyWords,Se_JobType,Se_City,Se_Time,Job_Type
Dim MoneyMin,FreeMoney,Job_Time,Job_Edu,Job_Language,MoneyMonthStr,Type_Job_Str
Dim CookRs,CommRs,CommID,Search_Result
Dim RetunValue
Dim TypeToCom,CityToCom,TimeToCom,Job_TypeStr,MoneyStr,ReeeMStr,TimeStr
Dim EduStr,LanguageStr,KeyStr,SexStr
Dim EndTimeStr,SqlStr
Dim SearchSql,SearchObj,AllinfoNum
'----AJAX����ֵ����
GetTypeStr = NoSqlHack(Request.QueryString("GetType"))
If GetTypeStr = "" Then
	Response.Write "��������"
	Response.End
End If
'-----��ȡ��������
Se_Type = CintStr(Request.QueryString("JobClass"))					'0Ϊ�ҹ�����1Ϊ���˲�
Se_JobType = CintStr(Request.QueryString("JobTypeID"))				'��ҵ/ְҵid
Se_City = NoSqlHack(Request.QueryString("JobCity"))					'����id
Se_Time = NoSqlHack(Request.QueryString("JobTimeID"))					'ʱ�䷶Χ
Se_KeyWords = NoSqlHack(Trim(Request.QueryString("Sekey")))	'�ؼ���
'-------�����ֶ�Ϊ�߼�����
Job_Type = NoSqlHack(Request.QueryString("SearchClass"))				'�������ͣ���ְ/ȫְ
MoneyMin = CintStr(Request.QueryString("MoneyMin"))					'н��Χ
FreeMoney =NoSqlHack(Request.QueryString("FreeMoney"))				'����
Job_Time = CintStr(Request.QueryString("JobYear"))					'��������/����
Job_Edu = CintStr(Request.QueryString("JobEdu"))						'�����̶�
Job_Language = NoSqlHack(Request.QueryString("Job_Langage"))			'Ҫ����������
'----AJAX����ֵ
Select Case GetTypeStr
	Case "SearchType"
		Response.Write "���β�ѯ���:"
		Response.End
	Case"Copy"
		Response.Write Request.Cookies("FoosunSearchCookie")("Cookie_Copyright")
		Response.End
	Case Else
		Response.Write Search
		Response.End
End Select

Function Search()
	'---���ݲ����жϲ�ѯsql�������
	If CintStr(Se_Type) = 0 Then
		'---������ҵ��ѯ���˵�λid��
		If Trim(Se_JobType) = "" Or Not IsNumeric(Se_JobType) Then
			TypeToCom = ""
		Else
			Set CommRs = User_Conn.ExeCute("Select FS_ME_Users.UserNumber From FS_ME_CorpUser,FS_ME_Users Where FS_ME_Users.UserNumber = FS_ME_CorpUser.UserNumber And FS_ME_Users.isLock = 0 And FS_ME_Users.IsCorporation = 1 And FS_ME_CorpUser.C_VocationClassID = " & Se_JobType & " Order By FS_ME_Users.UserID Desc")
			If CommRs.Eof Then
				TypeToCom = ""
			Else
				Do While Not CommRs.Eof
					CommID = CommID & "','" & CommRs(0)
				CommRs.MoveNext
				Loop
				If Left(CommID,3) = "','" Then
					CommID = Right(CommID,Clng(Len(CommID) - 3))
				End If	 
				TypeToCom = " And UserNumber In('" & CommID & "')"'fucxi����Ҫ����
			End If
			CommRs.Close : Set CommRs = Nothing	
		End If
		'---ƥ�����
		If Se_City <> ""  Then
			If Se_City <> "ѡ�����ص�" Then
				CityToCom = " And WorkCity like '" & Se_City & "'"
			Else
				CityToCom = ""
			End If		 
		Else
			CityToCom = ""
		End If
		'---ƥ��ʱ��
		If Se_Time = "" Or Not IsNumeric(Se_Time) Then
			TimeToCom = ""
		Else
			If G_IS_SQL_DB = 1 Then
				TimeToCom = " And DateDiff(d,PublicDate,getdate()) <= " & Se_Time
				EndTimeStr = " And DateDiff(d,EndDate,getdate()) < 0"
			Else
				TimeToCom = " And DateDiff('d',PublicDate,Now()) <= " & Se_Time
				EndTimeStr = " And DateDiff('d',EndDate,Now()) < 0"
			End If
		End If
		'---ƥ�乤������
		If Job_Type = "" Or Not IsNumeric(Job_Type) Then
			Job_TypeStr = ""
		Else
			If Cint(Job_Type) = 0 Then
				Job_TypeStr = " And JobType = 2"
			Else
				Job_TypeStr = " And JobType = 1"
			End If		
		End If 
		'---ƥ��н��
		If MoneyMin = "" Or Not IsNumeric(MoneyMin) Then
			MoneyStr = ""
		Else
			Select Case MoneyMin
				Case 1
					MoneyStr = " And MoneyMonth >0 And MoneyMonth <= 1500"
				Case 2
					MoneyStr = " And MoneyMonth >1500 And MoneyMonth <= 1999"
				Case 3
					MoneyStr = " And MoneyMonth >=2000 And MoneyMonth <= 2999"
				Case 4 
					MoneyStr = " And MoneyMonth >=3000 And MoneyMonth <= 4499"
				Case 5
					MoneyStr = " And MoneyMonth >=4500 And MoneyMonth <= 5999"
				Case 6
					MoneyStr = " And MoneyMonth >=6000 And MoneyMonth <= 7999"
				Case 7
					MoneyStr = " And MoneyMonth >=8000 And MoneyMonth <= 9999"
				Case 8
					MoneyStr = " And MoneyMonth >=10000 And MoneyMonth <= 14999"
				Case 9
					MoneyStr = " And MoneyMonth >=15000 And MoneyMonth <= 19999"
				Case 10
					MoneyStr = " And MoneyMonth >=20000 And MoneyMonth <= 29999"
				Case 11
					MoneyStr = " And MoneyMonth >=30000 And MoneyMonth <= 49999"
				Case 12
					MoneyStr = " And MoneyMonth >=50000"
				Case Else
					MoneyStr = ""
			End Select
		End If
		'---ƥ�乤������
		If FreeMoney = "" Or Not IsNumeric(FreeMoney) Then
			ReeeMStr = ""
		Else
			ReeeMStr = " And FreeMoney = 1"
		End If
		'---ƥ�乤������
		If Job_Time = "" Or Not IsNumeric(Job_Time) Then
			TimeStr = ""
		Else
			TimeStr = " And WorkAge = " & Cint(Job_Time)
		End If
		'---ƥ������̶�
		If Job_Edu = "" Or Not IsNumeric(Job_Edu) Then
			EduStr = ""
		Else
			EduStr = " And EducateExp = " & Job_Edu
		End If
		'---ƥ����������
		If Job_Language = "" Or Not IsNumeric(Job_Language) Then
			LanguageStr = ""
		Else
			Select Case Cint(Job_Language)
				Case 1 
					LanguageStr = " And ResumeLang = 'Ӣ��'"
				Case 2
					LanguageStr = " And ResumeLang = '����'"
				Case 3
					LanguageStr = " And ResumeLang = '����'"
				Case 4
					LanguageStr = " And ResumeLang = '����'"
				Case 5
					LanguageStr = " And ResumeLang = '����'"
				Case 6
					LanguageStr = " And ResumeLang = '��������'"
				Case 7
					LanguageStr = " And ResumeLang = '������'"
				Case 8
					LanguageStr = " And ResumeLang = '��������'"
				Case 9
					LanguageStr = " And ResumeLang = '����'"
				Case 10
					LanguageStr = " And ResumeLang = '����'"
				Case Else
					LanguageStr = ""	
			End Select	
		End If
		'---�����ؼ���ƥ��
		If Se_KeyWords <> "" Then
			If Se_KeyWords = "��ѯ�ؼ���" Then
				KeyStr = ""
			Else
				KeyStr = " And JobName like '%" & Se_KeyWords & "%'"
			End If
		Else
			KeyStr = ""
		End If				  
		'---��ϲ�ѯ����	��Ҫ���ط�
		SqlStr = TypeToCom & CityToCom & TimeToCom & EndTimeStr & Job_TypeStr & MoneyStr & ReeeMStr & TimeStr & EduStr & LanguageStr & KeyStr
		SearchSql = "Select PID,UserNumber,JobName,JobDescription,ResumeLang,WorkCity,PublicDate,EndDate,Jlmode,NeedNum,EducateExp,Sex,WorkAge,Age,JobType,OtherJobDes,MoneyMonth,FreeMoney,OtherMoneyDes,HolleType From FS_AP_Job_Public Where 1=1" & SqlStr & " Order By PID Desc"
		'----------------------
		Set SearchObj = Server.CreateObject(G_FS_RS)
		SearchObj.Open SearchSql,Conn,1,1
		If SearchObj.Eof Then
			Search_Result = "��Ǹ,û���ҵ����������Ľ��,�����������������"
		Else
			AllinfoNum = SearchObj.RecordCount
			SearchObj.PageSize=int_RPP
			cPageNo=Request.QueryString("Page")
			If cPageNo="" or not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo<1 Then cPageNo=1
			If cPageNo>SearchObj.PageCount Then cPageNo=SearchObj.PageCount 
			SearchObj.AbsolutePage=cPageNo
			'----------------------------------------
			Search_Result = "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbnewline
			Search_Result = Search_Result & "<tr>" & vbnewline
			Search_Result = Search_Result & "<td height=""30"" align=""left"" valign=""middle"" width=""30%"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:13px; color:#3399cc; font-weight:bold; margin:0px; padding:0px;"">ְλ</span></td>" & vbnewline
			Search_Result = Search_Result & "<td height=""30"" align=""left"" valign=""middle"" width=""30%"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:13px; color:#3399cc; font-weight:bold; margin:0px; padding:0px;"">��˾</span></td>" & vbnewline
			Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" width=""10%"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:13px; color:#3399cc; font-weight:bold; margin:0px; padding:0px;"">�����ص�</span></td>" & vbnewline
			Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" width=""10%"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:13px; color:#3399cc; font-weight:bold; margin:0px; padding:0px;"">��������</span></td>" & vbnewline
			Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" width=""10%"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:13px; color:#3399cc; font-weight:bold; margin:0px; padding:0px;"">��н</span></td>" & vbnewline
			Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" width=""10%"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:13px; color:#3399cc; font-weight:bold; margin:0px; padding:0px;"">��Ա����</span></td>" & vbnewline
			Search_Result = Search_Result & "</tr>" & vbnewline
			For int_Start = 1 To int_RPP
			If SearchObj.Eof Then Exit For	
				If SearchObj("MoneyMonth") = "" Or isNull(SearchObj("MoneyMonth")) Then
					MoneyMonthStr = "����"
				Else
					MoneyMonthStr = SearchObj("MoneyMonth")
				End If		
				Search_Result = Search_Result & "<tr>" & vbnewline
				Search_Result = Search_Result & "<td height=""30"" align=""left"" valign=""middle"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:12px; color:#3399cc;""><font color=#3399cc;><a title=""����鿴��ְλ��ϸ��Ϣ"" target=""_blank"" class=""Job"" href=""" & New_Cls.get_infoLink(SearchObj("PID")) & """>" & SearchObj("JobName") & "</a></font></span></td>"
				Search_Result = Search_Result & "<td height=""30"" align=""left"" valign=""middle"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:12px; color:#3399cc;"">" & GetCorpUserName(SearchObj("UserNumber"),"UName") & "</span></td>"
				Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:12px; color:#3399cc;"">" & SearchObj("WorkCity") & "</span></td>"
				Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:12px; color:#3399cc;"">" & SearchObj("PublicDate") & "</span></td>"
				Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:12px; color:#3399cc;"">" & MoneyMonthStr & "</span></td>"
				Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:12px; color:#3399cc;"">" & GetCorpUserName(SearchObj("UserNumber"),"UType") & "</span></td>"
				Search_Result = Search_Result & "</tr>"
			SearchObj.MoveNext
			Next
			Search_Result = Search_Result & "<tr>" & vbnewline
			Search_Result = Search_Result & "<td height=""30"" colspan=""6"" align=""right"" valign=""bottom""><span style=""font-size:12px; color:#3399cc; margin:0px; padding:0px;"">" & fPageCount(SearchObj,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo) & "</span></td>" & vbnewline
			Search_Result = Search_Result & "</tr>" & vbnewline
			Search_Result = Search_Result & "</table>"
			'------
			EndDisTimeStr = "����ѯ��<font color=red>" & AllinfoNum & "</font>����¼,��ѯ��ʱ<b>" & FormatNumber((Timer()-StarTimeStr),2,-1) & "</b>����"	
			SearchObj.Close : Set SearchObj = NOthing
		End If
	Else	'------���˲�
		'---ƥ����ҵ����λ
		If Trim(Se_JobType) = "" Or Not IsNumeric(Se_JobType) Then
			TypeToCom = ""
		Else
			Set CommRs = Conn.ExeCute("Select UserNumber From FS_AP_Resume_Position Where Job = (Select JOb From FS_AP_Job Where JID = " & Se_JobType & ") Order By BID Desc")
			If CommRs.Eof Then
				TypeToCom = ""
			Else
				Do While Not CommRs.Eof
					CommID = CommID & "','" & CommRs(0)
				CommRs.MoveNext
				Loop
				If Left(CommID,3) = "','" Then
					CommID = Right(CommID,Clng(Len(CommID) - 3))
				End If	 
				TypeToCom = " And FS_AP_Resume_BaseInfo.UserNumber In('" & FormatIntArr(CommID) & "')"'fucxi����Ҫ����
			End If
			CommRs.Close : Set CommRs = Nothing	
		End If
		'---ƥ�乤������
		If Se_City <> ""  Then
			If Se_City <> "ѡ�����ص�" Then
				CityToCom = " And FS_AP_Resume_WorkCity.City like '%" & Se_City & "%'"
			Else
				CityToCom = ""
			End If		 
		Else
			CityToCom = ""
		End If
		'---��������
		If Job_Type = "" Or Not IsNumeric(Job_Type) Then
			Job_TypeStr = ""
		Else
			If Cint(Job_Type) = 0 Then
				Job_TypeStr = " And FS_AP_Resume_Intention.WorkType = 1"
			Else
				Job_TypeStr = " And FS_AP_Resume_Intention.WorkType <> 1"
			End If		
		End If
		'---ƥ�乤������
		If Job_Time = "" Or Not IsNumeric(Job_Time) Then
			TimeStr = ""
		Else
			TimeStr = " And FS_AP_Resume_BaseInfo.WorkAge = " & CintStr(Job_Time)
		End If
		'---ƥ������̶�
		If Job_Edu = "" Or Not IsNumeric(Job_Edu) Then
			EduStr = ""
		Else
			Select Case Job_Edu
				Case 1
					EduStr = " And FS_AP_Resume_BaseInfo.XueLi Like '%��ר%'"
				Case 2
					EduStr = " And FS_AP_Resume_BaseInfo.XueLi Like '%��ר%'"
				Case 3
					EduStr = " And FS_AP_Resume_BaseInfo.XueLi Like '%����%'"
				Case 4
					EduStr = " And FS_AP_Resume_BaseInfo.XueLi Like '%˶ʿ%'"
				Case Else
					EduStr = ""
			End Select						
		End If
		'---�����ؼ���ƥ��
		If Se_KeyWords <> "" Then
			If Se_KeyWords = "��ѯ�ؼ���" Then
				KeyStr = ""
			Else
				KeyStr = " And FS_AP_Resume_BaseInfo.Uname like '%" & Se_KeyWords & "%'"
			End If
		Else
			KeyStr = ""
		End If				  
		'---��ϲ�ѯ����
		SqlStr = TypeToCom & CityToCom & Job_TypeStr & EduStr & KeyStr  
		SearchSql = "Select FS_AP_Resume_BaseInfo.UserNumber As UNum,FS_AP_Resume_BaseInfo.Sex As Usex,FS_AP_Resume_BaseInfo.Uname As NameUser,FS_AP_Resume_WorkCity.City As ucity,FS_AP_Resume_Intention.WorkType As UworkType From FS_AP_Resume_BaseInfo,FS_AP_Resume_WorkCity,FS_AP_Resume_Intention Where FS_AP_Resume_BaseInfo.UserNumber = FS_AP_Resume_WorkCity.UserNumber And FS_AP_Resume_BaseInfo.UserNumber = FS_AP_Resume_Intention.UserNumber" & SqlStr & " Order By FS_AP_Resume_BaseInfo.BID Desc"
		'---------------------------------------------
		Set SearchObj = Server.CreateObject(G_FS_RS)
		SearchObj.Open SearchSql,Conn,1,1
		If SearchObj.Eof Then
			Search_Result = "��Ǹ,û���ҵ����������Ľ��,�����������������"
		Else
			AllinfoNum = SearchObj.RecordCount
			SearchObj.PageSize=int_RPP
			cPageNo=CintStr(Request.QueryString("Page"))
			If cPageNo="" or not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo<1 Then cPageNo=1
			If cPageNo>SearchObj.PageCount Then cPageNo=SearchObj.PageCount 
			SearchObj.AbsolutePage=cPageNo
			'----------------------------------------
			Search_Result = "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbnewline
			Search_Result = Search_Result & "<tr>" & vbnewline
			Search_Result = Search_Result & "<td height=""30"" align=""left"" valign=""middle"" width=""30%"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:13px; color:#3399cc; font-weight:bold; margin:0px; padding:0px;"">ID</span></td>" & vbnewline
			Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" width=""10%"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:13px; color:#3399cc; font-weight:bold; margin:0px; padding:0px;"">����</span></td>" & vbnewline
			Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" width=""30%"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:13px; color:#3399cc; font-weight:bold; margin:0px; padding:0px;"">����ְλ</span></td>" & vbnewline
			Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" width=""10%"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:13px; color:#3399cc; font-weight:bold; margin:0px; padding:0px;"">�����ص�</span></td>" & vbnewline
			Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" width=""10%"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:13px; color:#3399cc; font-weight:bold; margin:0px; padding:0px;"">��������</span></td>" & vbnewline
			Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" width=""10%"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:13px; color:#3399cc; font-weight:bold; margin:0px; padding:0px;"">�Ա�</span></td>" & vbnewline
			Search_Result = Search_Result & "</tr>" & vbnewline
			For int_Start = 1 To int_RPP
			If SearchObj.Eof Then Exit For	
				If SearchObj("UworkType") = "" Or Not IsNumeric(SearchObj("UworkType")) Then
					Type_Job_Str = "����"
				Else
					If Cint(SearchObj("UworkType")) = 1 Then
						Type_Job_Str = "ȫְ"
					Else
						Type_Job_Str = "��ְ"
					End If
				End If
				
				If SearchObj("Usex") = "" Or Not IsNumeric(SearchObj("Usex")) Then
					SexStr = "��"
				Else
					If cint(SearchObj("Usex")) = 0 Then
						SexStr = "��"
					Else
						SexStr = "Ů"
					End if
				End If					
				Search_Result = Search_Result & "<tr>" & vbnewline
				Search_Result = Search_Result & "<td height=""30"" align=""left"" valign=""middle"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:12px; color:#3399cc;""><font color=#3399cc;><a target=""_blank"" class=""Job"" href=""http://"&request.Cookies("FoosunSearchCookie")("Cookie_Domain")&"/"&G_USER_DIR&"/Job/Person.asp?UID=" & SearchObj("UNum") & """ title=""����鿴���û�����"">" & SearchObj("UNum") &"</a></font></span></td>"
				Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:12px; color:#3399cc;""><font color=#3399cc;><a target=""_blank"" class=""Job"" href=""http://"&request.Cookies("FoosunSearchCookie")("Cookie_Domain")&"/job/Job_Read.asp?ID=" & SearchObj("UNum") & """ title=""����鿴���û���ϸ��Ϣ"">" & SearchObj("NameUser") &"</a></font></span></td>"
				Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:12px; color:#3399cc;"">" & GetOneUserJobType(SearchObj("UNum")) &"</span></td>"
				Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:12px; color:#3399cc;"">" & SearchObj("ucity") &"</span></td>"
				Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:12px; color:#3399cc;"">" & Type_Job_Str &"</span></td>"
				Search_Result = Search_Result & "<td height=""30"" align=""center"" valign=""middle"" style=""border-bottom:dashed 1px #cccccc""><span style=""font-size:12px; color:#3399cc;"">" & SexStr &"</span></td>"
				Search_Result = Search_Result & "</tr>"
			SearchObj.MoveNext
			Next
			Search_Result = Search_Result & "<tr>" & vbnewline
			Search_Result = Search_Result & "<td height=""30"" colspan=""6"" align=""right"" valign=""bottom""><span style=""font-size:12px; color:#3399cc; margin:0px; padding:0px;"">" & fPageCount(SearchObj,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo) & "</span></td>" & vbnewline
			Search_Result = Search_Result & "</tr>" & vbnewline
			Search_Result = Search_Result & "</table>"
			EndDisTimeStr = "����ѯ��<font color=red>" & AllinfoNum & "</font>����¼,��ѯ��ʱ<b>" & FormatNumber((Timer()-StarTimeStr),2,-1) & "</b>��"		
			SearchObj.Close : Set SearchObj = NOthing
		End If
	End If
	Search = Search_Result & "$$$" & EndDisTimeStr
End Function

'---------------
'---����ҵ�û�id�õ���ҵ����
Function GetCorpUserName(UserID,GetTypeStr)
	If UserID = "" Then : GetCorpUserName = "" : Exit Function : End If
	Dim Rs
	If GetTypeStr = "UName" Then
		Set Rs = User_Conn.ExeCute("Select C_Name From FS_ME_CorpUser Where UserNumber = '" & NoSqlHack(UserID) & "'" )
		IF Not Rs.Eof Then
			GetCorpUserName = Rs(0)
		Else
			GetCorpUserName = ""
		End If
		Rs.CLose : Set Rs = Nothing
	Else
		Set Rs = Conn.ExeCute("Select GroupLevel From FS_AP_UserList Where UserNumber = '" & NoSqlHack(UserID) & "'")
		If Not Rs.Eof Then
			If Rs(0) = "" Or Not IsNumeric(Rs(0)) Then
				GetCorpUserName = ""
			Else
				If Cint(Rs(0)) = 1 Then
					GetCorpUserName = "��ͨ�û�"
				ElseIf Cint(Rs(0)) = 2 Then
					GetCorpUserName = "�����û�"
				ElseIf Cint(Rs(0)) = 3 Then
					GetCorpUserName = "VIP�û�"
				End if
			End If					
		Else
			GetCorpUserName = ""
		End If
		Rs.Close : Set Rs = Nothing
	End If
	GetCorpUserName = GetCorpUserName						
End Function

'---���û���ŵõ����û�����ְ
Function GetOneUserJobType(UserID)
If UserID = "" Then : GetOneUserJobType = "" : Exit Function : End If
Dim Rs
	Set Rs = Conn.ExeCute("Select Top 1 Job From FS_AP_Resume_Position Where UserNumber = '" & NoSqlHack(UserID) & "' Order By BID Desc")
	If Not Rs.Eof Then
		GetOneUserJobType = Rs(0)
	Else
		GetOneUserJobType = ""
	End If		
	Rs.Close : Set Rs = Nothing
End Function

'---------------
Conn.Close : Set Conn = NOthing
User_Conn.Close : Set User_Conn = NOthing
%>







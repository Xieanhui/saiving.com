<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
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
<!--#include file="lib/cls_main.asp" -->
<%'Copyright (c) 2006 Foosun Inc. 
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,DS_Rs,DS_Sql ,DS_Rs1,DS_Sql1 ,sRootDir,str_CurrPath,fileDirRule,fileNameRule,sys_FileExtName
Dim AutoDelete,Months,classid,class_rs,icNum
Dim Fs_down,tmp_sFileExtName,tmp_sTemplets,str_ClassID         '���������Զ����ֶε�-------------
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
MF_Default_Conn 
MF_User_Conn        
MF_Session_TF
if not MF_Check_Pop_TF("Down_List") then Err_Show
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if

Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=15 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=10 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings>9</font>"   			'��ҳ 
toP10_=" <font face=webdings>7</font>"			'��ʮ 
toP1_=" <font face=webdings>3</font>"			'��һ
toN1_=" <font face=webdings>4</font>"			'��һ
toN10_=" <font face=webdings>8</font>"			'��ʮ
toL_="<font face=webdings>:</font>"				'βҳ
set DS_Rs = Conn.execute("select top 1 FileDirRule,fileNameRule from FS_DS_SysPara")
if not DS_Rs.eof then 
	fileDirRule = DS_Rs("FileDirRule")
	fileNameRule = DS_Rs("fileNameRule")
	if fileDirRule = "" or fileDirRule = "" then response.Redirect("../error.asp?ErrCodes=<li>Ŀ¼���ļ�������δ����.���Ƚ���ϵͳ��������.</li>") : response.End()
else
	response.Redirect("../error.asp?ErrCodes=<li>Ŀ¼���ļ�������δ����.���Ƚ���ϵͳ��������.</li>") : response.End()
end if	
DS_Rs.close
''------------------------�Զ����ֶβ������ݿ�ʼ-----------
set Fs_down = new Cls_News
	Fs_down.GetSysParam()
str_ClassID = NoSqlHack(Request.QueryString("ClassID"))
	if Trim(str_ClassID)<>"" then
		Dim tmp_class_obj,tmp_defineid
		set tmp_class_obj = conn.execute("select FileExtName,NewsTemplet,DefineID from FS_DS_Class where classID='"& NoSqlHack(str_ClassID) &"'")
		if tmp_class_obj.eof then
			tmp_class_obj.close:set tmp_class_obj = nothing
			response.Write "����Ĳ���,�Ҳ�����Ŀ"
			Response.end
		end if
		Select Case tmp_class_obj(0)
				Case "html"
					tmp_sFileExtName = 0
				Case "htm"
					tmp_sFileExtName =1
				Case "shtml"
					tmp_sFileExtName = 2
				Case "shtm"
					tmp_sFileExtName = 3
				Case "asp"
					tmp_sFileExtName = 4
		End Select
		tmp_sTemplets = tmp_class_obj(1)
		tmp_defineid = tmp_class_obj(2)
		set tmp_class_obj = nothing
	Else
		tmp_defineid = 0
		tmp_sFileExtName = Fs_down.fileExtName
		tmp_sTemplets = Replace("/"& G_TEMPLETS_DIR &"/down/down.htm","//","/")
	End if
	if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
	If Temp_Admin_Is_Super = 1 then
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	Else
		If Temp_Admin_FilesTF = 0 Then
			str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
		Else
			str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
		End If	
	End if
	'��ȡ�����ֶ���Ϣ,���浽����CustColumnArr��
	'(c)2002-2006 ��Ȩ���У�Foosun Inc.
	if not isnull(trim(tmp_defineid)) or trim(tmp_defineid)>0 then
		Dim CustColumnRs,CustSql,CustColumnArr
		CustSql="select DefineID,ClassID,D_Name,D_Coul,D_Type,D_isNull,D_Value,D_Content,D_SubType from [FS_MF_DefineTable] Where D_SubType='DS' and  Classid="& NoSqlHack(tmp_defineid) &""
		Set CustColumnRs=CreateObject(G_FS_RS)
		CustColumnRs.Open CustSql,Conn,1,3
		If Not CustColumnRs.Eof Then
			CustColumnArr=CustColumnRs.GetRows()
		End If
		CustColumnRs.close:Set CustColumnRs = Nothing
	end if
'------------------�Զ����ֶβ������ݽ���--------------------	

'�õ����ر���·��
Function SavePath(f_num)
SavePath = ""
Select Case f_num
		Case 0
			SavePath = "/" & year(now)&"-"&month(now)&"-"&day(now)
		Case 1
			SavePath = "/" & year(now)&"/"&month(now)&"/"&day(now)
		Case 2
			SavePath = "/" & year(now)&"/"&month(now)&"-"&day(now)
		Case 3
			SavePath = "/" & year(now)&"-"&month(now)&"/"&day(now)
		Case 4
			SavePath = "/"
		Case 5
			SavePath = "/" & year(now)&"/"&month(now)
		Case 6
			SavePath = "/" & year(now)&"/"&month(now)&day(now)
		Case 7
			SavePath = "/" & year(now)&month(now)&day(now)
End Select		
End Function
'����û��ļ���
Function strFileNameRule(str,f_idTF,f_id)
	strFileNameRule = ""
	Dim f_strFileNamearr,f_str0,f_str1,f_str2,f_str3,f_str4,Getstr,f_str5,f_str6
	f_strFileNamearr = split(str,"$")
	'FS$Y M D H I S$4$$_$0$0
	f_str0 = f_strFileNamearr(0)
	f_str1 = f_strFileNamearr(1)
	f_str2 = f_strFileNamearr(2)
	f_str3 = f_strFileNamearr(3)
	f_str4 = f_strFileNamearr(4)
	f_str5 = f_strFileNamearr(5)
	f_str6 = f_strFileNamearr(6)
	strFileNameRule = strFileNameRule & f_strFileNamearr(0)
	If Instr(1,f_strFileNamearr(1),"Y",1)<>0 then
		if Len(Trim(Cstr(f_strFileNamearr(4))))<>0 then
			strFileNameRule = strFileNameRule & right(year(now),2)&f_strFileNamearr(4)
		Else
			strFileNameRule = strFileNameRule & right(year(now),2)
		End if
	End if
	If Instr(1,f_strFileNamearr(1),"M",1)<>0 then
			if Len(Trim(Cstr(f_strFileNamearr(4))))<>0 then
				strFileNameRule = strFileNameRule & month(now)&f_strFileNamearr(4)
			Else
				strFileNameRule = strFileNameRule& month(now)
			End if
	End if
	If Instr(1,f_strFileNamearr(1),"D",1)<>0 then
			if Len(Trim(Cstr(f_strFileNamearr(4))))<>0 then
				strFileNameRule = strFileNameRule & day(now)&f_strFileNamearr(4)
			Else
				strFileNameRule = strFileNameRule& day(now)
			End if
	End if
	If Instr(1,f_strFileNamearr(1),"H",1)<>0 then
			if Len(Trim(Cstr(f_strFileNamearr(4))))<>0 then
				strFileNameRule = strFileNameRule & hour(now)&f_strFileNamearr(4)
			Else
				strFileNameRule = strFileNameRule& hour(now)
			End if
	End if

	If Instr(1,f_strFileNamearr(1),"I",1)<>0 then
			if Len(Trim(Cstr(f_strFileNamearr(4))))<>0 then
				strFileNameRule = strFileNameRule & minute(now)&f_strFileNamearr(4)
			Else
				strFileNameRule = strFileNameRule& minute(now)
			End if
	End if
	If Instr(1,f_strFileNamearr(1),"S",1)<>0 then
			if Len(Trim(Cstr(f_strFileNamearr(4))))<>0 then
				strFileNameRule = strFileNameRule & second(now)&f_strFileNamearr(4)
			Else
				strFileNameRule = strFileNameRule& second(now)
			End if
	End if
	Randomize
	Dim f_Randchar,f_Randchararr,f_RandLen,f_iR,f_Randomizecode
	f_Randchar="0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
	f_Randchararr=split(f_Randchar,",") 
	If f_strFileNamearr(2)="2" then
		if f_strFileNamearr(3)="1" then
			f_RandLen=2 
			for f_iR=1 to f_RandLen
			f_Randomizecode=f_Randomizecode&f_Randchararr(Int((21*Rnd)))
			next 
			strFileNameRule = strFileNameRule &  f_Randomizecode
		Else
			strFileNameRule = strFileNameRule &  CStr(Int((99 * Rnd) + 1))
		End if
	Elseif f_strFileNamearr(2)="3" then
		if f_strFileNamearr(3)="1" then
			f_RandLen=3 
			for f_iR=1 to f_RandLen
			f_Randomizecode=f_Randomizecode&f_Randchararr(Int((21*Rnd)))
			next 
			strFileNameRule = strFileNameRule &  f_Randomizecode
		Else
			strFileNameRule = strFileNameRule &  CStr(Int((999* Rnd) + 1))
		End if
	Elseif f_strFileNamearr(2)="4" then
		if f_strFileNamearr(3)="1" then
			f_RandLen=4 
			for f_iR=1 to f_RandLen
			f_Randomizecode=f_Randomizecode&f_Randchararr(Int((21*Rnd)))
			next 
			strFileNameRule = strFileNameRule &  f_Randomizecode
		Else
			strFileNameRule = strFileNameRule &  CStr(Int((9999* Rnd) + 1))
		End if
	Elseif f_strFileNamearr(2)="5" then
		if f_strFileNamearr(3)="1" then
			f_RandLen=5 
			for f_iR=1 to f_RandLen
			f_Randomizecode=f_Randomizecode&f_Randchararr(Int((21*Rnd)))
			next 
			strFileNameRule = strFileNameRule &  f_Randomizecode
		Else
			strFileNameRule = strFileNameRule &  CStr(Int((99999* Rnd) + 1))
		End if
   End if
 if f_strFileNamearr(5) = "1" then
	 strFileNameRule = strFileNameRule&f_strFileNamearr(4)&"�Զ����ID"
 End if
 if f_str6 = "1" then
	 strFileNameRule = strFileNameRule&f_strFileNamearr(4)&"ΨһDownID"
 End if
	 strFileNameRule = strFileNameRule
End Function

''�õ���ر���ֵ��
Function Get_OtherTable_Value(This_Fun_Sql)
	Dim This_Fun_Rs
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
	else
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	end if
	if instr(lcase(This_Fun_Sql)," in ")>0 then 
		do while not This_Fun_Rs.eof
			Get_OtherTable_Value = Get_OtherTable_Value & This_Fun_Rs(0) &","
			This_Fun_Rs.movenext
		loop
	else			
		if not This_Fun_Rs.eof then 
			Get_OtherTable_Value = This_Fun_Rs(0)
		else
			Get_OtherTable_Value = ""
		end if
	end if	
	set This_Fun_Rs=nothing 
End Function

Function Get_FildValue_List(This_Fun_Sql,EquValue,Get_Type)
'''This_Fun_Sql ����sql���,EquValue�����ݿ���ͬ��ֵ�����<option>�����selected,Get_Type=1Ϊ<option>
Dim Get_Html,This_Fun_Rs,Text
On Error Resume Next
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
	else
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	end if			
If Err<>0 then response.Redirect("../error.asp?ErrCodes=<li>"&Err.Number&"������"&Err.Description&"��Ǹ,�����Sql���"&This_Fun_Sql&"������.������ֶβ�����.</li>")
if isnull(EquValue) then EquValue = ""
do while not This_Fun_Rs.eof 
	select case Get_Type
	  case 1
		''<option>		
		if instr(This_Fun_Sql,",") >0 then 
			Text = This_Fun_Rs(1)
		else
			Text = This_Fun_Rs(0)
		end if	
		if cstr(EquValue) = cstr(This_Fun_Rs(0)) then 
			Get_Html = Get_Html & "<option value="""&This_Fun_Rs(0)&"""  style=""color:#0000FF"" selected>"&Text&"</option>"&vbNewLine
		else
			Get_Html = Get_Html & "<option value="""&This_Fun_Rs(0)&""">"&Text&"</option>"&vbNewLine
		end if		
	  case else
		exit do : Get_FildValue_List = "Get_Typeֵ�������" : exit Function
	end select
	This_Fun_Rs.movenext
loop
This_Fun_Rs.close
Get_FildValue_List = Get_Html
End Function

Function Get_While_Info(Add_Sql,orderby)
	Dim Get_Html,This_Fun_Sql,ii,db_ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp
	Str_Tmp = "ID,DownLoadID,ClassID,Description,Accredit,AddTime,Appraise,AuditTF,BrowPop,ClickNum,EditTime,EMail,FileExtName,FileName," _
		&"FileSize,[Language],Name,NewsTemplet,PassWord,Pic,Property,Provider,ProviderUrl,RecTF,ReviewTF,ShowReviewTF,SystemType,Types,Version,OverDue,ConsumeNum,speicalId"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_DS_List"
	if request.QueryString("Act")="SearchGo" then 
		Arr_Tmp = split(Str_Tmp,",")
		for each Str_Tmp in Arr_Tmp
			Req_Str = NoSqlHack(request(Str_Tmp))
			if Req_Str<>"" then 				
				select case Str_Tmp
					case "ID","Accredit","AddTime","Appraise","AuditTF","ClickNum","EditTime","Property","RecTF","ReviewTF","ShowReviewTF","Types","OverDue","ConsumeNum"
					''����,����
						regxp = "|<|>|=|<=|>=|<>|"
						if instr(regxp,"|"&left(Req_Str,1)&"|")>0 or instr(regxp,"|"&left(Req_Str,2)&"|")>0 then 
							New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" "& Req_Str
						elseif instr(Req_Str,"*")>0 then 
							if left(Req_Str,1)="*" then Req_Str = "%"&mid(Req_Str,2)
							if right(Req_Str,1)="*" then Req_Str = mid(Req_Str,1,len(Req_Str) - 1) & "%"							
							New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" like '"& Req_Str &"'"							
						else	
							New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" = "& Req_Str
						end if		
					case else
					''�ַ�
						New_Search_Str = and_where(New_Search_Str) & Search_TextArr(Req_Str,Str_Tmp,"")
				end select 		
			end if
		next
		if New_Search_Str<>"" then This_Fun_Sql = and_where(This_Fun_Sql) & replace(New_Search_Str," where ","")
	end if
	if Add_Sql<>"" then This_Fun_Sql = and_where(This_Fun_Sql) &" "& Decrypt(Add_Sql)	
	if orderby<>"" then 
		This_Fun_Sql = This_Fun_Sql &"  Order By "& replace(orderby,"csed"," Desc")
	else
		This_Fun_Sql = This_Fun_Sql &"  Order By AddTime Desc,id desc"
	end if	
	Str_Tmp = "" : ii = 0
	'response.Write(This_Fun_Sql)
	On Error Resume Next
	Set DS_Rs = CreateObject(G_FS_RS)
	DS_Rs.Open This_Fun_Sql,Conn,1,3	
	if Err<>0 then 
		response.Redirect("../error.asp?ErrCodes=<li>��ѯ����������ţ�"&Error.number&"����������"&Error.Description&"</li><li>�����ֶ������Ƿ�ƥ��.</li>")
		response.End()
	end if
	IF DS_Rs.eof THEN
	 	response.Write("<tr class=""hback""><td colspan=15>��������.</td></tr>") 
	else	
	DS_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>DS_Rs.PageCount Then cPageNo=DS_Rs.PageCount 
	DS_Rs.AbsolutePage=cPageNo
	
	  FOR int_Start=1 TO int_RPP 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" style=""cursor:hand"" onclick=""javascript:if(TD_U_"&DS_Rs("ID")&".style.display=='') TD_U_"&DS_Rs("ID")&".style.display='none'; else {TD_U_"&DS_Rs("ID")&".style.display='';ReImgSize('TD_Img_"&DS_Rs("ID")&"');}"" title='����鿴��ϸ���'>"&DS_Rs("ID")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""left""><a href=""DownloadList.asp?Act=Edit&ID="&DS_Rs("ID")&""" title=""����޸Ļ�鿴��ϸ"">"&DS_Rs("Name")&"</a></td>" & vbcrlf
		Str_Tmp = Get_OtherTable_Value("select ClassName as filda from FS_DS_Class where ClassID='"&NoSqlHack(DS_Rs("ClassID"))&"'")
		Get_Html = Get_Html & "<td align=""center""><a href=""DownloadList.asp?Add_Sql="&server.URLEncode(Encrypt("ClassID='"&DS_Rs("ClassID")&"'"))&""">"&Str_Tmp&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&DS_Rs("AddTime")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&DS_Rs("EditTime")&"</td>" & vbcrlf
		if DS_Rs("OverDue")=0 then 
			Str_Tmp = "��������"
		elseif datediff("d",DS_Rs("AddTime"),date())>DS_Rs("OverDue") then 
			Str_Tmp = DS_Rs("OverDue")&"��/�ѹ���"&datediff("d",DS_Rs("AddTime"),date())&"��"
		else 
			Str_Tmp = DS_Rs("OverDue")&"��/"&datediff("d",DS_Rs("AddTime"),date())&"��󽫹���"
		end if
		Get_Html = Get_Html & "<td align=""center"">"&Str_Tmp&"</td>" & vbcrlf 
		if DS_Rs("AuditTF") = 1 then 
			Get_Html = Get_Html & "<td align=""center""><input type=button name=Audited value=""��ͨ��"" onclick=""location='DownloadList.asp?Act=OtherSet&Sql="&server.URLEncode(Encrypt("Update FS_DS_List set AuditTF=0 where ID="&DS_Rs("ID")))&"';""></td>" & vbcrlf
		else
			Get_Html = Get_Html & "<td align=""center""><input type=button name=Audited value=""������"" style=""color:blue"" onclick=""location='DownloadList.asp?Act=OtherSet&Sql="&server.URLEncode(Encrypt("Update FS_DS_List set AuditTF=1 where ID="&DS_Rs("ID")))&"';""></td>" & vbcrlf
		end if
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><a href=""Down_Review.asp?DownID="&DS_Rs("DownLoadID")&""" target=""_blank"">Ԥ��</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" name=""DownLoadID"" id=""DownLoadID"" value="""&DS_Rs("DownLoadID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		''++++++++++++++++++++++++++++++++++++++�㿪ʱ��ʾ��ϸ��Ϣ��
		Get_Html = Get_Html & "<tr class=""hback"" id=""TD_U_"& DS_Rs("ID") &""" style=""display:'none'""><td colspan=20>" & vbcrlf
		db_ii = DS_Rs("Appraise")
		if db_ii = "" or isnull(db_ii) then  db_ii = 0
		if db_ii>6 then db_ii=6
		Str_Tmp = ""
		for ii = 1 to db_ii
			Str_Tmp = Str_Tmp & "<img border=0 src=""../Images/icon_star_2.gif"" title="""&DS_Rs("Appraise")&"��"">"
		next 
		for ii = 1 to 6 - db_ii
			Str_Tmp = Str_Tmp & "<img border=0 src=""../Images/icon_star_1.gif"" title="""&DS_Rs("Appraise")&"��"">"
		next 
		Get_Html = Get_Html & "<table width=""100%"" height=""30"" border=""0"" cellspacing=""1"" cellpadding=""2"" class=""table"">" & vbcrlf 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>������Ȩ:"&Replacestr(DS_Rs("Accredit"),"1:���,2:����,3:����,4:��ʾ,5:ע��,6:�ƽ�,7:����,8:����") & "</td><td>�Ǽ�����:"&Str_Tmp&"</td>"
		Str_Tmp = DS_Rs("BrowPop")
		if isnull(Str_Tmp) or Str_Tmp="" then 
			Str_Tmp = "������"
		else	
			Str_Tmp = Get_OtherTable_Value("select GroupName from FS_ME_Group where GroupID in ("&DS_Rs("BrowPop")&")")
		end if	
		Get_Html = Get_Html	&"<td>����Ȩ��(�û���):"&Str_Tmp&"</td>" & vbcrlf & "<td>��ϵ��EMAIL:"&DS_Rs("EMail")& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>������:"&DS_Rs("Provider")& "</td><td>�ṩ��(��ʾ):<a href="""&DS_Rs("ProviderUrl")&""" target=_blank>"&DS_Rs("ProviderUrl")&"</a></td><td>�ļ���:"&DS_Rs("FileName")&"</td><td>�ļ���չ��:"&Replacestr(DS_Rs("FileExtName"),"else:"&DS_Rs("FileExtName"))&"</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>�ļ���С:"&DS_Rs("FileSize")&"</td><td>����:"&DS_Rs("Language")& "</td><td>ģ���ļ���:"&DS_Rs("NewsTemplet")&"</td><td>��ѹ����:"&DS_Rs("PassWord")&"</td></tr>" & vbcrlf
		Str_Tmp = DS_Rs("Pic")
		if Str_Tmp<>"" then 
			Str_Tmp = "<img id=""TD_Img_"&DS_Rs("ID")&""" src="""&DS_Rs("Pic")&""" border=0>"
		else
			Str_Tmp = "��ͼƬ"
		end if		 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>��ʾͼƬ[<span style=""cursor:hand"" onClick=""if(!$('TD_Img_"&DS_Rs("ID")&"')) alert('û��ͼƬ'); else { if(TD_Img_"&DS_Rs("ID")&".width<=100) {TD_Img_"&DS_Rs("ID")&".width*=5;this.innerHTML='Сͼ'} else {TD_Img_"&DS_Rs("ID")&".width/=5;this.innerHTML='��ͼ'} }"" class=tx>��ͼ</span>] :<br />"&Str_Tmp&"</td><td>��������:"&DS_Rs("Property")&"</td><td>�Ƽ�:"&Replacestr(DS_Rs("RecTF"),"1:<span class=tx>��</span>,0:��")& "</td><td>����:"&Replacestr(DS_Rs("ReviewTF"),"1:<span class=tx>����</span>,0:������")& "</td></tr>"&vbNewLine
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>�����Ƿ������:"&Replacestr(DS_Rs("ShowReviewTF"),"1:<span class=tx>��</span>,0:��")& "</td><td>ϵͳƽ̨:"&DS_Rs("SystemType")&"</td><td>��������:"&Replacestr(DS_Rs("Types"),"1:ͼƬ,2:�ļ�,3:����,4:Flash,5:����,6:Ӱ��,7:����,else:"&DS_Rs("Types")&"")&"</td><td>�汾:"&DS_Rs("Version")& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>��������:"&Replacestr(DS_Rs("OverDue"),"0:��������,"&DS_Rs("OverDue")&":"&DS_Rs("OverDue")&"��")&"</td><td>���ѵ���:"&DS_Rs("ConsumeNum")& "</td><td colspan=3>����ID:<span class=tx>"&DS_Rs("DownLoadID")& "</span></td></tr>" & vbcrlf
		set DS_Rs1=Conn.execute("Select AddressName,Url from FS_DS_Address where DownLoadID='"&DS_Rs("DownLoadID")&"'  order by Number")
		if DS_Rs1.eof then 
			Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td colspan=4 class=tx>û�����ص�ַ����!!</td>" & vbcrlf & "</tr>"& vbcrlf	
		else	
			Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td class=xingmu colspan=2>��������</td><td  class=xingmu colspan=2>���ص�ַ</td></tr>" & vbcrlf
			do while not DS_Rs1.eof
				Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td  colspan=2>"&DS_Rs1("AddressName")&"</td><td colspan=2><a href="""&DS_Rs1("Url")&""" title=""�����������"" target=_blank>"&DS_Rs1("Url")&"</a></td></tr>" & vbcrlf
				DS_Rs1.movenext
			loop
		end if
		DS_Rs1.close	
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td  colspan=10>���:<br />"&DS_Rs("Description")& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "</table>" & vbcrlf
		Get_Html = Get_Html &"</td></tr>" & vbcrlf
		''+++++++++++++++++++++++++++++++++++++++		
		DS_Rs.MoveNext
 		if DS_Rs.eof or DS_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(DS_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""button"" value="" �������� "" onclick=""javascript:if (confirm('ȷ��Ҫȡ����ѡ��Ŀ�������?')) {document.viewform.action='?Act=NoAuditedAll';document.viewform.submit();}""></td>"
	Get_Html = Get_Html & "<td align=right><input type=""button"" value="" ������� "" onclick=""javascript:if (confirm('ȷ��Ҫͨ����ѡ��Ŀ�������?')) {document.viewform.action='?Act=AuditedAll';document.viewform.submit();}""></td>"
	Get_Html = Get_Html & "<td align=right><input type=""button"" value="" ����ɾ�� "" onclick=""javascript:if (confirm('ȷ��Ҫɾ����ѡ��Ŀ��?')) {document.viewform.action='?Act=Del';document.viewform.submit();}""></td>"
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	Get_Html = Get_Html &"</td></tr>"
	DS_Rs.close
	Get_While_Info = Get_Html
End Function

Sub OtherSet(Sql)
	Conn.execute(Decrypt(Sql))
	response.Redirect("DownloadList.asp")
end Sub

Sub NoAuditedAll()
	if not MF_Check_Pop_TF("DS005") then Err_Show
	Dim Str_Tmp
	Str_Tmp = FormatStrArr(request.form("DownLoadID"))
	if Str_Tmp="" then response.Redirect("../error.asp?ErrCodes=<li>���������ѡ��һ��������ˡ�</li>")
	Conn.execute("update FS_DS_List set AuditTF=0 where DownLoadID in ('"&Str_Tmp&"')")
	response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Down/DownloadList.asp?Act=View" )&"&ErrCodes=<li>       ��ϲ��ȡ����˳ɹ���</li>")
End Sub

Sub AuditedAll()
	if not MF_Check_Pop_TF("DS005") then Err_Show
	Dim Str_Tmp
	Str_Tmp = FormatStrArr(request.form("DownLoadID"))
	if Str_Tmp="" then response.Redirect("../error.asp?ErrCodes=<li>���������ѡ��һ��������ˡ�</li>")
	Conn.execute("update FS_DS_List set AuditTF=1 where DownLoadID in ('"&Str_Tmp&"')")
	response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Down/DownloadList.asp?Act=View" )&"&ErrCodes=<li>��ϲ����˳ɹ���</li>")
End Sub

Sub Del()
	if not MF_Check_Pop_TF("DS003") then Err_Show
	Dim Str_Tmp,Str_Tmp_
	if G_IS_SQL_DB=1 then
	Str_Tmp_="datediff(d,AddTime,'"&date()&"') > OverDue"
	else 
	Str_Tmp_="datediff('d',AddTime,'"&date()&"') > OverDue"
	end if
	if request.QueryString("sType") = "All_Over" then 
		Conn.execute("Delete from FS_DS_Address where DownLoadID in (select DownLoadID from FS_DS_List where OverDue>0 and "&Str_Tmp_&")")
		Conn.execute("Delete from FS_DS_List where OverDue>0 and "&Str_Tmp_&"")
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Down/DownloadList.asp?Act=View" )&"&ErrCodes=<li>��ϲ��ɾ���ɹ���</li>")
	end if
	if request.QueryString("DownLoadID")<>"" then 
		Conn.execute("Delete from FS_DS_Address where DownLoadID = '"&NoSqlHack(request.QueryString("DownLoadID"))&"'")
		Conn.execute("Delete from FS_DS_List where DownLoadID = '"&NoSqlHack(request.QueryString("DownLoadID"))&"'")
	else
		Str_Tmp = request.form("DownLoadID")
		if Str_Tmp="" then response.Redirect("../error.asp?ErrCodes=<li>���������ѡ��һ������ɾ����</li>")
		Str_Tmp = replace(Str_Tmp," ","")
		Str_Tmp = replace(Str_Tmp,",","','")
		Conn.execute("Delete from FS_DS_Address where DownLoadID in ('"&Str_Tmp&"')")
		Conn.execute("Delete from FS_DS_List where DownLoadID in ('"&Str_Tmp&"')")
	end if
	response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Down/DownloadList.asp?Act=View" )&"&ErrCodes=<li>��ϲ��ɾ���ɹ���</li>")
End Sub
''================================================================

Sub Save()
	Dim  lng_GroupID1,ConsumeNum,lng_PointNumber1
	lng_GroupID1 = request.form("strBrowPop")
	if right(lng_GroupID1,1)="," then lng_GroupID1=mid(lng_GroupID1,1,len(lng_GroupID1)-1)
	lng_PointNumber1 = request.form("ConsumeNum")
	Dim obj_insert_rs

	Dim Str_Tmp,Arr_Tmp,ID,Req_Other_Set,Errstr,Url,DownLoadID,AddressName,Number,form_ii,SuccessID,FileName
	form_ii = 0 : SuccessID = 0
	ID = NoSqlHack(request.Form("ID"))
	if not isnumeric(ID) or ID = "" then ID = 0 
	DownLoadID = NoSqlHack(request.Form("DownLoadID"))
	FileName  = NoSqlHack(request.Form("FileName"))
	if DownLoadID = "" then
		Dim Check_Down_IDTF,Last_Down_ID_Str,Temp_Down_ID_,Check_Obj_
		Check_Down_IDTF = False
		Do While Not Check_Down_IDTF
			Temp_Down_ID_ = GetRamCode(15)
			Set Check_Obj_ = Conn.ExeCute("select ID from FS_DS_List where DownLoadID='"&NoSqlHack(Temp_Down_ID_)&"'")
			IF Check_Obj_.Bof And Check_Obj_.Eof Then
				DownLoadID = Temp_Down_ID_
				Check_Down_IDTF = True
				Exit Do
			End if
			Check_Obj_.Close : Set Check_Obj_ = Nothing	
		Loop
	End if	
	On Error Resume Next
	for form_ii = 1 to request.Form("Url").Count	
		AddressName = NoSqlHack(request.Form("AddressName")(form_ii))
		Url = NoSqlHack(request.Form("Url")(form_ii))
		Number = NoSqlHack(request.Form("Number")(form_ii))
		if Number = "" or not isnumeric(Number) then Number = form_ii
		if Url<>"" and AddressName<>"" then
			if form_ii = 1 then Conn.execute("delete from FS_DS_Address where  DownLoadID='"&NoSqlHack(DownLoadID)&"'")
			Conn.execute("insert into FS_DS_Address (DownLoadID,AddressName,Url,[Number]) values ('"&NoSqlHack(DownLoadID)&"','"&NoSqlHack(AddressName)&"','"&NoSqlHack(Url)&"',"&CintStr(Number)&")")
		end if
		If Err.Number <> 0 then 
			Err.clear
		else
			SuccessID = SuccessID + 1
		end if	
	next
	if SuccessID=0 then  response.Redirect("../error.asp?ErrCodes=<li>û�����ص�ַ������,���ܼ�����</li>") : response.End()
	''==============================
	if ID=0 then
		Req_Other_Set = "&Accredit="&server.URLEncode(NoSqlHack(request.Form("Accredit")))&"&Appraise="&server.URLEncode(NoSqlHack(request.Form("Appraise")))&"&AuditTF="&server.URLEncode(NoSqlHack(request.Form("AuditTF"))) _
			&"&BrowPop="&server.URLEncode(NoSqlHack(request.Form("BrowPop")))&"&FileExtName="&server.URLEncode(NoSqlHack(request.Form("FileExtName")))&"&FileSize="&server.URLEncode(NoSqlHack(request.Form("FileSize"))) _
			&"&NewsTemplet="&server.URLEncode(NoSqlHack(request.Form("NewsTemplet")))&"&Language="&server.URLEncode(NoSqlHack(request.Form("Language"))) _
			&"&Pic="&server.URLEncode(NoSqlHack(request.Form("Pic")))&"&speicalId="&server.URLEncode(NoSqlHack(request.Form("speicalId")))&"&PassWord="&server.URLEncode(NoSqlHack(request.Form("PassWord")))&"&Provider="&server.URLEncode(NoSqlHack(request.Form("Provider")))&"&ProviderUrl=" _
			&server.URLEncode(NoSqlHack(request.Form("ProviderUrl")))&"&RecTF="&server.URLEncode(NoSqlHack(request.Form("RecTF")))&"&ReviewTF="&server.URLEncode(NoSqlHack(request.Form("ReviewTF")))&"&ShowReviewTF=" _
			&server.URLEncode(NoSqlHack(request.Form("ShowReviewTF")))&"&SystemType="&server.URLEncode(NoSqlHack(request.Form("SystemType")))&"&Types="&server.URLEncode(NoSqlHack(request.Form("Types")))&"&Version=" _
			&server.URLEncode(NoSqlHack(request.Form("Version")))&"&OverDue="&server.URLEncode(NoSqlHack(request.Form("OverDue")))&"&ConsumeNum="&server.URLEncode(NoSqlHack(request.Form("ConsumeNum")))&"&Hits="&server.URLEncode(NoSqlHack(request.Form("Hits")))
	end if		
	Str_Tmp = "ClassID,Description,Accredit,AddTime,Appraise,AuditTF,BrowPop,ClickNum,EditTime,EMail," _
		&"FileSize,[Language],Name,NewsTemplet,PassWord,Pic,Property,Provider,ProviderUrl,RecTF,ReviewTF," _
		&"ShowReviewTF,SystemType,Types,Version,OverDue,ConsumeNum,Hits,FileName,speicalId" 
	Arr_Tmp = split(Str_Tmp,",")
	DS_Sql = "select ID,DownLoadID,SavePath,FileName,FileExtName,"&Str_Tmp&" from FS_DS_List where ID="&CintStr(ID)
	Set DS_Rs = CreateObject(G_FS_RS)
	DS_Rs.Open DS_Sql,Conn,1,3
	if ID > 0 then 
		if not MF_Check_Pop_TF("DS002") then Err_Show
	''�޸�
		DS_Rs("DownLoadID") = DownLoadID
		DS_Rs("SavePath") = SavePath(fileDirRule)
		DS_Rs("FileExtName") = NoSqlHack(request.Form("FileExtName"))
		if instr(FileName,"�Զ����ID")>0 then 

			FileName = replace(FileName,"�Զ����ID",ID)
		end if
		if instr(FileName,"ΨһDownID")>0 then 
			FileName = replace(FileName,"ΨһDownID",DownLoadID)
		end if
		DS_Rs("FileName") = FileName	
		for each Str_Tmp in Arr_Tmp
			DS_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
		next	
		DS_Rs.update
		DS_Rs.close
		
		'����Ȩ�����ݱ�
	'	lng_GroupID1,lng_PointNumber1,flt_Money1
		if Trim(lng_GroupID1) <>"" or lng_PointNumber1 <> "" then 
			set obj_insert_rs = Server.CreateObject(G_FS_RS)
			obj_insert_rs.Open "select  GroupName,PointNumber,FS_Money,InfoID,PopType,isClass From FS_MF_POP  where InfoID='"& NoSqlHack(DownLoadID) &"' and PopType='DS'",Conn,1,3
			obj_insert_rs("InfoID")=DownLoadID
			obj_insert_rs("GroupName")=lng_GroupID1
			if lng_PointNumber1 <>""  then:obj_insert_rs("PointNumber")=lng_PointNumber1:Else:obj_insert_rs("PointNumber")=0:End if
			obj_insert_rs("FS_Money")=0
			obj_insert_rs("PopType")="DS"
			obj_insert_rs("isClass")=0
			obj_insert_rs.update
			obj_insert_rs.close:set obj_insert_rs = nothing
		End if
		Call Refresh("DS_download",ID)
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Down/DownloadList.asp?Act=Edit&ID="&ID )&"&ErrCodes=<li>��ϲ���޸ĳɹ���</li>")
	else
		if not MF_Check_Pop_TF("DS001") then Err_Show
	''����
		if Get_OtherTable_Value("select ID from FS_DS_List where DownLoadID='"&NoSqlHack(request.Form("DownLoadID"))&"'")<>"" then 
			Errstr = "<li>����ID�ظ�!!</li>"
		end if
		if Errstr<>"" then response.Redirect("../error.asp?ErrCodes="&Errstr) : response.End()
		''--------------------------
		DS_Rs.addnew
		DS_Rs("DownLoadID") = DownLoadID
		DS_Rs("SavePath") = SavePath(fileDirRule)		
		DS_Rs("FileExtName") = NoSqlHack(request.Form("FileExtName"))
		for each Str_Tmp in Arr_Tmp
			'response.Write(Str_Tmp&":"&NoSqlHack(request.Form(Str_Tmp))&"<br>")
			if request.Form(Str_Tmp)<>"" then DS_Rs(Str_Tmp) = request.Form(Str_Tmp)
		next
		'response.Write(Req_Other_Set)
		'response.End()	
		DS_Rs.update
		Dim Get_News_ID,rssql 'ȡ�Զ����ID
		if G_IS_SQL_DB = 0 then
			Get_News_ID = DS_Rs("ID")
		Else
			set rssql = Conn.execute("select top 1 id from FS_DS_List order by id asc")
			Get_News_ID = rssql(0)
			rssql.close:set rssql = nothing
		End if
		DS_Rs.close
		if instr(FileName,"�Զ����ID")>0 then 
			FileName = replace(FileName,"�Զ����ID",Conn.execute("select ID from FS_DS_List where DownLoadID='"&NoSqlHack(DownLoadID)&"'")(0))
			Conn.execute("update FS_DS_List set FileName='"&NoSqlHack(FileName)&"' where DownLoadID='"&NoSqlHack(DownLoadID)&"'")
		end if
		
		if instr(FileName,"ΨһDownID")>0 then 
			FileName = replace(FileName,"ΨһDownID",DownLoadID)
			Conn.execute("update FS_DS_List set FileName='"&NoSqlHack(FileName)&"' where DownLoadID='"&NoSqlHack(DownLoadID)&"'")
		end if
	
		'����Ȩ�����ݱ�
	'	lng_GroupID1,lng_PointNumber1,flt_Money1
		if Trim(lng_GroupID1) <>"" or lng_PointNumber1 <> "" then 
			set obj_insert_rs = Server.CreateObject(G_FS_RS)
			obj_insert_rs.Open "select GroupName,PointNumber,FS_Money,InfoID,PopType,isClass From FS_MF_POP",Conn,1,3
			obj_insert_rs.addnew
			obj_insert_rs("InfoID")=NoSqlHack(DownLoadID)
			obj_insert_rs("GroupName")=NoSqlHack(lng_GroupID1)
			if lng_PointNumber1 <>""  then:obj_insert_rs("PointNumber")=NoSqlHack(lng_PointNumber1):Else:obj_insert_rs("PointNumber")=0:End if
			obj_insert_rs("FS_Money")=0
			obj_insert_rs("PopType")="DS"
			obj_insert_rs("isClass")=0
			obj_insert_rs.update
			obj_insert_rs.close
		 end if 
		 
		Call Refresh("DS_download",Get_News_ID)
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Down/DownloadList.asp?Act="&Req_Other_Set ) &"&ErrCodes=<li>��ϲ�������ɹ���</li>")
	end if
End Sub
''=========================================================
classid = Request.QueryString("classid") ' ��classid�������ݽ���
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="js/Public.js"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<script language="JavaScript" src="../../FS_Inc/CheckJs.js"></script>
<script language="JavaScript" src="../../FS_Inc/coolWindowsCalendar.js"></script>
<script language="JavaScript" src="../../FS_Inc/Prototype.js"></script>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Get_Domain.asp"></script>
<script language="JavaScript">

<!--
//�����������
/////////////////////////////////////////////////////////
var Old_Sql = document.URL;

function CheckForm(FormObj)
{
	var nameStr = document.form1.Name.value;
	if (nameStr == '')
	{
		alert('�������Ʋ���Ϊ��!');
		return;
	}
	if (frames["NewsContent"].g_currmode!='EDIT') {alert('����ģʽ���޷����棬���л������ģʽ');return false;}
	FormObj.Description.value=frames["NewsContent"].GetNewsContentArray();
	FormObj.submit();
}



function OrderByName(FildName)
{
	var New_Sql='';
	var oldFildName="";
	if (Old_Sql.indexOf("&filterorderby=")==-1&&Old_Sql.indexOf("?filterorderby=")==-1)
	{
		if (Old_Sql.indexOf("=")>-1)
			New_Sql = Old_Sql+"&filterorderby=" + FildName + "csed";
		else
			New_Sql = Old_Sql+"?filterorderby=" + FildName + "csed";
	}
	else
	{	
		var tmp_arr_ = Old_Sql.split('?')[1].split('&');
		for(var ii=0;ii<tmp_arr_.length;ii++)
		{
			if (tmp_arr_[ii].indexOf("filterorderby=")>-1)
			{
				oldFildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby=") + "filterorderby=".length , tmp_arr_[ii].length);
				break;	
			}
		}
		oldFildName.indexOf("csed")>-1?New_Sql = Old_Sql.replace('='+oldFildName,'='+FildName):New_Sql = Old_Sql.replace('='+oldFildName,'='+FildName+"csed");
	}	
	//alert(New_Sql);
	location = New_Sql;
}
/////////////////////////////////////////////////////////
-->
</script>
<head><body>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr  class="hback">
    <td colspan="10" align="left" class="xingmu" ><a href="#" onMouseOver="this.T_BGCOLOR='#404040';this.T_FONTCOLOR='#FFFFFF';return escape('<div align=\'center\'>FoosunCMS5.0<br>  <BR>Copyright (c) 2006 Foosun Inc</div>')" class="sd"><strong>�����б�����</strong></a></td>
  </tr>
  <tr  class="hback">
    <td colspan="10" height="25"><a href="DownloadList.asp">������ҳ</a>
      <%if MF_Check_Pop_TF("DS001") then %>
      | <a href="DownloadList.asp?Act=Add&classid=<%= classid %>">����</a>
      <%end if%>
      | <a href="DownloadList.asp?Act=Search">��ѯ</a> | ����(<a href="DownloadList.asp?Act=View&Add_Sql=<%=server.URLEncode(Encrypt("DateDiff('d',AddTime,'"&date()&"')<7"))%>">һ����</a> | <a href="DownloadList.asp?Act=View&Add_Sql=<%=server.URLEncode(Encrypt("DateDiff('d',AddTime,'"&date()&"')<1*30"))%>">һ����</a> | <a href="DownloadList.asp?Act=View&Add_Sql=<%=server.URLEncode(Encrypt("DateDiff('d',AddTime,'"&date()&"')<3*30"))%>">������</a>) 
      | ����(<a href="DownloadList.asp?Act=View&Add_Sql=<%=server.URLEncode(Encrypt("OverDue>0 and datediff('d',AddTime,'"&date()&"') > OverDue"))%>">���й�������</a>
      <%if not MF_Check_Pop_TF("DS003") then%>
      | <a href="DownloadList.asp?Act=Del&sType=All_Over" onClick="return confirm('�⽫ɾ�����������Ϣ��ȷ��������');">ɾ�����й���</a>
      <%end if%>
      ) | <a  href="#" onClick="javascirp:history.back()">����</a> </td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="0" class="table">
  <tr class="hback">
    <%
	icNum = 0        '����һ������Ŀ����ʾ����
	if trim(classid)="" then
		Set class_rs=Conn.execute("Select ClassID,ParentID,ClassName from FS_DS_Class where ParentID='0'")
	else
		Set class_rs=Conn.execute("Select ClassID,ParentID,ClassName from FS_DS_Class where parentid='"&NoSqlHack(classid)&"'")'��ʾ����Ŀ
	end if
	dim prefix_img,tmp_rs
	Do while not class_rs.eof 
		Set tmp_rs=Conn.execute("select ClassID from FS_DS_Class where parentid='"&class_rs("ClassID")&"'")
		if not tmp_rs.eof then
			prefix_img="<img src=""../images/+.gif""></img>"
		else
			prefix_img="<img src=""../images/-.gif""></img>"		
		End if
		Response.Write("<td class=""hback"">"&prefix_img&"<a href='DownloadList.asp?classid="&class_rs("ClassID")&"&Add_Sql="&server.URLEncode(Encrypt("classid='"&class_rs("ClassID")&"'"))&"'>"&class_rs("ClassName")&"</a><a href='DownloadList.asp?Act=Add&classid="&class_rs("ClassID")&"'>(<img src=""../images/add.gif"" border=""0"">)</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>")
		class_rs.movenext
		icNum = icNum + 1
		if icNum mod 4 = 0 then
			Response.Write("</tr><tr class=""hback"">")
		End if
	loop
	Dim i
	For i=1 To icNum mod 4
		Response.Write "<td class=""hback""></td>"
	Next
%>
  </tr>
</table>
<%
'******************************************************************
select case request.QueryString("Act")
	case "Add","Edit","Search"
	Add_Edit_Search
	case "SearchGo","","View"
	View
	case "Save"
	Save
	case "Del"
	Del
	case "NoAuditedAll"
	NoAuditedAll
	case "AuditedAll"
	AuditedAll
	case "OtherSet"
	Call OtherSet(request.QueryString("Sql"))
	case else
	response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>����Ĳ������ݡ�</li>") : response.End()
end select
'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="viewform" id="viewform" method="post" action="?Act=Del">
    <tr  class="hback">
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('ID')" class="sd"><b>��ID��</b></a> <span id="Show_Oder_ID"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('Name')" class="sd"><b>��������</b></a> <span id="Show_Oder_Name"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('ClassID')" class="sd"><b>������Ŀ</b></a> <span id="Show_Oder_ClassID"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('AddTime')" class="sd"><b>����ʱ��</b></a> <span id="Show_Oder_AddTime"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('EditTime')" class="sd"><b>�޸�ʱ��</b></a> <span id="Show_Oder_EditTime"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('OverDue')" class="sd"><b>��������</b></a> <span id="Show_Oder_OverDue"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('AuditTF')" class="sd"><b>���</b></a> <span id="Show_Oder_AuditTF"></span></td>
	  <td align="center" class="xingmu">����</td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_While_Info( request.QueryString("Add_Sql"),request.QueryString("filterorderby") ) )
	%>
  </form>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr>
    <td height="18" class="hback"><div align="left">
        <p><span class="tx"><strong>˵��</strong></span>:�������ID������<br>
        </p>
      </div></td>
  </tr>
</table>
<%End Sub
Sub Add_Edit_Search()
Dim Bol_IsEdit,ID,TmpStr,Class_ExtName
Dim Accredit,Appraise,AuditTF,BrowPop,FileExtName,FileSize,NewsTemplet,Pic,PassWord,Provider,ProviderUrl,RecTF,ReviewTF,FileName
Dim ShowReviewTF,SystemType,Types,Version,OverDue,ConsumeNum  ,Language,Hits
Bol_IsEdit = false
TmpStr = ""
if request.QueryString("Act")="Edit" then
	ID = request.QueryString("ID")
	if ID="" then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>��Ҫ��IDû���ṩ��</li>") : response.End()
	DS_Sql = "select ID,DownLoadID,ClassID,Description,Accredit,AddTime,Appraise,AuditTF,BrowPop,ClickNum,EditTime,EMail,FileExtName,FileName," _
		&"FileSize,[Language],Name,NewsTemplet,PassWord,Pic,Property,Provider,ProviderUrl,RecTF,ReviewTF,ShowReviewTF,speicalId,SystemType,Types,Version,OverDue,ConsumeNum,Hits" _
		&" from FS_DS_List where ID = "&CintStr(ID)
	Set DS_Rs	= CreateObject(G_FS_RS)
	DS_Rs.Open DS_Sql,Conn,1,1
	if not DS_Rs.eof then 
	
		Bol_IsEdit = True
		
		Accredit = DS_Rs("Accredit")
		Appraise = DS_Rs("Appraise")
		AuditTF = DS_Rs("AuditTF")
		BrowPop = DS_Rs("BrowPop")
		FileSize = Ucase( DS_Rs("FileSize") )
		FileName = DS_Rs("FileName")
		NewsTemplet = DS_Rs("NewsTemplet")
		PassWord = DS_Rs("PassWord")
		Pic = DS_Rs("Pic")
		Provider = DS_Rs("Provider")
		ProviderUrl = DS_Rs("ProviderUrl")
		RecTF = DS_Rs("RecTF")
		ReviewTF = DS_Rs("ReviewTF")
		ShowReviewTF = DS_Rs("ShowReviewTF")
		SystemType = DS_Rs("SystemType")
		Types = DS_Rs("Types")
		Version = DS_Rs("Version")
		OverDue = DS_Rs("OverDue")
		ConsumeNum = DS_Rs("ConsumeNum")
		if ConsumeNum=0 then ConsumeNum = ""
		Language = DS_Rs("Language")
		NewsTemplet = DS_Rs("NewsTemplet") ''·��
		FileExtName = DS_Rs("FileExtName")''��չ��
		if isnull(FileExtName) or FileExtName="" then FileExtName = "html"
		Hits = DS_Rs("Hits")
		if isnull(Hits) or Hits="" then Hits = 0
	end if	
elseif request.QueryString("Act") = "Add" then
	Accredit = NoSqlHack(request.QueryString("Accredit"))
	if Accredit="" then Accredit = 1
	Appraise = NoSqlHack(request.QueryString("Appraise"))
	if Appraise="" then Appraise = 6	
	AuditTF = NoSqlHack(request.QueryString("AuditTF"))
	if AuditTF="" then AuditTF = 1	
	BrowPop = NoSqlHack(request.QueryString("BrowPop"))
	FileExtName = NoSqlHack(request.QueryString("FileExtName"))
	if FileExtName="" then 
		if sys_FileExtName="" then FileExtName = "html"
	end if	
	FileSize = NoSqlHack(request.QueryString("FileSize"))
	if FileSize="" then FileSize = "1024K"	
	FileName = strFileNameRule(fileNameRule,0,0)
	NewsTemplet = Replace("/"& G_TEMPLETS_DIR &"/Down/Down.htm","//","/") 
	PassWord = NoSqlHack(request.QueryString("PassWord"))
	Pic = NoSqlHack(request.QueryString("Pic"))
	Provider = NoSqlHack(request.QueryString("Provider"))
	ProviderUrl = NoSqlHack(request.QueryString("ProviderUrl"))
	RecTF = NoSqlHack(request.QueryString("RecTF"))
	if RecTF="" then RecTF = 1
	ReviewTF = NoSqlHack(request.QueryString("ReviewTF"))
	if ReviewTF="" then ReviewTF = 1
	ShowReviewTF = NoSqlHack(request.QueryString("ShowReviewTF"))
	if ShowReviewTF="" then ShowReviewTF = 0
	SystemType = NoSqlHack(request.QueryString("SystemType"))
	if SystemType="" then SystemType = "WIN9X/WIN2000/WINXP/2003"
	Types = NoSqlHack(request.QueryString("Types"))
	if Types="" then Types = 3
	Version = NoSqlHack(request.QueryString("Version"))
	OverDue = NoSqlHack(request.QueryString("OverDue"))
	if OverDue="" then OverDue = 0
	ConsumeNum = NoSqlHack(request.QueryString("ConsumeNum"))
	If IsNumeric(ConsumeNum) Then
		ConsumeNum=CInt(ConsumeNum)
	Else
		ConsumeNum=0
	End If
	
	if ConsumeNum=0 then ConsumeNum = ""
	Language  = NoSqlHack(request.QueryString("Language"))
	if Language = "" then Language = "��������"
	Hits = request.QueryString("Hits")
	if Hits = "" then 
		Hits = "0"	
		'randomize		
		'Hits = CStr(Int((9999* Rnd) + 1))
	end if	
end if
%>
<table id=secTable width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="table">
  <tr align=center height=30 class="hback">
    <td id="secTableTd0" class="xingmu" onClick="secBoard(0)" style="cursor:hand"> ������Ϣ </td>
    <td id="secTableTd1" onClick="secBoard(1)" style="cursor:hand"> ���ص�ַ </td>
  </tr>
  <tr>
    <td colspan="4"><!--������ʼ-->
      <form name="form1" id="form1" method="post"<%
  select case request.QueryString("Act")
  	case "Search"
	  response.Write(" action=""?Act=SearchGo""") 
	case "Edit","Add",""
	  response.Write(" onsubmit=""return chkinput();"" action=""?Act=Save""") 
  end select%>>
        <table id=mainTable width="99%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
          <tbody style="display:block;">
            <tr  class="hback">
              <td colspan="3" class="xingmu"><%if Bol_IsEdit then 
	  response.Write("�޸������б���Ϣ<input type=""hidden"" name=""ID"" id=""ID"" value="""&DS_Rs("ID")&""">") 
	  elseif request.QueryString("Act") = "Add" then 
	  response.Write("���������б���Ϣ<span class=""tx"" style=""cursor:help"" onclick=""if (help1.style.display='none') help1.style.display=''; else help1.style.display='none';"">����?</span>") 
	  else
	  response.Write("��ѯ�����б���Ϣ<span class=""tx"" style=""cursor:help"" onclick=""help2.style.display=''?help2.style.display='none':help2.style.display='';"">����?</span>") 
	  end if%>
              </td>
            </tr>
            <tr class="hback" id="help1" style="display:none">
              <td width="15%" align="right">���Ӱ���</td>
              <td width="85%"> ��������������ʱ,ϵͳ���ṩ����Ĭ�ϵ�����.��������һ�������Ժ�,���Ҽ�������ʱ,ϵͳ���Զ�������һ�����ݲ�����Ϣ.������������ʱ�����á����������ѡ�����������Ѿ����ӹ�������. </td>
            </tr>
            <tr class="hback" id="help2" style="display:none">
              <td align="right">��ѯ����</td>
              <td> ��ѯ:���ֺ������͵��ֶ�,֧��<=<>=><>�ȵ����������:���������>2 ; ��������֧�� A B ,A* *B ,*A* *B* ,AB��ģʽ. </td>
            </tr>
            <%if request.QueryString("Act")="Search" then %>
            <tr class="hback">
              <td align="right">�Ա��ID</td>
              <td><input type="text" name="ID" id="ID" size="40" maxlength="11">
              </td>
            </tr>
            <%end if%>
            <tr  class="hback">
              <td width="15%" align="right">����ID</td>
              <td width="85%"><%if request.QueryString("Act")<>"Search" then%>
                <input type="hidden" name="Property" value="0">
                <%end if%>
                <input type="text" size="40" maxlength="15" name="DownLoadID" id="DownLoadID" value="<%if Bol_IsEdit then response.Write(DS_Rs("DownLoadID")) else if request.QueryString("Act")<>"Search" then  response.Write(GetRamCode(15)) end if end if%>" 
	    onFocus="Do.these('DownLoadID',function(){return isEmpty('DownLoadID','DownLoadID_Alt')})" onKeyUp="Do.these('DownLoadID',function(){return isEmpty('DownLoadID','DownLoadID_Alt')});"
	    <%if request.QueryString("Act")="Add" then%> onBlur="if(this.value!='') new Ajax.Updater('DownLoadID_Chk','DownloadList_Ajax.asp?no-cache='+Math.random() , {method: 'get', parameters: 'Act=Check&stype=DownLoadID&value='+this.value });"<%end if%>>
                <span id="DownLoadID_Alt"></span>&nbsp;<span id="DownLoadID_Chk"></span>&nbsp;
                <%if request.QueryString("Act")<>"Search" then %>
                <span class="tx">�������Զ�����</span>
                <%end if%>
              </td>
            </tr>
            <tr>
              <td class="hback" align="right">ѡ����Ŀ</td>
              <td class="hback"><input name="ClassName" type="text" id="ClassName" style="width:50%" value="<%if Bol_IsEdit then response.Write(Get_OtherTable_Value("select ClassName from FS_DS_Class where ClassID='"&DS_Rs("ClassID")&"'")) else if request.QueryString("Act")<>"Search" then response.Write(Get_OtherTable_Value("select ClassName from FS_DS_Class where ClassID='"&request.QueryString("classid")&"'")) end if  end if%>" readonly 
	  onFocus="Do.these('ClassName',function(){return isEmpty('ClassName','ClassName_Alt')});if(this.value!='') new Ajax.Updater('Class_ExtName','DownloadList_Ajax.asp?no-cache='+Math.random() , {method: 'get', parameters: 'Act=GetExtName&ClassID='+ClassID.value });" 
	  onKeyUp="Do.these('ClassName',function(){return isEmpty('ClassName','ClassName_Alt')})" onBlur="if(Class_ExtName.innerHTML!='') FileExtName.value=$('Class_ExtName').innerHTML;">
                <input name="ClassID" type="hidden" id="ClassID" value="<%if Bol_IsEdit then response.Write(DS_Rs("ClassID")) else response.Write(request.QueryString("classid")) end if%>">
                <input type="button" name="Submit" value="ѡ����Ŀ"   onClick="SelectClass();ClassName.focus();">
                <span id="span_ClassName"></span> <span id="ClassName_Alt"></span><span id="Class_ExtName" style="display:none"></span></td>
            </tr>
            <tr>
              <td class="hback" align="right">����ר����</td>
              <%
			  if Bol_IsEdit then
				  if DS_RS("speicalId")>0 then
					  dim sRs,spname
					  Set sRs = Conn.execute("selecT SpecialCName From FS_DS_Special where SpecialID="& DS_RS("speicalId"))
					  if not sRs.eof then
						  spname = sRs(0)
					  end if
					  sRs.close:set sRs = nothing
				  end if
			  end if
			  %>
              <td class="hback"><input name="txt_specialname" type="text" id="txt_specialname" size="30" value="<% = spname %>" readonly="">
                <span id="alert_specialname"></span>
                <input name="speicalId" type="hidden" id="speicalId" value="<%if Bol_IsEdit then response.Write DS_RS("speicalId")%>">
                <button onClick="SelectSpecail();">ѡ��ר��</button></td>
            </tr>
            <tr  class="hback">
              <td align="right">��������</td>
              <td><input type="text" name="Name" id="Name" size="40" maxlength="100" value="<%if Bol_IsEdit then response.Write(DS_Rs("Name")) end if%>" 
	  onFocus="Do.these('Name',function(){return isEmpty('Name','Name_Alt')})" onKeyUp="Do.these('Name',function(){return isEmpty('Name','Name_Alt')});"
	  <%if request.QueryString("Act")="Add" then%> onBlur="if(ClassID.value!='' && this.value!='') new Ajax.Updater('Name_Chk','DownloadList_Ajax.asp?no-cache='+Math.random() , {method: 'get', parameters: 'Act=Check&stype=downname&classid='+ClassID.value+'&name='+this.value });"<%end if%>>
                <span id="Name_Alt"></span><span id="Name_Chk"></span> </td>
            </tr>
            <tr  class="hback">
              <td align="right">���ؼ��</td>
              <td colspan="3" class="hback">
			  <%
				Dim testdown
				If Bol_IsEdit Then
					If DS_Rs("Description")<>"" And Not IsNull(DS_Rs("Description")) Then
						testdown = DS_Rs("Description")
					else
						testdown = ""
					End If 
				End If 				
				%>
			  <!--�༭����ʼ-->
				<iframe id='NewsContent' src='../Editer/AdminEditer.asp?id=Description' frameborder=0 scrolling=no width='100%' height='440'></iframe>
				<input type="hidden" name="Description" value="<% = HandleEditorContent(testdown & "") %>">
                <!--�༭������-->
              </td>
            </tr>
            <!--�Զ����Զο�ʼ-->
            <%If IsArray(CustColumnArr) Then
			response.Write"<tr><td colspan=""4"" class=""hback_1""><b>�Զ��忪ʼ</b></td></tr>"
			Dim InputModeStr,AuxiInfoList,AuxiListArr,k,tmp_k,i,tmp_nulls_span,tmp_nulls
			For i = 0 to Ubound(CustColumnArr,2)
				if CustColumnArr(5,i)=0 then
					tmp_nulls="onFocus=""Do.these('FS_DS_Define_"&CustColumnArr(3,i)&"',function(){return isEmpty('FS_DS_Define_"&CustColumnArr(3,i)&"','span_FS_DS_Define_"&CustColumnArr(3,i)&"')})"" onKeyUp=""Do.these('FS_DS_Define_"&CustColumnArr(3,i)&"',function(){return isEmpty('FS_DS_Define_"&CustColumnArr(3,i)&"','span_FS_DS_Define_"&CustColumnArr(3,i)&"')})"""
					tmp_nulls_span="&nbsp;<span id=""span_FS_DS_Define_"&CustColumnArr(3,i)&"""></span>"
				else
					tmp_nulls=""
					tmp_nulls_span=""
				end if
				Select Case CStr(CustColumnArr(4,i))	'����ѡ������������������뷽ʽ
					Case 0
							Response.Write "<tr ><td class=""hback"" align=""right"">"&CustColumnArr(2,i)&"</td><td colspan=""3"" class=""hback""><input name=""FS_DS_Define_"&CustColumnArr(3,i)&""" type=""test"" style=""width:70%""  value="""&CustColumnArr(6,i)&""" "& tmp_nulls &">"& tmp_nulls_span &"&nbsp;<span class=""tx"">"&CustColumnArr(7,i)&"</span></td></tr>"&vbcrlf
					case 1
							Response.Write "<tr ><td class=""hback"" align=""right"">"&CustColumnArr(2,i)&"</td><td colspan=""3"" class=""hback""><textarea rows=""4"" name=""FS_DS_Define_"&CustColumnArr(3,i)&""" style=""width:70%"" "& tmp_nulls &">"&CustColumnArr(6,i)&"</textarea>"& tmp_nulls_span &"&nbsp;<span class=""tx"">"&CustColumnArr(7,i)&"</span></td></tr>"&vbcrlf
					Case 4
							Response.Write "<tr ><td class=""hback"" align=""right"">"&CustColumnArr(2,i)&"</td><td colspan=""3"" class=""hback""><Select Name=""FS_DS_Define_"&CustColumnArr(3,i)&""" style=""width:70%"">"&vbcrlf
							AuxiListArr=Split(CustColumnArr(6,i),vbcrlf)
							For tmp_k = 0 to UBound(AuxiListArr)	'�������ֵ��ѡ����Ϣ
								If AuxiListArr(tmp_k)<>"" Then 
									Response.Write"<Option value="""&AuxiListArr(tmp_k)&""">"&AuxiListArr(tmp_k)&"</option>"&vbcrlf
								End if
							Next
							Response.Write "</Select>&nbsp;<span class=""tx"">"&CustColumnArr(7,i)&"</span></td></tr>"&vbcrlf
					Case Else	'���У����֣�����
							Response.Write "<tr ><td class=""hback"" align=""right"">"&CustColumnArr(2,i)&"</td><td colspan=""3"" class=""hback""><input name=""FS_DS_Define_"&CustColumnArr(3,i)&""" type=""test"" style=""width:70%""  value="""&CustColumnArr(6,i)&""" "& tmp_nulls &">"& tmp_nulls_span &"&nbsp;<span class=""tx"">"&CustColumnArr(7,i)&"</span></td></tr>"&vbcrlf
				End Select
			Next
			response.Write"<tr><td colspan=""4"" class=""hback_1""><b>�Զ������</b></td></tr>"
		End If
	%>
            <!--�Զ����Զν���-->
            <tr  class="hback">
              <td colspan="2" class="xingmu" height="5"></td>
            </tr>
            <tr  class="hback">
              <td align="right">������Ȩ</td>
              <td><select name="Accredit" id="Accredit">
                  <%=PrintOption(Accredit,":��ѡ��,1:���,2:����,3:����,4:��ʾ,5:ע��,6:�ƽ�,7:����,8:����")%>
                </select>
                <span id="Accredit_Alt"></span> </td>
            </tr>
            <tr class="hback">
              <td align="right">�Ǽ�����</td>
              <td><select name="Appraise" id="Appraise">
                  <%=PrintOption(Appraise,":��ѡ��,1:һ��,2:����,3:����,4:����,5:����,6:����")%>
                </select>
                <span id="Appraise_Alt"></span> </td>
            </tr>
            <tr  class="hback">
              <td align="right">��ϵ��EMAIL</td>
              <td><input type="text" name="EMail" id="EMail" size="40" maxlength="25" value="<%if Bol_IsEdit then response.Write(DS_Rs("EMail")) end if%>">
                <span id="EMail_Alt"></span> </td>
            </tr>
            <tr  class="hback">
              <td align="right">�ļ���С</td>
              <td><input type="text" name="FileSize" id="FileSize" size="15" maxlength="50" value="<%=FileSize%>">
                <span id="FileSize_Alt"></span>
                <select onChange="FileSize.value=this.options[this.selectedIndex].value">
                  <option value="">��ѡ��</option>
                  <%=Get_FildValue_List("select distinct FileSize from FS_DS_List where FileSize<>''",FileSize,1)%>
                </select>
              </td>
            </tr>
            <tr  class="hback">
              <td align="right">����</td>
              <td><input type="text" name="Language" id="Language" size="15" maxlength="50" value="<%=Language%>">
                <span id="Language_Alt"></span>
                <select onChange="Language.value=this.options[this.selectedIndex].value">
                  <option value="">��ѡ��</option>
                  <%if Language<>"" then response.Write PrintOption(Language,Language&":"&Language)%>
                </select>
              </td>
            </tr>
            <tr  class="hback">
              <td align="right">��ѹ����</td>
              <td><input type="text" name="PassWord" id="PassWord" size="40" maxlength="50" value="<%=PassWord%>">
                <select onChange="PassWord.value=this.options[this.selectedIndex].value">
                  <option value="">��ѡ��</option>
                  <%=Get_FildValue_List("select distinct PassWord from FS_DS_List where PassWord<>'' ",PassWord,1)%>
                </select>
              </td>
            </tr>
            <tr  class="hback">
              <td align="right">ѡ��ģ��</td>
              <td><input name="NewsTemplet" type="text" id="NewsTemplet" style="width:60%" value="<%if request.QueryString("Act")="Edit" then Response.Write(NewsTemplet) else Response.Write(tmp_sTemplets)%>" maxlength="200" readonly>
                <input name="Submit5" type="button" id="selNewsTemplet" value="ѡ��ģ��"  onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir %>/<% = G_TEMPLETS_DIR %>',400,300,window,document.form1.NewsTemplet);document.form1.NewsTemplet.focus();">
              </td>
            </tr>
            <tr  class="hback">
              <td align="right">��ʾͼƬ</td>
              <td><input type="text" name="Pic" id="Pic" style="width:60%" maxlength="100" value="<%=Pic%>" readonly="">
                <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.form1.Pic);">
              </td>
            </tr>
            <tr  class="hback">
              <td align="right">������</td>
              <td><input type="text" name="Provider" id="Provider" size="40" maxlength="50" value="<%=Provider%>">
                <select onChange="Provider.value=this.options[this.selectedIndex].value">
                  <option value="">��ѡ��</option>
                  <%=Get_FildValue_List("select distinct Provider from FS_DS_List where Provider<>''",Provider,1)%>
                </select>
              </td>
            </tr>
            <tr  class="hback">
              <td align="right">�ṩ��Url��ַ(��ʾ��ַ)</td>
              <td><input type="text" name="ProviderUrl" id="ProviderUrl" size="40" maxlength="100" value="<%=ProviderUrl%>">
                <select onChange="ProviderUrl.value=this.options[this.selectedIndex].value">
                  <option value="">��ѡ��</option>
                  <%=Get_FildValue_List("select distinct ProviderUrl from FS_DS_List where ProviderUrl<>''",ProviderUrl,1)%>
                </select>
              </td>
            </tr>
            <tr  class="hback">
              <td align="right">ϵͳƽ̨</td>
              <td><input type="text" name="SystemType" id="SystemType" size="40" maxlength="100" value="<%=SystemType%>">
                <select onChange="SystemType.value=this.options[this.selectedIndex].value">
                  <option value="">��ѡ��</option>
                  <%=Get_FildValue_List("select distinct SystemType from FS_DS_List where SystemType<>''",SystemType,1)%>
                </select>
              </td>
            </tr>
            <tr  class="hback">
              <td align="right">��������</td>
              <td><select name="Types" id="Types">
                  <%=PrintOption(Types,":��ѡ��,1:ͼƬ,2:�ļ�,3:����,4:Flash,5:����,6:Ӱ��,7:����")%>
                </select>
              </td>
            </tr>
            <tr  class="hback">
              <td align="right">�汾</td>
              <td><input type="text" name="Version" id="Version" size="40" maxlength="50" value="<%=Version%>">
                <select onChange="Version.value=this.options[this.selectedIndex].value">
                  <option value="">��ѡ��</option>
                  <%=Get_FildValue_List("select distinct Version from FS_DS_List where Version<>''",Version,1)%>
                </select>
              </td>
            </tr>
            <tr  class="hback">
              <td colspan="2" class="xingmu" height="5"></td>
            </tr>
            <tr class="hback">
              <td align="right">�Ƿ����</td>
              <td><select name="AuditTF" id="AuditTF">
                  <%=PrintOption(AuditTF,":��ѡ��,1:ͨ��,0:����")%>
                </select>
                <span id="AuditTF_Alt"></span> </td>
            </tr>
            <tr class="hback">
              <td align="right">�Ƽ�</td>
              <td><select name="RecTF" id="RecTF">
                  <%=PrintOption(RecTF,":��ѡ��,1:��,0:��")%>
                </select>
                <span id="RecTF_Alt"></span> </td>
            </tr>
            <tr class="hback">
              <td align="right">����</td>
              <td><select name="ReviewTF" id="ReviewTF">
                  <%=PrintOption(ReviewTF,":��ѡ��,1:����,0:������")%>
                </select>
                <span id="ReviewTF_Alt"></span> </td>
            </tr>
            <tr class="hback">
              <td align="right">�����Ƿ������</td>
              <td><select name="ShowReviewTF" id="ShowReviewTF">
                  <%=PrintOption(ShowReviewTF,":��ѡ��,1:��,0:��")%>
                </select>
                <span id="ShowReviewTF_Alt"></span> </td>
            </tr>
            <tr  class="hback">
              <td colspan="2" class="xingmu" height="5"></td>
            </tr>
            <!--���ɿ�ʼ-->
            <tr class="hback">
              <td align="right">����Ȩ��</td>
              <td><input name="strBrowPop"  id="strBrowPop" value="<%if BrowPop<>"" then response.Write(Get_OtherTable_Value("select GroupName from FS_ME_Group where GroupID in ("&BrowPop&")")) end if%>" type="text" onMouseOver="this.title=this.value;" readonly>
                <input name="BrowPop"  id="BrowPop" value="<%=BrowPop%>" type="hidden">
                <select name="selectPop" id="selectPop" style="overflow:hidden;" onChange="ChooseExeName();">
                  <option value="" selected>ѡ���Ա��</option>
                  <option value="del" style="color:red;">���</option>
                  <%=Get_FildValue_List("select distinct GroupID,GroupName from FS_ME_Group",BrowPop,1)%>
                </select>
              </td>
            </tr>
            <tr  class="hback">
              <td align="right">���ѵ���</td>
              <td><input type="text" name="ConsumeNum" id="ConsumeNum" size="10" maxlength="6" value="<%=ConsumeNum%>"  onChange="ChooseExeName();">
                <select onChange="ConsumeNum.value=this.options[this.selectedIndex].value;ChooseExeName();">
                  <option value="">��ѡ��</option>
                  <%=Get_FildValue_List("select distinct ConsumeNum from FS_DS_List where ConsumeNum<>0",ConsumeNum,1)%>
                </select>
              </td>
            </tr>
            <tr  class="hback">
              <td align="right">�ļ���չ��</td>
              <td class="hback"><select name="FileExtName" id="FileExtName">
                  <%=PrintOption(FileExtName,":��ѡ��,html:html,htm:htm,shtml:shtml,shtm:shtm,asp:asp")%>
                </select>
                <span id="FileExtName_Alt"></span> </td>
            </tr>
            <tr  class="hback">
              <td align="right">�ļ���</td>
              <td><%
	  if request.QueryString("Act")="Search" then 
	 	 Response.Write("<input name=""FileName"" type=""text"" id=""FileName"" size=""40"" maxlength=""255"">")
	  else
		  Dim RoTF
		  if instr(strFileNameRule(fileNameRule,0,0),"�Զ����ID")>0 or instr(strFileNameRule(fileNameRule,0,0),"ΨһDownID") then:RoTF="Readonly":End if
		  Response.Write("<input name=""FileName"" type=""text"" id=""FileName"" size=""40"" "& RoTF &" maxlength=""255"" value="""&FileName&""" title=""��������������趨Ϊ�Զ���ţ��������޸�"">")
	  end if
	  %>
                <span id="FileName_Alt"></span> </td>
            </tr>
            <tr  class="hback">
              <td align="right">��������</td>
              <td><input type="text" name="OverDue" id="OverDue" size="10" maxlength="6" value="<%=OverDue%>" onChange="ChooseExeName();">
                <select onChange="OverDue.value=this.options[this.selectedIndex].value;ChooseExeName();">
                  <option value="">��ѡ��</option>
                  <%if OverDue = "0" then 
			response.Write("<option value=""0"" selected style=""color:blue"">��������</option>")
		else
			response.Write(Get_FildValue_List("select distinct OverDue from FS_DS_List",OverDue,1))
		end if
		%>
                </select>
              </td>
            </tr>
            <!--���ɽ���-->
            <tr  class="hback">
              <td align="right">�����</td>
              <td><input type="text" name="Hits" id="Hits" size="10" maxlength="6" value="<%=Hits%>">
              </td>
            </tr>
            <tr  class="hback">
              <td align="right">���ش���</td>
              <td><input type="text" name="ClickNum" id="ClickNum" size="10" maxlength="6" value="<%if Bol_IsEdit then response.Write(DS_Rs("ClickNum")) else if request.QueryString("Act")<>"Search" then response.Write("0") end if end if%>">
              </td>
            </tr>
            <tr class="hback">
              <td align="right">����ʱ��</td>
              <td><input type="text" name="AddTime" id="AddTime" onFocus="setday(this)" style="WIDTH: 100px; HEIGHT: 22px" maskType="shortDate" value="<%if Bol_IsEdit then response.Write(DS_Rs("AddTime")) else if request.QueryString("Act")<>"Search" then response.Write(date()) end if end if%>">
                <IMG onClick="AddTime.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom"><span id="AddTime_Alt"></span> </td>
            </tr>
            <tr class="hback">
              <td align="right">�޸�ʱ��</td>
              <td><input type="text" name="EditTime" id="EditTime" onFocus="setday(this)" style="WIDTH: 100px; HEIGHT: 22px" maskType="shortDate" value="<%if Bol_IsEdit then response.Write(date()) else if request.QueryString("Act")<>"Search" then response.Write(date()) end if end if%>">
                <IMG onClick="EditTime.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom"><span id="EditTime_Alt"></span> </td>
            </tr>
            <tr  class="hback">
              <td align="center" colspan="4"><input type="button" onClick="CheckForm(this.form);" value=" ȷ���ύ "/>
                &nbsp;
                <input type="reset" value=" ���� " />
                &nbsp;
                <input type="button" name="btn_todel" value=" ɾ�� " onClick="if(confirm('ȷ��ɾ������Ŀ��')) location='<%=server.URLEncode("DownloadList.asp?Act=Del&ID="&ID)%>'">
              </td>
            </tr>
          </tbody>
          <!--�ֽ�-->
          <tbody style="display:none;">
            <tr  class="hback">
              <td colspan="3" class="xingmu"><%if Bol_IsEdit then 
	  response.Write("�޸������б���Ϣ") 
	  elseif request.QueryString("Act") = "Add" then 
	  response.Write("���������б���Ϣ<span class=""tx"" style=""cursor:help"" onclick=""if (help3.style.display='none') help3.style.display=''; else help3.style.display='none';"">����?</span>") 
	  else
	  response.Write("��ѯ�����б���Ϣ<span class=""tx"" style=""cursor:help"" onclick=""help4.style.display=''?help4.style.display='none':help4.style.display='';"">����?</span>") 
	  end if%>
              </td>
            </tr>
            <tr class="hback" id="help3" style="display:none">
              <td align="right">���Ӱ���</td>
              <td> ��������������ʱ,ϵͳ���ṩ����Ĭ�ϵ�����.��������һ�������Ժ�,���Ҽ�������ʱ,ϵͳ���Զ�������һ�����ݲ�����Ϣ.������������ʱ�����á����������ѡ�����������Ѿ����ӹ�������. </td>
            </tr>
            <tr class="hback" id="help4" style="display:none">
              <td align="right">��ѯ����</td>
              <td> ��ѯ:���ֺ������͵��ֶ�,֧��<=<>=><>�ȵ����������:���������>2 ; ��������֧�� A B ,A* *B ,*A* *B* ,AB��ģʽ. </td>
            </tr>
            <%if request.QueryString("Act")="Search" then %>
            <tr class="hback">
              <td align="right">�Ա��ID</td>
              <td><input type="text" name="AddrID" id="AddrID" size="40" maxlength="11">
              </td>
            </tr>
            <%end if%>
            <tr  class="hback">
              <td id="Ajax_AddrInfo" colspan="4"><%if Bol_IsEdit then Call Edit_AddrList(DS_Rs("DownLoadID")) end if%>
              </td>
            </tr>
            <tr  class="hback">
              <td align="right">��������</td>
              <td><input name="FilesNum" type="text" value="1" size="10" maxlength="2" onBlur="Do.these('FilesNum',function(){return isNumber('FilesNum','FilesNum_Alt','����������!',true)})" onKeyUp="Do.these('FilesNum',function(){return isNumber('FilesNum','FilesNum_Alt','����������!',true)})">
                <input type="button" name="button42" value="�趨" onClick="ChooseOption();">
                <span id="FilesNum_Alt"></span> <span class="tx">�����趨�������ص�ַ</span></td>
            </tr>
            <tr  class="hback">
              <td colspan="4" id="OtherInput"></td>
            </tr>
            <tr  class="hback">
              <td align="center" colspan="4"><input type="button" onClick="CheckForm(this.form);" value=" ȷ���ύ "/>
                &nbsp;
                <input type="reset" value=" ���� " />
                &nbsp;
                <input type="button" name="btn_todel" value=" ɾ�� " onClick="if(confirm('ȷ��ɾ������Ŀ��')) location='<%=server.URLEncode("DownloadList.asp?Act=Del&DownLoadID=")%>'+DownLoadID.value">
              </td>
            </tr>
          </tbody>
        </table>
      </form>
      <!--��������-->
    </td>
  </tr>
</table>
<%
End Sub

Sub Edit_AddrList(DownID)
Dim rowii
rowii = 0
if DownID<>"" then
	DS_Sql1 = "select ID,AddressName,Url,Number from FS_DS_Address where DownLoadID = '"&NoSqlHack(DownID)&"' order by Number asc"
	set DS_Rs1 = Conn.execute(DS_Sql1)
	response.Write("<table border=""0"" width=""100%"" cellpadding=""3"" cellspacing=""1"" class=""table"">"&vbcrlf)
	do while not DS_Rs1.eof 
	rowii = rowii + 1%>
<tr  class="hback">
  <td align="right">���ص�ַ����</td>
  <td><input type="text" size="40" maxlength="50" name="AddressName" id="AddressName" value="<%=DS_Rs1("AddressName")%>"
	    onFocus="Do.these('AddressName',function(){return isEmpty('AddressName','AddressName_Alt<%=rowii%>')})" onKeyUp="Do.these('AddressName',function(){return isEmpty('AddressName','AddressName_Alt<%=rowii%>')})">
    <span id="AddressName_Alt<%=rowii%>"></span> </td>
</tr>
<tr>
  <td class="hback" align="right">���ص�ַ</td>
  <td colspan="3" class="hback"><input name="Url" type="text" id="Url" maxlength="100" style="width:50%" value="<%=DS_Rs1("Url")%>">
    <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ���ļ�" onClick="SelectFile();">
    <span id="Url_Alt"></span> </td>
</tr>
<tr  class="hback">
  <td align="right">���ص�ַ����</td>
  <td><input type="text" name="Number" id="Number" size="10" maxlength="1" value="<%=DS_Rs1("Number")%>">
    <%if rowii>1 then%>
    <input type="button" class="tx" value="ɾ����������" onClick="if(confirm('ȷ��ɾ������������')) {new Ajax.Updater('Ajax_AddrInfo','DownloadList_Ajax.asp?no-cache='+Math.random() , {method: 'get', parameters: 'Act=DelAddr&DownLoadID=<%=DownID%>&AddrID=<%=DS_Rs1("ID")%>' });disabled=true;}">
    <%end if%>
    <span class=tx>Ĭ�����������գ�</span></td>
</tr>
<%	
		DS_Rs1.movenext
	loop
	response.Write("</table>")
	DS_Rs1.close
end if

end Sub

set DS_Rs = Nothing
User_Conn.close
Set class_rs=nothing
Conn.close


%>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
<script language="JavaScript">
<!--//�жϺ���������.�ֶ���������ʾָʾ
//�򿪺���ݹ�����ʾ��ͷ
var Req_FildName;
if (Old_Sql.indexOf("filterorderby=")>-1)
{
	var tmp_arr_ = Old_Sql.split('?')[1].split('&');
	for(var ii=0;ii<tmp_arr_.length;ii++)
	{
		if (tmp_arr_[ii].indexOf("filterorderby=")>-1)
		{
			if(Old_Sql.indexOf("csed")>-1)
				{Req_FildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby=") + "filterorderby=".length , tmp_arr_[ii].indexOf("csed"));break;}
			else
				{Req_FildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby=") + "filterorderby=".length , tmp_arr_[ii].length);break;}	
		}
	}	
	if (document.getElementById('Show_Oder_'+Req_FildName)!=null)  
	{
		if(Old_Sql.indexOf(Req_FildName + "csed")>-1)
		{
			eval('Show_Oder_'+Req_FildName).innerText = '��';
		}
		else
		{
			eval('Show_Oder_'+Req_FildName).innerText = '��';
		}
	}	
}
/////////////////////////////////////////////////////////
function chkinput()
{
	var mainb=isEmpty('Name','Name_Alt') &&  isEmpty('ClassID','ClassName_Alt') && isEmpty('Description','Description_Alt') && isEmpty('FileExtName','FileExtName_Alt') ;
	if (mainb==true)
	{
		if (document.getElementById('AddressName')==null)
		{
			mainTable.tBodies[0].style.display="none";
			mainTable.tBodies[1].style.display="";
			alert('�������������ص�ַ��');
			ChooseOption();
			return false;
		}
		else if (document.getElementById('AddressName').value=='')
			{
				mainTable.tBodies[0].style.display="none";
				mainTable.tBodies[1].style.display="";
				alert('���������ص�ַ��');
				document.getElementById('AddressName').focus();
				return false;
			}
		else         
		return mainb;
	}
	return mainb;
}

function secBoard(n)
{  
  for(i=0;i<mainTable.tBodies.length;i++)
  {
  	mainTable.tBodies[i].style.display="none"; 
	document.getElementById('secTableTd'+i).className = ''; 	
  }
  event.srcElement.className='xingmu';
  mainTable.tBodies[n].style.display="block";
}

function SelectClass()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('lib/SelectClassFrame.asp',400,300,window);
	if (ReturnValue.indexOf('***')!=-1)
	{
		TempArray = ReturnValue.split('***');
		document.form1.ClassID.value=TempArray[0]
		document.form1.ClassName.value=TempArray[1]
	}
}

<%if request.QueryString("Act")="Add" or request.QueryString("Act")="Edit" then%>
function document.onreadystatechange()
{
	ChooseExeName();
}
<%end if%>
function ChooseExeName()
{
  var ObjValue = document.form1.selectPop.options[document.form1.selectPop.selectedIndex].value;
  var Objtext = document.form1.selectPop.options[document.form1.selectPop.selectedIndex].text;
  if (ObjValue!='')
  {
	if (document.form1.BrowPop.value=='')
		{document.form1.BrowPop.value = ObjValue;
		document.form1.strBrowPop.value = Objtext;}
	else if(document.form1.BrowPop.value.indexOf(ObjValue)==-1)
		{document.form1.BrowPop.value = document.form1.BrowPop.value+","+ObjValue;
		document.form1.strBrowPop.value = document.form1.strBrowPop.value+","+Objtext;}
	if (ObjValue=='del')
  		{document.form1.BrowPop.value ='';document.form1.strBrowPop.value ='';}
  }
   CheckNumber(document.form1.ConsumeNum,"����۵�ֵ");
  if (document.form1.ConsumeNum.value>32767||document.form1.ConsumeNum.value<-32768||document.form1.ConsumeNum.value=='0')
	{
		alert('����۵�ֵ����������Χ��\n���32767���Ҳ���Ϊ0');
		document.form1.ConsumeNum.value='';
		document.form1.ConsumeNum.focus();
	}
  if (isNaN(document.form1.OverDue.value)||document.form1.OverDue.value>32767||document.form1.OverDue.value<0)
	{
		alert('��������������ڵ���0');
		document.form1.OverDue.value='';
		document.form1.OverDue.focus();
	}
  if (document.form1.BrowPop.value!=''||document.form1.ConsumeNum.value!=''||(document.form1.OverDue.value!=''&&document.form1.OverDue.value!='0')){document.form1.FileExtName.options[5].selected=true;document.form1.FileExtName.readonly=true;}
  else {document.form1.FileExtName.readonly=false;}
}
function CheckFileExtName(Obj)
{
	if (Obj.value!='')
	{
		for (var i=0;i<document.all.FileExtName.length;i++)
		{
			if (document.all.FileExtName.options(i).value=='asp') document.all.FileExtName.options(i).selected=true;
		}
		document.all.FileExtName.readonly=true;
	}
	else
	{
		document.all.FileExtName.readonly=false;
	}
	Obj.Description.value=frames["Description"].GetNewsContentArray();
	Obj.submit();
}


function ChooseOption()
 {
  var FilesNum = document.all.FilesNum.value;
  if (FilesNum=='')
  	FilesNum=1;
  if (!isNumber('FilesNum','FilesNum_Alt','����������!',true)) {document.all.FilesNum.value='1';FilesNum=1;}		
  var i,Optionstr;
  Optionstr = '<table border="0" width="100%" cellpadding="3" cellspacing="1" class="table">';
  for (i=1;i<=FilesNum;i++)
      {
	   Optionstr += '    <tr  class="hback">\n' 
       Optionstr += '<td align="right">���ص�ַ����</td>\n'
       Optionstr += '<td>\n'
	   Optionstr += '<input type="text" size="40" maxlength="50" name="AddressName" id="AddressName" value="" '
	  //Optionstr += "onFocus=\"Do.these('AddressName',function(){return isEmpty('AddressName','AddressName_Alt"+i+"')})\" onKeyUp=\"Do.these('AddressName',function(){return isEmpty('AddressName','AddressName_Alt"+i+"')})\" "
	   Optionstr += "onBlur=\"if(this.value!='') new Ajax.Updater('Name_Chk"+i+"','DownloadList_Ajax.asp?no-cache='+Math.random() , {method: 'get', parameters: 'Act=Check&stype=addrname&value='+this.value });\">\n"
	   Optionstr += '<span id="Name_Chk'+i+'"></span>\n'
	   Optionstr += '</td>\n'
       Optionstr += '</tr>\n'
       Optionstr += '<tr>\n'
       Optionstr += '<td class="hback" align="right">���ص�ַ</td>\n'
       Optionstr += '<td colspan="3" class="hback" id="td_Url"+i>\n'
	   Optionstr += '<input name="Url" type="text" id="Url" style="width:50%"  maxlength="100" value="">\n' 
	   Optionstr += '<input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ���ļ�" onClick="SelectFile();">\n' 
	   Optionstr += '</td>\n'
       Optionstr += '</tr>\n'
       Optionstr += '<tr  class="hback"> \n'
       Optionstr += '<td align="right">���ص�ַ����</td>\n'
       Optionstr += '<td>\n'
	   Optionstr += '<input type="text" name="Number" id="Number" size="10" maxlength="1" value=""><span class=tx>Ĭ�����������գ�</span>\n'
       Optionstr += '</td>\n'
       Optionstr += '</tr>\n' ;
	   }
  Optionstr += '</table>\n' ;
  //alert(Optionstr);
  document.all.OtherInput.innerHTML = Optionstr;
  }

function SelectFile()     
{
 var returnvalue = OpenWindow('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window);
 if (returnvalue!='')
 {
 	event.srcElement.parentNode.firstChild.value=returnvalue;
 }
}

function ReImgSize(objstr)
{
	if ($(objstr).tagName=='IMG')
	if ($(objstr).src.indexOf("Files/")>-1)
	{	
		if ($(objstr).offsetWidth>100) 	$(objstr).width="100";
	}	
}  
function SelectSpecial()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('lib/SelectspecialFrame.asp',400,300,window);
	if (ReturnValue.indexOf('***')!=-1)
	{
		TempArray = ReturnValue.split('***');
		if (document.form1.SpecialID.value.search(TempArray[1])==-1)
		{
		if(document.all.SpecialID.value=='') document.all.SpecialID.value=TempArray[1];
		else document.all.SpecialID.value=document.all.SpecialID.value+','+TempArray[1];
		if(document.all.SpecialID_EName.value=='') document.all.SpecialID_EName.value=TempArray[0];
		else document.all.SpecialID_EName.value=document.all.SpecialID_EName.value+','+TempArray[0];
		}
	}
}
function SelectSpecail()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('lib/SelectSpecialFrame.asp',400,300,window);
	if(ReturnValue=='undefined')
	{
		return false;
	}
	if (ReturnValue.indexOf('***')!=-1)
	{
		TempArray = ReturnValue.split('***');
		$('speicalId').value=TempArray[0]
		$('txt_specialname').value=TempArray[1]
	}
}

    
-->
</script>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






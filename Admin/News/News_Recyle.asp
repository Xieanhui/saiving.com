<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,Fs_News
MF_Default_Conn
MF_Session_TF 
'Ȩ���ж�
'Call MF_Check_Pop_TF("NS_Class_000001")
Dim Recyle_Type,Rec_Table_Style,strShowErr
Dim Re_Sql_Str,Re_Temp,Re_ClassRecordset,Re_Temp_IsURL,Re_Class_Flag,Re_News_Flag
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_
Dim Page,cPageNo,fso_tmprs_,NewsSavePath
Recyle_Type=Request.QueryString("Recyle_Type")
Re_Class_Flag=False
Re_News_Flag=False

int_RPP=30 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"			'βҳ
set Fs_News=new Cls_News
%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title>ר�����___Powered by foosun Inc.</title>
	<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr class="hback">
			<td class="xingmu">
				<a href="#" class="sd"><strong>����վ����</strong></a>
			</td>
		</tr>
		<tr>
			<td width="100%" height="18" class="hback">
				<div align="left">
					<a href="News_Recyle.asp">������ҳ</a> | <a href="News_Recyle.asp?Recyle_Type=Class">��Ŀ����</a> | <a href="News_Recyle.asp?Recyle_Type=News">���Ź���</a> | <a href="News_Recyle.asp?Recyle_Type=DelAll">����</a></div>
			</td>
		</tr>
	</table>
	<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table" id="Rec_Class" style="display: none">
		<form name="for_Re_ClassOPType" method="post" action="?Recyle_Type=Class&Rec_OP_Class_Type=P">
		<tr class="xingmu">
			<input name="Hi_Re_OP_Class_Type" type="hidden" value="">
			<td height="20" class="xingmu" width="24%">
				<div align="center">
					��Ŀ����</div>
			</td>
			<td class="xingmu" width="24%">
				<div align="center">
					����Ŀ</div>
			</td>
			<td class="xingmu" width="15%">
				<div align="center">
					��Ŀ����</div>
			</td>
			<td class="xingmu" width="15%">
				<div align="center">
					������</div>
			</td>
			<td class="xingmu" width="20%">
				<div align="center">
					����</div>
			</td>
		</tr>
		<%
Dim Str_SYSPath
If G_VIRTUAL_ROOT_DIR<>"" Then
	Str_SYSPath="/"&G_VIRTUAL_ROOT_DIR
Else
	Str_SYSPath=""
End If
	If Recyle_Type="Class" then
		Dim Recyle_OP,Resume_Sql_Str,Resume_ClassID,Resume_Rs,ReCyle_Action,Re_Class_OPType,Re_Class_I,Re_Class_P_Flag
		Recyle_OP=NoSqlHack(Cstr(Request.QueryString("Recyle_OP")))
		Resume_ClassID=NoSqlHack(Cstr(Request.QueryString("ClassID")))
		Re_Class_OPType=NoSqlHack(Cstr(Request.QueryString("Rec_OP_Class_Type")))
		If Re_Class_OPType="P" Then
			Dim Re_Class_Temp_OPType,Re_Class_Temp_ID,Rs_ClassDel
			Re_Class_Temp_OPType=NoSqlHack(Cstr(Request.Form("Hi_Re_OP_Class_Type")))
			Re_Class_Temp_ID=NoSqlHack(Request.Form("Che_ClassOPType"))
			If Re_Class_Temp_ID="" or IsNull(Re_Class_Temp_ID) Then
				strShowErr = "<li>��ѡ��Ҫ��������������!</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End If
			If instr(Re_Class_Temp_ID,",")=0 Then
				Re_Class_Temp_ID=NoSqlHack(Cstr(trim(Re_Class_Temp_ID)))
				Re_Class_P_Flag=True
			Else
				Re_Class_Temp_ID=Split(Re_Class_Temp_ID,",")
				int_arr_Class_ub=Ubound(Re_Class_Temp_ID)+1
				Re_Class_P_Flag=False
			End If
			
			Select Case Re_Class_Temp_OPType
				Case "P_Class_Del"
					
					If Re_Class_P_Flag=True Then
						Set Rs_ClassDel=Conn.Execute("Select SavePath,ClassEName,IsURL FROM FS_NS_NewsClass Where ClassID In('"&Replace(Re_Class_Temp_ID,"'","")&"')")
						While Not Rs_ClassDel.Eof
							If Cint(Rs_ClassDel(2)) = 0 Then
								If Rs_ClassDel(0)="/" Then
									fso_DeleteFolder("/"&Rs_ClassDel(1))
								Else
									fso_DeleteFolder(Str_SYSPath&Rs_ClassDel(0)&"/"&Rs_ClassDel(1))
								End IF
							End If	
							Rs_ClassDel.movenext
						Wend
						Recyle_DelClassID(NoSqlHack(Re_Class_Temp_ID))
						strShowErr = "<li>����ɾ���ɹ�</li>"
						Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
						Response.end
					Else
					'Crazy 2008.10.29
						For Re_Class_I=0 To Ubound(Re_Class_Temp_ID)
							Set Rs_ClassDel=Conn.Execute("Select SavePath,ClassEName,IsURL FROM FS_NS_NewsClass Where ClassID In('"&Replace(Re_Class_Temp_ID(Re_Class_I),"'","")&"')")
							While Not Rs_ClassDel.Eof
								If Cint(Rs_ClassDel(2)) = 0 Then
									If Rs_ClassDel(0)="/" Then
										fso_DeleteFolder("/"&Rs_ClassDel(1))
									Else
										fso_DeleteFolder(Str_SYSPath&Rs_ClassDel(0)&"/"&Rs_ClassDel(1))
									End IF
								End If	
								Rs_ClassDel.movenext
							Wend
							
							Recyle_DelClassID(Re_Class_Temp_ID(Re_Class_I))
						Next
						strShowErr = "<li>����ɾ���ɹ�</li>"
						Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
						Response.end
					End If
				Case "P_Class_Resmue"
					Dim obj_Ser_ParendID_Rs,str_Ser_Par_Sql,int_Class_Count,int_Class_Error_Count,int_arr_Class_ub
					If Re_Class_P_Flag=True Then
						str_Ser_Par_Sql="Select ClassID from FS_NS_NewsClass Where ParentID in(Select ClassID from FS_NS_NewsClass Where ReycleTF=1) and ClassID='"&Cintstr(Re_Class_Temp_ID)&"'"
						Set obj_Ser_ParendID_Rs=Conn.execute(str_Ser_Par_Sql)
						If Not obj_Ser_ParendID_Rs.Eof Then
							strShowErr = "<li>��ǰ��Ŀ�ĸ���Ŀ�ڻ���վ�У����Ȼָ�����Ŀ!</li>"
							Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
							Response.end
						Else
							Conn.execute("Update FS_NS_NewsClass Set ReycleTF=0 Where ClassID='"&NoSqlHack(Re_Class_Temp_ID)&"'")
							strShowErr = "<li>�ָ��ɹ�!</li>"
							Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
							Response.end
						End If 
						Set obj_Ser_ParendID_Rs=Nothing
					Else
						Re_Class_Temp_ID=Split(Serch_ClassID(Re_Class_Temp_ID),",")
						For Re_Class_I=0 To Ubound(Re_Class_Temp_ID)
							Conn.execute("Update FS_NS_NewsClass Set ReycleTF=0 Where ClassID='"&NoSqlHack(Cstr(trim(Re_Class_Temp_ID(Re_Class_I))))&"'")		
						Next
						strShowErr = "<li>ѡ��"&int_arr_Class_ub&"��,�ѻָ�"&Ubound(Re_Class_Temp_ID)&",ʧ��"&int_arr_Class_ub-Ubound(Re_Class_Temp_ID)&"��</li>"
						strShowErr = strShowErr&"<li>�ָ��ɹ�!</li>"
						Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
						Response.end
					End If
			End Select
		End If
		
		If Recyle_OP="ResumeClass" then
			Resume_Sql_Str="Select ClassID from FS_NS_NewsClass Where ParentID in(Select ClassID from FS_NS_NewsClass Where ReycleTF=1) and ClassID='"&NoSqlHack(Resume_ClassID)&"'"
			Set Resume_Rs=Conn.execute(Resume_Sql_Str)	
			If Not Resume_Rs.Eof Then
				strShowErr = "<li>��ǰ��Ŀ�ĸ���Ŀ�ڻ���վ�У����Ȼָ�����Ŀ!</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			Else
				Conn.execute("Update FS_NS_NewsClass Set ReycleTF=0 Where ClassID='"&Resume_ClassID&"'")
				strShowErr = "<li>�ָ��ɹ�!</li>"
				Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			Set Resume_Rs=Nothing
		End if
		ReCyle_Action=Request.QueryString("Action")
		If ReCyle_Action="Submit" Then
			Set Rs_ClassDel=Conn.Execute("Select SavePath,ClassEName,IsURL FROM FS_NS_NewsClass Where ClassID In('"&FormatIntArr(Resume_ClassID)&"')")
			While Not Rs_ClassDel.Eof
				If Cint(Rs_ClassDel(2)) = 0 Then
					If Rs_ClassDel(0)="/" Then
						fso_DeleteFolder("/"&Rs_ClassDel(1))
					Else
						fso_DeleteFolder(Str_SYSPath&Rs_ClassDel(0)&"/"&Rs_ClassDel(1))
					End IF
				End If	
				Rs_ClassDel.movenext
			Wend
			Recyle_DelClassID(Resume_ClassID)
			strShowErr = "<li>ɾ���ɹ�!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if		
		Re_Sql_Str="Select ClassID,ClassName,IsURL from FS_NS_NewsClass where ReycleTF=1 order by parentID Asc"
		Set Re_ClassRecordset= CreateObject(G_FS_RS)
		Re_ClassRecordset.Open Re_Sql_Str,Conn,1,1
		If Not Re_ClassRecordset.Eof Then
			Re_Class_Flag=True
			
			Re_ClassRecordset.PageSize=int_RPP
			cPageNo=NoSqlHack(Request.QueryString("page"))
			If cPageNo="" Then cPageNo = 1
			If not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo<=0 Then cPageNo=1
			If cPageNo>Re_ClassRecordset.PageCount Then cPageNo=Re_ClassRecordset.PageCount 
			Re_ClassRecordset.AbsolutePage=cPageNo

			For int_Start=1 TO int_RPP  
				Re_Temp_IsURL=Re_ClassRecordset("IsURL")
				If Re_Temp_IsURL=0 Then
					Re_Temp_IsURL="�ڲ���Ŀ"
				Else
					Re_Temp_IsURL="�ⲿ��Ŀ"
				End if
		%>
		<tr class="hback">
			<td height="22" class="hback" align="center">
				<%=Re_ClassRecordset("ClassName")%>
			</td>
			<%
		Dim Re_Sql_ParName,Re_Rs_ParName,Re_Temp_Parname
		Re_Sql_ParName="Select ClassName from FS_NS_NewsClass where ClassID in(Select ParentID from FS_NS_NewsClass where ClassID='"&NoSqlHack(Re_ClassRecordset("ClassID"))&"')"
		Set Re_Rs_ParName=Conn.execute(Re_Sql_ParName)
		If not Re_Rs_ParName.Eof Then
			Re_Temp_Parname=Re_Rs_ParName("ClassName")
			If Re_Temp_Parname="0" then Re_Temp_Parname="������Ŀ"
		Else
			Re_Temp_Parname="������Ŀ"
		End if
		Set Re_Rs_ParName =Nothing
			%>
			<td align="center" class="hback">
				<%=Re_Temp_Parname%>
			</td>
			<td align="center" class="hback">
				<%=Re_Temp_IsURL%>
			</td>
			<td align="center" class="hback">
				<%=Conn.execute("Select Count(*) from FS_NS_News Where ClassID='"&Re_ClassRecordset("ClassID")&"'")(0)%>
			</td>
			<td class="hback" align="center">
				<a href="?Recyle_Type=Class&Page=<%=NoSqlHack(Request.QueryString("Page"))%>&Recyle_OP=ResumeClass&ClassID=<%=Re_ClassRecordset("ClassID")%>">�ָ�</a> | <a href="javascript:DelOneClass('<%=Re_ClassRecordset("ClassID")%>');">ɾ��</a> |
				<input type="checkbox" name="Che_ClassOPType" value="<%=Re_ClassRecordset("ClassID")%>">
			</td>
		</tr>
		<%
				Re_ClassRecordset.MoveNext
				If Re_ClassRecordset.Eof or Re_ClassRecordset.Bof Then Exit For
			Next
			Response.Write "<tr><td class=""hback"" colspan=""3"" align=""left"">"&fPageCount(Re_ClassRecordset,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"&nbsp;</td><td colspan=""2"" class=hback><table border=0 align=left cellpadding=0 cellspacing=0 class=table><tr><td class=""hback""  align=""left""><input type=""button"" value="" �����ָ� "" name=""But_P_Class_Resmue"" onclick=""javascript:P_Class_Resmue();"">&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" value="" ����ɾ�� "" name=""But_P_Class_Del"" onclick=""javascript:P_Class_Del();""></td><td class=hback>&nbsp;ȫѡ<input type=""checkbox"" name=""Che_ClassOPType"" onclick=""CheckAll('Che_ClassOPType');""></td></tr></table></td></tr>"
		Else
			Re_Class_Flag=False
   		 	Response.write"<table width=""98%"" border=0 align=center cellpadding=2 cellspacing=1 class=table><tr class=""hback""><td>����վ��û����Ŀ!</td></tr></table>"
		End if
	End If
	Set Re_ClassRecordset=Nothing
		%>
		</form>
	</table>
	<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table" id="Rec_News" style="display: none">
		<form name="for_Re_OPType" method="post" action="?Recyle_Type=News&Rec_OP_Type=P">
		<tr class="xingmu">
			<input name="Hi_Re_OP_Type" type="hidden" value="">
			<td height="25" class="xingmu" width="68%">
				<div align="center">
					���ű���</div>
			</td>
			<td class="xingmu" width="15%">
				<div align="center">
					����������Ŀ</div>
			</td>
			<td class="xingmu" width="15%">
				<div align="center">
					����</div>
			</td>
		</tr>
		<%
	If Recyle_Type="News" then
		Dim Rec_News_OP,Rec_News_Sql_Str,Rec_News_Rs,Rec_News_ClassRs,Rec_News_Class_TempName,Rec_News_Ser_Flag,ReCyle_News_Action,Rec_OP_Type,Rec_News_Temp_Serch,Rec_News_Temp_Serch_Keyword,News_Temp_Serch_Flag
		News_Temp_Serch_Flag=False
		Rec_News_OP=NoSqlHack(Request.QueryString("Recyle_News_OP"))
		Rec_News_NewsID=NoSqlHack(Cstr(Replace(Request.QueryString("NewsID"),"'","")))
		Rec_OP_Type=NoSqlHack(Cstr(Request.QueryString("Rec_OP_Type")))
		If Rec_OP_Type="P" Then
			Dim Rec_TempOp,Rec_TempID,Rec_Temp_I,Rec_Temp_OPFlag
			Rec_TempOp=NoSqlHack(Request.Form("Hi_Re_OP_Type"))
			Rec_TempID=FormatStrArr(Request.Form("Che_OPType"))
			If Rec_TempID="" or IsNull(Rec_TempID) Then
				strShowErr = "<li>��ѡ��Ҫ��������������!</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End If
			Select Case Rec_TempOp
				Case "P_Resmue"
						Conn.execute("Update FS_NS_News Set isRecyle=0 where NewsID in ('"&Rec_TempID&"')")
						strShowErr = "<li>�����ָ��ɹ�</li>"
						Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
						Response.end
				Case "P_Del"
					'ɾ����̬�ļ�
					set fso_tmprs_ = Conn.execute("select FS_NS_News.SaveNewsPath,FS_NS_News.FileName,FS_NS_News.FileExtName,FS_NS_NewsClass.SavePath,FS_NS_NewsClass.ClassEName from FS_NS_News,FS_NS_NewsClass where FS_NS_News.NewsID in ('"&Rec_TempID&"') and FS_NS_News.IsURL=0 and FS_NS_NewsClass.ClassID=FS_NS_News.ClassID")
					While Not fso_tmprs_.eof
						If G_VIRTUAL_ROOT_DIR = "" Then
							NewsSavePath = ""
						Else
							NewsSavePath = "/" & G_VIRTUAL_ROOT_DIR
						End If
						NewsSavePath=NewsSavePath&fso_tmprs_("SavePath")&"/"&fso_tmprs_("ClassEName")&fso_tmprs_("SaveNewsPath")&"/"&fso_tmprs_("FileName")&"."&fso_tmprs_("FileExtName")
						fso_DeleteFile(NewsSavePath)
						fso_tmprs_.movenext
					Wend
					'ɾ����̬�ļ�����
					Conn.execute("Delete from FS_NS_News where NewsID in ('"&Rec_TempID&"')")
					Call MF_Insert_oper_Log("����վ","����ɾ��������",now,session("admin_name"),"NS")
					strShowErr = "<li>����ɾ���ɹ�</li>"
					Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
			End Select
		End If
		If Rec_News_OP="ResumeNews" Then
			Dim Rec_News_NewsID,Rec_News_NewsIDRs
			Rec_News_Sql_Str="Select ID from FS_NS_NewsClass where ClassID in(Select ClassID from FS_NS_News Where NewsID='"&NoSqlHack(Rec_News_NewsID)&"' and ReycleTF=0) "
			Set Rec_News_NewsIDRs=Conn.execute(Rec_News_Sql_Str)
			If Not Rec_News_NewsIDRs.Eof Then
				Conn.execute("Update FS_NS_News Set isRecyle=0 where NewsID='"&NoSqlHack(Rec_News_NewsID)&"'")
				strShowErr = "<li>�ָ��ɹ�!</li>"
				Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			Else
				Conn.execute("Update FS_NS_News Set isRecyle=0,ClassID='0' where NewsID='"&NoSqlHack(Rec_News_NewsID)&"'")
				strShowErr = "<li>�ָ��ɹ�!</li>"
				Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End If
			Set Rec_News_NewsIDRs=Nothing
		End If
		ReCyle_News_Action=Request.QueryString("Action")
		If ReCyle_News_Action="Submit" Then
			'ɾ����̬�ļ�
			set fso_tmprs_ = Conn.execute("select FS_NS_News.SaveNewsPath,FS_NS_News.FileName,FS_NS_News.FileExtName,FS_NS_NewsClass.SavePath,FS_NS_NewsClass.ClassEName from FS_NS_News,FS_NS_NewsClass where FS_NS_News.NewsID in ('"&Replace(Replace(FormatIntArr(Rec_News_NewsID)," ",""),",","','")&"') and FS_NS_News.IsURL=0 and FS_NS_NewsClass.ClassID=FS_NS_News.ClassID")
			While Not fso_tmprs_.eof
				If G_VIRTUAL_ROOT_DIR = "" Then
					NewsSavePath = ""
				Else
					NewsSavePath = "/" & G_VIRTUAL_ROOT_DIR
				End If
				NewsSavePath=NewsSavePath&fso_tmprs_("SavePath")&"/"&fso_tmprs_("ClassEName")&fso_tmprs_("SaveNewsPath")&"/"&fso_tmprs_("FileName")&"."&fso_tmprs_("FileExtName")
				fso_DeleteFile(NewsSavePath)
				fso_tmprs_.movenext
			Wend
			'ɾ����̬�ļ�����
			Conn.execute("Delete from FS_NS_News where NewsID='"&NoSqlHack(Rec_News_NewsID)&"'")
			strShowErr = "<li>ɾ���ɹ�!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if		
		Rec_News_Ser_Flag=False
		Rec_News_Temp_Serch=NoSqlHack(Request.QueryString("Serch"))
		Rec_News_Temp_Serch_Keyword=NoSqlHack(Request.Form("Ser_Keyword"))
		If Rec_News_Temp_Serch="Submit" Then
			Rec_News_Sql_Str="Select NewsID,NewsTitle,ClassID From FS_NS_News where isRecyle=1 and NewsTitle like'%"&NoSqlHack(Rec_News_Temp_Serch_Keyword)&"%'"
			News_Temp_Serch_Flag=True
		Else
			Rec_News_Sql_Str="Select NewsID,NewsTitle,ClassID From FS_NS_News where isRecyle=1"
		End If
		Set Rec_News_Rs= CreateObject(G_FS_RS)
		Rec_News_Rs.Open Rec_News_Sql_Str,Conn,1,1
		If Not Rec_News_Rs.Eof Then
			
			Rec_News_Rs.PageSize=int_RPP
			cPageNo=NoSqlHack(Request.QueryString("page"))
			If cPageNo="" Then cPageNo = 1
			If not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo<=0 Then cPageNo=1
			If cPageNo>Rec_News_Rs.PageCount Then cPageNo=Rec_News_Rs.PageCount 
			Rec_News_Rs.AbsolutePage=cPageNo

			Re_News_Flag=True
			Rec_News_Ser_Flag=True
			For int_Start=1 TO int_RPP  
		%>
		<tr class="hback">
			<td height="22" class="hback">
				<%=Rec_News_Rs("NewsTitle")%>
			</td>
			<%
				Set Rec_News_ClassRs=Conn.execute("Select ClassName from FS_NS_NewsClass where ClassID='"&Rec_News_Rs("ClassID")&"'")
				If Not Rec_News_ClassRs.Eof Then
					Rec_News_Class_TempName=Rec_News_ClassRs("ClassName")
				Else
					Rec_News_Class_TempName="��Ŀ������"
				End if
				Set Rec_News_ClassRs=Nothing			
			%>
			<td align="center" class="hback" width="221">
				<%=Rec_News_Class_TempName%>
			</td>
			<td align="center" class="hback" width="151">
				<a href="?Recyle_Type=News&Page=<%=NoSqlHack(Request.QueryString("Page"))%>&Recyle_News_OP=ResumeNews&NewsID=<%=Rec_News_Rs("NewsID")%>">�ָ�</a> | <a href="javascript:DelOneNews('<%=Rec_News_Rs("NewsID")%>')">ɾ��</a> |
				<input type="checkbox" name="Che_OPType" value="<%=Rec_News_Rs("NewsID")%>">
			</td>
		</tr>
		<%
				Rec_News_Rs.MoveNext
				If Rec_News_Rs.Eof or Rec_News_Rs.Bof Then Exit For
			Next
			Response.Write "<tr><td class=""hback"" colspan=""6"" align=""left""><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td align=""left"">"&fPageCount(Rec_News_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td><td align=""right""><input type=""button"" value="" �����ָ� "" name=""But_P_UnLock"" onclick=""javascript:P_Resmue();"">&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" value="" ����ɾ�� "" name=""But_P_Del"" onclick=""javascript:P_Del();"">&nbsp;&nbsp;&nbsp;&nbsp;ȫѡ<input type=""checkbox"" name=""Che_OPType"" onclick=""CheckAll('Che_OPType');"" ></td></tr></table></td></tr>"
		Else
			If News_Temp_Serch_Flag=True Then
				strShowErr = "<li>û���ҵ���������������,������!</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			Else
  	 		 	Response.write"<table width=""98%"" border=0 align=center cellpadding=2 cellspacing=1 class=table><tr class=""hback""><td>����վ��û������!</td></tr></table>"
			End If
		End if
		Set Rec_News_Rs=Nothing
	End if
	Response.write "<script language=""javascript"">document.all.Rec_News.style.display=""none"";</script>"
		%>
		</form>
	</table>
	<%
	Recyle_Type=NoSqlHack(Request.QueryString("Recyle_Type"))
	If Recyle_Type="DelAll" Then
		Dim Recyle_Del_OP,Recyle_DelAll_Type
		Recyle_Del_OP=NoSqlHack(Request.QueryString("Recyle_Del_OP"))
		Recyle_DelAll_Type=NoSqlHack(Request.QueryString("Action"))
		If Recyle_DelAll_Type="S_DelAllClass" Then
			Conn.execute("Delete From FS_NS_News Where isRecyle=1")
			Call MF_Insert_oper_Log("����վ","����˻���վ������Ŀ",now,session("admin_name"),"NS")
			strShowErr = "<li>ɾ���ɹ�!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.End
		End If	
		If Recyle_DelAll_Type="S_DelAllNews" Then
			'ɾ����̬�ļ�
			set fso_tmprs_ = Conn.execute("select FS_NS_News.SaveNewsPath,FS_NS_News.FileName,FS_NS_News.FileExtName,FS_NS_NewsClass.SavePath,FS_NS_NewsClass.ClassEName from FS_NS_News,FS_NS_NewsClass where FS_NS_News.isRecyle=1 and FS_NS_News.IsURL=0 and FS_NS_NewsClass.ClassID=FS_NS_News.ClassID")
			While Not fso_tmprs_.eof
				If G_VIRTUAL_ROOT_DIR = "" Then
					NewsSavePath = ""
				Else
					NewsSavePath = "/" & G_VIRTUAL_ROOT_DIR
				End If
				NewsSavePath=NewsSavePath&fso_tmprs_("SavePath")&"/"&fso_tmprs_("ClassEName")&fso_tmprs_("SaveNewsPath")&"/"&fso_tmprs_("FileName")&"."&fso_tmprs_("FileExtName")
				fso_DeleteFile(NewsSavePath)
				fso_tmprs_.movenext
			Wend
			'ɾ����̬�ļ�����
			Conn.execute("Delete From FS_NS_News Where isRecyle=1")
			Call MF_Insert_oper_Log("����վ","����˻���վ��������",now,session("admin_name"),"NS")
			strShowErr = "<li>ɾ���ɹ�!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.End
		End If
		If Recyle_DelAll_Type="S_DelAll" Then
			'ɾ����̬�ļ�
			set fso_tmprs_ = Conn.execute("select FS_NS_News.SaveNewsPath,FS_NS_News.FileName,FS_NS_News.FileExtName,FS_NS_NewsClass.SavePath,FS_NS_NewsClass.ClassEName from FS_NS_News,FS_NS_NewsClass where FS_NS_News.isRecyle=1 and FS_NS_News.IsURL=0 and FS_NS_NewsClass.ClassID=FS_NS_News.ClassID")
			While Not fso_tmprs_.eof
				If G_VIRTUAL_ROOT_DIR = "" Then
					NewsSavePath = ""
				Else
					NewsSavePath = "/" & G_VIRTUAL_ROOT_DIR
				End If
				NewsSavePath=NewsSavePath&fso_tmprs_("SavePath")&"/"&fso_tmprs_("ClassEName")&fso_tmprs_("SaveNewsPath")&"/"&fso_tmprs_("FileName")&"."&fso_tmprs_("FileExtName")
				fso_DeleteFile(NewsSavePath)
				fso_tmprs_.movenext
			Wend
			'ɾ����̬�ļ�����
			Conn.execute("Delete From FS_NS_News Where isRecyle=1")
			Conn.execute("Delete From FS_NS_NewsClass Where ReycleTF=1")
			Call MF_Insert_oper_Log("����վ","����˻���վ������Ϣ",now,session("admin_name"),"NS")
			strShowErr = "<li>��ջ���վ�ɹ�!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.End
		End If
		If Recyle_Del_OP="ResumeAll" Then
			Conn.execute("Update FS_NS_News Set isRecyle=0 Where isRecyle=1")
			Conn.execute("Update FS_NS_NewsClass Set ReycleTF=0 Where ReycleTF=1")
			strShowErr = "<li>�ָ��ɹ�!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
		If Recyle_Del_OP="ResumeClass" Then
			Conn.execute("Update FS_NS_NewsClass Set ReycleTF=0 Where ReycleTF=1")
			strShowErr = "<li>�ָ��ɹ�!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
		If Recyle_Del_OP="ResumeNews" Then
			Conn.execute("Update FS_NS_News Set isRecyle=0 Where isRecyle=1")
			strShowErr = "<li>�ָ��ɹ�!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
	End If
	%>
	<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table" id="Rec_Op" style="display: none">
		<tr class="hback">
			<td height="47" align="center">
				<a href="?Recyle_Type=DelAll&Recyle_Del_OP=ResumeAll">ȫ���ָ�</a> | <a href="?Recyle_Type=DelAll&Recyle_Del_OP=ResumeClass">ֻ�ָ���Ŀ</a> | <a href="?Recyle_Type=DelAll&Recyle_Del_OP=ResumeNews">ֻ�ָ�����</a> | <a href="javascript:DelClass();">ɾ����Ŀ</a> | <a href="javascript:DelNews()">ɾ������</a> | <a href="javascript:DelAll()">��ջ���վ</a>
				<div align="center">
				</div>
			</td>
		</tr>
	</table>
	<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table" id="Rec_ZhuJie" style="display: none">
		<tr class="hback">
			<td height="47">
				<font color="#FF3300">ע��
					<p>
						1����Ŀ����:�ǶԷ������վ�е���Ŀ�����еĹ���</p>
					<p>
						2�����Ź���:�ǶԷ������վ�е����Ŷ����еĹ���</p>
					<p>
					3������:�Է������վ�е����ݽ��лָ���ɾ�������� </font>
			</td>
		</tr>
	</table>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table" id="Rec_Serch" style="display: none">
		<form name="for_Ser_News" method="post" action="?Recyle_Type=News&Serch=Submit">
		<tr>
			<td height="18" class="hback">
				�����������ؼ���
				<input name="Ser_Keyword" type="text" size="20">
				<select name="NewsType" id="NewsType">
					<option value="title" selected>����</option>
				</select>
				<input type="submit" name="But_Ser_Submit" value="  �� ��  ">
			</td>
		</tr>
		</form>
	</table>
	<%
If Rec_News_Ser_Flag=True Then
	Response.write "<script language=""javascript"">document.all.Rec_Serch.style.display="""";</script>"
End if
	%>
	<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
</body>
</html>
<%
Sub Recyle_DelClassID(ClassID)
	Dim Recyle_ClassID_Sql_Str,Recyle_ClassID_Rs
	Recyle_ClassID_Sql_Str="Select ClassID from FS_NS_NewsClass where ParentID='"&NoSqlHack(ClassID)&"'"
	Set Recyle_ClassID_Rs=Conn.execute(Recyle_ClassID_Sql_Str)
	While Not Recyle_ClassID_Rs.Eof 
		Recyle_DelClassID(Recyle_ClassID_Rs("ClassID"))
		Recyle_ClassID_Rs.MoveNext
	Wend
	Conn.execute("Delete from FS_NS_NewsClass where ClassID='"&Replace(ClassID,"'","")&"'")
	Response.write("Delete from FS_NS_NewsClass where ClassID='"&NoSqlHack(ClassID)&"'")
	Set Recyle_ClassID_Rs=Nothing
End Sub

Function Serch_ClassID(arr_CLassID)
	Dim int_i,str_s_sql,obj_ser_pa_Rs
	For int_i=0 To Ubound(arr_CLassID)
		str_s_sql="Select ClassID from FS_NS_NewsClass Where ParentID in(Select ClassID from FS_NS_NewsClass Where ReycleTF=1) and ClassID='"&NoSqlHack(Cstr(arr_CLassID(int_i)))&"'"
		Set obj_ser_pa_Rs=Conn.execute(str_s_sql)
		If obj_ser_pa_Rs.Eof Then
			Serch_ClassID=Serch_ClassID&arr_CLassID(int_i)&","
		End If 
		Set obj_ser_pa_Rs=Nothing	
	Next
End Function
%>
<script type="text/javascript">
function DelClass()
{
	if(confirm('�˲���������,����ɾ�����л���վ����Ŀ\n��ȷ��ɾ����'))
	{
		location='?Recyle_Type=DelAll&Action=S_DelAllClass';
	}	
}
function DelNews()
{
	if(confirm('�˲���������,����ɾ�����л���վ������\n��ȷ��ɾ����'))
	{
		location='?Recyle_Type=DelAll&Action=S_DelAllNews';
	}
}
function DelAll()
{
	if(confirm('�˲���������,����ɾ�����л���վ����Ŀ�Լ�����\n��ȷ��ɾ����'))
	{
		location='?Recyle_Type=DelAll&Action=S_DelAll';	
	}
}
function DelOneNews(NewsId)
{
	if(confirm('��ɾ��һ������\n��ȷ��ɾ����'))
	{
		location='?Recyle_Type=News&Page=<%=NoSqlHack(Request.QueryString("Page"))%>&NewsID='+NewsId+'&Action=Submit';
	}
}
function DelOneClass(ClassID)
{		
	if(confirm('�˲�����ɾ������Ŀ�µ���������Ŀ��\n��ȷ��ɾ����'))
	{
		location='?Recyle_Type=Class&Page=<%=NoSqlHack(Request.QueryString("Page"))%>&ClassID='+ClassID+'&Action=Submit';
	}
}
function P_Class_Del()
{
	if(confirm('�˲�����ɾ��ѡ����Ŀ�µ���������Ŀ��\n��ȷ��ɾ����'))
	{	
		document.for_Re_ClassOPType.Hi_Re_OP_Class_Type.value="P_Class_Del";
		document.for_Re_ClassOPType.submit();
	}
}
function P_Class_Resmue()
{
	if(confirm('�˲������ָ�ѡ�е���Ŀ��\n��ȷ���ָ���'))
	{
		document.for_Re_ClassOPType.Hi_Re_OP_Class_Type.value="P_Class_Resmue";
		document.for_Re_ClassOPType.submit();
	}
}
function SelectTableShow(ShowType)
{
	switch(ShowType)
		{
			case 0://��ʾ����վע����ҳ
				document.all.Rec_ZhuJie.style.display="";
				document.all.Rec_Class.style.display="none";
				document.all.Rec_News.style.display="none";
				document.all.Rec_Serch.style.display="none";
				document.all.Rec_Op.style.display="none";
				break;				
			case 1://��ʾ����վ��Ŀ������ҳ
				<%
				If Re_Class_Flag=False Then
						Response.write "document.all.Rec_Class.style.display=""none"";"
					Else
						Response.write"document.all.Rec_Class.style.display="""";"
				End If
				%>				
				document.all.Rec_ZhuJie.style.display="none";
				document.all.Rec_News.style.display="none";
				document.all.Rec_Serch.style.display="none";
				document.all.Rec_Op.style.display="none";				
				break;
			case 2://��ʾ����վ���Ź�����ҳ
				<%
					If Re_News_Flag=False Then
						Response.write "document.all.Rec_News.style.display=""none"";"
						Response.write "document.all.Rec_Serch.style.display=""none"";"
					Else
						Response.write"document.all.Rec_News.style.display="""";"
						Response.write "document.all.Rec_Serch.style.display="""";"
					End If
				%>				
				document.all.Rec_ZhuJie.style.display="none";
				document.all.Rec_Class.style.display="none";
				document.all.Rec_Op.style.display="none";
				break;
			case 3://��ʾ����վ���������ҳ
				document.all.Rec_Op.style.display="";
				document.all.Rec_ZhuJie.style.display="none";
				document.all.Rec_Class.style.display="none";
				document.all.Rec_News.style.display="none";
				document.all.Rec_Serch.style.display="none";
				break;				
		}
}
function CheckAll(CheckType)
{
	var checkBoxArray=document.all(CheckType)
	if(checkBoxArray[checkBoxArray.length-1].checked)
	{
		for(var i=0;i<checkBoxArray.length-1;i++)
		{
			checkBoxArray[i].checked=true;
			
		}
	}else
	{
		for(var i=0;i<checkBoxArray.length-1;i++)
		{
			checkBoxArray[i].checked=false;
		}
	}
}
function P_Resmue()
{
	if(confirm('�˲������ָ�ѡ�е����ţ�\n��ȷ���ָ���'))
	{
		document.for_Re_OPType.Hi_Re_OP_Type.value="P_Resmue";
		document.for_Re_OPType.submit();
	}
}
function P_Del()
{
		var chkJss_Array=document.all.Che_OPType;
		var chkTF=false;
		if (chkJss_Array[0]==null)
		{
			return 
		}
		else
		{
			//�ж��Ƿ�ѡ���˼�¼
			for(var i=0;i<chkJss_Array.length;i++)
			{
				if (chkJss_Array[i].checked)
				{
					chkTF=true
				}
			}
		}
		if(chkTF)
		{
			if(confirm("ȷ��Ҫɾ��ѡ�еļ�¼��"))
				document.for_Re_OPType.Hi_Re_OP_Type.value="P_Del";
				document.for_Re_OPType.submit();
		}else
		{
			alert("��ѡ����Ҫɾ���Ķ���");
		}
		//�����ж���ûѡ��ѡ��

}
</script>
<%
If Recyle_Type="" Then
	Response.write"<script language=""javascript"">SelectTableShow(0);</script>"
Else
	Recyle_Type=Cstr(Recyle_Type)
	Select Case Recyle_Type
		Case "Class"
			Response.write"<script language=""javascript"">SelectTableShow(1);</script>"
		Case "News"
			Response.write"<script language=""javascript"">SelectTableShow(2);</script>"
		Case "DelAll"
			Response.write"<script language=""javascript"">SelectTableShow(3);</script>"
		Case else
			Response.write"<script language=""javascript"">SelectTableShow(0);</script>"
	End Select
End IF
Set Conn=Nothing
%>
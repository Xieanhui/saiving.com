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
'权限判断
'Call MF_Check_Pop_TF("NS_Class_000001")
Dim Recyle_Type,Rec_Table_Style,strShowErr
Dim Re_Sql_Str,Re_Temp,Re_ClassRecordset,Re_Temp_IsURL,Re_Class_Flag,Re_News_Flag
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_
Dim Page,cPageNo,fso_tmprs_,NewsSavePath
Recyle_Type=Request.QueryString("Recyle_Type")
Re_Class_Flag=False
Re_News_Flag=False

int_RPP=30 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"			'尾页
set Fs_News=new Cls_News
%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title>专题管理___Powered by foosun Inc.</title>
	<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr class="hback">
			<td class="xingmu">
				<a href="#" class="sd"><strong>回收站管理</strong></a>
			</td>
		</tr>
		<tr>
			<td width="100%" height="18" class="hback">
				<div align="left">
					<a href="News_Recyle.asp">管理首页</a> | <a href="News_Recyle.asp?Recyle_Type=Class">栏目管理</a> | <a href="News_Recyle.asp?Recyle_Type=News">新闻管理</a> | <a href="News_Recyle.asp?Recyle_Type=DelAll">操作</a></div>
			</td>
		</tr>
	</table>
	<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table" id="Rec_Class" style="display: none">
		<form name="for_Re_ClassOPType" method="post" action="?Recyle_Type=Class&Rec_OP_Class_Type=P">
		<tr class="xingmu">
			<input name="Hi_Re_OP_Class_Type" type="hidden" value="">
			<td height="20" class="xingmu" width="24%">
				<div align="center">
					栏目中文</div>
			</td>
			<td class="xingmu" width="24%">
				<div align="center">
					父栏目</div>
			</td>
			<td class="xingmu" width="15%">
				<div align="center">
					栏目类型</div>
			</td>
			<td class="xingmu" width="15%">
				<div align="center">
					新闻数</div>
			</td>
			<td class="xingmu" width="20%">
				<div align="center">
					操作</div>
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
				strShowErr = "<li>请选择要批量操作的内容!</li>"
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
						strShowErr = "<li>批量删除成功</li>"
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
						strShowErr = "<li>批量删除成功</li>"
						Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
						Response.end
					End If
				Case "P_Class_Resmue"
					Dim obj_Ser_ParendID_Rs,str_Ser_Par_Sql,int_Class_Count,int_Class_Error_Count,int_arr_Class_ub
					If Re_Class_P_Flag=True Then
						str_Ser_Par_Sql="Select ClassID from FS_NS_NewsClass Where ParentID in(Select ClassID from FS_NS_NewsClass Where ReycleTF=1) and ClassID='"&Cintstr(Re_Class_Temp_ID)&"'"
						Set obj_Ser_ParendID_Rs=Conn.execute(str_Ser_Par_Sql)
						If Not obj_Ser_ParendID_Rs.Eof Then
							strShowErr = "<li>当前栏目的父栏目在回收站中，请先恢复父栏目!</li>"
							Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
							Response.end
						Else
							Conn.execute("Update FS_NS_NewsClass Set ReycleTF=0 Where ClassID='"&NoSqlHack(Re_Class_Temp_ID)&"'")
							strShowErr = "<li>恢复成功!</li>"
							Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
							Response.end
						End If 
						Set obj_Ser_ParendID_Rs=Nothing
					Else
						Re_Class_Temp_ID=Split(Serch_ClassID(Re_Class_Temp_ID),",")
						For Re_Class_I=0 To Ubound(Re_Class_Temp_ID)
							Conn.execute("Update FS_NS_NewsClass Set ReycleTF=0 Where ClassID='"&NoSqlHack(Cstr(trim(Re_Class_Temp_ID(Re_Class_I))))&"'")		
						Next
						strShowErr = "<li>选中"&int_arr_Class_ub&"条,已恢复"&Ubound(Re_Class_Temp_ID)&",失败"&int_arr_Class_ub-Ubound(Re_Class_Temp_ID)&"条</li>"
						strShowErr = strShowErr&"<li>恢复成功!</li>"
						Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
						Response.end
					End If
			End Select
		End If
		
		If Recyle_OP="ResumeClass" then
			Resume_Sql_Str="Select ClassID from FS_NS_NewsClass Where ParentID in(Select ClassID from FS_NS_NewsClass Where ReycleTF=1) and ClassID='"&NoSqlHack(Resume_ClassID)&"'"
			Set Resume_Rs=Conn.execute(Resume_Sql_Str)	
			If Not Resume_Rs.Eof Then
				strShowErr = "<li>当前栏目的父栏目在回收站中，请先恢复父栏目!</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			Else
				Conn.execute("Update FS_NS_NewsClass Set ReycleTF=0 Where ClassID='"&Resume_ClassID&"'")
				strShowErr = "<li>恢复成功!</li>"
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
			strShowErr = "<li>删除成功!</li>"
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
					Re_Temp_IsURL="内部栏目"
				Else
					Re_Temp_IsURL="外部栏目"
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
			If Re_Temp_Parname="0" then Re_Temp_Parname="顶级栏目"
		Else
			Re_Temp_Parname="顶级栏目"
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
				<a href="?Recyle_Type=Class&Page=<%=NoSqlHack(Request.QueryString("Page"))%>&Recyle_OP=ResumeClass&ClassID=<%=Re_ClassRecordset("ClassID")%>">恢复</a> | <a href="javascript:DelOneClass('<%=Re_ClassRecordset("ClassID")%>');">删除</a> |
				<input type="checkbox" name="Che_ClassOPType" value="<%=Re_ClassRecordset("ClassID")%>">
			</td>
		</tr>
		<%
				Re_ClassRecordset.MoveNext
				If Re_ClassRecordset.Eof or Re_ClassRecordset.Bof Then Exit For
			Next
			Response.Write "<tr><td class=""hback"" colspan=""3"" align=""left"">"&fPageCount(Re_ClassRecordset,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"&nbsp;</td><td colspan=""2"" class=hback><table border=0 align=left cellpadding=0 cellspacing=0 class=table><tr><td class=""hback""  align=""left""><input type=""button"" value="" 批量恢复 "" name=""But_P_Class_Resmue"" onclick=""javascript:P_Class_Resmue();"">&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" value="" 批量删除 "" name=""But_P_Class_Del"" onclick=""javascript:P_Class_Del();""></td><td class=hback>&nbsp;全选<input type=""checkbox"" name=""Che_ClassOPType"" onclick=""CheckAll('Che_ClassOPType');""></td></tr></table></td></tr>"
		Else
			Re_Class_Flag=False
   		 	Response.write"<table width=""98%"" border=0 align=center cellpadding=2 cellspacing=1 class=table><tr class=""hback""><td>回收站中没有栏目!</td></tr></table>"
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
					新闻标题</div>
			</td>
			<td class="xingmu" width="15%">
				<div align="center">
					新闻所属栏目</div>
			</td>
			<td class="xingmu" width="15%">
				<div align="center">
					操作</div>
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
				strShowErr = "<li>请选择要批量操作的内容!</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End If
			Select Case Rec_TempOp
				Case "P_Resmue"
						Conn.execute("Update FS_NS_News Set isRecyle=0 where NewsID in ('"&Rec_TempID&"')")
						strShowErr = "<li>批量恢复成功</li>"
						Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
						Response.end
				Case "P_Del"
					'删除静态文件
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
					'删除静态文件结束
					Conn.execute("Delete from FS_NS_News where NewsID in ('"&Rec_TempID&"')")
					Call MF_Insert_oper_Log("回收站","批量删除了新闻",now,session("admin_name"),"NS")
					strShowErr = "<li>批量删除成功</li>"
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
				strShowErr = "<li>恢复成功!</li>"
				Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			Else
				Conn.execute("Update FS_NS_News Set isRecyle=0,ClassID='0' where NewsID='"&NoSqlHack(Rec_News_NewsID)&"'")
				strShowErr = "<li>恢复成功!</li>"
				Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End If
			Set Rec_News_NewsIDRs=Nothing
		End If
		ReCyle_News_Action=Request.QueryString("Action")
		If ReCyle_News_Action="Submit" Then
			'删除静态文件
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
			'删除静态文件结束
			Conn.execute("Delete from FS_NS_News where NewsID='"&NoSqlHack(Rec_News_NewsID)&"'")
			strShowErr = "<li>删除成功!</li>"
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
					Rec_News_Class_TempName="栏目不存在"
				End if
				Set Rec_News_ClassRs=Nothing			
			%>
			<td align="center" class="hback" width="221">
				<%=Rec_News_Class_TempName%>
			</td>
			<td align="center" class="hback" width="151">
				<a href="?Recyle_Type=News&Page=<%=NoSqlHack(Request.QueryString("Page"))%>&Recyle_News_OP=ResumeNews&NewsID=<%=Rec_News_Rs("NewsID")%>">恢复</a> | <a href="javascript:DelOneNews('<%=Rec_News_Rs("NewsID")%>')">删除</a> |
				<input type="checkbox" name="Che_OPType" value="<%=Rec_News_Rs("NewsID")%>">
			</td>
		</tr>
		<%
				Rec_News_Rs.MoveNext
				If Rec_News_Rs.Eof or Rec_News_Rs.Bof Then Exit For
			Next
			Response.Write "<tr><td class=""hback"" colspan=""6"" align=""left""><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td align=""left"">"&fPageCount(Rec_News_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td><td align=""right""><input type=""button"" value="" 批量恢复 "" name=""But_P_UnLock"" onclick=""javascript:P_Resmue();"">&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" value="" 批量删除 "" name=""But_P_Del"" onclick=""javascript:P_Del();"">&nbsp;&nbsp;&nbsp;&nbsp;全选<input type=""checkbox"" name=""Che_OPType"" onclick=""CheckAll('Che_OPType');"" ></td></tr></table></td></tr>"
		Else
			If News_Temp_Serch_Flag=True Then
				strShowErr = "<li>没有找到符合条件的新闻,请重试!</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			Else
  	 		 	Response.write"<table width=""98%"" border=0 align=center cellpadding=2 cellspacing=1 class=table><tr class=""hback""><td>回收站中没有新闻!</td></tr></table>"
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
			Call MF_Insert_oper_Log("回收站","清空了回收站所有栏目",now,session("admin_name"),"NS")
			strShowErr = "<li>删除成功!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.End
		End If	
		If Recyle_DelAll_Type="S_DelAllNews" Then
			'删除静态文件
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
			'删除静态文件结束
			Conn.execute("Delete From FS_NS_News Where isRecyle=1")
			Call MF_Insert_oper_Log("回收站","清空了回收站所有新闻",now,session("admin_name"),"NS")
			strShowErr = "<li>删除成功!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.End
		End If
		If Recyle_DelAll_Type="S_DelAll" Then
			'删除静态文件
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
			'删除静态文件结束
			Conn.execute("Delete From FS_NS_News Where isRecyle=1")
			Conn.execute("Delete From FS_NS_NewsClass Where ReycleTF=1")
			Call MF_Insert_oper_Log("回收站","清空了回收站所有信息",now,session("admin_name"),"NS")
			strShowErr = "<li>清空回收站成功!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.End
		End If
		If Recyle_Del_OP="ResumeAll" Then
			Conn.execute("Update FS_NS_News Set isRecyle=0 Where isRecyle=1")
			Conn.execute("Update FS_NS_NewsClass Set ReycleTF=0 Where ReycleTF=1")
			strShowErr = "<li>恢复成功!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
		If Recyle_Del_OP="ResumeClass" Then
			Conn.execute("Update FS_NS_NewsClass Set ReycleTF=0 Where ReycleTF=1")
			strShowErr = "<li>恢复成功!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
		If Recyle_Del_OP="ResumeNews" Then
			Conn.execute("Update FS_NS_News Set isRecyle=0 Where isRecyle=1")
			strShowErr = "<li>恢复成功!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
	End If
	%>
	<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table" id="Rec_Op" style="display: none">
		<tr class="hback">
			<td height="47" align="center">
				<a href="?Recyle_Type=DelAll&Recyle_Del_OP=ResumeAll">全部恢复</a> | <a href="?Recyle_Type=DelAll&Recyle_Del_OP=ResumeClass">只恢复栏目</a> | <a href="?Recyle_Type=DelAll&Recyle_Del_OP=ResumeNews">只恢复新闻</a> | <a href="javascript:DelClass();">删除栏目</a> | <a href="javascript:DelNews()">删除新闻</a> | <a href="javascript:DelAll()">清空回收站</a>
				<div align="center">
				</div>
			</td>
		</tr>
	</table>
	<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table" id="Rec_ZhuJie" style="display: none">
		<tr class="hback">
			<td height="47">
				<font color="#FF3300">注：
					<p>
						1、栏目管理:是对放入回收站中的栏目而进行的管理。</p>
					<p>
						2、新闻管理:是对放入回收站中的新闻而进行的管理。</p>
					<p>
					3、操作:对放入回收站中的内容进行恢复和删除操作。 </font>
			</td>
		</tr>
	</table>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table" id="Rec_Serch" style="display: none">
		<form name="for_Ser_News" method="post" action="?Recyle_Type=News&Serch=Submit">
		<tr>
			<td height="18" class="hback">
				新闻搜索：关键字
				<input name="Ser_Keyword" type="text" size="20">
				<select name="NewsType" id="NewsType">
					<option value="title" selected>标题</option>
				</select>
				<input type="submit" name="But_Ser_Submit" value="  搜 索  ">
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
	if(confirm('此操作不可逆,将会删除所有回收站的栏目\n你确定删除吗？'))
	{
		location='?Recyle_Type=DelAll&Action=S_DelAllClass';
	}	
}
function DelNews()
{
	if(confirm('此操作不可逆,将会删除所有回收站的新闻\n你确定删除吗？'))
	{
		location='?Recyle_Type=DelAll&Action=S_DelAllNews';
	}
}
function DelAll()
{
	if(confirm('此操作不可逆,将会删除所有回收站的栏目以及新闻\n你确定删除吗？'))
	{
		location='?Recyle_Type=DelAll&Action=S_DelAll';	
	}
}
function DelOneNews(NewsId)
{
	if(confirm('将删除一条新闻\n你确定删除吗？'))
	{
		location='?Recyle_Type=News&Page=<%=NoSqlHack(Request.QueryString("Page"))%>&NewsID='+NewsId+'&Action=Submit';
	}
}
function DelOneClass(ClassID)
{		
	if(confirm('此操作将删除此栏目下的所有子栏目？\n你确定删除吗？'))
	{
		location='?Recyle_Type=Class&Page=<%=NoSqlHack(Request.QueryString("Page"))%>&ClassID='+ClassID+'&Action=Submit';
	}
}
function P_Class_Del()
{
	if(confirm('此操作将删除选中栏目下的所有子栏目？\n你确定删除吗？'))
	{	
		document.for_Re_ClassOPType.Hi_Re_OP_Class_Type.value="P_Class_Del";
		document.for_Re_ClassOPType.submit();
	}
}
function P_Class_Resmue()
{
	if(confirm('此操作将恢复选中的栏目？\n你确定恢复吗？'))
	{
		document.for_Re_ClassOPType.Hi_Re_OP_Class_Type.value="P_Class_Resmue";
		document.for_Re_ClassOPType.submit();
	}
}
function SelectTableShow(ShowType)
{
	switch(ShowType)
		{
			case 0://显示回收站注解首页
				document.all.Rec_ZhuJie.style.display="";
				document.all.Rec_Class.style.display="none";
				document.all.Rec_News.style.display="none";
				document.all.Rec_Serch.style.display="none";
				document.all.Rec_Op.style.display="none";
				break;				
			case 1://显示回收站栏目管理首页
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
			case 2://显示回收站新闻管理首页
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
			case 3://显示回收站管理操作首页
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
	if(confirm('此操作将恢复选中的新闻？\n你确定恢复吗？'))
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
			//判断是否选择了记录
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
			if(confirm("确认要删除选中的记录？"))
				document.for_Re_OPType.Hi_Re_OP_Type.value="P_Del";
				document.for_Re_OPType.submit();
		}else
		{
			alert("请选择你要删除的对象");
		}
		//以上判断有没选项选中

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
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<%session.CodePage="936"%>
<!--#include file="FS_Inc/Const.asp" -->
<!--#include file="FS_InterFace/MF_Function.asp" -->
<!--#include file="FS_Inc/Function.asp" -->
<!--#include file="FS_Inc/Func_Page.asp" -->
<%'Copyright (c) 2006 Foosun Inc. 
Server.ScriptTimeOut=999
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache"
response.Charset = "gb2312"
Dim starttime,endtime
starttime=timer()
function morestr(str,length)
	if len(str)>length then 
		morestr = left(str,length)&"<strong>...</strong>"
	else
		morestr = str
	end if	  
end function
Dim Conn,User_Conn,Search_Sql,Search_RS,strShowErr,Cookie_Domain,Cookie_Copyright,Cookie_eMail,Cookie_Site_Name
Dim Server_Name,Server_V1,Server_V2
Dim TmpStr,TmpArr,SqlDateType,FileSize,FileEditDate,TmpStr1
Dim Tags,s_type,SubSys,ClassId,s_date,e_date  ,GetType
Dim ChildDomain,ClassPath

GetType = request.QueryString("GetType") ''�ڲ�
if GetType = "" then response.Write("��ָ����Ҫ�Ĳ���.") : response.End()
''����
If G_IS_SQL_DB = 1 Then  
	SqlDateType = "'"
else
	SqlDateType = "#"
end if

Function Get_MF_Config()
	if request.Cookies("FoosunSearchCookie")("Cookie_Domain") = Get_MF_Domain() then exit Function
	set Search_RS=Conn.execute("select top 1 MF_Domain,MF_Site_Name,MF_eMail,MF_Copyright_Info  from FS_MF_Config")
	Response.Cookies("FoosunSearchCookie")("Cookie_Domain")=Search_RS("MF_Domain") 
	Response.Cookies("FoosunSearchCookie")("Cookie_Copyright")=Search_RS("MF_Copyright_Info") 
	Response.Cookies("FoosunSearchCookie")("Cookie_eMail")=Search_RS("MF_eMail") 
	Response.Cookies("FoosunSearchCookie")("Cookie_Site_Name")=Search_RS("MF_Site_Name") 
	Response.Cookies("FoosunSearchCookie").Expires=Date()+1
	Search_RS.close
End Function
''++++++++++++++++++++++++++++++++++++
'��鱾���ļ� ���ش�С���޸�����
Function CheckFile(PhFileName)
	On Error Resume Next
	FileEditDate="":FileSize=""
	if isnull(PhFileName) or PhFileName="" then CheckFile="|":exit Function
	Dim Fso,MyFile
	Set Fso = CreateObject(G_FS_FSO)
	If Fso.FileExists(server.MapPath(PhFileName)) Then
		set MyFile = Fso.GetFile(server.MapPath(PhFileName))
		FileEditDate = MyFile.DateLastModified
		FileSize = formatnumber(MyFile.Size/1024,1,-2)&"K"
		set MyFile = nothing 
	End if
	if Err<>0 then Err.clear : CheckFile="|":exit Function		
	Set Fso = Nothing
	CheckFile = FileSize&"|"&FileEditDate
End Function

MF_Default_Conn
MF_User_Conn
Get_MF_Config

Cookie_Domain = request.Cookies("FoosunSearchCookie")("Cookie_Domain")
Cookie_Copyright = request.Cookies("FoosunSearchCookie")("Cookie_Copyright")
Cookie_eMail = request.Cookies("FoosunSearchCookie")("Cookie_eMail")
Cookie_Site_Name = request.Cookies("FoosunSearchCookie")("Cookie_Site_Name")

if Cookie_Domain="" then 
	Cookie_Domain = "http://localhost"
else
	if left(lcase(Cookie_Domain),len("http://"))<>"http://" then Cookie_Domain = "http://"&Cookie_Domain
	if right(Cookie_Domain,1)="/" then Cookie_Domain = mid(Cookie_Domain,1,len(Cookie_Domain) - 1)
end if	


''�õ���ر��ֵ��
Function Get_OtherTable_Value(This_Fun_Sql)
	Dim This_Fun_Rs
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
	else
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	end if
	if instr(lcase(This_Fun_Sql)," in ")>0 then 
		do while not This_Fun_Rs.eof
			Get_OtherTable_Value = Get_OtherTable_Value & This_Fun_Rs(0) &"&nbsp;"
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

''������
Dim Main_Name,Name_Str1,V_MainName,V_Str
Server_Name = NoHtmlHackInput(NoSqlHack(LCase(Trim(Request.ServerVariables("SERVER_NAME")))))
IF Server_Name <> LCase(Cookie_Domain) Then
	Response.Write ("û��Ȩ�޷���")
	Response.End
End If
Server_V1 = NoHtmlHackInput(NoSqlHack(Trim(Replace(Lcase(Cstr(Request.ServerVariables("HTTP_REFERER"))),"http://",""))))
Server_V1 = Replace(Replace(Server_V1,"//","/"),"///","/")
IF Server_V1 = "" Then
	Response.Write ("û��Ȩ�޷���")
	Response.End
End If
IF Instr(Server_V1,"/") = 0 Then
	Server_V2 = Server_V1
Else
	Server_V2 = Split(Server_V1,"/")(0)
End If	
If Instr(Server_Name,".") = 0 Then
	Main_Name = Server_Name
Else
	Name_Str1 = Split(Server_Name,".")(0)
	Main_Name = Trim(Replace(Server_Name,Name_Str1 & ".",""))
End If
If Instr(Server_V2,".") = 0 Then
	V_MainName = Server_V2
Else
	V_Str = Split(Server_V2,".")(0)
	V_MainName = Trim(Replace(Server_V2,V_Str & ".",""))
End If
If Main_Name <> V_MainName And (Main_Name = "" OR V_MainName = "") Then
	Response.Write ("û��Ȩ�޷���")
	Response.End
End If

''+++++++++++++++++++++++++++++++++++++++++++
select case GetType
case "LoginHtml"
%>
<FONT size=-1><A href="<%=Cookie_Domain&"/User/Login.asp"%>" target="_blank">��¼</A></FONT>

<%case "FootHTML"%>

&nbsp;<BR>
        <BR>
        <FONT 
      size=-1><%=Cookie_Copyright%></FONT><BR><BR>
	  
<%case "CopyrightHTML"
	TmpStr = "<TABLE cellSpacing=0 cellPadding=2 width=""100%"" border=0>"&vbNewLine _ 
	&"<TR>"&vbNewLine _ 
	&"<TD align=right height=25><font size=-1>"&vbNewLine _
	&"<a class=fl href=""javascript:window.external.AddFavorite('"&Cookie_Domain&"', '"&Cookie_Site_Name&"')"">�����ղ�</a>"&vbNewLine _
	&" - <a class=fl href=""#"" onClick=""this.style.behavior='url(#default#homepage)';this.setHomePage('"&Cookie_Domain&"')"">��Ϊ��ҳ</a>"&vbNewLine _
	&" - <A class=fl href=""#"">Top</A>"&vbNewLine _
	&"</font></TD>"&vbNewLine _
	&"</TR>"&vbNewLine _
	&"</TABLE>"&vbNewLine _ 
	&"</CENTER>"&vbNewLine
	response.Write(TmpStr)  
case "MainInfo"
Tags = NoHtmlHackInput(request.QueryString("Tags"))
if Tags = "" then strShowErr=strShowErr&"<li>�ؼ��ֲ���Ϊ��</li>"&vbnewLine
Tags = replace(Tags,"��",",")
if strShowErr<>"" then strShowErr=strShowErr&"<li><a href="""&Cookie_Domain&""">"&Cookie_Domain&"</a>.</li>": response.Write(strShowErr):response.End()

Search_Sql = "select iLogID,iLogStyle,Title,UserName,HeadPic,Sex,Corner,Province,City,KeyWords,Content,iLogSource,MainID,ClassID,isTF,A.Hits,EmotFace,isTop,TempletID,savePath,FileName,FileExtName,Pic_1,Pic_2,Pic_3,Password,AddTime " _ 
	&" from FS_ME_Infoilog A,FS_ME_Users C where " _
	&" A.UserNumber=C.UserNumber and A.isLock=0 and A.isDraft=0 and A.adminLock=0 "
	TmpArr = split(Tags,",")
	TmpStr1 = ""
	for each TmpStr in TmpArr
		if trim(TmpStr)<>"" then 
			if ubound(TmpArr)>0 then 
				TmpStr1 = TmpStr1 & " or A.KeyWords like '%"&Trim(TmpStr)&"%' "
			else
				TmpStr1 =  " A.KeyWords like '%"&Trim(TmpStr)&"%' "
			end if 		
		end if	
	next
	if ubound(TmpArr)>0 then TmpStr1 = " ("& mid(TmpStr1,len(" or ")) &") " 
	Search_Sql = and_where(Search_Sql) & TmpStr1
On Error Resume Next
'response.Write(Search_Sql) '������
Set Search_RS = CreateObject(G_FS_RS)
Search_RS.Open Search_Sql,User_Conn,1,1	
if Err<>0 then 
	response.Write("<li>��ѯ������ƥ��.�޷�����."&Err.Description&"</li>")
	response.End()
end if
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

IF Search_RS.eof THEN%>
<TABLE class="t bt" cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD noWrap><FONT size=+1>&nbsp;<B><FONT size=+1>&nbsp;<B><%=TmpStr%></B></FONT>&nbsp;</B></FONT>&nbsp;</TD>
    <TD noWrap align=right>
	<FONT size=-1>����<B>0</B>�����<B><%=Tags%></B>�Ĳ�ѯ���
	��������ʱ <B><%=FormatNumber((timer()-starttime),1,-2)%></B>���룩&nbsp;</FONT>
	</TD>
   </TR>
  </TBODY>
</TABLE>
<p><font size=-1 color=#666666>δ��ѯ�����������ļ�¼��</font></p>
<%else
Dim UrlAndTitle,SaveNewsPath,Content,NewsSmallPicFile,NewsPicFile,addtime,NaviContent ,SysRs_Tmp,ChildPath,picShuXing,picShuXingB,NoNewsSmallPicFile	,ClassName
Dim EmotFace
Search_RS.PageSize=int_RPP
cPageNo=CintStr(Request.QueryString("Page"))
If cPageNo="" or not isnumeric(cPageNo) Then cPageNo = 1
cPageNo = Clng(cPageNo)
If cPageNo<1 Then cPageNo=1
If cPageNo>Search_RS.PageCount Then cPageNo=Search_RS.PageCount 
Search_RS.AbsolutePage=cPageNo
  FOR int_Start=1 TO int_RPP 
  
 	ChildPath = Cookie_Domain
	set SysRs_Tmp = User_Conn.execute("select top 1 Dir from FS_ME_iLogSysParam")
	if not SysRs_Tmp.eof then ChildPath = ChildPath & "/"&SysRs_Tmp(0)&"/"
	SysRs_Tmp.close
	
	SaveNewsPath = ChildPath &"blog.asp?id="&Search_RS("iLogID")
	UrlAndTitle = "<A class=l href="""&SaveNewsPath&""" target=_blank>"&Search_RS("Title")&"</A>"
	addtime = Search_RS("AddTime")
	if isnull(addtime) then addtime=""
	if isdate(addtime) then addtime = FormatDateTime(addtime,0)
	NewsSmallPicFile = Search_RS("HeadPic")
	if Search_Rs("Sex")=0 then 
		NoNewsSmallPicFile = "sys_images/man.gif"
	else
		NoNewsSmallPicFile = "sys_images/wom.gif"
	end if	
	if NewsSmallPicFile = "" then 
		NewsSmallPicFile = NoNewsSmallPicFile
	else
		NewsSmallPicFile = Cookie_Domain&NewsSmallPicFile
	end if		
	NaviContent = Search_RS("Content")
	if isnull(NaviContent) or NaviContent="" then 
		NaviContent = "����"
	else
		NaviContent = morestr(Lose_Html(NaviContent),200)
		EmotFace = Search_RS("EmotFace")
		if instr(EmotFace,"sys_images/emot")=0 then EmotFace = "sys_images/emot/"&replace(EmotFace,"/","")
		EmotFace = Replacestr(Search_RS("EmotFace"),":,else:<img border=0 src="""&EmotFace&""" />&nbsp;")
		
		TmpArr = split(Tags,",")
		TmpStr1 = ""
		for each TmpStr in TmpArr
			if trim(TmpStr)<>"" then 
				TmpStr1 = TmpStr1 & replace(NaviContent,Trim(TmpStr),"<font color=red>"&Trim(TmpStr)&"</font>")
			end if	
		next
		NaviContent = EmotFace & TmpStr1
	end if
	TmpStr = Replacestr(Search_RS("iLogStyle"),"0:�ռ�,1:��ժ")

	Content="<TABLE cellSpacing=1 cellPadding=1 border=0 width=""80%"">"&vbNewLine _
		&"<TBODY>"&vbNewLine _
		  &"<TR>"&vbNewLine _ 
			&"<TD style=""width:60px;height:60px"" rowspan=2 align=center>"&vbNewLine 
			picShuXing = CheckFile(NewsSmallPicFile)
			if picShuXing<>"|" then 
				Content=Content&"<img border=0 src="""&NewsSmallPicFile&""" alt=""ͼƬ����:["&picShuXing&"]"" onload=""if(this.offsetWidth>60)this.width=60;""></TD>"&vbNewLine
			else		
				Content=Content&"<img border=0 src="""&NoNewsSmallPicFile&""" onload=""if(this.offsetWidth>60)this.width=60;""></TD>"&vbNewLine
			end if
			picShuXing=""		
	Content=Content	&"<TD class=content valign=top>"&vbNewLine _
			&"<font size=-1>"&NaviContent&"</font>"&vbNewLine _
			&"</TD>"&vbNewLine _
		  &"</TR>"&vbNewLine _
		  &"<TR>"&vbNewLine _ 
			&"<TD height=21><font size=-1>"&vbNewLine _
			&"<font color=#008000>"&SaveNewsPath&" </font>" _
			&"<a class=fl href=""javascript:window.external.AddFavorite('"&SaveNewsPath&"', '"&Cookie_Site_Name&"')"">�����ղ�</a>"&vbNewLine _
			&" - <a class=fl href=""#"" onClick=""this.style.behavior='url(#default#homepage)';this.setHomePage('"&SaveNewsPath&"')"">��Ϊ��ҳ</a>"&vbNewLine _
			&"</font></TD>"&vbNewLine _
		  &"</TR>"&vbNewLine _
		&"</TBODY>"&vbNewLine _
	   &"</TABLE>"&vbNewLine
if int_Start = 1 then%>      
<TABLE class="t bt" cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD noWrap><FONT size=+1>&nbsp;<B><FONT size=+1>&nbsp;<B><%=TmpStr%></B></FONT>&nbsp;</B></FONT>&nbsp;</TD>
    <TD noWrap align=right>
	<FONT size=-1>����<B><%=Search_RS.recordcount%></B>����� <B><%=morestr(Tags,30)%></B> �Ĳ�ѯ�����
	�����ǵ� <B>1</B> - <B>10</B> ���������ʱ <B><%=FormatNumber((timer()-starttime),1,-2)%></B> �룩&nbsp;</FONT>
	</TD></TR></TBODY></TABLE>
<%end if%>
<DIV>
	<div>	
  <P class=g>
  <%
  ''����
  response.Write(UrlAndTitle)
  response.Write("<font size=-2 color=#666666>")
  response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&addtime&vbNewLine)
  response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
  ClassName = Get_OtherTable_Value("select ClassCName from FS_ME_iLogClass where ClassID = "&Search_Rs("ClassID"))
  response.Write(" | ����:"&Replacestr(ClassName,":δ����,else:"&ClassName))
  if Search_Rs("iLogSource")<>"" then response.Write(" | <a href="""&Search_Rs("iLogSource")&""" title="""&Search_Rs("iLogSource")&""" target=_blank>��Դ</a>"&vbNewLine)
  response.Write(" | ����:"&Replacestr(Search_Rs("UserName"),":,else:"&Search_Rs("UserName")))
  response.Write( Replacestr(Search_Rs("Sex"),"1:[Ů],0:[��]") )
  response.Write(" | ����:"&Replacestr(Search_Rs("Corner"),":,else:"&Search_Rs("Corner")))
  response.Write(Replacestr(Search_Rs("Province"),":,else: "&Search_Rs("Province")))
  response.Write(Replacestr(Search_Rs("City"),":,else: "&Search_Rs("City"))  &"]")
  response.Write(Replacestr(Search_Rs("isTF"),"0:,1: | �Ƽ�"))
  response.Write(" | ����:["&Replacestr(Search_Rs("Hits"),":,else:"&Search_Rs("Hits"))&"]")
  response.Write("</font>")
  response.Write(Content)
%>
</div>
<%
	''+++++++++++++++++++++++++++++++++++++++		
	Search_RS.MoveNext
	if Search_RS.eof or Search_RS.bof then exit for
  NEXT
%>
<BR clear=all>
<DIV class=n id=navbar> 
  <TABLE cellSpacing=0 cellPadding=0 width="1%" align=center border=0>
    <TBODY>
      <TR style="TEXT-ALIGN: center" vAlign=top align=middle> 
        <TD vAlign=bottom noWrap class=i><FONT size=-1>���ҳ��:&nbsp;</FONT> 
        <TD noWrap class="i"><font size=-1>&nbsp; 
		<%response.Write( fPageCount(Search_RS,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf )%>
		</font></TR>
    </TBODY>
  </TABLE>
</DIV> 
<%
END IF
RsClose

end select


Sub RsClose()
	Search_RS.Close
	Set Search_RS = Nothing
end Sub

set User_Conn=nothing
Set Conn = Nothing
response.End()
%>






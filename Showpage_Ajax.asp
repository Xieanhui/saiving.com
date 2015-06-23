<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="FS_Inc/Const.asp" -->
<!--#include file="FS_InterFace/MF_Function.asp" -->
<!--#include file="FS_Inc/Function.asp" -->
<!--#include file="FS_InterFace/CLS_Foosun.asp" -->
<%session.CodePage="936"%>
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	response.Charset = "gb2312"
	Dim Conn,User_Conn,Page_Sql,Page_RS,strShowErr,Cookie_Domain
	Dim Server_Name,Server_V1,Server_V2,TmpStr,TmpArr,tmprs_,tmpStr1,TmpStr_2
	Dim stype,Id,PageType,p_NewsLinkFields,p_NewsLink,p_NewsLinkFields
	MF_Default_Conn
	p_NewsLinkFields = "ClassEName,Domain,SavePath,IsURL,URLAddress,SaveNewsPath,FileName,FileExtName"
	Cookie_Domain = Get_MF_Domain()
	if Cookie_Domain="" then
		Cookie_Domain = "localhost"
	else
		if left(lcase(Cookie_Domain),len("http://"))="http://" then Cookie_Domain = mid(Cookie_Domain,len("http://")+1)
		if right(Cookie_Domain,1)="/" then Cookie_Domain = mid(Cookie_Domain,1,len(Cookie_Domain) - 1)
	end if
	''防盗连
	Dim Main_Name,Name_Str1,V_MainName,V_Str
	Server_Name = NoHtmlHackInput(NoSqlHack(LCase(Trim(Request.ServerVariables("SERVER_NAME")))))
	IF Server_Name <> LCase(Split(Cookie_Domain,"/")(0)) Then
		Response.Write ("没有权限访问")
		Response.End
	End If
	Server_V1 = NoHtmlHackInput(NoSqlHack(Trim(Replace(Lcase(Cstr(Request.ServerVariables("HTTP_REFERER"))),"http://",""))))
	Server_V1 = Replace(Replace(Server_V1,"//","/"),"///","/")
	IF Server_V1 = "" Then
		Response.Write ("没有权限访问")
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
		Response.Write ("没有权限访问")
		Response.End
	End If

	stype = NoSqlHack(request.QueryString("type")) 'NS
	Id = NoSqlHack(request.QueryString("Id")) 'NewsId
	PageType = NoSqlHack(request.QueryString("PageType")) 'PageType

	if stype="" then stype="NS"
	if Id="" then call response.Write("Error:Id is null!"):response.End()

	select case stype
		case "NS"
			'同时取栏目ID
			set tmprs_ = Conn.execute("select News.ID,News.ClassID,ClassEName,[Domain],Class.SavePath,News.IsURL,News.URLAddress,SaveNewsPath,News.FileName,News.FileExtName from FS_NS_News As News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And NewsID='"&NoSqlHack(Id)&"'")
			if tmprs_.eof then
				TmpStr = "错误：不存在的新闻ID."
			else
				Set p_NewsLinkRecordSet = New CLS_FoosunRecordSet
				Set p_NewsLinkRecordSet.Values(p_NewsLinkFields) = Page_RS
				p_NewsLink = get_NewsLink(p_NewsLinkRecordSet)
				Set p_NewsLinkRecordSet = Nothing
				If PageType = "PrevPage" Then
					set Page_RS=Conn.execute("select top 1 ID,NewsID,NewsTitle from FS_NS_News where ID < "&tmprs_("ID")&" and ClassID='"&tmprs_("ClassID")&"' And isRecyle=0  order by ID desc")
					if Page_RS.eof then
						TmpStr = "无"
					else
						if isnull(Page_RS("NewsTitle")) or Page_RS("NewsTitle")="" then
							TmpStr = "无标题！"
						else
							tmpStr1 = Page_RS("NewsTitle")
							TmpStr = "<a href="""&p_NewsLink&""">"&tmpStr1&"</a>"
						end if
					end If
				Else
					''''''''''''''''''''''''''''''''''''''''''
					set Page_RS=Conn.execute("select top 1 ID,NewsID,NewsTitle from FS_NS_News where ID > "&tmprs_("ID")&" and ClassID='"&tmprs_("ClassID")&"'  And isRecyle=0  order by ID")
					if Page_RS.eof then
						TmpStr = "无"
					else
						if isnull(Page_RS("NewsTitle")) or Page_RS("NewsTitle")="" then
							TmpStr= "无标题！"
						else
							tmpStr1 = Page_RS("NewsTitle")
							TmpStr = "<a href="""&p_NewsLink&""">"&tmpStr1&"</a>"
						end if
					end If
				End If
				Page_RS.close
			End If
			tmprs_.close
		case else
			response.Write("Error:"&stype&" is not found.")
	end select
	response.Write(TmpStr)
	ConnClose()

	Sub ConnClose()
		Set Conn = Nothing
		response.End()
	End Sub
%>







<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<%
	Response.Expires = 0
	Dim Conn,User_Conn,Old_News_Conn,str_Id,str_ClassId,str_where,strShowErr,SQL_Insearch,timeType,oldlimit,im
	Dim SearchSql,RsNewsObj,RsFileObj,c_rs
	Dim MoveNumber,array_news,i,FiledObj,old_sql
	MF_Default_Conn
	MF_Old_News_Conn
	MF_Session_TF 
	str_Id = server.HTMLEncode(NoSqlHack(Request.QueryString("Id")))
	str_ClassId = NoSqlHack(Request.QueryString("ClassId"))
	if trim(str_Id) = "" then
		im=0
		set c_rs = Conn.execute("select Oldtime From FS_NS_NewsClass where ClassId='"& str_ClassId &"'")
		if Not c_rs.eof then
			if trim(str_ClassId)<>"" then
				SQL_Insearch =" and ClassId='"& str_ClassId &"'"
			else
				SQL_Insearch =""
			end if
			
			if c_rs("Oldtime")=0 then
				oldlimit = 730'如果归档设置为0，则归档730天以前的新闻(2年前的新闻)
			else
				oldlimit = c_rs("Oldtime")
			end if
			If G_IS_SQL_DB = 1 Then
				timeType = " and DATEDIFF (d,addtime,GetDate())>"& oldlimit &""
			Else
				timeType = " and DATEDIFF ('d',addtime,Date())>"& oldlimit &""
			End If
			'Crazy2008.10.31
			Dim old_Obj,old_NewsCount,Maxid,Minid,old_News_Sql
			old_News_Sql = "Select Count(id) from FS_NS_News where ClassId='"& str_ClassId &"' "& timeType &""
			Set old_Obj = Server.CreateObject(G_FS_RS)
			old_Obj.Open old_News_Sql,Conn,1,1
			old_NewsCount = old_Obj(0)
			old_Obj.Close:Set old_Obj = Nothing
			For I = 1 To int(int(old_NewsCount)/200)+1
				SearchSql = "Select top 200 * from FS_NS_News where ClassId='"& str_ClassId &"' "& timeType
				Set RsNewsObj = Server.CreateObject(G_FS_RS)
				RsNewsObj.Open SearchSql,Conn,1,1
				If RsNewsObj.Eof Then
					strShowErr = "<li>该栏目下目前还没有满足归档条件要求的新闻！</li>"
					RsNewsObj.Close:Set RsNewsObj = Nothing
					Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../News_Manage.asp?ClassId="&str_ClassId&"")
					Response.end
				End if
				RsNewsObj.movelast
				Maxid = RsNewsObj(0)
				RsNewsObj.movefirst
				Minid = RsNewsObj(0)
				do while not RsNewsObj.eof 
					Set RsFileObj = Server.CreateObject(G_FS_RS)
					old_sql ="Select * from FS_Old_News Where 1=0"
					RsFileObj.Open old_sql,Old_News_Conn,3,3
					if RsFileObj.Eof then
						RsFileObj.AddNew
						for Each FiledObj in RsNewsObj.Fields
							if LCase(FiledObj.Name) <> "id" then
								RsFileObj(FiledObj.Name) = RsNewsObj(FiledObj.Name)
							end if
							RsFileObj("FileTime")=now()
						Next
						RsFileObj.Update
						RsFileObj.Close:set RsFileObj = nothing
					end if
				RsNewsObj.movenext
				loop
				RsNewsObj.close:set RsNewsObj = nothing

				Conn.Execute("Delete from FS_NS_News where ClassId='"& str_ClassId &"'And Id>"&MinId-1&" And Id<"&Maxid+1)
			Next
		end if
		'Crazy2008.10.31
		c_rs.close:set c_rs= nothing
		strShowErr = "<li>栏目新闻归档成功</li><li>共归档 "& old_NewsCount &" 个新闻到归档数据库中，可以在前台搜索!!</li>"
		set conn=nothing:set Old_News_Conn=nothing
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../News_Manage.asp?ClassId="&str_ClassId&"")
		Response.end
	else
		array_news = split(str_Id,",")
		for i = 0 to Ubound(array_news)
			SearchSql = "Select * from FS_NS_News where ID=" & CintStr(array_news(i))
			Set RsNewsObj = Conn.Execute(SearchSql)
			if not RsNewsObj.eof then
				Set RsFileObj = Server.CreateObject(G_FS_RS)
				old_sql ="Select * from FS_Old_News where NewsId='" & RsNewsObj("NewsId")&"'"
				RsFileObj.Open old_sql,Old_News_Conn,3,3
				if RsFileObj.Eof then
					RsFileObj.AddNew
					for Each FiledObj in RsNewsObj.Fields
						if LCase(FiledObj.Name) <> "id" Then
							RsFileObj(FiledObj.Name) = RsNewsObj(FiledObj.Name)
						end if
						RsFileObj("FileTime")=now()
					Next
					RsFileObj.Update
					RsFileObj.Close:set RsFileObj = nothing
					Conn.Execute("Delete from FS_NS_News where NewsID='" & RsNewsObj("NewsId") &"'")
				end if
			end if
		next
		strShowErr = "<li>选择的新闻归档成功</li><li>共归档 "& i &" 个新闻到归档数据库中，可以在前台搜索!</li>"
		set conn=nothing:set Old_News_Conn=nothing
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../News_Manage.asp?ClassId="&str_ClassId&"")
		Response.end
	end if
	set Conn = nothing
%>






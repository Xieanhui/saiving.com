<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp"-->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_InterFace/NS_Function.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
Dim Str_SysDir
Str_SysDir=""
if G_VIRTUAL_ROOT_DIR<>"" then
	Str_SysDir="/"&G_VIRTUAL_ROOT_DIR&"/Ads"
Else
	Str_SysDir="/Ads"
end if
Dim pic1,rsObj,picSql
pic1=NoSqlHack(request.QueryString("pic"))
set rsObj=server.createobject(G_FS_RS)
picSql="select * from FS_AD_Info where AdID="&CintStr(pic1)&""
rsObj.open picSql,Conn,1,1

If rsobj.eof or Err.number<>0 then
	rsObj.close
	set rsObj=nothing
	Set Conn=Nothing
	Response.Write "参数传递错误!!!"&Err.description
	Response.End()
Else
	If Cint(rsObj("AdLoopFactor"))=1 Then
		If Cint(rsobj("AdLock"))=1 or CLng(rsobj("AdClickNum")) > CLng(rsobj("AdMaxClickNum")) or CLng(rsobj("AdShowNum")) > CLng(rsobj("AdMaxShowNum")) Then
			If Trim(rsobj("AdEndDate"))<>"" And  not IsNull(rsobj("AdEndDate")) Then
				If Cdate(rsobj("AdEndDate"))<Now() Then
					rsObj.close
					set rsObj=nothing
					Set Conn=Nothing
					Response.Write "此广告已经暂停或是失效!"
					Response.End()
				End If		
			End If
			rsObj.close
			set rsObj=nothing
			Set Conn=Nothing
			Response.Write "此广告已经暂停或是失效!"
			Response.End()
		else
			Response.Write "此广告已经暂停或是失效!"
			Response.End()
		End If
	Else
		If	InStr(1,LCase(rsobj("AdPicPath")),".swf",1)<>0 Then
			If InStr(1,LCase(rsobj("AdPicPath")),"http://")=0 then
				Response.Write "<a href="&Conn.execute("Select MF_Domain From FS_MF_Config")(0)&"Ads/AdsClick.asp?Location="& Pic1 &" target=_blank><EMBED src="""& Conn.execute("Select MF_Domain From FS_MF_Config") &"/"&rsObj("AdPicPath") &""" quality=high WIDTH="""& rsObj("AdPicWidth") &""" HEIGHT="""& rsObj("AdPicHeight") &""" TYPE=""application/x-shockwave-flash"" PLUGINSPAGE=""http://www.macromedia.com/go/getflashplayer""></EMBED></a>"
			Else
				Response.Write "<EMBED src="""& rsObj("AdPicPath") &""" quality=high WIDTH="""& rsObj("AdPicWidth") &""" HEIGHT="""& rsObj("AdPicHeight") &""" TYPE=""application/x-shockwave-flash"" PLUGINSPAGE=""http://www.macromedia.com/go/getflashplayer""></EMBED>"
			End If 
		Else 
			Response.Write "<a href=http://"&Conn.execute("Select MF_Domain From FS_MF_Config")(0)&Str_SysDir&"/AdsClick.asp?Location="&Pic1&_
			" target=_blank><img src="&rsobj("AdPicPath")&" border=0 title="&rsobj("AdCaptionTxt")&"></img></a>"
		End If
	End If
End If
rsObj.close
set rsObj=nothing
Set Conn=Nothing
%>







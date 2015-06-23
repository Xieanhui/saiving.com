<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Dim Conn,User_Conn,tmp_type,strShowErr,strpage,ShowStr
	MF_Default_Conn
	MF_IP_Conn
	Dim AddrConn,AddrDBC,AddrConnStr,EnAddress
	Dim code,VisitIP,vSoft,vExplorer,vOS,EnVisitIP,RsCouObj,RsCouSql,vSource,ExpTime
	ExpTime = Conn.Execute("Select ExpTime from FS_SS_SysPara")(0)
	code = NoSqlHack(Request("code"))
	VisitIP = NoSqlHack(request.ServerVariables("HTTP_X_FORWARDED_FOR"))

	If VisitIP = "" then
		VisitIP = NoSqlHack(request.ServerVariables("REMOTE_ADDR"))
	End If
	vSoft = NoSqlHack(Request.ServerVariables("HTTP_USER_AGENT"))
	if vSource="" then
		vSource="直接输入网址进入的"
	else
		vSource=Mid(vSource,8)
		vSource="http://"&Mid(vSource,1,instr(vSource,"/"))
	end if

	if instr(vSoft,"NetCaptor") then
		vExplorer="NetCaptor"
	elseif instr(vSoft,"MSIE 6") then
		vExplorer="Internet Explorer 6.x"
	elseif instr(vSoft,"MSIE 5") then
		vExplorer="Internet Explorer 5.x"
	elseif instr(vSoft,"MSIE 4") then
		vExplorer="Internet Explorer 4.x"
	elseif instr(vSoft,"Netscape") then
		vExplorer="Netscape"
	elseif instr(vSoft,"Opera") then
		vExplorer="Opera"
	else
		vExplorer="Other"
	end if

	if instr(vSoft,"Windows NT 5.0") then
		vOS="Windows 2000"
	elseif instr(vSoft,"Windows NT 5.1") then
		vOS="Windows XP"
	elseif instr(vSoft,"Windows NT 5.2") then
		vOS="Windows 2003"
	elseif instr(vSoft,"Windows NT") then
		vOS="Windows NT"
	elseif instr(vSoft,"Windows 9") then
		vOS="Windows 9x"
	elseif instr(vSoft,"unix") or instr(vSoft,"linux") or instr(vSoft,"SunOS") or instr(vSoft,"SunOS") or instr(vSoft,"BSD") or instr(vSoft,"Mac") then
		vOS="Unix & Unix 类"
	else
		vOS="Other"
	end If
	VisitIP = CheckIpSafe(VisitIP)
	EnAddress = VisitIP
	EnAddress = EnIP(EnAddress)
	EnAddress = EnAddr(EnAddress)
	Set RsCouObj = Conn.Execute("Select ID from FS_SS_Stat where IP='"&NoSqlHack(VisitIP)&"'")
	If RsCouObj.eof then
		Response.Cookies("online") = false
	End if
	RsCouObj.Close
	Set RsCouObj = Nothing

	If code = "2" then
		If request.Cookies("FoosunCookie_stat")("online") <> "true" then
			Set RsCouObj = Server.CreateObject(G_FS_RS)
			RsCouSql = "Select VisitTime,OSType,ExploreType,IP,OSType,Area,Source,LoginNum from FS_SS_Stat where 1=0"
			RsCouObj.Open RsCouSql,Conn,3,3
			RsCouObj.AddNew
			RsCouObj("VisitTime") = Now()
			RsCouObj("OSType") = NoSqlHack(vOS)
			RsCouObj("ExploreType") = NoSqlHack(vExplorer)
			RsCouObj("IP") = NoSqlHack(Request.ServerVariables("Remote_Addr"))
			RsCouObj("OSType") = NoSqlHack(vOS)
			RsCouObj("Area") = NoSqlHack(EnAddress)
			RsCouObj("Source") = NoSqlHack(vSource)
			RsCouObj("LoginNum") = "1"
			RsCouObj.Update
			RsCouObj.Close
			Set RsCouObj = Nothing
		Else
			Conn.Execute("Update FS_SS_Stat Set LoginNum=LoginNum+1 where IP='"&NoSqlHack(VisitIP)&"' and day(VisitTime)='"&day(now())&"' and month(VisitTime)='"&month(now())&"' and year(VisitTime)='"&year(now())&"'")
		End If
		dim TempObj,TempObjs,VisitAllNums,VisitTodayNum
		Set TempObj = Conn.Execute("Select WebCountTime from FS_SS_SysPara")
		If G_IS_SQL_DB=0 then
			Set TempObjs = Conn.Execute("Select Count(ID) from FS_SS_Stat where VisitTime>#"&TempObj("WebCountTime")&"#")
		Else
			Set TempObjs = Conn.Execute("Select Count(ID) from FS_SS_Stat where VisitTime>'"&TempObj("WebCountTime")&"'")
		End if
			VisitAllNums = Clng(TempObjs(0))
		Set TempObjs = Conn.Execute("Select Count(ID) from FS_SS_Stat where day(VisitTime) = '"&Day(Now())&"' and month(VisitTime)='"&Month(Now())&"' and year(VisitTime)='"&Year(Now())&"'")
			VisitTodayNum = Clng(TempObjs(0))
		TempObjs.Close
		Set TempObjs = Nothing
		TempObj.Close
		Set TempObj = Nothing
		ShowStr = "总访问量: " & VisitAllNums & " &nbsp;今日访问: " & VisitTodayNum&""
	ElseIf code = "1" then
		If request.Cookies("FoosunCookie_stat")("online") <> "true" then
			Set RsCouObj = Server.CreateObject(G_FS_RS)
			RsCouSql = "Select VisitTime,OSType,ExploreType,IP,Area,Source,LoginNum from FS_SS_Stat where 1=0"
			RsCouObj.Open RsCouSql,Conn,3,3
			RsCouObj.AddNew
			RsCouObj("VisitTime") = Now()
			RsCouObj("OSType") = NoSqlHack(vOS)
			RsCouObj("ExploreType") = NoSqlHack(vExplorer)
			RsCouObj("IP") = NoSqlHack(Request.ServerVariables("Remote_Addr"))
			RsCouObj("Area") = NoSqlHack(EnAddress)
			RsCouObj("Source") = NoSqlHack(vSource)
			RsCouObj("LoginNum") = "1"
			RsCouObj.Update
			RsCouObj.Close
			Set RsCouObj = Nothing
		Else
			Conn.Execute("Update FS_SS_Stat Set LoginNum=LoginNum+1 where IP='"&NoSqlHack(VisitIP)&"' and day(VisitTime)='"&day(now())&"' and month(VisitTime)='"&month(now())&"' and year(VisitTime)='"&year(now())&"'")
		End If
		dim  tmp_1
		tmp_1 = Replace("/"&G_VIRTUAL_ROOT_DIR &"/stat","//","/")
		ShowStr = "<img src='"& tmp_1 &"/Img/mc.gif' border=0 alt='风讯(www.foosun.cn)统计'>"
	Else
		If request.Cookies("FoosunCookie_stat")("online") <> "true" then
			Set RsCouObj = Server.CreateObject(G_FS_RS)
			RsCouSql = "Select VisitTime,OSType,ExploreType,IP,OSType,Area,Source,LoginNum from FS_SS_Stat where 1=0"
			RsCouObj.Open RsCouSql,Conn,3,3
			RsCouObj.AddNew
			RsCouObj("VisitTime") = Now()
			RsCouObj("OSType") = NoSqlHack(vOS)
			RsCouObj("ExploreType") = NoSqlHack(vExplorer)
			RsCouObj("IP") = NoSqlHack(Request.ServerVariables("Remote_Addr"))
			RsCouObj("OSType") = NoSqlHack(vOS)
			RsCouObj("Area") = NoSqlHack(EnAddress)
			RsCouObj("Source") = NoSqlHack(vSource)
			RsCouObj("LoginNum") = "1"
			RsCouObj.Update
			RsCouObj.Close
			Set RsCouObj = Nothing
		Else
			Conn.Execute("Update FS_SS_Stat Set LoginNum=LoginNum+1 where IP='"&NoSqlHack(VisitIP)&"' and day(VisitTime)='"&day(now())&"' and month(VisitTime)='"&month(now())&"' and year(VisitTime)='"&year(now())&"'")
		End If
		ShowStr = ""
	End If
	Response.Cookies("FoosunCookie_stat")("online") = "true"
	Response.Cookies("FoosunCookie_stat").Expires = DateAdd("n", ExpTime, now())
	Response.Write "document.write(" & chr(34) & ShowStr & chr(34) & ")"
function EnIP(ip)
	dim ip1,ip2,ip3,ip4
	ip=cstr(ip)
	ip1=left(ip,cint(instr(ip,".")-1))
	ip=mid(ip,cint(instr(ip,".")+1))
	ip2=left(ip,cint(instr(ip,".")-1))
	ip=mid(ip,cint(instr(ip,".")+1))
	ip3=left(ip,cint(instr(ip,".")-1))
	ip4=mid(ip,cint(instr(ip,".")+1))
	EnIP=cint(ip1)*256*256*256+cint(ip2)*256*256+cint(ip3)*256+cint(ip4)
end function

Function EnAddr(IP)
	Dim EnAddrObj
    Set EnAddrObj = AddrConn.Execute("select Country,City from Address where StarIP <= "&NoSqlHack(IP)&" and EndIP >= "&NoSqlHack(IP)&"")
	if Not EnAddrObj.Eof then
		EnAddr = EnAddrObj("Country")&EnAddrObj("City")
	else
		EnAddr = "未知区域"
	end if
	EnAddrObj.close
	Set EnAddrObj = Nothing
End Function

Set AddrConn = Nothing
Set Conn = Nothing
%>
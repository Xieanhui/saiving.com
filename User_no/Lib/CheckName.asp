<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="Cls_user.asp" -->
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<%
Dim Conn,User_Conn
MF_Default_Conn
MF_User_Conn
Dim ReturnValue,Username,TempLenth,SysLength,CheckRs,CheckSql
CheckSql="Select LenUserName From FS_ME_SysPara"
Set CheckRs=server.CreateObject(G_FS_RS)
CheckRs.Open CheckSql,User_Conn,1,1
If Not CheckRs.eof then
	SysLength=CintStr(split(CheckRs("LenUserName"),",")(0))
Else
	SysLength=3
ENd If
CheckRs.close
set CheckRs=nothing

Username=Replace(replace(NoSqlHack(Request("Username")),"'","''"),Chr(39),"")
TempLenth=GotLenth(Username)
If TempLenth<SysLength then 
	Response.Write("用户名不能少于"&SysLength&"位")
	Response.end
End if
If Username="" then 
	Response.Write("请填写用户名")
	Response.end
End if
dim Fs_User
Set Fs_User = New Cls_User
ReturnValue = Fs_User.checkName(Username)
if ReturnValue then
	Response.Write(""& Username &" 此用户可以注册")
	Response.end
Else
	Response.Write(""& Username &" 已经被注册或者包含禁止注册字符")
	Response.end
End if
set Fs_User=nothing


Function GotLenth(Str)
	Dim l,t,c, i,LableStr,regEx,Match,Matches
	if IsNull(Str) then
		GotLenth =0
		Exit Function
	end if
	if Str = "" then
		GotLenth=0
		Exit Function
	end If
	Str=Replace(Replace(Replace(Replace(Str,"&nbsp;"," "),"&quot;",Chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
	Next
	GotLenth = t
End Function
%>







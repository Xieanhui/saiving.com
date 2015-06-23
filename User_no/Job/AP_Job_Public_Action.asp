<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<% 

Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

if not Session("FS_UserNumber")<>"" then response.Redirect("../lib/error.asp?ErrCodes=<li>你尚未登陆,或过期.</li>&ErrorUrl=../login.asp") : response.End()

Dim Ap_Rs,Ap_Sql
select case request.QueryString("Act")
	case "Del"
	Del
	case "Save"
	Save
	case else
	response.Redirect("job_applications.asp")
end select

''得到相关表的值。
Function Get_OtherTable_Value(This_Fun_Sql)
	Dim This_Fun_Rs
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
	else
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	end if			
	if not This_Fun_Rs.eof then 
		Get_OtherTable_Value = This_Fun_Rs(0)
	else
		Get_OtherTable_Value = ""
	end if
	if Err.Number>0 then 
		Err.Clear
		response.Redirect("../lib/error.asp?ErrCodes=<li>Get_OtherTable_Value未能得到相关数据。错误描述："&Err.Description&"</li>") : response.End()
	end if
	set This_Fun_Rs=nothing 
End Function
  
Sub Del()
	Dim Str_Tmp
	if request.QueryString("PID")<>"" then 
		Conn.execute("Delete from FS_AP_Job_Public where PID = "&CintStr(request.QueryString("PID")))
	else
		Str_Tmp = request.form("PID")
		if Str_Tmp="" then response.Redirect("../lib/error.asp?ErrCodes=<li>你必须至少选择一个进行删除。</li>")
		Str_Tmp = replace(Str_Tmp," ","")
		Conn.execute("Delete from FS_AP_Job_Public where PID in ("&FormatIntArr(Str_Tmp)&")")
	end if
	response.Redirect("../lib/Success.asp?ErrorUrl="&server.URLEncode( "../job/AP_Job_Public_List.asp" )&"&ErrCodes=<li>恭喜，删除成功。</li>")
End Sub
''================================================================

Sub Save()
	Dim Str_Tmp,Arr_Tmp,PID
	Str_Tmp = "JobName,JobDescription,ResumeLang,WorkCity,PublicDate,EndDate,NeedNum,jlmode,EducateExp,Sex,WorkAge,Age,JobType,OtherJobDes,MoneyMonth,FreeMoney,OtherMoneyDes,HolleType"
	Arr_Tmp = split(Str_Tmp,",")
	PID = CintStr(request.Form("PID"))
	if not isnumeric(PID) or PID = "" then PID = 0 
	Ap_Sql = "select UserNumber,"&Str_Tmp&"  from FS_AP_Job_Public  where UserNumber = '"&Session("FS_UserNumber")&"' and PID="&PID
	Set Ap_Rs = CreateObject(G_FS_RS)
	Ap_Rs.Open Ap_Sql,Conn,1,3
	if PID > 0 then 
	'修改
	on error resume next
		for each Str_Tmp in Arr_Tmp
			'response.Write(Str_Tmp&":"&NoSqlHack(request.Form(Str_Tmp))&"<br>")
			Ap_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
		next
		'response.End()	
		Ap_Rs.update
		Ap_Rs.close
		response.Redirect("../lib/Success.asp?ErrorUrl="&server.URLEncode( "../job/AP_Job_Public_AddUpdate.asp?Act=Edit&PID="&PID )&"&ErrCodes=<li>恭喜，修改成功。</li>")
	else
	'新增 允许重复
		on error resume next
		Ap_Rs.addnew
		Ap_Rs("UserNumber") = Session("FS_UserNumber")
		for each Str_Tmp in Arr_Tmp
			Ap_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
		next
		Ap_Rs.update
		Ap_Rs.close
		response.Redirect("../lib/Success.asp?ErrorUrl="&server.URLEncode( "../job/AP_Job_Public_AddUpdate.asp?Act=Add" ) &"&ErrCodes=<li>恭喜，新增成功。</li>")
	end if
End Sub

Set Ap_Rs=nothing
Conn.close
%>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






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

Dim Ap_Rs,Ap_Sql,UserNumber
UserNumber = Session("FS_UserNumber")
if not UserNumber<>"" then response.Redirect("../lib/error.asp?ErrCodes=<li>你尚未登陆,或过期.</li>&ErrorUrl=../login.asp") : response.End()
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
	if UserNumber<>"" then 
		Conn.execute("Delete from FS_AP_UserList where UserNumber = '"&NoSqlHack(UserNumber)&"'")
	end if
	response.Redirect("../lib/Success.asp?ErrorUrl="&server.URLEncode( "../job/AP_Job_Public_List.asp" )&"&ErrCodes=<li>恭喜，你撤消申请处理成功。</li>")
End Sub
''================================================================

Sub Save()
	Dim Str_Tmp,Arr_Tmp
	if request.Form("GroupLevel") = "1" then 
		Str_Tmp = "UserClass,GroupLevel,CompanyName,Introduct,DocExt,DescriptApply,Phone,Email"
	else
		Str_Tmp = "UserClass,GroupLevel,BeginDate,EndDate,CompanyName,Introduct,DocExt,DescriptApply,Phone,Email"
	end if	
	Arr_Tmp = split(Str_Tmp,",")	
	Ap_Sql = "select UserNumber,OrderID,Audited,"&Str_Tmp&"  from FS_AP_UserList  where UserNumber = '"&UserNumber&"'"
	'response.Write(Ap_Sql)
	Set Ap_Rs = CreateObject(G_FS_RS)
	Ap_Rs.Open Ap_Sql,Conn,1,3
	if not Ap_Rs.eof then 
	''修改
		for each Str_Tmp in Arr_Tmp
			if request.Form(Str_Tmp)<>"" then Ap_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
		next
		Ap_Rs.update
		Ap_Rs.close
		response.Redirect("../lib/Success.asp?ErrorUrl="&server.URLEncode( "../job/AP_Register.asp" )&"&ErrCodes=<li>恭喜，修改成功。</li>")
	else
	''新增 允许重复
		Ap_Rs.addnew
		Ap_Rs("UserNumber") = UserNumber
		Ap_Rs("OrderID") = 0
		Ap_Rs("Audited") = 0
		for each Str_Tmp in Arr_Tmp
			'response.Write(Str_Tmp&":"&NoSqlHack(request.Form(Str_Tmp))&"<br>")
			Ap_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
		next
		'response.End()	
		Ap_Rs.update
		Ap_Rs.close
		response.Redirect("../lib/Success.asp?ErrorUrl="&server.URLEncode( "../job/AP_Register.asp" ) &"&ErrCodes=<li>恭喜，新增成功。</li>")
	end if
End Sub

Set Ap_Rs=nothing
Conn.close
%>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






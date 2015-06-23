<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp"-->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_InterFace/NS_Function.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.CacheControl = "no-cache"
Dim obj_form_rs,form_sql,userGroup_Sql,obj_userGroup_Rs,i,Conn,VerifyLogin,IsOutDate,Validate,TimeLimited
Dim CustomFormID,formName,tableName,FormStyleID,DataStyleID,TextCSS,SelectCSS,OtherCSS,StyleRS,FormStyleContent,DataStyleContent,StartTime,EndTime
MF_Default_Conn

CustomFormID = NoSqlHack(request.QueryString("CustomFormID"))
FormStyleID = NoSqlHack(request.QueryString("FormStyleID"))
DataStyleID = NoSqlHack(request.QueryString("DataStyleID"))
TextCSS = NoSqlHack(request.QueryString("TextCSS"))
if TextCSS <> "" then TextCSS = " Class=""" & TextCSS & """"
SelectCSS = NoSqlHack(request.QueryString("SelectCSS"))
if SelectCSS <> "" then SelectCSS = " Class=""" & SelectCSS & """"
OtherCSS = NoSqlHack(request.QueryString("OtherCSS"))
if OtherCSS <> "" then OtherCSS = " Class=""" & OtherCSS & """"
if CustomFormID = "" then
	Response.Write("document.write('调用表单参数传递错误');" & vbcrlf)
	Response.End
end if
form_sql="select id,formName,tableName,VerifyLogin,StartTime,EndTime,Validate,TimeLimited from FS_MF_CustomForm where state=0 and id=" & CustomFormID
set obj_form_rs=conn.execute(form_sql)
if obj_form_rs.eof then
	obj_form_rs.Close
	Set obj_form_rs = Nothing
	Response.Write("document.write('调用表单不存在');" & vbcrlf)
	Response.End
else
	TimeLimited=obj_form_rs("TimeLimited") & ""
	if TimeLimited = 1 then
		StartTime=obj_form_rs("StartTime") & ""
		EndTime=obj_form_rs("EndTime") & ""
		if IsDate(StartTime) And IsDate(EndTime) then
			StartTime = CDate(StartTime)
			EndTime = CDate(EndTime)
			IsOutDate = (DateDiff("d",StartTime,Now) < 0) OR (DateDiff("d",Now,EndTime) < 0)
			if IsOutDate then
				obj_form_rs.Close
				Set obj_form_rs = Nothing
				Response.Write("document.write('表单调用时间过期');" & vbcrlf)
				Response.End
			end if
		end if
	end if
	formName=obj_form_rs("formName")
	tableName=obj_form_rs("tableName")
	VerifyLogin=obj_form_rs("VerifyLogin")
	Validate=obj_form_rs("Validate")
end if
obj_form_rs.Close
Set obj_form_rs = Nothing

if FormStyleID <> "" then
	FormStyleID = Trim(FormStyleID)
	Set StyleRS = Conn.Execute("Select Content from FS_MF_Labestyle Where ID=" & FormStyleID)
	if Not StyleRS.Eof then FormStyleContent = StyleRS("Content") else FormStyleContent = ""
	StyleRS.Close
	Set StyleRS = Nothing
end if
if DataStyleID <> "" then
	FormStyleID = Trim(FormStyleID)
	Set StyleRS = Conn.Execute("Select Content from FS_MF_Labestyle Where ID=" & DataStyleID)
	if Not StyleRS.Eof then DataStyleContent = StyleRS("Content") else DataStyleContent = ""
	StyleRS.Close
	Set StyleRS = Nothing
end if

Dim CustomFormHeader,CustomFormTailor,FormContent,DataContent,OneDataContent,DataField,ValidateCode
if FormStyleContent <> "" then
	CustomFormHeader="<form name=""" & formName & """ method=""post"" enctype=""multipart/form-data"" action=""" & Replace("/" & G_VIRTUAL_ROOT_DIR & "/customform/CustomFormSubmit.asp","//","/") & """>"
	CustomFormHeader = CustomFormHeader & "<input type=""hidden"" name=""formid"" value=""" & CustomFormID & """ >"
	CustomFormTailor = "</form>"
	FormContent = FormStyleContent
	FormContent = Replace(FormContent,"{CustomFormHeader}",CustomFormHeader)
	FormContent = Replace(FormContent,"{CustomFormTailor}",CustomFormTailor)
	ValidateCode = ""
	if Validate = 1 then ValidateCode = "<input name=""VerifyCode"" type=""text"" " & TextCSS & " id=""VerifyCode"" size=""10"" maxlength=""10""/>&nbsp;<img src=""" & Replace("/" & G_VIRTUAL_ROOT_DIR & "/customform/Validate.asp?","//","/") & """ onClick=""this.src+=Math.random()"" alt=""图片看不清？点击重新得到验证码"" style=""cursor:hand;"">"
	FormContent = Replace(FormContent,"{CustomFormValidate}",ValidateCode)
	
	form_sql = "select formitemid,ItemName,FieldName,IsNull,ItemType,MaxSize,DefaultValue,SelectItem,Remark from FS_MF_CustomForm_Item where formid="&CustomFormID&" and State=0 order by orderby"
	set obj_form_rs=conn.execute(form_sql)
	do while not obj_form_rs.eof
		FormContent = Replace(FormContent,"{CustomForm_" & obj_form_rs("FieldName") & "}",GetFieldStr(obj_form_rs))
		obj_form_rs.movenext
	loop
	obj_form_rs.Close
	Set obj_form_rs = Nothing
end if
FormContent = Replace(FormContent,Chr(13) & Chr(10),"")
FormContent = Replace(FormContent,"'","\'")

if DataStyleContent <> "" then
	form_sql = "Select * from " & tableName & " Where form_lock=0"
	Set obj_form_rs = Conn.Execute(form_sql)
	do while Not obj_form_rs.Eof
		OneDataContent = DataStyleContent
		For Each DataField In obj_form_rs.Fields
			OneDataContent = Replace(OneDataContent,"{CustomFormData_" & DataField.Name & "}",DataField.Value & "")
		Next
		DataContent = DataContent & OneDataContent
		obj_form_rs.MoveNext
	Loop
	obj_form_rs.Close
	Set obj_form_rs = Nothing
end if
DataContent = Replace(DataContent,Chr(13) & Chr(10),"")
DataContent = Replace(DataContent,"'","\'")
Response.Write("document.write('" & FormContent & "');" & vbcrlf)
Response.Write("document.write('" & DataContent & "');" & vbcrlf)

Function GetFieldStr(f_RS)
	Dim f_ItemType,f_SelectItem,f_selectItemArr
	f_ItemType = f_RS("ItemType")
	f_SelectItem = f_RS("SelectItem")
	select case f_ItemType
		case "SingleLineText" '单行文本
			GetFieldStr = GetFieldStr & "<input type=""text"" " & TextCSS & " name=""" & f_RS("FieldName") & """ value=""" & f_RS("DefaultValue") & """"
			if f_RS("MaxSize") <> 0 and cstr(f_RS("MaxSize")) <> "" then
				GetFieldStr = GetFieldStr & " maxsize=""" & f_RS("MaxSize") & """"
			end if
			GetFieldStr = GetFieldStr & ">" & f_RS("Remark") & ""
		case "MultiLineText" '多行文本
			GetFieldStr = GetFieldStr & "<textarea " & TextCSS & " name=""" & f_RS("FieldName") & """ cols=""40"" rows=""8"">" & f_RS("DefaultValue") & "</textarea>" & f_RS("Remark") & ""
		case "PassWordText" '密码
			GetFieldStr=GetFieldStr&"<input type=""password"" " & TextCSS & " name="""&f_RS("FieldName")&""" value="""&f_RS("DefaultValue")&""""
			if f_RS("MaxSize")<>0 and cstr(f_RS("MaxSize"))<>"" then
				GetFieldStr=GetFieldStr&" maxsize="""&f_RS("MaxSize")&""""
			end if
			GetFieldStr=GetFieldStr&">"&f_RS("Remark")&""
		case "DateTime" '日期时间
				GetFieldStr=GetFieldStr&"<input type=""text"" name="""&f_RS("FieldName")&""" " & TextCSS & " readOnly value="""&f_RS("DefaultValue")&"""> <input name="""&f_RS("FieldName")&"btn"" type=""button"" value=""选择时间"" onClick=""OpenWindowAndSetValue('"&replace("/"&G_VIRTUAL_ROOT_DIR&"/"&G_ADMIN_DIR&"/CommPages/SelectDate.asp","//","/")&"',300,130,window,document.all."&f_RS("FieldName")&");"" >"&f_RS("Remark")&""
		case "RadioBox" '单选项
			f_selectItemArr=split(f_SelectItem,Chr(13)&Chr(10))
			if isarray(f_selectItemArr) then 
				if f_RS("Remark")<>"" then
					GetFieldStr=GetFieldStr&""&f_RS("Remark")&"<br>"
				end if
			end if
			for i=0 to ubound(f_selectItemArr)
				if trim(f_selectItemArr(i))=f_RS("DefaultValue") then
					GetFieldStr=GetFieldStr&"<input " & OtherCSS & " type=""radio"" name="""&f_RS("FieldName")&""" value="""&f_selectItemArr(i)&""" checked>"&f_selectItemArr(i)&""
				else
					GetFieldStr=GetFieldStr&"<input " & OtherCSS & " type=""radio"" name="""&f_RS("FieldName")&""" value="""&f_selectItemArr(i)&""" >"&f_selectItemArr(i)&""
				end if
			next
		case "CheckBox" '多选项
			f_selectItemArr=split(f_SelectItem,Chr(13)&Chr(10))
			if isarray(f_selectItemArr) then 
				if f_RS("Remark")<>"" then
					GetFieldStr=GetFieldStr&""&f_RS("Remark")&"<br>"
				end if
			end if
			for i=0 to ubound(f_selectItemArr)
				if trim(f_selectItemArr(i))=f_RS("DefaultValue") then
					GetFieldStr=GetFieldStr&"<input " & OtherCSS & " type=""checkbox"" name="""&f_RS("FieldName")&""" value="""&f_selectItemArr(i)&""" checked>"&f_selectItemArr(i)&""
				else
					GetFieldStr=GetFieldStr&"<input " & OtherCSS & " type=""checkbox"" name="""&f_RS("FieldName")&""" value="""&f_selectItemArr(i)&""" >"&f_selectItemArr(i)&""
				end if
			next
		case "Numberic" '数字
			GetFieldStr=GetFieldStr&"<input type=""text"" " & TextCSS & " name="""&f_RS("FieldName")&""" value="""&f_RS("DefaultValue")&""""
			if f_RS("MaxSize")<>0 and cstr(f_RS("MaxSize"))<>"" then
				GetFieldStr=GetFieldStr&" maxsize="""&f_RS("MaxSize")&""""
			end if
			GetFieldStr=GetFieldStr&" onKeyUp=""value=value.replace(/[^0-9]/g,'') "">"&f_RS("Remark")&""
		case "UploadFile" '附件
			GetFieldStr=GetFieldStr&"<input type=""file"" " & TextCSS & " name="""&f_RS("FieldName")&""" value="""&f_RS("DefaultValue")&""">"&f_RS("Remark")&""
		case "DropList" '下拉框
			f_selectItemArr=split(f_SelectItem,Chr(13)&Chr(10))
			if isarray(f_selectItemArr) then 
				if f_RS("Remark")<>"" then
					GetFieldStr=GetFieldStr&""&f_RS("Remark")&"<br>"
				end if
				GetFieldStr=GetFieldStr&"<select " & SelectCSS & " name="""&f_RS("FieldName")&""">"
			end if
			for i=0 to ubound(f_selectItemArr)
				if trim(f_selectItemArr(i))=f_RS("DefaultValue") then
					GetFieldStr=GetFieldStr&"<option  value="""&f_selectItemArr(i)&""" checked>"&f_selectItemArr(i)&"</option>"
				else
					GetFieldStr=GetFieldStr&"<option  value="""&f_selectItemArr(i)&""" >"&f_selectItemArr(i)&"</option>"
				end if
			next
			if isarray(f_selectItemArr) then 
				GetFieldStr=GetFieldStr&"</select>"
			end if
		case "List" '列表框
			f_selectItemArr=split(f_SelectItem,Chr(13)&Chr(10))
			if isarray(f_selectItemArr) then 
				if f_RS("Remark")<>"" then
					GetFieldStr=GetFieldStr&""&f_RS("Remark")&"<br>"
				end if
				GetFieldStr=GetFieldStr&"<select " & SelectCSS & " name="""&f_RS("FieldName")&""" size=""8"" style=""height:150px"">"
			end if
			for i=0 to ubound(f_selectItemArr)
				if trim(f_selectItemArr(i))=f_RS("DefaultValue") then
					GetFieldStr=GetFieldStr&"<option  value="""&f_selectItemArr(i)&""" checked>"&f_selectItemArr(i)&"</option>"
				else
					GetFieldStr=GetFieldStr&"<option  value="""&f_selectItemArr(i)&""" >"&f_selectItemArr(i)&"</option>"
				end if
			next
			if isarray(f_selectItemArr) then 
				GetFieldStr=GetFieldStr&"</select>"
			end if
	end select
End Function
%>
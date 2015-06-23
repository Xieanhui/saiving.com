<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.CacheControl = "no-cache"
Dim Conn,CharIndexStr,strShowErr
Dim obj_form_rs,form_sql,userGroup_Sql,obj_userGroup_Rs
MF_Default_Conn
MF_Session_TF 
dim act,formid,formName,tableName,upfileSaveUrl,upfileSize,stateSet,TimeLimited,StartTime,EndTime,SubmitType,GoldFactor,PointFactor,UserGroup,UserOnce,Validate,remark,ArrUserGroup,VerifyLogin,DataInitStatus
act=NoSqlHack(Request.QueryString("act"))
formid=NoSqlHack(Request.Form("formid"))
formName=NoSqlHack(Request.Form("formName"))
tableName=NoSqlHack(Request.Form("tableName"))
tableName="FS_MF_CustomForm_"&tableName
upfileSaveUrl=NoSqlHack(Request.Form("upfileSaveUrl"))
upfileSize=NoSqlHack(Request.Form("upfileSize"))
if upfileSize = "" then upfileSize = "1"
stateSet=NoSqlHack(Request.Form("stateSet"))
TimeLimited=NoSqlHack(Request.Form("TimeLimited"))
StartTime=NoSqlHack(Request.Form("StartTime"))
EndTime=NoSqlHack(Request.Form("EndTime"))
SubmitType=NoSqlHack(Request.Form("SubmitType"))
GoldFactor=NoSqlHack(Request.Form("GoldFactor"))
PointFactor=NoSqlHack(Request.Form("PointFactor"))
UserGroup=NoSqlHack(Request.Form("UserGroup"))
UserOnce=NoSqlHack(Request.Form("UserOnce"))
if UserOnce="" then UserOnce=1
Validate=NoSqlHack(Request.Form("Validate"))
VerifyLogin=NoSqlHack(Request.Form("VerifyLogin"))
DataInitStatus=NoSqlHack(Request.Form("DataInitStatus"))
if Validate="" then Validate=0
if VerifyLogin="" then VerifyLogin=0
if DataInitStatus="" then DataInitStatus=0
Remark=NoSqlHack(Request.Form("remark"))
if formName = "" then
	strShowErr = "<li>请填写表单名！</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
end if
if tableName = "" then
	strShowErr = "<li>请填写表名！</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
end if
if act="edit" then
	if not MF_Check_Pop_TF("MF098") then Err_Show
	if formid = "" then
		strShowErr = "<li>参数传递错误！</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	set obj_form_rs=conn.execute("select tableName from FS_MF_CustomForm where ID="&formid&"")
	if obj_form_rs.eof then	
		obj_form_rs.Close
		Set obj_form_rs = Nothing
		strShowErr = "<li>修改的数据不存在！</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	obj_form_rs.Close
	Set obj_form_rs = Nothing
	form_sql="update FS_MF_CustomForm set upfileSize="&upfileSize
	form_sql=form_sql&",state="&stateSet
	form_sql=form_sql&",TimeLimited="&TimeLimited
	form_sql=form_sql&",StartTime='"&StartTime&"'"
	form_sql=form_sql&",EndTime='"&EndTime&"'"
	form_sql=form_sql&",SubmitType="&SubmitType
	form_sql=form_sql&",GoldFactor="&GoldFactor
	form_sql=form_sql&",PointFactor="&PointFactor
	form_sql=form_sql&",UserGroup='"&UserGroup&"'"
	form_sql=form_sql&",UserOnce="&UserOnce
	form_sql=form_sql&",Validate="&Validate
	form_sql=form_sql&",VerifyLogin="&VerifyLogin
	form_sql=form_sql&",DataInitStatus="&DataInitStatus
	form_sql=form_sql&",remark='"&remark&"'"
	form_sql=form_sql&" where formName='"&formName&"' And id="&formid
else
	if not MF_Check_Pop_TF("MF099") then Err_Show
	set obj_form_rs=conn.execute("select tableName from FS_MF_CustomForm where tableName='"&tableName&"'")
	if not obj_form_rs.eof then	
		obj_form_rs.Close
		Set obj_form_rs = Nothing
		strShowErr = "<li>数据表已经存在，请重新取名！</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	obj_form_rs.Close
	Set obj_form_rs = Nothing
	form_sql="create table "&tableName&"([id] int IDENTITY (1,1) PRIMARY KEY"
	form_sql=form_sql&",[form_usernum] int default 0"
	form_sql=form_sql&",[form_username] nvarchar(50) null"
	form_sql=form_sql&",[form_ip] nvarchar(100) null"
	form_sql=form_sql&",[form_time] datetime null"
	form_sql=form_sql&",[form_answer] ntext null"
	form_sql=form_sql&",[form_lock] int default 0)"
	
	conn.execute(form_sql)
	if Err then
		err.Clear
		strShowErr = "<li>创建数据表时发生错误，请与数据库管理员联系。确认是否有权限！</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	
	form_sql="insert into FS_MF_CustomForm(formName,tableName,upfileSaveUrl,upfileSize,state,TimeLimited,StartTime,EndTime,SubmitType,GoldFactor,PointFactor,UserGroup,UserOnce,Validate,VerifyLogin,DataInitStatus,remark) values('"&formName&"'"
	form_sql=form_sql&",'"&tableName&"'"
	form_sql=form_sql&",'"&upfileSaveUrl&"'"
	form_sql=form_sql&","&upfileSize
	form_sql=form_sql&","&stateSet
	form_sql=form_sql&","&TimeLimited
	if G_IS_SQL_DB = 1 then
		form_sql=form_sql&",'"&StartTime&"'"
		form_sql=form_sql&",'"&EndTime&"'"
	else
		if StartTime<>"" then
			form_sql=form_sql&",#"&StartTime&"#"
		else
			form_sql=form_sql&",now()"
		end if
		if EndTime<>"" then
			form_sql=form_sql&",#"&EndTime&"#"
		else
			form_sql=form_sql&",now()"
		end if
	end if
	form_sql=form_sql&","&SubmitType
	form_sql=form_sql&","&GoldFactor
	form_sql=form_sql&","&PointFactor
	form_sql=form_sql&",'"&UserGroup&"'"
	form_sql=form_sql&","&UserOnce
	form_sql=form_sql&","&Validate
	form_sql=form_sql&","&VerifyLogin
	form_sql=form_sql&","&DataInitStatus
	form_sql=form_sql&",'"&remark&"')"
end if
conn.execute(form_sql)

strShowErr = "<li>恭喜，自定义表单保存成功!</li>"
Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../FormManage.asp")
Response.end
%>
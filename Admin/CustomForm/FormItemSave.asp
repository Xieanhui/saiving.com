<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
on error resume next
Response.Buffer = True
Response.Expires = -1
Response.CacheControl = "no-cache"
Dim Conn,strShowErr
Dim obj_form_rs,form_sql,userGroup_Sql,obj_userGroup_Rs
dim act,itemID,formid,ItemName,FieldName,orderby,StateSet,IsNullset,ItemType,MaxSize,DefaultValue,SelectItem,Remark,tableName
MF_Default_Conn
MF_Session_TF 
act=NoSqlHack(Request.QueryString("act"))
itemID=NoSqlHack(Request.Form("itemID"))
formid=NoSqlHack(Request.Form("formid"))
ItemName=NoSqlHack(Request.Form("ItemName"))
FieldName=NoSqlHack(Request.Form("FieldName"))
orderby=NoSqlHack(Request.Form("orderby"))
StateSet=NoSqlHack(Request.Form("StateSet"))
IsNullset=NoSqlHack(Request.Form("IsNullset"))
ItemType=NoSqlHack(Request.Form("ItemType"))
MaxSize=NoSqlHack(Request.Form("MaxSize"))
DefaultValue=NoSqlHack(Request.Form("DefaultValue"))
SelectItem=NoSqlHack(Request.Form("SelectItem"))
Remark=NoSqlHack(Request.Form("Remark"))

form_sql="select tableName from FS_MF_CustomForm where id="&formid
set obj_form_rs=conn.execute(form_sql)
if obj_form_rs.eof then
	obj_form_rs.Close
	Set obj_form_rs = Nothing
	strShowErr = "<li>操作的数据不正确！</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
else
	tableName=obj_form_rs(0)
end if
obj_form_rs.Close
Set obj_form_rs = Nothing
if act="edit" then
	if not MF_Check_Pop_TF("MF093") then Err_Show
	form_sql="update FS_MF_CustomForm_Item set orderby='"&orderby&"'"
	form_sql=form_sql&",State="&StateSet
	form_sql=form_sql&",Remark='"&Remark&"' where FormItemID="&itemID
else
	if not MF_Check_Pop_TF("MF094") then Err_Show
	if MaxSize="" then
		strShowErr = "<li>文本长度不能为空！</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	if Not IsNumeric(orderby) then
		strShowErr = "<li>排序序号只能为数字！</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	form_sql="ALTER TABLE ["&tableName&"] ADD ["&FieldName&"]"
	if Not IsNumeric(Maxsize) then Maxsize = 0
	Maxsize = CInt(Maxsize)
	if Maxsize<>0 and MaxSize<=4000 then
		form_sql=form_sql&" Text("&Maxsize&")"
	else
		form_sql=form_sql&" Memo"
	end if
	if isnullSet=1 then
		form_sql=form_sql&" NULL"
	end if
	'不能给字段直接加默认值 ，否则删除时将因为默认的约束无法删除字段
	'if DefaultValue<>"" then
	'	form_sql=form_sql&" DEFAULT '"&DefaultValue&"'"
	'end if
	conn.execute(form_sql)
	if Err then
		strShowErr = "<li>操作数据表时发生错误！</li>"
		strShowErr = strShowErr&"<li>"&err.description&"</li>"
		err.Clear
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	form_sql="insert into FS_MF_CustomForm_Item(formid,ItemName,FieldName,orderby,State,IsNull,ItemType,MaxSize,DefaultValue,SelectItem,Remark) values("&formid
	form_sql=form_sql&",'"&ItemName&"'"
	form_sql=form_sql&",'"&FieldName&"'"
	form_sql=form_sql&","&orderby
	form_sql=form_sql&","&stateSet
	form_sql=form_sql&","&IsNullset
	form_sql=form_sql&",'"&ItemType&"'"
	form_sql=form_sql&","&MaxSize
	form_sql=form_sql&",'"&DefaultValue&"'"
	form_sql=form_sql&",'"&SelectItem&"'"
	form_sql=form_sql&",'"&remark&"')"
end if
conn.execute(form_sql)
Set Conn = Nothing
strShowErr = "<li>恭喜，自定义表单保存成功!</li>"
Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../FormItem.asp?formid="&formid)
Response.end
%>
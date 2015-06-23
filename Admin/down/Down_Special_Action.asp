<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Dim Conn,special_rs,specialid,action,tmp_rs
Dim SpecialEName,SpecialCName,SpecialTemplet,IsUrl,Addtime,Domain,IsLimited,naviPic,isLock,naviText,FileExtName,Savepath
MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("DS018") then Err_Show
action=NoSqlHack(request.QueryString("act"))
specialid=NoSqlHack(Request.QueryString("specialid"))
'锁定
Response.Charset="GB2312"
if action="lock" then
	Conn.execute("Update FS_DS_Special set isLock=1 where specialID="&CintStr(specialid))
	Response.Write("<a href='#' onclick=changeLockState(false,'"&specialid&"') style='color:red'>解锁</a>")
	response.End()
elseif action="unlock" then
	Conn.execute("Update FS_DS_Special set isLock=0 where specialID="&CintStr(specialid))
	Response.Write("<a href='#' onclick=changeLockState(true,'"&specialid&"')>锁定</a>")
	response.End()
elseif action="checkename" then
	Dim ename
	ename=request.QueryString("ename")
	Set tmp_rs=Conn.execute("select specialID from FS_DS_Special where SpecialEName='"&NoSqlHack(ename)&"'")
	if not tmp_rs.eof then
		Response.Write("重复，请重新填写")
	else
		Response.Write("可以使用")
	End if
	response.End()
elseif action="addaction" then
	SpecialEName=request.Form("SpecialEName")
	SpecialCName=request.Form("SpecialCName")
	SpecialTemplet=request.Form("SpecialTemplet")
	IsUrl=request.Form("IsUrl")
	Addtime=Now
	Domain=request.Form("Domain")
	IsLimited=request.Form("IsLimited")
	naviPic=request.Form("NaviPic")
	isLock=request.Form("isLock")
	naviText=request.Form("naviText")
	FileExtName=request.Form("FileExtName")
	Savepath=request.Form("Savepath")
	Set special_rs=Server.CreateObject(G_FS_RS)
	special_rs.open "select * from FS_DS_Special where 1=2",Conn,1,3
	special_rs.addnew
	special_rs("SpecialEName")=NoSqlHack(SpecialEName)
	special_rs("ParentID")=0
	special_rs("SpecialCName")=NoSqlHack(SpecialCName)
	special_rs("SpecialTemplet")=NoSqlHack(SpecialTemplet)
	if IsUrl<>"" then special_rs("IsUrl")=1 else special_rs("IsUrl")=0
	special_rs("Addtime")=Addtime
	special_rs("Domain")=NoSqlHack(Domain)
	if IsLimited<>"" then special_rs("IsLimited")=1 else special_rs("IsLimited")=0
	special_rs("naviPic")=NoSqlHack(naviPic)
	if isLock<>"" then special_rs("isLock")=1 else special_rs("isLock")=0
	special_rs("naviText")=NoSqlHack(naviText)
	special_rs("FileExtName")=NoSqlHack(FileExtName)
	special_rs("Savepath")=NoSqlHack(Savepath)
	special_rs.update
	special_rs.close
elseif action="editaction" then
	specialid=request.QueryString("specialID")
	SpecialCName=request.Form("SpecialCName")
	SpecialTemplet=request.Form("SpecialTemplet")
	IsUrl=trim(request.Form("IsUrl"))
	Addtime=Now
	Domain=request.Form("Domain")
	IsLimited=trim(request.Form("IsLimited"))
	naviPic=request.Form("NaviPic")
	isLock=trim(request.Form("isLock"))
	naviText=request.Form("naviText")
	FileExtName=request.Form("FileExtName")
	Savepath=request.Form("Savepath")
	Set special_rs=Server.CreateObject(G_FS_RS)
	special_rs.open "select SpecialEName,SpecialCName,SpecialTemplet,IsUrl,Addtime,[Domain],IsLimited,naviPic,isLock,naviText,FileExtName,Savepath,ParentID from FS_DS_Special where specialID="&CintStr(specialid),Conn,1,3
	special_rs("SpecialCName")=NoSqlHack(SpecialCName)
	special_rs("ParentID")=0
	special_rs("SpecialTemplet")=NoSqlHack(SpecialTemplet)
	if IsUrl<>"" then special_rs("IsUrl")=1 else special_rs("IsUrl")=0
	special_rs("Addtime")=NoSqlHack(Addtime)
	special_rs("Domain")=NoSqlHack(Domain)
	if IsLimited<>"" then special_rs("IsLimited")=1 else special_rs("IsLimited")=0
	special_rs("naviPic")=NoSqlHack(naviPic)
	if isLock<>"" then special_rs("isLock")=1 else special_rs("isLock")=0
	special_rs("naviText")=NoSqlHack(naviText)
	special_rs("FileExtName")=NoSqlHack(FileExtName)
	special_rs("Savepath")=NoSqlHack(Savepath)
	special_rs.update
	special_rs.close
elseif action="del" then
	Dim specialIDs
	specialIDs=FormatIntArr(request.QueryString("specialid"))
	Conn.execute("delete from FS_DS_Special where specialid in(" & FormatIntArr(specialIDs) & ")")
End if
set special_rs=nothing
set tmp_rs=nothing
Conn.close
Set Conn=nothing

if err.number<>0 then
	Response.Redirect("lib/Error.asp?ErrCodes=<li>发生异常，请在检查数据的有效性后重新操作</li>&ErrorUrl=")
	Response.end
else
	Response.Redirect("lib/Success.asp?ErrCodes=操作成功&ErrorUrl=../../down/Down_Special_Manage.asp")
	Response.end
End if
%>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->







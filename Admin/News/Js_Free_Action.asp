<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%
Dim Conn,FreeJs_Rs,jsid,act,Fs_News
Dim EName,CName,Js_Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode
MF_Default_Conn
MF_Session_TF 
'Call MF_Check_Pop_TF("NS_Class_000001") 
'得到会员组列表 
if not MF_Check_Pop_TF("NS_Freejs") then Err_Show
Set FreeJs_Rs=Server.CreateObject(G_FS_RS)
Set FS_News=New Cls_News
EName = NoSqlHack(Request.Form("txt_ename"))
CName=NoSqlHack(Request.Form("txt_cname"))
Js_Type=NoSqlHack(Request.Form("rad_type"))
act=NoSqlHack(request.QueryString("act"))
If Not Fs_news.IsSelfRefer Then response.write "非法提交数据":Response.end
if Js_Type="" then Js_Type=0
if Cint(Js_Type)=0  then'判断是否是文字类型
	Manner=NoSqlHack(Request.Form("sel_manner"))'文字 
else
	Manner=NoSqlHack(Request.Form("sel_manner_pic"))'图片 
End if
PicWidth=NoSqlHack(Request.Form("txt_picWidth"))
PicHeight=NoSqlHack(Request.Form("txt_picHeight"))
NewsNum=NoSqlHack(Request.Form("txt_newsNum"))
NewsTitleNum=NoSqlHack(Request.Form("txt_newsTitleNum"))
TitleCSS=NoSqlHack(Request.Form("txt_titleCSS"))
ContentCSS=NoSqlHack(Request.Form("txt_contentCSS"))
BackCSS=NoSqlHack(Request.Form("txt_backCSS"))
RowNum=NoSqlHack(Request.Form("txt_rowNum"))
PicPath=NoSqlHack(Request.Form("txt_picPath"))
AddTime=dateValue(Now)
ShowTimeTF=NoSqlHack(Request.Form("sel_showTimeTF"))
ContentNum=NoSqlHack(Request.Form("txt_contentNum"))
if ContentNum="" then
	ContentNum=0
End if
NaviPic=NoSqlHack(Request.Form("txt_naviPic"))
DateType=NoSqlHack(Request.Form("sel_dateType"))
DateCSS=NoSqlHack(Request.Form("sel_dateCSS"))
Info=NoSqlHack(Request.Form("txt_info"))
MoreContent=NoSqlHack(Request.Form("txt_moreContent"))
LinkWord=NoSqlHack(Request.Form("txt_linkWord"))
LinkCSS=NoSqlHack(Request.Form("txt_linkCSS"))
Info=NoSqlHack(Request.Form("txt_info"))
rowSpace=NoSqlHack(Request.Form("txt_rowSpace"))
RowBettween=NoSqlHack(Request.Form("txt_rowBettween"))
OpenMode=NoSqlHack(Request.Form("sel_OpenMode"))
'添加js
if act="add" then
	if not MF_Check_Pop_TF("NS037") then Err_Show
	FreeJs_Rs.open "select * from FS_NS_FreeJS where 1=2",Conn,1,3
	FreeJs_Rs.addnew
	FreeJs_Rs("EName")=EName
	FreeJs_Rs("CName")=CName
	FreeJs_Rs("Type")=Js_Type
	FreeJs_Rs("Manner")=Manner
	if Js_Type="1" then
		FreeJs_Rs("PicWidth")=PicWidth
		FreeJs_Rs("PicHeight")=PicHeight
		FreeJs_Rs("PicPath")=PicPath
	End if
	FreeJs_Rs("NewsNum")=NewsNum
	FreeJs_Rs("AddTime")=AddTime
	FreeJs_Rs("ShowTimeTF")=ShowTimeTF
	if ShowTimeTF="1" then
		FreeJs_Rs("DateType")=DateType
		FreeJs_Rs("DateCSS")=DateCSS
	End if
	FreeJs_Rs("ContentCSS")=ContentCSS
	FreeJs_Rs("TitleCSS")=TitleCSS
	FreeJs_Rs("NewsTitleNum")=NewsTitleNum
	FreeJs_Rs("rowSpace")=rowSpace
	FreeJs_Rs("BackCSS")=BackCSS
	FreeJs_Rs("RowNum")=RowNum
	FreeJs_Rs("ContentNum")=ContentNum
	FreeJs_Rs("NaviPic")=NaviPic
	FreeJs_Rs("Info")=Info
	FreeJs_Rs("MoreContent")=MoreContent
	FreeJs_Rs("LinkWord")=LinkWord
	FreeJs_Rs("LinkCSS")=LinkCSS
	FreeJs_Rs("Info")=Info
	FreeJs_Rs("RowBettween")=RowBettween
	FreeJs_Rs("OpenMode")=OpenMode
	FreeJs_Rs.update
	FreeJs_Rs.close
'修改js
Elseif act="edit" Then
	if not MF_Check_Pop_TF("NS038") then Err_Show
	jsid=Request.Form("hid_jsid")
	if not isnumeric(jsid) Then
		Response.Redirect("lib/error.asp?ErrCodes=<li>发生异常，请返回</li>")
		Response.End()
	End if
	FreeJs_Rs.open "select ID,EName,CName,Type,Manner,PicWidth,PicHeight,NewsNum,NewsTitleNum,TitleCSS,ContentCSS,BackCSS,RowNum,PicPath,AddTime,ShowTimeTF,ContentNum,NaviPic,DateType,DateCSS,Info,MoreContent,LinkWord,LinkCSS,RowSpace,RowBettween,OpenMode from FS_NS_FreeJS where id="&jsid,Conn,1,3
	FreeJs_Rs("EName")=EName
	FreeJs_Rs("CName")=CName
	FreeJs_Rs("Type")=Js_Type
	FreeJs_Rs("Manner")=Manner
	if Js_Type="1" then
		FreeJs_Rs("PicWidth")=PicWidth
		FreeJs_Rs("PicHeight")=PicHeight
		FreeJs_Rs("PicPath")=PicPath
	End if
	FreeJs_Rs("NewsNum")=NewsNum
	FreeJs_Rs("AddTime")=AddTime
	FreeJs_Rs("ShowTimeTF")=ShowTimeTF
	if ShowTimeTF="1" then
		FreeJs_Rs("DateType")=DateType
		FreeJs_Rs("DateCSS")=DateCSS
	End if
	FreeJs_Rs("ContentCSS")=ContentCSS
	FreeJs_Rs("TitleCSS")=TitleCSS
	FreeJs_Rs("NewsTitleNum")=NewsTitleNum
	FreeJs_Rs("rowSpace")=rowSpace
	FreeJs_Rs("BackCSS")=BackCSS
	FreeJs_Rs("RowNum")=RowNum
	FreeJs_Rs("ContentNum")=ContentNum
	FreeJs_Rs("NaviPic")=NaviPic
	FreeJs_Rs("Info")=Info
	FreeJs_Rs("MoreContent")=MoreContent
	FreeJs_Rs("LinkWord")=LinkWord
	FreeJs_Rs("LinkCSS")=LinkCSS
	FreeJs_Rs("Info")=Info
	FreeJs_Rs("RowBettween")=RowBettween
	FreeJs_Rs("OpenMode")=OpenMode
	FreeJs_Rs.update
	FreeJs_Rs.close
'删除js
ElseIf act="delete" Then
	if not MF_Check_Pop_TF("NS039") then Err_Show
	Dim d_js_ids,MyFile,TempRs,TempJsPath,TempID,TempEName,TempNameRs,TempArr
	
	Set TempRs=Conn.Execute("Select NewsDir from FS_NS_SysParam")
	If TempRs.eof Then
		Response.Redirect("lib/error.asp?ErrCodes=<li>出现异常</li>")
		Response.End()
	End if
	TempJsPath=TempRs("NewsDir")
	TempRs.Close
	Set TempRs=Nothing
	
	d_js_ids=Request.Form("chk_FreeJs")
	d_js_ids=DelHeadAndEndDot(d_js_ids)
	TempArr=Split(d_js_ids,",")
	For Each TempID In(TempArr)
		Set TempNameRs=Conn.execute("select EName From FS_NS_FreeJS where id="&CintStr(TempID))
		If Not TempNameRs.Eof Then 
			TempEName=TempNameRs("EName")
			Set MyFile=Server.CreateObject(G_FS_FSO)
			Dim Str_sysRoot
			If G_VIRTUAL_ROOT_DIR="" Then
				Str_sysRoot="/"
			Else
				Str_sysRoot="/"&G_VIRTUAL_ROOT_DIR&"/"
			End If			
			If MyFile.FileExists(Server.MapPath(replace(replace(Str_sysRoot&TempJsPath&"/JS/FreeJs","///","/"),"//","/"))&"\"& TempEName &".js") then
				MyFile.DeleteFile(Server.MapPath(replace(replace(Str_sysRoot&TempJsPath&"/JS/FreeJs","///","/"),"//","/"))&"\"& TempEName &".js")
			End If
			Set MyFile=nothing	
		End If
		TempNameRs.CLose
		Set TempNameRs=Nothing
	Next
	Conn.execute("Delete From FS_NS_FreeJsFile where JSName in (select EName From FS_NS_FreeJS where id in ("&FormatIntArr(d_js_ids)&"))")
	Conn.execute("Delete From FS_NS_FreeJS where id in ("&FormatIntArr(d_js_ids)&")")

	Call MF_Insert_oper_Log("自由JS","批量删除了自由JS,删除ID："& Replace(d_js_ids," ","") &"",now,session("admin_name"),"NS")
End if
if err.number=0 then
	Response.Redirect("lib/success.asp?ErrCodes=<li>操作成功</li>&ErrorURL=../Js_Free_Manage.asp")
	Response.End()
else
	Response.Redirect("lib/error.asp?ErrCodes=<li>发生异常，请返回</li>")
	Response.End()
end if
%>
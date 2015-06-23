<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%
Dim Conn,ns_sysPara_Rs,ns_siteName,ns_keyWords,ns_newsDir,ns_isDomain,ns_fileNameRule,ns_fileDirRule,ns_classSaveType,ns_fileExtName,ns_indexPage,ns_newsCheck,ns_refreshFile,ns_isOpen,ns_indexTemplet,ns_isPrintPic,ns_linkType,ns_fileChar,ns_isCheck,ns_isReviewCheck,ns_isConstrCheck,ns_addNewsType,ns_allInfotitle,ns_reycleTF,Fs_News
Dim CopyFileTF,EditFilesTF
Dim ns_InsideLink,ns_RSSTF,ns_rssNumber,ns_rssdescript,ns_RSSPIC,ns_rssContentNumber,SysParmTF
MF_Default_Conn'初始化Conn
MF_Session_TF
if not MF_Check_Pop_TF("NS_Param") then Err_Show
if not MF_Check_Pop_TF("NS049") then Err_Show
Set Fs_News=New Cls_News
If Not Fs_news.IsSelfRefer Then response.write "非法提交数据":Response.end
if NoSqlHack(Request.QueryString("act"))="SetSysPara_Action" then
	ns_SiteName=Request.Form("txt_SiteName")
	ns_keyWords=Request.Form("txt_KeyWords")
	ns_newsDir=Request.Form("txt_NewsDir")
	ns_isDomain=Request.Form("rad_IsDomain")
	if trim(Request.Form("txt_FileNameRule_Element_Separator"))<>"" then
		if not Fs_News.chkinputchar(trim(Request.Form("txt_FileNameRule_Element_Separator"))) then
			Response.Redirect("lib/error.asp?ErrCodes=<li>分割符号只允许为：""0-9""，""A-Z""，""-"",""_"","",""."",""@"",""#""</li>")
			Response.End()
		end if
	End if
	ns_fileNameRule=trim(Request.Form("txt_FileNameRule_Element_Prefix"))&"$"&trim(replace(Request.Form("chk_FileNameRule_Element"),",",""))&"$"&trim(Request.Form("rad_FileNameRule_Rnd"))&"$"&trim(Request.Form("chk_FileNameRule_UseWord"))&"$"&trim(Request.Form("txt_FileNameRule_Element_Separator"))&"$"&trim(Request.Form("rad_FileNameRule_UseNewsID"))&"$"&trim(Request.Form("rad_FileNameRule_NewsID"))
	ns_fileDirRule=Request.Form("rad_FileDirRule")
	ns_classSaveType=Request.Form("rad_ClassSaveType")
	ns_fileExtName=Request.Form("rad_FileExtName")
	ns_indexPage=trim(Request.Form("txt_IndexPage_Name"))
	ns_newsCheck=Request.Form("rad_NewsCheck")
	ns_isOpen=Request.Form("rad_isOpen")
	ns_indexTemplet=Request.Form("txt_IndexTemplet")
	ns_linkType=Request.Form("rad_LinkType")
	ns_isCheck=Request.Form("rad_isCheck")
	ns_isReviewCheck=Request.Form("rad_isReviewCheck")
	ns_isConstrCheck=Request.Form("rad_isConstrCheck")
	CopyFileTF = Request.Form("ISCopyFilesTF")
	EditFilesTF = Request.Form("EditFileTF")
	ns_addNewsType=Request.Form("rad_AddNewsType")
	ns_allInfotitle=Request.Form("txt_AllInfotitle")
	ns_reycleTF=Request.Form("rad_ReycleTF")
	ns_RSSTF= Request.Form("RSSTF")
	ns_rssNumber= Request.Form("rssNumber")
	ns_rssdescript= Request.Form("rssdescript")
	ns_RSSPIC= Request.Form("RSSPIC")
	ns_rssContentNumber=Request.Form("rssContentNumber")
	ns_InsideLink=Request.Form("InsideLink")
	SysParmTF = True
	if ns_reycleTF="" Then ns_reycleTF=1 
	Set ns_sysPara_Rs=Server.CreateObject(G_FS_RS)
	ns_sysPara_Rs.open "Select SiteName,Keywords,NewsDir,IsDomain,FileNameRule,FileDirRule,ClassSaveType,FileExtName,IndexPage,NewsCheck,isOpen,IndexTemplet,LinkType,isCheck,isReviewCheck,isConstrCheck,IsCopyFileTF,IsEditFileTF,AddNewsType,AllInfotitle,ReycleTF,RSSTF,rssNumber,rssdescript,RSSPIC,rssContentNumber,InsideLink From FS_ns_SysParam",Conn,1,3
	If ns_sysPara_Rs.Eof Then
		ns_sysPara_Rs.AddNew
	End If
	ns_sysPara_Rs("SiteName")=NoSqlHack(ns_SiteName)
	ns_sysPara_Rs("Keywords")=NoSqlHack(ns_keyWords)
	ns_sysPara_Rs("NewsDir")=NoSqlHack(ns_newsDir)
	ns_sysPara_Rs("IsDomain")=NoSqlHack(ns_isDomain)
	ns_sysPara_Rs("FileNameRule")= NoSqlHack(ns_fileNameRule)
	ns_sysPara_Rs("FileDirRule")=CintStr(ns_fileDirRule)
	ns_sysPara_Rs("ClassSaveType")=CintStr(ns_classSaveType)
	ns_sysPara_Rs("FileExtName")=CintStr(ns_fileExtName)
	ns_sysPara_Rs("IndexPage")=ns_indexPage
	'ns_sysPara_Rs("NewsCheck")=ns_newsCheck
	ns_sysPara_Rs("isOpen")=CintStr(ns_isOpen)
	ns_sysPara_Rs("IndexTemplet")=NoSqlHack(ns_indexTemplet)
	ns_sysPara_Rs("LinkType")=CintStr(ns_linkType)
	ns_sysPara_Rs("isCheck")=CintStr(ns_isCheck)
	ns_sysPara_Rs("isReviewCheck")=CintStr(ns_isReviewCheck)
	ns_sysPara_Rs("isConstrCheck")=CintStr(ns_isConstrCheck)
	ns_sysPara_Rs("IsCopyFileTF")=CintStr(CopyFileTF)
	ns_sysPara_Rs("IsEditFileTF")=CintStr(EditFilesTF)
	ns_sysPara_Rs("AddNewsType")=CintStr(ns_addNewsType)
	ns_sysPara_Rs("AllInfotitle")=NoSqlHack(ns_allInfotitle)
	ns_sysPara_Rs("ReycleTF")=CintStr(ns_reycleTF)
	if trim(ns_InsideLink)="1" then:ns_sysPara_Rs("InsideLink")=1:else:ns_sysPara_Rs("InsideLink")=0:end if
	'RSS
	if trim(ns_RSSTF)<>"" then:ns_sysPara_Rs("RSSTF")=1:else:ns_sysPara_Rs("RSSTF")=0:end if
	if isnumeric(ns_rssNumber) then
		ns_sysPara_Rs("rssNumber")=CintStr(ns_rssNumber)
	else
		ns_sysPara_Rs("rssNumber")=50
	end if
	ns_sysPara_Rs("rssdescript")=ns_rssdescript
	ns_sysPara_Rs("RSSPIC")=ns_RSSPIC
	if isnumeric(ns_rssContentNumber) then
		ns_sysPara_Rs("rssContentNumber")=CintStr(ns_rssContentNumber)
	else
		ns_sysPara_Rs("rssContentNumber")=400
	end if
	ns_sysPara_Rs.update
	ns_sysPara_Rs.close
	Set ns_sysPara_Rs=nothing
	if err.number=0 then
		NSConfig_Cookies
		Conn.close
		Set Conn=nothing
		Response.Redirect("lib/success.asp?ErrCodes=<li>操作成功</li>&ErrorURL=../SysParaSet.asp")
		Response.End()
	else
		Conn.close
		Set Conn=nothing
		Response.Redirect("lib/error.asp?ErrCodes=<li>请检查输入是否合法</li>")
		Response.End()
	end if
end if
%>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






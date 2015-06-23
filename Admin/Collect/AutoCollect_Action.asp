<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/WaterPrint_Function.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="CS_Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../News/lib/cls_main.asp" -->
<%
Dim Conn,CollectConn
MF_Default_Conn
MF_Collect_Conn
MF_Session_TF
Response.Buffer = true
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
if not MF_Check_Pop_TF("CS_collect") then Err_Show
Dim p_SYS_ROOT_DIR,SiteID,ErrorInfoStr,Action,SaveIMGPath,ListHeadSetting,ListFootSetting,LinkHeadSetting,LinkFootSetting
Dim PagebodyHeadSetting,PagebodyFootSetting,PageTitleHeadSetting,PageTitleFootSetting,OtherPageFootSetting,OtherPageHeadSetting
Dim OtherNewsPageHeadSetting,OtherNewsPageFootSetting,AuthorHeadSetting,AuthorFootSetting,SourceHeadSetting,SourceFootSetting
Dim AddDateHeadSetting,AddDateFootSetting,IndexRule,StartPageNum,EndPageNum,HandPageContent,OtherType
Dim IsStyle,IsDiv,IsA,IsClass,IsFont,IsSpan,IsObjectTF,IsIFrame,IsScript,HandSetAuthor,HandSetSource,HandSetAddDate,TextTF,SaveRemotePic,IsReverse
Dim ObjURL,ReturnValue,CollectStartLocation,CollectEndFlag,CollectObjURL,CollectedPageURL,p_DoMain_Str
Dim SiteName,CollectingSiteID,CollectSiteIndex,AllNewsNumber,CollectOKNumber,CollectPageNumber,Num,CollectType
Dim OtherNewsType,OtherNewsPageIndexSetting,OtherNewsPageIndexSettingStartPageNum,OtherNewsPageIndexSettingEndPageNum,OtherNewsPageIndexSettingHandPageContent
Dim WebCharset,WaterPrintTF,CS_SiteReKeyID,Temp_picPath,AuditTF
Dim AutoCollect
AutoCollect=True
If G_VIRTUAL_ROOT_DIR = "" Then
	p_SYS_ROOT_DIR = ""
Else
	p_SYS_ROOT_DIR = "/" & G_VIRTUAL_ROOT_DIR
End If
p_DoMain_Str = "http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
SiteID = Request("SiteID")
ErrorInfoStr = ""
CollectEndFlag = False
CollectedPageURL = Request("CollectedPageURL")
AllNewsNumber = Request("AllNewsNumber")
If AllNewsNumber = "" Then
	AllNewsNumber = 0
Else
	AllNewsNumber = CLng(AllNewsNumber)
End If
CollectOKNumber = Request("CollectOKNumber")
If CollectOKNumber = "" Then
	CollectOKNumber = 0
Else
	CollectOKNumber = CLng(CollectOKNumber)
End If
CollectSiteIndex = Request("CollectSiteIndex")
If CollectSiteIndex = "" Then
	CollectSiteIndex = 0
Else
	CollectSiteIndex = CInt(CollectSiteIndex)
End If
CollectPageNumber = Request("CollectPageNumber")
If CollectPageNumber = "" Then
	CollectPageNumber = 0
Else
	CollectPageNumber = CInt(CollectPageNumber)
End If
CollectStartLocation = Request("CollectStartLocation")
If CollectStartLocation = "" Then CollectStartLocation = 0
Num = Request("Num")
If Num = "allNews" Or Num="" Then
	Num = 10
Else
	If Not IsNumeric(Num) Then
		Num = 10
	Else
		Num = CInt(Num)
	End If
End If
'On Error Resume Next
CollectType = Request("CollectType")
If SiteID <> "" Then
	GetCollectPara
	If AllNewsNumber>=Num And Num<>0 Then
		CollectEndFlag = True
	End If
	If CollectEndFlag Then
		If ErrorInfoStr <> "" Then
			Response.Write(ErrorInfoStr)
		Else
			ReturnValue = "End||"&SiteID&"||<br><strong>采集结束</strong>：共读取" & AllNewsNumber & "条新闻，采集成功" & CollectOKNumber & "条新闻。"
			Response.Write(ReturnValue)
		End If
	ElseIf CollectType<>"ResumeCollect" Then
		GetNewsPageContent()
		Response.Write("Next||"&SiteID&"||CollectPageNumber=" & CollectPageNumber & "&SiteID=" & SiteID & "&CollectStartLocation=" & CollectStartLocation & "&CollectedPageURL=" & CollectedPageURL & "&CollectSiteIndex=" & CollectSiteIndex & "&Num=" & Num & "&AllNewsNumber=" & AllNewsNumber & "&CollectOKNumber=" & CollectOKNumber)
	Else
		ResumeGetNewsPageContent()
		Response.Write("Next||"&SiteID&"||CollectType=ResumeCollect&CollectPageNumber=" & CollectPageNumber & "&SiteID=" & SiteID & "&CollectStartLocation=" & CollectStartLocation & "&CollectedPageURL=" & CollectedPageURL & "&CollectSiteIndex=" & CollectSiteIndex & "&AllNewsNumber=" & AllNewsNumber & "&CollectOKNumber=" & CollectOKNumber)
	End If
	If Err Then
		Response.Write("Err||"&SiteID)
	End If
End If
%>






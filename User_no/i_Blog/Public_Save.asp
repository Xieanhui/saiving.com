<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="../../FS_Inc/Md5.asp" -->
<%
	If CheckBlogOpen=False Then 
		Response.write("<script language=""javascript"">alert('日志功能暂停使用,如需要使用请联系管理员.');history.back();</script>")
		Response.End()
	End If
	
	dim p_iLogStyle,p_Title,p_KeyWords,p_Content,p_iLogSource,p_MainID,str_Action,str_iLogID,rs,p_PutInPhoto,rs_sys,isCheck,Kcontent
	dim p_ClassID,p_EmotFace,p_isTop,p_FileName,p_FileExtName,p_Pic_1,p_Pic_2,p_Pic_3,p_Password,tmp_pic,tmp_picarr,tmp_i
	set rs_sys=User_Conn.execute("select top 1 isCheck,Kcontent From FS_ME_iLogSysParam")
	if rs_sys.eof then
		response.Write("系统配置出错!")
		response.End
		rs_sys.close:set rs_sys=nothing
	else
		isCheck=rs_sys("isCheck")
		Kcontent=rs_sys("Kcontent")
		rs_sys.close:set rs_sys=nothing
	end if
	str_Action=NoSqlHack(Request.Form("Action"))
	p_PutInPhoto=NoSqlHack(Request.Form("PutInPhoto"))
	tmp_pic=NoSqlHack(Request.Form("pic_1"))&"|"&NoSqlHack(Request.Form("pic_2"))&"|"&NoSqlHack(Request.Form("pic_3"))
	p_iLogStyle = NoSqlHack(Request.Form("iLogStyle"))
	p_Title = NoSqlHack(Request.Form("Title"))
	p_KeyWords = NoSqlHack(Request.Form("keyword1"))&","&NoSqlHack(Request.Form("keyword2"))&","&NoSqlHack(Request.Form("keyword3"))
	p_Content = NoSqlHack(Request.Form("Content"))
	p_iLogSource = NoSqlHack(Request.Form("iLogSource"))
	p_MainID = NoSqlHack(Request.Form("MainID"))
	p_ClassID = NoSqlHack(Request.Form("ClassID"))
	p_EmotFace = NoSqlHack(Request.Form("EmotFace"))
	p_isTop = NoSqlHack(Request.Form("isTop"))
	p_FileName = NoSqlHack(Request.Form("FileName"))
	p_FileExtName = NoSqlHack(Request.Form("FileExtName"))
	p_Pic_1 = NoSqlHack(Request.Form("Pic_1"))
	p_Pic_2 = NoSqlHack(Request.Form("Pic_2e"))
	p_Pic_3 = NoSqlHack(Request.Form("Pic_3"))
	p_Password = md5(NoSqlHack(Request.Form("Password")),16)
	str_iLogID = NoSqlHack(Request.Form("Id"))
	if p_Title="" or isnull(p_Title) or p_Content="" or isnull(p_Content) or p_FileName="" or isnull(p_FileName) or p_MainID="" or isnull(p_MainID) then
		strShowErr="<li>带*必须填写</li>"
		Response.Redirect("../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	if str_Action<>"Edit" then
		if p_FileName<>"自动编号" then
			dim rstf
			set rstf = User_Conn.execute("select FileName,FileExtName From FS_ME_Infoilog where UserNumber='"&Fs_User.UserNumber&"' and FileName='"&p_FileName&"' and FileExtName='"&P_FileExtName&"'")
			if not rstf.eof then
				strShowErr="<li>文件名已经存在</li>"
				Response.Redirect("../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			end if
			rstf.close:set rstf=nothing
		end if
	end if
	set rs= Server.CreateObject(G_FS_RS)
	if str_Action="Save" then
		rs.open "select * From FS_ME_Infoilog where 1=0",User_Conn,1,3
		rs.addnew
		rs("isDraft")=0
		rs("Addtime")=now
		if isCheck=1 then
			rs("adminLock")=1
		else
			rs("adminLock")=0
		end if
		rs("isTF")=0
		rs("Hits")=0
		rs("isLock")=0
		rs("savePath")=year(date)&"-"&month(date)
	elseif str_Action="isDraft" then
		rs.open "select * From FS_ME_Infoilog where 1=0",User_Conn,1,3
		rs.addnew
		rs("isDraft")=1
		rs("Addtime")=now
		rs("isLock")=0
		rs("savePath")=year(date)&"-"&month(date)
		rs("isTF")=0
		rs("Hits")=0
		if isCheck=1 then
			rs("adminLock")=1
		else
			rs("adminLock")=0
		end if
	elseif str_Action="Edit" then
		rs.open "select * From FS_ME_Infoilog where UserNumber='"&Fs_User.UserNumber&"' and iLogID="&CintStr(str_iLogID),User_Conn,1,3
		if request.Form("isDraft")<>"" then
			rs("isDraft")=1
		else
			rs("isDraft")=0
		end if
	end if
	if p_iLogStyle="1" then:rs("iLogStyle")=1:else:rs("iLogStyle")=0:end if
	rs("Title")=p_Title
	rs("KeyWords")=p_KeyWords
	rs("Content")=p_Content
	rs("iLogSource")=p_iLogSource
	if p_MainID<>"" then:rs("MainID")=CintStr(p_MainID):else:rs("MainID")=0:end if
	if p_ClassID<>"" then:rs("ClassID")=CintStr(p_ClassID):else:rs("ClassID")=0:end if
	rs("UserNumber")=Fs_User.UserNumber
	rs("EmotFace")=p_EmotFace
	if p_isTop<>"" then:rs("isTop")=1:else:rs("isTop")=0:end if
	rs("TempletID")=0
	rs("FileName")=p_FileName
	rs("FileExtName")=p_FileExtName
	if trim(p_Pic_1)<>"" then:rs("Pic_1")=p_Pic_1:end if
	if trim(p_Pic_2)<>"" then:rs("Pic_2")=p_Pic_2:end if
	if trim(p_Pic_3)<>"" then:rs("Pic_3")=p_Pic_3:end if
	if trim(Request.Form("Password"))<>"" then:rs("Password")=p_Password:end if
	rs.update
	Dim Get_News_ID '取自动编号ID，sqlserver中，有待休正
	if G_IS_SQL_DB = 0 then:Get_News_ID = rs("iLogID"):Else:Get_News_ID = "":End if
	rs.close:set rs=nothing
	If Instr(p_FileName,"自动编号") Then
		Dim TempRsObj
		p_FileName = Replace(p_FileName,"自动编号",Get_News_ID)
		Set TempRsObj=server.CreateObject(G_FS_RS)
		TempRsObj.open "select FileName From [FS_ME_Infoilog] where iLogID="&Get_News_ID&" and UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
		if not TempRsObj.eof Then
			TempRsObj("FileName") = Replace(TempRsObj("FileName"),"自动编号",Get_News_ID)
			TempRsObj.update
		End If
		TempRsObj.Close
	End IF
	'图片插入相册
	if p_PutInPhoto<>"" then
		tmp_picarr = split(tmp_pic,"|")
		for tmp_i = 0 to UBound(tmp_picarr)
			set rs= Server.CreateObject(G_FS_RS)
			rs.open "select * From FS_ME_Photo where PicSavePath='"&NoSqlHack(tmp_picarr(tmp_i))&"' and UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
			if rs.eof then
				if trim(tmp_picarr(tmp_i))<>"" then
					rs.addnew
					rs("title")=p_title
					rs("PicSavePath")=tmp_picarr(tmp_i)
					rs("Content")="无"
					rs("Addtime")=now
					rs("ClassID")=0
					rs("UserNumber")=Fs_User.UserNumber
					rs("PicSize")=0
					rs.update
				end if
			end if
		next
		rs.close:set rs=nothing
	end if
	set Fs_User=nothing
	if str_Action="Edit" then
		if request.Form("isDraft")<>"" then
			strShowErr = "<li>修改并保存到草稿箱中成功！</li><li><a href=../i_Blog/PublicLogEdit.asp?Id="&str_iLogID&">继续修改</a>&nbsp;&nbsp;<a href=../i_Blog/index.asp>返回日志管理</a></li>"
		else
			strShowErr = "<li>修改成功！</li><li><a href=../i_Blog/PublicLogEdit.asp?Id="&str_iLogID&">继续修改</a>&nbsp;&nbsp;<a href=../i_Blog/index.asp>返回日志管理</a></li>"
		end if
	else
		strShowErr = "<li>保存成功！</li><li><a href=../i_Blog/PublicLog.asp>继续添加</a>&nbsp;&nbsp;<a href=../i_Blog/index.asp>返回日志管理</a></li>"
	end if
	Response.Redirect("../lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->
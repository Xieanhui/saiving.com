<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
	Dim Conn,User_Conn
	Dim CharIndexStr
	Dim Fs_news,obj_news_rs,obj_news_rs_1,isUrlStr,str_Href,obj_cnews_rs,news_count,str_Href_title,str_action,str_ClassID,news_SQL
	Dim obj_newslist_rs,newslist_sql,strpage,str_showTF,str_ClassID_1,str_Editor,str_Keyword,str_GetKeyword,str_ktype
	Dim select_count,select_pagecount,i,Str_GetPopID,Str_PopID,str_check,str_UrlTitle,icNum,str_addType,str_addType_1
	Dim str_Rec,str_isTop,str_hot,str_pic,str_highlight,str_bignews,str_filt,str_Constr,str_Top,tmp_pictf
	Dim str_s_classIDarray,tmp_splitarrey_id,tmp_splitarrey_Classid,tmp_i,str_Move_type,str_t_classID,C_NewsIDarrey,Tmp_rs,Tmp_TF_Rs
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
	Dim StrSql,ArrSql(),temp_j,str_rep_type,str_rep_select_type,str_AdvanceTF,str_s_Content,str_t_Content,str_start_char,str_end_char,f_PLACE_OBJ
	Dim str_SpecialEname,SQL_SpecialEname,str_Get_Special,sp_i
	Dim fso_tmprs_,NewsSavePath,Str_Temp_Flag
	int_RPP=30 '����ÿҳ��ʾ��Ŀ
	int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
	showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
	str_nonLinkColor_="#999999" '����������ɫ
	toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ
	toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
	toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
	toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
	toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
	toL_="<font face=webdings title=""���һҳ"">:</font>"

	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF
	set Fs_news = new Cls_News
	Fs_News.GetSysParam()
	If Not Fs_news.IsSelfRefer Then response.write "�Ƿ��ύ����":Response.end
	str_ClassID = NoSqlHack(Request.QueryString("ClassID"))
	str_SpecialEname = NoSqlHack(Request.QueryString("specialEname"))
	if Request("Action")="makehtml" then
		if not Get_SubPop_TF(str_ClassID,"NS011","NS","news") then Err_Show
		'�޸ı���������ɡ�08.10.16.Fsj
		if ""=server.URLEncode(Replace(NoSqlHack(request.Form("C_NewsID"))," ","")) and ""=NoSqlHack(Request.QueryString("ClassId")) then
		    response.write("<script>alert('err���������Ų������ɣ�');</script>")
		    response.write("<script> location.replace(location);</script>")
		else
		    response.Redirect "Get_NewsHtml.asp?Id="&server.URLEncode(Replace(NoSqlHack(request.Form("C_NewsID"))," ",""))&"&ClassId="&NoSqlHack(Request.QueryString("ClassId"))&"&type=makenews"
	    end  if
	End if
	if Request("Action")="Toold" then
		response.Redirect "Get_OldNews.asp?Id="&server.URLEncode(Replace(NoSqlHack(request.Form("C_NewsID"))," ",""))&"&ClassId="&NoSqlHack(Request.QueryString("ClassId"))&""
	end if
	if Request("Action") = "signDel" then
		if not Get_SubPop_TF(str_ClassID,"NS003","NS","news") then Err_Show
		Dim strShowErr
		if fs_news.ReycleTF = 1 then
			Conn.execute("Update FS_NS_News set isRecyle = 1 where NewsID='"& NoSqlHack(Request.QueryString("NewsID"))&"'")
			strShowErr = "<li>"& Fs_news.allInfotitle &"�Ѿ�ɾ��</li><li>"& Fs_news.allInfotitle &"�Ѿ���ʱ�ŵ�����վ��</li>"
		Else
			strShowErr = "<li>"& Fs_news.allInfotitle &"�Ѿ�����ɾ��</li>"
			''===============
			''ɾ��ͼƬ�ļ�
			set fso_tmprs_ = Conn.execute("select NewsPicFile from FS_NS_News where NewsID='"& NoSqlHack(Request.QueryString("NewsID"))&"'")
			if not fso_tmprs_.eof then
				fso_DeleteFile(fso_tmprs_(0))
			end if
			fso_tmprs_.close
			''Сͼ
			set fso_tmprs_ = Conn.execute("select NewsSmallPicFile from FS_NS_News where NewsID='"& NoSqlHack(Request.QueryString("NewsID"))&"'")
			if not fso_tmprs_.eof then
				fso_DeleteFile(fso_tmprs_(0))
			end If
			'ɾ����̬�ļ�
			set fso_tmprs_ = Conn.execute("select FS_NS_News.SaveNewsPath,FS_NS_News.FileName,FS_NS_News.FileExtName,FS_NS_NewsClass.SavePath,FS_NS_NewsClass.ClassEName from FS_NS_News,FS_NS_NewsClass where NewsID='"& NoSqlHack(Request.QueryString("NewsID"))&"' and FS_NS_News.IsURL=0 and FS_NS_NewsClass.ClassID=FS_NS_News.ClassID")
			If G_VIRTUAL_ROOT_DIR = "" Then
				NewsSavePath = ""
			Else
				NewsSavePath = "/" & G_VIRTUAL_ROOT_DIR
			End If
			If Not fso_tmprs_.eof Then
				NewsSavePath=NewsSavePath&fso_tmprs_("SavePath")&"/"&fso_tmprs_("ClassEName")&fso_tmprs_("SaveNewsPath")&"/"&fso_tmprs_("FileName")&"."&fso_tmprs_("FileExtName")
				fso_DeleteFile(NewsSavePath)
			End If
			'ɾ����̬�ļ�����
			fso_tmprs_.close
			''===============
			Conn.execute("Delete From FS_NS_News where NewsID='"& NoSqlHack(Request.QueryString("NewsID"))&"'")
			'ɾ�����Ȩ�����ţ��Է�����������Ϣ
			Conn.execute("Delete From  FS_MF_Pop  where InfoID='"& NoSqlHack(Request.QueryString("NewsID"))&"' and PopType='NS'")
			'ɾ����̬�ļ�
		End if
		
		set conn=nothing:set user_conn=nothing
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="&Request.ServerVariables("HTTP_REFERER")&"")
		Response.end
	End if
	if Request("Action") = "singleCheck" then
		if not Get_SubPop_TF(str_ClassID,"NS004","NS","news") then Err_Show
		Conn.execute("Update FS_NS_News set isLock = 0 where NewsID='"& NoSqlHack(Request.QueryString("NewsID"))&"'")
		strShowErr = "<li>"& Fs_news.allInfotitle &"��˳ɹ�</li>"
		set conn=nothing:set user_conn=nothing
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../News_Manage.asp")
		Response.end
	Elseif Request("Action") = "singleUnCheck" then
		if not Get_SubPop_TF(str_ClassID,"NS005","NS","news") then Err_Show
		Conn.execute("Update FS_NS_News set isLock = 1 where NewsID='"& NoSqlHack(Request.QueryString("NewsID"))&"'")
		strShowErr = "<li>"& Fs_news.allInfotitle &"�����ɹ�</li>"
		set conn=nothing:set user_conn=nothing
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../News_Manage.asp")
		Response.end
	End if
	if Request("Action") = "signUnTop" then
		if not Get_SubPop_TF(str_ClassID,"NS006","NS","news") then Err_Show
		Conn.execute("Update FS_NS_News set popid =0 where NewsID='"& NoSqlHack(Request.QueryString("NewsID"))&"'")
		strShowErr = "<li>"& Fs_news.allInfotitle &"��̳ɹ�</li>"
		set conn=nothing:set user_conn=nothing
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../News_Manage.asp")
		Response.end
	End if
	if Request("Action") = "signTop" then
		if not Get_SubPop_TF(str_ClassID,"NS006","NS","news") then Err_Show
		Conn.execute("Update FS_NS_News set popid =5 where NewsID='"& NoSqlHack(Request.QueryString("NewsID"))&"'")
		set conn=nothing:set user_conn=nothing
		strShowErr = "<li>"& Fs_news.allInfotitle &"�̶ܹ��ɹ�</li><li>�����Ҫ����̶�,�����޸����趨</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../News_Manage.asp")
		Response.end
	End if
	If Request.Form("Action") = "Del" then
		if not Get_SubPop_TF(str_ClassID,"NS003","NS","news") then Err_Show
		Dim DelID,Str_Tmp,Str_Tmp1
		DelID = FormatIntArr(request.Form("C_NewsID"))
		if DelID = "" then
			strShowErr = "<li>�����ѡ��һ����ɾ��</li>"
			set conn=nothing:set user_conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		if fs_news.ReycleTF = 1 then
			Conn.execute("Update FS_NS_News set isRecyle=1 where ID in ("&DelID&")")
			set conn=nothing:set user_conn=nothing
			strShowErr = "<li>"& Fs_news.allInfotitle &"�Ѿ�ɾ��</li><li>"& Fs_news.allInfotitle &"�Ѿ���ʱ�ŵ�����վ��</li>"
		Else

			''===============
			''ɾ��ͼƬ�ļ�
			set fso_tmprs_ = Conn.execute("select NewsPicFile from FS_NS_News where ID in ("&DelID&")")
			do while not fso_tmprs_.eof
				fso_DeleteFile(fso_tmprs_(0))
				fso_tmprs_.movenext
			loop
			fso_tmprs_.close
			''Сͼ
			set fso_tmprs_ = Conn.execute("select NewsSmallPicFile from FS_NS_News where ID in ("&DelID&")")
			do while not fso_tmprs_.eof
				fso_DeleteFile(fso_tmprs_(0))
				fso_tmprs_.movenext
			loop
			fso_tmprs_.close
			''===============
			'ɾ����̬�ļ�
			set fso_tmprs_ = Conn.execute("select FS_NS_News.SaveNewsPath,FS_NS_News.FileName,FS_NS_News.FileExtName,FS_NS_NewsClass.SavePath,FS_NS_NewsClass.ClassEName from FS_NS_News,FS_NS_NewsClass where FS_NS_News.ID in ("& DelID &") and FS_NS_News.IsURL=0 and FS_NS_NewsClass.ClassID=FS_NS_News.ClassID")
			While Not fso_tmprs_.eof
				If G_VIRTUAL_ROOT_DIR = "" Then
					NewsSavePath = ""
				Else
					NewsSavePath = "/" & G_VIRTUAL_ROOT_DIR
				End If
				NewsSavePath=NewsSavePath&fso_tmprs_("SavePath")&"/"&fso_tmprs_("ClassEName")&fso_tmprs_("SaveNewsPath")&"/"&fso_tmprs_("FileName")&"."&fso_tmprs_("FileExtName")
				fso_DeleteFile(NewsSavePath)
				fso_tmprs_.movenext
			Wend
			'ɾ����̬�ļ�����
			Conn.execute("Delete From FS_NS_News where ID in ("&DelID&")")
			set conn=nothing:set user_conn=nothing
			strShowErr = "<li>��ѡ���"& Fs_news.allInfotitle &"�Ѿ�����ɾ����</li>"
		End if
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../News_manage.asp?ClassID="&Request("ClassID")&"")
		Response.end
	End if
	If Request.Form("Action") = "unlock" then
		if not Get_SubPop_TF(str_ClassID,"NS004","NS","news") then Err_Show
		Dim str_UnLockID
		str_UnLockID = FormatIntArr(request.Form("C_NewsID"))
		if str_UnLockID = "" then
			set conn=nothing:set user_conn=nothing
			strShowErr = "<li>�����ѡ��һ���ٲ���</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		Conn.execute("Update FS_NS_News set isLock = 0 where ID in ("& str_UnLockID &")")
		strShowErr = "<li>"& Fs_news.allInfotitle &"��˳ɹ�</li>"
		set conn=nothing:set user_conn=nothing
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	If Request.Form("Action") = "lock" then
		if not Get_SubPop_TF(str_ClassID,"NS005","NS","news") then Err_Show
		Dim str_LockID
		str_LockID = FormatIntArr(request.Form("C_NewsID"))
		if str_LockID = "" then
			strShowErr = "<li>�����ѡ��һ���ٲ���</li>"
			set conn=nothing:set user_conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		Conn.execute("Update FS_NS_News set isLock = 1 where ID in ("& str_LockID &")")
		strShowErr = "<li>"& Fs_news.allInfotitle &"�����ɹ�</li>"
		set conn=nothing:set user_conn=nothing
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	If Request.Form("Action") = "Move_News" then
		if not Get_SubPop_TF(str_ClassID,"NS009","NS","news") then Err_Show
		str_s_classIDarray =Replace(Request.Form("s_Classid")," ","")
		str_t_classID=Replace(Request.Form("t_Classid")," ","")
		str_Move_type = Trim(Replace(Request.Form("Move_type")," ",""))
		C_NewsIDarrey = Trim(Replace(Request.Form("C_NewsID")," ",""))
		'�ж��Ƿ����ⲿ��Ŀ
		If str_t_classID="" Then
			strShowErr = "<li>��ѡ��Ŀ����Ŀ</li>"
			set conn=nothing:set user_conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
		Set Tmp_TF_Rs = Conn.execute("select isUrl From FS_NS_NewsClass Where ClassID = '"& NoSqlHack(str_t_classID) &"'")
		if Tmp_TF_Rs(0)=1 then
			set conn=nothing:set user_conn=nothing
			strShowErr = "<li>Ŀ��"& Fs_news.allInfotitle &"��Ŀ����Ϊ�ⲿ��Ŀ</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		Tmp_TF_Rs.close:set Tmp_TF_Rs =nothing
		if str_Move_type = "" then
			strShowErr = "<li>��ѡ��ת������</li>"
			set conn=nothing:set user_conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
		if  str_Move_type = "id" then
			if Trim(C_NewsIDarrey)="" then
				strShowErr = "<li>��ѡ��Ҫת�Ƶ�"& Fs_news.allInfotitle &"!</li>"
				set conn=nothing:set user_conn=nothing
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			tmp_splitarrey_id = split(C_NewsIDarrey,",")
			for tmp_i = LBound(tmp_splitarrey_id) to UBound(tmp_splitarrey_id)
					Set Tmp_rs=server.CreateObject(G_FS_RS)
					Tmp_rs.open "select Classid From [FS_NS_News] where isRecyle=0 and ID="&CintStr(tmp_splitarrey_id(tmp_i))&" order by id desc",Conn,1,3
					if Not Tmp_rs.eof then
						Tmp_rs("ClassID") = NoSqlHack(str_t_classID)
						Tmp_rs.update
					End if
			Next
		Elseif  str_Move_type = "ClassID" then
			if Trim(str_s_classIDarray)="" then
				strShowErr = "<li>��ѡ��Ҫת����Ŀ�µ�"& Fs_news.allInfotitle &"!</li>"
				set conn=nothing:set user_conn=nothing
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			tmp_splitarrey_Classid = split(str_s_classIDarray,",")
			for tmp_i = LBound(tmp_splitarrey_Classid) to UBound(tmp_splitarrey_Classid)
					Set Tmp_rs=server.CreateObject(G_FS_RS)
					Tmp_rs.open "select Classid From [FS_NS_News] where isRecyle=0 and ClassID='"&NoSqlHack(tmp_splitarrey_Classid(tmp_i))&"' order by id desc",Conn,1,3
					do while Not Tmp_rs.eof
						Tmp_rs("ClassID") = NoSqlHack(str_t_classID)
						Tmp_rs.update
						Tmp_rs.movenext
					Loop
			Next
		End if
		Tmp_rs.close:set Tmp_rs=nothing
			strShowErr = "<li>ת�Ƴɹ�</li><li>��Ҫ�������ɲ���Ч!</li>"
			set conn=nothing:set user_conn=nothing
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	End if
	If Request.Form("Action") = "Replace_News" then
		if not Get_SubPop_TF(str_ClassID,"NS010","NS","news") then Err_Show
		Call Replace_News()
	End If
	If Request.Form("Action") = "Copy_News" then
		if not Get_SubPop_TF(str_ClassID,"NS007","NS","news") then Err_Show
		str_s_classIDarray =NoSqlHack(Replace(Request.Form("s_Classid")," ",""))
		str_t_classID=NoSqlHack(Replace(Request.Form("t_Classid")," ",""))
		str_Move_type = NoSqlHack(Trim(Replace(Request.Form("Move_type")," ","")))
		C_NewsIDarrey = NoSqlHack(Trim(Replace(Request.Form("C_NewsID")," ","")))
		'�ж��Ƿ����ⲿ��Ŀ
		If str_t_classID="" Then
			strShowErr = "<li>��ѡ��Ŀ��"& Fs_news.allInfotitle &"��Ŀ</li>"
			set conn=nothing:set user_conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.End
		End if
		Set Tmp_TF_Rs = Conn.execute("select isUrl From FS_NS_NewsClass Where ClassID = '"& NoSqlHack(str_t_classID) &"'")
		if Tmp_TF_Rs(0)=1 then
			set conn=nothing:set user_conn=nothing
			strShowErr = "<li>Ŀ��"& Fs_news.allInfotitle &"��Ŀ����Ϊ�ⲿ��Ŀ</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		Tmp_TF_Rs.close:set Tmp_TF_Rs =nothing
		if str_Move_type = "" then
			strShowErr = "<li>��ѡ��������</li>"
			set conn=nothing:set user_conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		if  str_Move_type = "id" then
			if Trim(C_NewsIDarrey)="" then
				strShowErr = "<li>��ѡ��Ҫ���Ƶ�"& Fs_news.allInfotitle &"!</li>"
				set conn=nothing:set user_conn=nothing
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			tmp_splitarrey_id = split(C_NewsIDarrey,",")
			for tmp_i = LBound(tmp_splitarrey_id) to UBound(tmp_splitarrey_id)
				StrSql="INSERT INTO FS_NS_News([NewsID],[PopId],[ClassID],[SpecialEName],[NewsTitle],[CurtTitle],[NewsNaviContent],[isShowReview],[TitleColor],[titleBorder],[TitleItalic],[IsURL],[URLAddress],[Content],[isPicNews],[NewsPicFile],[NewsSmallPicFile],[Templet],[isPop],[Source],[Editor],[Keywords],[Author],[Hits],[SaveNewsPath],[FileName],[FileExtName],[NewsProperty],[TodayNewsPic],[isLock],[isRecyle],[addtime],[isdraft],[IsAdPic],[AdPicWH],[AdPicLink],[AdPicAdress]) VALUES ("
				Set Tmp_rs=server.CreateObject(G_FS_RS)
				Tmp_rs.open "select * From [FS_NS_News] where isRecyle=0 and ID="&CintStr(tmp_splitarrey_id(tmp_i))&" order by id desc",Conn,1,3
				if Not Tmp_rs.eof then
					StrSql=StrSql & "'"&GetRamCode(15)&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("PopId"))&""
					StrSql=StrSql & ",'"&str_t_classID&"'"
					StrSql=StrSql & ",'"&Tmp_rs("SpecialEName")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("NewsTitle")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("CurtTitle")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("NewsNaviContent")&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("isShowReview"))&""
					StrSql=StrSql & ",'"&Tmp_rs("TitleColor")&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("titleBorder"))&""
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("TitleItalic"))&""
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("IsURL"))&""
					StrSql=StrSql & ",'"&Tmp_rs("URLAddress")&"'"
					StrSql=StrSql & ",'"&Replace(Tmp_rs("Content")&"","'","''")&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("isPicNews"))&""
					StrSql=StrSql & ",'"&Tmp_rs("NewsPicFile")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("NewsSmallPicFile")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("Templet")&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("isPop"))&""
					StrSql=StrSql & ",'"&Tmp_rs("Source")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("Editor")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("Keywords")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("Author")&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("Hits"))&""
					StrSql=StrSql & ",'"&Fs_news.SaveNewsPath(Fs_news.fileDirRule)&"'"
					StrSql=StrSql & ",'"&GetRamCode(4)&"'"
					StrSql=StrSql & ",'"&Tmp_rs("FileExtName")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("NewsProperty")&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("TodayNewsPic"))&""
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("isLock"))&""
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("isRecyle"))&""
					StrSql=StrSql & ",'"&Tmp_rs("addtime")&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("isdraft"))&""
					StrSql=StrSql & ",'"&NUllToStr(Tmp_rs("IsAdPic"))&"'"
					StrSql=StrSql & ",'"&NUllToStr(Tmp_rs("AdPicWH"))&"'"
					StrSql=StrSql & ",'"&NUllToStr(Tmp_rs("AdPicLink"))&"'"
					StrSql=StrSql & ",'"&NUllToStr(Tmp_rs("AdPicAdress"))&"'"
					StrSql=StrSql & ")"
					Conn.ExeCute(StrSql)
				End if
			Next
		Elseif  str_Move_type = "ClassID" then
			if Trim(str_s_classIDarray)="" then
				strShowErr = "<li>��ѡ��Ҫ������Ŀ�µ�"& Fs_news.allInfotitle &"!</li>"
				set conn=nothing:set user_conn=nothing
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if

			tmp_splitarrey_Classid = split(str_s_classIDarray,",")
			
			for tmp_i = LBound(tmp_splitarrey_Classid) to UBound(tmp_splitarrey_Classid)
				Set Tmp_rs=server.CreateObject(G_FS_RS)
				Tmp_rs.open "select * From [FS_NS_News] where isRecyle=0 and ClassID='"&tmp_splitarrey_Classid(tmp_i)&"' order by id desc",Conn,1,3
				
				ReDim ArrSql(0)
				Str_Temp_Flag=True
				
				while Not Tmp_rs.eof
					StrSql="INSERT INTO FS_NS_News([NewsID],[PopId],[ClassID],[SpecialEName],[NewsTitle],[CurtTitle],[NewsNaviContent],[isShowReview],[TitleColor],[titleBorder],[TitleItalic],[IsURL],[URLAddress],[Content],[isPicNews],[NewsPicFile],[NewsSmallPicFile],[Templet],[isPop],[Source],[Editor],[Keywords],[Author],[Hits],[SaveNewsPath],[FileName],[FileExtName],[NewsProperty],[TodayNewsPic],[isLock],[isRecyle],[addtime],[isdraft],[IsAdPic],[AdPicWH],[AdPicLink],[AdPicAdress]) VALUES ("
					StrSql=StrSql & "'"&GetRamCode(15)&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("PopId"))&""
					StrSql=StrSql & ",'"&str_t_classID&"'"
					StrSql=StrSql & ",'"&Tmp_rs("SpecialEName")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("NewsTitle")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("CurtTitle")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("NewsNaviContent")&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("isShowReview"))&""
					StrSql=StrSql & ",'"&Tmp_rs("TitleColor")&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("titleBorder"))&""
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("TitleItalic"))&""
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("IsURL"))&""
					StrSql=StrSql & ",'"&Tmp_rs("URLAddress")&"'"
					StrSql=StrSql & ",'"&Replace(Tmp_rs("Content")&"","'","''")&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("isPicNews"))&""
					StrSql=StrSql & ",'"&Tmp_rs("NewsPicFile")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("NewsSmallPicFile")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("Templet")&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("isPop"))&""
					StrSql=StrSql & ",'"&Tmp_rs("Source")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("Editor")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("Keywords")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("Author")&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("Hits"))&""
					StrSql=StrSql & ",'"&Fs_news.SaveNewsPath(Fs_news.fileDirRule)&"'"
					StrSql=StrSql & ",'"&GetRamCode(4)&"'"
					StrSql=StrSql & ",'"&Tmp_rs("FileExtName")&"'"
					StrSql=StrSql & ",'"&Tmp_rs("NewsProperty")&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("TodayNewsPic"))&""
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("isLock"))&""
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("isRecyle"))&""
					StrSql=StrSql & ",'"&Tmp_rs("addtime")&"'"
					StrSql=StrSql & ","&NUllToStr(Tmp_rs("isdraft"))&""
					StrSql=StrSql & ",'"&NUllToStr(Tmp_rs("IsAdPic"))&"'"
					StrSql=StrSql & ",'"&NUllToStr(Tmp_rs("AdPicWH"))&"'"
					StrSql=StrSql & ",'"&NUllToStr(Tmp_rs("AdPicLink"))&"'"
					StrSql=StrSql & ",'"&NUllToStr(Tmp_rs("AdPicAdress"))&"'"
					StrSql=StrSql & ")"
					
					If Str_Temp_Flag Then
						ArrSql(Ubound(ArrSql))=StrSql
						Str_Temp_Flag=False
					Else
						ReDim Preserve ArrSql(Ubound(ArrSql)+1)
						ArrSql(Ubound(ArrSql))=StrSql
					End If
					Tmp_rs.movenext
					
				Wend
				
				
				For temp_j=Lbound(ArrSql) to Ubound(ArrSql)
					If ArrSql(temp_j)<>"" Then
						Conn.Execute(ArrSql(temp_j))						
					End If
				Next
			Next
		End if
		Tmp_rs.close:set Tmp_rs=nothing
		strShowErr = "<li>���Ƴɹ�</li><li>��Ҫ�������ɲ���Ч!</li>"
		set conn=nothing:set user_conn=nothing
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	if Request.Form("Action") = "setup_News" then
		if not Get_SubPop_TF(str_ClassID,"NS002","NS","news") then Err_Show
		Dim str_set_type,C_Set_NewsIDarrey
		Dim str_NewsProperty_Rec,str_NewsProperty_mar,str_NewsProperty_rev,str_NewsProperty_constr,str_NewsProperty_tt,str_NewsProperty_hots,str_NewsProperty_jc,str_NewsProperty_unr,str_NewsProperty_ann,str_NewsProperty_filt,str_NewsProperty_Remote
		Dim str_NewsProperty_Rec_1,str_NewsProperty_mar_1,str_NewsProperty_rev_1,str_NewsProperty_constr_1,str_NewsProperty_tt_1,str_NewsProperty_hots_1,str_NewsProperty_jc_1,str_NewsProperty_unr_1,str_NewsProperty_ann_1,str_NewsProperty_filt_1,str_NewsProperty_Remote_1
		Dim Set_News_Sql,Property_Str,isShowReviewTF,News_Hits_num,Edit_News_time_Str
		str_s_classIDarray =NoSqlHack(Replace(Request.Form("s_Classid")," ",""))
		str_set_type = NoSqlHack(Replace(Request.Form("set_type")," ",""))
		C_Set_NewsIDarrey = NoSqlHack(Replace(Request.Form("Set_NewsID")," ",""))
		if str_set_type = "" then
			strShowErr = "<li>��ѡ����ѡ����"& Fs_news.allInfotitle &"�������û�����Ŀ��"& Fs_news.allInfotitle &"��������</li>"
			set conn=nothing:set user_conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		
		If str_set_type = "newsid" then
			if Trim(C_Set_NewsIDarrey)="" then
				strShowErr = "<li>��ѡ��Ҫ���õ�"& Fs_news.allInfotitle &"ID!</li>"
				set conn=nothing:set user_conn=nothing
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			Set_News_Sql = " And ID In(" & FormatIntArr(C_Set_NewsIDarrey) & ")"
		ElseIf str_set_type = "classid" then
			if Trim(str_s_classIDarray)="" then
				strShowErr = "<li>��ѡ��Ҫ�趨��Ŀ�µ�"& Fs_news.allInfotitle &"!</li>"
				set conn=nothing:set user_conn=nothing
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			Set_News_Sql = " and ClassID In('" & Replace(str_s_classIDarray,",","','") & "')"
		End IF
		Select Case Trim(Request.Form("Set_Act"))
			Case "Property"
				str_NewsProperty_Rec = NoSqlHack(Request.Form("NewsProperty_Rec"))
				str_NewsProperty_mar = NoSqlHack(Request.Form("NewsProperty_mar"))
				str_NewsProperty_rev = NoSqlHack(Request.Form("NewsProperty_rev"))
				str_NewsProperty_constr =  NoSqlHack(Request.Form("NewsProperty_constr"))
				str_NewsProperty_tt =   NoSqlHack(Request.Form("NewsProperty_tt"))
				str_NewsProperty_hots=   NoSqlHack(Request.Form("NewsProperty_hots"))
				str_NewsProperty_jc=   NoSqlHack(Request.Form("NewsProperty_jc"))
				str_NewsProperty_unr = NoSqlHack(Request.Form("NewsProperty_unr"))
				str_NewsProperty_ann = NoSqlHack(Request.Form("NewsProperty_ann"))
				str_NewsProperty_filt = NoSqlHack(Request.Form("NewsProperty_filt"))
				str_NewsProperty_Remote = NoSqlHack(Request.Form("NewsProperty_Remote"))
				if str_NewsProperty_Rec <>"" then:str_NewsProperty_Rec_1 = 1:else:str_NewsProperty_Rec_1 = 0:End if
				if str_NewsProperty_mar <>"" then:str_NewsProperty_mar_1 = 1:else:str_NewsProperty_mar_1 = 0:End if
				if str_NewsProperty_rev <>"" then:str_NewsProperty_rev_1 = 1:else:str_NewsProperty_rev_1 = 0:End if
				if str_NewsProperty_constr <>"" then:str_NewsProperty_constr_1 = 1:else:str_NewsProperty_constr_1 = 0:End if
				if str_NewsProperty_tt <>"" then:str_NewsProperty_tt_1 = 1:else:str_NewsProperty_tt_1 = 0:End if
				if str_NewsProperty_hots <>"" then:str_NewsProperty_hots_1 = 1:else:str_NewsProperty_hots_1 = 0:End if
				if str_NewsProperty_jc <>"" then:str_NewsProperty_jc_1 = 1:else:str_NewsProperty_jc_1 = 0:End if
				if str_NewsProperty_unr <>"" then:str_NewsProperty_unr_1 = 1:else:str_NewsProperty_unr_1 = 0:End if
				if str_NewsProperty_ann <>"" then:str_NewsProperty_ann_1 = 1:else:str_NewsProperty_ann_1 = 0:End if
				if str_NewsProperty_filt <>"" then:str_NewsProperty_filt_1 = 1:else:str_NewsProperty_filt_1 = 0:End if
				if str_NewsProperty_Remote <>"" then:str_NewsProperty_Remote_1 = 1:else:str_NewsProperty_Remote_1 = 0:End if
				Property_Str = str_NewsProperty_Rec_1&","&str_NewsProperty_mar_1&","&str_NewsProperty_rev_1&","&str_NewsProperty_constr_1&","&str_NewsProperty_Remote_1&","&str_NewsProperty_tt_1&","&str_NewsProperty_hots_1&","&str_NewsProperty_jc_1&","&str_NewsProperty_unr_1&","&str_NewsProperty_ann_1&","&str_NewsProperty_filt_1
				Rem ȡ��ͷ������ʱ��ͬ��ɾ��ͼƬͷ�����е�����
				If str_NewsProperty_tt_1 = 0 Then
					Conn.ExeCute("Delete From FS_NS_TodayPic where NewsID In (Select NewsID From FS_NS_News Where isRecyle=0 And TodayNewsPic = 1 " & Set_News_Sql & ")")
					Conn.ExeCute("Update FS_NS_News Set TodayNewsPic = 0 Where isRecyle=0 And TodayNewsPic = 1 " & Set_News_Sql)
				End If
				Rem 2007-09-05
				Conn.ExeCute("Update FS_NS_News Set NewsProperty = '" & Replace(Property_Str,"'","") & "' where isRecyle=0" & Set_News_Sql & "")
			Case "TempLets"
				IF Replace(Trim(Request.Form("Templet")),"'","") <> "" Then
					Conn.ExeCute("Update FS_NS_News Set Templet = '" & Replace(Trim(Request.Form("Templet")),"'","") & "' where isRecyle=0" & Set_News_Sql & "")
				End IF	
			Case "NewsPop"
				Conn.ExeCute("Update FS_NS_News Set PopID = " & Replace(Trim(Request.Form("PopID")),"'","") & " where isRecyle=0" & Set_News_Sql & "")
			Case "ShowReview"
				if Request.Form("isShowReview")<>"" then
					isShowReviewTF = 1
				Else
					isShowReviewTF = 0
				End if
				Conn.ExeCute("Update FS_NS_News Set isShowReview = " & isShowReviewTF & " where isRecyle=0" & Set_News_Sql & "")
			Case "KeyWords"
				Conn.ExeCute("Update FS_NS_News Set Keywords = '" & Replace(Replace(Trim(Request.Form("KeywordText")),"'","")," ","") & "' where isRecyle=0" & Set_News_Sql & "")
			Case "Hits"
				If Trim(Request.Form("hits")) <> "" And IsNumeric(Trim(Request.Form("hits"))) Then
					IF Clng(Trim(Request.Form("hits")))	> 0 Then
						Conn.ExeCute("Update FS_NS_News Set hits = " & CintStr(Request.Form("hits")) & " where isRecyle=0" & Set_News_Sql & "")
					End IF
				End IF
			Case "EditDate"
				If Trim(Request.Form("addtime")) <> "" And IsDate(Trim(Request.Form("addtime"))) Then
					IF G_IS_SQL_DB = 1 Then
						Conn.ExeCute("Update FS_NS_News Set addtime = '" & NoSqlHack(Request.Form("addtime")) & "' where isRecyle=0" & Set_News_Sql & "")
					Else
						Conn.ExeCute("Update FS_NS_News Set addtime = #" & NoSqlHack(Request.Form("addtime")) & "# where isRecyle=0" & Set_News_Sql & "")
					End IF
				End IF
			Case "ExName"					
				Conn.ExeCute("Update FS_NS_News Set FileExtName = '" & NoSqlHack(Request.Form("FileExtName")) & "' where isRecyle=0" & Set_News_Sql & "")
			Case Else
				str_NewsProperty_Rec = NoSqlHack(Request.Form("NewsProperty_Rec"))
				str_NewsProperty_mar = NoSqlHack(Request.Form("NewsProperty_mar"))
				str_NewsProperty_rev = NoSqlHack(Request.Form("NewsProperty_rev"))
				str_NewsProperty_constr =  NoSqlHack(Request.Form("NewsProperty_constr"))
				str_NewsProperty_tt =   NoSqlHack(Request.Form("NewsProperty_tt"))
				str_NewsProperty_hots=   NoSqlHack(Request.Form("NewsProperty_hots"))
				str_NewsProperty_jc=   NoSqlHack(Request.Form("NewsProperty_jc"))
				str_NewsProperty_unr = NoSqlHack(Request.Form("NewsProperty_unr"))
				str_NewsProperty_ann = NoSqlHack(Request.Form("NewsProperty_ann"))
				str_NewsProperty_filt = NoSqlHack(Request.Form("NewsProperty_filt"))
				str_NewsProperty_Remote = NoSqlHack(Request.Form("NewsProperty_Remote"))
				if str_NewsProperty_Rec <>"" then:str_NewsProperty_Rec_1 = 1:else:str_NewsProperty_Rec_1 = 0:End if
				if str_NewsProperty_mar <>"" then:str_NewsProperty_mar_1 = 1:else:str_NewsProperty_mar_1 = 0:End if
				if str_NewsProperty_rev <>"" then:str_NewsProperty_rev_1 = 1:else:str_NewsProperty_rev_1 = 0:End if
				if str_NewsProperty_constr <>"" then:str_NewsProperty_constr_1 = 1:else:str_NewsProperty_constr_1 = 0:End if
				if str_NewsProperty_tt <>"" then:str_NewsProperty_tt_1 = 1:else:str_NewsProperty_tt_1 = 0:End if
				if str_NewsProperty_hots <>"" then:str_NewsProperty_hots_1 = 1:else:str_NewsProperty_hots_1 = 0:End if
				if str_NewsProperty_jc <>"" then:str_NewsProperty_jc_1 = 1:else:str_NewsProperty_jc_1 = 0:End if
				if str_NewsProperty_unr <>"" then:str_NewsProperty_unr_1 = 1:else:str_NewsProperty_unr_1 = 0:End if
				if str_NewsProperty_ann <>"" then:str_NewsProperty_ann_1 = 1:else:str_NewsProperty_ann_1 = 0:End if
				if str_NewsProperty_filt <>"" then:str_NewsProperty_filt_1 = 1:else:str_NewsProperty_filt_1 = 0:End if
				if str_NewsProperty_Remote <>"" then:str_NewsProperty_Remote_1 = 1:else:str_NewsProperty_Remote_1 = 0:End if
				Property_Str = str_NewsProperty_Rec_1&","&str_NewsProperty_mar_1&","&str_NewsProperty_rev_1&","&str_NewsProperty_constr_1&","&str_NewsProperty_Remote_1&","&str_NewsProperty_tt_1&","&str_NewsProperty_hots_1&","&str_NewsProperty_jc_1&","&str_NewsProperty_unr_1&","&str_NewsProperty_ann_1&","&str_NewsProperty_filt_1
				if Request.Form("isShowReview")<>"" then
					isShowReviewTF = 1
				Else
					isShowReviewTF = 0
				End if	
				If Trim(Request.Form("hits")) = "" Or Not IsNumeric(Trim(Request.Form("hits")))	Then
					News_Hits_num = ""
				Else
					IF Clng(Trim(Request.Form("hits")))	<= 0 Then
						News_Hits_num = ""
					Else
						News_Hits_num = ",hits = " & Clng(Trim(Request.Form("hits"))) & ""
					End IF
				End IF
				If Trim(Request.Form("addtime")) <> "" And IsDate(NoSqlHack((Request.Form("addtime")))) Then
					IF G_IS_SQL_DB = 1 Then
						Edit_News_time_Str = ",addtime = '" & NoSqlHack(Request.Form("addtime")) & "'"
					Else
						Edit_News_time_Str = ",addtime = #" & NoSqlHack(Request.Form("addtime")) & "#"
					End IF
				Else
					Edit_News_time_Str = ""
				End IF
				Rem ȡ��ͷ������ʱ��ͬ��ɾ��ͼƬͷ�����е�����
				If str_NewsProperty_tt_1 = 0 Then
					Conn.ExeCute("Delete From FS_NS_TodayPic where NewsID In (Select NewsID From FS_NS_News Where isRecyle=0 And TodayNewsPic = 1 " & Set_News_Sql & ")")
					Conn.ExeCute("Update FS_NS_News Set TodayNewsPic = 0 Where isRecyle=0 And TodayNewsPic = 1 " & Set_News_Sql)
				End If
				Rem 2007-09-05
				Conn.ExeCute("Update FS_NS_News Set NewsProperty = '" & NoSqlHack(Property_Str) & "',Templet = '" & NoSqlHack(Request.Form("Templet")) & "',PopID = " & CintStr(Request.Form("PopID")) & ",isShowReview = " & CintStr(isShowReviewTF) & ",Keywords = '" & NoSqlHack(Request.Form("KeywordText")) & "'" & News_Hits_num & ",FileExtName = '" & NoSqlHack(Request.Form("FileExtName")) & "'" & Edit_News_time_Str & " Where isRecyle=0" & Set_News_Sql & "")				
		End Select
		strShowErr = "<li>�������óɹ�</li><li>��Ҫ�������ɲ���Ч!</li>"
		set conn=nothing:set user_conn=nothing
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���Ź���___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<script language="JavaScript" src="../../FS_Inc/coolWindowsCalendar.js"></script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="hback">
    <td colspan="3" class="xingmu"><a href="#" class="sd"><strong>���Ź���</strong></a><a href="../../help?Lable=NS_News_Manage" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a>
      <%
	Dim AndSQL
	AndSQL = GetAndSQLOfSearchClass("NS013")
	if str_ClassID <> "" then
		news_SQL = "Select Orderid,id,ClassID,ClassName,ClassEName,IsUrl,AddNewsType from FS_NS_NewsClass where Parentid  = '"& str_ClassID &"' and ReycleTF=0 " & AndSQL & " Order by Orderid desc,ID desc"
	Else
		news_SQL = "Select Orderid,id,ClassID,ClassName,ClassEName,IsUrl,AddNewsType from FS_NS_NewsClass where Parentid  = '0'  and ReycleTF=0 " & AndSQL & " Order by Orderid desc,ID desc"
	End if
	Set obj_news_rs = server.CreateObject(G_FS_RS)
	obj_news_rs.Open news_SQL,Conn,1,3
	if fs_news.addNewsType = 1 then str_addType_1 ="News_add.asp":else:str_addType_1 ="News_add_Conc.asp":end if
	%>
    </td>
  </tr>
  <tr>
    <form name="form1" method="post" action="">
      <td width="48%" height="18" class="hback"><div align="left"><a href="News_Manage.asp">��ҳ</a>��
          <%Response.Write"<a href="""& str_addType_1 &"?ClassID="& NoSqlHack(Request.QueryString("ClassID"))&""">���"& Fs_news.allInfotitle &"</a>|"%>
          <a href="News_Manage.asp?ClassID=<%=NoSqlHack(Request.QueryString("ClassID"))%>">����
          <% =  Fs_news.allInfotitle %>
          </a> ��<a href="News_Manage.asp?ClassID=<%=NoSqlHack(Request.QueryString("ClassID"))%>&isCheck=1&Keyword=<%=NoSqlHack(Request("keyword"))%>&ktype=<%=Request("ktype")%>&SpecialEName=<%=str_SpecialEName%>">�����</a>�� <a href="News_Manage.asp?ClassID=<%=Request.QueryString("ClassID")%>&SpecialEName=<%=str_SpecialEName%>&isCheck=0&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">δ���</a>��<a href="News_Manage.asp?ClassID=<%=Request.QueryString("ClassID")%>&SpecialEName=<%=str_SpecialEName%>&NewsTyp=Constr&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">Ͷ��</a>
          <%if Request.QueryString("NewsTyp")="Constr" then%>
          ��<a href="Constr_stat.asp">���ͳ��</a>
          <%end if%>
        </div></td>
      <td width="43%" class="hback"><div align="center"><a href="News_Manage.asp?ClassID=<%=NoSqlHack(Request.QueryString("ClassID"))%>&SpecialEName=<%=str_SpecialEName%>&NewsTyp=recTF">�Ƽ� </a> ��<a href="News_Manage.asp?ClassID=<%=Request.QueryString("ClassID")%>&SpecialEName=<%=str_SpecialEName%>&NewsTyp=isTop&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">�ö� </a> ��<a href="News_Manage.asp?ClassID=<%=Request.QueryString("ClassID")%>&SpecialEName=<%=str_SpecialEName%>&NewsTyp=hot&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">�ȵ� </a> ��<a href="News_Manage.asp?ClassID=<%=Request.QueryString("ClassID")%>&SpecialEName=<%=str_SpecialEName%>&NewsTyp=pic&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">ͼƬ </a> ��<a href="News_Manage.asp?ClassID=<%=Request.QueryString("ClassID")%>&SpecialEName=<%=str_SpecialEName%>&NewsTyp=highlight&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">���� </a> ��<a href="News_Manage.asp?ClassID=<%=Request.QueryString("ClassID")%>&SpecialEName=<%=str_SpecialEName%>&NewsTyp=bignews&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">ͷ�� </a> ��<a href="News_Manage.asp?ClassID=<%=Request.QueryString("ClassID")%>&SpecialEName=<%=str_SpecialEName%>&NewsTyp=filt&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">�õ�Ƭ</a>����</div></td>
      <td width="9%" class="hback"><a href="Class_Rss.asp?Class=<%=Request.QueryString("ClassID")%>"><img src="../Images/rss200.png" height="15" border="0"></a></td>
    </form>
  </tr>
</table>

<%if NoSqlHack(Request.QueryString("ClassId"))<>"" then%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr>
    <td class="hback"> λ�õ�����<A href="News_Manage.asp">ȫ������</a>
      <%response.write fs_news.navigation(NoSqlHack(Request.QueryString("ClassID")))%></td>
  </tr>
</table>
<%end if%>
<%
	  if Not obj_news_rs.eof then
		Response.Write("<table width=""98%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"" class=""table""> <tr class=""hback""><td>")
		Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" >")
		Response.Write("<tr>")
		icNum = 0
		Do while Not obj_news_rs.eof
			if obj_news_rs("AddNewsType") =1 then
				str_addType = "News_add.asp"
			Else
				str_addType ="News_add_Conc.asp"
			End if
			if obj_news_rs("IsUrl") = 1 then
				isUrlStr = "(<span class=""tx"">��</span>)"
				str_Href = ""
				str_Href_title = ""& obj_news_rs("ClassName") &""
			elseif obj_news_rs("IsUrl") = 2 then
				isUrlStr  = "(<span class=""tx"">��</span>)"
				str_Href = ""
				str_Href_title = ""& obj_news_rs("ClassName") &""
			Else
				isUrlStr = ""
				Rem#####################Ȩ���ж�#####################
				if Get_SubPop_TF(obj_news_rs("ClassID"),"NS001","NS","news") then
					str_Href = "<a href="& str_addType &"?ClassID="&obj_news_rs("ClassID")&"><img src=""../images/add.gif"" border=""0"" alt=""���"& Fs_news.allInfotitle &"""></a>"
				else
					str_Href = ""
				end if
				str_Href_title = "<a href=""News_Manage.asp?ClassID="& obj_news_rs("ClassID") &"&SpecialEName="&str_SpecialEName&""" title=""���������һ����Ŀ"">"& obj_news_rs("ClassName") &"</a>"
			End if
			Set obj_news_rs_1 = server.CreateObject(G_FS_RS)
			obj_news_rs_1.Open "Select Count(ID) from FS_NS_NewsClass where ParentID='"& obj_news_rs("ClassID") &"'",Conn,1,1
			
				'Rem#####################Ȩ���ж�#####################
				'If Get_SubPop_TF(obj_news_rs("Classid"),"NS001","NS","news") Then
					if obj_news_rs_1(0)>0 then
						str_action=  "<img src=""images/jia.gif""></img>"& str_Href_title &""
					Else
						str_action=  "<img src=""images/-.gif""></img>"& str_Href_title &""
					End if
				'Else
				'	str_action = ""
				'End if
				'Rem#####################Ȩ���ж�#####################
			
			obj_news_rs_1.close:set obj_news_rs_1 =nothing
			'�õ���������
			if obj_news_rs("IsUrl") = 0 then
				Set obj_cnews_rs = server.CreateObject(G_FS_RS)
			    obj_cnews_rs.Open "Select ID from FS_NS_News where ClassID='"& obj_news_rs("ClassID") &"' and  isRecyle=0 and isdraft=0 ",Conn,1,1
				news_count = "("&obj_cnews_rs.recordcount&"/"&fs_news.GetTodayNewsCount(obj_news_rs("ClassID"))
				obj_cnews_rs.close:set obj_cnews_rs = nothing
			Else
				news_count = ""
			End if
			'If Get_SubPop_TF(obj_news_rs("Classid"),"NS001","NS","news") Then
				Response.Write "<td height=""22"">"&str_action&isUrlStr&news_count&str_Href&"</td>"
			'End if
			obj_news_rs.MoveNext
			icNum = icNum + 1
			if icNum mod 4 = 0 then
				Response.Write("</tr><tr>")
			End if
		loop
		Response.Write("</tr></table></td></tr></table>")
	End if
	If Request.Form("Action") = "SetUp" then
		Call GetSetUp()
	ElseIf Request.Form("Action") = "move" then
		Call GetMove()
	Elseif Request.Form("Action") = "copy" then
		Call GetCopy()
	Elseif Request.Form("Action") = "replace" then
		Call Getreplace()
	Else
		Call Main()
	End if
	Sub Main()
		Call GetFunctionstr
		if Request("NewsTyp") = "recTF" Then:str_Rec=" and "& CharIndexStr &"(NewsProperty,1,1)='1'":Else:str_Rec="":End if
		if Request("NewsTyp") = "isTop" Then:str_isTop=" and PopID=4 or PoPID=5":Else:str_isTop="":End if
		if Request("NewsTyp") = "hot" Then:str_hot=" and "& CharIndexStr &"(NewsProperty,13,1)='1'":Else:str_hot="":End if
		if Request("NewsTyp") = "pic" Then:str_pic=" and  isPicNews=1":Else:str_pic="":End if
		if Request("NewsTyp") = "highlight" Then:str_highlight=" and "& CharIndexStr &"(NewsProperty,15,1)='1'":Else:str_highlight="":End if
		if Request("NewsTyp") = "bignews" Then:str_bignews="  and "& CharIndexStr &"(NewsProperty,11,1)='1'":Else:str_bignews="":End if
		if Request("NewsTyp") = "filt" Then:str_filt=" and "& CharIndexStr &"(NewsProperty,21,1)='1'":Else:str_filt="":End if
		if Request("NewsTyp") = "Constr" Then:str_Constr=" and "& CharIndexStr &"(NewsProperty,7,1)='1'":Else:str_Constr="":End if
		if Trim(Request("Editor")) <>"" then:str_Editor = " and Editor = '"& Request("Editor")&"'":Else:str_Editor = "":End if
		if str_ClassID<>"" and len(str_ClassID)<=15 then str_ClassID_1 = " and ClassID='"& str_ClassID &"'":Else:str_ClassID_1 = "":End if
		if str_SpecialEname<>"" and not isnull(str_SpecialEname) then
			if G_IS_SQL_DB=0 then
				SQL_SpecialEname = " and instr(1,SpecialEName,'"&str_SpecialEname&"',1)>0"
			else
				SQL_SpecialEname = " and charindex('"&str_SpecialEname&"',SpecialEName)>0"
			end if
		Else
			SQL_SpecialEname=""
		End if
		if Request("isCheck") = "1" then
			str_check = " and islock=0"
		elseif Request("isCheck") = "0" then
			str_check = " and islock=1"
		Else
			str_Check = ""
		End if
		str_Keyword = Trim(Request("keyword"))
		str_ktype =  Trim(Request("ktype"))
		if Trim(str_Keyword) <>"" then
			if str_ktype = "title" then
				str_GetKeyword = " and NewsTitle like '%"& str_Keyword &"%'"
			Elseif str_ktype = "content" then
				str_GetKeyword = " and content like '%"& str_Keyword &"%'"
			Elseif str_ktype = "author" then
				str_GetKeyword = " and author like '%"& str_Keyword &"%'"
			Elseif str_ktype = "editor" then
				str_GetKeyword = " and editor like '%"& str_Keyword &"%'"
			End if
		Else
			str_GetKeyword = ""
		End if
		strpage=request("page")
		if isnull(strpage) or strpage="" or not isnumeric(strpage) Then:strpage=1:end if
		if cbool(strpage)<1 then strpage=1
		
		'�Ż����Ź���ҳ���  08.10.16 Fsj,С������С��100000��¼��
	   ' obj_newslist_id.open "Select NewsID from FS_NS_News where isRecyle=0 and isdraft=0 "& str_Editor & str_Rec & str_isTop & str_hot & str_pic & str_highlight & str_bignews & str_filt & str_Constr & str_ClassID_1 & str_check & SQL_SpecialEname & str_GetKeyword &" Order by PopID desc,addtime desc,ID desc",Conn,1,1
        cPageNo=NoSqlHack(Request.QueryString("Page"))
		If cPageNo="'" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo<=0 Then cPageNo=1	
		
		if 1=cPageNo then 
		    newslist_sql="select top "&int_RPP&" ID,NewsID,PopID,ClassID,NewsTitle,SpecialEName,IsURL,isPicNews,URLAddress,Editor,Hits,NewsProperty,isLock,isRecyle,addtime,author,source from fs_ns_news where isRecyle=0 and isdraft=0 "&  str_Editor & str_Rec & str_isTop & str_hot & str_pic & str_highlight & str_bignews & str_filt & str_Constr & str_ClassID_1 & str_check & SQL_SpecialEname & str_GetKeyword &" order by ID desc"
       else
	        newslist_sql="select top "&int_RPP&" ID,NewsID,PopID,ClassID,NewsTitle,SpecialEName,IsURL,isPicNews,URLAddress,Editor,Hits,NewsProperty,isLock,isRecyle,addtime,author,source from fs_ns_news where id< (select min (id) from (select top "&int_RPP*(cPageNo-1)&" id from fs_ns_news where isRecyle=0 and isdraft=0 "&  str_Editor & str_Rec & str_isTop & str_hot & str_pic & str_highlight & str_bignews & str_filt & str_Constr & str_ClassID_1 & str_check & SQL_SpecialEname & str_GetKeyword &" order by ID desc) as T ) and isRecyle=0 and isdraft=0 "&  str_Editor & str_Rec & str_isTop & str_hot & str_pic & str_highlight & str_bignews & str_filt & str_Constr & str_ClassID_1 & str_check & SQL_SpecialEname & str_GetKeyword &" order by ID desc"

	    end if
	 
	    Set obj_newslist_rs = Server.CreateObject(G_FS_RS)
		obj_newslist_rs.Open newslist_sql,Conn,1,1
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="myForm" method="post" action="">
    <tr class="xingmu">
      <td colspan="3" class="xingmu"><div align="center"> </div>
        <div align="center">
          <% =  Fs_news.allInfotitle %>
          ����</div></td>
      <td width="12%" class="xingmu"><div align="center">¼����/�༭</div></td>
      <td width="9%" class="xingmu"><div align="center">���</div></td>
      <td width="5%" class="xingmu"><div align="center">���</div></td>
      <td width="22%" class="xingmu"><div align="center">����</div></td>
    </tr>
    <%
		if obj_newslist_rs.eof then
		   obj_newslist_rs.close
		   set obj_newslist_rs=nothing
		   Response.Write"<TR  class=""hback""><TD colspan=""7""  class=""hback"" height=""40"">û��"& Fs_news.allInfotitle &"��</TD></TR>"
		else
			str_showTF = 1
			for i=1 to int_RPP
				if obj_newslist_rs.eof Then exit For
					if Get_SubPop_TF(obj_newslist_rs("ClassID"),"NS013","NS","news") then
						Str_GetPopID = obj_newslist_rs("PopID")
						if Str_GetPopID = 5 then
							Str_PopID = "<IMG Src=""images/newstype/5.gif"" border=""0"" alt=""���ö�"& Fs_news.allInfotitle &",����鿴�������"">"
							str_Top = "<a href=News_Manage.asp?NewsID="& obj_newslist_rs("NewsID")&"&ClassId="& obj_newslist_rs("ClassId") &"&Action=signUnTop  onClick=""{if(confirm('ȷ������̶ܹ���')){return true;}return false;}"">���</a>"
						Elseif Str_GetPopID = 4 then
							Str_PopID = "<IMG Src=""images/newstype/4.gif"" border=""0"" alt=""��Ŀ�ö�"& Fs_news.allInfotitle &",����鿴�������"">"
							str_Top = "<a href=News_Manage.asp?NewsID="& obj_newslist_rs("NewsID")&"&ClassId="& obj_newslist_rs("ClassId") &"&Action=signUnTop  onClick=""{if(confirm('ȷ�������Ŀ�̶���')){return true;}return false;}"">���</a>"
						Elseif Str_GetPopID = 3 then
							Str_PopID = "<IMG Src=""images/newstype/3.gif"" border=""0"" alt=""���Ƽ�"& Fs_news.allInfotitle &",����鿴�������"">"
							str_Top = "<a href=News_Manage.asp?NewsID="& obj_newslist_rs("NewsID")&"&ClassId="& obj_newslist_rs("ClassId") &"&Action=signTop  onClick=""{if(confirm('ȷ���̶���')){return true;}return false;}"">�̶�</a>"
						Elseif Str_GetPopID = 2 then
							Str_PopID = "<IMG Src=""images/newstype/2.gif"" border=""0"" alt=""��Ŀ�Ƽ�"& Fs_news.allInfotitle &",����鿴�������"">"
							str_Top = "<a href=News_Manage.asp?NewsID="& obj_newslist_rs("NewsID")&"&ClassId="& obj_newslist_rs("ClassId") &"&Action=signTop onClick=""{if(confirm('ȷ���̶���')){return true;}return false;}"">�̶�</a>"
						Elseif Str_GetPopID = 0 then
							Str_PopID = "<IMG Src=""images/newstype/0.gif"" border=""0"" alt=""һ��"& Fs_news.allInfotitle &",����鿴�������"">"
							str_Top = "<a href=News_Manage.asp?NewsID="& obj_newslist_rs("NewsID")&"&ClassId="& obj_newslist_rs("ClassId") &"&Action=signTop onClick=""{if(confirm('ȷ���̶���')){return true;}return false;}"">�̶�</a>"
						End if
						if obj_newslist_rs("isUrl") = 1 then
							tmp_pictf = ""
							str_UrlTitle = "<a href="""& obj_newslist_rs("URLAddress") &""" target=""_blank""><Img src=""../images/folder/url.gif"" border=""0"" alt=""��������,���ת�������ַ""></img></a>"
						Else
							str_UrlTitle = ""
							if obj_newslist_rs("isPicNews") = 1 then
								tmp_pictf="<a href=""javascript:m_PicUrl('News_Pic_Modify.asp?NewsID="&obj_newslist_rs("NewsID")&"&ClassId="& obj_newslist_rs("ClassId") &"')""><Img src=""../images/folder/img.gif"" alt=""ͼƬ����,�������ͼƬ"" border=""0""></img></a>"
							else
								tmp_pictf="<Img src=""../images/folder/folder_1.gif"" alt=""��������""></img>"
							end if
						end if
		%>
    <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
      <td width="3%" class="hback"><div align="center">
          <input name="C_NewsID" type="checkbox" id="<%if obj_newslist_rs("isUrl") = 1 then %>C_TileID<%else%>C_NewsID<%end if %>" value="<% = obj_newslist_rs("ID")%>">
        </div></td>
      <td width="3%" height="18" class="hback" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(M_Newsid<% = obj_newslist_rs("ID")%>);"  language=javascript><% = Str_PopID %>
      </td>
      <td width="46%" class="hback"><% = str_UrlTitle %>
        <% = tmp_pictf %>
        <a href="News_edit.asp?NewsID=<% = obj_newslist_rs("NewsID")%>&ClassID=<% = obj_newslist_rs("ClassID")%>" title="������ڣ�<% = obj_newslist_rs("addtime")%>">
        <% = GotTopic(obj_newslist_rs("Newstitle"),50)%>
        </a></td>
      <td class="hback"><div align="center"> <a href="News_Manage.asp?ClassID=<% = Request.QueryString("ClassID")%>&Editor=<% = obj_newslist_rs("Editor")%>">
          <% = obj_newslist_rs("Editor")%>
          </a> </div></td>
      <td class="hback"><div align="center"> <font style="font-size:10px">
          <% = obj_newslist_rs("hits")%>
          </font> </div></td>
      <td class="hback"><div align="center">
          <%if obj_newslist_rs("isLock")=1 then response.Write"<a href=""News_Manage.asp?NewsID="& obj_newslist_rs("NewsId") &"&Action=singleCheck"" onClick=""{if(confirm('ȷ��ͨ�������')){return true;}return false;}""><span class=""tx""><b>��</b></span></a>":else response.Write"<a href=""News_Manage.asp?NewsID="& obj_newslist_rs("NewsId") &"&Action=singleUnCheck"" onClick=""{if(confirm('ȷ��������')){return true;}return false;}""><b>��</b></a>"%>
        </div></td>
      <td class="hback"><div align="center"><a href="News_Review.asp?NewsID=<% = obj_newslist_rs("NewsID")%>&ClassID=<% = obj_newslist_rs("ClassID")%>" target="_blank">Ԥ��</a>��
          <% = str_Top%>
          ��<a href="javascript:OpenWindow('lib/Frame.asp?FileName=NewsToJs.asp&Types=PicJs&PageTitle=��ӵ�JS&NewsID=<% = obj_newslist_rs("ID")%>',350,135,window)">����JS</a> ��<a href="News_Manage.asp?NewsID=<% = obj_newslist_rs("NewsID")%>&Action=signDel&ClassId=<% = obj_newslist_rs("ClassId")%>"  onClick="{if(confirm('ȷ��Ҫɾ����\n\n�������ϵͳ��������������ɾ��<% =  Fs_news.allInfotitle %>������վ\n<% =  Fs_news.allInfotitle %>��ɾ��������վ��!\n��Ҫʱ��ɻ�ԭ')){return true;}return false;}">ɾ��</a></div></td>
    </tr>
    <tr id="M_Newsid<% = obj_newslist_rs("ID")%>" style="display:none">
      <td height="35" colspan="7" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr class="hback_1">
            <td width="45%" height="20" class="hback_1"><font style="font-size:12px">
              <% =  Fs_news.allInfotitle %>
              ���ͣ�
              <%
								if trim(obj_newslist_rs("NewsProperty")) <>"" then
									if  split(obj_newslist_rs("NewsProperty"),",")(0) then Response.Write("����")
									if  split(obj_newslist_rs("NewsProperty"),",")(1) then Response.Write("����")
									if  split(obj_newslist_rs("NewsProperty"),",")(2) then Response.Write("����")
									if  split(obj_newslist_rs("NewsProperty"),",")(3) then Response.Write("���")
									if  split(obj_newslist_rs("NewsProperty"),",")(4) then Response.Write("Զͼ��")
									if  split(obj_newslist_rs("NewsProperty"),",")(5) then Response.Write("ͷ��")
									if  split(obj_newslist_rs("NewsProperty"),",")(6) then Response.Write("�ȣ�")
									if  split(obj_newslist_rs("NewsProperty"),",")(7) then Response.Write("����")
									if  split(obj_newslist_rs("NewsProperty"),",")(8) then Response.Write("���")
									if  split(obj_newslist_rs("NewsProperty"),",")(9) then Response.Write("����")
									if  split(obj_newslist_rs("NewsProperty"),",")(10) then Response.Write("�ã�")
								else
									response.Write("--")
								End if
								if trim(obj_newslist_rs("SpecialEName"))<>"" then
									str_Get_Special = ""
									for sp_i = 0 to Ubound(split(obj_newslist_rs("SpecialEName"),","))
										dim rs_speical
										set rs_speical = Conn.execute("select SpecialCName,SpecialEName From FS_NS_Special where SpecialEName='"& NoSqlHack(split(obj_newslist_rs("SpecialEName"),",")(sp_i))&"'")
										if not rs_speical.eof then
											str_Get_Special = str_Get_Special & "<a href=""News_Manage.asp?SpecialEName="&rs_speical("SpecialEName")&"&SpecialCName="&server.URLEncode(rs_speical("SpecialCName"))&""">" &rs_speical("SpecialCName") &"</a>��"
											rs_speical.close:set rs_speical=nothing
										else
											str_Get_Special = "��ר��"
											rs_speical.close:set rs_speical=nothing
										end if
									next
									str_Get_Special = str_Get_Special
								else
									str_Get_Special = "��ר��"
								end if
								%>
              </font></td>
            <td width="22%" class="hback_1"><font style="font-size:12px">���ڣ�
              <% = obj_newslist_rs("addtime")%>
              </font></td>
            <td width="14%" class="hback_1"><font style="font-size:12px">���ߣ� <a href="../../<%=G_USER_DIR%>/showuser.asp?UserName=<% = obj_newslist_rs("author")%>" target="_blank">
              <% = obj_newslist_rs("author")%>
              </a>
              <%
									Dim username
									username=Fs_News.GetUserNumber(obj_newslist_rs("author"))
									if username<>"" then
								%>
              (<%=Fs_News.newsStat(username,0)%>/<font color="#FF0000"><%=Fs_News.newsStat(username,1)%></font>)
              <%End if%>
              </font></td>
            <td width="19%" class="hback_1"><font style="font-size:12px">��Դ��
              <% = obj_newslist_rs("source")%>
              </font></td>
          </tr>
          <tr class="hback_1">
            <td height="20" class="hback_1"><font style="font-size:12px">����ר�⣺
              <% = str_Get_Special %>
              </font></td>
            <td class="hback_1">&nbsp;</td>
            <td class="hback_1">&nbsp;</td>
            <td class="hback_1">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <%
					end if
					obj_newslist_rs.MoveNext
			 Next			
	%>
    <tr>
      <td height="18" colspan="7" class="hback"><div align="right">
          <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form);">
          ѡ��/ȡ������
          <input name="Action" type="hidden" id="Action">
          <input name="ClassID" type="hidden" value="<%= Request.QueryString("ClassID") %>">
          <input type="button" name="Submit" value="ɾ��"  onClick="document.myForm.Action.value='Del';{if(confirm('ȷ���������ѡ��ļ�¼��')){this.document.myForm.submit();return true;}return false;}">
          <input type="submit" name="Submit2" value="��������" onClick="document.myForm.Action.value='SetUp'">
          <input type="submit" name="Submit3" value="�ƶ�" onClick="document.myForm.Action.value='move'">
          <input type="submit" name="Submit3" value="����" onClick="document.myForm.Action.value='copy'">
          <input type="button" name="Submit4" value="����" onClick="document.myForm.Action.value='lock';{if(confirm('ȷ����ѡ��ļ�¼������')){this.document.myForm.submit();return true;}return false;}">
          <input type="button" name="Submit42" value="ͨ�����" onClick="document.myForm.Action.value='unlock';{if(confirm('ȷ����ѡ��ļ�¼ͨ�������')){this.document.myForm.submit();return true;}return false;}">
          <input type="submit" name="Submit5" value="�����滻" onClick="document.myForm.Action.value='replace'">
          <input type="submit" name="Submit52" value="����HTML" onClick="CheckTile(this.form)">
          <input type="Button" name="Submit5222" value="�鵵" onClick="document.myForm.Action.value='Toold';{if(confirm('ȷ�Ϲ鵵��\n��[ȷ��]���鵵��ѡ������Ż��ߴ�����Ϲ鵵Ҫ�������\n�˲���������,��С��ʹ��')){this.document.myForm.submit();return true;}return false;}" <%if Request.QueryString("ClassId") = ""  then response.Write "disabled"%>>
          <input type="Button" name="Submit523" value="����JS" onClick="AddToJS()">
		  <input type="Button" name="Submit524" value="����ר��" onClick="AddToSpecial()">
        </div></td>
    </tr>
  </form>
  <tr>
    <td height="18" colspan="7" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="79%" align="right"><%
          '�Ż� ��ҳ����  Fsj09.10.22
            dim obj_newslist_id,NewsCount
            set obj_newslist_id= Server.CreateObject(G_FS_RS)
            obj_newslist_id.open "Select count(ID) as IDs from FS_NS_News where isRecyle=0 and isdraft=0 "&  str_Editor & str_Rec & str_isTop & str_hot & str_pic & str_highlight & str_bignews & str_filt & str_Constr & str_ClassID_1 & str_check & SQL_SpecialEname & str_GetKeyword &"",Conn,1,1
            NewsCount=obj_newslist_id("IDs")  
            obj_newslist_id.Close
            set obj_newslist_id=nothing
         	response.Write "<p>"&  fPageCountNews(NewsCount,int_RPP,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
		
			'response.Write "<p>"&  fPageCount(obj_newslist_id,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
			End if
	%>
          </td>
        </tr>
      </table></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr>
    <td height="18" colspan="7" class="hback"><span class="tx">˵������<img src="images/newstype/0.gif" width="18" height="12">�鿴<%=Fs_news.allInfotitle %>�����ԣ��� <img src="../Images/Folder/img.gif" width="20" height="16">ͼƬ���Զ�ͼƬ<%=Fs_news.allInfotitle %>���п���޸� !</span></td>
  </tr>
  <tr>
    <form name="form2" method="post" action="News_Manage.asp">
      <td height="18" colspan="7" class="hback"><% =  Fs_news.allInfotitle %>
        �������ؼ���
        <input name="keyword" type="text" id="keyword" value="<% = Request("keyword")%>" size="20">
        <select name="ktype" id="ktype">
          <option value="title" <%if Request("ktype")="title" then response.Write("selected")%>>����</option>
          <option value="content" <%if Request("ktype")="content" then response.Write("selected")%>>����</option>
          <option value="author" <%if Request("ktype")="author" then response.Write("selected")%>>����</option>
          <option value="editor" <%if Request("ktype")="editor" then response.Write("selected")%>>¼����/�༭</option>
        </select>
        <input name="ClassID" type="hidden" id="ClassID" value="<% = str_ClassID %>">
        <input type="submit" name="Submit" value=" �� �� ">
      </td>
    </form>
  </tr>
</table>
<%
End Sub
Sub GetMove()
		Dim str_LockID_move
		str_LockID_move = NoSqlHack(request.Form("C_NewsID"))
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form_m" method="post" action="">
    <tr class="xingmu">
      <td colspan="3" class="xingmu">��ѡ����Ҫת�Ƶ�����Ŀ </td>
    </tr>
    <tr>
      <td width="100%" colspan="3" class="hback">ָ�� I D
        <input name="Move_type" type="radio" value="id" <%if trim(str_LockID_move)<>"" then response.Write("checked")%>>
        <input name="C_NewsID" type="text" id="C_NewsID" value="<% = Replace(str_LockID_move," ","") %>" readonly style="width:80%">
      </td>
    </tr>
    <tr>
      <td class="hback" width="45%">ָ����Ŀ
        <input name="Move_type" type="radio" value="ClassID"  <%if trim(str_LockID_move)="" then response.Write("checked")%>>
      </td>
      <td width="10%" rowspan="3" align="center" class="hback">ת�Ƶ�>>></td>
      <td width="45%" class="hback">ѡ��Ҫת�Ƶ�����Ŀ<Br />
        ע�⣺Ŀ����Ŀ����Ϊ�ⲿ��Ŀ</td>
    </tr>
    <tr>
      <td rowspan="2" class="hback"><select name="s_Classid" id="select" multiple style="width:100%" size="18">
          <%
		  	Dim rs_movelist_rs,str_tmp_move
			Set rs_movelist_rs = server.CreateObject(G_FS_RS)
			rs_movelist_rs.Open "Select ID,ClassID,ClassName,ParentID,ReycleTF from FS_NS_NewsClass where ParentID='0'  and ReycleTF=0",Conn,1,3
			str_tmp_move = ""
			do while not rs_movelist_rs.eof
				str_tmp_move = str_tmp_move & "<option value="""& rs_movelist_rs ("ClassID") &""">"& rs_movelist_rs ("ClassName") &"</option>"
			   str_tmp_move = str_tmp_move & Fs_news.News_ChildNewsList(rs_movelist_rs("ClassID"),"")
			  rs_movelist_rs.movenext
		  Loop
		  	Response.Write str_tmp_move
		  rs_movelist_rs.close:set rs_movelist_rs=nothing
          %>
        </select>
        <input type="button" name="Submit" value="ѡ��������Ŀ" onClick="SelectAllClass()">
        <input type="button" name="Submit" value="ȡ��ѡ����Ŀ" onClick="UnSelectAllClass()">
      </td>
      <td class="hback"><select name="t_Classid" size="18" id="select"  style="width:100%">
          <% = str_tmp_move %>
        </select>
      </td>
    </tr>
    <tr>
      <td class="hback">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="3" class="hback"><div align="center">
          <input name="Action" type="hidden" id="Action" value="Move_News">
          <input type="submit" name="Submit6" value="ȷ����ʼת��">
        </div></td>
    </tr>
  </form>
</table>
<%
End Sub
Sub GetCopy()
		Dim str_LockID_move
		str_LockID_move = request.Form("C_NewsID")
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form_m" method="post" action="">
    <tr class="xingmu">
      <td colspan="3" class="xingmu">��ѡ����Ҫ���Ƶ�����Ŀ </td>
    </tr>
    <tr>
      <td width="44%" colspan="3" class="hback">ָ�� I D
        <input name="Move_type" type="radio" value="id" <%if trim(str_LockID_move)<>"" then response.Write("checked")%>>
        <input name="C_NewsID" type="text" id="C_NewsID" value="<% = Replace(str_LockID_move," ","") %>" readonly style="width:75%">
      </td>
    </tr>
    <tr>
      <td width="45%" class="hback">ָ����Ŀ
        <input name="Move_type" type="radio" value="ClassID"  <%if trim(str_LockID_move)="" then response.Write("checked")%>>
      </td>
      <td width="10%" rowspan="3" align="center" class="hback">���Ƶ�>>></td>
      <td width="45%" class="hback">ѡ��Ҫ���Ƶ�����Ŀ<br />
        ע�⣺Ŀ����Ŀ����Ϊ�ⲿ��Ŀ</td>
    </tr>
    <tr>
      <td rowspan="2" class="hback"><select name="s_Classid" id="select" multiple style="width:100%" size="18">
          <%
		  	Dim rs_movelist_rs,str_tmp_move
			Set rs_movelist_rs = server.CreateObject(G_FS_RS)
			rs_movelist_rs.Open "Select ID,ClassID,ClassName,ParentID,ReycleTF from FS_NS_NewsClass where ParentID='0'  and ReycleTF=0",Conn,1,3
	    	str_tmp_move = ""
			do while not rs_movelist_rs.eof
				str_tmp_move = str_tmp_move & "<option value="""& rs_movelist_rs ("ClassID") &""">"& rs_movelist_rs ("ClassName") &"</option>"
			   str_tmp_move = str_tmp_move & Fs_news.News_ChildNewsList(rs_movelist_rs("ClassID"),"")
			  rs_movelist_rs.movenext
		  Loop
		  	Response.Write str_tmp_move
		  rs_movelist_rs.close:set rs_movelist_rs=nothing
          %>
        </select>
        <input type="button" name="Submit" value="ѡ��������Ŀ" onClick="SelectAllClass()">
        <input type="button" name="Submit" value="ȡ��ѡ����Ŀ" onClick="UnSelectAllClass()">
      </td>
      <td class="hback"><select name="t_Classid" size="18" id="select"  style="width:100%">
          <% = str_tmp_move %>
        </select>
      </td>
    </tr>
    <tr>
      <td class="hback">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="3" class="hback"><div align="center">
          <input name="Action" type="hidden" id="Action" value="Copy_News">
          <input type="submit" name="Submit6" value="ȷ����ʼ����">
        </div></td>
    </tr>
  </form>
</table>
<%
End Sub
Sub GetSetUp()
Dim str_LockID_set
		str_LockID_set = NoSqlHack(request.Form("C_NewsID"))
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form_m" method="post" action=""> 
  <input name="Action" type="hidden" id="Action" value="setup_News">
  <input name="Set_Act" type="hidden" id="Set_Act" value="">
    <tr>
      <td width="36%" rowspan="8" align="center" class="hback"><div align="left">ָ���ɣ�
          <input name="set_type" type="radio" value="newsid" <%if trim(str_LockID_set)<>"" then response.Write("checked")%>>
          <br>
          ָ����Ŀ
          <input type="radio" name="set_type" value="classid" <%if trim(str_LockID_set)="" then response.Write("checked")%>>
          <select name="s_Classid" id="select" multiple style="width:100%" size="18">
            <%
		  	Dim rs_movelist_rs,str_tmp_move
			Set rs_movelist_rs = server.CreateObject(G_FS_RS)
			rs_movelist_rs.Open "Select ID,ClassID,ClassName,ParentID,ReycleTF from FS_NS_NewsClass where ParentID='0'  and ReycleTF=0",Conn,1,3
			str_tmp_move = ""
			do while not rs_movelist_rs.eof
				str_tmp_move = str_tmp_move & "<option value="""& rs_movelist_rs ("ClassID") &""">"& rs_movelist_rs ("ClassName") &"</option>"
			   str_tmp_move = str_tmp_move & Fs_news.News_ChildNewsList(rs_movelist_rs("ClassID"),"")
			  rs_movelist_rs.movenext
		  Loop
		  	Response.Write str_tmp_move
		  rs_movelist_rs.close:set rs_movelist_rs=nothing
          %>
          </select>
          <input type="button" name="Submit" value="ѡ��������Ŀ" onClick="SelectAllClass()">
          <input type="button" name="Submit" value="ȡ��ѡ����Ŀ" onClick="UnSelectAllClass()">
        </div></td>
      <td width="11%" rowspan="8" align="center" class="hback"><input name="Set_NewsID" type="hidden" id="Set_NewsID" style="width:95%" value="<% = Replace(str_LockID_set," ","") %>" readonly="readonly">
        ��������</td>
      <td width="53%" height="46" class="hback"> �����ԣ�
        <input name="NewsProperty_Rec" type="checkbox" id="NewsProperty" value="1">
        �Ƽ�
        <input name="NewsProperty_mar" type="checkbox" id="NewsProperty" value="1" checked>
        ����
        <input name="NewsProperty_rev" type="checkbox" id="NewsProperty" value="1" checked>
        ��������
        <input name="NewsProperty_constr" type="checkbox" id="NewsProperty" value="1">
        Ͷ��
        <input name="NewsProperty_tt" type="checkbox" id="NewsProperty" value="1">
        ͷ�� <br>
        <input name="NewsProperty_hots" type="checkbox" id="NewsProperty" value="1" disabled="disabled">
        �ȵ�
        <input name="NewsProperty_jc" type="checkbox" id="NewsProperty" value="1">
        ����
        <input name="NewsProperty_unr" type="checkbox" id="NewsProperty" value="1">
        ������
        <input name="NewsProperty_ann" type="checkbox" id="NewsProperty" value="1">
        ���� 
		<span id="str_filt" style="display1:none" title="��ע����ѡ����ȫ��ͼƬ���ţ�">
        <input name="NewsProperty_filt" type="checkbox" id="NewsProperty" value="1">
        �õ�</span>   
		<input type="button" name="button6" value="����" onClick="document.form_m.Set_Act.value='Property';document.form_m.submit();">
		</td>
    </tr>
    <tr>
      <td class="hback">ģ���壺
        <input type="text" name="Templet" value="<%=Replace("/"& G_TEMPLETS_DIR &"/NewsClass/news.htm","//","/")%>" style="width:50%">
        <input name="Submit53" type="button" id="selNewsTemplet" value="ѡ��ģ��"  onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=Replace(G_VIRTUAL_ROOT_DIR&"/"& G_TEMPLETS_DIR,"//","/") %>',400,300,window,document.form_m.Templet);document.form_m.Templet.focus();">
<input type="button" name="button6" value="����" onClick="document.form_m.Set_Act.value='TempLets';document.form_m.submit();">
      </td>
    </tr>
    <tr>
      <td class="hback">Ȩ���أ�
        <select name="PopID" id="PopID">
          <option value="5">���ö�</option>
          <option value="4">��Ŀ�ö�</option>
          <option value="0" selected>һ��</option>
        </select>
		<input type="button" name="button6" value="����" onClick="document.form_m.Set_Act.value='NewsPop';document.form_m.submit();">
      </td>
    </tr>
    <tr>
      <td class="hback"> �����ۣ�
        <input name="isShowReview" type="checkbox" id="isShowReview" value="1">
        �������ʾ&quot;����&quot;����
		<input type="button" name="button6" value="����" onClick="document.form_m.Set_Act.value='ShowReview';document.form_m.submit();"></td>
    </tr>
    <tr>
      <td class="hback"> �ؼ��֣�
        <input name="KeywordText" type="text" id="KeywordText" size="15" maxlength="255">
        <input name="KeyWords" type="hidden" id="KeyWords">
        <select name="selectKeywords" id="selectKeywords" style="width:120px" onChange=Dokesite_s(this.options[this.selectedIndex].value)>
          <option value="" selected>ѡ��ؼ���</option>
          <option value="Clean" style="color:red">���</option>
          <%=Fs_news.GetKeywordslist("",1)%>
        </select>
		<input type="button" name="button6" value="����" onClick="document.form_m.Set_Act.value='KeyWords';document.form_m.submit();">
      </td>
    </tr>
    <tr>
      <td class="hback">�������
        <input name="hits" type="text" id="hits" value="0" size="30">
		<input type="button" name="button6" value="����" onClick="document.form_m.Set_Act.value='Hits';document.form_m.submit();">
      </td>
    </tr>
    <tr>
      <td class="hback">�������ڣ�
          <input name="addtime" type="text" id="addtime" onFocus="setday(this)" value="<% = now()%>" size="30">
          <IMG onClick="addtime.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom"><span id="EditTime_Alt"></span>
		  <input type="button" name="button6" value="����" onClick="document.form_m.Set_Act.value='EditDate';document.form_m.submit();">
	  </td>
	 </tr>
	 <tr>
      <td class="hback">��չ����  
          <select name="FileExtName" id="FileExtName">
            <option value="html" <%if fs_news.fileExtName = 1 then response.Write("selected")%>>.html</option>
            <option value="htm" <%if fs_news.fileExtName = 0 then response.Write("selected")%>>.htm</option>
            <option value="shtml" <%if fs_news.fileExtName = 2 then response.Write("selected")%>>.shtml</option>
            <option value="shtm" <%if fs_news.fileExtName = 3 then response.Write("selected")%>>.shtm</option>
            <option value="asp" <%if fs_news.fileExtName = 4 then response.Write("selected")%>>.asp</option>
          </select>
		   <input type="button" name="button6" value="����" onClick="document.form_m.Set_Act.value='ExName';document.form_m.submit();">
        </td>
    </tr>
    <tr>
      <td colspan="3" class="hback">˵�������Ϊ����
        <% =  Fs_news.allInfotitle %>
        �����ò�������</td>
    </tr>
    <tr>
      <td colspan="3" class="hback"><div align="center">
		   <input type="button" name="button6" value="ȷ������ȫ��" onClick="document.form_m.Set_Act.value='Set_All';document.form_m.submit();">
          <input type="reset" name="Submit7" value="�����趨">
        </div></td>
    </tr>
  </form>
</table>
<%
End Sub
Sub Getreplace()
Dim str_LockID_rep
		str_LockID_rep = NoSqlHack(request.Form("C_NewsID"))
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form_m" method="post" action="">
    <tr>
      <td width="36%" rowspan="10" align="center" valign="top" class="hback"><div align="left">ָ���ɣ�
          <input name="rep_type" type="radio" value="newsid" <%if trim(str_LockID_rep)<>"" then response.Write("checked")%>>
          <br>
          ָ����Ŀ
          <input type="radio" name="rep_type" value="classid" <%if trim(str_LockID_rep)="" then response.Write("checked")%>>
          <select name="s_Classid" id="select" multiple style="width:100%" size="20">
            <%
		  	Dim rs_movelist_rs,str_tmp_move
			Set rs_movelist_rs = server.CreateObject(G_FS_RS)
			rs_movelist_rs.Open "Select ID,ClassID,ClassName,ParentID,ReycleTF from FS_NS_NewsClass where ParentID='0' and ReycleTF=0",Conn,1,3
			str_tmp_move = ""
			do while not rs_movelist_rs.eof
				str_tmp_move = str_tmp_move & "<option value="""& rs_movelist_rs ("ClassID") &"""  multiple>"& rs_movelist_rs ("ClassName") &"</option>"
			   str_tmp_move = str_tmp_move & Fs_news.News_ChildNewsList(rs_movelist_rs("ClassID"),"")
			  rs_movelist_rs.movenext
		  Loop
		  	Response.Write str_tmp_move
		  rs_movelist_rs.close:set rs_movelist_rs=nothing
          %>
          </select>
          <input type="button" name="Submit" value="ѡ��������Ŀ" onClick="SelectAllClass()">
          <input type="button" name="Submit" value="ȡ��ѡ����Ŀ" onClick="UnSelectAllClass()">
        </div></td>
      <td width="11%" rowspan="10" align="center" class="hback"><input name="rep_NewsID" type="hidden" id="rep_NewsID" style="width:95%" value="<% = Replace(str_LockID_rep," ","") %>" readonly="readonly">
        �����滻</td>
      <td width="53%" height="20" class="hback"><input name="rep_select_type" type="checkbox" id="rep_select_type" value="title">
        ����
        <input name="rep_select_type" type="checkbox" id="rep_select_type" value="Content" checked>
        ���� </td>
    </tr>
    <tr>
      <td class="hback"><input name="AdvanceTF" type="radio" id="radio" onClick="SwitchNewsType('snews');" value="snews" checked>
        һ���滻
        <input name="AdvanceTF" type="radio" id="AdvanceTF" value="adnews" onClick="SwitchNewsType('adnews');">
        �߼��滻</td>
    </tr>
    <tr  id="rep_1" style="display:">
      <td class="hback">Ҫ�滻���ַ�<br>
        <textarea name="s_Content" rows="8" id="s_Content" style="width:95%"></textarea>
      </td>
    </tr>
    <tr>
      <td class="hback">�滻����ַ�</td>
    </tr>
    <tr>
      <td class="hback"><textarea name="t_Content" rows="8" id="t_Content" style="width:95%"></textarea>
      </td>
    </tr>
    <tr id="rep_2" style="display:none">
      <td class="hback">��ʼ�ַ���<br>
        <textarea name="start_char" rows="8" id="start_char" style="width:95%"></textarea>
      </td>
    </tr>
    <tr id="rep_3" style="display:none">
      <td class="hback">�����ַ���<br>
        <textarea name="end_char" rows="8" id="end_char" style="width:95%"></textarea>
      </td>
    </tr>
    <tr>
      <td colspan="3" class="hback"><div align="center">
          <input name="Action" type="hidden" id="Action" value="Replace_News">
          <input type="submit" name="Submit6" value="ȷ����ʼ�滻">
          <input type="reset" name="Submit7" value="�����趨">
        </div></td>
    </tr>
  </form>
</table>
<script language="JavaScript" type="text/JavaScript">
function SwitchNewsType(AdvanceTF)
{
	switch (AdvanceTF)
	{
	case "snews":
		document.getElementById('rep_1').style.display='';
		document.getElementById('rep_2').style.display='none';
		document.getElementById('rep_3').style.display='none';
		break;
	case "adnews":
		document.getElementById('rep_1').style.display='none';
		document.getElementById('rep_2').style.display='';
		document.getElementById('rep_3').style.display='';
	}
}
</script>
<%End Sub%>
</body>
</html>
<%
set obj_newslist_rs = nothing
obj_news_rs.close
set obj_news_rs =nothing
set Fs_news = nothing
%>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
<script language="JavaScript" type="text/JavaScript" src="js/Public.js"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" type="text/JavaScript">
function opencat(cat)
{
  if(cat.style.display=="none"){
     cat.style.display="";
  } else {
     cat.style.display="none";
  }
}
<!---------------ѡ��ר��,���ר�⺯����ʼ----------------->
function SelectSpecial()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('lib/SelectspecialFrame.asp',400,300,window);
	if (ReturnValue.indexOf('***')!=-1)
	{
		TempArray = ReturnValue.split('***');
		if (document.form_m.SpecialID.value.search(TempArray[1])==-1)
		{
		if(document.all.SpecialID.value=='') document.all.SpecialID.value=TempArray[1];
		else document.all.SpecialID.value=document.all.SpecialID.value+','+TempArray[1];
		if(document.all.SpecialID_EName.value=='') document.all.SpecialID_EName.value=TempArray[0];
		else document.all.SpecialID_EName.value=document.all.SpecialID_EName.value+','+TempArray[0];
		}
	}
}
function dospclear1()
	{
	document.form_m.SpecialID.value = '';
	document.form_m.SpecialID_EName.value = '';
	}
<!---------------------ѡ��ר��,���ר�⺯������------------>
function SelectAllClass(){
  for(var i=0;i<document.form_m.s_Classid.length;i++){
    document.form_m.s_Classid.options[i].selected=true;}
}
function UnSelectAllClass(){
  for(var i=0;i<document.form_m.s_Classid.length;i++){
    document.form_m.s_Classid.options[i].selected=false;}
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = myForm.elements[i];
    if (e.name != 'chkall')
       e.checked = myForm.chkall.checked;
    }
	}
function CheckTile(form)
{
   for (var i=0;i<form.elements.length;i++)
    {
    var e = myForm.elements[i];
    if (e.id == 'C_TileID')
       e.checked ="";
    }
    document.myForm.Action.value='makehtml';

} 
	

function m_PicUrl(gotoURL) {
	   var open_url = gotoURL;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=1,width=550,height=480');
}
function AddToJS()
{
	var SelectedNews='';
	var ListObjArray=document.myForm.C_NewsID
	if (ListObjArray.length)
	{
		for (i=0;i<ListObjArray.length;i++)
		{
			if (ListObjArray[i].checked==true)
			{
				if (ListObjArray[i].value!=null)
				{
					if (!isNaN(ListObjArray[i].value))
					{
						if (SelectedNews=='') SelectedNews=ListObjArray[i].value;
						else  SelectedNews=SelectedNews+'***'+ListObjArray[i].value;
					}
				}
			}
		}
	}
	else
	{
		if (ListObjArray.checked)
		{
			SelectedNews=ListObjArray.value;
		}	
	}
	if (SelectedNews!='') OpenWindow('lib/Frame.asp?FileName=NewsToJs.asp&Types=PicJs&PageTitle=��ӵ�JS&NewsID='+SelectedNews,350,135,window);
	else alert('��ѡ������');
}

function AddToSpecial()
{
	var SelectedNews='';
	var ListObjArray=document.myForm.C_NewsID
	if (ListObjArray.length)
	{
		for (i=0;i<ListObjArray.length;i++)
		{
			if (ListObjArray[i].checked==true)
			{
				if (ListObjArray[i].value!=null)
				{
					if (!isNaN(ListObjArray[i].value))
					{
						if (SelectedNews=='') SelectedNews=ListObjArray[i].value;
						else  SelectedNews=SelectedNews+'***'+ListObjArray[i].value;
					}
				}
			}
		}
	}
	else
	{
		if (ListObjArray.checked)
		{
			SelectedNews=ListObjArray.value;
		}	
	}
	if (SelectedNews!='') OpenWindow('lib/Frame.asp?FileName=NewsToSpecial.asp&Types=ToSp&PageTitle=������ŵ�ר��&NewsID='+SelectedNews,400,135,window);
	else alert('��ѡ������');
}

</script>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->
<%
Sub Replace_News()
	str_s_classIDarray =NoSqlHack(Replace(Request.Form("s_Classid")," ",""))
	str_rep_type = NoSqlHack(Trim(Replace(Request.Form("rep_type")," ","")))
	C_NewsIDarrey = NoSqlHack(Trim(Replace(Request.Form("rep_NewsID")," ","")))
	str_rep_select_type = NoSqlHack(Trim(Replace(Request.Form("rep_select_type")," ","")))
	str_AdvanceTF = NoSqlHack(Trim(Replace(Request.Form("AdvanceTF")," ","")))
	str_s_Content = NoSqlHack(Trim(Replace(Request.Form("s_Content")," ","")))
	str_t_Content = NoSqlHack(Trim(Replace(Request.Form("t_Content")," ","")))
	str_start_char = NoSqlHack(Trim(Replace(Request.Form("start_char")," ","")))
	str_end_char = NoSqlHack(Trim(Replace(Request.Form("end_char")," ","")))

	If str_rep_type = "" Then
		strShowErr = "<li>��ѡ�� ID/��Ŀ �滻</li>"
	End If
	If str_rep_select_type = "" Then
		strShowErr = strShowErr & "<li>��ѡ���滻 ����/����</li>"
	End If

	If str_AdvanceTF = "snews" Then
		If str_s_Content="" Then
			strShowErr = strShowErr & "<li>����дҪ�滻���ַ�</li>"
		End If
	ElseIf str_AdvanceTF = "adnews" Then
		If str_start_char="" Then
			strShowErr = strShowErr & "<li>����дҪ�滻�Ŀ�ʼ�ַ���</li>"
		End If
		If str_end_char="" Then
			strShowErr = strShowErr & "<li>����дҪ�滻�Ľ����ַ���</li>"
		End If
	Else
		If str_t_Content="" Then
			strShowErr = strShowErr & "<li>��ѡ�� һ��/�߼� �滻</li>"
		End If
	End If
	If str_t_Content="" Then
		strShowErr = strShowErr & "<li>����д�滻����ַ�</li>"
	End If
	If strShowErr<>"" Then
		set conn=nothing:set user_conn=nothing
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.End
	End If
	If str_rep_type = "newsid" then
		If Trim(C_NewsIDarrey)="" Then
			strShowErr = "<li>��ѡ��Ҫ�滻��"& Fs_news.allInfotitle &"!</li>"
			set conn=nothing:set user_conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		'tmp_splitarrey_id = split(C_NewsIDarrey,",")
		'for tmp_i = LBound(tmp_splitarrey_id) to UBound(tmp_splitarrey_id)
		Set Tmp_rs=server.CreateObject(G_FS_RS)
		Tmp_rs.open "select NewsTitle,Content From [FS_NS_News] where isRecyle=0 and ID IN ("&FormatIntArr(C_NewsIDarrey)&") order by id desc",Conn,1,3
		While Not Tmp_rs.eof
			If str_AdvanceTF = "snews" Then
				If InStr(str_rep_select_type,"title")>0 Then Tmp_rs("NewsTitle")=Replace(Tmp_rs("NewsTitle"),str_s_Content,str_t_Content)
				If InStr(str_rep_select_type,"Content")>0 Then Tmp_rs("Content")=Replace(Tmp_rs("Content"),str_s_Content,str_t_Content)
			Else
				Set f_PLACE_OBJ = New RegExp
				f_PLACE_OBJ.Pattern = str_start_char&"(.*)"&str_end_char
				f_PLACE_OBJ.IgnoreCase = True
				f_PLACE_OBJ.Global = True
				f_PLACE_OBJ.Multiline = True
				If InStr(str_rep_select_type,"title")>0 Then Tmp_rs("NewsTitle")=f_PLACE_OBJ.Replace(Tmp_rs("NewsTitle"),str_t_Content)
				If InStr(str_rep_select_type,"Content")>0 Then Tmp_rs("Content")=f_PLACE_OBJ.Replace(Tmp_rs("Content"),str_t_Content)
				Set f_PLACE_OBJ = Nothing
			End If
			Tmp_rs.Update
			Tmp_rs.MoveNext
		Wend
		Tmp_rs.close:set Tmp_rs=nothing
		'Next
	Elseif str_rep_type = "classid" then
		if Trim(str_s_classIDarray)="" then
			strShowErr = "<li>��ѡ��Ҫ�滻��"& Fs_news.allInfotitle &"��Ŀ!</li>"
			set conn=nothing:set user_conn=nothing
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if

		tmp_splitarrey_Classid = split(str_s_classIDarray,",")
		for tmp_i = LBound(tmp_splitarrey_Classid) to UBound(tmp_splitarrey_Classid)
			Set Tmp_rs=server.CreateObject(G_FS_RS)
			Tmp_rs.open "select ID,NewsTitle,Content From [FS_NS_News] where isRecyle=0 and ClassID='"&tmp_splitarrey_Classid(tmp_i)&"' order by id desc",Conn,1,3
			While Not Tmp_rs.eof
				If str_AdvanceTF = "snews" Then
					If InStr(str_rep_select_type,"title")>0 Then Tmp_rs("NewsTitle")=Replace(Tmp_rs("NewsTitle"),str_s_Content,str_t_Content)
					If InStr(str_rep_select_type,"Content")>0 Then Tmp_rs("Content")=Replace(""&Tmp_rs("Content"),str_s_Content,str_t_Content)
				Else
					Set f_PLACE_OBJ = New RegExp
					f_PLACE_OBJ.Pattern = str_start_char&"(.*)"&str_end_char
					f_PLACE_OBJ.IgnoreCase = True
					f_PLACE_OBJ.Global = True
					f_PLACE_OBJ.Multiline = True
					If InStr(str_rep_select_type,"title")>0 Then Tmp_rs("NewsTitle")=f_PLACE_OBJ.Replace(Tmp_rs("NewsTitle"),str_t_Content)
					If InStr(str_rep_select_type,"Content")>0 Then Tmp_rs("Content")=f_PLACE_OBJ.Replace(""&Tmp_rs("Content"),str_t_Content)
					Set f_PLACE_OBJ = Nothing
				End If
				Tmp_rs.Update
				Tmp_rs.MoveNext
			Wend
		Next
		Tmp_rs.close:set Tmp_rs=nothing
	End if
	strShowErr = "<li>�滻�ɹ�</li><li>��Ҫ�������ɲ���Ч!</li>"
	set conn=nothing:set user_conn=nothing
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End Sub
Function NUllToStr(num)
	If IsNull(num) Then
		NUllToStr=Null
	Else
		NUllToStr=num
	End if
End Function

%>

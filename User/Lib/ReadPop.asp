<%
Sub GetNewsID(newsId,f_type,isClass)
	newsId  = NoSqlHack(newsId)
	f_type  = NoSqlHack(f_type)
	isClass = CintStr(isClass)
	dim JudgeTF,RsNews,user_rs,into_rs
	Set RsNews = Conn.Execute("select GroupName,PointNumber,FS_Money,InfoID,PopType,isClass from FS_MF_POP where InfoID='"&NoSqlHack(newsId)&"' and PopType='"&NoSqlHack(f_type)&"' and isClass="& isClass &"")
	if Not RsNews.eof then
		if RsNews("PointNumber")>0 then
			set user_rs = User_Conn.execute("select id,UserNumber,InfoId,SubType,addtime,isClass From FS_ME_POP where UserNumber='"&session("FS_UserNumber")&"' and InfoId='"& NoSqlHack(newsId) &"' and SubType='"& NoSqlHack(f_type) &"' and isClass="& isClass &"")
			if user_rs.eof then
				if Fs_User.NumIntegral<RsNews("PointNumber") then
					Response.Write("<script>alert(""����\n���ĵ������㣬�����������Ϣ.\n���������Ҫ����:"& RsNews("PointNumber") &"\n���ȷ��ǰ����ֵ����"");location.href='" & s_savepath & "/card.asp';</script>")
					Response.End
				else
					User_Conn.execute("Update FS_ME_Users Set Integral=Integral-"&RsNews("PointNumber")&" where UserNumber='"&session("FS_UserNumber")&"'")
					set into_rs = Server.CreateObject(G_FS_RS)
					into_rs.open "select * From FS_ME_POP where 1=0",User_Conn,1,3
					into_rs.addnew
					into_rs("UserNumber")=session("FS_UserNumber")
					into_rs("InfoId")=newsId
					into_rs("SubType")=f_type
					into_rs("addtime")=now
					into_rs("isClass")=isClass
					into_rs.update
					into_rs.close:set into_rs = nothing
				end if
			end if
			user_rs.close:set user_rs = nothing
		end if
		if RsNews("FS_Money")>0 then
			set user_rs = User_Conn.execute("select id,UserNumber,InfoId,SubType,addtime,isClass From FS_ME_POP where UserNumber='"&session("FS_UserNumber")&"' and InfoId='"& newsId &"' and SubType='"& f_type &"' and isClass="&isClass&"")
			if user_rs.eof then
				if Fs_User.NumFS_Money<RsNews("FS_Money") then
					Response.Write("<script>alert(""����\n���Ľ�Ҳ��㣬�����������Ϣ.\n���������Ҫ���:"& RsNews("FS_Money") &"\n�뵽��Ա���ĳ�ֵ"");location.href='" & s_savepath & "/card.asp';</script>")
					Response.End
				else
					User_Conn.execute("Update FS_ME_Users Set FS_Money=FS_Money-"&RsNews("FS_Money")&" where UserNumber='"&session("FS_UserNumber")&"'")
					set into_rs = Server.CreateObject(G_FS_RS)
					into_rs.open "select * From FS_ME_POP where 1=0",User_Conn,1,3
					into_rs.addnew
					into_rs("UserNumber")=session("FS_UserNumber")
					into_rs("InfoId")=newsId
					into_rs("SubType")=f_type
					into_rs("addtime")=now
					into_rs("isClass")=isClass
					into_rs.update
					into_rs.close:set into_rs = nothing
				end if
			end if
			user_rs.close:set user_rs = nothing
		end if
		if trim(rsNews("GroupName"))<>"" and not isNull(trim(rsNews("GroupName"))) then
			dim GroupArray,i,GroupRs,UGroupName
			Set GroupRs = User_Conn.execute("select GroupName from FS_ME_Group where GroupId="& Fs_User.NumGroupID &"")
			If GroupRs.Eof Then
				UGroupName = "$NothingGroup$"
			Else
				UGroupName = GroupRs("GroupName")
			End If
			if instr(rsNews("GroupName"),UGroupName)=0 then
				Response.Write("<script>alert(""����\n��û���Ȩ��.\n���������Ҫ:  ��"& rsNews("GroupName") &"��  ��Ա�鼶��������.\n���ȷ���رո�ҳ."");window.close();</script>")
				Response.End
			end if
		end if
	end if
	RsNews.close:set RsNews =nothing
End Sub
%>






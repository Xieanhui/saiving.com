<% Option Explicit %>
<%session.CodePage="936"%>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	response.Charset = "gb2312"
	Dim Conn,VS_Sql,VS_RS,VS_RS1,strShowErr,IID,TID,ItemValue,VisitIP ,IPInterView,IsSigned,Steps,MaxNum
	Dim TmpStr,TmpArr
	
	TID = NoSqlHack(request.QueryString("TID"))
	IID = NoSqlHack(request.form("Items"))
	MaxNum = request.QueryString("MaxNum")
	ItemValue =  NoSqlHack(request.QueryString("ItemsInput"))
	if TID="" or not isnumeric(TID) then response.Write("ͶƱ�����������!"&TID)  :  response.End()
	if IID="" then response.Write("��������ѡ��һ��!") : response.End()
	IID = replace(IID," ","")
	TmpArr = split(IID,",")
	if MaxNum="" or not isnumeric(MaxNum) then MaxNum = 1
	if ubound(TmpArr)+1 > cint(MaxNum) then response.Write("ѡ��ܳ���"&MaxNum&"��!") : response.End()
	
	MF_Default_Conn 
	
	''''�õ���������
	VS_Sql = "select top 1 IPInterView,IsSigned from FS_VS_SysPara"
	set VS_RS = Conn.execute(VS_Sql)
	if Not VS_RS.eof then 
		IPInterView = VS_RS(0)
		IsSigned = VS_RS(1)
	else
		IPInterView = 1
		IsSigned = 0	
	end if
	VS_RS.close

	if IsSigned = 1 then 
		''����ע��
		if isnull(session("FS_UserNumber")) or session("FS_UserNumber")="" then response.Write("<a href=""/user/login.asp"" target=_blank>���ȵ�½��ͶƱ!</a>") : ConnClose
	end if

	'''''''''''''''''''''''''''''''''''
	VisitIP = NoSqlHack(request.ServerVariables("HTTP_X_FORWARDED_FOR"))
	If VisitIP = "" then
		VisitIP = NoSqlHack(request.ServerVariables("REMOTE_ADDR"))
	End If
	VisitIP = CheckIpSafe(VisitIP)
	
	Set VS_RS = Conn.Execute("Select top 1 VoteTime from FS_VS_Items_Result where TID = "&CintStr(TID)&" and VoteIp='"&NoSqlHack(VisitIP)&"' order by RID desc")
	If VS_RS.eof then
		for each TmpStr in TmpArr
			''''�õ�������д��ѡ���IID
			if ItemValue<>"" then 
				set VS_RS1 = Conn.execute("select IID from FS_VS_Items where ItemMode = 2 and TID="&CintStr(TID)&" and IID="&CintStr(TmpStr))
				if Not VS_RS1.eof then 
					Conn.execute("insert into FS_VS_Items_Result (IID,TID,ItemValue,VoteIp,VoteTime,UserNumber) values ("&CintStr(TmpStr)&","&CintStr(TID)&",'"&NoSqlHack(ItemValue)&"','"&NoSqlHack(VisitIP)&"','"&now&"','"&session("FS_UserNumber")&"')")
				else
					Conn.execute("insert into FS_VS_Items_Result (IID,TID,ItemValue,VoteIp,VoteTime,UserNumber) values ("&CintStr(mpStr)&","&CintStr(TID)&",'"&ItemValue&"','"&NoSqlHack(VisitIP)&"','"&now&"','"&session("FS_UserNumber")&"')")
				end if
				VS_RS1.close
			else
				Conn.execute("insert into FS_VS_Items_Result (IID,TID,ItemValue,VoteIp,VoteTime,UserNumber) values ("&CintStr(TmpStr)&","&CintStr(TID)&",'"&ItemValue&"','"&VisitIP&"','"&now&"','"&session("FS_UserNumber")&"')")	
			end if
		next
		session("OldTID") = TID
		response.Cookies("FS_Cookies")(cstr(TID)&"_IID") = ","&IID&","
		response.Cookies("FS_Cookies")(cstr(TID)&"_ItemValue") = ItemValue
		response.Write("��л���ͶƱ!")
		RsClose : ConnClose
	else
		if cstr(session("OldTID")) = cstr(TID) then 
			if datediff("n",VS_RS("VoteTime"),now) < IPInterView then 
				response.Cookies("FS_Cookies")(cstr(TID)&"_IID") = ""
				response.Cookies("FS_Cookies")(cstr(TID)&"_ItemValue") = ""
				response.Write("�����ظ�ͶƱ!"&IPInterView&"���Ӻ���Լ���.")
				RsClose : ConnClose
			end if
		end if
		for each TmpStr in TmpArr
			''''�õ�������д��ѡ���IID
			if ItemValue<>"" then 
				set VS_RS1 = Conn.execute("select IID from FS_VS_Items where ItemMode = 2 and TID="&CintStr(TID)&" and IID="&CintStr(TmpStr))
				if Not VS_RS1.eof then 
					Conn.execute("insert into FS_VS_Items_Result (IID,TID,ItemValue,VoteIp,VoteTime,UserNumber) values ("&CintStr(TmpStr)&","&CintStr(TID)&",'"&NoSqlHack(ItemValue)&"','"&NoSqlHack(VisitIP)&"','"&now&"','"&session("FS_UserNumber")&"')")
				else
					Conn.execute("insert into FS_VS_Items_Result (IID,TID,ItemValue,VoteIp,VoteTime,UserNumber) values ("&CintStr(TmpStr)&","&CintStr(TID)&",'"&ItemValue&"','"&NoSqlHack(VisitIP)&"','"&now&"','"&session("FS_UserNumber")&"')")
				end if
				VS_RS1.close
			else
				Conn.execute("insert into FS_VS_Items_Result (IID,TID,ItemValue,VoteIp,VoteTime,UserNumber) values ("&CintStr(TmpStr)&","&CintStr(TID)&",'"&ItemValue&"','"&NoSqlHack(VisitIP)&"','"&now&"','"&session("FS_UserNumber")&"')")	
			end if
		next
		response.Cookies("FS_Cookies")(cstr(TID)&"_IID") = ","&IID&","
		response.Cookies("FS_Cookies")(cstr(TID)&"_ItemValue") = ItemValue
		session("OldTID") = TID
		response.Write("��л���ͶƱ!")
		RsClose : ConnClose
	End if


Sub RsClose()
	VS_RS.Close
	Set VS_RS = Nothing
end Sub
Sub ConnClose()
	Set Conn = Nothing
	response.End()
End Sub
%>

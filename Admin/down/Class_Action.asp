<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%
Dim Conn,User_Conn,strShowErr
MF_Default_Conn
MF_User_Conn
'session�ж�
MF_Session_TF 
'Ȩ���ж�
if not MF_Check_Pop_TF("DS_Class") then Err_Show
dim Fs_news
set Fs_news = new Cls_News
Fs_News.GetSysParam()
If Not Fs_news.IsSelfRefer Then response.write "�Ƿ��ύ����":Response.end
if Request.Form("actionType") = "xml" then
	Response.Write("����xml")
	Response.end
Elseif Request.Form("actionType") = "html" then
	Response.Write("����html")
	Response.end
End if
if Request("Action") = "makeXML" then
	dim cid
	cid=NoSqlHack(Replace(Request("Cid")," ",""))
	if trim(cid)=empty then
			strShowErr = "<li>��ѡ����Ŀ</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	end if
	Response.Redirect "Class_makerss.asp?cid="& cid &""
	Response.end
End if
%> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ŀ����___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu">��Ŀ����<a href="../../help?Lable=NS_Class_Action" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr> 
    <td height="18" class="hback"><div align="left"><a href="Class_Manage.asp">������ҳ</a>��<%if MF_Check_Pop_TF("DS010") then%><a href="Class_add.asp?ClassID=&Action=add">��Ӹ���Ŀ</a>��<%end if%><%if MF_Check_Pop_TF("DS013") then%><a href="Class_Action.asp?Action=one">һ����Ŀ����</a>��<%end if%><%if MF_Check_Pop_TF("DS014") then%><a href="Class_Action.asp?Action=n">N����Ŀ����</a>��<%end if%><%if MF_Check_Pop_TF("DS013") then%><a href="Class_Action.asp?Action=reset"   onClick="{if(confirm('ȷ�ϸ�λ������Ŀ��\n\n���ѡ��ȷ�������е���Ŀ������Ϊһ������!!')){return true;}return false;}">��λ������Ŀ</a>��<%end if%><%if MF_Check_Pop_TF("DS015") then%><a href="Class_Action.asp?Action=unite">��Ŀ�ϲ�</a>��<%end if%><%if MF_Check_Pop_TF("DS_Class") then%><a href="Class_Action.asp?Action=allmove">��Ŀת��</a>��<%end if%><%if MF_Check_Pop_TF("DS012") then%><a href="Class_Action.asp?Action=clearClass"  onClick="{if(confirm('ȷ�����������Ŀ���������\n\n���ѡ��ȷ��,���е���Ŀ�����ؽ����ŵ�����վ��!!')){return true;}return false;}">ɾ��������Ŀ</a> <%end if%>
        <a href="../../help?Lable=NS_Class_Action_1" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></div></td>
  </tr>
</table>
<%
Dim str_Action,obj_news_rs,obj_news_rs_1,isUrlStr
str_Action = Request("Action")
Select Case str_Action
		Case "one"
			Call OrderOne()
		Case "n"
			Call OrderN()
		Case "Order_one"
			Call UpdateOrderID()
		Case "Order_n"
			Call UpdateOrderIDN()
		Case "del"
			Call delclass()
		Case "reset"
			Call resetClass()
		Case "unite"
			Call Classunite()
		Case "Saveunite"
			Call Saveunite()
		Case "allmove"
			Call allmove()
		Case "clearClass"
			Call clearClass()
		Case "clear"
			Call one_clear()
		Case "SaveAllmove"
			Call allmove_save()
End Select
Sub OrderOne()
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td height="22" class="xingmu">��Ŀ����</td>
    <td height="22" class="xingmu"><div align="center">ID</div></td>
    <td class="xingmu"><div align="center">����</div></td>
  </tr>
  <%
	Set obj_news_rs = server.CreateObject(G_FS_RS)
	obj_news_rs.Open "Select Orderid,id,ClassID,ParentID,ClassName from FS_DS_Class where Parentid  = '0' Order by Orderid desc,ID desc",Conn,1,3
	Do while Not obj_news_rs.eof 
	%>
  <form name="ClassForm" method="post" action="Class_Action.asp">
    <tr class="hback"> 
      <td width="39%" height="31" class="hback"><img src="images/%2B.gif" width="15" height="15"> 
        <% = obj_news_rs("ClassName") %> </td>
      <td width="21%" class="hback"><div align="center"> 
          <% = obj_news_rs("ID") %>
        </div></td>
      <td width="40%" class="hback"><div align="center"> 
          <input name="OrderID" type="text" id="OrderID" value="<% = obj_news_rs("OrderID") %>" size="4" maxlength="3">
          <input name="ClassID" type="hidden" id="ClassID" value="<% = obj_news_rs("ClassID") %>">
          <input name="Action" type="hidden" id="ClassID" value="Order_one">
          <input type="submit" name="Submit" value="����Ȩ��(�������)">
        </div></td>
    </tr>
  </form>
  <%
		obj_news_rs.MoveNext
	loop
%>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr>
    <td class="hback">˵����Ȩ��(�������)����Խ���ŵ�Խ��ǰ.���Ȩ��(�������)������ͬ���͸���ID������</td>
  </tr>
</table>
<%
obj_news_rs.close
set obj_news_rs =nothing
End Sub
Sub OrderN()
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td height="22" class="xingmu">��Ŀ����</td>
    <td height="22" class="xingmu"><div align="center">ID</div></td>
    <td class="xingmu"><div align="center">����</div></td>
  </tr>
  <%
	Set obj_news_rs = server.CreateObject(G_FS_RS)
	obj_news_rs.Open "Select Orderid,id,ClassID,ParentID,ClassName from FS_DS_Class where Parentid  = '0' Order by Orderid desc,ID desc",Conn,1,3
	Do while Not obj_news_rs.eof 
	%>
    <tr class="hback"> 
      
    <td width="39%" height="28" class="hback"><img src="images/%2B.gif" width="15" height="15"> 
      <% = obj_news_rs("ClassName") %> </td>
      <td width="21%" class="hback"><div align="center"> 
          <% = obj_news_rs("ID") %>
        </div></td>
      <td width="40%" class="hback"><div align="center"> </div></td>
    </tr>
  <%
  		Response.Write(Fs_news.GetChildNewsList_order(obj_news_rs("ClassID"),""))
		obj_news_rs.MoveNext
loop
%>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr>
    <td class="hback">˵����Ȩ��(�������)����Խ���ŵ�Խ��ǰ.���Ȩ��(�������)������ͬ���͸���ID������</td>
  </tr>
</table>
<%
obj_news_rs.close
set obj_news_rs =nothing
End Sub
Dim obj_unite_rs,tmp_str_list


Sub Classunite()
Set obj_unite_rs = server.CreateObject(G_FS_RS)
obj_unite_rs.Open "Select Orderid,id,ClassID,ClassName,ParentID from FS_DS_Class where Parentid  = '0' and ReycleTF=0 Order by Orderid desc,ID desc",Conn,1,3
%>
<form name="form1" method="post" action="Class_Action.asp">
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr class="hback"> 
      <td height="22" colspan="2" class="xingmu"><div align="center">Դ��Ŀ</div></td>
      <td height="22" colspan="2" class="xingmu"><div align="center">Ŀ����Ŀ</div></td>
    </tr>
    <tr class="hback"> 
      <td width="45%" height="28" class="hback"> <div align="center">
<select name="SourceClassID" id="SourceClassID" style="width:80%">
            <%
		  tmp_str_list  = ""
		  do while Not obj_unite_rs.eof 
            tmp_str_list = tmp_str_list &"<option value="""& obj_unite_rs("ClassID") &","& obj_unite_rs("ParentID") &""">+"& obj_unite_rs("ClassName") &"</option>"& Chr(13) & Chr(10)
			tmp_str_list = tmp_str_list &Fs_news.UniteChildNewsList(obj_unite_rs("ClassID"),"")
			 obj_unite_rs.movenext
		 Loop
		 obj_unite_rs.close
		 set obj_unite_rs = nothing
		 Response.Write tmp_str_list
		 %>
          </select>
        </div></td>
      <td colspan="2" class="hback"> <div align="center">�ϲ���&gt;&gt;&gt; </div></td>
      <td width="46%" class="hback"><select name="TargetClassID" id="SourceClassID" style="width:80%">
          <%= tmp_str_list%> </select></td>
    </tr>
    <tr class="hback"> 
      <td height="28" colspan="4" class="hback"><div align="center"> 
          <input type="button" name="Submit2" value="ȷ���ϲ���Ŀ" onClick="{if(confirm('ȷ���ϲ���?\n\n�ϲ��󽫲��ܻ�ԭ!!!')){this.document.form1.submit();return true;}return false;}">
          <input name="Action" type="hidden" id="Action" value="Saveunite">
        </div></td>
    </tr>
  </table>
</form>
<%End Sub%>
<%
Sub allmove()
Set obj_unite_rs = server.CreateObject(G_FS_RS)
obj_unite_rs.Open "Select Orderid,id,ClassID,ClassName,ParentID from FS_DS_Class where Parentid  = '0' and ReycleTF=0 Order by Orderid desc,ID desc",Conn,1,1
%>
<form name="form1" method="post" action="Class_Action.asp">
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr class="hback"> 
      <td height="22" colspan="2" class="xingmu"><div align="center">Դ��Ŀ</div></td>
      <td height="22" colspan="2" class="xingmu"><div align="center">Ŀ����Ŀ</div></td>
    </tr>
    <tr class="hback"> 
      <td width="45%" height="28" class="hback"> <div align="center">
<select name="SourceClassID" id="SourceClassID" style="width:80%">
            <%
		  tmp_str_list  = ""
		  do while Not obj_unite_rs.eof 
            tmp_str_list = tmp_str_list &"<option value="""& obj_unite_rs("ClassID") &","&  obj_unite_rs("ParentID") &""">+"& obj_unite_rs("ClassName") &"</option>"& Chr(13) & Chr(10)
			tmp_str_list = tmp_str_list &Fs_news.UniteChildNewsList(obj_unite_rs("ClassID"),"")
			 obj_unite_rs.movenext
		 Loop
		 obj_unite_rs.close
		 set obj_unite_rs = nothing
		 Response.Write tmp_str_list
		 %>
          </select>
        </div></td>
      <td colspan="2" class="hback"> <div align="center">ת�Ƶ�&gt;&gt;&gt;</div></td>
      <td width="46%" class="hback"><select name="TargetClassID" id="SourceClassID" style="width:80%">
          <option value="0000000000">ת�Ƶ���Ŀ¼</option>
          <%= tmp_str_list%> </select></td>
    </tr>
    <tr class="hback"> 
      <td height="28" colspan="4" class="hback"><div align="center"> 
          <input type="button" name="Submit2" value="ȷ��ת����Ŀ" onClick="{if(confirm('ȷ��ת����Ŀ��?')){this.document.form1.submit();return true;}return false;}">
          <input name="Action" type="hidden" id="Action" value="SaveAllmove">
        </div></td>
    </tr>
  </table>
</form>
<%End Sub%>
</body>
</html>
<%
sub UpdateOrderID()
		if not MF_Check_Pop_TF("DS013") then Err_Show
		Dim ClassID,OrderID
		ClassID = NoSqlHack(Request.Form("ClassID"))
		OrderID = NoSqlHack(Request.Form("OrderID"))
		if ClassID="" then
			strShowErr = "<li>�������:ClassID</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		else
			ClassID=ClassID
		end if
		if isnumeric(OrderID)= false then
			strShowErr = "<li>�������:�����������д��ȷ������</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		if OrderID="" then
			strShowErr = "<li>�������:OrderID</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		Conn.execute "update FS_DS_Class set OrderID=" & CintStr(OrderID) & " where ClassID='" & NoSqlHack(ClassID) &"'"
		Response.Redirect "Class_Action.asp?Action=one"
end sub
sub UpdateOrderIDN()
		if not MF_Check_Pop_TF("DS013") then Err_Show
		Dim N_ClassID,N_OrderID
		N_ClassID = NoSqlHack(Request.Form("ClassID"))
		N_OrderID = CintStr(Request.Form("OrderID"))
		if N_ClassID="" then
			strShowErr = "<li>�������:ClassID</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		else
			N_ClassID=N_ClassID
		end if
		if isnumeric(N_OrderID)= false then
			strShowErr = "<li>�������:�����������д��ȷ������</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		if N_OrderID="" then
			strShowErr = "<li>�������:OrderID</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		Conn.execute "update FS_DS_Class set OrderID=" & N_OrderID & " where ClassID='" & NoSqlHack(N_ClassID) &"'"
		Response.Redirect "Class_Action.asp?Action=n"
end sub
Sub delclass()
		if not MF_Check_Pop_TF("DS012") then Err_Show
		Dim str_delClassID,str_tmp_DelclassTF,rs_tmp
		str_delClassID = NoSqlHack(Request.QueryString("ClassID"))
		if str_delClassID="" then
			strShowErr = "<li>�������:û����ĿClassID</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		Conn.execute("Delete From FS_DS_Class Where ClassID='"& NoSqlHack(str_delClassID) &"'")
		'ɾ������
		set rs_tmp = Conn.execute("select DownLoadID from FS_DS_List where ClassID='"& NoSqlHack(str_delClassID) &"'")
		do while not rs_tmp.eof 
			Conn.execute("Delete From FS_DS_Address Where DownLoadID='"& NoSqlHack(rs_tmp(0)) &"'")
			rs_tmp.movenext
		loop
		rs_tmp.close
		Conn.execute("Delete From FS_DS_List Where ClassID='"& NoSqlHack(str_delClassID) &"'")
		str_tmp_DelclassTF=0
		Response.write Fs_news.DelChildNewsList(str_delClassID,str_tmp_DelclassTF)
		'����ɾ����¼��־
		'**************���� 
		strShowErr = "<li>��Ŀ�Ѿ�����ɾ���ɹ�</li>"
		Call MF_Insert_oper_Log("ɾ����Ŀ","ɾ����ĿClassID:"& str_delClassID &"��������Ŀ,ͬʱɾ���˴���Ŀ�µ�������Ϣ,����ɾ��",now,session("admin_name"),"DS")
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Class_Manage.asp")
		Response.end
End Sub
Sub resetClass()
	if not MF_Check_Pop_TF("DS014") then Err_Show
	Conn.Execute("Update FS_DS_Class Set ParentID ='0'")
	'���������־
		Call MF_Insert_oper_Log("��λ��Ŀ","��������Ŀ��λΪһ����Ŀ",now,session("admin_name"),"DS")
		strShowErr = "<li>��λ������Ŀ�ɹ�</li><li>���е���Ŀ�Ѿ�����Ϊһ������</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Class_Manage.asp")
		Response.end
End Sub
'�ϲ�Ŀ¼
Sub Saveunite()
		if not MF_Check_Pop_TF("DS015") then Err_Show
		Dim str_SourceClassID,str_TargetClassID
		str_SourceClassID = NoSqlHack(Request.Form("SourceClassID"))
		str_TargetClassID = NoSqlHack(Request.Form("TargetClassID"))
		if split(str_SourceClassID,",")(0) = split(str_TargetClassID,",")(0) then
			strShowErr = "<li>ԴĿ¼��Ŀ��Ŀ¼����һ��</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		Dim obj_Url_rs
		set obj_Url_rs = Conn.execute("select isUrl from FS_DS_Class where ClassID='"& NoSqlHack(split(str_TargetClassID,",")(0)) &"'")
		if obj_Url_rs("isUrl") = 1 then
			strShowErr = "<li>���ܰ���Ŀ�ϲ����ⲿ��Ŀ!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		'Conn.execute("Update FS_DS_List set DownLoadID='"& split(str_TargetClassID,",")(0)&"'  where ClassID ='"& split(str_SourceClassID,",")(0)&"'")
		Conn.execute("Update FS_DS_List set ClassID='"& NoSqlHack(split(str_TargetClassID,",")(0))&"'  where ClassID ='"& NoSqlHack(split(str_SourceClassID,",")(0))&"'")
		'��������������ݿ�
		'Conn.execute("Delete From  FS_DS_Address  where ClassID ='"& split(str_SourceClassID,",")(0)&"'")
		'ɾ��Դ��Ŀ��
		Conn.execute("Delete From FS_DS_Class  where ClassID ='"& NoSqlHack(split(str_SourceClassID,",")(0))&"'")
		Dim ob_Tmp_rs
		set ob_Tmp_rs = Conn.execute("select ClassID,ParentID From FS_DS_Class Where ParentID='"& NoSqlHack(split(str_SourceClassID,",")(0)) &"' order by id desc")
		if Not ob_Tmp_rs.eof then
			do while Not ob_Tmp_rs.eof 
				Conn.execute("Update FS_DS_Class Set ParentID ='"&  NoSqlHack(split(str_SourceClassID,",")(1)) &"' where ClassID ='"& NoSqlHack(ob_Tmp_rs("ClassID"))&"'")
				ob_Tmp_rs.movenext
			Loop
		End if
		ob_Tmp_rs.close:set ob_Tmp_rs = nothing
		'������������
		'******************����
		Call MF_Insert_oper_Log("�ϲ���Ŀ","����ĿClassID:"&  split(str_SourceClassID)(0)&"�ϲ���ClassID:"&  split(str_TargetClassID)(0) &"��",now,session("admin_name"),"DS")
		strShowErr = "<li>�ϲ���Ŀ�ɹ�</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Class_Manage.asp")
		Response.end
End Sub
'�ƶ���Ŀ
Sub allmove_save()
		if not MF_Check_Pop_TF("DS_Class") then Err_Show
		Dim str_SourceClassID_move,str_TargetClassID_move
		str_SourceClassID_move = NoSqlHack(Request.Form("SourceClassID"))
		str_TargetClassID_move = NoSqlHack(Request.Form("TargetClassID"))
		if split(str_SourceClassID_move,",")(0) = split(str_TargetClassID_move,",")(0) then
			strShowErr = "<li>ԴĿ¼��Ŀ��Ŀ¼����һ��</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		if str_TargetClassID_move ="0000000000" then
			Conn.execute("Update FS_DS_Class Set ParentID ='0' where ClassID ='"&  NoSqlHack(split(str_SourceClassID_move,",")(0))&"'")
		Else
			Dim obj_Url_rs_1
			set obj_Url_rs_1 = Conn.execute("select isUrl from FS_DS_Class where ClassID='"& NoSqlHack(split(str_TargetClassID_move,",")(0)) &"'")
			if obj_Url_rs_1("isUrl") = 1 then
				strShowErr = "<li>���ܰ���Ŀת���ⲿ��Ŀ!</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			Conn.execute("Update FS_DS_Class Set ParentID ='"&  split(str_TargetClassID_move,",")(0) &"' where ClassID ='"&  split(str_SourceClassID_move,",")(0)&"'")
			Dim ob_Tmp_rs_1
			set ob_Tmp_rs_1 = Conn.execute("select ClassID,ParentID From FS_DS_Class Where ParentID='"& NoSqlHack(split(str_SourceClassID_move,",")(0)) &"' order by id desc")
			if Not ob_Tmp_rs_1.eof then
				do while Not ob_Tmp_rs_1.eof 
					Conn.execute("Update FS_DS_Class Set ParentID ='"&  NoSqlHack(split(str_SourceClassID_move,",")(1)) &"' where ClassID ='"& NoSqlHack(ob_Tmp_rs_1("ClassID"))&"'")
					ob_Tmp_rs_1.movenext
				Loop
			End if
			ob_Tmp_rs_1.close:set ob_Tmp_rs_1 = nothing
		End if
		'������������
		'******************����
		Call MF_Insert_oper_Log("ת����Ŀ","����ĿClassID:"& split(str_SourceClassID_move)(0) &"ת�Ƶ�ClassID:"& split(str_TargetClassID_move)(0)&"��",now,session("admin_name"),"DS")
		strShowErr = "<li>��Ŀת�Ƴɹ���</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Class_Manage.asp")
		Response.end
End Sub
Sub clearClass()
		if not MF_Check_Pop_TF("DS016") then Err_Show
		Conn.execute("delete from FS_DS_Address")
		Conn.execute("delete from FS_DS_List")
		Conn.execute("Delete From FS_DS_Class")
		Call MF_Insert_oper_Log("ɾ��������Ŀ","ɾ������վ����Ŀ",now,session("admin_name"),"DS")
		strShowErr = "<li>����ɾ����Ŀ�ɹ�</li><li>ɾ��������Ŀ�µ����سɹ�</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Class_Manage.asp")
		Response.end
End Sub
Sub one_clear()
		if not MF_Check_Pop_TF("DS016") then Err_Show
		Conn.execute("Delete From FS_DS_Address where DownLoadID = '"& NoSqlHack(Request.QueryString("DownLoadID"))&"'")
		Conn.execute("Delete From FS_DS_List where ClassID = '"& NoSqlHack(Request.QueryString("ClassID"))&"'")
		Call MF_Insert_oper_Log("�������","�����Classid:"& Request.QueryString("ClassID") &"�µ���������",now,session("admin_name"),"DS")
		strShowErr = "<li>��Ŀ�µ���������ɹ�</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
End Sub
set Fs_news = nothing
%>






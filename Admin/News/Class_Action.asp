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
'Call MF_Check_Pop_TF("NS_Class_000001")
'�õ���Ա���б�
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
	if not Get_SubPop_TF("","NS022","NS","class") then Err_Show
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
'---2007-02-06 By Ken ---����ɾ����Ŀ-------
If Request.Form("Action") = "DelClass" then
	Dim NewsConfigTF,DleClassObj,VRootStr,Str_DelClassID
	Dim DelClassIDStr,DelNum,DelXmlStr,DelAllClassIDStr
	Dim Arr_Num,Arr_i,Class_i,Class_Num
	Dim ArrC_Num,ArrC_i,CheckTopClassRs
	'---��ȡɾ����Ŀ��id
	DelClassIDStr = Request.Form("Cid")
	'---��ȡϵͳ��������ֱ��ɾ������ɾ��������վ��
	NewsConfigTF = fs_news.ReycleTF
	'---ȡ��ϵͳ������Ŀ¼
	If G_VIRTUAL_ROOT_DIR <> "" Then
		VRootStr = "/" & G_VIRTUAL_ROOT_DIR
	Else
		VRootStr = ""
	End If
	'---���ɾ����ĿidΪ�շ��ش���
	If DelClassIDStr = "" Then
		strShowErr = "<li><font color=""#ff0000"">û��ѡ��Ҫɾ������Ŀ</font></li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
		if Instr(DelClassIDStr,",") > 0 Then
		    Dim Idlist,listi
		    Idlist = Split(DelClassIDStr,",")
		    For listi=0 To ubound(Idlist)
	            if not Get_SubPop_TF(NoSqlHack(Idlist(listi)),"NS021","NS","class") then
                    Err_Show
	            Else
	                if not Get_SubPop_TF(NoSqlHack(Idlist(listi)),"NS003","NS","news") then Err_Show
	            End if
			 Next
		Else
            if not Get_SubPop_TF(NoSqlHack(DelClassIDStr),"NS021","NS","class") then
                Err_Show
            Else
               if not Get_SubPop_TF(NoSqlHack(DelClassIDStr),"NS003","NS","news") then Err_Show
            End if
		End if
		'---����Ҫɾ������Ŀid���и�ʽ��
		If Instr(DelClassIDStr,",") > 0 Then
			DelClassIDStr = Replace(Replace(DelClassIDStr," ,",","),", ",",")
			Class_Num = Split(DelClassIDStr,",")
			For Class_i = LBound(Class_Num) To UBound(Class_Num)
				Set CheckTopClassRs = Conn.ExeCute("Select ParentID,ClassID From FS_NS_NewsClass Where ClassID = '" & Class_Num(Class_i) & "' And ReycleTF = 0 Order By OrderID Desc,ID Desc")
				If CheckTopClassRs(0) <> "0" Then
					If InStrTopIDTF(Class_Num(Class_i),DelClassIDStr) = True Then
						DelClassIDStr = Replace(DelClassIDStr,"," & Class_Num(Class_i),"")
					End If
				End If
				CheckTopClassRs.CLose : Set CheckTopClassRs = Nothing	
			Next
			Str_DelClassID = DelClassIDStr
			ArrC_Num = Split(Str_DelClassID,",")
			For ArrC_i = LBound(ArrC_Num) To Ubound(ArrC_Num)
				DelAllClassIDStr = DelAllClassIDStr & "," & GetAllChildRenClassID(ArrC_Num(ArrC_i))
			Next
		Else
			DelAllClassIDStr = GetAllChildRenClassID(DelClassIDStr)
		End If
		If Left(DelAllClassIDStr,1) = "," Then
			DelAllClassIDStr = Right(DelAllClassIDStr,Len(DelAllClassIDStr) - 1)
		End If
		If Instr(DelAllClassIDStr,",") > 0 Then
			Arr_Num = Split(DelAllClassIDStr,",")
			DelNum = Clng(Ubound(Arr_Num) + 1)
		Else
			Arr_Num = 0
			DelNum = 1
		End If		
		'---���Ϊ����ɾ��
		If NewsConfigTF = 0 Then
			On Error Resume Next
			'---����ɾ����ѡ��Ŀ�����ɵľ�̬�ļ�
			Set DleClassObj = Conn.ExeCute("Select SavePath,ClassEName,IsURL From Fs_Ns_NewsClass Where ClassID In(" & DelAllClassIDStr & ")")
			While Not DleClassObj.Eof
				If Cint(DleClassObj(2)) = 0 Then
					If DleClassObj(0) = "/" Then
						fso_DeleteFolder("/" & DleClassObj(1))
					Else
						fso_DeleteFolder(VRootStr & DleClassObj(0) & "/" & DleClassObj(1))
					End If
				End IF	
				DleClassObj.MoveNext		
			Wend
			'---�����ݿ���ɾ����ѡ��Ŀ
			Conn.ExeCute("Delete From Fs_Ns_NewsClass Where ClassID In(" & DelAllClassIDStr & ")")
			Conn.ExeCute("Delete From Fs_Ns_News Where ClassID In(" & DelAllClassIDStr & ")")
			'---ͬʱ����xml�ļ�
			If Arr_Num = 0 Then
				DelXmlStr = fso_DeleteFile("../../FS_InterFace/xml/NS/" & Replace(DelAllClassIDStr,"'","") & ".xml")
			Else
				For Arr_i = LBound(Arr_Num) To UBound(Arr_Num)
					DelXmlStr = fso_DeleteFile("../../FS_InterFace/xml/NS/" & Replace(Arr_Num(Arr_i),"'","") & ".xml")
				Next
			End If	
			Call Makexml("0")
			If DelXmlStr = True Then
				DelXmlStr = "ͬʱ������xml�ļ�"
			Else
				DelXmlStr = "ͬʱ������xml�ļ���״̬�����²��ɹ�.����������Ŀ¼��֧��д��Ȩ��"
			End If
			If Error.Number = 0 Then
				'---������д��ϵͳ������־
				Call MF_Insert_oper_Log("ɾ����Ŀ","ɾ����ĿClassID:"& Replace(DelAllClassIDStr,"'","") &"��������Ŀ,ͬʱɾ���˴���Ŀ�µ�������Ϣ,����ɾ��",now,session("admin_name"),"NS")
				'---ɾ���ɹ����ؽ��
				strShowErr = "<li><font color=""#ff0000"">ɾ���ɹ�</font></li><li>��ѡ��Ŀ(��<strong style=""color:#FF0000"">" & DelNum & "</strong>��)�ѳ���ɾ��</li><li>" & DelXmlStr & "</li>"
				Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Class_Manage.asp")
				Response.end
			Else
				Error.Clear
				'---ɾ��ʧ�ܷ��ؽ��
				strShowErr = "<li><font color=""#ff0000"">�������󣬲�����Ŀɾ��ʧ��</font></li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End If	
		Else  '---���ɾ��������վ
			On Error Resume Next
			Conn.ExeCute("Update Fs_Ns_NewsClass Set ReycleTF = 1 Where ClassID In(" & DelAllClassIDStr & ")")
			Conn.ExeCute("Update Fs_Ns_News Set ReycleTF = 1 Where ClassID In(" & DelAllClassIDStr & ")")
			'---ͬʱ����xml�ļ�
			If Arr_Num = 0 Then
				DelXmlStr = fso_DeleteFile("../../FS_InterFace/xml/NS/" & Replace(DelAllClassIDStr,"'","") & ".xml")
			Else
				For Arr_i = LBound(Arr_Num) To UBound(Arr_Num)
					DelXmlStr = fso_DeleteFile("../../FS_InterFace/xml/NS/" & Replace(Arr_Num(Arr_i),"'","") & ".xml")
				Next
			End If	
			Call Makexml("0")
			If DelXmlStr = True Then
				DelXmlStr = "ͬʱ������xml�ļ�"
			Else
				DelXmlStr = "ͬʱ������xml�ļ���״̬�����²��ɹ�.����������Ŀ¼��֧��д��Ȩ��"
			End If
			If Error.Number = 0 Then
				'---������д��ϵͳ������־
				Call MF_Insert_oper_Log("ɾ����Ŀ","ɾ����ĿClassID:"& Replace(DelAllClassIDStr,"'","") &"��������Ŀ,ͬʱɾ���˴���Ŀ�µ�������Ϣ,ɾ��������վ",now,session("admin_name"),"NS")
				'---ɾ���ɹ����ؽ��
				strShowErr = "<li><font color=""#ff0000"">ɾ���ɹ�</font></li><li>��ѡ��Ŀ(��<strong style=""color:#FF0000"">" & DelNum & "</strong>��)�ѷ������վ,���Դӻ���վ�лָ�</li><li>" & DelXmlStr & "</li>"
				Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Class_Manage.asp")
				Response.end
			Else
				Error.Clear
				'---ɾ��ʧ�ܷ��ؽ��
				strShowErr = "<li><font color=""#ff0000"">�������󣬲�����Ŀɾ��ʧ��</font></li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End If
		End If	
	End If
End If

Function InStrTopIDTF(ClassID,InStrID)
	Dim GetTopIDRs,StrID,ArrIDNum,Arr_i,InstrNum
	Set GetTopIDRs = Conn.ExeCute("Select ParentID,ClassID From FS_NS_NewsClass Where ClassID = '" & NoSqlHack(ClassID) & "' And ReycleTF = 0 Order By OrderID Desc,ID Desc")
	If Not GetTopIDRs.Eof Then
		Do While GetTopIDRs(0) <> "0"
			Set GetTopIDRs = Conn.ExeCute("Select ParentID,ClassID From FS_NS_NewsClass Where ClassID = '" & NoSqlHack(GetTopIDRs(0)) & "' And ReycleTF = 0 Order By OrderID Desc,ID Desc")
			StrID = StrID & "," & GetTopIDRs(1)
		Loop
	End If
	GetTopIDRs.Close : Set GetTopIDRs = Nothing
	StrID = Replace(Replace(StrID," ,",","),", ",",")
	If Left(StrID,1) = "," Then
		StrID = Right(StrID,Len(StrID) - 1)
	End If
	If Instr(StrID,",") > 0 Then
		ArrIDNum = Split(StrID,",")
		For Arr_i = LBound(ArrIDNum) To UBound(ArrIDNum)
			InstrNum = Clng(InstrNum + Instr(InStrID,ArrIDNum(Arr_i)))
		Next
	Else
		InstrNum = Instr(InStrID,StrID)
	End If	
	If InstrNum > 0 Then
		InStrTopIDTF = True
	Else
		InStrTopIDTF = False
	End If
End Function
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
    <td height="18" class="hback"><div align="left"><a href="Class_Manage.asp">������ҳ</a>��<a href="Class_add.asp?ClassID=&Action=add">��Ӹ���Ŀ</a>��<a href="Class_Action.asp?Action=one">һ����Ŀ����</a>��<a href="Class_Action.asp?Action=n">N����Ŀ����</a>��<a href="Class_Action.asp?Action=reset"   onClick="{if(confirm('ȷ�ϸ�λ������Ŀ��\n\n���ѡ��ȷ�������е���Ŀ������Ϊһ������!!')){return true;}return false;}">��λ������Ŀ</a>��<a href="Class_Action.asp?Action=unite">��Ŀ�ϲ�</a>��<a href="Class_Action.asp?Action=allmove">��Ŀת��</a>��<a href="Class_Action.asp?Action=clearClass"  onClick="{if(confirm('ȷ�����������Ŀ���������\n\n���ѡ��ȷ��,���е���Ŀ�����Ž����ŵ�����վ��!!')){return true;}return false;}">ɾ��������Ŀ</a> 
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
	obj_news_rs.Open "Select Orderid,id,ClassID,ParentID,ClassName from FS_NS_NewsClass where Parentid  = '0' Order by Orderid desc,ID desc",Conn,1,3
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
	obj_news_rs.Open "Select Orderid,id,ClassID,ParentID,ClassName from FS_NS_NewsClass where Parentid  = '0' Order by Orderid desc,ID desc",Conn,1,3
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
obj_unite_rs.Open "Select Orderid,id,ClassID,ClassName,ParentID from FS_NS_NewsClass where Parentid  = '0' and ReycleTF=0 Order by Orderid desc,ID desc",Conn,1,1
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
obj_unite_rs.Open "Select Orderid,id,ClassID,ClassName,ParentID from FS_NS_NewsClass where Parentid  = '0' and ReycleTF=0 Order by Orderid desc,ID desc",Conn,1,1
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
		if not Get_SubPop_TF(NoSqlHack(Request.QueryString("ClassID")),"NS024","NS","class") then Err_Show
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
		Conn.execute "update FS_NS_NewsClass set OrderID=" & NoSqlHack(OrderID) & " where ClassID='" & NoSqlHack(ClassID) &"'"
		Response.Redirect "Class_Action.asp?Action=one"
end sub
sub UpdateOrderIDN()
		if not Get_SubPop_TF(NoSqlHack(Request.QueryString("ClassID")),"NS024","NS","class") then Err_Show
		Dim N_ClassID,N_OrderID
		N_ClassID = NoSqlHack(Request.Form("ClassID"))
		N_OrderID = NoSqlHack(Request.Form("OrderID"))
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
		Conn.execute "update FS_NS_NewsClass set OrderID=" & N_OrderID & " where ClassID='" & N_ClassID &"'"
		Response.Redirect "Class_Action.asp?Action=n"
end sub
Sub delclass()
	Dim Rs_ClassDel,Str_SYSPath
	If G_VIRTUAL_ROOT_DIR<>"" Then
		Str_SYSPath="/"&G_VIRTUAL_ROOT_DIR
	Else
		Str_SYSPath=""
	End If
	if not Get_SubPop_TF(NoSqlHack(Request.QueryString("ClassID")),"NS021","NS","class") then
        Err_Show
	Else
	    if not Get_SubPop_TF(NoSqlHack(Request.QueryString("ClassID")),"NS003","NS","news") then Err_Show
	End if

	Dim str_delClassID,str_tmp_DelclassTF,Str_Tree_ClassID
	str_delClassID = NoSqlHack(Request.QueryString("ClassID"))
	if str_delClassID="" then
		strShowErr = "<li>�������:û����ĿClassID</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	Str_Tree_ClassID = GetAllChildRenClassID(str_delClassID )
	if fs_news.ReycleTF = 1 then
		Conn.execute("Update FS_NS_NewsClass set ReycleTF=1 Where ClassID In("&Str_Tree_ClassID&")")
		Conn.execute("Update FS_NS_News set isRecyle =1 Where ClassID In("&Str_Tree_ClassID&")")
		str_tmp_DelclassTF = 1
	Else
		Set Rs_ClassDel=Conn.Execute("Select SavePath,ClassEName,IsURL FROM FS_NS_NewsClass Where ClassID In("&FormatStrArr(Str_Tree_ClassID)&")")
		While Not Rs_ClassDel.Eof
			If Cint(Rs_ClassDel(2)) = 0 Then
				If Rs_ClassDel(0)="/" Then
					fso_DeleteFolder("/"&Rs_ClassDel(1))
				Else
					fso_DeleteFolder(Str_SYSPath&Rs_ClassDel(0)&"/"&Rs_ClassDel(1))
				End IF
			End If	
			Rs_ClassDel.movenext
		Wend
		Conn.execute("Delete From FS_NS_NewsClass Where ClassID In("&FormatStrArr(Str_Tree_ClassID)&")")
		Conn.execute("Delete From FS_NS_News Where ClassID In("&FormatStrArr(Str_Tree_ClassID)&")")
		str_tmp_DelclassTF=0
	End if
	dim returnvalues
	returnvalues = fso_DeleteFile("../../FS_InterFace/xml/NS/"&str_delClassID&".xml")
	Call Makexml("0")
	if returnvalues = true then:returnvalues = "ͬʱ������xml�ļ�":else:returnvalues = "ͬʱ������xml�ļ���״̬�����²��ɹ�.����������Ŀ¼��֧��д��Ȩ��":end if
	If str_tmp_DelclassTF =1 then
		strShowErr = "<li>��Ŀɾ���Ѿ�ɾ��������վ</li><li>"& returnvalues &"</li>" 
		Call MF_Insert_oper_Log("ɾ����Ŀ","ɾ����ĿClassID:"& str_delClassID &"��������Ŀ,ͬʱɾ���˴���Ŀ�µ�������Ϣ,ɾ��������վ��",now,session("admin_name"),"NS")
	Else
		strShowErr = "<li>��Ŀ�Ѿ�����ɾ���ɹ�</li><li>"& returnvalues &"</li>"
		Call MF_Insert_oper_Log("ɾ����Ŀ","ɾ����ĿClassID:"& str_delClassID &"��������Ŀ,ͬʱɾ���˴���Ŀ�µ�������Ϣ,����ɾ��",now,session("admin_name"),"NS")
	End if
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Class_Manage.asp")
	Response.end
End Sub
Sub resetClass()
	if not Get_SubPop_TF(NoSqlHack(Request.QueryString("ClassID")),"NS018","NS","class") then Err_Show
	Conn.Execute("Update FS_NS_NewsClass Set ParentID ='0'")
	'���������־
		Call MF_Insert_oper_Log("��λ��Ŀ","��������Ŀ��λΪһ����Ŀ",now,session("admin_name"),"NS")
		strShowErr = "<li>��λ������Ŀ�ɹ�</li><li>���е���Ŀ�Ѿ�����Ϊһ������</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Class_Manage.asp")
		Response.end
End Sub
'�ϲ�Ŀ¼
Sub Saveunite()
		if not Get_SubPop_TF(NoSqlHack(Request.QueryString("ClassID")),"NS019","NS","class") then 
			Response.Redirect("lib/error.asp?ErrCodes=ȱ��Ȩ��&ErrorUrl=")
			Response.end
		End if
		Dim str_SourceClassID,str_TargetClassID
		str_SourceClassID = NoSqlHack(Trim(Request.Form("SourceClassID")))
		str_TargetClassID = NoSqlHack(Trim(Request.Form("TargetClassID")))
		if split(str_SourceClassID,",")(0) = split(str_TargetClassID,",")(0) then
			strShowErr = "<li>ԴĿ¼��Ŀ��Ŀ¼����һ��</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		Dim obj_Url_rs
		set obj_Url_rs = Conn.execute("select isUrl from FS_NS_NewsClass where ClassID='"& NoSqlHack(split(str_TargetClassID,",")(0)) &"'")
		if obj_Url_rs("isUrl") = 1 then
			strShowErr = "<li>���ܰ���Ŀ�ϲ����ⲿ��Ŀ!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		Conn.execute("Update FS_NS_News set ClassID='"& NoSqlHack(split(str_TargetClassID,",")(0))&"'  where ClassID ='"& NoSqlHack(split(str_SourceClassID,",")(0))&"'")
		'��������������ݿ�
		'Conn.execute("Delete From  FS_NS_NewsPop  where ClassID ='"& split(str_SourceClassID,",")(0)&"'")
		'ɾ��Դ��Ŀ��
		Conn.execute("Delete From FS_NS_NewsClass  where ClassID ='"& NoSqlHack(split(str_SourceClassID,",")(0))&"'")
		Dim ob_Tmp_rs
		set ob_Tmp_rs = Conn.execute("select ClassID,ParentID From FS_NS_NewsClass Where ParentID='"& NoSqlHack(split(str_SourceClassID,",")(0)) &"' order by id desc")
		if Not ob_Tmp_rs.eof then
			do while Not ob_Tmp_rs.eof 
				Conn.execute("Update FS_NS_NewsClass Set ParentID ='"&  NoSqlHack(split(str_SourceClassID,",")(1)) &"' where ClassID ='"& NoSqlHack(ob_Tmp_rs("ClassID"))&"'")
				ob_Tmp_rs.movenext
			Loop
		End if
		ob_Tmp_rs.close:set ob_Tmp_rs = nothing
		'������������
		'******************����
		Call MF_Insert_oper_Log("�ϲ���Ŀ","����ĿClassID:"&  split(str_SourceClassID)(0)&"�ϲ���ClassID:"&  split(str_TargetClassID)(0) &"��",now,session("admin_name"),"NS")
		strShowErr = "<li>�ϲ���Ŀ�ɹ�</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Class_Manage.asp")
		Response.end
End Sub
'�ƶ���Ŀ
Sub allmove_save()
		if not Get_SubPop_TF(NoSqlHack(Request.QueryString("ClassID")),"NS020","NS","class") then Err_Show
		Dim str_SourceClassID_move,str_TargetClassID_move
		str_SourceClassID_move = NoSqlHack(Trim(Request.Form("SourceClassID")))
		str_TargetClassID_move = NoSqlHack(Trim(Request.Form("TargetClassID")))
		if split(str_SourceClassID_move,",")(0) = split(str_TargetClassID_move,",")(0) then
			strShowErr = "<li>ԴĿ¼��Ŀ��Ŀ¼����һ��</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		if str_TargetClassID_move ="0000000000" then
			Conn.execute("Update FS_NS_NewsClass Set ParentID ='0' where ClassID ='"&  split(str_SourceClassID_move,",")(0)&"'")
		Else
			Dim obj_Url_rs_1
			set obj_Url_rs_1 = Conn.execute("select isUrl from FS_NS_NewsClass where ClassID='"& split(str_TargetClassID_move,",")(0) &"'")
			if obj_Url_rs_1("isUrl") = 1 then
				strShowErr = "<li>���ܰ���Ŀת���ⲿ��Ŀ!</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End if
			Conn.execute("Update FS_NS_NewsClass Set ParentID ='"&  split(str_TargetClassID_move,",")(0) &"' where ClassID ='"&  split(str_SourceClassID_move,",")(0)&"'")
			Dim ob_Tmp_rs_1
			set ob_Tmp_rs_1 = Conn.execute("select ClassID,ParentID From FS_NS_NewsClass Where ParentID='"& split(str_SourceClassID_move,",")(0) &"' order by id desc")
			if Not ob_Tmp_rs_1.eof then
				do while Not ob_Tmp_rs_1.eof 
					Conn.execute("Update FS_NS_NewsClass Set ParentID ='"&  split(str_SourceClassID_move,",")(1) &"' where ClassID ='"& ob_Tmp_rs_1("ClassID")&"'")
					ob_Tmp_rs_1.movenext
				Loop
			End if
			ob_Tmp_rs_1.close:set ob_Tmp_rs_1 = nothing
		End if
		'������������
		'******************����
		Call MF_Insert_oper_Log("ת����Ŀ","����ĿClassID:"& split(str_SourceClassID_move)(0) &"ת�Ƶ�ClassID:"& split(str_TargetClassID_move)(0)&"��",now,session("admin_name"),"NS")
		strShowErr = "<li>��Ŀת�Ƴɹ���</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Class_Manage.asp")
		Response.end
End Sub
Sub clearClass()
		if not Get_SubPop_TF(NoSqlHack(Request.QueryString("ClassID")),"NS021","NS","class") then Err_Show
		Conn.execute("Update FS_NS_NewsClass set ReycleTF =1")
		Conn.execute("Update FS_NS_News  set  isRecyle =1")
		strShowErr = "<li>������Ŀ�Ѿ��ŵ�����վ��</li><li>���е���Ŀ�������Ѿ��ŵ�����վ��</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Class_Manage.asp")
		Response.end
End Sub
Sub one_clear()
		if not Get_SubPop_TF(NoSqlHack(Request.QueryString("ClassID")),"NS023","NS","class") then 
			Response.Redirect("lib/error.asp?ErrCodes=ȱ��Ȩ��&ErrorUrl=")
			Response.end
		End if

		Conn.execute("Delete From FS_NS_News where ClassID = '"& NoSqlHack(Request.QueryString("ClassID"))&"'")
		Call MF_Insert_oper_Log("�������","�����Classid:"& Request.QueryString("ClassID") &"�µ���������",now,session("admin_name"),"NS")
		strShowErr = "<li>��Ŀ�µ���������ɹ�</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
End Sub
set Fs_news = nothing

'----2006-12--06 by ken------
Function GetAllChildRenClassID(ClassID)
If ClassID = "" Then Exit Function
Dim GetClassIDRs,AllClassID
Set GetClassIDRs = Conn.ExeCute("Select ClassID From FS_NS_NewsClass Where ParentID = '"&NoSqlHack(ClassID)&"' And ReycleTF = 0 Order By OrderID Desc,ID Desc") 
If GetClassIDRs.Eof Then
	AllClassID = "'" & NoSqlHack(ClassID) & "'"
Else
	AllClassID = ""
	Do While Not GetClassIDRs.Eof
		AllClassID = AllClassID & "," & GetAllChildRenClassID(GetClassIDRs(0))
		GetClassIDRs.MoveNext
	Loop
	AllClassID = "'" & ClassID & "'" & AllClassID
End If
GetAllChildRenClassID = AllClassID
GetClassIDRs.Close
Set GetClassIDRs = Nothing				
End Function
%>






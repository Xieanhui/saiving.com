<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("NS_Genal") then Err_Show
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,strShowErr,showMorePageGo_Type_
Dim Page,cPageNo
Call MF_Check_Pop_TF("NS_Class_000001")
int_RPP=30 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"			'βҳ
 
Dim OM_OP,OM_OP_Sql_Str,obj_OM_OP_Rs,OM_OP_ID,obj_OM_OP_Rs_EofFlag,obj_OM_OP_Up_Rs,OM_OP_IDFlag,ShowList
OM_OP=Request.QueryString("OM_OP")
obj_OM_OP_Rs_EofFlag=False
If OM_OP="" Or Isnull(OM_OP) Then
	OM_OP="OM_ShowList"
End If
OM_OP_ID=Request.QueryString("ID")
If IsNull(OM_OP_ID) or OM_OP_ID="" Then
	OM_OP_IDFlag=True
Else
	OM_OP_IDFlag=False
End If

If OM_OP="S_AddNew" Then
	Dim OM_Temp_Add_Type,OM_Temp_Add_Name,OM_Temp_Add_LinkUrl,OM_Temp_Add_Email,OM_Temp_Add_Explanation,Re_ID,Up_Action
	OM_Temp_Add_Type=CintStr(Request.Form("Sel_Type"))
	Up_Action=NoSqlHack(Request.QueryString("Up_Action"))
	OM_Temp_Add_Name=NoSqlHack(Replace(Replace(Request.Form("Txt_Name"),"<","&lt;"),">","&gt;"))
	OM_Temp_Add_LinkUrl=NoSqlHack(Replace(Replace(Request.Form("Txt_LinkUrl"),"<","&lt;"),">","&gt;"))
	OM_Temp_Add_Email=NoSqlHack(Replace(Replace(Request.Form("Txt_Email"),"<","&lt;"),">","&gt;"))
	OM_Temp_Add_Explanation=NoSqlHack(Replace(Replace(Request.Form("Txt_Explanation"),"<","&lt;"),">","&gt;"))

	If Up_Action="Submit" Then
		Conn.execute("Update FS_NS_General Set G_Type="&NoSqlHack(OM_Temp_Add_Type)&",G_Name='"&NoSqlHack(OM_Temp_Add_Name)&"',G_URL='"&NoSqlHack(OM_Temp_Add_LinkUrl)&"',G_Email='"&NoSqlHack(OM_Temp_Add_Email)&"',G_Content='"&NoSqlHack(OM_Temp_Add_Explanation)&"' where GID="&CintStr(OM_OP_ID)&"")
		strShowErr = "<li>�޸ĳɹ�</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
		Conn.Execute("insert into FS_NS_General(G_Type,G_Name,G_URL,G_Email,G_Content,isLock) values ("&OM_Temp_Add_Type&",'"&OM_Temp_Add_Name&"','"&OM_Temp_Add_LinkUrl&"','"&OM_Temp_Add_Email&"','"&OM_Temp_Add_Explanation&"',0)")	
		strShowErr = "<li>��ӳɹ�</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If
End If
If OM_OP="S_DelAll" Then
	Conn.execute("Delete From FS_NS_General")
	strShowErr = "<li>ɾ���ɹ�</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=\admin\news\other_manage.asp?OM_OP=OM_ShowList")
	Response.end
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu"><a href="#" class="sd"><strong>�������</strong></a></td>
  </tr>
  <tr> 
      <td width="100%" height="18" class="hback"><div align="left"> <a href="javascript:Sy();">������ҳ</a> 
		|<a href="javascript:KeyWord();"> �ؼ���</a> | <a href="javascript:Source();">��Դ</a> | 
		<a href="javascript:ZouZe();">����</a> | <a href="javascript:LinkUrl();">�ڲ�����</a> 
		| 
		<a href="javascript:AddOne();">���</a> | 
		<a href="javascript:DelAll();">ɾ��ȫ��</a>  </div></td>
  </tr>
</table>

<div align="center">
<%
		Dim OM_OP_Up_Type,OM_OP_Up_Name,OM_OP_Up_Email,OM_OP_Up_Content,OM_OP_Up_Url,Temp_Action
		If OM_OP_IDFlag=False Then
			Set obj_OM_OP_Up_Rs=Conn.Execute("SELECT G_Type,G_Name,G_URL,G_Email,G_Content from FS_NS_General where GID="&Cint(OM_OP_ID)&"")
			If Not obj_OM_OP_Up_Rs.Eof Then
				OM_OP_Up_Type=CintStr(obj_OM_OP_Up_Rs("G_Type"))
				OM_OP_Up_Name=obj_OM_OP_Up_Rs("G_Name")
				OM_OP_Up_Url=obj_OM_OP_Up_Rs("G_URL")
				OM_OP_Up_Email=obj_OM_OP_Up_Rs("G_Email")
				OM_OP_Up_Content=obj_OM_OP_Up_Rs("G_Content")
			End If
			Set obj_OM_OP_Up_Rs=Nothing
			Temp_Action="&Up_Action=Submit&ID="&OM_OP_ID
		End If
%>
<table width="98%" border="0" cellpadding="5" cellspacing="1" class="table" id="OM_AddNew" style="display:none">
  <tr class="hback"> 
    <td class="xingmu" colspan="2">���</td>
  </tr>
 <form name="form_OM_AddNew" id="form_OM_AddNew" method="post" action="other_manage.asp?OM_OP=S_AddNew<%=Temp_Action%>" onSubmit="javascript:return Check_Type();">
  <tr> 
      <td width="290" height="29" class="hback" align="right">����</td>
      <td width="652" height="29" class="hback">
		<select size="1" name="Sel_Type" onChange="javascript:SelectOpType(this.value);">
		<% If IsNull(OM_OP_ID)=True Or OM_OP_ID="" Then %>
		<option value="0" selected>��ѡ��</option>
		<% End If %>		
		<option value="1" <% If OM_OP_Up_Type=1 Then Response.write "selected" %>>�ؼ���</option>
		<option value="2" <% If OM_OP_Up_Type=2 Then Response.write "selected" %>>��Դ</option>
		<option value="3" <% If OM_OP_Up_Type=3 Then Response.write "selected" %>>����</option>
		<option value="4" <% If OM_OP_Up_Type=4 Then Response.write "selected" %>>�ڲ�����</option>
		</select></td>
  </tr>
	<tr id="Tr_Title" style="display:none">
      <td height="28" class="hback" align="right">����</td>
      <td height="28" class="hback">
		<input type="text" name="Txt_Name" size="50" maxlength="50" value="<%=OM_OP_Up_Name%>">(<font color=red size=2>*</font>)</td>
    </tr>
	<tr id="Tr_Url" style="display:none">
      <td height="-5" class="hback" align="right">���ӵ�ַ</td>
      <td height="-5" class="hback">
		<input type="text" name="Txt_LinkUrl" size="50" maxlength="100" value="<%=OM_OP_Up_Url%>"></td>
    </tr>   
  <tr id="Tr_Email" style="display:none">
      <td height="-5" class="hback" align="right">�����ʼ�</td>
      <td height="-5" class="hback">
		<input type="text" name="Txt_Email" size="50" maxlength="50" value="<%=OM_OP_Up_Email%>"></td>
    </tr>
	<tr id="Tr_Explanation" style="display:none"> 
      <td height="-5" class="hback" align="right">˵��</td>
      <td height="-5" class="hback">
		<textarea rows="4" name="Txt_Explanation" cols="50" maxlength="200"><%=OM_OP_Up_Content%></textarea></td>
    </tr>
	<tr id="Tr_SubMit" style="display:none">
      <td  height="-5" class="hback" align="center" colspan="2">
		<input type="submit" value=" �� �� " name="But_AddNew">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="reset" value=" �� �� " name="But_Clear">(ע��(<font color=red size=2>*</font>)Ϊ����)</td>
    </tr>
  </form>
</table>
<table width="98%" border="0" cellpadding="5" cellspacing="1" class="table" id="OM_ShowList" style="display:none">
  <tr class="hback"> 
    <td class="xingmu" width="20%" align="center">����</td>
    <td class="xingmu" width="8%" align="center">����</td>
    <td class="xingmu" width="21%" align="center">���ӵ�ַ</td>
    <td class="xingmu" width="20%" align="center">�����ʼ�</td>
    <td class="xingmu" width="5%" align="center">״̬</td>
    <td class="xingmu" width="26%" align="center">����</td>
  </tr>
  <form name="form_OM_OP" id="form_OM_OP" method="post" action="other_manage.asp?OM_OP=OM_ShowList&OM_OP_P=P"><input name="Hi_OP_Type" type="hidden" value="">
<%
	If OM_OP="OM_ShowList" Then
		Dim OM_OP_Action,OM_OP_P
		OM_OP_Action=Request.QueryString("Action")
		OM_OP_P=Request.QueryString("OM_OP_P")
		If OM_OP_Action="Lock" Then
			Conn.execute("Update FS_NS_General Set isLock=1 Where GID="&CintStr(OM_OP_ID)&"")
		'Response.Write("<script language=javascript>alert('"&Request.QueryString("ShowList")&"')<script>")
		'Response.End()

		End If
		If OM_OP_Action="UnLock" Then
			Conn.execute("Update FS_NS_General Set isLock=0 Where GID="&CintStr(OM_OP_ID)&"")
		End If
		'--------------by chen  �޸ĵ���ɾ��  ----2/6--------------------
		If OM_OP_Action="DelOne" Then
			Conn.execute("Delete From FS_NS_General Where GID="&CintStr(OM_OP_ID)&"")
			strShowErr = "<li>ɾ���ɹ�</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
			response.Write(OM_OP_P)
			'response.end

		If OM_OP_P="P" Then
			Dim OM_OP_Temp_P_Type,OM_OP_Temp_P_ID,OM_OP_IOM_OP_Temp_P_Flag,OM_OP_IOM_OP_Temp_P_I,i,int_tempID
			OM_OP_Temp_P_Type=NoSqlHack(Request.Form("hi_OP_Type"))
			OM_OP_Temp_P_ID=NoSqlHack(Request.Form("Che_OPType"))
			If OM_OP_Temp_P_ID="" or IsNull(OM_OP_Temp_P_ID) Then
				strShowErr = "<li>��ѡ��Ҫ��������������!</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End If
			'int_tempID=Split(OM_OP_Temp_P_ID,",")
'			For i=1 To Ubound(OM_OP_Temp_P_ID)
'				int_tempID=int_tempID&OM_OP_Temp_P_ID(i)
'			Next
			'Response.Write(int_tempID)
			'Response.End()
			Select Case OM_OP_Temp_P_Type
				Case "P_Lock"
					Conn.execute("Update FS_NS_General Set isLock=1 Where GID in("&FormatIntArr(OM_OP_Temp_P_ID)&")")
					strShowErr = "<li>���������ɹ�</li>"
					Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
				Case "P_UnLock"
					Conn.execute("Update FS_NS_General Set isLock=0 Where GID in("&FormatIntArr(OM_OP_Temp_P_ID)&")")
					strShowErr = "<li>���������ɹ�</li>"
					Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
				'-----------by  chen �޸�����ɾ�� ---2/6-------------------
				Case "P_Del"
					Conn.execute("Delete From FS_NS_General Where GID in("&FormatIntArr(OM_OP_Temp_P_ID)&")")
					strShowErr = "<li>����ɾ���ɹ�</li>"
					Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=\admin\news\other_manage.asp?OM_OP=OM_ShowList")
					Response.end
			End Select
		End If
		ShowList=Request.QueryString("ShowList")
		If ShowList="" or Isnull(ShowList) Then
			OM_OP_Sql_Str="Select GID,G_Type,G_Name,G_URL,G_Email,isLock from FS_NS_General Order By G_Type"
		Else
			Select Case Cstr(ShowList)
				Case "Keyword"
					OM_OP_Sql_Str="Select GID,G_Type,G_Name,G_URL,G_Email,isLock from FS_NS_General Where G_Type=1"
				Case "Source"
					OM_OP_Sql_Str="Select GID,G_Type,G_Name,G_URL,G_Email,isLock from FS_NS_General Where G_Type=2"
				Case "ZouZe"
					OM_OP_Sql_Str="Select GID,G_Type,G_Name,G_URL,G_Email,isLock from FS_NS_General Where G_Type=3"
				Case "LinkUrl"
					OM_OP_Sql_Str="Select GID,G_Type,G_Name,G_URL,G_Email,isLock from FS_NS_General Where G_Type=4"
			End Select
		End If
		Set obj_OM_OP_Rs=CreateObject(G_FS_RS)
		
		obj_OM_OP_Rs.Open OM_OP_Sql_Str,Conn,1,1			
		
		Dim OM_Temp_Type,OM_Temp_IsLock,OM_Temp_Url,OM_Temp_Email,OM_Temp_Name

		If Not obj_OM_OP_Rs.Eof Then
			obj_OM_OP_Rs_EofFlag=True
			
			obj_OM_OP_Rs.PageSize=int_RPP
			cPageNo=NoSqlHack(Request.QueryString("page"))
			If cPageNo="" Then cPageNo = 1
			If not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo>obj_OM_OP_Rs.PageCount Then cPageNo=obj_OM_OP_Rs.PageCount 
			If cPageNo<=0 Then cPageNo=1
			obj_OM_OP_Rs.AbsolutePage=cPageNo
			
			For int_Start=1 TO int_RPP  
%>
 	<tr>
<%
				OM_Temp_Name=obj_OM_OP_Rs("G_Name")
				OM_Temp_Type=Cint(obj_OM_OP_Rs("G_Type"))
				Select Case OM_Temp_Type
					Case 1
						OM_Temp_Type="�ؼ���"
					Case 2
						OM_Temp_Type="��Դ"
					Case 3
						OM_Temp_Type="����"
					Case 4
						OM_Temp_Type="�ڲ�����"
				End Select
				OM_Temp_IsLock=Cint(obj_OM_OP_Rs("isLock"))
				Select Case OM_Temp_IsLock
					Case 0
						OM_Temp_IsLock="����"
					Case 1
						OM_Temp_IsLock="<font color=""red"">����</font>"
				End Select 
				If obj_OM_OP_Rs("G_URL")="" or Isnull(obj_OM_OP_Rs("G_URL")) then
					OM_Temp_Url="-"
				Else
					OM_Temp_Url=obj_OM_OP_Rs("G_URL")
				End If
				If obj_OM_OP_Rs("G_Email")="" or Isnull(obj_OM_OP_Rs("G_Email")) Then
					OM_Temp_Email="-"
				Else
					OM_Temp_Email=obj_OM_OP_Rs("G_Email")
				End If
%>
    <td class="hback" align="center" height="20"><%=OM_Temp_Name%></td>
    <td class="hback" align="center"><%=OM_Temp_Type%></td>
    <td class="hback" align="center"><%=OM_Temp_Url%></td>
    <td class="hback" align="center"><%=OM_Temp_Email%></td>
    <td class="hback" align="center"><%=OM_Temp_IsLock%></td>
    <td class="hback" align="center"><a href="javascript:Update('<%=obj_OM_OP_Rs("GID")%>');">�޸�</a> | <a href="javascript:DelOne('<%=obj_OM_OP_Rs("GID")%>');">ɾ��</a> | <a href="javascript:Lock('<%=obj_OM_OP_Rs("GID")%>');">����</a> | <a href="javascript:UnLock('<%=obj_OM_OP_Rs("GID")%>');">����</a>  |<input type="checkbox" value="<%=obj_OM_OP_Rs("GID")%>" name="Che_OPType"></td>
  </tr>
<%		
				obj_OM_OP_Rs.MoveNext
				If obj_OM_OP_Rs.Eof or obj_OM_OP_Rs.Bof Then Exit For
			Next
			Response.Write "<tr><td class=""hback"" colspan=""6"" align=""left"">"&fPageCount(obj_OM_OP_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"&nbsp;&nbsp;<input type=""button"" value="" �������� "" name=""But_P_Lock"" onclick=""javascript:P_Lock();"">&nbsp;<input type=""button"" value="" �������� "" name=""But_P_UnLock"" onclick=""javascript:P_UnLock();"">&nbsp;<input type=""button"" value="" ����ɾ�� "" name=""But_P_Del"" onclick=""javascript:P_Del();"">&nbsp;ȫѡ<input type=""checkbox"" name=""Che_OPType"" onclick=""CheckAll();"" value=""-1""></td></tr>"
		Else
			Response.write"<table width=""98%"" border=0 align=center cellpadding=2 cellspacing=1 class=table><tr><td>��ǰû������!</td></tr></table>"
		End If
		obj_OM_OP_Rs.Close
		Set obj_OM_OP_Rs=Nothing
	End If
	Response.write "<script language=""javascript"">document.all.OM_ShowList.style.display=""none"";</script>"

%>
</form> 
</table>
</div>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
</body>
</html>
<script language="javascript">
function SelectTableShow(ShowType)
{
	switch(ShowType)
		{
			case 0://��ʾ���������ҳ
				document.all.OM_ShowList.style.display="";
				document.all.OM_AddNew.style.display="none";
				break;				
			case 1://��ʾ����������ҳ
				document.all.OM_ShowList.style.display="none";
				document.all.OM_AddNew.style.display="";
				break;
			case 2://��ʾ��������޸�ҳ
				break;
			//case 3://��ʾɾ����ʾ�Ի���
			//	if(confirm('�㽫ɾ��������Ϣ��\n��ȷ��ɾ����'))location='?OM_OP=S_DelAll';
			//	break;	
			case 4://��ʾ�б���Ϣ
				<%
					If obj_OM_OP_Rs_EofFlag=True Then
				%>
				document.all.OM_ShowList.style.display="";
				<%
					Else
				%>
				document.all.OM_ShowList.style.display="none";
				<%
					End If
				%>
				document.all.OM_AddNew.style.display="none";
				break;							
		}
}
function SelectOpType(OpType)
{

	switch(parseInt(OpType))
		{
			case 0://Ĭ��ȫ������ʾ
				document.all.Tr_Url.style.display="none";
				document.all.Tr_Title.style.display="none";
				document.all.Tr_Email.style.display="none";
				document.all.Tr_Explanation.style.display="none";
				document.all.Tr_SubMit.style.display="none";
				break;
			case 1://�ؼ���
				document.all.Tr_Url.style.display="none";
				document.all.Tr_Title.style.display="";
				document.all.Tr_Email.style.display="none";
				document.all.Tr_Explanation.style.display="";
				document.all.Tr_SubMit.style.display="";
				break;
			case 2://��Դ
				document.all.Tr_Url.style.display="";
				document.all.Tr_Title.style.display="";
				document.all.Tr_Email.style.display="none";
				document.all.Tr_Explanation.style.display="";
				document.all.Tr_SubMit.style.display="";
				break;
			case 3://����
				document.all.Tr_Url.style.display="none";
				document.all.Tr_Title.style.display="";
				document.all.Tr_Email.style.display="";
				document.all.Tr_Explanation.style.display="";
				document.all.Tr_SubMit.style.display="";
				document.form_OM_AddNew.Sel_Type.value="3";
				break;	
			case 4://�ڲ�����
				document.all.Tr_Url.style.display="";
				document.all.Tr_Title.style.display="";
				document.all.Tr_Email.style.display="none";
				document.all.Tr_Explanation.style.display="";
				document.all.Tr_SubMit.style.display="";
				break;	
		}
}
function Check_Type()
{
	if (document.form_OM_AddNew.Sel_Type.value=="0") 
		{
			alert("��ѡ�����ͣ�");
			document.form_OM_AddNew.Sel_Type.focus();
			return false;
		}
	if	(document.form_OM_AddNew.Txt_Name.value=="")
		{
			alert("����ӱ��⣡");
			document.form_OM_AddNew.Txt_Name.focus();
			return false;
		}
	if (document.form_OM_AddNew.Sel_Type.value=="3"&&document.form_OM_AddNew.Txt_Email.value=="")
		{
			alert("����������ʼ���ַ��");
			document.form_OM_AddNew.Txt_Email.focus();
			return false;	
		}
	if( document.form_OM_AddNew.Sel_Type.value=="3"&&document.form_OM_AddNew.Txt_Email.value.length<6 || document.form_OM_AddNew.Sel_Type.value=="3"&&document.form_OM_AddNew.Txt_Email.value.length>36 || document.form_OM_AddNew.Sel_Type.value=="3"&&!validateEmail() ) 
		{
		      alert("\����������ȷ�������ַ !");
		     document.form_OM_AddNew.Txt_Email.focus();
		     return false;	
	    }
	return true
}
//�������ʼ���ʽ
function validateEmail(){
	var re=/^[\w-]+(\.*[\w-]+)*@([0-9a-z]+[0-9a-z-]*[0-9a-z]+\.)+[a-z]{2,3}$/i;
	if(re.test(document.form_OM_AddNew.Txt_Email.value))
		return true;
	else
		return false;
}
function CheckAll()
{
	var checkBoxArray=document.all("Che_OPType")
	if(checkBoxArray[checkBoxArray.length-1].checked)
	{
		for(var i=0;i<checkBoxArray.length-1;i++)
		{
			checkBoxArray[i].checked=true;
			
		}
	}else
	{
		for(var i=0;i<checkBoxArray.length-1;i++)
		{
			checkBoxArray[i].checked=false;
		}
	}
}
function P_Lock()
{
	document.form_OM_OP.Hi_OP_Type.value="P_Lock";
	document.form_OM_OP.action='?OM_OP=OM_ShowList&OM_OP_P=P&ShowList=<%=NoSqlHack(Request.QueryString("Showlist"))%>';
	document.form_OM_OP.submit();
}
function P_UnLock()
{
	document.form_OM_OP.Hi_OP_Type.value="P_UnLock";
	document.form_OM_OP.action='?OM_OP=OM_ShowList&OM_OP_P=P&ShowList=<%=NoSqlHack(Request.QueryString("Showlist"))%>';
	document.form_OM_OP.submit();
}
function P_Del()
{
	if(confirm('�˲�����ɾ��ѡ�е����ݣ�\n��ȷ��ɾ����'))
	{
		document.form_OM_OP.Hi_OP_Type.value="P_Del";
		document.form_OM_OP.action='?OM_OP=OM_ShowList&OM_OP_P=P&ShowList=<%=NoSqlHack(Request.QueryString("Showlist"))%>';
		document.form_OM_OP.submit();
	}
}
function DelAll()
{
	if(confirm('��ɾ���������ݣ�\n��ȷ��ɾ����'))
	{
		location='?OM_OP=S_DelAll';
	}
}
function Lock(GID)
{
	location='?OM_OP=OM_ShowList&ID='+GID+'&Action=Lock&Page=<%=NoSqlHack(Request.QueryString("Page"))%>&ShowList=<%=NoSqlHack(Request.QueryString("ShowList"))%>';
}
function UnLock(GID)
{
	location='?OM_OP=OM_ShowList&ID='+GID+'&Action=UnLock&Page=<%=NoSqlHack(Request.QueryString("Page"))%>&ShowList=<%=NoSqlHack(Request.QueryString("ShowList"))%>';
}
function Update(GID)
{
	location='?OM_OP=AddNew&ID='+GID+'&Action=Update&Page=<%=NoSqlHack(Request.QueryString("Page"))%>&ShowList=<%=NoSqlHack(Request.QueryString("ShowList"))%>';
}
function DelOne(GID)
{
	if (confirm('ȷ��Ҫɾ���ü�¼��?'))
	{
		location='?OM_OP=OM_ShowList&ID='+GID+'&Action=DelOne&Page=<%=Request.QueryString("Page")%>&ShowList=<%=Request.QueryString("ShowList")%>';
	}
}
function KeyWord()
{
	location="other_manage.asp?OM_OP=OM_ShowList&ShowList=Keyword";
}
function Source()
{
	location="other_manage.asp?OM_OP=OM_ShowList&ShowList=Source";
}
function ZouZe()
{
	location='other_manage.asp?OM_OP=OM_ShowList&ShowList=ZouZe';
}
function LinkUrl()
{
	location='other_manage.asp?OM_OP=OM_ShowList&ShowList=LinkUrl';
}
function AddOne()
{
	location='other_manage.asp?OM_OP=AddNew';
}
function Sy()
{
	location='other_manage.asp?OM_OP=OM_ShowList';
}
</script>
<%
If OM_OP="" Then
	Response.write"<script language=""javascript"">SelectTableShow(0);</script>"
Else
	OM_OP=Cstr(OM_OP)
	Select Case OM_OP
		Case "AddNew"
			Response.write"<script language=""javascript"">SelectTableShow(1);</script>"
		Case "Update"
			Response.write"<script language=""javascript"">SelectTableShow(2);</script>"
		Case "DelAll"
			Response.write"<script language=""javascript"">SelectTableShow(3);</script>"
		Case "OM_ShowList"
			Response.write"<script language=""javascript"">SelectTableShow(4);</script>"
		Case else
			Response.write"<script language=""javascript"">SelectTableShow(0);</script>"
	End Select
End IF

If OM_OP_Up_Type=4 Then Response.write "<script language=""javascript"">SelectOpType(4);</script>" 
If OM_OP_Up_Type=1 Then Response.write "<script language=""javascript"">SelectOpType(1);</script>"
If OM_OP_Up_Type=2 Then Response.write "<script language=""javascript"">SelectOpType(2);</script>"
If OM_OP_Up_Type=3 Then Response.write "<script language=""javascript"">SelectOpType(3);</script>"
Set Conn=Nothing
%><!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






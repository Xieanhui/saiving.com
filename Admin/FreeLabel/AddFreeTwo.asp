<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="FieldsArr.asp" -->
<%
'========================================================
'ϵͳ����
'========================================================
Dim Conn,strShowErr
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
MF_Default_Conn
MF_Session_TF 
If Not MF_Check_Pop_TF("MF_sPublic") Then Err_Show
Dim sRootDir,str_CurrPath
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if
'====================================================
'��ȡ�ӵ�һ���ύ������
'====================================================
Dim LabelName,Label_DesStr,LabelID,ActStr,LabelType,NS_TableName,NC_TableName,Select_num
Dim Sql_ConStr,Ns_FieldsStr,Nc_fieldsStr,Str_Lable_ConStr,Lable_Sql_Str
Dim Auto_Fields_Str,Auto_Str_N,Auto_Str_C,LNameStr,CheckNameRs,Auto_Str_DA
Dim Dis_AllFields
Dim Lable_Content
Dim InfoTitle,InfoType,InfoContent,ReturnUrl
Dim AddRs,AddSql
ActStr = Request.Form("To_Act")
If ActStr = "Add" Then
	LabelID = ""
	Label_DesStr = ""
	Str_Lable_ConStr = ""
ElseIf ActStr = "Edit" Then
	LabelID = Request.Form("LabelID")
	Label_DesStr = Request.Form("Label_Des")
	Str_Lable_ConStr = Request.Form("Lable_ConStr")
End If	
LabelName = "FS400_" & Request.Form("LabelName")
LabelType = Request.Form("SysType")
NS_TableName = Request.Form("NTable")
NC_TableName = Request.Form("CTable")
Select_num = Request.Form("SelectNum")
Sql_ConStr = ReplaceMidFildes(Request.Form("DisSql"))
Ns_FieldsStr =ReplaceMidFildes(Request.Form("Fist_TF_All"))
Nc_fieldsStr = ReplaceMidFildes(Request.Form("Sec_TF_All"))


If LabelType = "NS" Then 
	Auto_Str_N = "<span style=""cursor:hand;"" onclick=""InsertToHTMl('[#NewsUrl#]')"">[���������ַ]</span>"
	Auto_Str_C = "<span style=""cursor:hand;"" onclick=""InsertToHTMl('[#NewsClassUrl#]')"">[��Ŀ�����ַ]</span>"
	Auto_Str_DA = ""		
ElseIf LabelType = "DS" Then
	Auto_Str_N = "<span style=""cursor:hand;"" onclick=""InsertToHTMl('[#DownUrl#]')"">[���������ַ]</span>"
	Auto_Str_C = "<span style=""cursor:hand;"" onclick=""InsertToHTMl('[#DownClassUrl#]')"">[��Ŀ�����ַ]</span>"
	Auto_Str_DA = "<span style=""cursor:hand;"" onclick=""InsertToHTMl('[#DownAddress#]')"">[������ص�ַ]</span>"
ElseIf LabelType = "MS" Then
	Auto_Str_N = "<span style=""cursor:hand;"" onclick=""InsertToHTMl('[#MallUrl#]')"">[��Ʒ�����ַ]</span>"
	Auto_Str_C = "<span style=""cursor:hand;"" onclick=""InsertToHTMl('[#MallClassUrl#]')"">[��Ŀ�����ַ]</span>"
	Auto_Str_DA = ""
End If
If Ns_FieldsStr <> "" Then
	If LabelType = "DS" Then
		Auto_Fields_Str = Auto_Str_N & " �� " & Auto_Str_DA
	Else
		Auto_Fields_Str = Auto_Str_N	
	End If	
Else
	Auto_Fields_Str = ""
End If
If Nc_fieldsStr <> "" Then
	If Auto_Fields_Str <> "" Then
		Auto_Fields_Str = Auto_Fields_Str & " �� " & Auto_Str_C
	Else
		Auto_Fields_Str = Auto_Str_C
	End If		
Else
	Auto_Fields_Str = Auto_Fields_Str & ""
End If

If Ns_FieldsStr <> "" And Nc_fieldsStr <> "" Then
	Dis_AllFields = GetLeftSeFields(Ns_FieldsStr,"And",LabelType) & " �� " & GetRightSeFields(Nc_fieldsStr,"And",LabelType)
ElseIf Ns_FieldsStr <> "" And Nc_fieldsStr = "" Then
	Dis_AllFields = GetLeftSeFields(Ns_FieldsStr,"No",LabelType)
ElseIf Ns_FieldsStr = "" And Nc_fieldsStr <> "" Then
	Dis_AllFields =  GetRightSeFields(Nc_fieldsStr,"No",LabelType)
Else
	Dis_AllFields = ""	
End If


'====================================================
'��������mid����substring���ֶ��е�,�滻��*
'====================================================
Function ReplaceMidFildes(SqlStr)
	Dim Str_Sql,FlagSql,FormBack,InstrMidStr,ReplaceMidStr,LeftStr
	Dim Mid_arr,Mid_i
	Str_Sql = SqlStr & ""
	IF Instr(Str_Sql,"From") > 0 Then
		FlagSql = Trim(Split(Str_Sql,"From")(0))
		FormBack = " From " & Trim(Split(Str_Sql,"From")(1))
	Else
		FlagSql = Trim(Str_Sql)
		FormBack = ""
	End If
	IF Instr(FlagSql,GetStrFun) <> 0 Then
		Mid_arr = Split(FlagSql,GetStrFun)
		For Mid_i = 1 To UBound(Mid_arr)
			InstrMidStr = Trim(Mid_arr(Mid_i))
			LeftStr = Trim(Split(InstrMidStr,")")(0))
			ReplaceMidStr = Replace(LeftStr,",","*")
			FlagSql = Replace(FlagSql,LeftStr,ReplaceMidStr)
		Next
		ReplaceMidFildes = FlagSql & FormBack
	Else
		ReplaceMidFildes = Str_Sql
	End IF		
End Function


'====================================================
'���ݻ�ȡ��ֵ���õ����ű�����ֶε������б�
'====================================================
Function GetLeftSeFields(FidldsStr,StrType,TableType)
	Dim TableName,i,Arr,Fname,FNameStr,FNum,FCName
	If FidldsStr = "" Then 
		GetLeftSeFields = ""
		Exit Function
	End If
	GetLeftSeFields = "<select name=""InsertToConN"" id=""InsertToConN"" onChange=""InsertToHTMl(this.options[this.selectedIndex].value)"">" & vbnewline
	GetLeftSeFields = GetLeftSeFields & "<option value="""">��ѡ���ֶ�</option>" & vbnewline
	If Instr(FidldsStr,",") > 0 Then
		Arr = Split(FidldsStr,",")
		For i = LBound(Arr) To UBound(Arr)
			FNameStr = Replace(Replace(Arr(i),"*",","),NS_TableName & ".","")
			If TableType = "NS" Then
				FNum = GetInnerFieldsNum(FNameStr,NSAllFENArr)
				FCName = NSAllFCNArr(FNum)
				TableName = "����"
			ElseIf TableType = "DS" Then
				FNum = GetInnerFieldsNum(FNameStr,DSAllFENArr)
				FCName = DSAllFCNArr(FNum)
				TableName = "����"
			ElseIF TableType = "MS" Then
				FNum = GetInnerFieldsNum(FNameStr,MSAllFENArr)
				FCName = MSAllFCNArr(FNum)
				TableName = "��Ʒ"
			End IF
			If StrType = "And" Then 
				Fname = Replace(Arr(i),"*",",")
				FCName = TableName & "." & FCName
			ElseIF StrType = "No" Then 	
				Fname = Replace(Replace(Arr(i),"*",","),NS_TableName & ".","")
				FCName = FCName
			End If
			GetLeftSeFields = GetLeftSeFields & "<option value=""[*" & Fname & "*]"">" & FCName & "</option>" & vbnewline
		Next
	Else
		FNameStr = Replace(Replace(FidldsStr,"*",","),NS_TableName & ".","")
		If TableType = "NS" Then
			FNum = GetInnerFieldsNum(FNameStr,NSAllFENArr)
			FCName = NSAllFCNArr(FNum)
			TableName = "����"
		ElseIf TableType = "DS" Then
			FNum = GetInnerFieldsNum(FNameStr,DSAllFENArr)
			FCName = DSAllFCNArr(FNum)
			TableName = "����"
		ElseIF TableType = "MS" Then
			FNum = GetInnerFieldsNum(FNameStr,MSAllFENArr)
			FCName = MSAllFCNArr(FNum)
			TableName = "��Ʒ"
		End IF
		If StrType = "And" Then 
			Fname = Replace(FidldsStr,"*",",")
			FCName = TableName & "." & FCName
		ElseIF StrType = "No" Then 	
			Fname = Replace(Replace(FidldsStr,"*",","),NS_TableName & ".","")
			FCName = FCName
		End If
		GetLeftSeFields = GetLeftSeFields & "<option value=""[*" & Fname & "*]"">" & FCName & "</option>" & vbnewline
	End If
	GetLeftSeFields = GetLeftSeFields & "</select>" & vbnewline	
End Function


'====================================================
'���ݻ�ȡ��ֵ���õ���Ŀ������ֶε������б�
'====================================================
Function GetRightSeFields(FidldsStr,StrType,TableType)
	Dim TableName,i,Arr,Fname,FNameStr,FNum,FCName
	If FidldsStr = "" Then 
		GetRightSeFields = ""
		Exit Function
	End If
	GetRightSeFields = "<select name=""InsertToConC"" id=""InsertToConC"" onChange=""InsertToHTMl(this.options[this.selectedIndex].value)"">" & vbnewline
	GetRightSeFields = GetRightSeFields & "<option value="""">��ѡ���ֶ�</option>" & vbnewline
	If Instr(FidldsStr,",") > 0 Then
		Arr = Split(FidldsStr,",")
		For i = LBound(Arr) To UBound(Arr)
			FNameStr = Replace(Replace(Arr(i),"*",","),NC_TableName & ".","")
			If TableType = "NS" Then
				FNum = GetInnerFieldsNum(FNameStr,NS_CAllENArr)
				FCName = NS_CAllFCNArr(FNum)
				TableName = "������Ŀ"
			ElseIf TableType = "DS" Then
				FNum = GetInnerFieldsNum(FNameStr,DCAllFENArr)
				FCName = DCAllFCNArr(FNum)
				TableName = "������Ŀ"
			ElseIF TableType = "MS" Then
				FNum = GetInnerFieldsNum(FNameStr,MCAllFENArr)
				FCName = MCAllFCNArr(FNum)
				TableName = "��Ʒ��Ŀ"
			End IF
			If StrType = "And" Then 
				Fname = Replace(Arr(i),"*",",")
				FCName = TableName & "." & FCName
			ElseIF StrType = "No" Then 	
				Fname = Replace(Replace(Arr(i),"*",","),NC_TableName & ".","")
				FCName = FCName
			End If
			GetRightSeFields = GetRightSeFields & "<option value=""[*" & Fname & "*]"">" & FCName & "</option>" & vbnewline
		Next
	Else
		FNameStr = Replace(Replace(FidldsStr,"*",","),NC_TableName & ".","")
		If TableType = "NS" Then
			FNum = GetInnerFieldsNum(FNameStr,NS_CAllENArr)
			FCName = NS_CAllFCNArr(FNum)
			TableName = "������Ŀ"
		ElseIf TableType = "DS" Then
			FNum = GetInnerFieldsNum(FNameStr,DCAllFENArr)
			FCName = DCAllFCNArr(FNum)
			TableName = "������Ŀ"
		ElseIF TableType = "MS" Then
			FNum = GetInnerFieldsNum(FNameStr,MCAllFENArr)
			FCName = MCAllFCNArr(FNum)
			TableName = "��Ʒ��Ŀ"
		End IF
		If StrType = "And" Then 
			Fname = Replace(FidldsStr,"*",",")
			FCName = TableName & "." & FCName
		ElseIF StrType = "No" Then 	
			Fname = Replace(Replace(FidldsStr,"*",","),NC_TableName & ".","")
			FCName = FCName
		End If
		GetRightSeFields = GetRightSeFields & "<option value=""[*" & Fname & "*]"">" & FCName & "</option>" & vbnewline
	End If
	GetRightSeFields = GetRightSeFields & "</select>" & vbnewline	
End Function


'====================================================
'��ҳ���ύ
'====================================================
If Request.Form("Action") = "submit" Then
	Label_DesStr = Request.Form("Label_Des")
	Lable_Content = Request.Form("Style_Txt")
	Lable_Sql_Str = Request.Form("Sql_ConStr")
	LabelID = Request.Form("LabelID")
	LNameStr = Trim(Request.Form("LableName"))
	If Lable_Sql_Str = "" Then
		InfoTitle = Server.URLEncode("������")
		InfoType = Server.URLEncode("ER")
		InfoContent = Server.URLEncode("<li>��ѯSQL��䲻��Ϊ��</li>")
		ReturnUrl = Server.URLEncode("")
		Response.Redirect "ShowInfo.asp?Str_T=" & InfoTitle & "&Str_C=" & InfoType & "&Con_Str=" & InfoContent & "&Str_U=" & ReturnUrl & ""
		Response.End
	End If		
	If Len(Label_DesStr) > 100 Then
		InfoTitle = Server.URLEncode("������")
		InfoType = Server.URLEncode("ER")
		InfoContent = Server.URLEncode("<li>��ǩ����̫���ˣ����ܳ���100���ַ�</li>")
		ReturnUrl = Server.URLEncode("")
		Response.Redirect "ShowInfo.asp?Str_T=" & InfoTitle & "&Str_C=" & InfoType & "&Con_Str=" & InfoContent & "&Str_U=" & ReturnUrl & ""
		Response.End
	End If
	If Lable_Content <> "" Then
		Lable_Content = Replace(Lable_Content,"'","")
	Else
		InfoTitle = Server.URLEncode("������")
		InfoType = Server.URLEncode("ER")
		InfoContent = Server.URLEncode("<li>��ǩ���ݲ���Ϊ��</li>")
		ReturnUrl = Server.URLEncode("")
		Response.Redirect "ShowInfo.asp?Str_T=" & InfoTitle & "&Str_C=" & InfoType & "&Con_Str=" & InfoContent & "&Str_U=" & ReturnUrl & ""
		Response.End
	End If		
	On Error Resume Next
	Set AddRs = Server.CreateObject(G_FS_RS)
	If LabelID <> "" And Len(LabelID) = 15 Then
		Set CheckNameRs = Conn.ExeCute("Select LabelName From FS_MF_FreeLabel Where LabelName = '" & NoSqlHack(LNameStr) & "' And LabelID <> '" & NoSqlHack(LabelID) & "'")
		IF Not CheckNameRs.Eof Then
			InfoTitle = Server.URLEncode("������")
			InfoType = Server.URLEncode("ER")
			InfoContent = Server.URLEncode("<li>��ǩ���Ʋ����ظ�</li>")
			ReturnUrl = Server.URLEncode("")
			Response.Redirect "ShowInfo.asp?Str_T=" & InfoTitle & "&Str_C=" & InfoType & "&Con_Str=" & InfoContent & "&Str_U=" & ReturnUrl & ""
			Response.End
		End If
		CheckNameRs.Close : Set CheckNameRs = Nothing
		AddSql = "Select LabelID,LabelName,LabelSQl,NSFields,NCFields,LabelContent,selectNum,DesCon,SysType From FS_MF_FreeLabel Where LabelID = '" & LabelID & "'"
		AddRs.Open AddSql,Conn,1,3
	Else
		Set CheckNameRs = Conn.ExeCute("Select LabelName From FS_MF_FreeLabel Where LabelName = '" & NoSqlHack(LNameStr) & "'")
		IF Not CheckNameRs.Eof Then
			InfoTitle = Server.URLEncode("������")
			InfoType = Server.URLEncode("ER")
			InfoContent = Server.URLEncode("<li>��ǩ���Ʋ����ظ�</li>")
			ReturnUrl = Server.URLEncode("")
			Response.Redirect "ShowInfo.asp?Str_T=" & InfoTitle & "&Str_C=" & InfoType & "&Con_Str=" & InfoContent & "&Str_U=" & ReturnUrl & ""
			Response.End
		End If
		CheckNameRs.Close : Set CheckNameRs = Nothing
		AddSql = "Select LabelID,LabelName,LabelSQl,NSFields,NCFields,LabelContent,selectNum,DesCon,SysType From FS_MF_FreeLabel Where 1=2"
		AddRs.Open AddSql,Conn,1,3
		AddRs.AddNew
		AddRs(0) = GetRamCode(15)
	End If
	AddRs(1) = LNameStr
	AddRs(2) = Lable_Sql_Str
	AddRs(3) = Request.Form("Ns_FieldsStr")
	AddRs(4) = Request.Form("Nc_fieldsStr")
	AddRs(5) = Lable_Content
	AddRs(6) = Request.Form("Select_num")
	AddRs(7) = Label_DesStr
	AddRs(8) = Request.Form("LabelType")
	AddRs.Update
	AddRs.Close : Set AddRs = Nothing
	If Err.Number <> 0 Then
		InfoTitle = Server.URLEncode("������")
		InfoType = Server.URLEncode("ER")
		InfoContent = Server.URLEncode("<li>" & Err.Description & "</li>")
		ReturnUrl = Server.URLEncode("")
		Response.Redirect "ShowInfo.asp?Str_T=" & InfoTitle & "&Str_C=" & InfoType & "&Con_Str=" & InfoContent & "&Str_U=" & ReturnUrl & ""
		Response.End
	Else
		InfoTitle = Server.URLEncode("�����ɹ�")
		InfoType = Server.URLEncode("OK")
		InfoContent = Server.URLEncode("")
		ReturnUrl = Server.URLEncode("FreeLabelList.asp")
		Response.Redirect "ShowInfo.asp?Str_T=" & InfoTitle & "&Str_C=" & InfoType & "&Con_Str=" & InfoContent & "&Str_U=" & ReturnUrl & ""
		Response.End
	End If		
End If
%>
<html>
<head>
<title>���ɱ�ǩ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" src="../../FS_Inc/Prototype.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Get_Domain.asp"></script>
<body>
<table width="98%" height="40" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="hback" >
    <td width="100%" height="20"  align="Left" class="xingmu" valign="middle">�������ɱ�ǩ</td>
  </tr>
  <tr class="hback" >
    <td height="20" align="center" class="hback" valign="middle"><div style="line-height:20px; text-align:left;"><span onClick="javascript:history.back();" style="cursor:hand;">��һ��</span>��<span onClick="SubmitFun()" style="cursor:hand;">����</span></div></td>
  </tr>
</table>
<form name="List_Form" id="List_Form" action="" method="post" style="margin:0px;">
<input name="Action" id="Action" type="hidden" value="submit">
<table width="98%" height="50" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="hback">
    <td width="10%" height="20" align="right" valign="middle">��ǩ���ƣ�</td>
    <td height="20" align="left" valign="middle">
	<input name="LableName" id="LableName" type="text" readonly value="<% = LabelName %>">
	<input name="LabelType" id="LabelType" type="hidden" value="<% = LabelType %>">
	<input name="Select_num" id="Select_num" type="hidden" value="<% = Select_num %>">
	<input name="Nc_fieldsStr" id="Nc_fieldsStr" type="hidden" value="<% = Nc_fieldsStr %>">
	<input name="Ns_FieldsStr" id="Ns_FieldsStr" type="hidden" value="<% = Ns_FieldsStr %>">
	<input name="LabelID" id="LabelID" type="hidden" value="<% = LabelID %>">
	<textarea name="Sql_ConStr" id="Sql_ConStr" style="display:none" ><% = Sql_ConStr %></textarea>
	</td>
  </tr>
  <tr class="hback">
    <td width="10%" height="20" align="right" valign="middle">��ǩ˵����</td>
    <td height="20" align="left" valign="middle">
		<input name="Label_Des" id="Label_Des" type="text" style="width:40%" value="<% = Label_DesStr %>">
		<span class="tx" style="margin-left:20px;">��ǩ˵�����֣�100������</span>
	</td>
  </tr>
  <tr class="hback">
    <td width="10%" height="20" align="right" valign="middle">Ԥ�����ֶΣ�</td>
    <td height="20" align="left" valign="middle">
	<% = Auto_Fields_Str %>
	<span class="tx" style="margin-left:20px;">��˵�� 1</span>
	</td>
  </tr>
  <tr class="hback">
    <td width="10%" height="20" align="right" valign="middle">�����ֶΣ�</td>
    <td height="20" align="left" valign="middle">
	<% = Dis_AllFields %>
	<span class="tx" style="margin-left:20px;">��Ҫ��һ��ѡ���ֶΡ�</span>
	</td>
  </tr>
  <tr class="hback">
    <td width="10%" height="20" align="right" valign="middle">������ʽ��</td>
    <td height="20" align="left" valign="middle">
	<input name="Date_Style" id="Date_Style" type="text" style="width:20%" value="YY02��MM��DD��">
	<input type="button" value="����" name="InsertTime" id="InsertTime" onClick="InsertTimeToHTML()">
	<span class="tx" style="margin-left:20px;">��Ҫѡ��ʱ���ֶΣ���ʽ��˵�� 2</span>
	</td>
  </tr>
 <tr class="hback">
    <td width="10%" height="20" align="right" valign="middle">��ǩ���ݣ�</td>
    <td height="20" align="left" valign="middle">
	<span class="tx">��HTML�������ѡ���ֶΡ��Զ��庯����ɣ����������ѯ��¼����ʾ��ʽ</span>
	</td>
  </tr>
  <tr class="hback">
    <td colspan="2" align="center" valign="middle" height="200">
		<!--�༭����ʼ-->
		<iframe id='NewsContent' src='../Editer/AdminEditer.asp?id=Style_Txt' frameborder=0 scrolling=no width='100%' height='440'></iframe>
        <textarea name="Style_Txt" rows="15" id="Style_Txt" style="width:100%;display:none;" ><% = HandleEditorContent(Str_Lable_ConStr) %></textarea>
        <!--�༭������-->	</td>
  </tr>
</table>
</form>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="hback" >
    <td width="100%" height="20"  align="Left" valign="middle"><span class="tx">˵����</span></td>
  </tr>
 <tr class="hback" >
    <td width="100%" height="20"  align="Left" valign="middle"><span class="tx">1.Ԥ�����ֶ���Ҫѡ����Զ�Ӧ��š����������·����Ҫѡ�����ű�ţ���Ŀ���·����Ҫѡ����Ŀ���(ע�⣺�����ű�ţ����Ǳ��)��</span></td>
  </tr>
  <tr class="hback" >
    <td width="100%" height="20"  align="Left" valign="middle"><span class="tx">2.���ڸ�ʽ:YY02����2λ�����(��06��ʾ2006��),YY04��ʾ4λ�������(2006)��MM�����£�DD�����գ�HH����Сʱ��MI����֣�SS�����롣</span></td>
  </tr>
  <tr class="hback" >
    <td width="100%" height="20"  align="Left" valign="middle"><span class="tx">3.�Զ��庯����ѭ������{#...#}����ѭ������{*n...*}(n>0)�����¼��š�����(#...#)����(#Left([*FS_News.Title*],20)#)</span></td>
  </tr>
</table>
</body>
</html>

<script language="javascript" type="text/javascript">
//�ֶ�ֵ����༭��
function InsertToHTMl(str)
{
	InsertHTML(str,"NewsContent"); 
}

//�ֶ�ֵ����༭��
function InsertTimeToHTML()
{
	var Str = '';
	var time_Str = $('Date_Style').value;
	if (time_Str == '')
	{
		return;
	}
	Str = '[$' + time_Str + '$]';
	InsertToHTMl(Str)
}

//�ύ��
function SubmitFun()
{
	document.List_Form.Style_Txt.value=frames["NewsContent"].GetNewsContentArray()
	if (document.List_Form.Style_Txt.value == '')
	{
		alert('��ǩ���ݲ���Ϊ��');
		return;
	}
	document.List_Form.submit();
}
</script>
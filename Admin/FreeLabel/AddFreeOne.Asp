<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="FieldsArr.asp" -->
<%
Dim Conn,strShowErr
MF_Default_Conn
MF_Session_TF 
If Not MF_Check_Pop_TF("MF_sPublic") Then Err_Show

'===========================
'判断系统是否有商城子系统
'===========================
Dim HaveMallTF,GetConfigRs
Set GetConfigRs = Conn.ExeCute("Select Sub_Sys_ID From FS_MF_Sub_Sys Where Sub_Sys_ID = 'MS'")
IF GetConfigRs.Eof Then
	HaveMallTF = False
Else
	HaveMallTF = True
End If
GetConfigRs.Close : Set GetConfigRs = Nothing		



Dim ActionStr
Dim LeftFieldLIstStr,RightFieldListStr,NTable,CTable
Dim LabelName,SysTypeStr,SelectNum,DisSQlStr,NNameStr,CNameStr
Dim LeftOrderStr,LeftRulerStr,RightOrderStr,RightRulerStr,Button_VName
Dim Select_Table_Fields_left,Select_Table_Fields_Right,Free_Label_ID
Dim To_Two_Act,Label_Des_Str
Dim CheckRs,Lable_ContentStr
ActionStr = Trim(Request.QueryString("Act"))
'=========================
'新建和编辑
'=========================
If ActionStr = "Add" Then
	Button_VName = "设置"
	LeftFieldLIstStr = GetFieldsList(NSConArr,"ConStr",NSAllFCNArr,NSAllFENArr,"FirstList","NS")
	RightFieldListStr = GetFieldsList(NS_CConArr,"ConStr",NS_CAllFCNArr,NS_CAllENArr,"SecList","NC")
	NTable = "FS_NS_News"
	CTable = "FS_NS_NewsClass"
	LabelName = ""    
	SysTypeStr = "NS"
	SelectNum = "10"
	DisSQlStr = ""
	NNameStr = "新闻表"
	CNameStr = "新闻栏目表"
	LeftOrderStr = GetFieldsList(NSOrderArr,"SetOrder",NSAllFCNArr,NSAllFENArr,"LeftOrder","NS")
	LeftRulerStr = GetFieldsList(NSAllFCNArr,"SetV",NSAllFCNArr,NSAllFENArr,"LeftRuler","NS")
	RightOrderStr = GetFieldsList(NS_COrderArr,"SetOrder",NS_CAllFCNArr,NS_CAllENArr,"RightOrder","NC")
	RightRulerStr = GetFieldsList(NS_CAllFCNArr,"SetV",NS_CAllFCNArr,NS_CAllENArr,"RightRuler","NC")
	Select_Table_Fields_left = ""
	Select_Table_Fields_Right = ""
	Free_Label_ID = ""
	To_Two_Act = "Add"
	Label_Des_Str = ""
	Lable_ContentStr = ""
Else
	Free_Label_ID = Request.QueryString("LableID")
	Set CheckRs = Conn.ExeCute("Select LabelID,LabelName,LabelSQl,NSFields,NCFields,LabelContent,selectNum,DesCon,SysType From FS_MF_FreeLabel Where LabelID = '" & NoSqlHack(Free_Label_ID) & "'")
	SysTypeStr = CheckRs(8)
	Button_VName = "修改"
	LabelName = Replace(CheckRs(1),"FS400_","")    
	SelectNum = CheckRs(6)  
	DisSQlStr = Replace(CheckRs(2),"*",",")  
	Select_Table_Fields_left = CheckRs(3)   
	Select_Table_Fields_Right = CheckRs(4)  
	To_Two_Act = "Edit"
	Label_Des_Str = CheckRs(7)
	Lable_ContentStr = CheckRs(5)
	If SysTypeStr = "NS" Then
		NTable = "FS_NS_News"
		CTable = "FS_NS_NewsClass"
		NNameStr = "新闻表"
		CNameStr = "新闻栏目表"
		LeftFieldLIstStr = GetFieldsList(NSConArr,"ConStr",NSAllFCNArr,NSAllFENArr,"FirstList","NS")
		RightFieldListStr = GetFieldsList(NS_CConArr,"ConStr",NS_CAllFCNArr,NS_CAllENArr,"SecList","NC")
		LeftOrderStr = GetFieldsList(NSOrderArr,"SetOrder",NSAllFCNArr,NSAllFENArr,"LeftOrder","NS")
		LeftRulerStr = GetFieldsList(NSAllFCNArr,"SetV",NSAllFCNArr,NSAllFENArr,"LeftRuler","NS")
		RightOrderStr = GetFieldsList(NS_COrderArr,"SetOrder",NS_CAllFCNArr,NS_CAllENArr,"RightOrder","NC")
		RightRulerStr = GetFieldsList(NS_CAllFCNArr,"SetV",NS_CAllFCNArr,NS_CAllENArr,"RightRuler","NC")
	ElseIF SysTypeStr = "DS" Then
		NTable = "FS_DS_List"
		CTable = "FS_DS_Class"
		NNameStr = "下载表"
		CNameStr = "下载栏目表"
		LeftFieldLIstStr = GetFieldsList(DSConArr,"ConStr",DSAllFCNArr,DSAllFENArr,"FirstList","DS")
		RightFieldListStr = GetFieldsList(DCConArr,"ConStr",DCAllFCNArr,DCAllFENArr,"SecList","DC")
		LeftOrderStr = GetFieldsList(DSOrderArr,"SetOrder",DSAllFCNArr,DSAllFENArr,"LeftOrder","DS")
		LeftRulerStr = GetFieldsList(DSAllFCNArr,"SetV",DSAllFCNArr,DSAllFENArr,"LeftRuler","DS")
		RightOrderStr = GetFieldsList(DCOrderArr,"SetOrder",DCAllFCNArr,DCAllFENArr,"RightOrder","DC")
		RightRulerStr = GetFieldsList(DCAllFCNArr,"SetV",DCAllFCNArr,DCAllFENArr,"RightRuler","DC")
	ElseIf SysTypeStr = "MS" Then
		NTable = "FS_MS_Products"
		CTable = "FS_MS_ProductsClass"
		NNameStr = "商品表"
		CNameStr = "商品栏目表"
		LeftFieldLIstStr = GetFieldsList(MSConArr,"ConStr",MSAllFCNArr,MSAllFENArr,"FirstList","MS")
		RightFieldListStr = GetFieldsList(MCConArr,"ConStr",MCAllFCNArr,MCAllFENArr,"SecList","MC")
		LeftOrderStr = GetFieldsList(MSOrderArr,"SetOrder",MSAllFCNArr,MSAllFENArr,"LeftOrder","MS")
		LeftRulerStr = GetFieldsList(MSAllFCNArr,"SetV",MSAllFCNArr,MSAllFENArr,"LeftRuler","MS")
		RightOrderStr = GetFieldsList(MCOrderArr,"SetOrder",MCAllFCNArr,MCAllFENArr,"RightOrder","MC")
		RightRulerStr = GetFieldsList(MCAllFCNArr,"SetV",MCAllFCNArr,MCAllFENArr,"RightRuler","MC")
	End If
	CheckRs.Close : Set CheckRs = Nothing		
End If
'==================================================
'取得字段列表
'==================================================
Function GetFieldsList(MainArr,Str_Type,AllFieldsArr,AllFieldsENArr,NameID,FS_SysType)
	Dim IsMu,SelectStyle,IsJavaStr,AutoOP
	Dim i,ArrNum,FieldsValue,NameIDStr
	If Not IsArray(MainArr) Then Exit Function
	If Str_Type = "ConStr" Then
		IsMu = " multiple"
		SelectStyle = " style=""width:100%; height:150px; margin:0px; padding:0px;"""
		IsJavaStr = " onChange=""AddFieldsToSQL();"""
		AutoOP = ""
	Else
		IsMu = ""
		SelectStyle = " style=""margin:0px; padding:0px;"""
		If Str_Type = "SetV" Then
			IsJavaStr = " onChange=""DisFields_Type(this.options[this.selectedIndex].value,'" & FS_SysType & "');"""
			AutoOP = "<option value="""">条件字段</option>" & Chr(13)
		Else
			IsJavaStr = ""
			AutoOP = "<option value="""">排序字段</option>" & Chr(13)
		End If		
	End If
	NameIDStr = NameID
	GetFieldsList = "<select name=""" & NameIDStr & """ id=""" & NameIDStr & """" & IsMu & SelectStyle & IsJavaStr & ">" & Chr(13)
	GetFieldsList = GetFieldsList & AutoOP
	For i = LBound(MainArr) To UBound(MainArr)
		ArrNum = GetInnerFieldsNum(MainArr(i),AllFieldsArr)
		FieldsValue = AllFieldsENArr(ArrNum)
		GetFieldsList = GetFieldsList & "<option value=""" & FieldsValue & """>" & MainArr(i) & "</option>" & Chr(13)
	Next
	GetFieldsList = GetFieldsList & "</select>" & Chr(13)
End Function	
%>
<html>
<head>
<title>自由标签管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" src="../../FS_Inc/Prototype.js" type="text/JavaScript"></script>
<body oncontextmenu="return false;">
<table width="98%" height="50" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="hback" >
    <td width="100%" height="25"  align="Left" class="xingmu" valign="middle">创建自由标签</td>
  </tr>
  <tr class="hback" >
    <td height="25" align="center" class="hback" valign="middle"><div style="line-height:25px; text-align:left;"><span onClick="SubmitFun()" style="cursor:hand;">下一步</span>┆<span onClick="javascript:history.back();" style="cursor:hand;">返回</span></div></td>
  </tr>
</table>
<form name="List_Form" id="List_Form" action="AddFreeTwo.asp" method="post" style="margin:0px;">
  <table width="98%" height="30" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
    <tr class="hback" >
       <td width="30%" height="25" align="left" valign="middle" class="hback">
	   标签名称：	   
  	   <input name="LabelName" id="LabelName" type="text" value="<% = LabelName %>" style="width:160px;">
	   </td>
      <td width="40%" align="center" valign="middle" class="hback"> 选择子系统： 新闻
        <input name="SysType" type="radio" value="NS" <% If SysTypeStr = "NS" Then Response.Write "checked" %> onClick="SetTableName('NS')" />
        下载
        <input name="SysType" type="radio" value="DS" <% If SysTypeStr = "DS" Then Response.Write "checked" %> onClick="SetTableName('DS')" />
        <% If HaveMallTF Then %>
		商城
        <input name="SysType" type="radio" value="MS" <% If SysTypeStr = "MS" Then Response.Write "checked" %> onClick="SetTableName('MS')" />
		<% End If %>
        <input name="NTable" id="NTable" type="hidden" value="<% = NTable %>" />
        <input name="CTable" id="CTable" type="hidden" value="<% = CTable %>" />
		<input name="LabelID" id="LabelID" type="hidden" value="<% = Free_Label_ID %>" />
		<input name="To_Act" id="To_Act" type="hidden" value="<% = To_Two_Act %>" />
		<input name="Label_Des" id="Label_Des" type="hidden" value="<% = Label_Des_Str %>" />
		<textarea name="Lable_ConStr" id="Lable_ConStr" style="display:none;" readonly="readonly"><% = Lable_ContentStr %></textarea>
	  </td>
	  <td width="30%" align="center" valign="middle" class="hback">
	  查询数量：	  
	    <input name="SelectNum" id="SelectNum" type="text" value="<% = SelectNum %>" style="width:60px;" onKeyUp=if(isNaN(value)||event.keyCode==32)execCommand('undo');  onafterpaste=if(isNaN(value)||event.keyCode==32)execCommand('undo');>
        <input name="buttonNum" id="buttonNum" type="button" value="<% = Button_VName %>" onClick="SetNumToSql();">
	  </td>
    </tr>
    <tr class="hback" >
      <td height="85" colspan="3" align="left" valign="middle" class="hback" style="WORD-BREAK: break-all; TABLE-LAYOUT: fixed;"> 
	  	<input type="button" value="SQL语句预览：" disabled="disabled">
        <input type="button" value="检测SQL语句" onClick="CheckSql()">
		<!--input type="button" value="刷新" onClick="javascript:location.reload();"-->
		<span id="CheckSql_Txt" class="tx" style="text-align:center;"></span>
		<textarea name="DisSql" id="DisSql" style="width:100%; height:60px;" readonly="readonly"><% = DisSQlStr %></textarea>      </td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
    <tr class="hback" >
      <td height="25" colspan="2" align="left" valign="middle" class="hback">
	  <span id="FirstTableName"><% = NNameStr %></span>
	  <input name="Fist_TF_All" id="Fist_TF_All" type="hidden" value="<% = Select_Table_Fields_left %>" />
	  </td>
      <td colspan="2" align="left" valign="middle" class="hback">
	  <span id="SecTableName"><% = CNameStr %></span>
	  <input name="Sec_TF_All" id="Sec_TF_All" type="hidden" value="<% = Select_Table_Fields_Right %>" />
	  </td>
    </tr>
    <tr class="hback" >
      <td colspan="2" align="center" valign="middle" class="hback" id="FistFieldsListTable">
	  <img src="../images/load.gif" width="39" height="39" border="0" align="absmiddle">正在加载...请稍微候...	  
	  </td>
      <td colspan="2" align="center" valign="middle" class="hback" id="SecFieldsListTable">
	  <img src="../images/load.gif" width="39" height="39" border="0" align="absmiddle">正在加载...请稍微候...  
	  </td>
    </tr>
    <tr class="hback" >
      <td colspan="2" align="left" valign="middle" class="hback">
	  设置排序：
		<span id="LeftOrderListSpan">
		 <img src="../images/load.gif" width="19" height="19" border="0" align="absmiddle">正在加载...请稍微候...
		</span>
		<select name="LeftOrderRuler" id="LeftOrderRuler" style="margin:0px; padding:0px;">
			<option value="">设置排序</option>
			<option value="Asc">升序</option>
			<option value="Desc">降序</option>
		</select>
		<input type="button" name="OrderbuttonLeft" value="<% = Button_VName %>" onClick="SetOrderToSQl('left')">
	  </td>
      <td colspan="2" align="left" valign="middle" class="hback">
	  	设置排序：
		<span id="RightOrderListSpan">
		<img src="../images/load.gif" width="19" height="19" border="0" align="absmiddle">正在加载...请稍微候...
		</span>
		<select name="RightOrderRuler" id="RightOrderRuler" style="margin:0px; padding:0px;">
			<option value="">设置排序</option>
			<option value="Asc">升序</option>
			<option value="Desc">降序</option>
		</select>
		<input type="button" name="OrderbuttonRight" value="<% = Button_VName %>" onClick="SetOrderToSQl('right')">  
	 </td>
    </tr>
    <tr class="hback" >
      <td width="9%" align="left" valign="middle" class="hback">
	  	查询条件：	  
	  </td>
      <td width="41%" align="left" valign="middle" class="hback">
	  	<span id="LeftRulerSetSpan">
		<img src="../images/load.gif" width="19" height="19" border="0" align="absmiddle">正在加载...请稍微候...
		</span>
		<span class="tx" id="LeftDisFidldsType_Txt" style="margin-left:30px;"></span>
		<br />
		<select name="LeftRulers" id="LeftRulers" style="margin:0px; padding:0px;">
			<option value="">数据关系</option>
			<option value="<>">&lt;&gt;</option>
			<option value="<=">&lt;=</option>
			<option value=">=">&gt;=</option>
			<option value="<">&lt;</option>
			<option value=">">&gt;</option>
			<option value="=">=</option>
			<option value="In">In</option>
			<option value="Like">Like</option>
		</select>
		<input name="LeftRulerTxt" id="LeftRulerTxt" type="text" value="" style="width:100px;">
		<input name="RulerbuttonLeft" id="RulerbuttonLeft" type="button" value="<% = Button_VName %>" onClick="AddWhereToSQL('left')">
	  </td>
      <td width="9%" align="left" valign="middle" class="hback">
	  	查询条件：	  
	  </td>
      <td width="41%" align="left" valign="middle" class="hback">
	  	<span id="RightRulerSetSpan"> 
		<img src="../images/load.gif" width="19" height="19" border="0" align="absmiddle">正在加载...请稍微候...
		</span>
		<span class="tx" id="RightDisFidldsType_Txt" style="margin-left:30px;"></span>
		<br />
		<select name="RightRulers" id="RightRulers" style="margin:0px; padding:0px;">
			<option value="">数据关系</option>
			<option value="<>">&lt;&gt;</option>
			<option value="<=">&lt;=</option>
			<option value=">=">&gt;=</option>
			<option value="<">&lt;</option>
			<option value=">">&gt;</option>
			<option value="=">=</option>
			<option value="In">In</option>
			<option value="Like">Like</option>
		</select>
		<input name="RightRulerTxt" id="RightRulerTxt" type="text" value="" style="width:100px;">
		<input name="RulerbuttonRight" id="RulerbuttonRight" type="button" value="<% = Button_VName %>" onClick="AddWhereToSQL('right')">
	  </td>
    </tr>
  </table>
</form>
<table width="98%" height="20" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="hback" >
    <td width="33%" height="20"  align="center" valign="middle"><span id="N_C_Type_Txt" style="cursor:hand;" onClick="Dis_Data_Table(Fs_NewsTable)">新闻栏目对照表</span></td>
    <td width="33%"  align="center" valign="middle"><span id="D_C_Type_Txt" style="cursor:hand;" onClick="Dis_Data_Table(Fs_DownTable)">下载栏目对照表</span></td>
    <% If HaveMallTF Then %>
	<td width="33%"  align="center" valign="middle"><span id="M_C_Type_Txt" style="cursor:hand;" onClick="Dis_Data_Table(Fs_MallTable)">商城栏目对照表</span></td>
	<% End If %>
  </tr>
</table>
<table width="98%" height="20" border="0" align="center" cellpadding="4" cellspacing="1" class="table" id="Fs_NewsTable" style="display:none;">
  <tr class="hback" >
    <td width="25%" height="20" colspan="4"  align="left" valign="middle">新闻栏目对照表</td>
  </tr>
  <tr class="hback" >
    <td width="25%" height="20"  align="center" valign="middle">栏目中文名</td>
    <td width="25%"  align="center" valign="middle">栏目编号</td>
    <td width="25%"  align="center" valign="middle">栏目中文名</td>
    <td width="25%"  align="center" valign="middle">栏目编号</td>
  </tr>
<%
Dim News_Rs,News_I,N_C_num,JO_Num
Set News_Rs = Server.CreateObject(G_FS_RS)
News_Rs.open "Select ClassName,ClassID From Fs_Ns_NewsClass Where ID > 0 Order By ID Desc",Conn,1,1
If News_Rs.Eof Then
	With Response
		.Write "<tr class=""hback"" >" & vbnewline
		.Write "<td  height=""25"" colspan=""4""  align=""center"" valign=""middle"">暂无数据</td>" & vbnewline
		.Write "</tr> " & vbnewline
	End With
Else
	N_C_num = 0
	Do While Not News_Rs.Eof
		N_C_num = N_C_num + 1
		Response.Write "<tr class=""hback"">" & vbnewline
		For News_I = 1 To 2
			If News_Rs.Eof Then Exit For 	
			With Response
				.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle"">" & News_Rs(0) & "</td>" & vbnewline
				.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle"">" & News_Rs(1) & "</td>" & vbnewline
			End With
			News_Rs.MoveNext
		Next
		If News_Rs.RecordCount = 1 Then
			Response.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle""></td>" & vbnewline
			Response.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle""></td>" & vbnewline
		Else
			If Cint(News_Rs.RecordCount + 1) Mod 4 = 0 Then
				JO_Num = Cint(News_Rs.RecordCount/2)
			Else
				JO_Num = Cint(News_Rs.RecordCount/2) + 1
			End If	
			If News_Rs.RecordCount Mod 2 <> 0 And Cint(N_C_num) = Cint(JO_Num) Then
				Response.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle""></td>" & vbnewline
				Response.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle""></td>" & vbnewline
			End If
		End If	
		Response.Write "</tr>" & vbnewline
	Loop
End If
News_Rs.Close : Set News_Rs = NOthing			
%>  
</table>
<table width="98%" height="20" border="0" align="center" cellpadding="4" cellspacing="1" class="table" id="Fs_DownTable" style="display:none;">
  <tr class="hback" >
    <td width="25%" height="25" colspan="4"  align="left" valign="middle">下载栏目对照表</td>
  </tr>
  <tr class="hback" >
    <td width="25%" height="25"  align="center" valign="middle">栏目中文名</td>
    <td width="25%"  align="center" valign="middle">栏目编号</td>
    <td width="25%"  align="center" valign="middle">栏目中文名</td>
    <td width="25%"  align="center" valign="middle">栏目编号</td>
  </tr>
<%
Dim DS_Rs,DS_I,D_C_num,D_JO_Num
Set DS_Rs = Server.CreateObject(G_FS_RS)
DS_Rs.open "Select ClassName,ClassID From FS_DS_Class Where ID > 0 Order By ID Desc",Conn,1,1
If DS_Rs.Eof Then
	With Response
		.Write "<tr class=""hback"" >" & vbnewline
		.Write "<td  height=""25"" colspan=""4""  align=""center"" valign=""middle"">暂无数据</td>" & vbnewline
		.Write "</tr> " & vbnewline
	End With
Else
	Do While Not DS_Rs.Eof
	D_C_num = D_C_num + 1
		Response.Write "<tr class=""hback"">" & vbnewline
		For DS_I = 1 To 2
		If DS_Rs.Eof Then Exit For 	
			With Response
				.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle"">" & DS_Rs(0) & "</td>" & vbnewline
				.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle"">" & DS_Rs(1) & "</td>" & vbnewline
			End With
		DS_Rs.MoveNext
		Next
		If DS_Rs.RecordCount = 1 Then
			Response.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle""></td>" & vbnewline
			Response.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle""></td>" & vbnewline
		Else
			If Cint(DS_Rs.RecordCount + 1) Mod 4 = 0 Then
				D_JO_Num = Cint(DS_Rs.RecordCount/2)
			Else
				D_JO_Num = Cint(DS_Rs.RecordCount/2) + 1
			End If	
			If DS_Rs.RecordCount Mod 2 <> 0 And Cint(D_C_num) = Cint(D_JO_Num) Then
				Response.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle""></td>" & vbnewline
				Response.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle""></td>" & vbnewline
			End If
		End If	
		Response.Write "</tr>" & vbnewline
	Loop
End If
DS_Rs.Close : Set DS_Rs = NOthing			
%>  
</table>
<% If HaveMallTF Then %>
<table width="98%" height="20" border="0" align="center" cellpadding="4" cellspacing="1" class="table" id="Fs_MallTable" style="display:none;">
  <tr class="hback" >
    <td width="25%" height="25" colspan="4"  align="left" valign="middle">商城栏目对照表</td>
  </tr>
  <tr class="hback" >
    <td width="25%" height="25"  align="center" valign="middle">栏目中文名</td>
    <td width="25%"  align="center" valign="middle">栏目编号</td>
    <td width="25%"  align="center" valign="middle">栏目中文名</td>
    <td width="25%"  align="center" valign="middle">栏目编号</td>
  </tr>
<%
Dim MS_Rs,MS_I,M_C_num,M_JO_Num
Set MS_Rs = Server.CreateObject(G_FS_RS)
MS_Rs.open "Select ClassCName,ClassID From FS_MS_ProductsClass Where ID > 0 Order By ID Desc",Conn,1,1
If MS_Rs.Eof Then
	With Response
		.Write "<tr class=""hback"" >" & vbnewline
		.Write "<td  height=""25"" colspan=""4""  align=""center"" valign=""middle"">暂无数据</td>" & vbnewline
		.Write "</tr> " & vbnewline
	End With
Else
	Do While Not MS_Rs.Eof
	M_C_num = M_C_num + 1
		Response.Write "<tr class=""hback"">" & vbnewline
		For MS_I = 1 To 2
		If MS_Rs.Eof Then Exit For 	
			With Response
				.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle"">" & MS_Rs(0) & "</td>" & vbnewline
				.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle"">" & MS_Rs(1) & "</td>" & vbnewline
			End With
		MS_Rs.MoveNext
		Next
		If MS_Rs.RecordCount = 1 Then
			Response.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle""></td>" & vbnewline
			Response.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle""></td>" & vbnewline
		Else
			If Cint(MS_Rs.RecordCount + 1) Mod 4 = 0 Then
				M_JO_Num = Cint(MS_Rs.RecordCount/2)
			Else
				M_JO_Num = Cint(MS_Rs.RecordCount/2) + 1
			End If	
			If MS_Rs.RecordCount Mod 2 <> 0 And Cint(M_C_num) = Cint(M_JO_Num) Then
				Response.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle""></td>" & vbnewline
				Response.Write "<td width=""25%"" height=""25"" align=""center"" valign=""middle""></td>" & vbnewline
			End If
		End If	
		Response.Write "</tr>" & vbnewline
	Loop
End If
MS_Rs.Close : Set MS_Rs = NOthing			
%>  
</table>
<% End If %>


<table width="98%" height="40" border="0" align="center" cellpadding="4" cellspacing="1">
  <tr class="hback" >
    <td width="100%" height="40"  align="Left" valign="middle">
	<span class="tx" style="line-height:20px;">说明：<br />1.此功能建议具有一定sql基础的人使用。如果不熟悉sql，请谨慎使用，以防误操作引起的数据损坏。<br />
	2.数据库表字段可按CTRL、shift进行多选和取消。<br />
	3.排序和查询条件，同一字段，第一次设置为设置该条件，同条件设置第二次则取消该条件。<br />
	4.判断型数字赋值：0为否，1为是。<br />
	5.如果设置条件的字段为文本、文本型ID等类型，赋值时请前后加 ' ，如：= '内容' 、In('内容1'，'内容2'),否则会出错；Like关系除外。<br />
	6.在下一步设置之前，建议先点击测试按钮确认sql语句没有问题再继续。
	
	</span>
	</td>
  </tr>
</table>
</body>
</html>
<%
Conn.CLose : Set Conn = Nothing
%>
<script language="javascript" type="text/javascript">
<!--
//新建标签时初始化页面
var LeftList_Str_Fields = '<% = Replace(Replace(LeftFieldLIstStr,Chr(13),""),"'","\'") %>';
var RightList_Str_Fields = '<% = Replace(Replace(RightFieldListStr,Chr(13),""),"'","\'") %>';
var LeftList_Str_Order = '<% = Replace(Replace(LeftOrderStr,Chr(13),""),"'","\'") %>';
var RightList_Str_Order = '<% = Replace(Replace(RightOrderStr,Chr(13),""),"'","\'") %>';
var LeftList_Str_Rulers = '<% = Replace(Replace(LeftRulerStr,Chr(13),""),"'","\'") %>';
var RightList_Str_Rulers = '<% = Replace(Replace(RightRulerStr,Chr(13),""),"'","\'") %>';
document.getElementById('FistFieldsListTable').innerHTML = LeftList_Str_Fields;
document.getElementById('SecFieldsListTable').innerHTML = RightList_Str_Fields;
document.getElementById('LeftOrderListSpan').innerHTML = LeftList_Str_Order;
document.getElementById('RightOrderListSpan').innerHTML = RightList_Str_Order;
document.getElementById('LeftRulerSetSpan').innerHTML = LeftList_Str_Rulers;
document.getElementById('RightRulerSetSpan').innerHTML = RightList_Str_Rulers;
 
 
//子系统切换时更新显示操作字段
function SetTableName(str)
{
	var LeftList_Str_Fields = '';
	var RightList_Str_Fields = '';
	var LeftList_Str_Order = '';
	var RightList_Str_Order = '';
	var LeftList_Str_Rulers = '';
	var RightList_Str_Rulers = '';
	var NTable_Name = '';
	var CTable_Name = '';
	var N_TableName = '';
	var C_TableName = '';
	var Type_Str_Auto = '<% = SysTypeStr %>';
	var Auto_SQl_Str = '<% = Replace(DisSQlStr,"'","\'") %>';
	var Auto_Left_AllFields = '<% = Select_Table_Fields_left %>';
	var Auto_Right_AllFields = '<% = Select_Table_Fields_Right %>';
	var Auto_Lable_ConStr = '<% = Replace(Replace(Replace(Replace(Lable_ContentStr,"'","\'"),"/","\/"),Chr(10),""),Chr(13),"") %>';
	var Auto_SeNum = '<% = SelectNum %>';
	if (Type_Str_Auto == str) 
	{
		document.getElementById('DisSql').value = Auto_SQl_Str;
		document.getElementById('Fist_TF_All').value = Auto_SQl_Str;
		document.getElementById('Sec_TF_All').value = Auto_SQl_Str;
		document.getElementById('Lable_ConStr').value = Auto_Lable_ConStr;
		document.getElementById('SelectNum').value = Auto_SeNum;
	}    
	else
	{
		document.getElementById('DisSql').value = '';
		document.getElementById('Fist_TF_All').value = '';
		document.getElementById('Sec_TF_All').value = '';
		document.getElementById('Lable_ConStr').value = '';
		document.getElementById('SelectNum').value = '10';
	}
	if (str == 'DS')
	{
		LeftList_Str_Fields = '<% = Replace(Replace(GetFieldsList(DSConArr,"ConStr",DSAllFCNArr,DSAllFENArr,"FirstList","DS"),Chr(13),""),"'","\'") %>';
		RightList_Str_Fields = '<% = Replace(Replace(GetFieldsList(DCConArr,"ConStr",DCAllFCNArr,DCAllFENArr,"SecList","DC"),Chr(13),""),"'","\'") %>';
		LeftList_Str_Order = '<% = Replace(Replace(GetFieldsList(DSOrderArr,"SetOrder",DSAllFCNArr,DSAllFENArr,"LeftOrder","DS"),Chr(13),""),"'","\'") %>';
		RightList_Str_Order = '<% = Replace(Replace(GetFieldsList(DCOrderArr,"SetOrder",DCAllFCNArr,DCAllFENArr,"RightOrder","DC"),Chr(13),""),"'","\'") %>';
		LeftList_Str_Rulers = '<% = Replace(Replace(GetFieldsList(DSAllFCNArr,"SetV",DSAllFCNArr,DSAllFENArr,"LeftRuler","DS"),Chr(13),""),"'","\'") %>';
		RightList_Str_Rulers = '<% = Replace(Replace(GetFieldsList(DCAllFCNArr,"SetV",DCAllFCNArr,DCAllFENArr,"RightRuler","DC"),Chr(13),""),"'","\'") %>';
		NTable_Name = 'FS_DS_List';
		CTable_Name = 'FS_DS_Class';
		N_TableName = '下载表';
		C_TableName = '下载栏目表';
	}
	else if (str == 'MS')
	{
		LeftList_Str_Fields = '<% = Replace(Replace(GetFieldsList(MSConArr,"ConStr",MSAllFCNArr,MSAllFENArr,"FirstList","MS"),Chr(13),""),"'","\'") %>'; 
		RightList_Str_Fields = '<% = Replace(Replace(GetFieldsList(MCConArr,"ConStr",MCAllFCNArr,MCAllFENArr,"SecList","MC"),Chr(13),""),"'","\'") %>';
		LeftList_Str_Order = '<% = Replace(Replace(GetFieldsList(MSOrderArr,"SetOrder",MSAllFCNArr,MSAllFENArr,"LeftOrder","MS"),Chr(13),""),"'","\'") %>';
		RightList_Str_Order = '<% = Replace(Replace(GetFieldsList(MCOrderArr,"SetOrder",MCAllFCNArr,MCAllFENArr,"RightOrder","MC"),Chr(13),""),"'","\'") %>';
		LeftList_Str_Rulers = '<% = Replace(Replace(GetFieldsList(MSAllFCNArr,"SetV",MSAllFCNArr,MSAllFENArr,"LeftRuler","MS"),Chr(13),""),"'","\'") %>';
		RightList_Str_Rulers = '<% = Replace(Replace(GetFieldsList(MCAllFCNArr,"SetV",MCAllFCNArr,MCAllFENArr,"RightRuler","MC"),Chr(13),""),"'","\'") %>';
		NTable_Name = 'FS_MS_Products';
		CTable_Name = 'FS_MS_ProductsClass';
		N_TableName = '商品表';
		C_TableName = '商品栏目表';
	}
	else if (str == 'NS')
	{
		LeftList_Str_Fields = '<% = Replace(Replace(GetFieldsList(NSConArr,"ConStr",NSAllFCNArr,NSAllFENArr,"FirstList","NS"),Chr(13),""),"'","\'") %>'; 
		RightList_Str_Fields = '<% = Replace(Replace(GetFieldsList(NS_CConArr,"ConStr",NS_CAllFCNArr,NS_CAllENArr,"SecList","NC"),Chr(13),""),"'","\'") %>';
		LeftList_Str_Order = '<% = Replace(Replace(GetFieldsList(NSOrderArr,"SetOrder",NSAllFCNArr,NSAllFENArr,"LeftOrder","NS"),Chr(13),""),"'","\'") %>';
		RightList_Str_Order = '<% = Replace(Replace(GetFieldsList(NS_COrderArr,"SetOrder",NS_CAllFCNArr,NS_CAllENArr,"RightOrder","NC"),Chr(13),""),"'","\'") %>';
		LeftList_Str_Rulers = '<% = Replace(Replace(GetFieldsList(NSAllFCNArr,"SetV",NSAllFCNArr,NSAllFENArr,"LeftRuler","NS"),Chr(13),""),"'","\'") %>';
		RightList_Str_Rulers = '<% = Replace(Replace(GetFieldsList(NS_CAllFCNArr,"SetV",NS_CAllFCNArr,NS_CAllENArr,"RightRuler","NC"),Chr(13),""),"'","\'") %>';
		NTable_Name = 'FS_NS_News';
		CTable_Name = 'FS_NS_NewsClass';
		N_TableName = '新闻表';
		C_TableName = '新闻栏目表';
	}    
	document.getElementById('FirstTableName').innerHTML = N_TableName;
	document.getElementById('SecTableName').innerHTML = C_TableName;
	document.getElementById('NTable').value = NTable_Name;
	document.getElementById('CTable').value = CTable_Name;
	document.getElementById('LeftDisFidldsType_Txt').innerHTML = '';
	document.getElementById('RightDisFidldsType_Txt').innerHTML = '';
	document.getElementById('CheckSql_Txt').innerHTML = '';
	document.getElementById('FistFieldsListTable').innerHTML = LeftList_Str_Fields;
	document.getElementById('SecFieldsListTable').innerHTML = RightList_Str_Fields;
	document.getElementById('LeftOrderListSpan').innerHTML = LeftList_Str_Order;
	document.getElementById('RightOrderListSpan').innerHTML = RightList_Str_Order;
	document.getElementById('LeftRulerSetSpan').innerHTML = LeftList_Str_Rulers;
	document.getElementById('RightRulerSetSpan').innerHTML = RightList_Str_Rulers;
}         
 
 
//显示数据关系字段类型
function DisFields_Type(str,Type_Str)
{
	var Temp_Str = '';
	var TempStr = '';
	var Temp_Arr = '';
	var TempArr = '';
	var DisInfoID_Str = '';
	switch (Type_Str)
	{
		case "NS":
			Temp_Str = '<% = NewAllENameFields %>';
			TempStr = '<% = NewAllENameFieldsType %>';
			DisInfoID_Str = document.getElementById('LeftDisFidldsType_Txt');
			break;
		case "NC":
			Temp_Str = '<% = NCAllENameFields %>';
			TempStr = '<% = NCAllENameFieldsType %>';
			DisInfoID_Str = document.getElementById('RightDisFidldsType_Txt');
			break;
		case "DS":
			Temp_Str = '<% = DownAllENameFields %>';
			TempStr = '<% = DownAllENameFieldsType %>';
			DisInfoID_Str = document.getElementById('LeftDisFidldsType_Txt');
			break;
		case "DC":
			Temp_Str = '<% = D_CAllENameFields %>';
			TempStr = '<% = D_CAllENameFieldsType %>';
			DisInfoID_Str = document.getElementById('RightDisFidldsType_Txt');
			break;
		case "MS":
			Temp_Str = '<% = MallAllENameFields %>';
			TempStr = '<% = MallAllENameFieldsType %>';
			DisInfoID_Str = document.getElementById('LeftDisFidldsType_Txt');
			break;
		case "MC":
			Temp_Str = '<% = M_CAllENameFields %>';
			TempStr = '<% = M_CAllENameFieldsType %>';
			DisInfoID_Str = document.getElementById('RightDisFidldsType_Txt');
			break;	
	}
	Temp_Arr = Temp_Str.split('||');
	TempArr = TempStr.split(',');
	for (i=0;i<=Temp_Arr.length-1;i++)
	{
		if (Temp_Arr[i] == str)
		{
			DisInfoID_Str.innerHTML = '该字段类型：' + TempArr[i];
			return;
		}
	}
}


//添加所选字段到sql语句中       
function AddFieldsToSQL()
{
	document.List_Form.DisSql.value = '';
	document.List_Form.Fist_TF_All.value = '';
	document.List_Form.Sec_TF_All.value = '';  
	var LeftFields_Str = document.List_Form.FirstList;
	var RightFields_Str = document.List_Form.SecList;
	var LeftTable_Name = document.List_Form.NTable.value;
	var RightTable_Name = document.List_Form.CTable.value;
	var SelectNum = document.List_Form.SelectNum.value;
	var V_StrArr = '';
	var V_StartStr = '';
	var V_EndStr = '';
	if (SelectNum == '')
	{
		alert('查询数量请为正整数');
		document.List_Form.SelectNum.focus();
		return;
	}
	if (LeftFields_Str.value != '')
	{
		for(i=0;i<LeftFields_Str.length;i++)
		{
			if (LeftFields_Str[i].selected == true )
			{
				if (document.List_Form.Fist_TF_All.value == '')
				{
					if (LeftFields_Str[i].value.indexOf('Mid') == -1 && LeftFields_Str[i].value.indexOf('SubString') == -1)
					{
						document.List_Form.Fist_TF_All.value = LeftTable_Name + '.' + LeftFields_Str[i].value;
					}
					else
					{
						V_StrArr = LeftFields_Str[i].value.split('(');
						V_StartStr = V_StrArr[0].replace(/^\s+/,'').replace(/\s+$/,'');
						V_EndStr = V_StrArr[1].replace(/^\s+/,'').replace(/\s+$/,'');
						document.List_Form.Fist_TF_All.value = V_StartStr + '(' + LeftTable_Name + '.' + V_EndStr;
					}
				}
				else
				{
					if (LeftFields_Str[i].value.indexOf('Mid') == -1 && LeftFields_Str[i].value.indexOf('SubString') == -1)
					{
						document.List_Form.Fist_TF_All.value = document.List_Form.Fist_TF_All.value + ',' + LeftTable_Name + '.' + LeftFields_Str[i].value;
					}
					else
					{
						V_StrArr = LeftFields_Str[i].value.split('(');
						V_StartStr = V_StrArr[0].replace(/^\s+/,'').replace(/\s+$/,'');
						V_EndStr = V_StrArr[1].replace(/^\s+/,'').replace(/\s+$/,'');
						document.List_Form.Fist_TF_All.value = document.List_Form.Fist_TF_All.value + ',' + V_StartStr + '(' + LeftTable_Name + '.' + V_EndStr;
					}
				}
				if (RightFields_Str.value == '')
				{
					if (document.List_Form.DisSql.value == '')
					{
						document.List_Form.DisSql.value = LeftFields_Str[i].value;
					}
					else
					{
						document.List_Form.DisSql.value = document.List_Form.DisSql.value + ',' + LeftFields_Str[i].value;
					}
				}	
				else
				{
					if (document.List_Form.DisSql.value == '')
					{
						if (LeftFields_Str[i].value.indexOf('Mid') == -1 && LeftFields_Str[i].value.indexOf('SubString') == -1)
						{
							document.List_Form.DisSql.value = LeftTable_Name + '.' + LeftFields_Str[i].value;
						}
						else
						{
							V_StrArr = LeftFields_Str[i].value.split('(');
							V_StartStr = V_StrArr[0].replace(/^\s+/,'').replace(/\s+$/,'');
							V_EndStr = V_StrArr[1].replace(/^\s+/,'').replace(/\s+$/,'');
							document.List_Form.DisSql.value = V_StartStr + '(' + LeftTable_Name + '.' + V_EndStr;
						}
					}
					else
					{
						if (LeftFields_Str[i].value.indexOf('Mid') == -1 && LeftFields_Str[i].value.indexOf('SubString') == -1)
						{
							document.List_Form.DisSql.value = document.List_Form.DisSql.value + ',' + LeftTable_Name + '.' + LeftFields_Str[i].value;
						}
						else
						{
							V_StrArr = LeftFields_Str[i].value.split('(');
							V_StartStr = V_StrArr[0].replace(/^\s+/,'').replace(/\s+$/,'');
							V_EndStr = V_StrArr[1].replace(/^\s+/,'').replace(/\s+$/,'');
							document.List_Form.DisSql.value = document.List_Form.DisSql.value + ',' + V_StartStr + '(' + LeftTable_Name + '.' + V_EndStr;
						}
					}
				}	
			}
		}
	}	
	if (RightFields_Str.value != '')
	{
		for(i=0;i<RightFields_Str.length;i++)
		{
			if (RightFields_Str[i].selected == true )
			{
				if (document.List_Form.Sec_TF_All.value == '')
				{
					if (RightFields_Str[i].value.indexOf('Mid') == -1 && RightFields_Str[i].value.indexOf('SubString') == -1)
					{
						document.List_Form.Sec_TF_All.value = RightTable_Name + '.' + RightFields_Str[i].value;
					}
					else
					{
						V_StrArr = RightFields_Str[i].value.split('(');
						V_StartStr = V_StrArr[0].replace(/^\s+/,'').replace(/\s+$/,'');
						V_EndStr = V_StrArr[1].replace(/^\s+/,'').replace(/\s+$/,'');
						document.List_Form.Sec_TF_All.value = V_StartStr + '(' + RightTable_Name + '.' + V_EndStr;
					}
				}
				else
				{
					if (RightFields_Str[i].value.indexOf('Mid') == -1 && RightFields_Str[i].value.indexOf('SubString') == -1)
					{
						document.List_Form.Sec_TF_All.value = document.List_Form.Sec_TF_All.value + ',' + RightTable_Name + '.' + RightFields_Str[i].value;
					}
					else
					{
						V_StrArr = RightFields_Str[i].value.split('(');
						V_StartStr = V_StrArr[0].replace(/^\s+/,'').replace(/\s+$/,'');
						V_EndStr = V_StrArr[1].replace(/^\s+/,'').replace(/\s+$/,'');
						document.List_Form.Sec_TF_All.value = document.List_Form.Sec_TF_All.value + ',' + V_StartStr + '(' + RightTable_Name + '.' + V_EndStr;
					}
				}
				if (LeftFields_Str.value == '')
				{
					if (document.List_Form.DisSql.value == '')
					{
						document.List_Form.DisSql.value = RightFields_Str[i].value;
					}
					else
					{
						document.List_Form.DisSql.value = document.List_Form.DisSql.value + ',' + RightFields_Str[i].value;
					}
				}	
				else
				{
					if (document.List_Form.DisSql.value == '')
					{
						if (RightFields_Str[i].value.indexOf('Mid') == -1 && RightFields_Str[i].value.indexOf('SubString') == -1)
						{
							document.List_Form.DisSql.value = RightTable_Name + '.' + RightFields_Str[i].value;
						}
						else
						{
							V_StrArr = RightFields_Str[i].value.split('(');
							V_StartStr = V_StrArr[0].replace(/^\s+/,'').replace(/\s+$/,'');
							V_EndStr = V_StrArr[1].replace(/^\s+/,'').replace(/\s+$/,'');
							document.List_Form.DisSql.value = V_StartStr + '(' + RightTable_Name + '.' + V_EndStr;
						}
					}
					else
					{
						if (RightFields_Str[i].value.indexOf('Mid') == -1 && RightFields_Str[i].value.indexOf('SubString') == -1)
						{
							document.List_Form.DisSql.value = document.List_Form.DisSql.value + ',' + RightTable_Name + '.' + RightFields_Str[i].value;
						}
						else
						{
							V_StrArr = RightFields_Str[i].value.split('(');
							V_StartStr = V_StrArr[0].replace(/^\s+/,'').replace(/\s+$/,'');
							V_EndStr = V_StrArr[1].replace(/^\s+/,'').replace(/\s+$/,'');
							document.List_Form.DisSql.value = document.List_Form.DisSql.value + ',' + V_StartStr + '(' + RightTable_Name + '.' + V_EndStr;
						}
					}
				}	
			}
		}
	}
	if (document.List_Form.DisSql.value != '')
	{
		if (RightFields_Str.value != '' && LeftFields_Str.value != '')
		{
			document.List_Form.DisSql.value = 'Select Top ' + SelectNum + ' ' + document.List_Form.DisSql.value + ' From ' + LeftTable_Name + ',' + RightTable_Name + ' Where ' + LeftTable_Name + '.ClassID = ' + RightTable_Name + '.ClassID';
		}
		else if (LeftFields_Str.value != '' && RightFields_Str.value == '')
		{
			document.List_Form.DisSql.value = 'Select Top ' + SelectNum + ' ' + document.List_Form.DisSql.value + ' From ' + LeftTable_Name;
		}
		else if (LeftFields_Str.value == '' && RightFields_Str.value != '')
		{
			document.List_Form.DisSql.value = 'Select Top ' + SelectNum + ' ' + document.List_Form.DisSql.value + ' From ' + RightTable_Name;
		}
	}
}


//添加排序
function SetOrderToSQl(Str)
{
	var Sql_Str = document.List_Form.DisSql.value;
	var Order_Txt_left = document.List_Form.LeftOrder.value;
	var Order_Txt_right = document.List_Form.RightOrder.value;
	var Order_Ruler_left = document.List_Form.LeftOrderRuler.value;
	var Order_Ruler_right = document.List_Form.RightOrderRuler.value;
	var Select_Fields_Left = document.List_Form.Fist_TF_All.value;
	var Select_Fields_right = document.List_Form.Sec_TF_All.value;
	var LeftTable_Name = document.List_Form.NTable.value;
	var RightTable_Name = document.List_Form.CTable.value;
	var Main_Sql_Str = '';
	var Order_Str = '';
	var Replace_order_Str = '';
	var Sql_Str_Arr = '';
	if (Sql_Str != '')
	{
		if (Str == 'left')
		{
			if (Order_Txt_left != '' && Order_Ruler_left != '')
			{
				if (Select_Fields_right != '')
				{
					if (Sql_Str.indexOf('Order By') == -1)
					{
						document.List_Form.DisSql.value = Sql_Str + ' Order By ' + LeftTable_Name + '.' + Order_Txt_left + ' ' + Order_Ruler_left;
					}
					else
					{
						Sql_Str_Arr = Sql_Str.split('Order By');
						Main_Sql_Str = Sql_Str_Arr[0].replace(/^\s+/,'').replace(/\s+$/,'');
						Order_Str = Sql_Str_Arr[1].replace(/^\s+/,'').replace(/\s+$/,'');
						if (Order_Str.indexOf(LeftTable_Name + '.' + Order_Txt_left) == -1)
						{
							document.List_Form.DisSql.value = Sql_Str + ',' + LeftTable_Name + '.' + Order_Txt_left + ' ' + Order_Ruler_left;
						}
						else
						{
							Replace_order_Str = Order_Str.replace(LeftTable_Name + '.' + Order_Txt_left + ' Asc,','').replace(',' + LeftTable_Name + '.' + Order_Txt_left + ' Asc','').replace(LeftTable_Name + '.' + Order_Txt_left + ' Asc','').replace(LeftTable_Name + '.' + Order_Txt_left + ' Desc,','').replace(',' + LeftTable_Name + '.' + Order_Txt_left + ' Desc','').replace(LeftTable_Name + '.' + Order_Txt_left + ' Desc','').replace(/^\s+/,'').replace(/\s+$/,'');
							if (Replace_order_Str == '')
							{
								document.List_Form.DisSql.value = Main_Sql_Str;
							}
							else
							{
								document.List_Form.DisSql.value = Main_Sql_Str + ' Order By ' + Replace_order_Str;
							}
						}
					}
				}
				else
				{
					if (Sql_Str.indexOf('Order By') == -1)
					{
						document.List_Form.DisSql.value = Sql_Str + ' Order By ' + Order_Txt_left + ' ' + Order_Ruler_left;
					}
					else
					{
						Sql_Str_Arr = Sql_Str.split('Order By');
						Main_Sql_Str = Sql_Str_Arr[0].replace(/^\s+/,'').replace(/\s+$/,'');
						Order_Str = Sql_Str_Arr[1].replace(/^\s+/,'').replace(/\s+$/,'');
						if (Order_Str.indexOf(Order_Txt_left) == -1)
						{
							document.List_Form.DisSql.value = Sql_Str + ',' + Order_Txt_left + ' ' + Order_Ruler_left;
						}
						else
						{
							Replace_order_Str = Order_Str.replace(Order_Txt_left + ' Asc,','').replace(',' + Order_Txt_left + ' Asc','').replace(Order_Txt_left + ' Asc','').replace(Order_Txt_left + ' Desc,','').replace(',' + Order_Txt_left + ' Desc','').replace(Order_Txt_left + ' Desc','').replace(/^\s+/,'').replace(/\s+$/,'');
							if (Replace_order_Str == '')
							{
								document.List_Form.DisSql.value = Main_Sql_Str;
							}
							else
							{
								document.List_Form.DisSql.value = Main_Sql_Str + ' Order By ' + Replace_order_Str;
							}
						}	
					}
				}
			}
			else
			{
				return;
			}
		}
		else if (Str == 'right')
		{
			if (Order_Txt_right != '' && Order_Ruler_right != '')
			{
				if (Select_Fields_Left != '')
				{
					if (Sql_Str.indexOf('Order By') == -1)
					{
						document.List_Form.DisSql.value = Sql_Str + ' Order By ' + RightTable_Name + '.' + Order_Txt_right + ' ' + Order_Ruler_right;
					}
					else
					{
						Sql_Str_Arr = Sql_Str.split('Order By');
						Main_Sql_Str = Sql_Str_Arr[0].replace(/^\s+/,'').replace(/\s+$/,'');
						Order_Str = Sql_Str_Arr[1].replace(/^\s+/,'').replace(/\s+$/,'');
						if (Order_Str.indexOf(RightTable_Name + '.' + Order_Txt_right) == -1)
						{
							document.List_Form.DisSql.value = Sql_Str + ',' + RightTable_Name + '.' + Order_Txt_right + ' ' + Order_Ruler_right;
						}
						else
						{
							Replace_order_Str = Order_Str.replace(RightTable_Name + '.' + Order_Txt_right + ' Asc,','').replace(',' + RightTable_Name + '.' + Order_Txt_right + ' Asc','').replace(RightTable_Name + '.' + Order_Txt_right + ' Asc','').replace(RightTable_Name + '.' + Order_Txt_right + ' Desc,','').replace(',' + RightTable_Name + '.' + Order_Txt_right + ' Desc','').replace(RightTable_Name + '.' + Order_Txt_right + ' Desc','').replace(/^\s+/,'').replace(/\s+$/,'');
							if (Replace_order_Str == '')
							{
								document.List_Form.DisSql.value = Main_Sql_Str;
							}
							else
							{
								document.List_Form.DisSql.value = Main_Sql_Str + ' Order By ' + Replace_order_Str;
							}
						}
					}
				}
				else
				{
					if (Sql_Str.indexOf('Order By') == -1)
					{
						document.List_Form.DisSql.value = Sql_Str + ' Order By ' + Order_Txt_right + ' ' + Order_Ruler_right;
					}
					else
					{
						Sql_Str_Arr = Sql_Str.split('Order By');
						Main_Sql_Str = Sql_Str_Arr[0].replace(/^\s+/,'').replace(/\s+$/,'');
						Order_Str = Sql_Str_Arr[1].replace(/^\s+/,'').replace(/\s+$/,'');
						if (Order_Str.indexOf(Order_Txt_right) == -1)
						{
							document.List_Form.DisSql.value = Sql_Str + ',' + Order_Txt_right + ' ' + Order_Ruler_right;
						}
						else
						{
							Replace_order_Str = Order_Str.replace(Order_Txt_right + ' Asc,','').replace(',' + Order_Txt_right + ' Asc','').replace(Order_Txt_right + ' Asc','').replace(Order_Txt_right + ' Desc,','').replace(',' + Order_Txt_right + ' Desc','').replace(Order_Txt_right + ' Desc','').replace(/^\s+/,'').replace(/\s+$/,'');
							if (Replace_order_Str == '')
							{
								document.List_Form.DisSql.value = Main_Sql_Str;
							}
							else
							{
								document.List_Form.DisSql.value = Main_Sql_Str + ' Order By ' + Replace_order_Str;
							}
						}	
					}
				}
			}
			else
			{
				return;
			}
		}	
	}
	else
	{
		return;
	}	
}


//添加查询条件到sql语句中
function AddWhereToSQL(str)
{
	var Sql_Str = document.List_Form.DisSql.value;
	var Select_Fields_Left = document.List_Form.Fist_TF_All.value;
	var Select_Fields_right = document.List_Form.Sec_TF_All.value;
	var LeftTable_Name = document.List_Form.NTable.value;
	var RightTable_Name = document.List_Form.CTable.value;
	var LeftRulerStr = document.List_Form.LeftRuler.value;
	var RightRulerStr = document.List_Form.RightRuler.value;
	var Left_Ruler = document.List_Form.LeftRulers.value;
	var Right_Ruler = document.List_Form.RightRulers.value;
	var Left_Ruler_Txt = document.List_Form.LeftRulerTxt.value;
	var Right_Ruler_Txt = document.List_Form.RightRulerTxt.value;
	var Ruler_Str = '';
	var Where_Str = '';
	var Sql_Start_Str = '';
	var Sql_End_Str = '';
	var Sql_Split_Arr = '';
	var Replace_Sql_Str = '';
	var split_order_str = '';
	var split_order_Arr = '';
	var Where_No_And = '';
	var Order_End_Str = '';
	var R_Str_Arr = '';
	var R_Str_Str = '';
	var R_Str_End = '';
	if (Sql_Str != '')
	{
		if (str == 'left')
		{
			if (LeftRulerStr != '' && Left_Ruler != '' && Left_Ruler_Txt != '')
			{
				if (Select_Fields_right != '')
				{
					if (LeftRulerStr.indexOf('Mid') == -1 && LeftRulerStr.indexOf('SubString') == -1)
					{
						Ruler_Str = LeftTable_Name + '.' + 	LeftRulerStr;
					}
					else
					{
						R_Str_Arr = LeftRulerStr.split('(');
						R_Str_Str = R_Str_Arr[0].replace(/^\s+/,'').replace(/\s+$/,'');
						R_Str_End = R_Str_Arr[1].replace(/^\s+/,'').replace(/\s+$/,'');
						Ruler_Str = R_Str_Str + '(' + LeftTable_Name + '.' + R_Str_End;
					}
				}
				else
				{
					Ruler_Str = LeftRulerStr;
				}	
				if (Sql_Str.indexOf('Where') == -1)
				{
					if (Sql_Str.indexOf('Order By') == -1)
					{
						switch (Left_Ruler)
						{
							case "In":
								Where_Str = Ruler_Str + ' In(' + Left_Ruler_Txt + ')';
								break;	
							case "Like":
								Where_Str = Ruler_Str + ' Like \'%' + Left_Ruler_Txt + '%\'';
								break;
							default:
								Where_Str = Ruler_Str + ' ' + Left_Ruler + ' ' + Left_Ruler_Txt;
								break;
						}
						document.List_Form.DisSql.value = Sql_Str + ' Where ' + Where_Str;
					}
					else
					{
						Sql_Split_Arr = Sql_Str.split('Order By');
						Sql_Start_Str = Sql_Split_Arr[0].replace(/^\s+/,'').replace(/\s+$/,'');
						Sql_End_Str = Sql_Split_Arr[1].replace(/^\s+/,'').replace(/\s+$/,'');
						switch (Left_Ruler)
						{
							case "In":
								Where_Str = Ruler_Str + ' In(' + Left_Ruler_Txt + ')';
								break;	
							case "Like":
								Where_Str = Ruler_Str + ' Like \'%' + Left_Ruler_Txt + '%\'';
								break;
							default:
								Where_Str = Ruler_Str + ' ' + Left_Ruler + ' ' + Left_Ruler_Txt;
								break;
						}
						document.List_Form.DisSql.value = Sql_Start_Str + ' Where ' + Where_Str + ' Order By ' + Sql_End_Str;
					}
				}
				else
				{
					Sql_Split_Arr = Sql_Str.split('Where');
					Sql_Start_Str = Sql_Split_Arr[0].replace(/^\s+/,'').replace(/\s+$/,'');
					Sql_End_Str = Sql_Split_Arr[1].replace(/^\s+/,'').replace(/\s+$/,'');
					switch (Left_Ruler)
					{
						case "In":
							Where_No_And = Ruler_Str + ' In(' + Left_Ruler_Txt + ')';
							break;	
						case "Like":
							Where_No_And = Ruler_Str + ' Like \'%' + Left_Ruler_Txt + '%\'';
							break;
						default:
							Where_No_And = Ruler_Str + ' ' + Left_Ruler + ' ' + Left_Ruler_Txt;
							break;
					}
					if (Sql_End_Str.indexOf(Where_No_And) == -1)
					{
						Where_Str = Where_No_And + ' And ';
						document.List_Form.DisSql.value = Sql_Start_Str + ' Where ' + Where_Str + Sql_End_Str;
					}
					else
					{
						if (Sql_End_Str.indexOf('Order By') == -1)
						{
							split_order_str = Sql_End_Str
							Order_End_Str = '';
						}
						else
						{
							split_order_Arr = Sql_End_Str.split('Order By');
							split_order_str = split_order_Arr[0].replace(/^\s+/,'').replace(/\s+$/,'');
							Order_End_Str = ' Order By ' + split_order_Arr[1].replace(/^\s+/,'').replace(/\s+$/,'');
						}
						Replace_Sql_Str = split_order_str.replace(Where_No_And + ' And ','').replace(' And ' + Where_No_And,'').replace(Where_No_And,'').replace(/^\s+/,'').replace(/\s+$/,'');
						if (Replace_Sql_Str != '')
						{
							document.List_Form.DisSql.value = Sql_Start_Str + ' Where ' + Replace_Sql_Str + Order_End_Str;
						}
						else
						{
							document.List_Form.DisSql.value = Sql_Start_Str + Order_End_Str;
						}	
					}
				}
			}
			else
			{
				return;
			}
		}
		else if (str == 'right')
		{
			if (RightRulerStr != '' && Right_Ruler != '' && Right_Ruler_Txt != '')
			{
				if (Select_Fields_Left != '')
				{
					if (RightRulerStr.indexOf('Mid') == -1 && RightRulerStr.indexOf('SubString') == -1)
					{
						Ruler_Str = RightTable_Name + '.' + RightRulerStr;
					}
					else
					{
						R_Str_Arr = RightRulerStr.split('(');
						R_Str_Str = R_Str_Arr[0].replace(/^\s+/,'').replace(/\s+$/,'');
						R_Str_End = R_Str_Arr[1].replace(/^\s+/,'').replace(/\s+$/,'');
						Ruler_Str = R_Str_Str + '(' + RightTable_Name + '.' + R_Str_End;
					}
				}
				else
				{
					Ruler_Str = RightRulerStr;
				}	
				if (Sql_Str.indexOf('Where') == -1)
				{
					if (Sql_Str.indexOf('Order By') == -1)
					{
						switch (Right_Ruler)
						{
							case "In":
								Where_Str = Ruler_Str + ' In(' + Right_Ruler_Txt + ')';
								break;	
							case "Like":
								Where_Str = Ruler_Str + ' Like \'%' + Right_Ruler_Txt + '%\'';
								break;
							default:
								Where_Str = Ruler_Str + ' ' + Right_Ruler + ' ' + Right_Ruler_Txt;
								break;
						}
						document.List_Form.DisSql.value = Sql_Str + ' Where ' + Where_Str;
					}
					else
					{
						Sql_Split_Arr = Sql_Str.split('Order By');
						Sql_Start_Str = Sql_Split_Arr[0].replace(/^\s+/,'').replace(/\s+$/,'');
						Sql_End_Str = Sql_Split_Arr[1].replace(/^\s+/,'').replace(/\s+$/,'');
						switch (Right_Ruler)
						{
							case "In":
								Where_Str = Ruler_Str + ' In(' + Right_Ruler_Txt + ')';
								break;	
							case "Like":
								Where_Str = Ruler_Str + ' Like \'%' + Right_Ruler_Txt + '%\'';
								break;
							default:
								Where_Str = Ruler_Str + ' ' + Right_Ruler + ' ' + Right_Ruler_Txt;
								break;
						}
						document.List_Form.DisSql.value = Sql_Start_Str + ' Where ' + Where_Str + ' Order By ' + Sql_End_Str;
					}
				}
				else
				{
					Sql_Split_Arr = Sql_Str.split('Where');
					Sql_Start_Str = Sql_Split_Arr[0].replace(/^\s+/,'').replace(/\s+$/,'');
					Sql_End_Str = Sql_Split_Arr[1].replace(/^\s+/,'').replace(/\s+$/,'');
					switch (Right_Ruler)
					{
						case "In":
							Where_No_And = Ruler_Str + ' In(' + Right_Ruler_Txt + ')';
							break;	
						case "Like":
							Where_No_And = Ruler_Str + ' Like \'%' + Right_Ruler_Txt + '%\'';
							break;
						default:
							Where_No_And = Ruler_Str + ' ' + Right_Ruler + ' ' + Right_Ruler_Txt;
							break;
					}
					if (Sql_End_Str.indexOf(Where_No_And) == -1)
					{
						Where_Str = Where_No_And + ' And ';
						document.List_Form.DisSql.value = Sql_Start_Str + ' Where ' + Where_Str + Sql_End_Str;
					}
					else
					{
						if (Sql_End_Str.indexOf('Order By') == -1)
						{
							split_order_str = Sql_End_Str
							Order_End_Str = '';
						}
						else
						{
							split_order_Arr = Sql_End_Str.split('Order By');
							split_order_str = split_order_Arr[0].replace(/^\s+/,'').replace(/\s+$/,'');
							Order_End_Str = ' Order By ' + split_order_Arr[1].replace(/^\s+/,'').replace(/\s+$/,'');
						}
						Replace_Sql_Str = split_order_str.replace(Where_No_And + ' And ','').replace(' And ' + Where_No_And,'').replace(Where_No_And,'').replace(/^\s+/,'').replace(/\s+$/,'');
						if (Replace_Sql_Str != '')
						{
							document.List_Form.DisSql.value = Sql_Start_Str + ' Where ' + Replace_Sql_Str + Order_End_Str;
						}
						else
						{
							document.List_Form.DisSql.value = Sql_Start_Str + Order_End_Str;
						}	
					}
				}
			}
			else
			{
				return;
			}
		}
	}
	else
	{
		return;
	}
}

//设置查询数量
function SetNumToSql()
{
	var Sql_Str = document.List_Form.DisSql.value;
	var Num_Str = document.List_Form.SelectNum.value;
	var Sql_Arr = '';
	var Replace_Str = '';
	var Top_Str = '';
	var End_Str = '';
	if (Sql_Str == '')
	{
		return;
	}
	else
	{
		if (Num_Str == '')
		{
			alert('查询数量请为正整数');
			document.List_Form.SelectNum.focus();
			return;
		}
		Sql_Arr = Sql_Str.split(' ');
		Top_Str = Sql_Arr[0] + ' ' + Sql_Arr[1] + ' ' + Sql_Arr[2] + ' ';
		End_Str = Sql_Str.replace(Top_Str,'');
		Replace_Str = Top_Str.replace(Sql_Arr[2],Num_Str);
		document.List_Form.DisSql.value = Replace_Str + End_Str;
	}
}

//检测sql语句
function CheckSql()
{
	var Sql_Str = document.List_Form.DisSql.value.toLowerCase();
	if (Sql_Str == '')
	{
		return;
	}
	var PageUrl = 'CheckSql.asp';
	var action = '?Act=' + Sql_Str;
	var myAjax = new Ajax.Request(
		PageUrl + action,
		{method:'get',
		parameters:'',
		onComplete:Check_Recive
		}
		);
}
function Check_Recive(XmlObj)
{
	$('CheckSql_Txt').innerHTML=XmlObj.responseText;
}



//显示和隐藏栏目对照表
function Dis_Data_Table(str)
{
  if(str.style.display=="none")
  {
     str.style.display="";
  } 
  else
  {
     str.style.display="none"; 
  }
}


//提交表单
function SubmitFun()
{
	var Name_Str = document.List_Form.LabelName.value;
	if (Name_Str == '')
	{
		alert('标签名称不能为空');
		document.List_Form.LabelName.focus();
		return;
	}
	else
	{
		if (Name_Str.length > 20)
		{
			alert('标签名称不能超过20个字符');
			document.List_Form.LabelName.focus();
			return;
		}
	}
	document.List_Form.submit();
}
-->
</script>







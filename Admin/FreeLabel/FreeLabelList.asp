<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
'ҳ�����ü�Ȩ���ж�
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,strShowErr
MF_Default_Conn
MF_Session_TF 
If Not MF_Check_Pop_TF("MF_sPublic") Then Err_Show
'��ҳ��������
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage
int_RPP = 20 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_ = 8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_ = "<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_ = " <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_ = " <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_ = " <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_ = " <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_ = "<font face=webdings title=""���һҳ"">:</font>"

'===========================================
'����ɾ����¼
Dim ActionStr,DelID
ActionStr = Request.QueryString("Act")
If ActionStr = "Del" Then
	DelID = Request.QueryString("LableID")
	IF DelID = "" Then
		Response.Write "<script>alert('û��ѡ����Ҫɾ���ļ�¼');</script>"
		Response.End
	Else
		Conn.ExeCute("Delete From FS_MF_FreeLabel Where LabelID = '" & NoSqlHack(DelID) & "'")
	End If
End If

'����ɾ��
Dim AllDelIDStr,IDStr,i
If Request.Form("Action") = "del" Then
	AllDelIDStr = NoSqlHack(Request.Form("FreeID"))
	If AllDelIDStr = "" Then
		Response.Write "<script>alert('û��ѡ����Ҫɾ���ļ�¼');</script>"
		Response.End
	Else
		AllDelIDStr = Replace(Replace(AllDelIDStr," ,",","),", ",",")
		If Right(AllDelIDStr,1) = "," Then
			AllDelIDStr = Left(AllDelIDStr,Len(AllDelIDStr) - 1)
		End If
	End If
	If Instr(AllDelIDStr,",") > 0 Then
		IDStr = "'" & Replace(AllDelIDStr,",","','") & "'"
	Else
		IDStr = "'" & AllDelIDStr & "'"
	End If
	AllDelIDStr = "'" & Replace(AllDelIDStr,",","','") & "'"
	if AllDelIDStr <> "" then Conn.ExeCute("Delete From FS_MF_FreeLabel Where LabelID In(" & AllDelIDStr & ")")	
	Response.Write "<script>alert('ɾ���ɹ�');parent.location.reload();</script>"	
End If
%>
<html>
<head>
<title>���ɱ�ǩ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<body>
<table width="98%" height="66" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="hback" > 
    <td width="100%" height="20"  align="Left" class="xingmu">��ǩ��</td>
  </tr>
  <tr class="hback" > 
     <td height="27" align="center" class="hback"><div align="left"><a href="../Label/All_Label_Stock.asp">���б�ǩ</a>��<a href="FreeLabelList.asp"><font color="#FF0000">���ɱ�ǩ</font></a>��<a href="../Label/All_Label_Stock.asp?isDel=1">���ݿ�</a>��<a href="../Label/label_creat.asp">������ǩ</a>��<a href="../Label/Label_Class.asp" target="_self">��ǩ����</a>&nbsp;��<a href="../Label/All_label_style.asp">��ʽ����</a>&nbsp;<a href="../../help?Label=MF_Label_Stock" target="_blank" style="cursor:help;"><img src="../Images/_help.gif" width="50" height="17" border="0"></a></div></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr>
    <td width="10%" class="xingmu"><div align="center">���</div></td>
    <td width="30%" class="xingmu"><div align="center">��ǩ����</div></td>
    <td width="15%" class="xingmu"><div align="center">��ϵͳ</div></td>
    <td width="35%" class="xingmu"><div align="center">����</div></td>
	<td width="10%" class="xingmu"><div align="center">ѡ��</div></td>
  </tr>
  <tr class="hback_1">
	<td colspan="5" height="2"></td>
  </tr>
 <form name="FreeListForm" id="FreeListForm" method="post" action="" style="margin:0px;" target="TempFream"> 
<%
'ȡ�������б�
Dim FreeRs,FreeSql,FreeListNum,HaveVTF
Set FreeRs = Server.CreateObject(G_FS_RS)
FreeSql = "Select LabelID,LabelName,LabelSQl,NSFields,NCFields,LabelContent,selectNum,DesCon,SysType From FS_MF_FreeLabel Where ID > 0 Order By ID Desc"
FreeRs.Open FreeSql,Conn,1,1
IF FreeRs.Eof Then
	HaveVTF = False
%>  
  <tr class="hback_1">
	<td colspan="5" height="20">��ʱ������.</td>
  </tr>
<%
Else
	HaveVTF = True
	FreeRs.PageSize = int_RPP
	cPageNo = NoSqlHack(Request.QueryString("Page"))
	If cPageNo = "" Then cPageNo = 1
	If Not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo > FreeRs.PageCount Then cPageNo = FreeRs.PageCount 
	FreeRs.AbsolutePage = cPageNo
	For FreeListNum = 1 To FreeRs.pagesize
	If FreeRs.Eof Then Exit For
%>  
  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
    <td class="hback"><div align="center"><% = FreeListNum %></div></td>
    <td class="hback"><div align="left"><span style="cursor:hand;" onClick="opencat(lable_<% = FreeRs(0) %>)"><% = FreeRs(1) %></span></div></td>
    <td class="hback"><div align="center">
	<%
		If FreeRs(8) = "NS" Then
			Response.Write "����"
		ElseiF FreeRs(8) = "DS" Then
			Response.Write "����"
		ElseIf FreeRs(8) = "MS" Then
			Response.Write "�̳�"
		End If			
	%>
	</div></td>
	<td class="hback"><div align="center"><a href="AddFreeOne.asp?Act=Edit&LableID=<% = FreeRs(0) %>">�޸�</a>��<a href="FreeLabelList.asp?Act=Del&LableID=<% = FreeRs(0) %>" onClick="{if(confirm('ȷ��ɾ���˼�¼��')){return true;}return false;}">ɾ��</a>��<span onClick="Dis_Style('../<%=G_ADMIN_DIR%>/FreeLabel/FlableDIsStyle.asp?ConStr=<% = Server.URLEncode(FreeRs(5)) %>',500,350,'obj');" style="cursor:hand;">�鿴��ʽ</span></div></td>
    <td class="hback"><div align="center"><input name="FreeID" type="checkbox" id="FreeID" value="<% = FreeRs(0) %>"></div></td>
  </tr>
  <tr id="lable_<% = FreeRs(0) %>" style="display:none;">
    <td colspan="5" class="hback" style="WORD-BREAK: break-all; TABLE-LAYOUT: fixed;">
	<div style="width:100%; line-height:20px; text-align:left;">
		<span style="width:100%; line-height:20px; text-align:left;">��ǩ������</span><br />
		<span style="width:100%; line-height:20px; text-align:left;"><% = Server.HTMLEncode(FreeRs(7)) %></span><br />
		<span style="width:100%; line-height:20px; text-align:left;">��ѯ��䣺</span><br />
		<span style="width:100%; line-height:20px; text-align:left;"><% = Server.HTMLEncode(Replace(FreeRs(2),"*",",")) %></span>
	</div>
	</td>
  </tr>
<%
	FreeRs.MoveNext
	Next
End If	
%>  
   <tr>
     <td height="26" colspan="5" class="hback" valign="middle"><div style="height:25px;"><span style="width:50%; height:25px; line-height:25px; text-align:left; float:left;"><input type="button" name="AddNew" value="�½�" onClick="JavaScript:location.href='AddFreeOne.asp?Act=Add'"></span>
	 <span style="width:48%; height:25px; line-height:25px; text-align:right; float:left;"> <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
	 ѡ������
	 <input type="button" name="Submit" value="ɾ��"  onClick="document.FreeListForm.Action.value='del';if(confirm('ȷ���������ѡ��ļ�¼��')){this.document.FreeListForm.submit();}"></span>
	 </div></td>
   </tr>
   <input name="Action" type="hidden" id="Action" value="">
  </form>
</table>
<% If HaveVTF = True Then %>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">  
   <tr class="hback">
    <td height="30" align="right" valign="middle">
<% response.Write fPageCount(FreeRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo) %>
    </td>
   </tr>
</table>
<% End If %>
<iframe name="TempFream" id="TempFream" src="" width="0" height="0" frameborder="0"></iframe>
</body>
<% 
FreeRs.Close : Set FreeRs = Nothing
Conn.Close : Set Conn=nothing
%>
</html>
<script language="JavaScript" type="text/JavaScript">
function opencat(cat)
{
  if(cat.style.display=="none")
  {
     cat.style.display="";
  } 
  else
  {
     cat.style.display="none"; 
  }
}
function CheckAll(form)  
{  
  	for (var i=0;i<form.elements.length;i++)  
	{  
		var e = FreeListForm.elements[i];  
		if (e.name != 'chkall')
		{
			e.checked = FreeListForm.chkall.checked;  
		}  
	}  
}

function Dis_Style(URL,widthe,heighte,obj)
{
  var obj=window.OpenWindowAndSetValue("../../Fs_Inc/convert.htm?"+URL,widthe,heighte,'window',obj)
  if (obj==undefined)return false;
}
</script>
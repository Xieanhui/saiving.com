<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="Cls_ads.asp"-->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("AS_site") then Err_Show
if not MF_Check_Pop_TF("AS001") then Err_Show
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_
Dim Page,cPageNo

int_RPP=5 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"			'βҳ

Dim Ad_OP,ID,strShowErr,CheckAllID,i,TempID,str_ClassName,str_ClassMode,o_ClassInfo_Rs,str_ClassInfo,str_AdRemarks,Temp_C_ID
Dim o_Ad_Rs,str_Ad_Sql,str_TempAType,str_Temp_Lock,str_date,str_ClassType,str_Classtype_Sql,f_AdclassFlag,MyFile_Obj,G_Ads_FILES_DIR
G_Ads_FILES_DIR="/ads"
Dim Str_SysDir
Str_SysDir=""
if G_VIRTUAL_ROOT_DIR<>"" then
	Str_SysDir="/"&G_VIRTUAL_ROOT_DIR&"/Ads"
Else
	Str_SysDir="/Ads"
end if
str_ClassType=NoSqlHack(Request.queryString("ClassID"))
Ad_OP=Request.QueryString("Ad_OP")
ID=Request.QueryString("ID")
CheckAllID=Request.Form("Checkallbox")
Select Case Ad_OP
	Case "AdLock"
			Conn.execute("update FS_AD_Info set AdLock=1 where AdID="&CintStr(ID)&"")
			UpdateAdsJsContent CintStr(ID)
	Case "AdUnLock"
		Set f_AdclassFlag=Conn.execute("Select Lock From FS_AD_Class Where AdClassID=(Select AdClassID from FS_AD_Info where AdID="&CintStr(ID)&")")
		If Not f_AdclassFlag.Eof Then
			If f_AdclassFlag("Lock")=1 Then
				strShowErr = "<li>������Ŀ�ѱ�����,Ҫ������˹��,���Ƚ����˹��������Ŀ</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			Else
				Conn.execute("update FS_AD_Info set AdLock=0 where AdID="&CintStr(ID)&"")
				UpdateAdsJsContent CintStr(ID)
			End If
		Else
			Conn.execute("update FS_AD_Info set AdLock=0 where AdID="&CintStr(ID)&"")
			UpdateAdsJsContent CintStr(ID)
		End If
		Set f_AdclassFlag=Nothing
	Case "AdDelOne"
		Conn.execute("delete  from FS_AD_Info where AdID="&CintStr(ID)&"")
		Conn.execute("delete  from FS_AD_TxtInfo where AdID="&CintStr(ID)&"")
		Set MyFile_Obj=Server.CreateObject(G_FS_FSO)
		on error resume next
		If MyFile_Obj.FileExists(Server.MapPath(G_Ads_FILES_DIR)&"\"& CintStr(ID) &".js") then
			MyFile_Obj.DeleteFile(Server.MapPath(G_Ads_FILES_DIR)&"\"& CintStr(ID) &".js")
		End If
		Set MyFile_Obj=Nothing
		strShowErr = "<li>ɾ���ɹ�!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Manage.asp?Page="&Request.QueryString("Page")&"")
		Response.end
	Case "P_Lock"
		CheckAllIDFLag("����")
		Conn.execute("update FS_AD_Info set AdLock=1 where AdID in (" & FormatIntArr(CheckAllID) & ")")
		Dim Ads_p_Lock_i,Ads_p_Lock_Arr
		If DelHeadAndEndDot(CheckAllID) <> "" Then
			Ads_p_Lock_Arr = Split(DelHeadAndEndDot(CheckAllID),",")
			For Ads_p_Lock_i = LBound(Ads_p_Lock_Arr) To UBound(Ads_p_Lock_Arr)
				If Ads_p_Lock_Arr(Ads_p_Lock_i) <> "" And IsNumeric(Ads_p_Lock_Arr(Ads_p_Lock_i)) Then
					UpdateAdsJsContent Clng(Ads_p_Lock_Arr(Ads_p_Lock_i))
				End if	
			Next
		End If
		strShowErr = "<li>���������ɹ�!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Manage.asp?Page="&NoSqlHack(Request.QueryString("Page"))&"")
		Response.end
	Case "P_UnLock"
		Dim YesLockNum,temp_num
		YesLockNum=0
		CheckAllIDFLag("����")
		Temp_C_ID=Replace(CheckAllID," ","")
		CheckAllID=split(Temp_C_ID,",")
		For i=0 to Ubound(CheckAllID)
			Set f_AdclassFlag=Conn.execute("Select Lock From FS_AD_Class Where AdClassID In (Select AdClassID from FS_AD_Info where AdID = "&CintStr(CheckAllID(i))&")")
			If Not f_AdclassFlag.Eof Then
				If f_AdclassFlag("Lock")=0 Then
					YesLockNum=YesLockNum+1
					Conn.execute("update FS_AD_Info set AdLock=0 where AdID = "&CintStr(CheckAllID(i))&"")
					UpdateAdsJsContent CintStr(CheckAllID(i))
				End If
			Else
				YesLockNum=YesLockNum+1
				Conn.execute("update FS_AD_Info set AdLock=0 where AdID = "&CintStr(CheckAllID(i))&"")
				UpdateAdsJsContent CintStr(CheckAllID(i))
			End If
			Set f_AdclassFlag=Nothing
		Next
		If YesLockNum=0 Then
			temp_num=0
		Else
			temp_num=YesLockNum-1
		End if
		strShowErr = "<li>���������ɹ�!</li>"
		strShowErr = strShowErr&"<li>��ѡ��<font color=red>"&Ubound(CheckAllID)&"</font>�����</li>"
		strShowErr = strShowErr&"<li>������<font color=red>"&temp_num&"</font>�����</li>"
		If Ubound(CheckAllID)-YesLockNum+1<>0 Then
			strShowErr = strShowErr&"<li><font color=red>"&Ubound(CheckAllID)-YesLockNum+1&"</font>��������ʧ��,ԭ�򣺴˹����Ŀ������</li>"
		End If
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Manage.asp?Page="&NoSqlHack(Request.QueryString("Page"))&"")
		Response.end
	Case "P_Del"
		on error resume next
		CheckAllIDFLag("ɾ��")		
		Set MyFile_Obj=Server.CreateObject(G_FS_FSO)
		Temp_C_ID=CheckAllID
		CheckAllID=split(CheckAllID,",")
		For i=0 to Ubound(CheckAllID)
			TempID=TempID&CheckAllID(i)&","		
			If MyFile_Obj.FileExists(Server.MapPath(G_Ads_FILES_DIR)&"\"& CheckAllID(i) &".js") then
				MyFile_Obj.DeleteFile(Server.MapPath(G_Ads_FILES_DIR)&"\"& CheckAllID(i) &".js")
			End If
		Next
		'Response.Write("delete  from FS_AD_Info where AdID in ("&DelHeadAndEndDot(Temp_C_ID)&")")
		'Response.End()
		Conn.execute("delete  from FS_AD_Info where AdID in (" & FormatIntArr(Temp_C_ID) & ")")
		Conn.execute("delete  from FS_AD_TxtInfo where AdID in (" & FormatIntArr(Temp_C_ID) & ")")
		Set MyFile_Obj=Nothing
		strShowErr = "<li>����ɾ���ɹ�!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Manage.asp?Page="&NoSqlHack(Request.QueryString("Page"))&"")
		Response.end
End Select
Sub CheckAllIDFLag(Showstr)
	If CheckAllID="" or IsNull(CheckAllID) Then
		strShowErr = "<li>��ѡ��Ҫ"&Showstr&"���ļ�!</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If
End Sub
if str_ClassType="" Then str_ClassType=0
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script src="Public.js" language="JavaScript"></script>
</head>
<body>
<%
	Dim str_AdClass_Sql,o_AdClass_Rs,str_AdClass_str,str_Selected
	str_AdClass_Sql="Select AdClassID,AdClassName,Lock from FS_AD_Class"
	Set o_AdClass_Rs=Conn.execute(str_AdClass_Sql)

	If Not o_AdClass_Rs.Eof Then
		While Not o_AdClass_Rs.Eof
			str_Selected=""
			If CintStr(str_ClassType)=CintStr(o_AdClass_Rs("AdClassID")) Then
				str_Selected=" selected"
			End If
			str_AdClass_str=str_AdClass_str&"<option value="&o_AdClass_Rs("AdClassID")&str_Selected&">"&o_AdClass_Rs("AdClassName")&"</option>"			
		o_AdClass_Rs.MoveNext
		Wend			
	End If
	Set o_AdClass_Rs=Nothing	
%>
<form name="AdInfo" action="" method="post">
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu">������</td>
  </tr>
  <tr> 
      <td width="100%" height="18" class="hback"><div algin="Center"><a href="Ads_Add.asp?OpType=Add">��ӹ��</a> | <a href="javascript:P_Lock();">��������</a> | <a href="javascript:P_UnLock();">��������</a> | <a href="javascript:P_Del();">����ɾ��</a> | 
        <select name="ClassID" size="1" style="width:150px" onChange="javascript:location='Ads_Manage.asp?ClassID='+this.value;">
		  <option value="-1">�鿴������Ŀ</option>
		  <%=str_AdClass_str%>
        </select>
        </div></td>
  </tr>
</table>
<%

If Cint(str_ClassType) >0 Then
	str_Classtype_Sql="Where AdClassID="&CintStr(str_ClassType)&""
Else
	str_Classtype_Sql=""
End If
str_Ad_Sql="Select AdID,AdName,AdType,AdAddDate,AdMaxClickNum,AdMaxShowNum,AdLock,AdRemarks,AdClassID from FS_AD_Info "&str_Classtype_Sql&" order by AdID Desc"
Set o_Ad_Rs= CreateObject(G_FS_RS)
o_Ad_Rs.Open str_Ad_Sql,Conn,1,1
If Not o_Ad_Rs.Eof Then
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback">
    <td width="20%" align="center" class="xingmu" height="43">�������</td>
    <td width="12%" align="center" class="xingmu" height="43">�������</td>
    <td width="8%" align="center" class="xingmu" height="43">���ʱ��</td>
    <td width="10%" align="center" class="xingmu" title="���������" height="43">�������</td>
    <td width="10%" align="center" class="xingmu" title="�����ʾ����" height="43">�����ʾ��</td>
    <td width="6%" align="center" class="xingmu" height="43">״̬</td>
    <td width="34%" align="center" class="xingmu" height="43">����</td>
  </tr>
<%
	o_Ad_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>o_Ad_Rs.PageCount Then cPageNo=o_Ad_Rs.PageCount 
	o_Ad_Rs.AbsolutePage=cPageNo

	For int_Start=1 TO int_RPP  
		Select Case Clng(o_Ad_Rs("AdType"))
			Case 0
				str_TempAType="��ͨ��ʾ���"
			Case 1
				str_TempAType="�����´���"
			Case 2
				str_TempAType="���´���"
			Case 3
				str_TempAType="������ʧ"
			Case 4
				str_TempAType="��ҳ�Ի���"
			Case 5
				str_TempAType="͸���Ի���"
			Case 6
				str_TempAType="��������"
			Case 7
				str_TempAType="���µ׶�"
			Case 8
				str_TempAType="���µ׶�"
			Case 9
				str_TempAType="�������"
			Case 10
				str_TempAType="ѭ�����"
			Case 11
				str_TempAType="���ֹ��"
		End Select
		Select Case Clng(o_Ad_Rs("AdLock"))
			Case 0
				str_Temp_Lock="����"
			Case 1
				str_Temp_Lock="<font color=""red"">����</font>"
		End Select
%>
  <tr class="hback">
    <td width="20%" align="center" class="hback" onClick="javascript:ShowClassInfo('<%=o_Ad_Rs("AdID")%>')" title="�����ʾ��Ŀ��Ϣ"><%=o_Ad_Rs("AdName")%></td>
    <td width="12%" align="center" class="hback" onClick="javascript:ShowClassInfo('<%=o_Ad_Rs("AdID")%>')" title="�����ʾ��Ŀ��Ϣ"><%=str_TempAType%></td>
	<%
		str_date = Year(o_Ad_Rs("AdAddDate"))&"-"&Month(o_Ad_Rs("AdAddDate"))&"-"&Day(o_Ad_Rs("AdAddDate"))
	%>
    <td width="8%" align="center" class="hback" onClick="javascript:ShowClassInfo('<%=o_Ad_Rs("AdID")%>')" title="�����ʾ��Ŀ��Ϣ"><%=str_date%></td>
    <td width="10%" align="center" class="hback" onClick="javascript:ShowClassInfo('<%=o_Ad_Rs("AdID")%>')" title="�����ʾ��Ŀ��Ϣ"><%=o_Ad_Rs("AdMaxClickNum")%></td>
    <td width="10%" align="center" class="hback" onClick="javascript:ShowClassInfo('<%=o_Ad_Rs("AdID")%>')" title="�����ʾ��Ŀ��Ϣ"><%=o_Ad_Rs("AdMaxShowNum")%></td>
    <td width="6%" align="center" class="hback" onClick="javascript:ShowClassInfo('<%=o_Ad_Rs("AdID")%>')" title="�����ʾ��Ŀ��Ϣ"><%=str_Temp_Lock%></td>
    <td width="34%" align="center" class="hback"><a href="javascript:AdLock('<%=Clng(o_Ad_Rs("AdID"))%>');">����</a> | <a href="javascript:AdUnLock('<%=Clng(o_Ad_Rs("AdID"))%>');">����</a> | <a href="javascript:AdDelOne('<%=Clng(o_Ad_Rs("AdID"))%>');">ɾ��</a> | <a href="javascript:AdUpdate('<%=Clng(o_Ad_Rs("AdID"))%>','<%=o_Ad_Rs("AdClassID")%>','<%=Request.QueryString("Page")%>');">�޸�</a> | <a href="javascript:Use('<%=Clng(o_Ad_Rs("AdID"))%>');">���ô���</a> |
      <input type="checkbox" name="Checkallbox" value="<%=o_Ad_Rs("AdID")%>"></td>
  </tr>
  	<%
  		str_ClassInfo=""
  		str_ClassMode=""
  		str_ClassName=""
  		str_AdRemarks=""
  		If o_Ad_Rs("AdRemarks")="" Or IsNull(o_Ad_Rs("AdRemarks")) Then
  			str_AdRemarks=""
  		Else
  			str_AdRemarks="&nbsp;&nbsp;��汸ע��"&Left(o_Ad_Rs("AdRemarks"),50)
  		End If
  		If Clng(o_Ad_Rs("AdClassID"))=-1 Then
  			str_ClassInfo="�˹����������Ŀ"&str_AdRemarks
  		Else
  			Set o_ClassInfo_Rs=Conn.execute("Select AdClassID,AdClassName,Lock From FS_AD_Class Where AdClassID="&CintStr(o_Ad_Rs("AdClassID"))&"")
  			If Not o_ClassInfo_Rs.Eof Then
  				str_ClassName=o_ClassInfo_Rs("AdClassName")
  				If o_ClassInfo_Rs("Lock")=0 Then
  					str_ClassMode="����"
  				Else
  					str_ClassMode="<font color=""red"">������</font>"
  				End If
  				str_ClassInfo="������Ŀ��"&str_ClassName&"&nbsp;&nbsp;��Ŀ״̬��"&str_ClassMode&""&str_AdRemarks

	  		Else
  				str_ClassInfo="�˹����������Ŀ"&str_AdRemarks
  			End If
  			Set o_ClassInfo_Rs=Nothing
  		End If
  	%>
	<tr id="<%=o_Ad_Rs("AdID")%>" style="display:none" title="���������Ŀ��Ϣ">
    <td width="100%" align="left" class="hback" onClick="javascript:HideClassInfo('<%=o_Ad_Rs("AdID")%>')" colspan="7" ><%=str_ClassInfo%></td>
    </tr>
<%
		o_Ad_Rs.MoveNext
		If o_Ad_Rs.Eof or o_Ad_Rs.Bof Then Exit For
	Next
	Response.Write "<tr><td class=""hback"" colspan=""7"" align=""left"">"&fPageCount(o_Ad_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ȫѡ<input type=""checkbox"" name=""Checkallbox"" onclick=""javascript:CheckAll('Checkallbox');"" value=""0""></td></tr>"
%>
  </table>
<%
	Else
		Response.write"<table width=""98%"" border=0 align=center cellpadding=2 cellspacing=1 class=table><tr class=""hback""><td>��ǰû�й��!</td></tr></table>"
End If
%>	
</form>
</body>
</html>
<script language="javascript">
function AdLock(ID)
{
	ID=parseInt(ID);
	location='?Ad_OP=AdLock&ID='+ID+'&Page=<%=NoSqlHack(Request.QueryString("Page"))%>';
}
function AdUnLock(ID)
{
	ID=parseInt(ID);
	location='?Ad_OP=AdUnLock&ID='+ID+'&Page=<%=NoSqlHack(Request.QueryString("Page"))%>';
}
function AdDelOne(ID)
{
	if(confirm('�˲�����ɾ��ѡ�е����ݣ�\n��ȷ��ɾ����'))
	{
		ID=parseInt(ID);
		location='?Ad_OP=AdDelOne&ID='+ID+'&Page=<%=NoSqlHack(Request.QueryString("Page"))%>';
	}
}
function CheckAll(CheckType)
{
	var checkBoxArray=document.all(CheckType)
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
	if(confirm('�˲���������ѡ�еĹ�棿\n��ȷ��������'))
	{
		document.AdInfo.action="?Ad_OP=P_Lock"+'&Page=<%=NoSqlHack(Request.QueryString("Page"))%>';
		document.AdInfo.submit();
	}
}
function P_UnLock()
{
	if(confirm('�˲������������ѡ�еĹ�棿\n��ȷ�����������'))
	{
		document.AdInfo.action="?Ad_OP=P_UnLock"+'&Page=<%=NoSqlHack(Request.QueryString("Page"))%>';
		document.AdInfo.submit();
	}
}
function P_Del()
{
	if(confirm('�˲�����ɾ��ѡ�еĹ�棿\n��ȷ��ɾ����'))
	{
		document.AdInfo.action="?Ad_OP=P_Del"+'&Page=<%=NoSqlHack(Request.QueryString("Page"))%>';
		document.AdInfo.submit();
	}
}
function AdUpdate(ID,ClassID,Page)
{
	ID=parseInt(ID);
	location='Ads_Add.asp?OpType=Update&ID='+ID+'&AdClassID='+ClassID+'&OpPage='+Page;
}
function Use(ID)
{
	OpenWindow('Ad_UseShow.asp?PageTitle=��ȡ���ô���&ID='+ID+"&rnd="+Math.random(),300,130,window);
}
function ShowClassInfo(TrID)
{
	document.all(TrID).style.display="";
}
function HideClassInfo(TrID)
{
	document.all(TrID).style.display="none";
}
</script>
<%
Set Conn=nothing
%><!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






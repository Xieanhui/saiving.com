<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp"-->
<%
Dim Conn,CollectConn
MF_Default_Conn
MF_Collect_Conn
MF_Session_TF
Dim SiteID,Action,NewsSql,RsNewsObj,CurrPage,AllPageNum,RecordNum,i,SiteName,RsTempObj,AttributeStr,DelID,DelIDArray,str_History
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
if not MF_Check_Pop_TF("CS003") then Err_Show
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

Action = Request("Action")
SiteID = Request("SiteID")
if Action = "Del" then
	DelID = Request("ID")
	if DelID <> "" then
		if DelID = "record" then
			CollectConn.Execute("Delete from FS_News where History=1")
		else
			DelIDArray = Split(DelID,"***")
			for i = LBound(DelIDArray) to UBound(DelIDArray)
				if DelIDArray(i) <> "" then
					CollectConn.Execute("Delete from FS_News where ID=" & CintStr(DelIDArray(i)))
				end if
			Next
		end if
	end if	
end if
CurrPage = NoSqlHack(Request("Page"))
if Request("check")="1" then
	str_History = " and History=1"
elseif Request("check")="0" then
	str_History = " and History=0"
elseif Request("check") = "all" Then
	str_History = ""
Else
	str_History = " and History=0"
end if
NewsSql = "Select * from FS_News where 0=0"&NoSqlHack(str_History)&" Order by ID Desc"
Set RsNewsObj = Server.CreateObject(G_FS_RS)
RsNewsObj.Open NewsSql,CollectConn,1,1
%>
<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>[site] �����̨ -- ��Ѷ���ݹ���ϵͳ FoosunCMS V5.0</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<BODY topmargin="2" leftmargin="2">
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr>
    <td class="hback">ѡ��״̬��<a href="Check.asp?check=0">�ɼ�����</a> -- <a href="Check.asp?check=1">��ʷ��¼</a> -- <a href="Check.asp?check=all">ȫ���ɼ���¼</a> </td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr> 
    <td height="26" nowrap class="xingmu">
      <div align="center"> ����</div></td>
    <td width="15%" height="20" nowrap class="xingmu"> 
      <div align="center">״̬</div></td>
    <td width="15%" height="20" nowrap class="xingmu"> 
      <div align="center">�ɼ�վ��</div></td>
    <td width="15%" height="20" nowrap class="xingmu"> 
      <div align="center">�������</div></td>
    <td width="15%" height="20" nowrap class="xingmu"> 
      <div align="center">����</div></td>
  </tr>
  <%
if Not RsNewsObj.Eof then
	if CurrPage = "" then
		CurrPage = 1
	else
		CurrPage = CInt(CurrPage)
	end if
	RsNewsObj.PageSize = int_RPP
	RecordNum = RsNewsObj.RecordCount
	AllPageNum = RsNewsObj.PageCount
	if CurrPage > AllPageNum then CurrPage = AllPageNum
	RsNewsObj.AbsolutePage = Cint(CurrPage)
	for i = 1 to RsNewsObj.PageSize
		if RsNewsObj.Eof then Exit For
		Set RsTempObj = CollectConn.Execute("Select SiteName from FS_Site where ID=" & RsNewsObj("SiteID"))
		if Not RsTempObj.Eof then
			SiteName = RsTempObj("SiteName")
		else
			SiteName = "δ֪"
		end if
		RsTempObj.Close
		Set RsTempObj = Nothing
%>
  <tr class="hback"> 
    <td height="26" nowrap> 
      <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><input type="checkbox" CName="<% = RsNewsObj("Title") %>" value="<% = RsNewsObj("ID") %>" name="NewsID"></td>
          <td><% = Left(RsNewsObj("Title"),20) %></td>
        </tr>
      </table></td>
    <td nowrap><div align="center"> 
        <%
		AttributeStr = ""
		if RsNewsObj("History") = True then
			AttributeStr = "<font color=""red"">�����</fonr>"
		else
			AttributeStr = "δ���"
		end if
		Response.Write(AttributeStr)
		%>
        </div></td>
    <td nowrap><div align="center">
        <% = SiteName %>
      </div></td>
    <td nowrap><div align="center">
        <% = RsNewsObj("AddDate") %>
      </div></td>
    <td nowrap class="SpanStyle"> <div align="center"><span style="cursor:hand;" onClick="if (confirm('ȷ��Ҫ�޸���')) location='EditNews.asp?NewsIDStr=<% = RsNewsObj("ID") %>';">�޸�</span>&nbsp;&nbsp;<span style="cursor:hand;" onClick="if (confirm('ȷ��Ҫɾ����')) location='?Action=Del&ID=<% = RsNewsObj("ID") %>';">ɾ��</span></div></td>
  </tr>
  <%
		RsNewsObj.MoveNext
	next
%>
  <tr  class="hback"> 
    <td height="30" colspan="5" nowrap>
	  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td> <div align="right">����������
              <input type="checkbox" name="checkAll" value="" onClick="selectAll(document.all.NewsID,this.checked)">
              ȫѡ&nbsp;&nbsp;
			  <input onClick="DeleteNews('record');" type="button" name="Submit" value=" ɾ��ȫ����������� ">
			  &nbsp;&nbsp;
			  <input onClick="MoveNews('all');" type="button" name="Submit" value=" ȫ����� ">
			  &nbsp;&nbsp;
              <input onClick="MoveNews('');" type="button" name="Submit" value=" �� �� ">
			  &nbsp;&nbsp;<input onClick="DeleteNews('');" type="button" name="Submit" value=" ɾ �� ">
              <%
			  Response.Write"<br>"
			Response.Write(fPageCount(RsNewsObj,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,CurrPage))
			Response.Write"<br>"
		%>
            </div></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
%>
</table>
</BODY>
</HTML>
<%
Set CollectConn = Nothing
Set Conn = Nothing
Set RsNewsObj = Nothing
%>
<script language="JavaScript">
function MoveNews(f_All_ID)
{
	var ID_Str='',CName_Str='';
	if (f_All_ID=='')
	{
		if (document.all.NewsID.length)
		{
			for (var i=0;i<document.all.NewsID.length;i++)
			{	
				if(document.all.NewsID(i).checked)
				{
					if (ID_Str!=''){ID_Str=ID_Str+'***'+document.all.NewsID(i).value;}else{ID_Str=document.all.NewsID(i).value;}
					if (CName_Str!=''){CName_Str=CName_Str+'***'+document.all.NewsID(i).CName;}else{CName_Str=document.all.NewsID(i).CName;}
				}
			}
		}
		else{if(document.all.NewsID.checked){ID_Str=document.all.NewsID.value;CName_Str=document.all.NewsID.CName;}}
	}
	else {ID_Str='all';CName_Str='ȫ������';}
	if (ID_Str!='')
	{
		if(confirm('ȷ��Ҫ�����'))location='MoveNews.asp?ID='+ID_Str+'&CName='+CName_Str;
	}
	else alert('��ѡ��Ҫ��������');
}

function DeleteNews(f_ID_Type)
{
	var ID_Str='';
	if (f_ID_Type=='')
	{
		if (document.all.NewsID.length)
		{
			for (var i=0;i<document.all.NewsID.length;i++)
			{	
				if(document.all.NewsID(i).checked)
				{
					if (ID_Str!=''){ID_Str=ID_Str+'***'+document.all.NewsID(i).value;}else{ID_Str=document.all.NewsID(i).value;}
				}
			}
		}
		else{if(document.all.NewsID.checked)ID_Str=document.all.NewsID.value;}
	}
	else
	{
		ID_Str=f_ID_Type
	}
	if (ID_Str!='')
	{
		if(confirm('ȷ��Ҫɾ����')) location='?Action=Del&ID='+ID_Str;
	}
	else alert('��ѡ��Ҫɾ��������');
}

function selectAll(f_OBJ,f_Flag)
{
	if (f_OBJ.length)
	{
		for (var i=0;i<f_OBJ.length;i++)
		{	
			f_OBJ(i).checked=f_Flag;
		}
	}
	else{f_OBJ.checked=f_Flag;}
}
</script>






<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_InterFace/CLS_Foosun.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="lib/cls_js.asp"-->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
Dim Conn,FS_JsObj,Js_Rs,Manner,i
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
'Ȩ���ж�
'Call MF_Check_Pop_TF("NS_Class_000001") 
'�õ���Ա���б� 
MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("NS_Sysjs") then Err_Show
Set FS_JsObj=New Cls_Js
'---------------------------------��ҳ����
int_RPP=15 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"			'βҳ
'-----------------------------------------
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>CMS5.0</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<script src="js/Public.js" language="JavaScript"></script>
<form action="Js_Free_Action.asp?act=delete" method="post" name="JSForm">
  <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
	<tr> 
	  <td class="xingmu" colspan="11"><a href="#" onMouseOver="this.T_BGCOLOR='#404040';this.T_FONTCOLOR='#FFFFFF';return escape('<div align=\'center\'>FoosunCMS5.0<br> Code by ���ϱ� <BR>Copyright (c) 2006 Foosun Inc</div>')" class="sd"><strong>����JS����</strong></a><a href="../../help?Lable=NS_Class_Templet" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0">
	</tr>
	<tr class="hback_1">
	<td align="center" width="30%">����</td>
	<td align="center">����</td>
	<td align="center">��ʽ</td>
	<td align="center" width="20%">��������[����鿴����]</td>
	<td align="center">����</td>
	<td align="center" width="20%">���ʱ��</td>
	<td align="center" width="5%"><input type="checkbox" name="chk_FreeJs" id="chk_FreeJs" value='' onClick="javascript:checkAll()"></td> 
	</tr> 
	<%
		Set Js_Rs=Server.CreateObject(G_FS_RS)
		Js_Rs.open "Select id from FS_NS_FreeJS",Conn,1,1
		If Not Js_Rs.eof then
		'��ҳʹ��-----------------------------------
			Js_Rs.PageSize=int_RPP
			cPageNo=CintStr(Request.QueryString("page"))
			If cPageNo="" Then cPageNo = 1
			If not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo<=0 Then cPageNo=1
			If cPageNo>Js_Rs.PageCount Then cPageNo=Js_Rs.PageCount 
			Js_Rs.AbsolutePage=cPageNo
		End if
		'--------------------------------------
		For i=0 To int_RPP
			If Js_Rs.eof Then Exit For
			Response.Write("<tr class='hback'>"&vbcrlf)
			FS_JsObj.getFreeJsParam(Js_Rs("id"))
			Response.Write("<td align='center'><a href='Js_Free_Add.asp?act=edit&jsid="&FS_JsObj.id&"' title='��������޸�'>"&FS_JsObj.cname&"</a></td>")
			if FS_JsObj.js_type=1 then
				Response.Write("<td align='center'>ͼƬ</td>"&vbcrlf)
			else
				Response.Write("<td align='center'>����</td>"&vbcrlf)
			End if
			Select Case FS_JsObj.manner
				Case 1  Manner="��ʽA"
				Case 2  Manner="��ʽB"
				Case 3  Manner="��ʽC"
				Case 4  Manner="��ʽD"
				Case 5  Manner="��ʽE"
				Case 6  Manner="��ʽA"
				Case 7  Manner="��ʽB"
				Case 8  Manner="��ʽC"
				Case 9  Manner="��ʽD"
				Case 10 Manner="��ʽE"
				Case 11 Manner="��ʽF"
				Case 12 Manner="��ʽG"
				Case 13 Manner="��ʽH"
				Case 14 Manner="��ʽI"
				Case 15 Manner="��ʽJ"
				Case 16 Manner="��ʽK"
				Case 17 Manner="��ʽL"
			End Select
			Response.Write("<td align='center'>"&Manner&"</td>"&vbcrlf)
			Dim NewsNumberRs,NewsNumber
			Set NewsNumberRs=Conn.execute("Select count(ID) from FS_NS_FreeJsFile where jsname='"&FS_JsObj.ename&"'")
			if not NewsNumberRs.eof then
				NewsNumber=NewsNumberRs(0)
			Else
				NewsNumber=0
			End if
			Response.Write("<td align='center'><a href='#' onClick=""newsDetial('"&FS_JsObj.ename&"')"">"&NewsNumber&"(<font color=""red"">"&FS_JsObj.newsNum&"</font>)</a></td>"&vbcrlf)
			Response.Write("<td align='center'><a href=""javascript:getCode('"&FS_JsObj.id&"')"">����</a></td>"&vbcrlf)
			Response.Write("<td align='center'>"&FS_JsObj.addtime&"</td>"&vbcrlf)
			Response.Write("<td align='center'><input type='checkbox' name='chk_FreeJs' value='"&FS_JsObj.id&"' id='chk_FreeJs_"&FS_JsObj.id&"'></td>")
			Response.Write("</tr>"&vbcrlf)
			Js_Rs.movenext
		next
	%>
	<tr class="hback"> 
	<td colspan="11" align="right">
	<button onClick="javascript:location='Js_Free_Add.asp'">����</button> &nbsp;
	<button onClick="javascript:DeleteJs()">ɾ��</button> &nbsp;
	<!--<button onClick="">�鿴����</button>-->&nbsp;
	</td>
	</tr>
	<%
	Response.Write("<tr>"&vbcrlf)
	Response.Write("<td align='right' colspan='11'  class=""hback"">"&fPageCount(js_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td>"&vbcrlf)
	Response.Write("</tr>"&vbcrlf)
	%>
  </table> 
</form>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
</body>
<script  language="JavaScript">  
<!--  
	function DeleteJs()
	{
		var chkJs_Array=document.all.chk_FreeJs;
		var chkTF=false;
		if (chkJs_Array[0]==null)
		{
			return 
		}
		else
		{
			//�ж��Ƿ�ѡ���˼�¼
			for(var i=1;i<chkJs_Array.length;i++)
			{
				if (chkJs_Array[i].checked)
				{
					chkTF=true
				}
			}
		}
		if(chkTF)
		{
			if(confirm("ȷ��Ҫɾ��ѡ�еļ�¼��"))
				document.JSForm.submit();
		}else
		{
			alert("��ѡ����Ҫɾ���Ķ���");
		}
	}
	function checkAll()
	{
		var chkJs_Array=document.all.chk_FreeJs;
		if (chkJs_Array[0]==null)
		{
			return 
		}
		if(chkJs_Array[0].checked)
		{
			for(var i=1;i<chkJs_Array.length;i++)
			{
				chkJs_Array[i].checked=true;
			}
		}else if(!chkJs_Array[0].checked)
		{
			for(var i=1;i<chkJs_Array.length;i++)
			{
				chkJs_Array[i].checked=false;
			}
		}
	}
	function getCode(jsid)
	{
		if (jsid!=""&&!isNaN(jsid))
		{
			OpenWindow('lib/Frame.asp?PageTitle=��ȡJS���ô���&FileName=showJsPath.asp&JsID='+jsid,360,140,window);
		}else
		{
			alert("���ִ�������ϵ�ͷ���Ա��")
		}
	}
	function newsDetial(JsName)
	{
		if (JsName!=""&&JsName!=null)
		{
			OpenWindow('lib/Frame.asp?PageTitle=���������б�&FileName=showJsNews.asp&JsName='+JsName,360,400,window);
		}else
		{
			alert("���ִ�������ϵ�ͷ���Ա��")
		}
	}
//-->  
</script>  
<%
Conn.close
Set Conn=nothing
Set Js_Rs=nothing
%>
</html>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->
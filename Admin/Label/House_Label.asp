<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/HS_Function.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Dim Conn
	MF_Default_Conn
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show
	Dim obj_Lable_style_Rs,Lable_style_List
	Lable_style_List=""
	Set  obj_Lable_style_Rs = server.CreateObject(G_FS_RS)
	obj_Lable_style_Rs.Open "Select ID,StyleName from FS_MF_Labestyle where StyleType='HS' Order by  id desc",Conn,1,3
	do while Not obj_Lable_style_Rs.eof 
		Lable_style_List = Lable_style_List&"<option value="""& obj_Lable_style_Rs(0)&""">"& obj_Lable_style_Rs(1)&"</option>"
		obj_Lable_style_Rs.movenext
	loop
	obj_Lable_style_Rs.close:set obj_Lable_style_Rs = nothing
%>
<html>
<head>
<title>������ǩ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<base target=self>
</head>
<body class="hback">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<table width="98%" height="100" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
  <form  name="form1" method="post">
    <tr class="hback" > 
      <td colspan="3"  align="Left" class="xingmu"><a href="House_Lable.asp" class="sd" target="_self"><strong><font color="#FF0000">������ǩ</font></strong></a>��<a href="All_Lable_style.asp?Lable_Sub=HS&TF=HS" target="_self" class="sd"><strong>��ʽ����</strong></a></td>
      <td  align="Left" class="xingmu"><div align="right">
          <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="�ر�">
        </div></td>
    </tr>
    <tr class="hback"  style="font-family:����" > 
      <td  align="center" class="hback" ><div align="right">��ʾ��ʽ</div></td>
      <td colspan="3" class="hback" > <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select></td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  align="center"><div align="right">��ǩ����</div></td>
      <td colspan="3" class="hback" > <select  name="labelStyle" onChange="ChooseInfoType(this.options[this.selectedIndex].value);">
          <option value="" style="background:#DEDEDE">---�б���----------</option>
          <option value="ClassInfo" selected>������Ŀ�б�</option>
          <!-- <option value="ReadInfo">�������(����ҳ��)</option>-->
          <option value="LastInfo">��������</option>
<!--          <option value="RecInfo">�����Ƽ�</option>
          <option value="HotInfo">�����ȵ�</option>
          <option value="FiltInfo">�����õ�</option>
          <option value="MarInfo">��������</option>-->
          <option value="" style="background:#DEDEDE">---�ռ���----------</option>
          <option value="ClassList">�����ռ������б�</option>
        </select> <span class="tx">������ռ����࣬����ѡ����Ŀ</span> </td>
    </tr>
    <tr class="hback" style="display:none"> 
      <td width="14%"  align="center" class="hback"><div align="right">��ǩ����</div></td>
      <td colspan="3" class="hback" ><input name="labelName"  type="text" size="12" maxlength="25"> 
        <span class="tx">�����Ժ���ұ�ǩ,����25���ַ�(ֻ��Ϊ���ġ����֡�Ӣ�ġ��»��ߺ��л���)��</span></td>
    </tr>
    <tr class="hback" id="ClassName_col"> 
      <td class="hback"  align="center"><div align="right">��Ŀ�б�</div></td>
      <td colspan="3" class="hback" >
	  <select name="ClassID" onChange="setOrderTypeID(this)">
	  	<option value="Quotation">¥����Ϣ</option>
	  	<option value="Second">������Ϣ</option>
	  	<option value="Tenancy">������Ϣ</option>
	  	<option value="ToRent">---������Ϣ</option>
	  	<option value="Rent">---������Ϣ</option>
	  	<option value="ToSell">---������Ϣ</option>
	  	<option value="Sell">---����Ϣ</option>
	  	<option value="AddRent">---������Ϣ</option>
	  	<option value="Transfer">---ת����Ϣ</option>
	  </select>
	  </td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:����;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV����</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID" disabled type="text" id="DivID" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��"> 
        &quot; class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6"  disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid"  disabled type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" id="ulclass" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot; <input name="liid" disabled type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="liclass"  type="text" id="liclass" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  align="center"><div align="right">��ʾ��Χ</div></td>
      <td colspan="3" class="hback" >ֻ��ʾ 
        <input name="DateNumber"  type="text" id="DateNumber" value="0" size="5">
        ���ڵ�����. <span class="tx">���Ϊ0������ʾ����ʱ���ڵ�����</span></td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  id="InfoNum" align="center"><div align="right">��������</div></td>
      <td colspan="3" class="hback" ><input name="TitleNumber" type="text" id="TitleNumber" value="30" size="5"> 
        <span class="tx">������2���ַ�</span> ����ͼ�ı�־ 
        <select name="PicTF" id="PicTF">
          <option value="1" selected>��ʾ</option>
          <option value="0">����ʾ</option>
        </select>
        �����򿪴��� 
        <select name="Openstyle" id="Openstyle">
          <option value="0" selected>ԭ����</option>
          <option value="1">�´���</option>
        </select> </td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  id="InfoNum" align="center"><div align="right">��������</div></td>
      <td width="35%" class="hback" ><input id="InfoNumber"  name="InfoNumber" type="text" size="12" value="10"> 
        <span class="tx">���õ�ǰ̨��ʾ��������</span> �������� </td>
      <td width="15%" class="hback" >��������</td>
      <td width="36%" class="hback" >
	  	  <select name="ColsNumber" id="select">
          <option value="1" selected>1</option>
          <option value="2">2</option>
          <option value="3">3</option>
          <option value="4">4</option>
          <option value="5">5</option>
          <option value="6">6</option>
          <option value="7">7</option>
          <option value="8">8</option>
          <option value="9">9</option>
          <option value="10">10</option>
          <option value="11">11</option>
          <option value="12">12</option>
          <option value="13">13</option>
          <option value="14">14</option>
          <option value="15">15</option>
          <option value="16">16</option>
          <option value="17">17</option>
          <option value="18">18</option>
          <option value="19">19</option>
          <option value="20">20</option>
        </select> <span class="tx">����һ����ʾ����</span></td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  id="ColsNum" align="center"><div align="right">��������</div></td>
      <td colspan="3" class="hback" ><select name="SubTF" id="SubTF">
          <option value="1">��</option>
          <option value="0" selected>��</option>
        </select> <span class="tx">���鲻Ҫѡ�񣬷������������ٶȻ����Ƚ���</span></td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  id="Num" align="center"><div align="right">�����ֶ�</div></td>
      <td class="hback" > <select id="OrderType"  name="OrderType">
          <option value="ID">�Զ����</option>
          <option value="PubDate">�������</option>
        </select></td>
      <td class="hback"  id="InfoType" align="center">����ʽ</td>
      <td class="hback" ><select id="orderby"  name="orderby">
          <option value="ASC">����</option>
          <option value="DESC" selected>����</option>
        </select> </td>
    </tr>
    <tr class="hback" id="More_col"> 
      <td class="hback"  id="PageStyle_1" align="center"><div align="right">��������</div></td>
      <td colspan="3" class="hback" ><input  name="More_char" type="text" id="More_char" value="��" size="20"> 
        <span class="tx">������html�﷨,�磺&lt;img src=&quot;/files/more.gif&quot; border='0'&gt;</span></td>
    </tr>
    <tr class="hback" id="PageStyle_col" style="display:none"> 
      <td class="hback"   align="center"><div align="right">�Ƿ��ҳ</div></td>
      <td colspan="3" class="hback" ><input name="PageTF" type="checkbox" id="PageTF" value="1" disabled>
        �� ��ҳ��ʽ 
        <input id="PageStyle"  name="PageStyle" type="text" size="9" value='3,CC0066'  disabled> 
        <input name="button" type="button" id="SetPage" onClick="OpenPageStyle(this.form.PageStyle)" value="����" disabled>
        ��ÿҳ���� 
        <input id="PageNumber"  name="PageNumber" type="text" size="4" value="30" disabled>
        ��<span id="page_css" style="display:">��ҳCSS 
        <input id="PageCSS" name="PageCSS" type="text" size="5">
        </span> <span class="tx">�������ռ�����</span></td>
    </tr>
    <tr class="hback" id="Mar_cols" style="display:none"> 
      <td class="hback"  id="ScrollSpeed" align="center"><div align="right">�����ٶ�</div></td>
      <td colspan="3" class="hback" ><input id="MarqueeSpeed"  name="MarqueeSpeed" type="text" size="12" value='20'  disabled>
        ���������� 
        <select id="MarqueeDirection"  name="MarqueeDirection" disabled>
          <option value="up">����</option>
          <option value="down">����</option>
          <option value="left" selected>����</option>
          <option value="right">����</option>
        </select> </td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  id="NewsStyle" align="center"><div align="right">���ڸ�ʽ</div></td>
      <td colspan="3" class="hback" ><input name="DateType" type="text" id="DateType" value="YY02��MM��DD��" size="20"> 
        <span class="tx">��ʽ:YY02����2λ�����(��06��ʾ2006��),YY04��ʾ4λ�������(2006)��MM�����£�DD�����գ�HH����Сʱ��MI����֣�SS������</span></td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  align="center"><div align="right">������ʽ</div></td>
      <td class="hback"  colspan="3" align="center"><div align="left"> 
          <select id="InfoStyle"  name="InfoStyle" style="width:40%">
            <% = Lable_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('House_Label_styleread.asp?ID='+document.form1.InfoStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="�鿴">
          <span class="tx">���ڸ�����ϵͳ�н���ǰ̨ҳ��������ʾ��ʽ</span></div></td>
    </tr>
    <tr class="hback" > 
      <td class="hback" align="center" height="30"><div align="right">����(����2���ַ�)</div></td>
      <td height="30" align="left" class="hback" colspan="3">
	  ������ʾ����ַ�<input name="contentnumber" type="text" id="contentnumber" value="200" size="12">
      ������ʾ����ַ�<input name="navinumber" type="text" id="navinumber" value="200" size="12"/>
	  </td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  colspan="4" align="center" height="30"> <input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "> 
      </td>
    </tr>
  </form>
</table>

</body>
<% 
Set Conn=nothing
%>
</html>
<script language="JavaScript" type="text/JavaScript">
function ChooseInfoType(InfoType)
{
	switch (InfoType)
	{
		case "ClassInfo":
			returnChooseType();
			break;
		case "LastInfo":
			returnChooseType();
			break;
		case "HotInfo":
			returnChooseType();
			break;
		case "RecInfo":
			returnChooseType();
			break;
		case "DayInfo":
			returnChooseType();
			break;
		case "BriInfo":
			returnChooseType();
			break;
		case "AnnInfo":
			returnChooseType();
			break;
		case "ConstrInfo":
			returnChooseType();
			break;
		case "FiltInfo":
			returnChooseType();
			break;
		case "MarInfo":
			document.getElementById('PageStyle_col').disabled=true;
			document.getElementById('SetPage').disabled=true;
			document.getElementById('PageTF').disabled=true;
			document.getElementById('PageStyle').disabled=true;
			document.getElementById('PageNumber').disabled=true;
			document.getElementById('Mar_cols').disabled=false;
			document.getElementById('Mar_cols').style.display='';
			document.getElementById('PageStyle_col').style.display='none';
			document.getElementById('More_col').style.display='';
			document.getElementById('More_char').disabled=false;
			document.getElementById('MarqueeSpeed').disabled=false;
			document.getElementById('MarqueeDirection').disabled=false;
			document.getElementById('ColsNumber').value='10';
			document.getElementById('ClassID').disabled=false;
			break;
		case "ClassList":
			document.getElementById('Mar_cols').disabled=true;
			document.getElementById('PageStyle_col').disabled=false;
			document.getElementById('SetPage').disabled=false;
			document.getElementById('PageTF').disabled=false;
			document.getElementById('PageStyle').disabled=false;
			document.getElementById('PageNumber').disabled=false;
			document.getElementById('Mar_cols').disabled='';
			document.getElementById('PageStyle_col').style.display='';
			document.getElementById('More_col').style.display='none';
			document.getElementById('MarqueeSpeed').disabled=true;
			document.getElementById('MarqueeDirection').disabled=true;
			document.getElementById('More_char').disabled=true;
			document.getElementById('ColsNumber').value='1';
			document.getElementById('ClassID').disabled=false;
			document.getElementById('InfoNumber').disabled=true;
			break;
	}
 }
function returnChooseType()
	{
			document.getElementById('PageStyle_col').disabled=true;
			document.getElementById('SetPage').disabled=true;
			document.getElementById('PageTF').disabled=true;
			document.getElementById('PageStyle').disabled=true;
			document.getElementById('PageNumber').disabled=true;
			document.getElementById('Mar_cols').disabled=true;
			document.getElementById('Mar_cols').style.display='none';
			document.getElementById('PageStyle_col').style.display='none';
			document.getElementById('More_col').style.display='';
			document.getElementById('More_char').disabled=false;
			document.getElementById('ClassID').disabled=false;
			document.getElementById('ColsNumber').value='1';
			document.getElementById('MarqueeSpeed').disabled=true;
			document.getElementById('MarqueeDirection').disabled=true;
			document.getElementById('InfoNumber').disabled=false;
	}
function selectHtml_express(Html_express)
{
	switch (Html_express)
	{
	case "out_Table":
		document.getElementById('div_id').style.display='none';
		document.getElementById('li_id').style.display='none';
		document.getElementById('ul_id').style.display='none';
		document.getElementById('DivID').disabled=true;
		document.getElementById('Divclass').disabled=true;
		document.getElementById('ulid').disabled=true;
		document.getElementById('ulclass').disabled=true;
		document.getElementById('liid').disabled=true;
		document.getElementById('liclass').disabled=true;
		document.getElementById('page_css').style.display='';
		break;
	case "out_DIV":
		document.getElementById('div_id').style.display='';
		document.getElementById('li_id').style.display='';
		document.getElementById('ul_id').style.display='';
		document.getElementById('page_css').style.display='none';
		document.getElementById('DivID').disabled=false;
		document.getElementById('Divclass').disabled=false;
		document.getElementById('ulid').disabled=false;
		document.getElementById('ulclass').disabled=false;
		document.getElementById('liid').disabled=false;
		document.getElementById('liclass').disabled=false;
		break;
	}
}
function ShowTempLayer(oDiv){
	var obj=document.getElementById(oDiv);
	obj.style.display = '';
	obj.style.width = 280;
	obj.style.height = 120;
	obj.style.top = parseInt(document.body.scrollTop+event.clientY)-100;
	obj.style.left = parseInt(event.clientX)-280;
}
function OpenPageStyle(obj){
	var ret = OpenWindow("setPage.asp",200,200,"page");
	if(ret!='')obj.value = ret;
}
function GetObj(s){
	return document.getElementById(s);
}
function IsDisabled(oName){
	if(GetObj(oName).disabled == true){
		return true;
	}else{
		return false;
	}
}
var G_CanSave = true;
var G_Save_Msg = '';
function ok(obj){
	//if(obj.labelName.value == ''){alert('�������ǩ����');obj.labelName.focus();return false;}
	//if(obj.labelName.value.length >25){alert('��ǩ��������̫��');obj.labelName.focus();return false;}
	if(obj.labelStyle.value=='MarInfo')
	{
		if(isNaN(obj.MarqueeSpeed.value)==true){alert('�����ٶ�����д����');obj.MarqueeSpeed.focus();return false;}
	}
	//if(obj.out_char.value=='out_DIV')
	//{
	//	if(obj.DivID.value==''){alert('����дDIV��ID');obj.DivID.focus();return false;}
	//}
	if(obj.PageStyle.value == '')obj.PageStyle.value=',CC0066';
	var retV = '{FS:HS=';
	retV+=obj.labelStyle.value + '��';
	if(!IsDisabled('labelName')) retV+='����$' + obj.labelName.value + '��';
	if(!IsDisabled('ClassID')) retV+='��Ŀ$' + obj.ClassID.value + '��';
	if(!IsDisabled('InfoNumber')) retV+='Loop$'+ obj.InfoNumber.value + '��';
	if(!IsDisabled('out_char')) retV+='�����ʽ$' + obj.out_char.value + '��';
	if(!IsDisabled('DivID')) retV+='DivID$' + obj.DivID.value + '��';
	if(!IsDisabled('Divclass')) retV+='DivClass$' + obj.Divclass.value + '��';
	if(!IsDisabled('ulid')) retV+='ulid$' + obj.ulid.value + '��';
	if(!IsDisabled('ulclass')) retV+='ulclass$' + obj.ulclass.value + '��';
	if(!IsDisabled('liid')) retV+='liid$' + obj.liid.value + '��';
	if(!IsDisabled('liclass')) retV+='liclass$' + obj.liclass.value + '��';
	if(!IsDisabled('DateNumber')) retV+='������$' + obj.DateNumber.value + '��';
	if(!IsDisabled('TitleNumber')) retV+='������$' + obj.TitleNumber.value + '��';
	if(!IsDisabled('PicTF')) retV+='ͼ�ı�־$' + obj.PicTF.value + '��';
	if(!IsDisabled('Openstyle')) retV+='�򿪴���$' + obj.Openstyle.value + '��';
	if(!IsDisabled('SubTF')) retV+='��������$' + obj.SubTF.value + '��';
	if(!IsDisabled('OrderType')) retV+='�����ֶ�$' + obj.OrderType.value + '��';
	if(!IsDisabled('orderby')) retV+='���з�ʽ$' + obj.orderby.value + '��';
	if(!IsDisabled('More_char')) retV+='��������$' + obj.More_char.value + '��';
	if(!IsDisabled('PageTF')) retV+='��ҳ$' + obj.PageTF.value + '��';
	if(!IsDisabled('PageStyle')) retV+='��ҳ��ʽ$' + obj.PageStyle.value + '��';
	if(!IsDisabled('PageNumber')) retV+='ÿҳ����$' + obj.PageNumber.value + '��';
	if(!IsDisabled('PageCSS')) retV+='PageCSS$' + obj.PageCSS.value + '��';
	if(!IsDisabled('MarqueeSpeed')) retV+='�����ٶ�$' + obj.MarqueeSpeed.value + '��';
	if(!IsDisabled('MarqueeDirection')) retV+='��������$' + obj.MarqueeDirection.value + '��';
	if(!IsDisabled('DateType')) retV+='���ڸ�ʽ$' + obj.DateType.value + '��';
	if(!IsDisabled('InfoStyle')) retV+='������ʽ$' + obj.InfoStyle.value + '��';
	retV+='������$' + obj.ColsNumber.value +'��';
	retV+='��������$' + obj.contentnumber.value + '��';
	if(obj.labelStyle.value=='MarNews')
	{retV+='��������$' + obj.navinumber.value + '��';}
	else
	{retV+='��������$' + obj.navinumber.value +'';}
	retV+='}';
	window.parent.returnValue = retV;
	window.close();
}
//��id�滻Ϊ��Ӧ�ı������
function setOrderTypeID(Obj)
{
	var Select_OrderTypeID_Item=document.all("OrderType").options[0];
	switch (Obj.value)
	{
		case "Quotation": Select_OrderTypeID_Item.value="ID";break;
		case "Second": Select_OrderTypeID_Item.value="SID";break;
		default: Select_OrderTypeID_Item.value="TID";break;
	}
}
</script>

<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Dim Conn,User_Conn
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show
%>
<html>
<head>
<title>��ǩ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<base target=self>
</head>
<body class="hback">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
  <form  name="form1" method="post">
  <table width="98%" height="29" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
    <tr class="hback" > 
      <td height="27"  align="Left" class="xingmu"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td colspan="2" class="xingmu">������ñ�ǩ</td>
    </tr>
		  <tr> 
            <td width="41%" class="xingmu"><strong>��־��ǩ</strong></td>
            <td width="59%"><div align="right"> 
                <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="�ر�">
            </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="hback"> 
      <td height="50%"><div align="left"><a href="FL_C_Label.asp?type=WordFL" target="_self"></a><a href="News_C_Label.asp?type=OldNews" target="_self"></a>��־��ǩ���ù̶�ģʽ���ȴ��Ժ�����,���Ҫ�޸Ĳ�������ֱ�������ɵı�ǩ������ʾ�޸�<br>
        <a href="Log_c_label.asp?type=0" target="_self">�����ǩ</a> | <a href="Log_c_label.asp?type=1" target="_self">�����ǩ</a> | <a href="Log_c_label.asp?type=2" target="_self">�����־��ǩ</a>  | <a href="Log_c_label.asp?type=3" target="_self">�û��б����</a> </div></td>
    </tr>
  </table>
  <%If Request.QueryString("type")="1" then%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback">
    <td width="21%"><div align="right">ѡ�����</div></td>
    <td width="79%">
	<select name="ClassID" id="ClassID">
	<%
	dim rs
	set rs = User_Conn.execute("select id,ClassName From FS_ME_iLogClass Order by id asc")
	do while not rs.eof
		response.Write"<option value="""&rs("id")&""">"&rs("ClassName")&"</option>"
	rs.movenext
	loop
	rs.close:set rs=nothing
	%>
    </select>    </td>
  </tr>
  <tr class="hback">
    <td><div align="right">����CSS</div></td>
    <td><input name="TitleCSS" type="text" id="TitleCSS"></td>
  </tr>
  <tr class="hback">
    <td><div align="right">��������</div></td>
    <td><input name="CodeNumber" type="text" id="CodeNumber" value="10"></td>
  </tr>
  <tr class="hback">
    <td><div align="right">��������</div></td>
    <td><input name="leftTitle" type="text" id="leftTitle" value="40">
      ����ռ2���ֽ�</td>
  </tr>
    <tr> 
      <td class="hback"><div align="right">�����ʽ</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:����;display:none;" > 
      <td rowspan="3"  align="center" class="hback">
        <div align="right">DIV����</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��"> 
        &quot;class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6" disabled  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid" disabled  type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" disabled id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">        &quot; class=&quot; <input name="liclass" disabled type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
  <tr class="hback">
    <td><div align="right">��ʾ����</div></td>
    <td><select name="DateTF" id="DateTF">
      <option value="1">��ʾ</option>
      <option value="0" selected>����ʾ</option>
    </select>    </td>
  </tr>
  <tr class="hback">
    <td><div align="right">���ڸ�ʽ</div></td>
    <td><input name="DateType" type="text" id="DateType" value="MM��DD��">
      <span class="tx">��ʽ:YY02����2λ�����(��06��ʾ2006��),YY04��ʾ4λ�������(2006)��MM�����£�DD�����գ�HH����Сʱ��MI����֣�SS������</span></td>
  </tr>
  <tr class="hback">
    <td><div align="right"></div></td>
    <td><input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
      <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
  </tr>
</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:ME=ClassList��';
		retV+='��������$' + obj.CodeNumber.value + '��';
		retV+='ClassId$' + obj.ClassID.value + '��';
		retV+='��������$' + obj.leftTitle.value + '��';
		retV+='����CSS$' + obj.TitleCSS.value + '��';
		if(!IsDisabled('DivID')) retV+='DivID$' + obj.DivID.value + '��';
		if(!IsDisabled('Divclass')) retV+='DivClass$' + obj.Divclass.value + '��';
		if(!IsDisabled('ulid')) retV+='ulid$' + obj.ulid.value + '��';
		if(!IsDisabled('ulclass')) retV+='ulclass$' + obj.ulclass.value + '��';
		if(!IsDisabled('liid')) retV+='liid$' + obj.liid.value + '��';
		if(!IsDisabled('liclass')) retV+='liclass$' + obj.liclass.value + '��';
		retV+='��ʾ����$' + obj.DateTF.value + '��';
		retV+='���ڸ�ʽ$' + obj.DateType.value + '��';
		if(!IsDisabled('out_char')) retV+='�����ʽ$' + obj.out_char.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%elseif Request.QueryString("type")="2" then%>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    
    <tr>
      <td colspan="2" class="xingmu">��־���ҳ��ǩ</td>
    </tr>
    <tr>
      <td width="26%" class="hback"><div align="right">ѡ����־��ǩ</div></td>
      <td width="74%" class="hback">
	  	<select name="LogType" id="LogType">
        <option value="Log_title">��־����</option>
        <option value="Log_Content">��־����</option>
        <option value="Log_Author">��־����</option>
        <option value="Log_hits">��־�����</option>
        <option value="Log_keywords">tags(�ؼ���)</option>
        <option value="Log_AddTime">��������</option>
        <option value="Log_LogType">��־����</option>
        <option value="Log_ReviewList">�����б�</option>
        <option value="Log_ReviewForm">�������۱�</option>
      </select>      </td>
    </tr>
	<tr>
      <td class="hback">&nbsp;</td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:ME=LogPage:' + obj.LogType.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%elseif Request.QueryString("type")="3" then%>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    
    <tr>
      <td colspan="2" class="xingmu">�û���ҳ����</td>
    </tr>
    <tr>
      <td width="26%" class="hback"><div align="right">ѡ����־��ǩ</div></td>
      <td width="74%" class="hback">
	  	<select name="LogType" id="LogType">
        <option value="Log_Usertitle">վ������</option>
        <option value="Log_UserName">�û���</option>
        <option value="Log_NickName">�û��ǳ�</option>
        <option value="Log_UserContent">վ������</option>
      </select>      </td>
    </tr>
	<tr>
      <td class="hback">&nbsp;</td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:ME=LogUserPage:' + obj.LogType.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%else%>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">�����ǩ</td>
    </tr>
    <tr>
      <td width="26%" class="hback"><div align="right">ѡ����־��ǩ</div></td>
      <td width="74%" class="hback">
	  	<select name="LogType" id="LogType" onChange="ChooseNewsType(this.options[this.selectedIndex].value);">
        <option value="LastLog" selected>������־</option>
        <option value="TopLog">��־����</option>
        <option value="HotLog">�ȵ���־</option>
        <option value="TopSubject">��־ר��</option>
        <option value="InfoClass">��־����</option>
        <option value="InfoList">��־�б�(�ռ�)</option>
        <option value="Log_LastReview">��������</option>
        <option value="Log_LastForm">���۱�</option>
        <!--<option value="Log_MyInfo">��������</option>-->
        <option value="Log_PublicLog">������־����</option>
        <option value="Log_Search">����</option>
        <option value="Log_Navi">����</option>

      </select>      </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��������,��������</div></td>
      <td class="hback"><input name="TitleNumber" type="text" id="TitleNumber" value="10,40">
        ��ʽ&quot;10,40&quot;,���ð��&quot;,&quot;</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">���ڸ�ʽ</div></td>
      <td class="hback"><input name="DateType" type="text" id="DateType" value="MM-DD">
      <span class="tx">��ʽ:YY02����2λ�����(��06��ʾ2006��),YY04��ʾ4λ�������(2006)��MM�����£�DD�����գ�HH����Сʱ��MI����֣�SS������</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">�����ʽ</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:����;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV����</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��"> 
        &quot;class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6" disabled  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid" disabled  type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" disabled id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">        &quot; class=&quot; <input name="liclass" disabled type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
	<tr>
      <td class="hback">&nbsp;</td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:ME=' + obj.LogType.value + '��';
		if(!IsDisabled('TitleNumber')) retV+='��������,��������$' + obj.TitleNumber.value + '��';
		if(!IsDisabled('DivID')) retV+='DivID$' + obj.DivID.value + '��';
		if(!IsDisabled('Divclass')) retV+='DivClass$' + obj.Divclass.value + '��';
		if(!IsDisabled('ulid')) retV+='ulid$' + obj.ulid.value + '��';
		if(!IsDisabled('ulclass')) retV+='ulclass$' + obj.ulclass.value + '��';
		if(!IsDisabled('liid')) retV+='liid$' + obj.liid.value + '��';
		if(!IsDisabled('liclass')) retV+='liclass$' + obj.liclass.value + '��';
		retV+='CSS$��';
		retV+='��ʾ����$1��';
		if(!IsDisabled('DateType')) retV+='������ʽ$' + obj.DateType.value + '��';
		if(!IsDisabled('out_char')) retV+='�����ʽ$' + obj.out_char.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end if%>
  </form>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function IsDisabled(oName){
	if(GetObj(oName).disabled == true){
		return true;
	}else{
		return false;
	}
}
var G_CanSave = true;
var G_Save_Msg = '';
function GetObj(s){
	return document.getElementById(s);
}
function SelectClass()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('../News/lib/SelectClassFrame.asp',400,300,window);
	try {
		document.getElementById('ClassID').value = ReturnValue[0][0];
		document.getElementById('ClassName').value = ReturnValue[1][0];
	}
	catch (ex) { }
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
		break;
	case "out_DIV":
		document.getElementById('div_id').style.display='';
		document.getElementById('li_id').style.display='';
		document.getElementById('ul_id').style.display='';
		document.getElementById('DivID').disabled=false;
		document.getElementById('Divclass').disabled=false;
		document.getElementById('ulid').disabled=false;
		document.getElementById('ulclass').disabled=false;
		document.getElementById('liid').disabled=false;
		document.getElementById('liclass').disabled=false;
		break;
	}
}
function ChooseNewsType(NewsType)
{
	switch (NewsType)
	{
		case "LastLog":
			document.getElementById('TitleNumber').disabled=false;
			document.getElementById('DateType').disabled=false;
			document.getElementById('out_char').disabled=false;
			break;
		case "TopLog":
			document.getElementById('TitleNumber').disabled=false;
			document.getElementById('DateType').disabled=false;
			document.getElementById('out_char').disabled=false;
			break;
		case "HotLog":
			document.getElementById('TitleNumber').disabled=false;
			document.getElementById('DateType').disabled=false;
			document.getElementById('out_char').disabled=false;
			break;
		case "TopSubject":
			document.getElementById('TitleNumber').disabled=false;
			document.getElementById('DateType').disabled=false;
			document.getElementById('out_char').disabled=false;
			break;
		case "InfoClass":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=false;
			break;
		case "InfoList":
			document.getElementById('TitleNumber').disabled=false;
			document.getElementById('DateType').disabled=false;
			document.getElementById('out_char').disabled=false;
			break;
		case "Log_LastReview":
			document.getElementById('TitleNumber').disabled=false;
			document.getElementById('out_char').disabled=false;
			break;
		case "Log_LastForm":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_MyInfo":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_PublicLog":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_Search":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_Navi":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=false;
			break;
		case "Log_InfoStat":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_PageTitle":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_title":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_Content":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_Author":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_hits":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_keywords":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_AddTime":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_LogType":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_ReviewList":
			document.getElementById('TitleNumber').disabled=false;
			document.getElementById('DateType').disabled=false;
			document.getElementById('out_char').disabled=false;
			break;
		case "Log_ReviewContent":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
	}
 }
</script>







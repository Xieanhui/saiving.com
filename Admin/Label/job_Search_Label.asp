<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/AP_Function.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
%>
<html>
<head>
<title>�˲ű�ǩ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<base target=self>
</head>
<body class="hback">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="javascript" src="../../FS_Inc/prototype.js" type="text/javascript"></script>
<table width="98%" height="100" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
  <form  name="form1" method="post">
    <tr class="hback" > 
      <td colspan="2"  align="Left" class="xingmu"><a href="Job_label.asp" class="sd" target="_self"><strong><font color="#FF0000">������ǩ</font></strong></a>��<a href="All_label_style.asp?label_Sub=AP&TF=AP" target="_self" class="sd"><strong>��ʽ����</strong></a></td>
      <td width="38%"  align="Left" class="xingmu"><div align="right"> 
          <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="�ر�">
        </div></td>
    </tr>
    <tr class="hback"  style="font-family:����" > 
      <td  align="center" class="hback" ><div align="right">������ʽ</div></td>
      <td colspan="2" class="hback" >
	    <select name="SearchType" id="SearchType" onChange="DisTypeFun(this.options[this.selectedIndex].value,'DisTypeTR');">
          <option value="0" selected="selected">һ������</option>
          <option value="1">�߼�����</option>
        </select></td>
    </tr>
	<tr class="hback" id="DisTypeTR" style="display:;"> 
      <td  align="center" class="hback" ><div align="right">��ʾ��ʽ</div></td>
      <td colspan="2" class="hback">
	    <select name="DisType" id="DisType" onChange="DisTypeFun(this.options[this.selectedIndex].value,'RowTR');">
          <option value="1" selected="selected">����</option>
          <option value="0">����</option>
        </select></td>
    </tr>
    <tr class="hback"> 
      <td width="21%"  align="center" class="hback"><div align="right">�ı�����ʽ</div></td>
      <td colspan="2" class="hback" >
	  <input name="TextStyle"  type="text" size="40" maxlength="25" id="TextStyle"> 
      </td>
    </tr>
	 <tr class="hback"> 
      <td width="21%"  align="center" class="hback"><div align="right">ѡ��˵���ʽ</div></td>
      <td colspan="2" class="hback" >
	  <input name="SelectStyle"  type="text" size="40" maxlength="25" id="SelectStyle"> 
      </td>
    </tr>
	</tr>
	 <tr class="hback"> 
      <td width="21%"  align="center" class="hback"><div align="right">��ť��ʽ</div></td>
      <td colspan="2" class="hback" >
	  <input name="ButtonStyle"  type="text" size="40" maxlength="25" id="ButtonStyle"> 
	  ������ͼƬ��ַ
      </td>
    </tr>
	 <tr class="hback" id="RowTR" style="display:none;"> 
      <td width="21%"  align="center" class="hback"><div align="right">�����о�</div></td>
      <td colspan="2" class="hback" >
	   <input name="RowHieght"  type="text" size="40" maxlength="25" id="RowHieght" value="30">
      </td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  colspan="3" align="center" height="30"> <input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� ">      </td>
    </tr>
  </form>
</table>

</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function ok(obj){
	var retV = '{FS:AP=APSearch��';
	retV+='������ʽ$' + obj.SearchType.value + '��';
	retV+='��ʾ��ʽ$' + obj.DisType.value + '��';  
	retV+='�ı�����ʽ$' + obj.TextStyle.value + '��';
	retV+='ѡ��˵���ʽ$' + obj.SelectStyle.value + '��';
	retV+='��ť��ʽ$' + obj.ButtonStyle.value + '��';
	retV+='�о�$' + obj.RowHieght.value;
	retV+='}';
	window.parent.returnValue = retV;
	window.close();
}
function DisTypeFun(DisType,TRID)
{
	if(DisType==1)
	{
		document.getElementById(TRID).style.display='none';
	}
	else
	{
		document.getElementById(TRID).style.display='';
	}
}
</script>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






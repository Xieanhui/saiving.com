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
	Dim Conn
	MF_Default_Conn
	'session�ж�
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show
	Dim str_Lable_AdSql,o_Lable_AdRs,str_Ad_TempStr
	str_Lable_AdSql="Select AdID,AdName From FS_AD_Info Order by AdID Desc"
	Set o_Lable_AdRs=Conn.execute(str_Lable_AdSql)
	If Not o_Lable_AdRs.Eof Then
		str_Ad_TempStr=""
		While Not o_Lable_AdRs.Eof 
			str_Ad_TempStr=str_Ad_TempStr&"<option value="&o_Lable_AdRs("AdID")&">"&o_Lable_AdRs("AdName")&"</option>"
		o_Lable_AdRs.MoveNext
		Wend
	Else
		str_Ad_TempStr="<option value=""ADNameIsNull"">��ǰû�й�棬����ӹ�棡</option>"
	End If
	Set o_Lable_AdRs=Nothing
%>
<html>
<head>
<title>���ű�ǩ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<base target=self>
</head>
<body class="hback">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<table width="98%" height="100" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
  <form  name="form1" method="post">
    <tr class="hback" > 
      <td colspan="2"  align="Left" class="xingmu"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="49%" class="xingmu"><strong>����ǩ����</strong></td>
            <td width="51%"><div align="right"> 
                <input name="Wclose" type="button" onClick="window.returnValue='';window.close();" value="�ر�">
            </div></td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback" > 
      <td width="28%"  align="right" class="hback"><div align="right">ѡ����</div></td>
      <td width="72%"  align="left" class="hback"><select name="AdName" style="width:100%"><%=str_Ad_TempStr%>
      </select>
      </td>
    </tr>
    <tr class="hback" >
      <td  align="right" class="hback"><div align="right">���SPIN ID </div></td>
      <td  align="left" class="hback"><input name="spanid" type="text" id="spanid" value="Ad_HTML_ID" size="15" maxlength="100">
        (�����Ҫ��һ��ҳ����ö��ͶƱ�����ID����Ϊ��ͬ)</td>
    </tr>
    
    
    
    <tr class="hback" > 
      <td class="hback"  colspan="2" align="center" height="30"> <input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� ">      </td>
    </tr>
  </form>
</table>

</body>
<% 
Set Conn=nothing
%>
</html>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{	
		var retV;
		if (obj.AdName.value=="ADNameIsNull")
		{
			retV="";
			window.parent.returnValue = retV;
			window.close();
		}
		retV= '{FS:AS=AdLIST��';
		retV+='���ID$' + obj.AdName.value+'��';
		retV+='SpanId$' + obj.spanid.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






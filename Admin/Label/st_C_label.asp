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
  <form  name="form1" method="post">
  <table width="98%" height="29" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
    <tr class="hback" > 
      <td height="27"  align="Left" class="xingmu"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="41%" class="xingmu"><strong>�����ǩ����</strong></td>
            <td width="59%"><div align="right"> 
                <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="�ر�">
            </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">ͳ��ϵͳ</td>
    </tr>
    <tr>
      <td width="19%" class="hback"><div align="right">ͳ������</div></td>
      <td width="81%" class="hback">
	  	<select name="sstype" id="sstype">
        <option value="0" selected>��ͼ��</option>
        <option value="1">��ͼ��</option>
        <option value="2">����ͳ��</option>
      </select>      </td>
    </tr>
    
    <tr> 
      <td class="hback"><div align="right">·��</div></td>
      <td class="hback"><select name="Path" id="Path">
        <option value="0" selected>���·��</option>
        <option value="1">����·��</option>
      </select>
      </td>
    </tr>
    
	<tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:SS=SSTYPE��';
		retV+='��ʽ$' + obj.sstype.value + '��';
		retV+='·��$' + obj.Path.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
  </form>
</body>
</html>







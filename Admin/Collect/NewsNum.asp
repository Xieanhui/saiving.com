<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ȷ�ϲɼ�����</title>
</head>
<link href="FS_css.css" rel="stylesheet">
<body topmargin="0" bgcolor="buttonface" leftmargin="0">
<div align="center">
  <table width="90%" border="0" cellspacing="3" cellpadding="0">
	<tr align="center"><td>
	<br>��ӭʹ�÷�Ѷ���Ųɼ�ϵͳ<br>
	�����Ƶ���Ȩ�������Ĵ���Ѷ�Ƽ���չ���޹�˾�޹�<br>
	ȷ��Ҫʹ����<br>
	</td></tr>
     <tr align="center"> 
       <td>���βɼ�����������<input type="text" name='PageNum' value=''></td>
    </tr>
    <tr align="center"> 
      <td height="30">
          <input type="button" onClick="InsertScript()" name="Submit2" value=" ȷ �� ">
          <input type="button" onClick="window.returnValue='back';window.close();" name="Submit" value=" ȡ �� ">
     </td>
    </tr>
  </table>
</div>
</body>
</html>
<script language="JavaScript">
function InsertScript()
{
	var PageNum='';
	PageNum=document.all.PageNum.value;
	window.returnValue=PageNum;
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>






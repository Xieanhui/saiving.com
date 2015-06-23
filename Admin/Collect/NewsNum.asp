<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>确认采集新闻</title>
</head>
<link href="FS_css.css" rel="stylesheet">
<body topmargin="0" bgcolor="buttonface" leftmargin="0">
<div align="center">
  <table width="90%" border="0" cellspacing="3" cellpadding="0">
	<tr align="center"><td>
	<br>欢迎使用风讯新闻采集系统<br>
	如果设计到版权问题与四川风讯科技发展有限公司无关<br>
	确定要使用吗？<br>
	</td></tr>
     <tr align="center"> 
       <td>本次采集新闻数量：<input type="text" name='PageNum' value=''></td>
    </tr>
    <tr align="center"> 
      <td height="30">
          <input type="button" onClick="InsertScript()" name="Submit2" value=" 确 定 ">
          <input type="button" onClick="window.returnValue='back';window.close();" name="Submit" value=" 取 消 ">
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






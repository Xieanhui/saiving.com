<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
%>
<html>
<head>
<title>���÷�ҳ��ʽ</title> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../Fs_Inc/PublicJS.js"></script>
<script language="javascript">
function getColorOptions(){
	var color= new Array("00","33","66","99","CC","FF");
	for (var i=0;i<color.length ;i++ )
	{
		for (var j=0;j<color.length ;j++ )
		{
			for (var k=0;k<color.length ;k++ )
			{
				if(k==0&&j==0&&i==0)
				document.write('<option style="background:#'+color[j]+color[k]+color[i]+'" value="'+color[j]+color[k]+color[i]+'" selected>����</option>');
				else
				document.write('<option style="background:#'+color[j]+color[k]+color[i]+'" value="'+color[j]+color[k]+color[i]+'">����</option>');
			}
		}
	}
}
function getChecked(o){
	for (var i=0;i<o.length;i++ )
	{
		if(o[i].checked) return o[i].value;
	}
}
function setIT(obj){
	var retV='';
	retV += '' + getChecked(obj.list2);
	//var pageColor=obj.pageColor.options[obj.pageColor.selectedIndex].value;
	//if(pageColor='undefined')pageColor='CC0099';
	retV += ',' + obj.pageColor.options[obj.pageColor.selectedIndex].value;
	window.returnValue = retV;
	window.close();	
}
</script>
</head>

<body topmargin="0" leftmargin="0" style="overflow-Y:auto;background:menu">
<iframe width="260" height="165" id="colourPalette" src="../selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></iframe>

<table width="100%" border="0" height="100%" cellspacing="1" cellpadding="1">
<form name="form1">
<tr> 
<td valign="top"><table width="100%"  border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td><font color="#000000"> 
              <input name="list2" type="radio" value="1">
              ǰһҳ ��һҳ</font></td>
          </tr>
          <tr> 
            <td><font color="#000000"> 
              <input name="list2" type="radio" value="2">
              ��Nҳ����1ҳ,��2ҳ </font></td>
          </tr>
          <tr> 
            <td><font color="#000000"> 
              <input name="list2" type="radio" value="3" checked>
              ��Nҳ��1 2 3 </font></td>
          </tr>
          <tr> 
            <td><font color="#000000"> 
              <input name="list2" type="radio" value="4">
              <font face=webdings title="��ҳ"> 9</font></font><font color="#000000" face=webdings title="��ʮҳ">7</font><font color="#000000" face=webdings title="��һҳ">3</font><font color="#000000" face=webdings title="��һҳ">4</font><font color="#000000" face=webdings title="��ʮҳ">8</font><font color="#000000" face=webdings title="���һҳ">:</font> 
            </td>
          </tr>
          <tr> 
            <td>
                <input name="list2" type="radio" value="5" checked> <img src="../../sys_images/page.gif" /> ��DIV����
                <br />����DIVCSS��pagecontent����ǰҳ��CSS��currentPageCSS ,ѡ���˴��������ɫ������Ч��
                
            </td>
          </tr>          
         
          <tr> 
            <td><select name="pageColor" size=8>
                <script>getColorOptions();</script>
              </select> </td>
          </tr>
        </table>
</td>
<td width="1" bgcolor="menu"></td>
<td valign="top"><input type="button" value=" ȷ �� " onClick="setIT(this.form);">
  <br>
  <input type="button" value=" ȡ �� " onClick="window.returnValue='';window.close();"></td>
</tr>
</form>
</table>
</body>
</html>






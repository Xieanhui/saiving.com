<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<%
Dim Conn,User_Conn
ConnDataBase
User_ConnDataBase
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title><%=GetUserSystemTitle%>-�����ļ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="keywords" content="��Ѷ,��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="javascript" src="class_liandong.js"></script>
</head>
<body>

<form name="form1" method="post">
	<SELECT NAME="vclass1" ID="vclass1" style="width:120px">
    	<OPTION selected></OPTION>
    </SELECT>
	<SELECT NAME="vclass2" ID="vclass2" style="width:120px">
		<OPTION selected></OPTION>
    </SELECT>
	<SELECT NAME="vclass3" ID="vclass3" style="width:120px">
    	<OPTION selected></OPTION>
    </SELECT>
	<SELECT NAME="vclass4" ID="vclass4" style="width:120px">
    	<OPTION selected></OPTION>
    </SELECT>
</form>


</body>
</html>

<script language="javascript">
//�����˵�---��ҵ���
//���ݸ�ʽ ID������ID������
var array=new Array();
<%dim sql,rs,i
  sql="select VCID,ParentID,vClassName from FS_ME_VocationClass"
  set rs=User_Conn.execute(sql)
  i=0
  do while not rs.eof
%>
array[<%=i%>]=new Array("<%=rs("VCID")%>","<%=rs("ParentID")%>","<%=rs("vClassName")%>"); 
<%
	rs.movenext
	i=i+1
loop
rs.close
%>

var liandong=new CLASS_LIANDONG_YAO(array)
liandong.firstSelectChange("0","vclass1");
liandong.subSelectChange("vclass1","vclass2");
liandong.subSelectChange("vclass2","vclass3");
liandong.subSelectChange("vclass3","vclass4");
</script>








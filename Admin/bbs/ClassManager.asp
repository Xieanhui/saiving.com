<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
'session�ж�
MF_Session_TF 
if not MF_Check_Pop_TF("WS001") then Err_Show
%>
<html>
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<style type="text/css">
<!--
.style1 {font-weight: bold}
.style2 {color: #FF0000}
-->
</style>
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY>
<%
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
	int_RPP=15 '����ÿҳ��ʾ��Ŀ
	int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
	showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
	str_nonLinkColor_="#999999" '����������ɫ
	toF_="<font face=webdings>9</font>"  			'��ҳ 
	toP10_=" <font face=webdings>7</font>"			'��ʮ
	toP1_=" <font face=webdings>3</font>"			'��һ
	toN1_=" <font face=webdings>4</font>"			'��һ
	toN10_=" <font face=webdings>8</font>"			'��ʮ
	toL_="<font face=webdings>:</font>"				'βҳ
Dim ClassSql,i,ClassRs
Set ClassRs=server.createobject(G_FS_RS)
%>
<script language="JavaScript" src="js/CheckJs.js" type="text/JavaScript"></script>
<script language="JavaScript" src="js/Public.js" type="text/javascript"></script>
<script language="javascript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = ClassForm.elements[i];  
    if (e.name != 'chkall')  
       e.checked = ClassForm.chkall.checked;  
    }  
 }
</script>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table"> 
  <tr> 
    <td align="left" colspan="2" class="xingmu">����ϵͳ���������&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="../../help?Lable=Message" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td> 
  </tr>
  <tr>
  	<td><a href="ClassManager.asp">������ҳ</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="ClassAdd.asp">�����Ŀ</a></td>
  </tr>
 </table>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table"> 
<form name="ClassForm" method="post" action="ClassDel.asp?Act=del">
	<tr>
		<td class="xingmu" align="center" colspan="2" width="20%">��Ŀ����</td>
		<td class="xingmu" align="center" width="20%">��Ŀ˵��</td>
		<td class="xingmu" align="center" width="7%">�����</td>
		<td class="xingmu" align="center" width="18%">���ʱ��</td>
		<td class="xingmu" align="center" width="30%">��    ��</td>
	</tr>
	<%
	ClassRs.open "Select ID,ClassID,ClassName,ClassExp,Pid,Author,AddDate from FS_WS_Class order by Pid,id desc",Conn,1,1
	IF not ClassRs.eof THEN 
		ClassRs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("Page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo<=0 Then cPageNo=1
		If cPageNo>ClassRs.PageCount Then cPageNo=ClassRs.PageCount 
		ClassRs.AbsolutePage=cPageNo
		 FOR int_Start=1 TO int_RPP 
		%>
	<tr>
		<td class="hback" align="center" width="4%"><input type="checkbox" name="ID" value="<%=ClassRs("ID")%>"></td>
		<td class="hback" align="center" width="16%">
		<%=left(ClassRs("ClassName"),15)%>		</td>
		<td class="hback" align="center" widht="25%">
		<%=Left(ClassRs("ClassExp"),30)%>		</td>
		<td class="hback" align="center" widht="25%">
		<%=ClassRs("Author")%>		</td>
		<td class="hback" align="center" width="20%">
		<%=ClassRs("AddDate")%>		</td>
		<td class="hback" align="center" width="35%"><a href="ClassEdit.asp?ID=<%=ClassRs("ID")%>">�޸�</a> �� <a href="ClassDel.asp?ID=<%=ClassRs("ID")%>&Act=single" onClick="{if(confirm('ȷ��Ҫɾ����')){return true;}return false;}">ɾ��</a>		</td>
	</tr>
	<%
		ClassRs.MoveNext
		if ClassRs.eof or ClassRs.bof then exit for
      NEXT
	END IF	  
	%>
	<tr> 
      <td colspan="6" class="hback"><table width="100%"><tr><td width="40%" align="center"><input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
          ѡ������ 
          <input name="Submit" type="submit" id="Submit"  onClick="{if(confirm('ȷʵҪ����ɾ����?')){this.document.ClassForm.submit();return true;}return false;}" value=" ɾ�� "></td><%
	 Response.Write("<td class=""hback"" colspan=""6"" align=""right"" width=""60%"">"&fPageCount(ClassRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo) &"</td>")
	Set ClassRs=nothing
	%></tr></table>
</td>
    </tr>
</form>
</table>
<%
Set Conn=nothing
%>
</body>
</html>







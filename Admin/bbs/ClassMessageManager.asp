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
Dim int_Start,int_RPP,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=15 '����ÿҳ��ʾ��Ŀ
toF_="<font face=webdings>9</font>"   			'��ҳ 
str_nonLinkColor_="#999999" '����������ɫ
int_showNumberLink_=10 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
toF_="<font face=webdings>9</font>"   			'��ҳ 
toP10_=" <font face=webdings>7</font>"			'��ʮ
toP1_=" <font face=webdings>3</font>"			'��һ
toN1_=" <font face=webdings>4</font>"			'��һ
toN10_=" <font face=webdings>8</font>"			'��ʮ
toL_="<font face=webdings>:</font>"
%>
<html>
<HEAD>
<TITLE>FoosunCMS����ϵͳ</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="javascript">
function opencat(cat)
{	//alert(cat);
  if(document.getElementById(cat).style.display=="none"){
     document.getElementById(cat).style.display="";
	 document.getElementById("Img"+cat).src="images/nofollow.gif";
  } else {
     document.getElementById(cat).style.display="none"; 
	 document.getElementById("Img"+cat).src="images/plus.gif";
  }
}
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = form.elements[i];  
    if (e.name != 'chkall')  
       e.checked = form.chkall.checked;  
    }  
 }
function ShowNote(NoteID,ClassName,ClassID)
{
location="ShowNote.asp?NoteID="+NoteID+"&ClassName="+ClassName+"&ClassID="+ClassID;
}
</script>
<body>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table"> 
  <tr> 
    <td align="left" class="xingmu">���Թ���&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="../../help?Lable=Message" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td> 
  </tr>
</table>
<%
Dim TempRs
'�������
if Request.queryString("Au")="Y" and Request.QueryString("ID")<>"" then
	Set TempRs=Server.CreateObject(G_FS_RS)
	TempRs.open "Select State From FS_WS_BBS Where ID="&CintStr(Request.QueryString("ID"))&"",Conn,3,3
	if not TempRs.eof then
		TempRs(0)="1"
		TempRs.Update
	end if
	Set TempRs=nothing
elseif Request.queryString("Au")="N" and Request.QueryString("ID")<>"" then
Set TempRs=Server.CreateObject(G_FS_RS)
	TempRs.open "Select State From FS_WS_BBS Where ID="&CintStr(Request.QueryString("ID"))&"",Conn,3,3
	if not TempRs.eof then
		TempRs(0)="0"
		TempRs.Update
	end if
	Set TempRs=nothing
End if
Dim ClassRs,ClassSql,NoteRs,NoteSql,MsRs,MsSql,NoteAct,NoteSqlEnd
Set ClassRs=Server.CreateObject(G_FS_RS)
Set MsRs=Server.CreateObject(G_FS_RS)
ClassRs.open "Select ID,ClassID,ClassName,ClassExp,Pid,Author from FS_WS_Class order by Pid,id desc",Conn,1,1
If not ClassRs.eof then
	Do While not ClassRs.eof
%><table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr>
    <td width="3%" height="30" align="center" valign="middle" class="xingmu" onMouseUp="opencat('<%=ClassRs("ClassID")%>')"> <img src="images/nofollow.gif" name="Img" id="Img<%=ClassRs("ClassID")%>"> </td>
    <td  height="30" align="left" class="xingmu" colspan="3"> <font size="2"><%=ClassRs("ClassName")%></font> </td>
  </tr>
  <tr>
    <td class="hback" colspan="3" width="14%">&nbsp;<img src="images/forum_readme.gif"><%=ClassRs("ClassExp")%></td>
	<td width="59%" class="hback" ><a href="?Act=all">��������</a>��<a href="AddNewNote.asp?ClassID=<%=ClassRS("ClassID")%>">�������</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="?Act=Adime">����Ա�ɼ���</a>��<a href="?Act=Y">�����</a>��<a href="?Act=N">δ���</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="?Act=Before">�Ƽ�����</a>��<a href="?Act=Person">��������</a></td>
  </tr>
  <tr id="<%=ClassRs("ClassID")%>" style="display:">
    <td colspan="4" class="hback">
	<%
	NoteSql="Select ID,ClassID,[User],Topic,Body,AddDate,IsTop,State,Style,IsAdmin,Answer,Hit,LastUpdateDate,LastUpdateUser,Face from FS_WS_BBS"
	NoteSqlEnd=" order by IsTop DESC,AddDate DESC"
	if Request.queryString("Act")<>"" then
		NoteAct=Request.queryString("Act")
		Select Case NoteAct
		Case "All"
			NoteSql=NoteSql&NoteSqlEnd
		Case "Adime"
			NoteSql=NoteSql&" Where ClassID='"&NoSqlHack(ClassRs("ClassID"))&"' and IsAdmin='1' and ParentID='0'  order by IsTop DESC,AddDate DESC"
		Case "Y"
			NoteSql=NoteSql&" Where ClassID='"&NoSqlHack(ClassRs("ClassID"))&"' and State='0' and ParentID='0'  " &NoteSqlEnd
		Case "N"
			NoteSql=NoteSql&" Where ClassID='"&NoSqlHack(ClassRs("ClassID"))&"' and State='1' and ParentID='0' " &NoteSqlEnd
		Case "Before"	
			NoteSql=NoteSql&" where ClassID='"&NoSqlHack(ClassRs("ClassID"))&"' and IsTop='1' and ParentID='0' "  &NoteSqlEnd
		Case "Person"
			NoteSql=NoteSql&" where ClassID='"&NoSqlHack(ClassRs("ClassID"))&"' and Hit>0 and ParentID='0'  order by Hit DESC,IsTop DESC,AddDate DESC"
		Case Else
			NoteSql=NoteSql&" Where ClassID='"&NoSqlHack(ClassRs("ClassID"))&"' and ParentID='0' "&NoteSqlEnd
		End Select
	else
			NoteSql=NoteSql&" Where ClassID='"&NoSqlHack(ClassRs("ClassID"))&"' and ParentID='0'  "&NoteSqlEnd
	end if
	Set NoteRs=Server.CreateObject(G_FS_RS)
	NoteRs.open NoteSql,Conn,1,1
	%>
      <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="table" id="<%=ClassRs("ClassID")%>">
	  	<form method="post" action="NoteDel.asp?Act=del" name="mainform"  id="mainform">
       <%
	   if not NoteRs.eof then	
		   NoteRs.PageSize=int_RPP
			cPageNo=NoSqlHack(Request.QueryString("Page"))
			If cPageNo="" Then 
				cPageNo = 1
			End if
			If not isnumeric(cPageNo) Then 
				cPageNo = 1
			End If
			cPageNo = Clng(cPageNo)
			If cPageNo<=0 Then 
				cPageNo=1
			End If
			If cPageNo>NoteRs.PageCount Then 
				cPageNo=NoteRs.PageCount 
			End IF
			NoteRs.AbsolutePage=cPageNo		
			%>
		<tr>
		 <td class="hback" width="4%">&nbsp;</td>
		 <td class="hback" width="7%" align="center">״ ̬</td>		
		 <td class="hback" width="20%" align="center">���ӱ���</td>
		 <td class="hback" width="11%" align="center">�� ��</td>
		 <td class="hback" width="10%"  align="center">�� ��</td>
		 <td class="hback" width="11%"  align="center">�� ��</td>
		 <td class="hback" width="21%"  align="center">������</td>
		 <td class="hback" width="5%"  align="center">�� ��</td>
		 <td class="hback" width="11%"  align="center">��  ��</td>
		<%
			FOR int_Start=1 TO int_RPP 	
		 %>
	    <tr>
          <td class="hback" align="center" width="4%"><input type="checkbox" id="NoteID" name="NoteID" value="<%=NoteRs("ID")%>"></td>
		  <td class="hback" align="center" width="7%" style="CURSOR: hand"  onmouseup="ShowNote(<%=NoteRs("ID")%>,'<%=ClassRs("ClassName")%>','<%=ClassRs("ClassID")%>')">
		  <%
		  	if NoteRs("State")="1" then
			 	Response.write("<img src=""images\lock.gif"" alt=""��������"">")
			elseif NoteRs("IsAdmin")="1" then
			 		Response.write("<img src=""images\Admin.gif"" alt=""����Ա�ɼ�����"">")
			elseif NoteRs("IsTop")="1" then
			 		Response.write("<img src=""images\top.gif"" alt=""�Ƽ�����"">")
			else				
			 	Response.write("<img src=""images\gogo.gif"" alt=""��ͨ����"">")
			end if
		  %>
		  </td>
		  <td class="hback" align="center"><a href="#" onClick="ShowNote(<%=NoteRs("ID")%>,'<%=ClassRs("ClassName")%>','<%=ClassRs("ClassID")%>')"><font color="red"><%=left(NoteRs("Topic"),30)%></font></a></td>
		  <td class="hback" align="center"><%=NoteRs("User")%></td>
		  <td class="hback"	align="center"><%=NoteRs("Answer")%></td>
		  <td class="hback"	align="center"><%=NoteRs("Hit")%></td>
		  <td class="hback"	align="center"><%=NoteRS("LastUpdateDate")%>��<%=NoteRS("LastUpdateUser")%></td>
		  <td class="hback"	align="center">
		  <%
		  if NoteRs("State")=0 then
		  	Response.write("<a href='?Au=Y&ID="&NoteRs("ID")&"'><font color=red>��</font></a>")
		  else
		  	Response.write("<a href='?Au=N&ID="&NoteRs("ID")&"'><font color=red>δ</font></a>")
		  end if
		  %></td>
		  <td class="hback" align="center"><a href="NoteEdit.asp?Act=NoteEdit&ID=<%=NoteRs("ID")%>">�޸�</a> �� <a href="NoteDel.asp?ID=<%=NoteRs("ID")%>&Act=single" onClick="{if(confirm('���ɾ���û���,��ô��ص����۶�����ɾ��,ȷ��Ҫɾ����?')){return true;}return false;}">ɾ��</a>
		  </td>
        </tr>
		<%
		 NoteRs.MoveNext
		 if NoteRs.eof or NoteRs.bof then exit for
     	 NEXT		   	  
	    %>
		<tr> 
      <td colspan="9" class="hback">
<table width="100%"><tr><td width="40%" align="center">
          <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
          ѡ������ 
          <input name="Submit" type="submit" id="Submit"  onClick="{if(confirm('ȷʵҪ����ɾ����?')){this.document.ClassForm.submit();return true;}return false;}" value=" ɾ�� ">
        </td><%
	 Response.Write("<td class=""hback"" colspan=""9"" align=""right"" width=""60%"">"&fPageCount(NoteRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo) &"</td>")

	else
		Response.Write("<tr><td class=""hback"" colspan=""8"" align=""left"">������</td>")
	END IF
	Set NoteRs=nothing
	%>	</tr></table></td>
    </tr>
	  </form>
      </table>
    </td>
  </tr>
</table>
  <%  
  ClassRs.movenext
  Loop
Else
	Response.Write("��������")
End If
Set ClassRs=nothing
Set Conn=nothing
%>
</body>
</html>







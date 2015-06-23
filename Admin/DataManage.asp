<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%'Copyright (c) 2006 Foosun Inc
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,strShowErr,tmp_getPath
MF_Default_Conn
MF_Session_TF
Dim MF_DataPath,ME_DataPath,Data_FolderPath,Data_BakPath
if not MF_Check_Pop_TF("MF_DataFix") then Err_Show
if Request("Action")="Fix" then
	if not MF_Check_Pop_TF("MF015") then Err_Show
	'-------
	If G_IS_SQL_DB = 0 Then
		IF G_VIRTUAL_ROOT_DIR = "" Then
			MF_DataPath = "/" & G_DATABASE_CONN_STR
		Else
			MF_DataPath = "/" & G_VIRTUAL_ROOT_DIR & "/" & G_DATABASE_CONN_STR
		End If
		MF_DataPath = Replace(MF_DataPath,"//","/")	
	Else
		Response.Write "SQL���ݿ����SQL SERVER����ѹ��" : Response.End
	End If
	If G_IS_SQL_User_DB = 0 Then
		IF G_VIRTUAL_ROOT_DIR = "" Then
			ME_DataPath = "/" & G_User_DATABASE_CONN_STR
		Else
			ME_DataPath = "/" & G_VIRTUAL_ROOT_DIR & "/" & G_User_DATABASE_CONN_STR
		End If
		ME_DataPath = Replace(ME_DataPath,"//","/")	
	Else
		Response.Write "SQL���ݿ����SQL SERVER����ѹ��" : Response.End
	End If
	'-------
	'On Error Resume Next
	Dim oldDB,bakDB,newDB,FSO,Engine,EditFile,prov
	if Request.QueryString("Sub") = "MF" then
		oldDB = MF_DataPath
		Data_FolderPath = Replace(MF_DataPath,Split(MF_DataPath,"/")(Ubound(Split(MF_DataPath,"/"))),"")
		bakDB = Data_FolderPath & "DataBase_BackUp/MF_Fix.Mdb" 
		newDB = Data_FolderPath & "MF_Fixed.Mdb"
		
	Else
		oldDB = ME_DataPath
		Data_FolderPath = Replace(ME_DataPath,Split(ME_DataPath,"/")(Ubound(Split(ME_DataPath,"/"))),"")
		bakDB = Data_FolderPath & "DataBase_BackUp/ME_Fix.Mdb" 
		newDB = Data_FolderPath & "ME_Fixed.Mdb"
	End if
	Data_BakPath = Data_FolderPath & "DataBase_BackUp/"
	Set FSO = Server.CreateObject(G_FS_FSO) 
	If FSO.FolderExists(Server.MapPath(Data_BakPath)) = False Then
		FSO.createFolder(Server.MapPath(Data_BakPath))
	End If
	oldDB = Server.MapPath(oldDB)
	bakDB = Server.MapPath(bakDB)
	newDB = Server.MapPath(newDB)	
    Conn.Close : Set Conn = Nothing
	FSO.CopyFile oldDB,bakDB,true 
    Set Engine = Server.CreateObject(G_JRO_JETENGINE) 
    prov = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    Engine.CompactDatabase prov & OldDB,prov & newDB 
    set Engine = nothing 
	FSO.DeleteFile oldDB 
	FSO.MoveFile newDB,oldDB
	FSO.DeleteFile bakDB  
    set FSO = nothing  
	MF_Default_Conn
	Call MF_Insert_oper_Log("���ݿ�ά��","ѹ�����ݿ�",now,session("admin_name"),"MF")
	strShowErr = "<li>���ݿ�ѹ���ɹ�.</li>"
	Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
if Request("Action")="bak" then
	if not MF_Check_Pop_TF("MF016") then Err_Show
	'-------
	If G_IS_SQL_DB = 0 Then
		IF G_VIRTUAL_ROOT_DIR = "" Then
			MF_DataPath = "/" & G_DATABASE_CONN_STR
		Else
			MF_DataPath = "/" & G_VIRTUAL_ROOT_DIR & "/" & G_DATABASE_CONN_STR
		End If
		MF_DataPath = Replace(MF_DataPath,"//","/")	
	Else
		Response.Write "SQL���ݿ����SQL SERVER���б���" : Response.End
	End If
	If G_IS_SQL_User_DB = 0 Then
		IF G_VIRTUAL_ROOT_DIR = "" Then
			ME_DataPath = "/" & G_User_DATABASE_CONN_STR
		Else
			ME_DataPath = "/" & G_VIRTUAL_ROOT_DIR & "/" & G_User_DATABASE_CONN_STR
		End If
		ME_DataPath = Replace(ME_DataPath,"//","/")	
	Else
		Response.Write "SQL���ݿ����SQL SERVER���б���" : Response.End
	End If
	'-------
	Randomize
	Dim tmp_GetRamCode
	tmp_GetRamCode = GetRamCode(16)
	Set FSO = Server.CreateObject(G_FS_FSO) 
	if Request.QueryString("Sub") = "MF" then
		oldDB = MF_DataPath
		Data_FolderPath = Replace(MF_DataPath,Split(MF_DataPath,"/")(Ubound(Split(MF_DataPath,"/"))),"")
		bakDB = Data_FolderPath & "DataBase_BackUp/MF_" & tmp_GetRamCode & ".mdb"
	Else
		oldDB = ME_DataPath
		Data_FolderPath = Replace(ME_DataPath,Split(ME_DataPath,"/")(Ubound(Split(ME_DataPath,"/"))),"")
		bakDB = Data_FolderPath & "DataBase_BackUp/ME_" & tmp_GetRamCode & ".mdb"
	End if
	Data_BakPath = Data_FolderPath & "DataBase_BackUp/"
	If FSO.FolderExists(Server.MapPath(Data_BakPath)) = false then 
		FSO.createFolder Server.MapPath(Data_BakPath)
	End If	 
	oldDB = Server.MapPath(oldDB)
	bakDB = Server.MapPath(bakDB)
	FSO.CopyFile oldDB,bakDB,true 
    set FSO = nothing  
	if  Request.QueryString("Sub") = "MF" then
		tmp_getPath = Data_BakPath & "MF_" & tmp_GetRamCode & ".mdb"
	Else
		tmp_getPath = Data_BakPath & "ME_" & tmp_GetRamCode & ".mdb"
	End if
		Call MF_Insert_oper_Log("���ݿ�ά��","�������ݿ⣬���ƣ�"& bakDB &"",now,session("admin_name"),"MF")
		strShowErr = "<li>���ݳɹ�.</li><li>�ļ���:"& bakDB &"</li><li>�뼰ʱ����<a href="""&Replace(tmp_getPath,"//","/")&"""><b><<����>></b></a>&nbsp;&nbsp;���ؽ�������<a href=""DataManage.asp?Action=DelData&File="& Replace(tmp_getPath,"//","/") &"""><b><<ɾ��>></b></a></li>"
	Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
if Request("Action")="DelData" then
	if not MF_Check_Pop_TF("MF016") then Err_Show
    Set FSO = Server.CreateObject(G_FS_FSO) 
    FSO.DeleteFile Server.MapPath(Request.QueryString("File")) 
	strShowErr = "<li>ɾ���ɹ�.</li>"
	Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=DataManage.asp?Type=bak")
	Response.end
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="xingmu">
    <td class="xingmu"><a href="#" onMouseOver="this.T_BGCOLOR='#404040';this.T_FONTCOLOR='#FFFFFF';return escape('<div align=\'center\'>FoosunCMS5.0<br> Copyright (c) 2006 Foosun Inc</div>')" class="sd"><strong>���ݿ�ά��</strong></a>
      </td>
  </tr>
  <tr class="hback">
    <td><a href="DataManage.asp">��ҳ</a><%if G_IS_SQL_DB = 0 then%>&nbsp;��&nbsp;<a href="DataManage.asp?Type=fix">���ݿ�ѹ��</a>&nbsp;��&nbsp;<a href="DataManage.asp?Type=bak">���ݿⱸ��</a><%end if%>&nbsp;��&nbsp;<a href="DataManage.asp?Type=SQLexe">SQL����ѯ����</a></td>
  </tr>
</table>
<%
Dim tmp_type
tmp_type = NoSqlHack(Trim(Request.QueryString("Type")))
select Case tmp_type 
		Case "fix"
			if not MF_Check_Pop_TF("MF015") then Err_Show
			Call fixs()
		Case "bak"
			if not MF_Check_Pop_TF("MF016") then Err_Show
			Call Bak()
		Case "SQLexe"
			if not MF_Check_Pop_TF("MF017") then Err_Show
			Call SQLexe()
		Case else
			if not MF_Check_Pop_TF("MF017") then Err_Show
			Call SQLexe()
End Select
Sub fixs()
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu"> 
    <td class="xingmu">���ݿ�ѹ��</td>
  </tr>
  <tr> 
    <td height="52" class="hback">
	 <%
	 	If G_IS_SQL_DB = 0 Then
	 %>
	  <input type="button" name="Submit" value="��ʼѹ�������ݿ�" onClick="window.location.href='DataManage.asp?Action=Fix&Type=fix&Sub=MF'">
	 <%
	 	End If
		If G_IS_SQL_User_DB = 0 Then
	 %>  
      <input type="button" name="Submit2" value="��ʼѹ����Ա���ݿ�" onClick="window.location.href='DataManage.asp?Action=Fix&Type=fix&Sub=ME'">
	 <%
	 	End if
	 %> 
	  </td>
  </tr>
  <tr>
    <td height="22" class="hback">˵����ѹ��ǰ�뱸���������ݿ⡣�Է�ֹ��һ</td>
  </tr>
</table>
<%End Sub%>
<%Sub bak()%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu"> 
    <td class="xingmu">���ݿⱸ��</td>
  </tr>
  <tr> 
    <td height="51" class="hback">
	<%
	 	If G_IS_SQL_DB = 0 Then
	 %>
	<input type="button" name="Submit22" value="��ʼ���������ݿ�" onClick="window.location.href='DataManage.asp?Action=bak&Type=fix&Sub=MF'"> 
     <%
	 	End If
		If G_IS_SQL_User_DB = 0 Then
	 %> 
	  <input type="button" name="Submit222" value="��ʼ���ݻ�Ա���ݿ�" onClick="window.location.href='DataManage.asp?Action=bak&Type=bak&Sub=ME'">
	 <%
	 	End if
	 %> 
	  </td>
  </tr>
  <tr>
    <td height="31" class="hback">˵�������ڱ�����ɺ󣬼�ʱɾ�������ļ����Է�ֹ���˶����������ݿ��ļ�</td>
  </tr>
</table>
<%End sub%>
<%Sub SQLexe()%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="form1" method="post" action=""><tr class="xingmu"> 
    <td class="xingmu">���ݿ�ӣѣ�����ѯ����</td>
  </tr>
  <tr> 
      <td class="hback">˵����ע��һ��ֻ��ִ��һ��Sql��䡣������SQL����Ϥ���뾡����Ҫʹ�á�����һ���������������ġ�<br>
        ����ʹ�ò�ѯ���.�磺select count(id) From FS_MF_Admin order by id desc,������Ҫʹ��delete,update������</td>
  </tr>
  <tr> 
    <td class="hback"><textarea name="Content" rows="5" wrap="OFF" style="width:100%;"></textarea></td>
  </tr>
  <tr> 
    <td class="hback"><iframe id="ResultShowFrame" scrolling="yes" src="DataSqlResult.asp" style="width:100%;" frameborder=1></iframe></td>
  </tr>
  <tr>
    <td class="hback">
        <input type="button" name="Submit3" value="ִ��SQL���" onClick="{if(confirm('��ȷ��ִ��SQL�����\nһ��SQLִ����ɾ�����߲�������������������!!!')){ExecuteSql();return true;}return false;}">
        <input name="Result" type="hidden" id="Result" value="Submit">
      </td>
  </tr></form>
</table>
<%End sub%>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
</body>
</html>
<%
Set Conn = Nothing
%><script language="JavaScript">
function ExecuteSql()
{
	var FormObj=frames["ResultShowFrame"].document.ExecuteForm;
	if (document.all.Content.value!='')
	{
		FormObj.Sql.value=document.all.Content.value;
		FormObj.Result.value='Submit';
		FormObj.submit();
		FormObj.Result.value='';
	}
	else alert('����дSQL���');
}
</script>







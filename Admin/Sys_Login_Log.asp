<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim Conn
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo

int_RPP=20'����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("MF_Log") then Err_Show
if not MF_Check_Pop_TF("MF019") then Err_Show
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<%
Dim obj_Log_Rs,strpage,select_count,select_pagecount,i,Tmp_adminname,Tmp_super,Tmp_Lock,tmp_my,SQL,tmp_admin_name,tmp_Type
strpage=CintStr(request("page"))
if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
Set  obj_Log_Rs = server.CreateObject(G_FS_RS)
if tRIM(Request("Type"))<>"" then:tmp_Type=" and Log_TF='"& NoSqlHack(tRIM(Request("Type"))) &"'":else:tmp_Type="":end if
SQL = "Select ID,Admin_Name,Log_IP,Log_OS_Sys,Log_TF,Log_Error_Pass,Log_Time  from FS_MF_Login_Log where ID>0  "& tmp_Type &" Order by id desc"
obj_Log_Rs.Open SQL,Conn,1,3
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="xingmu">
    <td class="xingmu"><p>��ȫ������־</p>
      </td>
  </tr>
  <tr class="hback">
    <td><a href="Sys_Login_Log.asp">��ҳ</a>��<a href="Sys_Login_Log.asp?Type=0">������־</a> 
      ��<a href="Sys_Login_Log.asp?Type=1">�ɹ���־</a><strong>��</strong><a href="Sys_Oper_Log.asp"><strong>������־</strong></a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="LogForm" method="post" action="">
    <tr class="hback"> 
      <td width="20%" height="25" class="xingmu"> <div align="left">��½��</div></td>
      <td width="16%" height="25" class="xingmu"> <div align="center">״̬</div></td>
      <td width="19%" height="25" class="xingmu"> <div align="center">��������</div></td>
      <td width="17%" height="25" class="xingmu"> <div align="center">IP</div></td>
      <td width="28%" class="xingmu"><div align="center">����</div></td>
    </tr>
    <%
if obj_Log_Rs.eof then
   obj_Log_Rs.close
   set obj_Log_Rs=nothing
   Response.Write"<TR  class=""hback""><TD colspan=""6""  class=""hback"" height=""40"">û�в�����</TD></TR>"
else
	obj_Log_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>obj_Log_Rs.PageCount Then cPageNo=obj_Log_Rs.PageCount 
	obj_Log_Rs.AbsolutePage=cPageNo
	for i=1 to obj_Log_Rs.pagesize
		if obj_Log_Rs.eof Then exit For 
%>
    <tr class="hback"> 
      <td height="25" >��
        <% = obj_Log_Rs("Admin_Name")%></td>
      <td height="25"><div align="center"> 
          <% if obj_Log_Rs("Log_TF")="1" then:response.Write("�ɹ�"):else:Response.Write("<span class=""tx"">ʧ��</span>"):end if%>
        </div></td>
      <td height="25"><div align="center"> 
          <% if obj_Log_Rs("Log_TF")="1" then:response.Write"":else:Response.Write("<span class=""tx"">"&obj_Log_Rs("Log_Error_Pass")&"</span>"):end if%></a>
          </div></td>
      <td height="25"><div align="center"> 
          <% = obj_Log_Rs("Log_IP")%>
        </div></td>
      <td height="25"><div align="center"> 
          <% = obj_Log_Rs("Log_Time")%>
        </div></td>
    </tr>
    <%
		obj_Log_Rs.movenext
	Next
	%>
    <tr class="hback"> 
      <td height="25" colspan="5"><div align="right"> 
          <input name="Action" type="hidden" id="Action">
          <input type="button" name="Submit222" value="ɾ��������־��ֻ��ɾ�����<% = G_HOLD_LOG_DAY_NUM %>����ǰ����־"   onClick="document.LogForm.Action.value='Del';{if(confirm('ȷ�������־��')){this.document.LogForm.submit();return true;}return false;}">
        </div></td>
    </tr>
  </form>
</table>

<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr height="1" class="hback"> 
    <td height="25"> 
      <%
		response.Write "<p>"&  fPageCount(obj_Log_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
    End if%>
    </td>
  </tr>
</table>
</body>
</html>
<%
set obj_Log_Rs = nothing
if Request.Form("Action")="Del" then
	if not MF_Check_Pop_TF("MF020") then Err_Show
	Dim obj_DLog_rs,strShowErr,tmp_i,DatAllowDate,f_Sql
	Set  obj_DLog_rs = server.CreateObject(G_FS_RS)
	DatAllowDate=dateadd("d",-G_HOLD_LOG_DAY_NUM,date())
	If G_IS_SQL_DB=1 then
		f_Sql = "Select ID from FS_MF_Login_Log Where Log_Time<='" & DatAllowDate & " 23:59:59' Order by  id desc"
	Else
		f_Sql = "Select ID from FS_MF_Login_Log Where  Log_Time<=#" & DatAllowDate & " 23:59:59# Order by  id desc"
	End If
	obj_DLog_rs.Open f_Sql,Conn,1,3
	tmp_i = 0 
	Do while not obj_DLog_rs.eof 
		Conn.execute("Delete From FS_MF_Login_Log where ID="&obj_DLog_rs("ID"))
		obj_DLog_rs.movenext
		tmp_i = tmp_i +1 
	Loop
	if tmp_i>0 then
		Call MF_Insert_oper_Log("ɾ����ȫ��־","ɾ�����а�ȫ��־,ɾ����"& tmp_i &"����־",now,session("admin_name"),"MF")
	End if
	strShowErr = "<li>��־ɾ���ɹ�,��ɾ��"& tmp_i &"����־</li>"
	Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
	obj_DLog_rs.close:set obj_DLog_rs = nothing
End if
Set Conn = Nothing
%>
<script language="JavaScript" type="text/JavaScript">
function opencat(cat)
{
  if(cat.style.display=="none"){
     cat.style.display="";
  } else {
     cat.style.display="none"; 
  }
}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->






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
if not MF_Check_Pop_TF("MF018") then Err_Show
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
Dim obj_Log_Rs,strpage,select_count,select_pagecount,i,Tmp_adminname,Tmp_super,Tmp_Lock,tmp_my,SQL,tmp_admin_name,tmp_SubType
strpage=NoSqlHack(request("page"))
'if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
Set  obj_Log_Rs = server.CreateObject(G_FS_RS)
if tRIM(Request("Admin_Name"))<>"" then:tmp_admin_name=" and Admin_Name='"& NoSqlHack(Trim(Request("Admin_Name"))) &"'":else:tmp_admin_name="":end if
if tRIM(Request("SubType"))<>"" then:tmp_SubType=" and Logtype='"& NoSqlHack(tRIM(Request("SubType"))) &"'":else:tmp_SubType="":end if
SQL = "Select ID,LogTitle,LogContent,LogTime,Admin_Name,Logtype  from FS_MF_Oper_Log where ID>0  "& tmp_SubType & tmp_admin_name &" Order by id desc"
obj_Log_Rs.Open SQL,Conn,1,1
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="xingmu">
    <td class="xingmu"><a href="#" class="sd"><strong>����Ա������־</strong></a>
      </td>
  </tr>
  <tr class="hback">
    <td><a href="Sys_Oper_Log.asp">��ҳ</a>��<a href="Sys_Oper_Log.asp?SubType=MF">��ϵͳ</a>��<a href="Sys_Oper_Log.asp?SubType=NS">����</a>��<%if Request.Cookies("FoosunSUBCookie")("FoosunSUBMS")=1 then%><a href="Sys_Oper_Log.asp?SubType=MS">�̳�</a>��<%end if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBDS")=1 then%><a href="Sys_Oper_Log.asp?SubType=DS">����</a>��<%end if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBCS")=1 then%><a href="Sys_Oper_Log.asp?SubType=CS">�ɼ�</a>��<%end if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBAP")=1 then%><a href="Sys_Oper_Log.asp?SubType=AP">�˲�</a>��<%end if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBHS")=1 then%><a href="Sys_Oper_Log.asp?SubType=HS">����</a>��<%End if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBSS")=1 then%><a href="Sys_Oper_Log.asp?SubType=SS">ͳ��</a>��<%End if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBSD")=1 then%><a href="Sys_Oper_Log.asp?SubType=SD">����</a>��<%End if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBME")=1 then%><a href="Sys_Oper_Log.asp?SubType=ME">��Ա</a>��<%end if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBVS")=1 then%><a href="Sys_Oper_Log.asp?SubType=VS">ͶƱ</a>��<%End if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBAS")=1 then%><a href="Sys_Oper_Log.asp?SubType=AS">���</a>��<%End if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBWS")=1 then%><a href="Sys_Oper_Log.asp?SubType=WS">���Ա�</a>��<%end if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBFL")=1 then%><a href="Sys_Oper_Log.asp?SubType=FL">��������</a><%End if%>��<a href="Sys_Login_Log.asp"><strong>��ȫ��־</strong></a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="LogForm" method="post" action="">
    <tr class="hback"> 
      <td width="17%" height="25" class="xingmu"> <div align="left">����(�����鿴��������)</div></td>
      <td width="13%" height="25" class="xingmu"> <div align="center">����</div></td>
      <td width="18%" height="25" class="xingmu"> <div align="center">����Ա</div></td>
      <td width="11%" height="25" class="xingmu"> <div align="center">��ϵͳ</div></td>
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
	If cPageNo>obj_Log_Rs.PageCount Then cPageNo=obj_Log_Rs.PageCount 
	If cPageNo<=0 Then cPageNo=1
	obj_Log_Rs.AbsolutePage=cPageNo
	for i=1 to obj_Log_Rs.pagesize
		if obj_Log_Rs.eof Then exit For 
%>
    <tr class="hback"> 
      <td height="25"> <div  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(Logid<% = obj_Log_Rs("ID")%>);"  language=javascript>��<a href="#"><% = obj_Log_Rs("LogTitle")%></a></div></td>
      <td height="25"><div align="center"> 
          <% = obj_Log_Rs("LogTime")%>
        </div></td>
      <td height="25"><div align="center"> <a href="Sys_Oper_Log.asp?Admin_Name=<% = obj_Log_Rs("Admin_Name")%>"> 
          <% = obj_Log_Rs("Admin_Name")%>
          </a> </div></td>
      <td height="25"><div align="center"> 
          <% = obj_Log_Rs("Logtype")%>
        </div></td>
    </tr>
    <tr valign="top" class="hback" id="Logid<% = obj_Log_Rs("ID")%>" style="display:none"> 
      <td height="32" colspan="4"> �� 
        <% = obj_Log_Rs("LogContent")%>
      </td>
    </tr>
    <%
		obj_Log_Rs.movenext
	Next
	%>
    <tr class="hback"> 
      <td height="25" colspan="4"><div align="right"> 
          <input name="Action" type="hidden" id="Action">
          <input type="button" name="Submit222" value="ɾ��������־��ֻ��ɾ�����<% = G_HOLD_LOG_DAY_NUM %>����ǰ����־"   onClick="document.LogForm.Action.value='Del';{if(confirm('ȷ�������־��')){this.document.LogForm.submit();return true;}return false;}">
        </div></td>
    </tr>
    <tr class="hback">
      <td height="25" colspan="4"><span class="tx">��ϵͳ����˵��</span><br>
        MF-��ϵͳ NS-����ϵͳ MS-�̳� DS-���� CS-�ɼ�ϵͳ AP-�˲�ϵͳ HS-����ϵͳ <br>
        SS-ͳ��ϵͳ ME-��Աϵͳ VS-ͶƱϵͳ AS-���ϵͳ WS-����ϵͳ FL-��������</td>
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
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
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
		f_Sql = "Select ID from FS_MF_Oper_Log Where logtime<='" & DatAllowDate & " 23:59:59' Order by  id desc"
	Else
		f_Sql = "Select ID from FS_MF_Oper_Log Where  logtime<=#" & DatAllowDate & " 23:59:59# Order by  id desc"
	End If
	obj_DLog_rs.Open f_Sql,Conn,1,3
	tmp_i = 0 
	Do while not obj_DLog_rs.eof 
		Conn.execute("Delete From FS_MF_Oper_Log where ID="&CintStr(obj_DLog_rs("ID")))
		obj_DLog_rs.movenext
		tmp_i = tmp_i +1 
	Loop
	if tmp_i>0 then
		Call MF_Insert_oper_Log("ɾ��������־","ɾ�����в�����־,ɾ����"& tmp_i &"����־",now,session("admin_name"),"MF")
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






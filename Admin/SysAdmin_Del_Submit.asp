<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Cls_Cache.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%
Dim Conn,strShowErr
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("MF_Pop") then Err_Show
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo

int_RPP=20 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"
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
Dim obj_admin_get_rs
set obj_admin_get_rs= Conn.execute("select ID,Admin_Name,Admin_Parent_Admin,Admin_Is_Super,Admin_Real_Name,Admin_Is_Locked From FS_MF_Admin where ID="&CintStr(Request.QueryString("Id")))
Dim obj_admin_Rs,strpage,select_count,select_pagecount,i,Tmp_adminname,Tmp_super,Tmp_Lock,tmp_my,SQL
strpage=NoSqlHack(request("page"))
if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
Set obj_admin_Rs = server.CreateObject(G_FS_RS)
SQL = "Select ID,Admin_Name,Admin_Parent_Admin,Admin_Is_Super,Admin_Real_Name,Admin_Is_Locked  from FS_MF_Admin where Admin_Name<>'"& obj_admin_get_rs("Admin_Name")&"' and  Admin_Parent_Admin<>'"& obj_admin_get_rs("Admin_Name")&"' Order by id desc"
obj_admin_Rs.Open SQL,Conn,1,3
%>
  <form name="form1" method="post" action="">
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="xingmu">
    <td class="xingmu">����Ա����</td>
  </tr>
  <tr class="hback">
    <td><a href="SysAdmin_List.asp">����Ա��ҳ</a> | <a href="SysAdmin_List.asp?Is_Super=1">��������Ա</a> 
      | <a href="SysAdmin_List.asp?islock=1">�����Ĺ���Ա</a> | <a href="SysAdmin_List.asp?islock=0">���ŵĹ���Ա</a> 
      | <a href="SysAdmin_List.asp?my=1">�ҵĹ���Ա</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">

    <tr class="hback"> 
      <td width="32%" height="25" class="xingmu"> <div align="left">ָ��ɾ���˹���Ա<span class="tx">(<% = obj_admin_get_rs("Admin_Name")%>
          )</span>����¼�����Ա�����Ĺ���Ա,����ѡ��һ�� </div></td>
    </tr>
    <tr class="hback"> 
      <td height="25"><input type="radio" name="Parent_Admin_Name" value="0">
        ���ô˹���Ա���������κι���Ա</td>
    </tr>
    <%
	Response.Write"<tr class=""hback""><td class=""hback""><table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"">"
if obj_admin_Rs.eof then
   obj_admin_Rs.close
   set obj_admin_Rs=nothing
   Response.Write"<table width=""98%"" class=""table"" align=""center""><tr  class=""hback""><td  class=""hback"" height=""40"">û�з��������Ĺ���Ա��</td></tr></table>"
else
	obj_admin_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>obj_admin_Rs.PageCount Then cPageNo=obj_admin_Rs.PageCount 
	obj_admin_Rs.AbsolutePage=cPageNo
	Response.Write"<tr class=""hback"">"
	Dim i_tmp_n
	i_tmp_n = 0
	for i=1 to obj_admin_Rs.pagesize
		if obj_admin_Rs.eof Then exit For 
			Response.Write"<td class=""hback""><input type=""radio"" name=""Parent_Admin_Name"" value="""& obj_admin_Rs("Admin_Name") &""">"& obj_admin_Rs("Admin_Name") &"-"& obj_admin_Rs("Admin_Real_Name") &"</td>"
		obj_admin_Rs.movenext
		i_tmp_n = i_tmp_n +1 
		if i_tmp_n mod 4 = 0 then
			Response.Write("</tr>")
		End if
	Next
		Response.Write"</tr>"
		Response.Write"</table></td></tr>"
	%>
    <tr class="hback"> 
      <td height="25"> <%
			response.Write "<p>"&  fPageCount(obj_admin_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
	%>
      </td>
    </tr>
    <%
	obj_admin_Rs.close
	set obj_admin_Rs = nothing
End if
%>
</table>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr>
    <td class="hback"><div align="right">
          <input name="AdminID" type="hidden" id="AdminID" value="<% = Request.QueryString("ID")%>">
          <input name="Action" type="hidden" id="Action" value="del_p">
          <input type="submit" name="Submit" value="ȷ��ɾ��">
      </div></td>
  </tr>
</table></form>
</body>
</html>
<%
if Request.Form("Action")="del_p" then
	if NoSqlHack(Request.Form("Parent_Admin_Name"))="" then
		strShowErr = "<li>��ѡ��һ����������Ա!!!</li>"
		Response.Redirect("Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
		Dim obj_admin_rs_2,tmp_str_d
		Set obj_admin_rs_2 = Conn.execute("Select Admin_Parent_Admin,Admin_Name,Admin_Is_Super,Admin_Add_Admin From FS_MF_Admin where ID="&CintStr(Request.Form("AdminID")))
		tmp_str_d = obj_admin_rs_2("Admin_Name")
		if session("Admin_Is_Super")<>1 then
			if obj_admin_rs_2("Admin_Name")<>session("Admin_Name") then
				if obj_admin_rs_2("Admin_Is_Super")=1 then
					strShowErr = "<li>������ɾ��ϵͳ����Ա</li>"
					Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
					Response.end
				end if
				if obj_admin_rs_2("Admin_Add_Admin")<>session("Admin_Name") then
					strShowErr = "<li>������ɾ�����˵Ĺ���Ա��</li>"
					Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
					Response.end
				end if
			end if
		end if
		'�ж��Ƿ�����������Ա
		Conn.execute("Update FS_MF_Admin set Admin_Parent_Admin ='"&NoSqlHack(Request.Form("Parent_Admin_Name"))&"' where Admin_Parent_Admin='"& NoSqlHack(tmp_str_d) &"'")
		Conn.execute("Delete From FS_MF_Admin where id="&CintStr(Request.Form("AdminID")))
		'������־
		'ɾ����̬Ŀ¼
		Dim p_FSO,tmp_path
		Set p_FSO = Server.CreateObject(G_FS_FSO)
		tmp_path = "..\"& G_UP_FILES_DIR &"\adminFiles\"& tmp_str_d
		tmp_path = Server.MapPath(Replace(tmp_path,"\\","\"))
		if p_FSO.FolderExists(tmp_path) = true then p_FSO.DeleteFolder tmp_path
		set p_FSO = nothing
		Call MF_Insert_oper_Log("ɾ������Ա","ɾ���˹���ԱID("& tmp_str_d &")��"&Request.Form("AdminID")&",ͬʱ�����˴˹���Ա�����е���������Ա",now,session("admin_name"),"MF")
		obj_admin_rs_2.close:set obj_admin_rs_2 = nothing
		strShowErr = "<li>ɾ������Ա�ɹ�</li>"
		Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
		Response.end
	End if
End if
Set Conn = Nothing
%>






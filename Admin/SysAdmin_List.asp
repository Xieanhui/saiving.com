<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Cls_Cache.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<!--#include file="../FS_Inc/md5.asp" -->
<%
Dim Conn,strShowErr
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("MF029") then Err_Show
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
Dim obj_admin_Rs,strpage,select_count,select_pagecount,i,Tmp_adminname,Tmp_super,Tmp_Lock,tmp_my,SQL
strpage=NoSqlHack(request("page"))
'if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
Set obj_admin_Rs = server.CreateObject(G_FS_RS)
if Trim(Request.QueryString("Parent_Admin"))<>"" then:Tmp_adminname = " and Admin_Parent_Admin = '"& NoSqlHack(Trim(Request.QueryString("Parent_Admin"))) &"'":Else:Tmp_adminname = "":End if
if Trim(Request.QueryString("Is_Super"))="1" then:Tmp_super =  " and Admin_Is_Super=1":Else:Tmp_super = "":End if
if Trim(Request.QueryString("islock"))="1" then
	Tmp_Lock =  " and Admin_Is_Locked=1"
Elseif  Trim(Request.QueryString("islock"))="0" then
	Tmp_Lock =  " and Admin_Is_Locked=0"
Else
	Tmp_Lock =  ""
End if
dim keys,wh
keys = Request("key")
if keys<>"" then
wh = " and (Admin_Name like '%"+keys+"%' or Admin_Real_Name like '%"+keys+"%')"
end if
Dim Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
if NoSqlHack(trim(Request.QueryString("my")))="1" then:tmp_my = " and Admin_Parent_Admin='"& Temp_Admin_Name &"'":else:tmp_my="":end if
SQL = "Select ID,Admin_Name,Admin_Parent_Admin,Admin_Is_Super,Admin_Real_Name,Admin_Is_Locked  from FS_MF_Admin where id>0 "& Tmp_adminname & wh &Tmp_super&Tmp_Lock&tmp_my&" Order by id desc"
obj_admin_Rs.Open SQL,Conn,1,1
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="xingmu">
    <td class="xingmu"><a href="#" class="sd"><strong>����Ա����</strong></a></td>
  </tr>
  <tr class="hback">
    <td><a href="SysAdmin_List.asp">����Ա��ҳ</a> | <a href="SysAdmin_List.asp?Is_Super=1">��������Ա</a> | <a href="SysAdmin_List.asp?islock=1">�����Ĺ���Ա</a> | <a href="SysAdmin_List.asp?islock=0">���ŵĹ���Ա</a> | <a href="SysAdmin_List.asp?my=1">�ҵĹ���Ա</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" method="post" action="">
    <tr class="hback">
      <td width="17%" height="25" class="xingmu"><div align="center">����Ա</div></td>
      <td width="19%" height="25" class="xingmu"><div align="center">��������Ա</div></td>
      <td width="14%" height="25" class="xingmu"><div align="center">ϵͳ����Ա</div></td>
      <td width="15%" height="25" class="xingmu"><div align="center">��ʵ����</div></td>
      <td width="8%" height="25" class="xingmu"><div align="center">״̬</div></td>
      <td width="27%" height="25" class="xingmu"><div align="center">����</div></td>
    </tr>
    <%
if obj_admin_Rs.eof then
   obj_admin_Rs.close
   set obj_admin_Rs=nothing
   Response.Write"<tr  class=""hback""><td colspan=""6""  class=""hback"" height=""40"">û�й���Ա��</td></tr>"
else
	obj_admin_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo>obj_admin_Rs.PageCount Then cPageNo=obj_admin_Rs.PageCount 
	If cPageNo<=0 Then cPageNo=1
	obj_admin_Rs.AbsolutePage=cPageNo
	for i=1 to obj_admin_Rs.pagesize
		if obj_admin_Rs.eof Then exit For 
%>
    <tr class="hback">
      <td height="25"><% = obj_admin_Rs("Admin_Name")%></td>
      <td height="25"><div align="center">
          <%
	  Dim obj_admin_Rs_1
	  set obj_admin_Rs_1 = Conn.execute("select Admin_Name from FS_MF_Admin where Admin_Name = '"& obj_admin_Rs("Admin_Parent_Admin")&"'")
	 if Not obj_admin_Rs_1.eof then
		 Response.Write "<a href=""SysAdmin_List.asp?Parent_Admin="& obj_admin_Rs_1("Admin_Name") &""">" & obj_admin_Rs_1("Admin_Name")&"</a>"
	 Else
	 	Response.Write("--")
	 End if
%>
        </div></td>
      <td height="25"><div align="center">
          <%if  obj_admin_Rs("Admin_Is_Super")=1 then:response.Write("��"):else:response.Write("��"):end if%>
        </div></td>
      <td height="25"><div align="center">
          <% = obj_admin_Rs("Admin_Real_Name")%>
        </div></td>
      <td height="25"><div align="center">
          <%if obj_admin_Rs("Admin_Is_Locked")=0 then:response.Write("����"):else:response.Write("<span class=""tx"">����</span>"):end if%>
        </div></td>
      <td height="25"><div align="left"><a href="SysAdmin_SetPop.asp?AdminID=<% = obj_admin_Rs("id")%>">����Ȩ��</a>��<a href="SysAdmin_Add.asp?Action=edit&AdminID=<% = obj_admin_Rs("id")%>">�޸�</a>
          <% 
		  '---------------------------------------------------------------------------------------
		  '�ж��Ƿ����������Ա���������������ʾɾ���������������
		  If Cstr(obj_admin_Rs("Admin_Name"))<>Cstr(Temp_Admin_Name) Then 
		  %>
          ��<a href="SysAdmin_List.asp?Action=del&AdminID=<% = obj_admin_Rs("id")%>"  onClick="{if(confirm('ȷ��Ҫɾ����?\n����˹���Ա�����ӹ���Ա\n�˹���Ա�µ���������Ա������ָ����������Ա!!\nͬʱ��ɾ���˹���Ա�ϴ��ļ�Ŀ¼�������ϴ����ļ�')){return true;}return false;}">ɾ��</a> ��<a href="SysAdmin_List.asp?Action=Lock&AdminID=<% = obj_admin_Rs("id")%>">����</a>��<a href="SysAdmin_List.asp?Action=UnLock&AdminID=<% = obj_admin_Rs("id")%>">����</a>
          <% 
		  End If
		  '---------------------------------------------------------------------------------------
		  %>
        </div></td>
    </tr>
    <%
		obj_admin_Rs.movenext
	Next
 %>
    <tr class="hback">
      <td height="25" colspan="6"><div align="right">
          <input name="Action" type="hidden" id="Action">
          <input type="button" name="Submit" value="����¹���Ա" onClick="location.href='SysAdmin_Add.asp'">
        </div></td>
    </tr>
    <tr class="hback">
      <td height="25" colspan="6"><%
			response.Write "<p>"&  fPageCount(obj_admin_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
	%>
      </td>
    </tr>
    <%
	obj_admin_Rs.close
	set obj_admin_Rs = nothing
End if
%>
  </form>
  <tr>
  <td height="25" colspan="6">
    <form name="Label_Form" method="get" action="" target="_self" style="margin:0;padding:0;" onSubmit="return false;">
        ��������Ա��<input type="text" id="key" name="keyw" /><input type="button" name="se" value="����" onClick="searcha();" />
  </form>
  </td>
  </tr>
  
<script type="text/javascript">
    function searcha()
       {
            if(document.getElementById("key").value=="")
            {
                alert("��д�ؼ���");
            return false;
            } 
            window.location.href="sysadmin_list.asp?key="+escape(document.getElementById("key").value)+"";
       } 
</script>
</table>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
</body>
</html>
<%
if Request.QueryString("Action")="del" then
	Dim obj_admin_rs_2,tmp_str_d,tmp_str_adminname
	Set obj_admin_rs_2 = Conn.execute("Select Admin_Parent_Admin,Admin_Name,Admin_Is_Super,Admin_Add_Admin From FS_MF_Admin where ID="&CintStr(Request.QueryString("AdminID")))
	tmp_str_d = obj_admin_rs_2("Admin_Name")
	tmp_str_adminname= obj_admin_rs_2("Admin_Name")
	'�ж��Ƿ�����������Ա
	Dim obj_adminTF_rs_2
	Set obj_adminTF_rs_2 = Conn.execute("Select Admin_Parent_Admin,Admin_Name From FS_MF_Admin where Admin_Parent_Admin='"& tmp_str_d &"'")
	if Not obj_adminTF_rs_2.eof then
		if session("Admin_Is_Super")<>1 then
			if obj_admin_rs_2("Admin_Name")<>session("Admin_Name") then
				if obj_admin_rs_2("Admin_Is_Super")=1 then
					strShowErr = "<li>������ɾ��ϵͳ����Ա</li>"
					Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
					Response.end
				end if
				if obj_admin_rs_2("Admin_Parent_Admin")<>session("Admin_Name") then
					strShowErr = "<li>������ɾ�����˵Ĺ���Ա��</li>"
					Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
					Response.end
				end if
			end if
		end if
		Response.Redirect "SysAdmin_Del_submit.asp?id="& NoSqlHack(Request.QueryString("AdminID")) &""
		Response.end
	Else
		if session("Admin_Is_Super")<>1 then
			if obj_admin_rs_2("Admin_Name")<>session("Admin_Name") then
				if obj_admin_rs_2("Admin_Is_Super")=1 then
					strShowErr = "<li>������ɾ��ϵͳ����Ա</li>"
					Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
					Response.end
				end if
				if obj_admin_rs_2("Admin_Parent_Admin")<>session("Admin_Name") then
					strShowErr = "<li>������ɾ�����˵Ĺ���Ա��</li>"
					Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
					Response.end
				end if
			end if
		end if
		Conn.execute("Delete From FS_MF_Admin where id="&CintStr(Request.QueryString("AdminID")))
	End if
	'������־
	'ɾ����̬Ŀ¼
'	Dim p_FSO,tmp_path,Temo_AdminPath,Admin_NameFolder,Admin_Md5Folder
'	Set p_FSO = Server.CreateObject(G_FS_FSO)
'	Temo_AdminPath = "..\"& G_UP_FILES_DIR &"\adminFiles"
'	If p_FSO.FolderExists(Server.MapPath(Temo_AdminPath)) Then
'		Admin_NameFolder = Temo_AdminPath & "\" & tmp_str_adminname
'		Admin_Md5Folder = Temo_AdminPath & "\" & UCase(md5(tmp_str_adminname,16))
'		If p_FSO.FolderExists(Server.MapPath(Admin_NameFolder)) Then
'			p_FSO.DeleteFolder Server.MapPath(Admin_NameFolder)
'		End If
'		If p_FSO.FolderExists(Server.MapPath(Admin_Md5Folder)) Then
'			p_FSO.DeleteFolder Server.MapPath(Admin_Md5Folder)
'		End If
'	End If
'	set p_FSO = nothing
	Call MF_Insert_oper_Log("ɾ������Ա","ɾ���˹���ԱID("& tmp_str_d &")��"&Request.QueryString("AdminID")&",ͬʱ�����˴˹���Ա�����е���������Ա",now,session("admin_name"),"MF")
	obj_admin_rs_2.close:set obj_admin_rs_2 = nothing
	strShowErr = "<li>ɾ������Ա�ɹ�</li>"
	Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
	Response.end
End if
if Request.QueryString("Action")="Lock" then
	if Request.QueryString("AdminID")="" then
		strShowErr = "<li>�������</li>"
		Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
		Response.end
	else
		dim rs
		set rs= Server.CreateObject(G_FS_RS)
		rs.open "select Admin_Add_Admin,Admin_Is_Super,Admin_Name,Admin_Parent_Admin From FS_MF_Admin where Id="&CintStr(Request.QueryString("AdminId")),Conn,1,3
		if session("Admin_Is_Super")<>1 then
			if rs("Admin_Name")<>session("Admin_Name") then
				if rs("Admin_Is_Super")=1 then
					strShowErr = "<li>����������ϵͳ����Ա</li>"
					Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
					Response.end
				end if
				if rs("Admin_Parent_Admin")<>session("Admin_Name") then
					strShowErr = "<li>�������������˵��ӹ���Ա��</li>"
					Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
					Response.end
				end if
			end if
		end if
		rs.close:set rs=nothing
		Conn.execute("Update FS_MF_Admin set Admin_Is_Locked=1 where ID="&CintStr(Request.QueryString("AdminId")))
		strShowErr = "<li>�����ɹ�</li>"
		Response.Redirect("success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
		Response.end
	end if
end if
if Request.QueryString("Action")="UnLock" then
	if Request.QueryString("AdminID")="" then
		strShowErr = "<li>�������</li>"
		Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
		Response.end
	else
		set rs= Server.CreateObject(G_FS_RS)
		rs.open "select Admin_Add_Admin,Admin_Is_Super,Admin_Name,Admin_Parent_Admin From FS_MF_Admin where Id="&CintStr(Request.QueryString("AdminId")),Conn,1,3
		if session("Admin_Is_Super")<>1 then
			if rs("Admin_Name")<>session("Admin_Name") then
				if rs("Admin_Is_Super")=1 then
					strShowErr = "<li>�����ܲ���ϵͳ����Ա</li>"
					Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
					Response.end
				end if
				if rs("Admin_Parent_Admin")<>session("Admin_Name") then
					strShowErr = "<li>�����ܽ������˵��ӹ���Ա��</li>"
					Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
					Response.end
				end if
			end if
		end if
		rs.close:set rs=nothing
		Conn.execute("Update FS_MF_Admin set Admin_Is_Locked=0 where ID="&CintStr(Request.QueryString("AdminId")))
		strShowErr = "<li>�����ɹ�</li>"
		Response.Redirect("success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
		Response.end
	end if
end if
Set Conn = Nothing
%>







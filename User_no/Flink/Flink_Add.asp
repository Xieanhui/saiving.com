<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../Lib/strlib.asp" -->
<%
Response.Buffer = True
Response.CacheControl = "no-cache"
Dim tmp_type,sRootDir,str_CurrPath,obj_u_fl_rs
MF_Default_Conn
MF_User_Conn
set obj_u_fl_rs = conn.execute("select top 1 IsOpen,IsRegister,ArrSize,isLock,Content From FS_FL_SysPara")
if obj_u_fl_rs.eof then
	response.Write("�Ҳ���������Ϣ���������Ա��ϵ"):response.end
else
	if obj_u_fl_rs("IsOpen")=0 then
		response.Write("û�п���������������")
		response.end
	end if
if obj_u_fl_rs("IsRegister") = 1 then
%>
<!--#include file="../lib/UserCheck.asp" -->
<%
end if
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
if obj_u_fl_rs("IsRegister") = 1 then
	str_CurrPath = Replace(sRootDir &"/"&G_USERFILES_DIR&"/"&session("FS_UserNumber"),"//","/")
End if
if trim(Request.Form("edit_save"))<>"" then
	dim obj_save_Rs
	if trim(Request.Form("isread"))="" then
			Response.Write("<script>alert(""����\n��ͬ��ע����֪"");location.href=""javascript:history.back()"";</script>")
			Response.End
	end if
	if trim(Request.Form("F_Name"))="" or trim(Request.Form("F_Type"))="" or trim(Request.Form("F_Url"))="" or trim(Request.Form("F_Url"))="http://" or trim(Request.Form("F_Content"))="" then
			Response.Write("<script>alert(""����\n��*������д"");location.href=""javascript:history.back()"";</script>")
			Response.End
	end if
	if Request.Form("F_Type")="0" then
		if Request.Form("F_PicUrl")="" then
			Response.Write("<script>alert(""����\n����дͼƬ"");location.href=""javascript:history.back()"";</script>")
			Response.End
		end if
	end if
	if len(trim(Request.Form("F_Content")))>300 then
		Response.Write("<script>alert(""����\nվ���������ܳ���300��"");location.href=""javascript:history.back()"";</script>")
		Response.End
	end if
	if len(trim(Request.Form("F_Content_des")))>300 then
		Response.Write("<script>alert(""����\n�������ɲ��ܳ���300��"");location.href=""javascript:history.back()"";</script>")
		Response.End
	end if
	set obj_save_Rs = Server.CreateObject(G_FS_RS)
	obj_save_Rs.open "select * from FS_FL_FrendList where 1=2",Conn,1,3
	obj_save_Rs.addnew
	obj_save_Rs("F_Name") = NosqlHack(Request.Form("F_Name"))
	if Request.Form("F_Type")="0" then:obj_save_Rs("F_Type") =0:else:obj_save_Rs("F_Type") =1:end if
	obj_save_Rs("F_Url") = NosqlHack(Request.Form("F_Url"))
	obj_save_Rs("F_Content") = NosqlHack(Request.Form("F_Content"))
	if Request.Form("F_Type")="0" then:obj_save_Rs("F_PicUrl") = NosqlHack(Request.Form("F_PicUrl")):end if
	if trim(obj_u_fl_rs("isLock"))=1 then:obj_save_Rs("F_Lock") =1:else:obj_save_Rs("F_Lock") = 0:end if
	obj_save_Rs("F_OrderID") = 0
	obj_save_Rs("F_IsUser") = 1
	if obj_u_fl_rs("IsRegister") = 1 then
		obj_save_Rs("F_Author") = session("FS_UserNumber")
	else
		obj_save_Rs("F_Author") = NosqlHack(Request.Form("F_Author"))
		obj_save_Rs("F_Mail") = NosqlHack(Request.Form("F_Mail"))
	end if
	obj_save_Rs("F_Content_des") = NosqlHack(Request.Form("F_Content_des"))
	obj_save_Rs("F_LinkContent") = NosqlHack(Request.Form("F_LinkContent"))
	obj_save_Rs("F_isAdmin") = 0
	if trim(Request.Form("ClassID"))="" then:obj_save_Rs("ClassID") = 0:else:obj_save_Rs("ClassID") = NosqlHack(Request.Form("ClassID")):end if
	obj_save_Rs("Addtime") = now
	obj_save_Rs.update
	obj_save_Rs.close:set obj_save_Rs = nothing
		Response.Write("<script>alert(""��ϲ��\n�ύ�ɹ�"");location.href=""Flink_Add.asp"";</script>")
		Response.End
end if
%>
<html>
<HEAD>
<TITLE><%=GetUserSystemTitle%></TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<%if Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")<>"" then%>
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<%else%>
<link href="../images/skin/Css_2/2.css" rel="stylesheet" type="text/css">
<%end if%>
<BODY>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="form1" method="post" action="" onSubmit="return checksinput();">
    <tr> 
      <td colspan="2" class="xingmu">������������</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">�Ķ�������֪</div></td>
      <td class="hback"><input name="isread" type="checkbox" id="isread" value="1">
        ���Ѿ��Ķ� ��<a href="#" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(id_fl);" onMouseOver="this.className='bg'" onMouseOut="this.className='bg1'" language=javascript>�鿴�Ķ���֪</a></td>
    </tr>
    <tr id="id_fl" style="display:none;"> 
      <td class="hback"><div align="right">������֪</div></td>
      <td class="hback"><% = obj_u_fl_rs("Content") %></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">���</div></td>
      <td class="hback"> <select name="ClassID" id="ClassID">
          <option value="">��ѡ����Ŀ</option>
          <%dim obj_list_rs
	  set obj_list_rs=conn.execute("select id,F_ClassCName from FS_FL_Class order by id desc")
	  do while not obj_list_rs.eof
		Response.Write "<option value="""& obj_list_rs("id") &""">"& obj_list_rs("F_ClassCName") &"</option>"
		  obj_list_rs.movenext
	  loop
	  obj_list_rs.close:set obj_list_rs = nothing
	  %>
        </select></td>
    </tr>
    <tr> 
      <td width="27%" class="hback"><div align="right">վ������</div></td>
      <td width="73%" class="hback"><input name="F_Name" type="text" id="F_Name" size="40">
        *</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">��������</div></td>
      <td class="hback"> <input name="F_Type" type="radio" value="1" checked >
        �������� 
          <input type="radio" name="F_Type" value="0">
        ͼƬ����*</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">���ӵ�ַ</div></td>
      <td class="hback"><input name="F_Url" type="text" id="F_Url" size="40">
        *</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">վ��˵��</div></td>
      <td class="hback"><textarea name="F_Content" rows="5" id="F_Content" style="width:60%"></textarea>
        ���300���ַ�*</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">ͼƬ��ַ</div></td>
      <td class="hback"><input name="F_PicUrl" type="text" id="F_PicUrl" size="42"> 
        <%if obj_u_fl_rs("IsRegister") = 1 then %>
        <input type="button" name="PPPChoose"  value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.form1.F_PicUrl);">
        <%end if%></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">����������(���)</div></td>
      <td class="hback"><input name="F_Author" type="text" id="F_Author" value="<%=session("FS_UserNumber")%>"></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">�����˵����ʼ�</div></td>
      <td class="hback"><input name="F_Mail" type="text" id="F_Mail2" value="<%=session("FS_UserEmail")%>"></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">��������</div></td>
      <td class="hback"><textarea name="F_Content_des" rows="5" id="textarea" style="width:60%"></textarea>
        ���300���ַ�</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input type="submit" name="Submit" value="�ύ��Ϣ" > <input type="reset" name="Submit2" value="����"> 
        <input name="edit_save" type="hidden" id="edit_save" value="add_save"></td>
    </tr>
  </form>
</table>

</body>
</html>
<%
set obj_u_fl_rs=nothing
end if
%><script language="JavaScript" type="text/JavaScript">
function opencat(cat)
{
  if(cat.style.display=="none"){
     cat.style.display="";
  } else {
     cat.style.display="none"; 
  }
}
function checksinput()
{
	if(document.form1.isread.checked==false)
	{
	 alert('��ͬ���Ķ�Э��');
	 return false;
	}
	if(document.form1.F_Name.value=='')
	{
	 alert('��дվ�����');
	 form1.F_Name.focus();
	 return false;
	}
	if(document.form1.F_Url.value=='')
	{
	 alert('��дվ�����ӵ�ַ');
	 form1.F_Url.focus();
	 return false;
	}
	if(document.form1.F_Content.value=='')
	{
	 alert('��дվ��˵��');
	 form1.F_Content.focus();
	 return false;
	}
}
</script>







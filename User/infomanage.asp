<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
User_GetParm
if Request.QueryString("type")="delinfo" then
	if Request.QueryString("Id")="" then
		strShowErr = "<li>����Ĳ���</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../InfoManage.asp")
		Response.end
	else
		User_Conn.execute("Delete From FS_ME_InfoClass where UserNumber = '" & Fs_User.UserNumber & "' And ClassID="&CintStr(Request.QueryString("Id")))
		'����������Class����Ϣ
		strShowErr = "<li>ɾ���ɹ�</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../InfoManage.asp")
		Response.end
	end if
end if 
if request.Form("Action")<>"" then
	dim rs_svae
	set rs_svae= Server.CreateObject(G_FS_RS)
	if request.Form("Action")="add" then
		rs_svae.open "select * from FS_ME_InfoClass where UserNumber='"&Fs_User.UserNumber&"'",User_conn,1,3
		if not rs_svae.eof then
			if rs_svae.recordcount>cint(p_LimitClass) then
				strShowErr = "<li>����ཨ��"&p_LimitClass&"�����ࡣ</li>"
				Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../InfoManage.asp")
				Response.end
			end if
		end if
		rs_svae.addnew
		rs_svae("ClassEName") = NoSqlHack(Request.Form("ClassEName"))
		rs_svae("ParentID") = 0
		rs_svae("UserNumber") = Fs_User.UserNumber
		rs_svae("AddTime") = now
	elseif request.Form("Action")="edit" then
		rs_svae.open "select * from FS_ME_InfoClass where ClassID="&CintStr(Request.Form("ClassID")),User_conn,1,3
	end if
		rs_svae("ClassCName") = NoSqlHack(Request.Form("ClassCName"))
		If NoSqlHack(Trim(Request.Form("ClassTypes"))) = "" Then
			rs_svae("ClassTypes") = 0
		Else
			rs_svae("ClassTypes") = CintStr(Request.Form("ClassTypes"))
		End If
		rs_svae("ClassContent") = NoSqlHack(NoHtmlHackInput(Request.Form("ClassContent")))
		rs_svae.update
		strShowErr = "<li>����ɹ�</li>"
		Response.Redirect("lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../InfoManage.asp")
		Response.end
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>��ӭ�û�<%=Fs_User.UserName%>����<%=GetUserSystemTitle%>-ר������</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<body>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td class="hback"><strong>λ�ã�</strong><a href="../">��վ��ҳ</a> &gt;&gt; 
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; ר��</td>
        </tr>
      </table> 
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td class="Hback_1"><a href="infoManage.asp">ר����ҳ</a>��<a href="infomanage.asp?type=add">����ר��</a></td>
        </tr>
        <tr class="hback"> 
          <td class="hback">
		  <%
		  dim rs,class_list,tmp_k,s_type
		  class_list = "<table width=""100%"" cellpadding=""3""><tr>"
		  set rs = Server.CreateObject(G_FS_RS)
		  rs.open "select Classid,ClassTypes,ClassCName,ParentID,ClassContent,AddTime From FS_ME_InfoClass where UserNumber='"& Fs_User.UserNumber &"' and ParentID=0 order by ClassTypes asc,ClassID desc",User_Conn,1,3
		  if rs.eof then
			  class_list = class_list & "<td>ûר��</td>"
		  else
		  		tmp_k=0
		  		do	 while not rs.eof 
					select case rs("ClassTypes")
						case 0
							s_type = "(����)"
						case 3
							s_type = "(����)"
						case 4
							s_type = "(����)"
						case 5
							s_type = "(��ְ)"
						case 6
							s_type = "(��Ƹ)"
						case 7
							s_type = "(��־/��ժ)"
					end select
					class_list =  class_list & "<td width=""1"" align=right><img src=""Images/folderopened.gif"" ></td><td width=""33%"" valign=""bottom""><a href=""ShowInfoClass.asp?ClassID="&rs("Classid")&"&UserNumber="& Fs_User.UserNumber &""" title=""������"&rs("ClassContent")&"���������ڣ�"&rs("Addtime")&""">"&rs("ClassCName")&"</a><span class=""tx"">"&s_type&"</span>[<a href=Infomanage.asp?id="&rs("Classid")&"&type=edit>�޸�</a>]<<A href=Infomanage.asp?Id="&rs("Classid")&"&type=delinfo onClick=""{if(confirm('ȷ��ɾ������ѡ��ļ�¼��?')){return true;}return false;}"">ɾ��</a>></td>"
					rs.movenext
					tmp_k=tmp_k+1
					if tmp_k mod 3 = 0 then
						class_list =  class_list & "</tr>"
					end if
				loop
				class_list =  class_list & "</tr></table>"
		  end if
		  Response.Write class_list
		  %>

		  </td>
        </tr>
        <tr class="hback">
          <td class="hback"><span class="tx">�ر�˵����ר������Ϊ����������Ż���С˵�ȹ���������ʹ��</span></td>
        </tr>
      </table>
		 <%
	  		if Request.QueryString("type")="add" then
				call add()
			end if
			if request.QueryString("type")="edit" then
				call edit()
			end if
	    %>
		<%sub add()%>
        <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <form name="form1" method="post" action="InfoManage.asp" onSubmit="return checkdata();"><tr> 
            <td colspan="2" class="hback_1"><strong>����ר��</strong></td>
          </tr>
          <tr class="hback"> 
            <td width="19%"><div align="right">ר����������</div></td>
            <td width="81%"><input name="ClassCName" type="text" id="ClassCName" size="50" maxlength="50">
              * </td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">ר��Ӣ������</div></td>
            <td><input name="ClassEName" type="text" id="ClassEName" value="<%=Fs_User.UserNumber&"@"&month(now)&day(now)&"-"&hour(now)&minute(now)&second(now)%>" size="50" maxlength="50" readonly>
              * </td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">�û���</div></td>
            <td><%=Fs_User.UserName%></td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">����</div></td>
            <td><select name="ClassTypes" id="ClassTypes">
                <option value="0" selected>������ѧ��</option>
                <%if IsExist_SubSys("HS") Then%><option value="3">����</option><%end if%>
                <%if IsExist_SubSys("SD") Then%><option value="4">����</option><%end if%>
                <%if IsExist_SubSys("AP") Then%><option value="5">��ְ</option>
               <option value="6">��Ƹ</option><%end if%>
                <option value="7">��־/��ժ</option>
              </select>
            </td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">ר������</div></td>
            <td><textarea name="ClassContent" rows="6" id="ClassContent" style="width:60%"></textarea>
              500���ַ����� </td>
          </tr>
          <tr class="hback"> 
            <td>&nbsp;</td>
            <td><input type="submit" name="Submit" value=" ����ר�� ">
              <input name="Action" type="hidden" id="Action" value="add">
              <input type="reset" name="Submit2" value="����"></td>
          </tr></form> 
        </table>
      <script language="JavaScript" type="text/JavaScript">
	  <!--
	function checkdata()
	{
		if(f_trim(document.form1.ClassCName.value)=='')
		{
		alert('��д��������');
			form1.ClassCName.focus();
			return false;
		}
		if(document.form1.ClassEName.value=='')
		{
		alert('��дӢ������');
			form1.ClassEName.focus();
			return false;
		}
		if(document.form1.ClassContent.value=='')
		{
		alert('��д����');
			form1.ClassContent.focus();
			return false;
		}
		if(document.form1.ClassContent.value.length>500)
		{
		alert('�������500���ַ�');
			form1.ClassContent.focus();
			return false;
		}
	}
	-->
</script>

      <%end sub%>
		<%
		sub edit()
			dim edit_rs
			set edit_rs= Server.CreateObject(G_FS_RS)
			edit_rs.open "select * From FS_ME_InfoClass where UserNumber='"&Fs_User.UserNumber&"' and Classid="&CintStr(Request.QueryString("id")),User_conn,1,3
			if edit_rs.eof then
				strShowErr = "<li>�Ҳ�����¼</li>"
				Response.Redirect("lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../InfoManage.asp")
				Response.end
			else
		%><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <form name="form1" method="post" action="InfoManage.asp" onSubmit="return checkdata();"><tr> 
            <td colspan="2" class="hback_1"><strong>�޸�ר��
              <input name="ClassID" type="hidden" id="ClassID" value="<% = edit_rs("Classid")%>">
              </strong></td>
          </tr>
          <tr class="hback"> 
            <td width="19%"><div align="right">ר����������</div></td>
            <td width="81%"><input name="ClassCName" type="text" id="ClassCName" size="50" maxlength="50" value="<% = edit_rs("ClassCName")%>">
              * </td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">ר��Ӣ������</div></td>
            <td><input name="ClassEName" type="text" id="ClassEName" value="<% = edit_rs("ClassEName")%>" size="50" maxlength="50" readonly>
              * </td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">�û���</div></td>
            <td><%=Fs_User.UserName%></td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">����</div></td>
            <td><select name="ClassTypes" id="ClassTypes">
                <option value="0" <%if edit_rs("ClassTypes")=0 then response.Write("selected")%>>������ѧ��</option>
                <option value="3" <%if edit_rs("ClassTypes")=3 then response.Write("selected")%>>����</option>
                <option value="4" <%if edit_rs("ClassTypes")=4 then response.Write("selected")%>>����</option>
                <option value="5" <%if edit_rs("ClassTypes")=5 then response.Write("selected")%>>��ְ</option>
                <option value="6" <%if edit_rs("ClassTypes")=6 then response.Write("selected")%>>��Ƹ</option>
                <option value="7" <%if edit_rs("ClassTypes")=7 then response.Write("selected")%>>��־/��ժ</option>
              </select>
            </td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">ר������</div></td>
            <td><textarea name="ClassContent" rows="6" id="ClassContent" style="width:60%"><% = edit_rs("ClassContent")%></textarea>
              500���ַ����� </td>
          </tr>
          <tr class="hback"> 
            <td>&nbsp;</td>
            <td><input type="submit" name="Submit" value=" ����ר�� ">
              <input name="Action" type="hidden" id="Action" value="edit">
              <input type="reset" name="Submit2" value="����"></td>
          </tr></form> 
        </table>
		
      <script language="JavaScript" type="text/JavaScript">
	  <!--
	function checkdata()
	{
		if(f_trim(document.form1.ClassCName.value)=='')
		{
		alert('��д��������');
			form1.ClassCName.focus();
			return false;
		}
		if(document.form1.ClassEName.value=='')
		{
		alert('��дӢ������');
			form1.ClassEName.focus();
			return false;
		}
		if(document.form1.ClassContent.value=='')
		{
		alert('��д����');
			form1.ClassContent.focus();
			return false;
		}
		if(document.form1.ClassContent.value.length>500)
		{
		alert('�������500���ַ�');
			form1.ClassContent.focus();
			return false;
		}
	}
	-->
</script>
      <%
			end if
		end sub
		%>
			  </td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="Copyright.asp" -->
        </div></td>
    </tr>
 
</table>
</body>
</html>
<%
Set Fs_User = Nothing
%>
<script language="javascript">
//ȥ���ִ���ߵĿո� 
function lTrim(str) 
{ 
	if (str.charAt(0) == " ") 
	{ 
		//����ִ���ߵ�һ���ַ�Ϊ�ո� 
		str = str.slice(1);//���ո���ִ���ȥ�� 
		//��һ��Ҳ�ɸĳ� str = str.substring(1, str.length); 
		str = lTrim(str); //�ݹ���� 
	} 
	return str; 
} 

//ȥ���ִ��ұߵĿո� 
function rTrim(str) 
{ 
	var iLength; 

	iLength = str.length; 
	if (str.charAt(iLength - 1) == " ") 
	{ 
		//����ִ��ұߵ�һ���ַ�Ϊ�ո� 
		str = str.slice(0, iLength - 1);//���ո���ִ���ȥ�� 
		//��һ��Ҳ�ɸĳ� str = str.substring(0, iLength - 1); 
		str = rTrim(str); //�ݹ���� 
	} 
	return str; 
} 
//ȥ�����ҿո�
/*
����ֵ:ȥ�����ֵ
����˵��:_str,ԭֵ
*/
function f_trim(_str)
{
	return lTrim(rTrim(_str)); 
}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->






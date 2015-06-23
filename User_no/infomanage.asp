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
		strShowErr = "<li>错误的参数</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../InfoManage.asp")
		Response.end
	else
		User_Conn.execute("Delete From FS_ME_InfoClass where UserNumber = '" & Fs_User.UserNumber & "' And ClassID="&CintStr(Request.QueryString("Id")))
		'更新其他有Class的信息
		strShowErr = "<li>删除成功</li>"
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
				strShowErr = "<li>你最多建立"&p_LimitClass&"个分类。</li>"
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
		strShowErr = "<li>保存成功</li>"
		Response.Redirect("lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../InfoManage.asp")
		Response.end
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>欢迎用户<%=Fs_User.UserName%>来到<%=GetUserSystemTitle%>-专栏管理</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
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
          <td class="hback"><strong>位置：</strong><a href="../">网站首页</a> &gt;&gt; 
            <a href="main.asp">会员首页</a> &gt;&gt; 专栏</td>
        </tr>
      </table> 
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td class="Hback_1"><a href="infoManage.asp">专栏首页</a>┆<a href="infomanage.asp?type=add">创建专栏</a></td>
        </tr>
        <tr class="hback"> 
          <td class="hback">
		  <%
		  dim rs,class_list,tmp_k,s_type
		  class_list = "<table width=""100%"" cellpadding=""3""><tr>"
		  set rs = Server.CreateObject(G_FS_RS)
		  rs.open "select Classid,ClassTypes,ClassCName,ParentID,ClassContent,AddTime From FS_ME_InfoClass where UserNumber='"& Fs_User.UserNumber &"' and ParentID=0 order by ClassTypes asc,ClassID desc",User_Conn,1,3
		  if rs.eof then
			  class_list = class_list & "<td>没专栏</td>"
		  else
		  		tmp_k=0
		  		do	 while not rs.eof 
					select case rs("ClassTypes")
						case 0
							s_type = "(新闻)"
						case 3
							s_type = "(房产)"
						case 4
							s_type = "(供求)"
						case 5
							s_type = "(求职)"
						case 6
							s_type = "(招聘)"
						case 7
							s_type = "(日志/网摘)"
					end select
					class_list =  class_list & "<td width=""1"" align=right><img src=""Images/folderopened.gif"" ></td><td width=""33%"" valign=""bottom""><a href=""ShowInfoClass.asp?ClassID="&rs("Classid")&"&UserNumber="& Fs_User.UserNumber &""" title=""描述："&rs("ClassContent")&"┆创建日期："&rs("Addtime")&""">"&rs("ClassCName")&"</a><span class=""tx"">"&s_type&"</span>[<a href=Infomanage.asp?id="&rs("Classid")&"&type=edit>修改</a>]<<A href=Infomanage.asp?Id="&rs("Classid")&"&type=delinfo onClick=""{if(confirm('确定删除您所选择的记录吗？?')){return true;}return false;}"">删除</a>></td>"
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
          <td class="hback"><span class="tx">特别说明：专栏可以为您发表的新闻或者小说等归类做连载使用</span></td>
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
            <td colspan="2" class="hback_1"><strong>创建专栏</strong></td>
          </tr>
          <tr class="hback"> 
            <td width="19%"><div align="right">专栏中文名称</div></td>
            <td width="81%"><input name="ClassCName" type="text" id="ClassCName" size="50" maxlength="50">
              * </td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">专栏英文名称</div></td>
            <td><input name="ClassEName" type="text" id="ClassEName" value="<%=Fs_User.UserNumber&"@"&month(now)&day(now)&"-"&hour(now)&minute(now)&second(now)%>" size="50" maxlength="50" readonly>
              * </td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">用户名</div></td>
            <td><%=Fs_User.UserName%></td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">类型</div></td>
            <td><select name="ClassTypes" id="ClassTypes">
                <option value="0" selected>新闻文学类</option>
                <%if IsExist_SubSys("HS") Then%><option value="3">房产</option><%end if%>
                <%if IsExist_SubSys("SD") Then%><option value="4">供求</option><%end if%>
                <%if IsExist_SubSys("AP") Then%><option value="5">求职</option>
               <option value="6">招聘</option><%end if%>
                <option value="7">日志/网摘</option>
              </select>
            </td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">专栏描述</div></td>
            <td><textarea name="ClassContent" rows="6" id="ClassContent" style="width:60%"></textarea>
              500个字符以内 </td>
          </tr>
          <tr class="hback"> 
            <td>&nbsp;</td>
            <td><input type="submit" name="Submit" value=" 保存专栏 ">
              <input name="Action" type="hidden" id="Action" value="add">
              <input type="reset" name="Submit2" value="重置"></td>
          </tr></form> 
        </table>
      <script language="JavaScript" type="text/JavaScript">
	  <!--
	function checkdata()
	{
		if(f_trim(document.form1.ClassCName.value)=='')
		{
		alert('填写中文名称');
			form1.ClassCName.focus();
			return false;
		}
		if(document.form1.ClassEName.value=='')
		{
		alert('填写英文名称');
			form1.ClassEName.focus();
			return false;
		}
		if(document.form1.ClassContent.value=='')
		{
		alert('填写描述');
			form1.ClassContent.focus();
			return false;
		}
		if(document.form1.ClassContent.value.length>500)
		{
		alert('描述最多500个字符');
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
				strShowErr = "<li>找不到记录</li>"
				Response.Redirect("lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../InfoManage.asp")
				Response.end
			else
		%><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <form name="form1" method="post" action="InfoManage.asp" onSubmit="return checkdata();"><tr> 
            <td colspan="2" class="hback_1"><strong>修改专栏
              <input name="ClassID" type="hidden" id="ClassID" value="<% = edit_rs("Classid")%>">
              </strong></td>
          </tr>
          <tr class="hback"> 
            <td width="19%"><div align="right">专栏中文名称</div></td>
            <td width="81%"><input name="ClassCName" type="text" id="ClassCName" size="50" maxlength="50" value="<% = edit_rs("ClassCName")%>">
              * </td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">专栏英文名称</div></td>
            <td><input name="ClassEName" type="text" id="ClassEName" value="<% = edit_rs("ClassEName")%>" size="50" maxlength="50" readonly>
              * </td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">用户名</div></td>
            <td><%=Fs_User.UserName%></td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">类型</div></td>
            <td><select name="ClassTypes" id="ClassTypes">
                <option value="0" <%if edit_rs("ClassTypes")=0 then response.Write("selected")%>>新闻文学类</option>
                <option value="3" <%if edit_rs("ClassTypes")=3 then response.Write("selected")%>>房产</option>
                <option value="4" <%if edit_rs("ClassTypes")=4 then response.Write("selected")%>>供求</option>
                <option value="5" <%if edit_rs("ClassTypes")=5 then response.Write("selected")%>>求职</option>
                <option value="6" <%if edit_rs("ClassTypes")=6 then response.Write("selected")%>>招聘</option>
                <option value="7" <%if edit_rs("ClassTypes")=7 then response.Write("selected")%>>日志/网摘</option>
              </select>
            </td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">专栏描述</div></td>
            <td><textarea name="ClassContent" rows="6" id="ClassContent" style="width:60%"><% = edit_rs("ClassContent")%></textarea>
              500个字符以内 </td>
          </tr>
          <tr class="hback"> 
            <td>&nbsp;</td>
            <td><input type="submit" name="Submit" value=" 保存专栏 ">
              <input name="Action" type="hidden" id="Action" value="edit">
              <input type="reset" name="Submit2" value="重置"></td>
          </tr></form> 
        </table>
		
      <script language="JavaScript" type="text/JavaScript">
	  <!--
	function checkdata()
	{
		if(f_trim(document.form1.ClassCName.value)=='')
		{
		alert('填写中文名称');
			form1.ClassCName.focus();
			return false;
		}
		if(document.form1.ClassEName.value=='')
		{
		alert('填写英文名称');
			form1.ClassEName.focus();
			return false;
		}
		if(document.form1.ClassContent.value=='')
		{
		alert('填写描述');
			form1.ClassContent.focus();
			return false;
		}
		if(document.form1.ClassContent.value.length>500)
		{
		alert('描述最多500个字符');
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
//去掉字串左边的空格 
function lTrim(str) 
{ 
	if (str.charAt(0) == " ") 
	{ 
		//如果字串左边第一个字符为空格 
		str = str.slice(1);//将空格从字串中去掉 
		//这一句也可改成 str = str.substring(1, str.length); 
		str = lTrim(str); //递归调用 
	} 
	return str; 
} 

//去掉字串右边的空格 
function rTrim(str) 
{ 
	var iLength; 

	iLength = str.length; 
	if (str.charAt(iLength - 1) == " ") 
	{ 
		//如果字串右边第一个字符为空格 
		str = str.slice(0, iLength - 1);//将空格从字串中去掉 
		//这一句也可改成 str = str.substring(0, iLength - 1); 
		str = rTrim(str); //递归调用 
	} 
	return str; 
} 
//去除左右空格
/*
返回值:去除后的值
参数说明:_str,原值
*/
function f_trim(_str)
{
	return lTrim(rTrim(_str)); 
}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






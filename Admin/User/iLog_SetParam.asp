<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
dim Conn,User_Conn,rs,str_c_isp,str_c_user,str_c_pass,str_c_url,str_domain,rs_param,str_c_gurl,strShowErr
dim id,siteName,Domain,keywords,FileName,FileExtName,isCheck,isOpen,Kcontent,dir
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("ME_Log") then Err_Show 
if not MF_Check_Pop_TF("ME039") then Err_Show 

set rs = Server.CreateObject(G_FS_RS)
rs.open "select top 1 id,siteName,[Domain],dir,keywords,FileName,FileExtName,isCheck,isOpen,Kcontent From FS_ME_iLogSysParam",User_Conn,1,3
if rs.eof then
	id=""
	siteName="风讯日志"
	Domain="www.foosun.cn"
	keywords="风讯,CMS,Foosun,FoosunCMS"
	FileName=0
	FileExtName="html"
	isCheck=0
	isOpen=1
	Kcontent="我日,我靠,Fuck you,共产,中共"
	dir="blog"
else
	id=rs("id")
	siteName=rs("siteName")
	Domain=rs("Domain")
	keywords=rs("keywords")
	FileName=rs("FileName")
	FileExtName=rs("FileExtName")
	isCheck=rs("isCheck")
	isOpen=rs("isOpen")
	Kcontent=rs("Kcontent")
	dir=rs("dir")
end if
rs.close:set rs=nothing
if Request.Form("Action")="save" then
	if Request.Form("siteName")="" or Request.Form("Domain")="" or Request.Form("dir")="" then
		strShowErr = "<li>带*号的必须填写</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	set rs = Server.CreateObject(G_FS_RS)
	rs.open "select top 1 id,siteName,[Domain],dir,keywords,FileName,FileExtName,isCheck,isOpen,Kcontent From FS_ME_iLogSysParam",User_Conn,3,3
	if trim(Request.Form("id"))="" then
		rs.addnew
	end if
	rs("siteName")=NoSqlHack(Request.Form("siteName"))
	rs("Domain")=NoSqlHack(Request.Form("Domain"))
	rs("keywords")=NoSqlHack(Request.Form("keywords"))
	rs("FileName")=CintStr(Request.Form("FileName"))
	rs("FileExtName")=NoSqlHack(Request.Form("FileExtName"))
	if Request.Form("isCheck")<>"" then:rs("isCheck")=1:else:rs("isCheck")=0:end if
	if Request.Form("isOpen")<>"" then:rs("isOpen")=1:else:rs("isOpen")=0:end if
	rs("Kcontent")=NoSqlHack(Request.Form("Kcontent"))
	rs("dir")=NoSqlHack(Request.Form("dir"))
	rs.update
	rs.close:set rs=nothing
	set conn=nothing
	set user_conn=nothing
	strShowErr = "<li>更新成功</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog_SetParam.asp")
	Response.end
end if
%>
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr> 
    <td width="100%" class="xingmu">日志网摘管理</td>
  </tr>
  <tr> 
    <td class="hback"><a href="iLog.asp">日志管理</a>┆<a href="iLog_Templet.asp">模板设置</a>┆<a href="iLog_Class.asp">系统栏目</a>┆<a href="iLog_SetParam.asp">参数设置</a></td>
  </tr>
</table>
  
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="form1" method="post" action="">
    <tr> 
      <td width="20%" class="hback"><div align="right">站点名称</div></td>
      <td width="80%" class="hback"><input name="siteName" type="text" id="siteName" value="<%=siteName%>" size="30">
        <input name="id" type="hidden" id="id" value="<%=id%>">
        *</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">捆绑域名</div></td>
      <td class="hback"><input name="domain" type="text" id="domain" value="<%=domain%>" size="30">
        *</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">日志关键字</div></td>
      <td class="hback"><input name="keywords" type="text" id="keywords" value="<%=keywords%>" size="30">
        100个字符</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">日志前台目录</div></td>
      <td class="hback"><input name="dir" type="text" id="dir" value="<%=dir%>" size="30">
        *</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">生成静态文件文件名</div></td>
      <td class="hback"><input name="FileName" type="radio" value="0" <%if FileName=0 then response.Write"checked"%>>
        8位随机数　 <input type="radio" name="FileName" value="1" <%if FileName=1 then response.Write"checked"%>>
        ID　 
        <input type="radio" name="FileName" value="2" <%if FileName=2 then response.Write"checked"%>>
        时间*</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">扩展名</div></td>
      <td class="hback"><select name="FileExtName" id="FileExtName">
          <option value="html" <%if FileExtName="html" then response.Write"selected"%>>html</option>
          <option value="htm" <%if FileExtName="htm" then response.Write"selected"%>>htm</option>
          <option value="shtml" <%if FileExtName="shtml" then response.Write"selected"%>>shtml</option>
          <option value="shtm" <%if FileExtName="shtm" then response.Write"selected"%>>shtm</option>
          <option value="asp" <%if FileExtName="asp" then response.Write"selected"%>>asp</option>
        </select></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">发表日志需要审核</div></td>
      <td class="hback"><input name="isCheck" type="checkbox" id="isCheck" value="1" <%if isCheck=1 then response.Write"checked"%>></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">开放日志</div></td>
      <td class="hback"><input name="isOpen" type="checkbox" id="isOpen" value="1" <%if isOpen=1 then response.Write"checked"%>></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">敏感字过滤管理</div></td>
      <td class="hback"><textarea name="Kcontent" rows="6" id="Kcontent" style="width:80%"><%=Kcontent%></textarea></td>
    </tr>
    <tr> 
      <td class="hback">&nbsp;</td>
      <td class="hback"><input type="submit" name="Submit" value="保存参数">
        <input name="Action" type="hidden" id="Action" value="save"></td>
    </tr>
  </form>
</table>

</body>
</html>
<%
Conn.close:set conn=nothing
User_Conn.close:set User_Conn=nothing
%>






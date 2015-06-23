<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
If CheckBlogOpen=False Then 
	Response.write("<script language=""javascript"">alert('日志功能暂停使用,如需要使用请联系管理员.');history.back();</script>")
	Response.End()
End If
Dim rs,ClassID
Dim siteName,ShowContentNumber,ShowReviewNumber,ShowLogNumber,Content,ShowNewNumber,TempletID
set rs= Server.CreateObject(G_FS_RS)
rs.open "select * From FS_ME_InfoilogParam where UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
if not rs.eof then
	siteName = rs("siteName")
	ClassID = rs("ClassID")
	Content = rs("Content")
	TempletID=rs("TempletID")
else
	siteName = Fs_User.UserName&"的站点"
	ClassID = 1
	Content = Fs_User.UserName&"的站点"
	TempletID=2
end if
rs.close:set rs=nothing
if Request.Form("Action")="Edit" then
	set rs= Server.CreateObject(G_FS_RS)
	rs.open "select * From FS_ME_InfoilogParam where UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
	if rs.eof then
		rs.addnew
	end if
	rs("siteName")= NoSqlHack(Request.Form("siteName"))
	rs("ClassID") = CintStr(Request.Form("ClassID"))
	rs("Content") = NoSqlHack(NoHtmlHackInput(Request.Form("Content")))
	rs("UserNumber")=Fs_User.UserNumber
	rs("TempletID")=NoSqlHack(Request.Form("TempletID"))
	rs.update
	rs.close:set rs=nothing
	dim templet_rs,u_tmplet_savepath
	'set templet_rs = User_Conn.execute("select id,TempletSavePath From FS_ME_InfoiLogTemplet where id="&cint(Request.Form("TempletID")))
	u_tmplet_savepath = NoSqlHack(Request.Form("TempletID"))
	'templet_rs.close:set templet_rs = nothing
	dim u_FSO,u_savePath,rs_sys_blogobj,u_dir,u_FileExtName,u_savePath_index,u_savepath_list,u_savepath_page,u_savepath_photo,oStream
	set rs_sys_blogobj = User_Conn.execute("select Top 1 Dir,FileExtName From FS_ME_iLogSysParam")
	if rs_sys_blogobj.eof then
		response.Write("找不到系统配置信息，请与管理员联系。创建目录失败")
		response.End
		rs_sys_blogobj.close:set rs_sys_blogobj = nothing
	else
		u_dir = rs_sys_blogobj(0)
		u_FileExtName = rs_sys_blogobj(1)
		rs_sys_blogobj.close:set rs_sys_blogobj = nothing
	end if
'	set U_FSO=server.CreateObject(G_FS_FSO)
'	u_savePath = server.MapPath("../../"& u_dir & "/" & Fs_User.UserNumber)
'	u_savePath_index = server.MapPath("../../"& u_dir & "/" & Fs_User.UserNumber & "/index."& u_FileExtName &"")
'	u_savepath_list = server.MapPath("../../"& u_dir & "/" & Fs_User.UserNumber & "/list."& u_FileExtName &"")
'	u_savepath_page = server.MapPath("../../"& u_dir & "/" & Fs_User.UserNumber & "/page."& u_FileExtName &"")
'	u_savepath_photo = server.MapPath("../../"& u_dir & "/" & Fs_User.UserNumber & "/photo."& u_FileExtName &"")
'	if U_FSO.FolderExists(u_savePath) = False Then U_FSO.CreateFolder(u_savePath)
'	dim File_Obj,FileStreamObj,FileContent
'	if U_FSO.FileExists(u_savePath_index) = False Then
'		Set oStream = U_FSO.CreateTextFile(u_savePath_index, false)
'		Set File_Obj = U_FSO.GetFile(server.MapPath(replace("/"&G_VIRTUAL_ROOT_DIR&u_tmplet_savepath&"/index.htm","//","/")))
'		Set FileStreamObj = File_Obj.OpenAsTextStream(1)
'		FileContent = FileStreamObj.ReadAll
'		oStream.Write FileContent
'		oStream.Close
'		Set oStream = Nothing
'	end if
'	if U_FSO.FileExists(u_savepath_list) = False Then
'		Set oStream = U_FSO.CreateTextFile(u_savepath_list, false)
'		'U_FSO.CreateFile(u_savepath_list)
'		Set File_Obj = U_FSO.GetFile(server.MapPath(replace("/"&G_VIRTUAL_ROOT_DIR&u_tmplet_savepath&"/list.htm","//","/")))
'		Set FileStreamObj = File_Obj.OpenAsTextStream(1)
'		FileContent = FileStreamObj.ReadAll
'		oStream.Write FileContent
'		oStream.Close
'		Set oStream = Nothing
'	end if
'	if U_FSO.FileExists(u_savepath_page) = False Then
'		Set oStream = U_FSO.CreateTextFile(u_savepath_page, false)
'		Set File_Obj = U_FSO.GetFile(server.MapPath(replace("/"&G_VIRTUAL_ROOT_DIR&u_tmplet_savepath&"/page.htm","//","/")))
'		Set FileStreamObj = File_Obj.OpenAsTextStream(1)
'		FileContent = FileStreamObj.ReadAll
'		oStream.Write FileContent
'		oStream.Close
'		Set oStream = Nothing
'	end if
'	if U_FSO.FileExists(u_savepath_photo) = False Then
'		Set oStream = U_FSO.CreateTextFile(u_savepath_photo, false)
'		Set File_Obj = U_FSO.GetFile(server.MapPath(replace("/"&G_VIRTUAL_ROOT_DIR&u_tmplet_savepath&"/photo.htm","//","/")))
'		Set FileStreamObj = File_Obj.OpenAsTextStream(1)
'		FileContent = FileStreamObj.ReadAll
'		oStream.Write FileContent
'		oStream.Close
'		Set oStream = Nothing
'	end if
'	set U_FSO = nothing
	strShowErr = "<li>保存参数成功！</li>"
	Response.Redirect("../lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../i_Blog/PublicParam.asp")
	Response.end
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head></head>
<body>

<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="../top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="../Top_navi.asp" -->
    </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="../menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback">
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td  valign="top">你的位置：<a href="../../">网站首页</a> &gt;&gt; <a href="../main.asp">会员首页</a> 
            &gt;&gt; <a href="index.asp">日志管理</a> &gt;&gt;日志管理</td>
        </tr>
        <tr class="hback"> 
          <td  valign="top"><a href="index.asp">日志首页</a>┆<a href="PublicLog.asp">发表日志</a>┆<a href="index.asp?type=box">草稿箱</a>┆<a href="../PhotoManage.asp">相册管理</a>┆<a href="PublicParam.asp">参数设置</a>┆<a href="../Review.asp">评论管理</a></td>
        </tr>
      </table>
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="s_form" method="post" action="">
          <tr> 
            <td colspan="2" class="xingmu">参数设置</td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">站点标题</div></td>
            <td class="hback"><input name="siteName" type="text" id="siteName" onFocus="Do.these('siteName',function(){return CheckContentLen('siteName','span_siteName','3-50')})" onKeyUp="Do.these('siteName',function(){return CheckContentLen('siteName','span_siteName','3-50')})" value="<%=siteName%>" size="57" maxlength="50"> 
              <span id="span_siteName"></span> </td>
          </tr>
          <tr> 
            <td width="22%" class="hback"><div align="right">站点类别</div></td>
            <td width="78%" class="hback"><select name="ClassID" id="ClassID">
              <!--<option value="">选择系统分类</option>-->
              <%
				dim c_rs
				set c_rs = Server.CreateObject(G_FS_RS)
				c_rs.open "select ID,ClassName From FS_ME_iLogClass Order by id asc",User_Conn,1,3
				do while not c_rs.eof
				if c_rs("Id")=ClassID then
				%>
              <option value="<%=c_rs("id")%>" selected="selected"><%=c_rs("ClassName")%></option>
			  <%else%>
              <option value="<%=c_rs("id")%>"><%=c_rs("ClassName")%></option>
			  <%end if%>
               <%
				c_rs.movenext
				loop
				c_rs.close:set c_rs=nothing
				%>
            </select></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">站点描述</div></td>
            <td class="hback"><textarea name="Content" rows="12" style="width:80%"><%=Content%></textarea></td>
          </tr>
          <tr>
            <td class="hback"><div align="right">模板选择</div></td>
            <td class="hback">
			 <select name="TempletID" id="TempletID">
			<%
			set rs = User_Conn.execute("select id,TempletName,TempletSavePath From FS_ME_InfoiLogTemplet Order by id desc")
			do while not rs.eof
				if TempletID=rs("TempletSavePath") then
					response.Write"<option value="""&rs("TempletSavePath")&""" selected>"&rs("TempletName")&"</option>"
				else
					response.Write"<option value="""&rs("TempletSavePath")&""">"&rs("TempletName")&"</option>"
				end if
				rs.movenext
			loop
			rs.close:set rs=nothing
			%>
              </select>
			 　　预览模板
			 <select name="select_s" id="select_s">
             <%
			set rs = User_Conn.execute("select TempletSavePath,TempletName From FS_ME_InfoiLogTemplet Order by id desc")
			do while not rs.eof
				response.Write"<option value="""&rs("TempletSavePath")&""">"&rs("TempletName")&"</option>"
				rs.movenext
			loop
			rs.close:set rs=nothing
			%>
             </select>
			 <input name="button3" type="button" id="button" onClick="showModalDialog(''+document.s_form.select_s.value+'/Index.htm','WindowObj','dialogWidth:600pt;dialogHeight:500pt;status:yes;help:no;scroll:yes;');" value="查看"></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right"></div></td>
            <td class="hback"><input name="Action" type="hidden" id="Action" value="Edit"> 
              <input type="submit" name="Submit" value="保存站点设置"> </td>
          </tr>
        </form>
      </table>
      </td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="../Copyright.asp" -->
        </div></td>
    </tr>
</table>
</body>
</html>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






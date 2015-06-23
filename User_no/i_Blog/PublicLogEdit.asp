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
Dim str_CurrPath,rs
str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
set rs= Server.CreateObject(G_FS_RS)
rs.open "select * From FS_ME_Infoilog where iLogID="&CintStr(Request.QueryString("id")),User_Conn,1,3
if rs.eof then
	strShowErr = "<li>错误的参数！</li>"
	Response.Redirect("../lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../i_Blog/index.asp")
	Response.end
	rs.close:set rs = nothing
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
<script language="JavaScript" src="../../Editor/FS_scripts/editor.js"></script>
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
        <form name="s_form" method="post" action="Public_Save.asp">
          <tr> 
            <td colspan="2" class="xingmu">发表日志</td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">日志标题</div></td>
            <td class="hback"><input name="Title" type="text" id="Title" size="57" value="<% = rs("Title")%>" maxlength="100" onFocus="Do.these('Title',function(){return CheckContentLen('Title','span_Title','3-100')})" onKeyUp="Do.these('Title',function(){return CheckContentLen('Title','span_Title','3-100')})"> 
              <span id="span_Title"></span></td>
          </tr>
          <tr> 
            <td width="13%" class="hback"><div align="right">主分类</div></td>
            <td width="87%" class="hback">
			<select name="MainID" id="MainID">
                <!--<option value="">选择系统分类</option>-->
                <%
				dim c_rs
				set c_rs = Server.CreateObject(G_FS_RS)
				c_rs.open "select ID,ClassName From FS_ME_iLogClass Order by id asc",User_Conn,1,3
				do while not c_rs.eof
				if rs("MainId")=c_rs("id") then
				%>
                <option value="<%=c_rs("id")%>" selected><%=c_rs("ClassName")%></option>
				<%else%>
                <option value="<%=c_rs("id")%>"><%=c_rs("ClassName")%></option>
                <%
				end if
				c_rs.movenext
				loop
				c_rs.close:set c_rs=nothing
				%>
              </select>
              专栏 
              <select name="ClassID" id="ClassID" style="width:38%">
                <option value="0">选择我的专栏</option>
                <%
			  dim obj_zl_rs
			  set obj_zl_rs= Server.CreateObject(G_FS_RS)
			  obj_zl_rs.open "select ClassID,ClassCName,ClassEName,ParentID,UserNumber,ClassTypes From FS_ME_InfoClass  where ParentID=0 and ClassTypes=7 and UserNumber='"& Fs_User.UserNumber&"' order by ClassID desc",User_Conn,1,3
			  do while not obj_zl_rs.eof 
			  		if rs("ClassID")=obj_zl_rs("ClassID") then
						Response.Write"<option value="""& obj_zl_rs("ClassID")&""" selected>"& obj_zl_rs("ClassCName")&"</option>"
					else
						Response.Write"<option value="""& obj_zl_rs("ClassID")&""">"& obj_zl_rs("ClassCName")&"</option>"
					end if
				  obj_zl_rs.movenext
			  Loop
			  obj_zl_rs.close:set obj_zl_rs =nothing
			  %>
              </select></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">类型</div></td>
            <td class="hback"><select name="iLogStyle" id="iLogStyle">
                <option value="0" <%if rs("iLogStyle")=0 then response.Write"selected"%>>日志</option>
                <option value="1" <%if rs("iLogStyle")=1 then response.Write"selected"%>>网摘</option>
              </select> </td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">关键字</div></td>
            <td class="hback"><input name="keyword1" type="text" id="keyword1" size="15" maxlength="15" value="<%=split(rs("Keywords"),",")(0)%>">
              , 
              <input name="keyword2" type="text" id="keyword2" size="15" maxlength="15" value="<%=split(rs("Keywords"),",")(1)%>">
              , 
              <input name="keyword3" type="text" id="keyword3" size="15" maxlength="15" value="<%=split(rs("Keywords"),",")(2)%>"></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">来源</div></td>
            <td class="hback"><input name="iLogSource" type="text" id="iLogSource" size="57" maxlength="60" value="<%=rs("iLogSource")%>"></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">内容</div></td>
            <td class="hback">
						<!--编辑器开始-->
						<iframe id='NewsContent' src='../Editer/UserEditer.asp?id=Content' frameborder=0 scrolling=no width='100%' height='280'></iframe>
						<input type="hidden" name="Content" value="<% = HandleEditorContent(rs("Content")) %>">
						<!--编辑器结束-->
						</td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">首页置顶</div></td>
            <td class="hback"><input name="isTop" type="checkbox" id="isTop" value="1" <%if rs("isTop")=1 then response.Write("checked")%>></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">心情</div></td>
            <td class="hback"><input name="EmotFace" type="radio" value="face1.gif" checked> 
              <img src="../../sys_images/emot/face1.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face3.gif" <%if rs("EmotFace")="face3.gif" then response.Write("checked")%>> 
              <img src="../../sys_images/emot/face3.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face5.gif" <%if rs("EmotFace")="face5.gif" then response.Write("checked")%>> 
              <img src="../../sys_images/emot/face5.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face8.gif" <%if rs("EmotFace")="face8.gif" then response.Write("checked")%>> 
              <img src="../../sys_images/emot/face8.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face18.gif" <%if rs("EmotFace")="face18.gif" then response.Write("checked")%>> 
              <img src="../../sys_images/emot/face18.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face11.gif" <%if rs("EmotFace")="face11.gif" then response.Write("checked")%>> 
              <img src="../../sys_images/emot/face11.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face47.gif" <%if rs("EmotFace")="face47.gif" then response.Write("checked")%>> 
              <img src="../../sys_images/emot/face47.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face15.gif" <%if rs("EmotFace")="face15.gif" then response.Write("checked")%>> 
              <img src="../../sys_images/emot/face15.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face22.gif" <%if rs("EmotFace")="face22.gif" then response.Write("checked")%>> 
              <img src="../../sys_images/emot/face22.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face29.gif" <%if rs("EmotFace")="face29.gif" then response.Write("checked")%>> 
              <img src="../../sys_images/emot/face29.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face34.gif" <%if rs("EmotFace")="face34.gif" then response.Write("checked")%>> 
              <img src="../../sys_images/emot/face34.gif" width="19" height="19"></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">图片地址</div></td>
            <td class="hback"><table width="100%" border="0" cellspacing="1" cellpadding="5">
                <tr> 
                  <td width="29%"><div align="center"> 
                      <table width="10" border="0" cellspacing="1" cellpadding="2" class="table">
                        <tr> 
                          <%if not isnull(rs("pic_1")) then%>
                          <td class="hback"><img src="<%=rs("pic_1")%>" width="90" height="90" id="pic_p_1"></td>
                          <%else%>
                          <td class="hback"><img src="../Images/nopic_supply.gif" width="90" height="90" id="pic_p_1"></td>
                          <%end if%>
                        </tr>
                      </table>
                      <input name="pic_1" type="hidden" id="pic_1" size="40" value="<%=rs("pic_1")%>">
                    </div></td>
                  <td width="36%"><div align="center"> 
                      <table width="10" border="0" cellspacing="1" cellpadding="2" class="table">
                        <tr> 
                          <%if not isnull(rs("pic_2")) then%>
                          <td class="hback"><img src="<%=rs("pic_2")%>" width="90" height="90" id="pic_p_2"></td>
                          <%else%>
                          <td class="hback"><img src="../Images/nopic_supply.gif" width="90" height="90" id="pic_p_2"></td>
                          <%end if%>
                        </tr>
                      </table>
                      <input name="pic_2" type="hidden" id="pic_2" size="40" value="<%=rs("pic_2")%>">
                    </div></td>
                  <td width="35%"><div align="center"> 
                      <table width="10" border="0" cellspacing="1" cellpadding="2" class="table">
                        <tr> 
                          <%if not isnull(rs("pic_3")) then%>
                          <td class="hback"><img src="<%=rs("pic_3")%>" width="90" height="90" id="pic_p_3"></td>
                          <%else%>
                          <td class="hback"><img src="../Images/nopic_supply.gif" width="90" height="90" id="pic_p_3"></td>
                          <%end if%>
                        </tr>
                      </table>
                      <input name="pic_3" type="hidden" id="pic_3" size="40" value="<%=rs("pic_3")%>">
                    </div></td>
                </tr>
                <tr> 
                  <td><div align="center"><img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.s_form.pic_1);" style="cursor:hand;"> 
                      　<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_1();" style="cursor:hand;"> 
                    </div></td>
                  <td><div align="center"><img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.s_form.pic_2);" style="cursor:hand;"> 
                      　<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_2();" style="cursor:hand;"> 
                    </div></td>
                  <td><div align="center"><img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.s_form.pic_3);" style="cursor:hand;"> 
                      　<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_3();" style="cursor:hand;"> 
                    </div></td>
                </tr>
              </table></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">日志密码</div></td>
            <td class="hback"><input name="Password" type="password" id="Password">
              不修改请保持为空</td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">文件名</div></td>
            <td class="hback"><input name="FileName" type="text" id="FileName" value="<%=rs("FileName")%>" readonly maxlength="20"  onFocus="Do.these('FileName',function(){return CheckContentLen('FileName','span_FileName','3-18')})" onKeyUp="Do.these('FileName',function(){return CheckContentLen('FileName','span_FileName','3-18')})">
              <span id="span_FileName"></span> 扩展名 
              <select name="FileExtName" id="FileExtName">
                <option value="html" <%if rs("FileExtName")="html" then response.Write("selected")%>>html</option>
                <option value="htm" <%if rs("FileExtName")="htm" then response.Write("selected")%>>htm</option>
                <option value="shtm" <%if rs("FileExtName")="shtm" then response.Write("selected")%>>shtm</option>
                <option value="shtml" <%if rs("FileExtName")="shtml" then response.Write("selected")%>>shtml</option>
                <option value="asp" <%if rs("FileExtName")="asp" then response.Write("selected")%>>asp</option>
              </select></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td class="hback"><div align="right">保存到草稿箱</div></td>
            <td class="hback"><input name="isDraft" type="checkbox" id="isDraft" value="1" <%if rs("isDraft")=1 then response.Write("checked")%>></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right"></div></td>
            <td class="hback"><input name="Action" type="hidden" id="Action" value="Edit"> 
              <input name="Id" type="hidden" id="Id" value="<%=rs("iLogID")%>"> 
							<input type="button" name="button1" value="修改"  onClick="CheckSubmit(this.form);">
							</td>
          </tr>
        </form>
      </table>
      <%rs.close:set rs=nothing%>
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
<script language="JavaScript" type="text/JavaScript">
function CheckSubmit(FormObj){
	if (frames["NewsContent"].g_currmode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return false;}
	FormObj.Content.value=frames["NewsContent"].GetNewsContentArray();
	FormObj.submit();
	return true;
}
new Form.Element.Observer($('pic_1'),1,pics_1);
	function pics_1()
		{
			if ($('pic_1').value=='')
			{
				$('pic_p_1').src='../Images/nopic_supply.gif'
			}
			else
			{
			$('pic_p_1').src=$('pic_1').value
			}
		} 
new Form.Element.Observer($('pic_2'),1,pics_2);
	function pics_2()
		{
			if($('pic_2').value=='')
			{
			$('pic_p_2').src='../Images/nopic_supply.gif'
			}
			else
			{
			$('pic_p_2').src=$('pic_2').value
			}
		} 
new Form.Element.Observer($('pic_3'),1,pics_3);
	function pics_3()
		{
			if($('pic_3').value=='')
			{
			$('pic_p_3').src='../Images/nopic_supply.gif'
			}
			else
			{
			$('pic_p_3').src=$('pic_3').value
			}
		}
function dels_1()
	{
		document.s_form.pic_1.value=''
	}
function dels_2()
	{
		document.s_form.pic_2.value=''
	}
function dels_3()
	{
		document.s_form.pic_3.value=''
	}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp" -->
<%
Dim str_CurrPath,rs
str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
set rs= Server.CreateObject(G_FS_RS)
rs.open "select * From FS_ME_Photo where UserNumber='"&Fs_User.UserNumber&"' and id="&CintStr(Request.QueryString("id")),User_Conn,1,3
if rs.eof then
	rs.close:set rs=nothing
	strShowErr="<li>错误的参数</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../PhotoManage.asp")
	Response.end
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-相册管理</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt; <a href="PhotoManage.asp">相册管理</a> 
            &gt;&gt;修改图片</td>
        </tr>
      </table> 
		  
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr> 
          <td class="hback"><a href="PhotoManage.asp">相册首页</a>┆<a href="Photo_add.asp">增加图片</a>┆<a href="PhotoManage.asp?isRec=1">被推荐的图片</a>┆<a href="Photo_filt.asp">幻灯片播放</a>┆<a href="Photo_Class.asp">相册分类</a></td>
        </tr>
      </table>
      
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="s_form" method="post" action="Photo_save.asp" onSubmit="return checkinput();">
          <tr> 
            <td colspan="2" class="xingmu">增加相册</td>
          </tr>
          <tr> 
            <td width="18%" class="hback"> <div align="right">相片标题：</div></td>
            <td width="82%" class="hback"><input name="title" type="text" id="title" value="<%=rs("title")%>" size="45"></td>
          </tr>
          <tr> 
            <td class="hback"><div align="right">图片：</div></td>
            <td class="hback"><table width="27%" border="0" cellspacing="1" cellpadding="5">
                <tr> 
                  <td width="33%"><div align="center"> 
                      <table width="10" border="0" cellspacing="1" cellpadding="2" class="table">
                        <tr> 
                          <td class="hback"><img src="<%=rs("PicSavePath")%>" width="90" border="0" id="pic_p_1"></td>
                        </tr>
                      </table>
                      <input name="pic_1" type="hidden" id="pic_1" value="<%=rs("PicSavePath")%>" size="40" >
                    </div></td>
                </tr>
                <tr> 
                  <td><div align="center"><img  src="Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.s_form.pic_1);" style="cursor:hand;"> 
                      　<img src="Images/del_supply.gif" width="44" height="22" onClick="dels_1();" style="cursor:hand;"> 
                    </div></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td class="hback"><div align="right">类别：</div></td>
            <td class="hback"><select name="Classid">
                <option value="0">选择相册分类</option>
                <%
				dim srs
				set srs=User_Conn.execute("select id,title From FS_ME_PhotoClass where UserNumber='"&session("FS_UserNumber")&"' order by id desc")
				do while not srs.eof
						if rs("Classid")=srs("id") then
							response.Write"		<option value="""&srs("id")&""" selected>"&srs("title")&"</option>"&chr(13)
						else
							response.Write"		<option value="""&srs("id")&""">"&srs("title")&"</option>"&chr(13)
						end if
					srs.movenext
				loop
				srs.close:set srs=nothing
				%>
              </select></td>
          </tr>
          <tr> 
            <td class="hback"><div align="right">图片说明：</div></td>
            <td class="hback"><textarea name="content" rows="8" id="content" style="width:80%"><%=rs("content")%></textarea></td>
          </tr>
          <tr> 
            <td class="hback"><div align="right"></div></td>
            <td class="hback"><input type="submit" name="Submit" value="保存图片到相册">
              <input name="Action" type="hidden" id="Action" value="edit">
              <input name="id" type="hidden" id="id" value="<%=rs("id")%>"></td>
          </tr>
        </form>
      </table>
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
set rs=nothing
Set Fs_User = Nothing
%>
<script language="JavaScript" type="text/JavaScript">
new Form.Element.Observer($('pic_1'),1,pics_1);
	function pics_1()
		{
			if ($('pic_1').value=='')
			{
				$('pic_p_1').src='Images/nopic_supply.gif'
			}
			else
			{
			$('pic_p_1').src=$('pic_1').value
			}
		} 

function dels_1()
	{
		document.s_form.pic_1.value=''
	}
function checkinput()
{
	if(document.s_form.title.value=='')
	{
		alert('填写相册标题');
		document.s_form.title.focus();
		return false;
	}
	if(document.s_form.pic_1.value=='')
	{
		alert('填写至少输入一个图片地址');
		//document.s_form.pic_1.focus();
		return false;
	}
	if(document.s_form.content.value=='')
	{
		alert('填写图片描述');
		document.s_form.content.focus();
		return false;
	}
}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






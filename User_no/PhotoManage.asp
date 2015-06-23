<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp" -->
<%
if request("Action")="del" then
	if Request("id")="" then
		strShowErr = "<li>错误的参数！</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		User_Conn.execute("Delete from FS_ME_Photo where ID in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>删除成功！</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../PhotoManage.asp")
		Response.end
	end if
end if
if request("Action")="delall" then
	if Request("chkall")="" then
		strShowErr = "<li>错误的参数！</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		User_Conn.execute("Delete from FS_ME_PhotoClass where UserNumber='"&Fs_User.UserNumber&"'")
		User_Conn.execute("Delete from FS_ME_Photo")
		strShowErr = "<li>删除成功！</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../PhotoManage.asp")
		Response.end
	end if
	chkall
end if
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,rs,i
int_RPP=15 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"
strpage=request("page")
'if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-相册管理</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" href="lib/css/lightbox.css" type="text/css" media="screen" />
<script type="text/javascript" src="../FS_INC/prototype.js"></script>
<script type="text/javascript" src="lib/js/scriptaculous.js?load=effects"></script>
<script type="text/javascript" src="lib/js/lightbox.js"></script>
<head>
<body onLoad="initLightbox()">
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
            &gt;&gt;</td>
        </tr>
      </table> 
		  
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr>
          <td class="hback"><a href="PhotoManage.asp">相册首页</a>┆<a href="Photo_add.asp">增加图片</a>┆<a href="PhotoManage.asp?isRec=1">被推荐的图片</a>┆<a href="Photo_filt.asp">幻灯片播放</a>┆<a href="Photo_Class.asp">相册分类</a></td>
        </tr>
        <tr> 
          <td class="hback"> 
            <%
		  response.Write("	<table width=""98%"" align=center cellpadding=""2"" cellspacing=""1""><tr>")
		  dim t_k,rec_str
		  t_k=0
		  set rs = Server.CreateObject(G_FS_RS)
		  rs.open "select id,title,UserNumber From FS_ME_PhotoClass where UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
		  do while not rs.eof 
		  	Response.Write("	<td width=""24%"" valign=bottom><img src=""images/folderopened.gif""></img><a href=PhotoManage.asp?classid="&rs("id")&">"&rs("title")&"</a></td>")
		  rs.movenext
		  t_k = t_k+1
		  if t_k mod 4 =0 then
		  	Response.Write("	</tr>")
		  end if
		  loop
		  response.Write("	</table>")
		  rs.close:set rs=nothing
		  %>
          </td>
        </tr>
      </table> 
      
        <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <form name="myform" method="post" action="">
		  <%
		  dim o_class
		  if NosqlHack(Request.QueryString("Classid"))<>"" then
		  	if not isnumeric(NosqlHack(Request.QueryString("Classid"))) then
				strShowErr = "<li>错误的参数！</li>"
				Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			end if
		  	o_class=" and classid="&CintStr(Request("Classid"))&""
		  else
		  	o_class =""
		  end if
		  if NoSqlHack(Request.QueryString("isRec"))="1" then
		  	 rec_str = " and isRec=1"
		  else
		  	 rec_str = ""
		  end if
			set rs = Server.CreateObject(G_FS_RS)
			rs.open "select * from FS_ME_Photo where UserNumber='"& Fs_User.UserNumber&"' "&o_class&rec_str&" order by id desc",User_Conn,1,1
			if rs.eof then
			   rs.close
			   set rs=nothing
			   Response.Write"<tr  class=""hback""><td colspan=""1""  class=""hback"" height=""40"">没有记录。</td></tr>"
			else
				rs.PageSize=int_RPP
				cPageNo=NoSqlHack(Request.QueryString("Page"))
				If cPageNo="" Then cPageNo = 1
				If not isnumeric(cPageNo) Then cPageNo = 1
				cPageNo = Clng(cPageNo)
				If cPageNo>rs.PageCount Then cPageNo=rs.PageCount 
				If cPageNo<=0 Then cPageNo=1
				rs.AbsolutePage=cPageNo
				for i=1 to rs.pagesize
					if rs.eof Then exit For
			%>
          <tr class="hback"> 
            <td width="21%" rowspan="5" class="hback_1">
			<table border="0" align="center" cellpadding="2" cellspacing="1" class="table">
                <tr> 
                  <td class="hback"><%if isnull(trim(rs("PicSavePath"))) then%><img src="Images/nopic_supply.gif" width="90"  id="pic_p_1"><%else%><a href="<%=rs("PicSavePath")%>" rel="lightbox" title="<%=rs("title")%>"><img src="<%=rs("PicSavePath")%>" width="90" border="0" id="pic_p_1"></a><%end if%>
                  </td>
                </tr>
              </table></td>
            <td width="12%" class="hback"><div align="center"><strong>相片名称：</strong></div></td>
            <td width="40%" class="hback"><%if rs("isRec")=1 then response.Write"<span class=""tx"">[此图片已经被推荐]</span>"%><font style="font-size:14px"><span class="hback_1"><strong><%=rs("title")%></strong>;</span></font></td>
            <td width="10%"><div align="center">浏览次数：</div></td>
            <td width="17%"><%=rs("hits")%></td>
          </tr>
          <tr class="hback"> 
            <td><div align="center">创建日期：</div></td>
            <td><%=rs("Addtime")%></td>
            <td><div align="center">图片大小：</div></td>
            <td><%=rs("PicSize")%>byte</td>
          </tr>
          <tr class="hback"> 
            <td><div align="center">相片地址：</div></td>
            <td colspan="3"><%=rs("PicSavePath")%></td>
          </tr>
          <tr class="hback"> 
            <td><div align="center">相片描述：</div></td>
            <td colspan="3"><%=rs("Content")%></td>
          </tr>
          <tr class="hback"> 
            <td><div align="center">相片分类：</div></td>
            <td><%
			dim c_rs
			if rs("ClassID")=0 then
				response.Write("没分类")
			else
				set c_rs=User_Conn.execute("select ID,title From Fs_me_photoclass where id="&rs("classid"))
				Response.Write "<a href=PhotoManage.asp?ClassiD="&c_rs("ID")&">"&c_rs("title")&"</a>"
				c_rs.close:set c_rs=nothing
			end if
			%></td>
            <td colspan="2"><div align="center"><a href="Photo_Edit.asp?Id=<%=rs("id")%>">修改</a>┆<a href="PhotoManage.asp?id=<%=rs("id")%>&Action=del" onClick="{if(confirm('确定通过删除吗？')){return true;}return false;}">删除</a> 
                <input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>">
              </div></td>
          </tr>
          <tr class="hback"> 
            <td colspan="5" height="3" class="xingmu"></td>
          </tr>
          <%
			rs.movenext
		next
		%>
          <tr class="hback"> 
            <td colspan="5"> 
              <%
			response.Write "<p>"&  fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
		  end if
		  %>
              　
<input name="chkall" type="checkbox" id="chkall" onClick="CheckAll(this.form);" value="checkbox" title="点击选择所有或者撤消所有选择">
              全选择
              <input name="Action" type="hidden" id="Action"> <input type="button" name="Submit1" value="删除"  onClick="document.myform.Action.value='delall';{if(confirm('确定清除您所选择的记录吗？')){this.document.myform.submit();return true;}return false;}"></td>
          </tr></form>
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
Set Fs_User = Nothing
%>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = myform.elements[i];  
    if (e.name != 'chkall')  
       e.checked = myform.chkall.checked;  
    }  
}
</script>

<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






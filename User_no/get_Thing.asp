<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp" -->
<%
Dim strpage,int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
dim obj_nothing_Rs,i,str_action,str_id,str_Url_down,rs_d_obj
str_action = NoSqlHack(Request("Action"))
str_id = CintStr(Request.QueryString("id"))
if str_action = "Down" then
	set rs_d_obj = User_Conn.execute("select * From FS_ME_getThing where UserNumber='"&NoSqlHack(Fs_User.UserNumber)&"' and UserDel=0 and id="&CintStr(str_id))
	if rs_d_obj.eof then
		strShowErr = "<li>找不到记录</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
		rs_d_obj.close:set rs_d_obj = nothing
	else
		if rs_d_obj("useNum")>=rs_d_obj("MaxNum") then
			strShowErr = "<li>您已经下载了"&rs_d_obj("MaxNum")&"次,不能再下载!</li>"
			Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
			rs_d_obj.close:set rs_d_obj = nothing
		end if
		if rs_d_obj("isLock")=1 then
			strShowErr = "<li>此记录已经被锁定</li>"
			Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
			rs_d_obj.close:set rs_d_obj = nothing
		else
			'更新数据库
			DIM up_d_rs
			set up_d_rs = Server.CreateObject(G_FS_RS)
			up_d_rs.open "select * From FS_ME_getThing where UserNumber='"&Fs_User.UserNumber&"' and id="&CintStr(str_id),User_conn,1,3
			up_d_rs("isUse")=1
			up_d_rs("useNum")=up_d_rs("useNum")+1
			up_d_rs("UpdateTime")=now
			up_d_rs("IP")=NoSqlHack(Request.ServerVariables("Remote_Addr"))
			up_d_rs.update
			up_d_rs.close:set up_d_rs=nothing
			Response.Redirect rs_d_obj("URL_1")
			rs_d_obj.close:set rs_d_obj = nothing
			response.end
		end if
	end if
elseif str_action = "Del" then
	User_Conn.execute("Update FS_ME_getThing set UserDel=1 where UserNumber='"&Fs_User.UserNumber&"' and Id="&CintStr(str_Id))
	strShowErr = "<li>删除成功</li>"
	Response.Redirect("lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Get_Thing.asp")
	Response.end
end if
strpage=request("page")
if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
int_RPP=20 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"
Set obj_nothing_Rs = server.CreateObject(G_FS_RS)
SQL = "Select *  from FS_ME_getThing where UserNumber='"&Fs_User.UserNumber&"' and islock=0 and UserDel=0 Order by id desc"
obj_nothing_Rs.Open SQL,User_Conn,1,3
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-获取商品</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt; 获取商品 </td>
        </tr>
      </table> 
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback">
          <td colspan="7" class="xingmu">获取商品</td>
        </tr>
        <tr class="hback"> 
          <td width="25%" class="hback"><div align="left"><strong>名称</strong></div></td>
          <td width="13%" class="hback"><div align="left"><strong>版本</strong></div></td>
          <td width="13%" class="hback"><div align="center"><strong>型号</strong></div></td>
          <td width="11%" class="hback"><div align="center"><strong>已下载</strong></div></td>
          <td width="10%" class="hback"><div align="center"><strong>最大下载</strong></div></td>
          <td width="18%" class="hback"><div align="center"><strong>最后下载时间</strong></div></td>
          <td width="10%" class="hback"><div align="center"><strong>操作</strong></div></td>
        </tr>
		<%
		if obj_nothing_Rs.eof then
		   obj_nothing_Rs.close
		   set obj_nothing_Rs=nothing
		   Response.Write"<tr  class=""hback""><td colspan=""6""  class=""hback"" height=""40"">没有记录。</td></tr>"
		else
			obj_nothing_Rs.PageSize=int_RPP
			cPageNo=NoSqlHack(Request.QueryString("Page"))
			If cPageNo="" Then cPageNo = 1
			If not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo<=0 Then cPageNo=1
			If cPageNo>obj_nothing_Rs.PageCount Then cPageNo=obj_nothing_Rs.PageCount 
			obj_nothing_Rs.AbsolutePage=cPageNo
			for i=1 to obj_nothing_Rs.pagesize
				if obj_nothing_Rs.eof Then exit For 
		%>
        <tr class="hback"> 
          <td class="hback"><a href="#" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(down_<% = obj_nothing_Rs("id")%>);" onMouseOver="this.className='bg'" onMouseOut="this.className='bg1'" language=javascript><% = obj_nothing_Rs("ProductID")%></a></td>
          <td class="hback"><% = obj_nothing_Rs("Version")%></td>
          <td class="hback"><div align="center">
            <% = obj_nothing_Rs("PType")%>
          </div></td>
          <td class="hback"><div align="center">
            <% = obj_nothing_Rs("useNum")%>次</div></td>
          <td class="hback"><div align="center">
            <% = obj_nothing_Rs("MaxNum")%>次</div></td>
          <td class="hback"><div align="center">
            <% = obj_nothing_Rs("UpdateTime")%>
          </div></td>
          <td class="hback"><div align="center"><a href="get_Thing.asp?Action=Down&Id=<%=obj_nothing_Rs("id")%>">下载</a>|<a href="get_Thing.asp?Action=Del&Id=<%=obj_nothing_Rs("id")%>" onClick="{if(confirm('确定要删除吗?\n删除后将不能恢复!!')){return true;}return false;}">删除</a></div></td>
        </tr>
         <tr class="hback" id="down_<% = obj_nothing_Rs("id")%>" style="display:none;">
           <td colspan="7" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="5" class="hback_1">
             <tr>
               <td width="31%">最后下载IP：
                 <% = obj_nothing_Rs("IP")%></td>
               <td width="69%">添加日期：
                 <% = obj_nothing_Rs("AddTime")%></td>
             </tr>
             <tr>
               <td colspan="2">描述：
                <% = obj_nothing_Rs("Content")%></td>
              </tr>
           </table>             </td>
         </tr>
		 <%
				obj_nothing_Rs.movenext
			Next
		 %>
		<tr class="hback"> 
          <td colspan="7" class="hback">
			<%
					response.Write "<p>"&  fPageCount(obj_nothing_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
				   obj_nothing_Rs.close
				   set obj_nothing_Rs=nothing
			end if
			%>	
		 </td>
        </tr>
      </table></td>
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






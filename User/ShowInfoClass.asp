<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp" -->
<%
		Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,i_ii
		Dim SQL_list,rs_l,str_title,str_id,str_addtime,str_lockTF,type_s,novalue
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
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-专栏管理</title>
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
          <td colspan="2" class="xingmu">专栏管理</td>
        </tr>
        <tr class="hback"> 
          <%
		  dim rs,classid,usernumber,s_type
		  classid = CintStr(Request.QueryString("ClassID"))
		  usernumber = NoSqlHack(Request.QueryString("UserNumber"))
		  if classid="" or not isnumeric(classid) or usernumber="" then
				strShowErr = "<li>错误的参数</li>"
				Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
		  end if
		  set rs = Server.CreateObject(G_FS_RS)
		  rs.open "select ClassID,ClassCName,ClassEName,ParentID,UserNumber,ClassTypes,AddTime,ClassContent From FS_ME_InfoClass where Classid="& classid &" and UserNumber='"& usernumber &"'",User_Conn,1,3
		  if rs.eof then
				strShowErr = "<li>找不到记录</li>"
				Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
		  end if
		  %>
          <td colspan="2" class="hback"> </td>
        </tr>
        <tr class="hback"> 
          <td width="20%" class="hback_1"><div align="right">专栏名称</div></td>
          <td width="80%" class="hback"><%=rs("ClassCName")%></td>
        </tr>
        <tr class="hback"> 
          <td class="hback_1"><div align="right">专栏创建日期</div></td>
          <td class="hback"><%=rs("addtime")%></td>
        </tr>
        <tr class="hback"> 
          <td class="hback_1"><div align="right">专栏所属类型</div></td>
          <td class="hback">
		  <%
				select case rs("ClassTypes")
					case 0
						s_type = "新闻文学类"
					case 1
						s_type = "下载"
					case 2
						s_type = "商品"
					case 3
						s_type = "房产"
					case 4
						s_type = "供求"
					case 5
						s_type = "求职"
					case 6
						s_type = "招聘"
					case 7
						s_type = "生活日志"
				end select
				Response.Write s_type
		 %>
		 </td>
        </tr>
        <tr class="hback">
          <td height="38" class="hback_1"> 
            <div align="right">专栏描述</div></td>
          <td class="hback"><%=rs("ClassContent")%></td>
        </tr>
      </table>
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr> 
          <td colspan="4" class="hback_1">专栏下列表<span class="tx">(如果要修改，请在相应子系统下去修改)</span></td>
        </tr>
        <tr class="hback_1"> 
          <td width="38%"><div align="center"><strong>名称</strong></div></td>
          <td width="35%"><div align="center"><strong>创建日期</strong></div></td>
          <td width="16%"><div align="center"><strong>类型</strong></div></td>
          <td width="11%"><div align="center"><strong>状态</strong></div></td>
        </tr>
        <%
			strpage=request("page")
			if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
			set rs_l= Server.CreateObject(G_FS_RS)
			select case rs("ClassTypes")
				case 0
					SQL_list = "select ContID,ContTitle,AddTime,AuditTF From FS_ME_InfoContribution where UserNumber='"& usernumber &"' and ClassID="&rs("ClassID")&""
					rs_l.open SQL_list,User_Conn,1,3
					if rs_l.eof then
						call no_s()
					else
						call no_s_1()
					end if
				case 1
					 
				case 2
					 
				case 3
					 
				case 4
					SQL_list = "select ID,PubTitle,AddTime,isPass,MyClassID From FS_SD_News where UserNumber='"& usernumber &"' and MyClassID="&rs("ClassID")&""
					rs_l.open SQL_list,Conn,1,3
					if rs_l.eof then
						call no_s()
					else
						call no_s_1()
					end if
				case 5
					 
				case 6
				
				case 7
					 
			end select
			sub no_s()
				   rs_l.close
				   set rs_l=nothing
				   Response.Write"<tr  class=""hback""><td colspan=""4""  class=""hback"" height=""40"">没有记录。</td></tr>"
			end sub
			sub no_s_1()
						rs_l.PageSize=int_RPP
						cPageNo=NoSqlHack(Request.QueryString("Page"))
						If cPageNo="" Then cPageNo = 1
						If not isnumeric(cPageNo) Then cPageNo = 1
						cPageNo = Clng(cPageNo)
						If cPageNo<=0 Then cPageNo=1
						If cPageNo>rs_l.PageCount Then cPageNo=rs_l.PageCount 
						rs_l.AbsolutePage=cPageNo
						for i_ii=1 to rs_l.pagesize
							if rs_l.eof Then exit For 
						  select case rs("ClassTypes")
								case 0
									str_title=rs_l("ContTitle")
									str_id = rs_l("ContID")
									str_addtime =rs_l("AddTime")
									str_lockTF=rs_l("AuditTF")
									type_s = "投稿"
								case 1
									 
								case 2
									 
								case 3
									 
								case 4
									str_title=rs_l("PubTitle")
									str_id = rs_l("ID")
									str_addtime =rs_l("AddTime")
									str_lockTF=rs_l("isPass")
									type_s = "供求"
								case 5
									 
								case 6
								
								case 7
									 
							end select
	%>
        <tr class="hback"> 
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="7%"><img src="Images/folderopened.gif" width="16" height="16"></td>
                <td width="93%" valign="bottom"> <% = str_title %> </td>
              </tr>
            </table></td>
          <td><div align="center"> 
              <% = str_addtime %>
            </div></td>
          <td><div align="center"> 
              <% = type_s %>
            </div></td>
          <td><div align="center"> 
              <% 
			  select case rs("ClassTypes")
					case 0
						if str_lockTF=1 then:response.Write("已审核"):else:response.Write("<span class=""tx"">未审核</span>"):end if
					case 1
						 
					case 2
						 
					case 3
						 
					case 4
						if str_lockTF=1 then:response.Write("已审核"):else:response.Write("<span class=""tx"">未审核</span>"):end if
					case 5
						 
					case 6
					
					case 7
						 
				end select
			    %>
            </div></td>
        </tr>
        <%
			rs_l.movenext
		next
		%>
        <tr class="hback"> 
          <td colspan="4"><div align="right"> 
              <%response.Write "<p>"&  fPageCount(rs_l,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo) %>
            </div></td>
        </tr>
        <%
		rs_l.close:set rs_l=nothing
		end sub
		%>
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
rs.close:set rs=nothing
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






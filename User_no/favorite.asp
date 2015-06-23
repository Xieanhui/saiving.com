<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp" -->
<%
	User_Conn.execute("Select UserSystemName From FS_ME_SysPara")
	if request("Action")="del" then
		if Request("id")="" then
			strShowErr = "<li>错误的参数！</li>"
			Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		else
			User_Conn.execute("Delete from FS_ME_Favorite where FavoID in ("&FormatIntArr(Request("id"))&")")
			strShowErr = "<li>删除成功！</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Favorite.asp")
			Response.end
		end If
	Elseif Request("Action")="sort" Then
		if Request("id")="" Or Request("classID")="" then
			strShowErr = "<li>错误的参数！</li>"
			Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		Else
			User_Conn.execute("Update FS_ME_Favorite set FavoClassID="&CintStr(Request("ClassID"))&" where FavoID in ("&FormatIntArr(Request("id"))&")")
			strShowErr = "<li>转移成功</li>"	
		    Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Favorite.asp")
			Response.end
		end If
	end if
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,rs,i
	int_RPP=20 '设置每页显示数目
	int_showNumberLink_=5 '数字导航显示数目
	showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
	str_nonLinkColor_="#999999" '非热链接颜色
	toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
	toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
	toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
	toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
	toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
	toL_="<font face=webdings title=""最后一页"">:</font>"
	strpage=CintStr(request("page"))
	if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="top.asp" -->
    </td>
  </tr>
</table>
<form name="myForm" method="post" action="">
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="Top_navi.asp" -->
    </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback">
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td  valign="top">你的位置：<a href="../">网站首页</a> &gt;&gt; <a href="main.asp">会员首页</a> 
            &gt;&gt; <a href="Favorite.asp">收藏夹管理</a> &gt;&gt; </td>
        </tr>
        <tr class="hback"> 
          <td  valign="top"><a href="Favorite.asp">全部</a>┆<a href="Favorite.asp?Type=0">新闻</a>┆<%if IsExist_SubSys("DS") Then%><a href="Favorite.asp?Type=1">下载</a>┆<%end if%><a href="Favorite.asp?Type=2">企业会员</a>┆<%if IsExist_SubSys("SD") Then%><a href="Favorite.asp?Type=3">供求信息</a>┆<%end if%><%if IsExist_SubSys("MS") Then%><a href="Favorite.asp?Type=4">商品</a>┆<%end if%><%if IsExist_SubSys("HS") Then%><a href="Favorite.asp?Type=5">房产信息</a>┆<%end if%><%if IsExist_SubSys("AP") Then%><a href="Favorite.asp?Type=6">招聘</a>┆<%end if%><a href="Favorite.asp?Type=7">日志</a>┆<a href="Favorite_Class.asp">收藏夹(分类)管理</a></td>
        </tr>
        <tr class="hback">
          <td>
          <%
		  response.Write("	<table width=""98%"" align=center cellpadding=""2"" cellspacing=""1""><tr>")
		  dim t_k
		  t_k=0
		  set rs = Server.CreateObject(G_FS_RS)
		  rs.open "select ClassID,ClassCName,UserNumber From FS_ME_FavoriteClass where UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,1
		  do while not rs.eof 
		  	Response.Write("<td width=""24%"" valign=bottom><input type=""radio"" name=""classID"" value="""&rs("ClassID")&"""/><a href=""Favorite.asp?classid="&rs("ClassID")&""">"&rs("ClassCName")&"</a></td>")
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
       <%
	   
		  	dim o_type,o_class
		  	select case NoSqlHack(Request.QueryString("type"))
				case "1"
					o_type = " and FavoriteType=1"
				case "2"
					o_type = " and FavoriteType=2"
				case "3"
					o_type = " and FavoriteType=3"
				case "4"
					o_type = " and FavoriteType=4"
				case "5"
					o_type = " and FavoriteType=5"
				case "6"
					o_type = " and FavoriteType=6"
				case "7"
					o_type = " and FavoriteType=7"
				case "0"
					o_type = " and FavoriteType=0"
				case else
					o_type = ""
			end select
			if Request.QueryString("ClassId")<>"" then
				o_class=" and FavoClassID="&CintStr(Request.QueryString("ClassId"))&""
			else
				o_class=""
			end if
			set rs = Server.CreateObject(G_FS_RS)
			rs.open "select * from FS_ME_Favorite where UserNumber='"& Fs_User.UserNumber&"' "& o_type & o_class &" order by FavoID desc",User_Conn,1,1
			if rs.eof then
			   Response.Write"<tr  class=""hback""><td colspan=""6""  class=""hback"" height=""40"">没有记录。</td></tr>"
			else
			%>
          <tr> 
            <td width="4%" class="hback_1"><div align="center"> 
                <input name="chkall" type="checkbox" id="chkall" onClick="CheckAll(this.form);" value="checkbox" title="点击选择所有或者撤消所有选择">
              </div></td>
            <td class="hback_1"><div align="left"><strong>信息标题</strong></div></td>
            <td width="7%" class="hback_1"><div align="center"><strong>类型</strong></div></td>
            <td width="17%" class="hback_1"><div align="center"><strong>日期</strong></div></td>
            <td width="15%" class="hback_1"><div align="center"><strong>分类</strong></div></td>
            <td width="12%" class="hback_1"><div align="center"><strong>操作</strong></div></td>
          </tr>
          <%
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
          <tr  onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="center"> 
                <input name="id" type="checkbox" id="id" value="<%=rs("FavoID")%>">
              </div></td>
            <td class="hback"> 
              <%
			dim f_rs
			select case rs("FavoriteType")
				case 0
					set f_rs=Conn.execute("select id,NewsTitle From FS_NS_News where id="&CintStr(rs("FID")))
					if f_rs.eof then:Response.Write("<span class=tx>信息已经被管理员删除</span>"):else:response.Write"<a href=Public_info.asp?type=NS&Url="&rs("FID")&" target=_blank>"&f_rs("NewsTitle")&"</a>":end if
					f_rs.close:set f_rs=nothing
				case 1
					set f_rs=Conn.execute("select id,Name From FS_DS_List where id="&CintStr(rs("FID")))
					if f_rs.eof then:Response.Write("<span class=tx>信息已经被管理员删除</span>"):else:response.Write"<a href=Public_info.asp?type=DS&Url="&rs("FID")&" target=_blank>"&f_rs("Name")&"</a>":end if
					f_rs.close:set f_rs=nothing
				case 2
					response.Write"<a href=Public_info.asp?type=AP_1&Url="&rs("FID")&" target=_blank>查看</a>"
				case 3
					response.Write"<a href=../Supply/Supply.asp&id="&rs("FID")&" target=_blank>查看</a>"
				case 4
					response.Write"<a href=Public_info.asp?type=MS&Url="&rs("FID")&" target=_blank>查看</a>"
				case 5
					response.Write"<a href=../House/house.asp?ID="&rs("FID")&" target=_blank>查看</a>"
				case 6
					response.Write"<a href=Public_info.asp?type=AP_2&Url="&rs("FID")&" target=_blank>查看</a>"
				case 7
					response.Write"<a href=../Blog/Blog.asp?id="&rs("FID")&" target=_blank>查看</a>"
				case else
			end select
			%>
            </td>
            <td class="hback"><div align="center"> 
                <%
			select case rs("FavoriteType")
				case 0
					response.Write"<a href=Favorite.asp?type=0>新闻</a>"
				case 1
					response.Write"<a href=Favorite.asp?type=1>下载</a>"
				case 2
					response.Write"<a href=Favorite.asp?type=2>企业</a>"
				case 3
					response.Write"<a href=Favorite.asp?type=3>供求</a>"
				case 4
					response.Write"<a href=Favorite.asp?type=4>商品</a>"
				case 5
					response.Write"<a href=Favorite.asp?type=5>房产</a>"
				case 6
					response.Write"<a href=Favorite.asp?type=6>招聘</a>"
				case 7
					response.Write"<a href=Favorite.asp?type=6>日志</a>"
				case else
					response.Write"<a href=Favorite.asp>-</a>"
			end select
			
			%>
              </div></td>
            <td class="hback"><div align="center"><%=rs("AddTime")%></div></td>
            <td class="hback"> <div align="center">
			<%
			if rs("FavoClassID")=0 then
				response.Write"<a href=Favorite.asp?ClassID=0>未分类</a>"
			else
				dim crs
				set crs=user_Conn.execute("select ClassID,ClassCName,UserNumber From FS_ME_FavoriteClass where ClassID="&rs("FavoClassID"))
				Response.Write "<a href=""Favorite.asp?ClassID="&crs("ClassID")&""">"&crs("ClassCName")&"</a>"
			end if
			%> 
                </div></td>
            <td class="hback"><div align="center">
			<a href="Favorite.asp?Action=del&id=<%=rs("FavoID")%>" onClick="{if(confirm('确定要删除吗?')){return true;}return false;}">删除</a>
			| <a href="#" onclick="sort('<%=rs("FavoID")%>')">转移</>
			</div></td>
          </tr>
          <%
			  rs.movenext
		  next
		  %>
          <tr  class="hback"> 
            <td height="31" colspan="6">
			<div align="right"> 
            <%
			response.Write "<p>"&  fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
			%>
                <input name="Action" type="hidden" id="Action">
                <input type="button" name="Submit" value="删除"  onClick="document.myForm.Action.value='del';{if(confirm('确定清除您所选择的记录吗？')){this.document.myForm.submit();return true;}return false;}">
				 <input type="button" name="Submit" value="转移"  onClick="document.myForm.Action.value='sort';this.document.myForm.submit();">
              </div></td>
          </tr>
          <% 
		  end if
		  %>
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
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = myForm.elements[i];  
    if (e.name != 'chkall')  
       e.checked = myForm.chkall.checked;  
    }  
	}
function sort(id)
{
	var elements=document.all("classID")
	var classid="";
	if(elements.length=='undefined')
	{
		classid=elements.value;
	}
	else
	{
		for(var i=0;i<elements.length;i++)
		{
			if (elements[i].checked)
			{
				classid=elements[i].value;
			}
		}
	}
	if (classid=="")
	{
		alert("请选择目标分类！");
		return false;
	}
	location.href="Favorite.asp?Action=sort&id="+id+"&classiD="+classid;
}
</script>
<%
Set Fs_User = Nothing
rs.close
set rs=nothing
set conn=nothing
set User_Conn=nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






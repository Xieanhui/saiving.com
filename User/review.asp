<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp" -->
<%
Dim DleIDStr
if Trim(Request.QueryString("Url"))<>"" then
	response.Write NoSqlHack(request.QueryString("Url"))&"-" &NoSqlHack(request.QueryString("type"))
	response.end
end if
if request("Action")="del" then
	if Request("id")="" then
		strShowErr = "<li>错误的参数！</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		DleIDStr = NoHtmlHackInput(NoSqlHack(Trim(Request("id"))))
		User_Conn.execute("Delete from FS_ME_Review where UserNumber = '" & Fs_User.UserNumber & "' And ReviewID in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>删除成功！</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Review.asp")
		Response.end
	end if
end if
if request("Action")="Lock" then
	if Request("id")="" then
		strShowErr = "<li>错误的参数！</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		DleIDStr = NoHtmlHackInput(NoSqlHack(Trim(Request("id"))))
		User_Conn.execute("Update FS_ME_Review set islock=1 where UserNumber = '" & Fs_User.UserNumber & "' And ReviewID in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>锁定成功！</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Review.asp")
		Response.end
	end if
end if
if request("Action")="UnLock" then
	if Request("id")="" then
		strShowErr = "<li>错误的参数！</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		DleIDStr = NoHtmlHackInput(NoSqlHack(Trim(Request("id"))))
		User_Conn.execute("Update FS_ME_Review set islock=0 where UserNumber = '" & Fs_User.UserNumber & "' And ReviewID in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>解琐成功！</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Review.asp")
		Response.end
	end if
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
strpage=request("page")
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
            &gt;&gt; <a href="Review.asp">评论管理</a> &gt;&gt; </td>
        </tr>
        <tr class="hback"> 
          <td  valign="top"><a href="Review.asp">全部评论</a>┆<a href="Review.asp?type=0">新闻评论</a>┆<%if IsExist_SubSys("DS") Then%><a href="Review.asp?type=1">下载评论</a>┆<%End if%><%if IsExist_SubSys("MS") Then%><a href="Review.asp?type=2">商品评论</a>┆<%end if%><%if IsExist_SubSys("HS") Then%><a href="Review.asp?type=3">房产评论</a>┆<%end if%><%if IsExist_SubSys("SD") Then%><a href="Review.asp?type=4">供求评论</a>┆<%end if%><a href="Review.asp?type=5">日记评论</a>┆<a href="Review.asp?type=6">相册评论</a></td>
        </tr>
      </table>
     
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="myForm" method="post" action="">
		  <tr> 
            <td width="6%" class="hback_1"><div align="center"> 
                <input name="chkall" type="checkbox" id="chkall" onClick="CheckAll(this.form);" value="checkbox" title="点击选择所有或者撤消所有选择">
              </div></td>
            <td width="27%" class="hback_1"><div align="left"><strong>标题</strong></div></td>
            <td width="14%" class="hback_1"><div align="center"><strong>类型</strong></div></td>
            <td width="10%" class="hback_1"><div align="center"><strong>被评论信息</strong></div></td>
            <td width="17%" class="hback_1"><div align="center"><strong>日期</strong></div></td>
            <td width="6%" class="hback_1"><div align="center"><strong>状态</strong></div></td>
            <td width="5%" class="hback_1"><strong>通过</strong></td>
            <td width="15%" class="hback_1"><strong>操作</strong></td>
          </tr>
          <%
		  	dim o_type
		  	select case NoSqlHack(Request.QueryString("type"))
				case "1"
					o_type = " and ReviewTypes=1"
				case "2"
					o_type = " and ReviewTypes=2"
				case "3"
					o_type = " and ReviewTypes=3"
				case "4"
					o_type = " and ReviewTypes=4"
				case "5"
					o_type = " and ReviewTypes=5"
				case "6"
					o_type = " and ReviewTypes=6"
				case "0"
					o_type = " and ReviewTypes=0"
				case else
					o_type = ""
			end select
			set rs = Server.CreateObject(G_FS_RS)
			rs.open "select * from FS_ME_Review where UserNumber='"& Fs_User.UserNumber&"' "& o_type &" order by Addtime desc,ReviewID desc",User_Conn,1,3
			if rs.eof then
			   rs.close
			   set rs=nothing
			   set conn=nothing
			   set fs_user=nothing
			   Response.Write"<tr  class=""hback""><td colspan=""8""  class=""hback"" height=""40"">没有记录。</td></tr>"
			else
				rs.PageSize=int_RPP
				cPageNo=NoSqlHack(Request.QueryString("Page"))
				If cPageNo="" Then cPageNo = 1
				If not isnumeric(cPageNo) Then cPageNo = 1
				cPageNo = Clng(cPageNo)
				If cPageNo<=0 Then cPageNo=1
				If cPageNo>rs.PageCount Then cPageNo=rs.PageCount 
				rs.AbsolutePage=cPageNo
				for i=1 to rs.pagesize
					if rs.eof Then exit For 
		  %>
          <tr  onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="center"> 
                <input name="id" type="checkbox" id="id" value="<%=rs("ReviewID")%>">
              </div></td>
            <td class="hback" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(rid<%=rs("ReviewID")%>);" language=javascript><%=rs("title")%></td>
            <td class="hback"><div align="center">
			<%'0为新闻评论，1为下载评论，2为商品，3为房产评论，4为供求评论，5日记,6相册
			select case rs("ReviewTypes")
				case 0
					response.Write"<a href=Review.asp?type=0>新闻评论</a>"
				case 1
					response.Write"<a href=Review.asp?type=1>下载评论</a>"
				case 2
					response.Write"<a href=Review.asp?type=2>商品评论</a>"
				case 3
					response.Write"<a href=Review.asp?type=3>房产评论</a>"
				case 4
					response.Write"<a href=Review.asp?type=4>供求评论</a>"
				case 5
					response.Write"<a href=Review.asp?type=5>日志评论</a>"
				case 6
					response.Write"<a href=Review.asp?type=6>相册评论</a>"
				case else
					response.Write"<a href=Review.asp>-</a>"
			end select
			%></div></td>
            <td class="hback"><div align="center">
			<%'0为新闻评论，1为下载评论，2为商品，3为房产评论，4为供求评论，5日记,6相册
			select case rs("ReviewTypes")
				case 0
					response.Write"<a href=Public_info.asp?type=NS&Url="&rs("InfoID")&" target=_blank>查看</a>"
				case 1
					response.Write"<a href=Public_info.asp?type=DS&Url="&rs("InfoID")&" target=_blank>查看</a>"
				case 2
					response.Write"<a href=Public_info.asp?type=MS&Url="&rs("InfoID")&" target=_blank>查看</a>"
				case 3
					response.Write"<a href=Public_info.asp?type=HS&Url="&rs("InfoID")&" target=_blank>查看</a>"
				case 4
					response.Write"<a href=Public_info.asp?type=SD&Url="&rs("InfoID")&" target=_blank>查看</a>"
				case 5
					response.Write"<a href=Public_info.asp?type=LS&Url="&rs("InfoID")&" target=_blank>查看</a>"
				case 6
					response.Write"<a href=Public_info.asp?type=PH&Url="&rs("InfoID")&" target=_blank>查看</a>"
				case else
					
			end select
			%>
		</div></td>
            <td class="hback"><div align="center"><%=rs("AddTime")%></div></td>
            <td class="hback"><div align="center"> 
                <%if rs("isLock")=1 then:response.Write"<span class=tx>锁定</span>":else:response.Write"开放":end if%>
              </div></td>
            <td class="hback"> 
              <div align="center"><b><%if rs("AdminLock")=1 then:response.Write"<span class=tx>×</span>":else:response.Write"√":end if%></b></div>
            </td>
            <td class="hback"><div align="center"><a href="Review.asp?Action=del&id=<%=rs("ReviewID")%>" onClick="{if(confirm('确定要删除吗?')){return true;}return false;}">删除</a>┆<a href="Review.asp?Action=UnLock&id=<%=rs("ReviewID")%>">解锁</a>┆<a href="Review.asp?Action=Lock&id=<%=rs("ReviewID")%>" onClick="{if(confirm('确定锁定评论吗？？\n锁定后将不能显示')){return true;}return false;}">锁定</a></div></td>
          </tr>
          <tr  class="hback" id="rid<%=rs("ReviewID")%>" style="display:none"> 
            <td height="40"><div align="center">内容:</div></td>
            <td height="40" colspan="7"><%=rs("Content")%></td>
          </tr>
          <%
			  rs.movenext
		  next
		  %>
          <tr  class="hback"> 
            <td colspan="8"><div align="right"> 
                <%
			response.Write "<p>"&  fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
			%>
                <input name="Action" type="hidden" id="Action">
                <input type="button" name="Submit" value="删除"  onClick="document.myForm.Action.value='del';{if(confirm('确定清除您所选择的记录吗？')){this.document.myForm.submit();return true;}return false;}">
                <input type="button" name="Submit2" value="批量解锁"  onClick="document.myForm.Action.value='UnLock';{if(confirm('确定解锁评论吗')){this.document.myForm.submit();return true;}return false;}">
                <input name="Submit3" type="button"  onClick="document.myForm.Action.value='Lock';{if(confirm('确定锁定评论吗？？\n锁定后将不能显示')){this.document.myForm.submit();return true;}return false;}" value="批量锁定">
              </div></td>
          </tr>
          <% 
		  rs.close:set rs=nothing
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
</script>

<%
Set Fs_User = Nothing
set user_conn=nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






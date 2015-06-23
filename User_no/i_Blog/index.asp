<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
dim rs_mysys
If CheckBlogOpen=False Then 
	Response.write("<script language=""javascript"">alert('日志功能暂停使用,如需要使用请联系管理员.');history.back();</script>")
	Response.End()
End If
set rs_mysys = User_Conn.execute("select id From FS_ME_InfoiLogParam where UserNumber='"& Fs_User.UserNumber&"'")
if rs_mysys.eof then
	Response.write("<br>要发布日志，请开通您的日志空间,5秒后转向...")
	Response.Write("<meta http-equiv=""refresh"" content=""5;url=PublicParam.asp"">")
	response.end
end if
if request("Action")="del" then
	if Request("id")="" then
		strShowErr = "<li>错误的参数！</li>"
		Response.Redirect("../lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		User_Conn.execute("Delete from FS_ME_Infoilog where iLogID in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>删除成功！</li>"
		Response.Redirect("../lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../i_Blog/index.asp")
		Response.end
	end if
end if
if request("Action")="Lock" then
	if Request("id")="" then
		strShowErr = "<li>错误的参数！</li>"
		Response.Redirect("../lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		User_Conn.execute("Update FS_ME_Infoilog set islock=1 where iLogID in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>锁定成功！</li>"
		Response.Redirect("../lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../i_Blog/index.asp")
		Response.end
	end if
end if
if request("Action")="UnLock" then
	if Request("id")="" then
		strShowErr = "<li>错误的参数！</li>"
		Response.Redirect("../lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		User_Conn.execute("Update FS_ME_Infoilog set islock=0 where iLogID in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>解琐成功！</li>"
		Response.Redirect("../lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../i_Blog/index.asp")
		Response.end
	end if
end if

Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,rs,i
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
strpage=request("page")
if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
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
        <form name="myform" method="post" action="">
          <tr> 
            <td width="4%" class="xingmu"><div align="center"> 
                <input name="chkall" type="checkbox" id="chkall" onClick="CheckAll(this.form);" value="checkbox" title="点击选择所有或者撤消所有选择">
              </div></td>
            <td width="31%" class="xingmu"><div align="center">标题</div></td>
            <td width="15%" class="xingmu"><div align="center">分类</div></td>
            <td width="17%" class="xingmu"><div align="center">日期</div></td>
            <td width="14%" class="xingmu"><div align="center">状态</div></td>
            <td width="19%" class="xingmu"><div align="center">操作</div></td>
          </tr>
          <%
		  	dim o_class,o_draff
		  	if request.QueryString("classid")<>"" then
				o_class= " and ClassId="&CintStr(request.QueryString("classid"))&""
			else
				o_class= ""
			end if
		  	if request.QueryString("type")="box" then
				o_draff= " and isDraft=1"
			else
				o_draff= ""
			end if
			set rs = Server.CreateObject(G_FS_RS)
			rs.open "select * from FS_ME_Infoilog where UserNumber='"& Fs_User.UserNumber&"' "&o_draff&o_class&" order by isTop desc,AddTime desc,iLogID desc",User_Conn,1,3
			if rs.eof then
			   rs.close
			   set rs=nothing
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
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="center"> 
                <input name="id" type="checkbox" id="id" value="<%=rs("iLogID")%>">
              </div></td>
            <td class="hback"><a href="PublicLogEdit.asp?id=<%=rs("iLogID")%>"><%=rs("Title")%></a></td>
            <td class="hback"><div align="center">
			<%
			if rs("ClassID")=0 then
				response.Write"<a href=index.asp?classid=0>未分类</a>"
			else
				dim c_rs
				set c_rs= Server.CreateObject(G_FS_RS)
				c_rs.open "select ClassID,ClassCName From FS_ME_InfoClass where UserNumber='"&Fs_User.UserNumber&"' and ClassTypes=7 and ClassID="&rs("ClassID"),User_Conn,1,3
				if not c_rs.eof then
					Response.Write "<a href=index.asp?classid="&rs("ClassID")&">"&c_rs("ClassCName")&"</a>"
					c_rs.close:set c_rs=nothing
				else
					response.Write"<a href=index.asp?classid=0>未分类</a>"
					c_rs.close:set c_rs=nothing
				end if
			end if
			%> </div></td>
            <td class="hback"><div align="center"><%=rs("addtime")%> </div></td>
            <td class="hback"> 
              <div align="center">
			  <%
			if rs("adminLock")=1 then
				Response.Write("<span class=tx>管理员审核中.或用户锁定</span>")
			else
				if rs("islock")=1 then
					response.Write("用户锁定")
				else
					response.Write("正常")
				end if
			end if
			%>
              </div></td>
            <td class="hback"><div align="center"><a href="PublicLogEdit.asp?id=<%=rs("iLogID")%>">修改</a>┆<a href="index.asp?id=<%=rs("iLogID")%>&Action=Lock">锁定</a>┆<a href="index.asp?id=<%=rs("iLogID")%>&Action=UnLock">解锁</a>┆<a href="index.asp?id=<%=rs("iLogID")%>&Action=del">删除</a> 
              </div></td>
          </tr>
          <%
			rs.movenext
		next
		%>
          <tr> 
            <td colspan="6" class="hback"> 
              <%
			response.Write "<p>"&  fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
		  end if
		  %>
              　 
              <input name="Action" type="hidden" id="Action">
			  <input type="button" name="Submit" value="删除"  onClick="document.myform.Action.value='del';{if(confirm('确定清除您所选择的记录吗？')){this.document.myform.submit();return true;}return false;}"> 
              <input type="button" name="Submit2" value="批量解锁"  onClick="document.myform.Action.value='UnLock';{if(confirm('确定解锁吗')){this.document.myform.submit();return true;}return false;}"> 
              <input name="Submit3" type="button"  onClick="document.myform.Action.value='Lock';{if(confirm('确定锁定吗？？\n锁定后将不能显示')){this.document.myform.submit();return true;}return false;}" value="批量锁定"> 
            </td>
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






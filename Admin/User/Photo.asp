<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
	dim Conn,User_Conn,strShowErr
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF
	if not MF_Check_Pop_TF("ME_Photo") then Err_Show 
	if not MF_Check_Pop_TF("ME040") then Err_Show 

	Function GetFriendName(f_strNumber)
		Dim RsGetFriendName
		Set RsGetFriendName = User_Conn.Execute("Select UserName From FS_ME_Users Where UserNumber = '"& f_strNumber &"'")
		If  Not RsGetFriendName.eof  Then 
			GetFriendName = RsGetFriendName("UserName")
		Else
			GetFriendName = 0
		End If 
		set RsGetFriendName = nothing
	End Function 
	if Request("Action")="del" then
		if trim(Request("Id"))="" then
			strShowErr = "<li>请选择至少一项</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		User_Conn.execute("Delete From FS_ME_Photo where id in ("&FormatIntArr(Request("Id"))&")")
		strShowErr = "<li>删除成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Photo.asp")
		Response.End
	ElseIf Request("Action")="rec" Then
		User_Conn.execute("Update FS_ME_Photo set isRec=1 where id in ("&FormatIntArr(Request("Id"))&")")
		strShowErr = "<li>修改成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Photo.asp")
		Response.End
	Elseif Request("Action")="unrec" Then 
		User_Conn.execute("Update FS_ME_Photo set isRec=0 where id in ("&FormatIntArr(Request("Id"))&")")
		strShowErr = "<li>修改成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Photo.asp")
		Response.End
	end If
	if Request("Action")="all" then
		User_Conn.execute("Delete From FS_ME_Photo")
			strShowErr = "<li>删除所有成功</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Photo.asp")
			Response.end
	end if
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,i
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
	if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td width="100%" class="xingmu">相册管理</td>
  </tr class="hback">
  <tr class="hback">
    <td class="hback"><a href="UserReport.asp">返回</a>┆<a href="Photo.asp?Action=all" onClick="{if(confirm('确定删除吗？')){return true;}return false;}">清空所有相册</a></td>
  </tr class="hback">
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="myform" method="post" action="">
    <%
			dim rs,isrec
			set rs = Server.CreateObject(G_FS_RS)
			rs.open "select * from FS_ME_Photo  order by id desc",User_Conn,1,1
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
					isrec=rs("isrec")
			%>
    <tr class="hback"> 
      <td width="21%" rowspan="5" class="hback_1"> <table border="0" align="center" cellpadding="2" cellspacing="1" class="table">
          <tr> 
            <td class="hback">
              <%if isnull(trim(rs("PicSavePath"))) then%>
              <img src="Images/nopic_supply.gif" width="90"  id="pic_p_1">
              <%else%>
              <a href="<%=rs("PicSavePath")%>" target="_blank"><img src="<%=rs("PicSavePath")%>" width="90" border="0" id="pic_p_1"></a>
              <%end if%>
            </td>
          </tr>
        </table></td>
      <td width="12%" class="hback"><div align="center"><strong>相片名称：</strong></div></td>
      <td width="40%" class="hback"><font style="font-size:14px"><span class="hback_1">
	  <%if isrec=1 Then Response.write("<img src=""../images/award.gif"" alt=""推荐"" />")%>
	  <strong><%=rs("title")%></strong></span></font><a href="../../<%=G_USER_DIR%>/ShowUser.asp?UserNumber=<% = rs("UserNumber")%>" target="_blank">(<%=GetFriendName(rs("UserNumber"))%>)</a></td>
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
      <td>
        <%
			dim c_rs
			if rs("ClassID")=0 then
				response.Write("没分类")
			else
				set c_rs=User_Conn.execute("select ID,title From Fs_me_photoclass where id="&rs("classid"))
				Response.Write c_rs("title")
				c_rs.close:set c_rs=nothing
			end if
			%>
      </td>
      <td colspan="2"><div align="center"><a href="Photo.asp?id=<%=rs("id")%>&Action=del" onClick="{if(confirm('确定通过删除吗？')){return true;}return false;}">删除</a>  |  
	  <a href="Photo.asp?id=<%=rs("id")%>&Action=rec">推荐</a> |
	  <a href="Photo.asp?id=<%=rs("id")%>&Action=unrec">取消推荐</a> |
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
        <input name="Action" type="hidden" id="Action"> <input type="button" name="Submit" value="删除"  onClick="document.myform.Action.value='del';{if(confirm('确定清除您所选择的记录吗？')){this.document.myform.submit();return true;}return false;}"></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
Conn.close:set conn=nothing
User_Conn.close:set User_Conn=nothing
%>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<myform.elements.length;i++)  
    {  
    var e = myform.elements[i];  
    if (e.name != 'chkall')  
       e.checked = myform.chkall.checked;  
    }  
  }
</script>







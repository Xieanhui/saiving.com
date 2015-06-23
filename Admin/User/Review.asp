<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<!--#include file="../../FS_InterFace/CLS_Foosun.asp" -->
<!--#include file="../../FS_InterFace/MS_Public.asp" -->
<!--#include file="../../FS_InterFace/DS_Public.asp" -->
<!--#include file="../../FS_InterFace/ME_Public.asp" -->
<!--#include file="../../FS_InterFace/SD_Public.asp" -->
<!--#include file="../../FS_InterFace/HS_Public.asp" -->
<%
	dim Conn,User_Conn,strShowErr
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF
	if not MF_Check_Pop_TF("ME_Review") then Err_Show 
	if not MF_Check_Pop_TF("ME038") then Err_Show 

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
	if Request.Form("Action")="Del" then
		if trim(Request.Form("Id"))="" then
			strShowErr = "<li>请选择至少一项</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		User_Conn.execute("Delete From FS_ME_Review where ReviewID in ("&FormatIntArr(Request.Form("Id"))&")")
			strShowErr = "<li>删除成功</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Review.asp")
			Response.end
	end if
	if Request.Form("Action")="DelAll" then
		User_Conn.execute("Delete From FS_ME_Review")
		Call MF_Insert_oper_Log("删除评论","删除了所有评论",now,session("admin_name"),"ME")
			strShowErr = "<li>删除所有评论成功</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Review.asp")
			Response.end
	end if
	if Request.Form("Action")="UnLock" then
		if trim(Request.Form("Id"))="" then
			strShowErr = "<li>请选择至少一项</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		User_Conn.execute("Update FS_ME_Review set AdminLock=0 where IsLock=0 and ReviewID in ("&FormatIntArr(Request.Form("Id"))&")")
			strShowErr = "<li>批量审核成功</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Review.asp")
			Response.end
	end if
	if Request.Form("Action")="Lock" then
		if trim(Request.Form("Id"))="" then
			strShowErr = "<li>请选择至少一项</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		User_Conn.execute("Update FS_ME_Review set AdminLock=1 where IsLock=0 and ReviewID in ("&FormatIntArr(Request.Form("Id"))&")")
			strShowErr = "<li>批量锁定成功</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Review.asp")
			Response.end
	end if
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,i
	
	int_RPP=25 '设置每页显示数目
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
    <td width="100%" class="xingmu">评论管理</td>
  </tr class="hback">
  <tr class="hback">
    <td class="hback"><a href="Review.asp">全部</a>┆<a href="Review.asp?isLock=1">未审核</a>┆<a href="Review.asp?isLock=0">已审核</a>┆<a href="Review.asp?type=0">新闻评论</a>┆<a href="Review.asp?type=1">下载评论</a>┆<a href="Review.asp?type=2">商品评论</a>┆<a href="Review.asp?type=3">房产评论</a>┆<a href="Review.asp?type=4">供求评论</a>┆<a href="Review.asp?type=5">日记评论</a>┆<a href="Review.asp?type=6">相册评论</a></td>
  </tr class="hback">
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form action="Review.asp"  method="post" name="form1" id="form1">
    <tr class="hback"> 
      <td width="25%" class="xingmu"><div align="left"><strong> 标题</strong></div></td>
      <td width="14%" class="xingmu"><div align="center"><strong>IP</strong></div></td>
      <td width="13%" class="xingmu"><div align="center"><strong>发表人</strong></div></td>
      <td width="16%" class="xingmu"><div align="center"><strong>日期</strong></div></td>
      <td width="7%" class="xingmu"><div align="center"><strong>类型</strong></div></td>
      <td width="14%" class="xingmu"><div align="center"><strong>状态</strong></div></td>
	  <td width="14%" class="xingmu"><div align="center"><strong>主题</strong></div></td>
      <td width="4%" class="xingmu">&nbsp;</td>
    </tr>
    <%
		dim rs_reviewsql,rs_review,str_type,str_isLock
		strpage=request("page")
		if len(strpage)=0 Or strpage<1 or trim(strpage)=""  Then strpage="1"
		if NoSqlHack(Request.QueryString("type"))<>"" then:str_type=" and ReviewTypes="&NoSqlHack(Request.QueryString("type"))&"":else:str_type="":end if
		if NoSqlHack(Request.QueryString("isLock"))<>"" then:str_isLock=" and AdminLock="&NoSqlHack(Request.QueryString("isLock"))&"":else:str_isLock="":end if
		Set rs_review = Server.CreateObject(G_FS_RS)
		rs_reviewsql = "Select * From FS_ME_Review  where 1=1 "& str_type & str_isLock &"  order by  Addtime desc, ReviewID desc"
		rs_review.Open rs_reviewsql,User_Conn,1,1
		if rs_review.eof then
		   rs_review.close
		   set rs_review=nothing
		   Response.Write"<tr  class=""hback""><td colspan=""8""  class=""hback"" height=""40"">没有记录。</td></tr>"
		else
			rs_review.PageSize=int_RPP
			cPageNo=NoSqlHack(Request.QueryString("Page"))
			If cPageNo="" Then cPageNo = 1
			If not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo>rs_review.PageCount Then cPageNo=rs_review.PageCount 
			If cPageNo<=0 Then cPageNo=1
			rs_review.AbsolutePage=cPageNo
			for i=1 to int_RPP
				if rs_review.eof Then exit For 
		%>
    <tr class="hback"> 
      <td class="hback"><div align="left" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(rid<%=rs_review("ReviewID")%>);" language=javascript> 
          <a href="#"> 
          <% = rs_review("title")%>
          </a></div></td>
      <td class="hback"><div align="center"><% = rs_review("ReviewIP")%></div></td>
      <td class="hback"><div align="center"><a href="../../<%=G_USER_DIR%>/ShowUser.asp?UserNumber=<% = rs_review("UserNumber")%>" target="_blank"> 
          <% = GetFriendName(rs_review("UserNumber"))%>
          </a> </div></td>
      <td class="hback"><div align="center"> 
          <% = rs_review("addtime")%>
        </div></td>
      <td class="hback"><div align="center"> 
          <%'0为新闻评论，1为下载评论，2为商品，3为房产评论，4为供求评论，5日记,6相册
			select case rs_review("ReviewTypes")
				case 0
					response.Write"<a href=Review.asp?type=0>新闻</a>"
				case 1
					response.Write"<a href=Review.asp?type=1>下载</a>"
				case 2
					response.Write"<a href=Review.asp?type=2>商品</a>"
				case 3
					response.Write"<a href=Review.asp?type=3>房产</a>"
				case 4
					response.Write"<a href=Review.asp?type=4>供求</a>"
				case 5
					response.Write"<a href=Review.asp?type=5>日志</a>"
				case 6
					response.Write"<a href=Review.asp?type=6>相册</a>"
				case else
					response.Write"<a href=Review.asp>-</a>"
			end select
			%>
        </div></td>
      <td class="hback"><div align="center"> 
          <%
			if rs_review("IsLock")="1" then
				response.Write("用户锁定┆")
			else
				response.Write("用户开放┆")
			end if
			if rs_review("AdminLock")="1" then
				response.Write("<span class=tx>未审核</span>")
			else
				response.Write("已审核")
			end if
		  %>
        </div></td>
      <td class="hback"><div align="center"> 
         	<span onMouseUp="opencat(Zhuti<%=rs_review("ReviewID")%>);" language=javascript><a href="#" title="点击查看该评论所属主题">点击查看</a></span>
        </div></td>
	  <td class="hback"><div align="center"> 
          <input name="ID" type="checkbox" id="ID" value="<% = rs_review("ReviewID")%>">
        </div></td>
    </tr>
    <tr class="hback" id="rid<%=rs_review("ReviewID")%>" style="display:none"> 
      <td height="31" colspan="8" class="hback"> <strong>评论内容:</strong> 
        <% = rs_review("Content")%>
      </td>
    </tr>
	<tr class="hback" id="Zhuti<%=rs_review("ReviewID")%>" style="display:none"> 
      <td height="31" colspan="8" class="hback"> <strong>所属主题:</strong> 
        <% = Get_Info(rs_review("InfoID"),rs_review("ReviewTypes"))%>
      </td>
    </tr>
    <%
		  rs_review.MoveNext
	  Next
	  %>
    <tr class="hback"> 
      <td width="80%" colspan="8"><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left"><%
			response.Write "<p>"&  fPageCount(rs_review,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
			%></td>
    <td><input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)"><input name="Action" type="hidden" id="Action" value="">
              <input type="button" name="Submit2" value="删除"  onClick="document.form1.Action.value='Del';{if(confirm('确定清除您所选择的记录吗？')){this.document.form1.submit();return true;}return false;}">
              <input type="button" name="Submit22" value="批量审核"  onClick="document.form1.Action.value='UnLock';{if(confirm('您确定要批量审核通过吗？')){this.document.form1.submit();return true;}return false;}">
              <input type="button" name="Submit23" value="批量锁定"  onClick="document.form1.Action.value='Lock';{if(confirm('您确定要锁定？')){this.document.form1.submit();return true;}return false;}">
              <input type="button" name="Submit232" value="删除所有"  onClick="document.form1.Action.value='DelAll';{if(confirm('您确定要删除所有评论吗？\n删除后将不能恢复!!!')){this.document.form1.submit();return true;}return false;}"></td>
  </tr>
</table>
 </td>
    </tr>
  </FORM>
</table>
		  </body>
</html>
<%
'------2007-01-05 by ken
Function Get_Info(InfoID,TypeStr)
	Dim Sql,Rs,NewClassObj,f_NewsLinkRecordSet,f_NewsLink
	If InfoID = "" Then
		Get_Info = "未知"
		Exit Function
	End If
	Select Case TypeStr
		Case 0
			Sql = "Select NewsID,NewsTitle,ClassEName,[Domain],SavePath,News.IsURL,News.URLAddress,SaveNewsPath,FileName,News.FileExtName From Fs_Ns_News as News,FS_NS_NewsClass as Class Where News.ClassID=Class.ClassID And News.ID = " & CintStr(InfoID)
			Set Rs = Conn.ExeCute(Sql)
			If Not Rs.Eof Then
				Set f_NewsLinkRecordSet = New CLS_FoosunRecordSet
				Set f_NewsLinkRecordSet.Values("ClassEName,Domain,SavePath,IsURL,URLAddress,SaveNewsPath,FileName,FileExtName") = Rs
				Set f_NewsLink = New CLS_FoosunLink
				Get_Info = "<a target=""_blank"" href=""" & f_NewsLink.NewsLink(f_NewsLinkRecordSet) & """>" & Rs(1) & "</a>"
				Set f_NewsLink = Nothing
				Set f_NewsLinkRecordSet = Nothing
			Else
				Get_Info = "未知"
			End If
			Rs.Close : Set Rs = Nothing		 	
		Case 1
			Set NewClassObj = New cls_DS
			Sql = "Select DownLoadID,[Name] From FS_DS_List Where ID = " & CintStr(InfoID)
			Set Rs = Conn.ExeCute(Sql)
			If Not Rs.Eof Then
				Get_Info = "<a target=""_blank"" href=""" & NewClassObj.get_DownLink(Rs(0)) & """>" & Rs(1) & "</a>"
			Else
				Get_Info = "未知"
			End If
			Rs.Close : Set Rs = Nothing	
		Case 2
			Set NewClassObj = New cls_MS
			Sql = "Select ID,ProductTitle From FS_MS_Products Where ID = " & InfoID
			Set Rs = Conn.ExeCute(Sql)
			If Not Rs.Eof Then
				Get_Info = "<a target=""_blank"" href=""" & NewClassObj.get_productsLink(Rs(0)) & """>" & Rs(1) & "</a>"
			Else
				Get_Info = "未知"
			End If
			Rs.Close : Set Rs = Nothing	
		Case 3
			Get_Info = "房产"
		Case 4
			Set NewClassObj = New cls_SD
			Sql = "Select ID,PubTitle From FS_SD_News Where ID = " & InfoID
			Set Rs = Conn.ExeCute(Sql)
			If Not Rs.Eof Then
				Get_Info = "<a target=""_blank"" href=""" & NewClassObj.getPageLink(Rs(0)) & """>" & Rs(1) & "</a>"
			Else
				Get_Info = "未知"
			End If
			Rs.Close : Set Rs = Nothing
		Case 5
			Set NewClassObj = New cls_ME
			Sql = "Select iLogID,Title From FS_ME_Infoilog Where iLogID = " & InfoID
			Set Rs = User_Conn.ExeCute(Sql)
			If Not Rs.Eof Then
				Get_Info = "<a target=""_blank"" href=""" & NewClassObj.Get_LogLink(Rs(0)) & """>" & Rs(1) & "</a>"
			Else
				Get_Info = "未知"
			End If
			Rs.Close : Set Rs = Nothing
		Case 6
			Sql = "Select ID,title From FS_ME_Photo Where ID = " & InfoID
			Set Rs = User_Conn.ExeCute(Sql)
			If Not Rs.Eof Then
				Get_Info = "<a target=""_blank"" href=""http://" & request.Cookies("FoosunMFCookies")("FoosunMFDomain")& "/" & Request.Cookies("FoosunMECookies")("FoosunMELogDir")&"/ShowPhoto.asp?Id=" &Rs(0)&""">" & Rs(1) & "</a>"
			Else
				Get_Info = "未知"
			End If
			Rs.Close : Set Rs = Nothing
		Case Else
			Get_Info = "未知"							
	End Select
	Get_Info = Get_Info
End Function
'-----------------------------
set conn=nothing
set User_Conn=nothing
end if
%>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form1.elements.length;i++)  
    {  
    var e = form1.elements[i];  
    if (e.name != 'chkall')  
       e.checked = form1.chkall.checked;  
    }  
  }
function opencat(cat)
{
  if(cat.style.display=="none"){
     cat.style.display="";
  } else {
     cat.style.display="none"; 
  }
}
</script> 







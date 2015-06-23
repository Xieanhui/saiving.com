<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.CacheControl = "no-cache"
	Dim Conn,User_Conn
	Dim CharIndexStr,strShowErr
	Dim Fs_news,obj_news_rs,obj_news_rs_1,isUrlStr,str_Href,obj_cnews_rs,news_count,str_Href_title,str_action,str_ClassID,news_SQL
	Dim obj_newslist_rs,newslist_sql,strpage,str_showTF,str_ClassID_1,str_Editor,str_Keyword,str_GetKeyword,str_ktype,tmp_draft
	Dim select_count,select_pagecount,i,Str_GetPopID,Str_PopID,str_check,str_UrlTitle,icNum,str_addType,str_addType_1
	Dim str_Rec,str_isTop,str_hot,str_pic,str_highlight,str_bignews,str_filt,str_Constr,str_Top,tmp_pictf,tmp_isRecyle
	Dim str_s_classIDarray,tmp_splitarrey_id,tmp_splitarrey_Classid,tmp_i,str_Move_type,str_t_classID,C_NewsIDarrey,Tmp_rs,Tmp_TF_Rs
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
	
	int_RPP=20'设置每页显示数目
	int_showNumberLink_=8 '数字导航显示数目
	showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
	str_nonLinkColor_="#999999" '非热链接颜色
	toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
	toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
	toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
	toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
	toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
	toL_="<font face=webdings title=""最后一页"">:</font>"

	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF 
	'权限判断
	'Call MF_Check_Pop_TF("NS_Class_000001") 
	'得到会员组列表 
	set Fs_news = new Cls_News
	Fs_News.GetSysParam()
	If Not Fs_news.IsSelfRefer Then response.write "非法提交数据":Response.end
	if Request("Action") = "signDel" then
		if fs_news.ReycleTF = 1 then
			Conn.execute("Update FS_NS_News set isRecyle = 1 where NewsID='"& NoSqlHack(Request.QueryString("NewsID"))&"'")
			strShowErr = "<li>"& Fs_news.allInfotitle &"已经删除</li><li>"& Fs_news.allInfotitle &"已经临时放到回收站中</li>"
		Else
			strShowErr = "<li>"& Fs_news.allInfotitle &"已经彻底删除</li>"
			Conn.execute("Delete From FS_NS_News where NewsID='"& NoSqlHack(Request.QueryString("NewsID"))&"'")
			'删除浏览权限新闻，以防产生垃圾信息
			Conn.execute("Delete From  FS_MF_POP  where InfoID='"& NoSqlHack(Request.QueryString("NewsID"))&"' and PopType='NS'")
		End if
		'删除静态文件
		'******************保留
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	if Request("Action") = "draftDel" then
		strShowErr = "<li>"& Fs_news.allInfotitle &"已经彻底删除</li>"
		Conn.execute("Delete From FS_NS_News where NewsID='"& NoSqlHack(Request.QueryString("NewsID"))&"'")
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新闻管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="hback"> 
    <td width="100%" class="xingmu"><a href="#" class="sd"><strong>我的工作目录</strong></a><a href="../../help?Lable=NS_News_MyFolder" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a>　　　　　　　　　　　　　　 
    </td>
  </tr>
  <tr> 
    <td height="18" class="hback"><div align="left"><img src="../Images/all_article_icon.gif" width="11" height="10" border="0">&nbsp;<a href="News_MyFolder.asp" 您所有的文章都在这里>所有<% =  Fs_news.allInfotitle %>
        </a>&nbsp;┆<img src="../Images/draft_icon.gif" width="10" height="9">&nbsp;<a href="News_MyFolder.asp?Action=draft" title="草稿箱里最多放20篇<% =  Fs_news.allInfotitle %>在这里">草稿夹</a>&nbsp;┆<img src="../Images/recycle_icon.gif" width="10" height="11">&nbsp;<a href="News_MyFolder.asp?Action=reycle">回收站</a>&nbsp;┆<a href="Class_add.asp?ClassID=<%=Request.QueryString("classid")%>&Action=add">创建栏目</a>&nbsp;┆<a href="News_add.asp?ClassID=<%=Request.QueryString("classid")%>">&nbsp;创建<% =  Fs_news.allInfotitle %>
        </a> </div></td>
  </tr>
</table>
<%if Request.QueryString("ClassId")<>"" then%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr>
    <td class="hback"> 位置导航：<A href="News_Manage.asp">全部新闻</a>
      <%response.write fs_news.navigation(NoSqlHack(Request.QueryString("ClassID")))%></td>
  </tr>
</table>
<%end if%>
<%
	Dim AndSQL
	AndSQL = GetAndSQLOfSearchClass("NS013")

	If NoSqlHack(Trim(Request.QueryString("classid"))) <> "" Then
		str_ClassID_1 = NoSqlHack(Trim(Request.QueryString("classid")))
	Else
		str_ClassID_1 = 0
	End If	
	Set obj_news_rs = server.CreateObject(G_FS_RS)
	news_SQL = "Select Orderid,id,ClassID,ClassName,ClassEName,IsUrl,AddNewsType from FS_NS_NewsClass where Parentid  = '" & str_ClassID_1 & "'  and ReycleTF=0 " & AndSQL & "  Order by Orderid desc,ID desc"
	obj_news_rs.Open news_SQL,Conn,1,3
	if fs_news.addNewsType = 1 then str_addType_1 ="News_add.asp":else:str_addType_1 ="News_add_Conc.asp":end if
	IF Not obj_news_rs.Eof Then
		With Response
			.Write "<table width=""98%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"" class=""table"">"
			.Write "<tr class=""hback"">"
			.Write "<td>"
			.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"">"
			.Write "<tr>"
		End With	
		icNum = 0
		Do while Not obj_news_rs.eof 
			if obj_news_rs("AddNewsType") =1 then
				str_addType = "News_add.asp"
			Else
				str_addType ="News_add_Conc.asp"
			End if
			if obj_news_rs("IsUrl") = 1 then
				isUrlStr = "(<span class=""tx"">外</span>)"
				str_Href = ""
				str_Href_title = ""& obj_news_rs("ClassName") &""
			elseif obj_news_rs("IsUrl") = 2 then
				isUrlStr = "(<span class=""tx"">单</span>)"
				str_Href = ""
				str_Href_title = ""& obj_news_rs("ClassName") &""
			Else
				isUrlStr = ""
				if Get_SubPop_TF(obj_news_rs("ClassID"),"NS001","NS","news") then
					str_Href = "("&"<a href="& str_addType &"?ClassID="&obj_news_rs("ClassID")&"><img src=""../images/add.gif"" border=""0"" alt=""添加"& Fs_news.allInfotitle &"""></a>"&")"
				else
					str_Href = ""
				end if
				str_Href_title = "<a href=""News_MyFolder.asp?ClassID="& obj_news_rs("ClassID") &"&Action="& Request.QueryString("Action")&""" title=""点击进入下一级栏目"">"& obj_news_rs("ClassName") &"</a>"
			End if
			Set obj_news_rs_1 = server.CreateObject(G_FS_RS)
			obj_news_rs_1.Open "Select Count(ID) from FS_NS_NewsClass where ParentID='"& obj_news_rs("ClassID") &"'",Conn,1,1
			if obj_news_rs_1(0)>0 then
				str_action=  "<img src=""images/+.gif""></img>"& str_Href_title &""
			Else
				str_action=  "<img src=""images/-.gif""></img>"& str_Href_title &""
			End if
			obj_news_rs_1.close:set obj_news_rs_1 =nothing
			'得到新闻数量
			if obj_news_rs("IsUrl") = 0 then
				Set obj_cnews_rs = server.CreateObject(G_FS_RS)
				obj_cnews_rs.Open "Select ID from FS_NS_News where ClassID='"& obj_news_rs("ClassID") &"' and  isRecyle=0 and Editor='"& session("Admin_Name") &"' ",Conn,1,1
				obj_cnews_rs.close:set obj_cnews_rs = nothing
			Else
				news_count = ""
			End if
			Response.Write"<td height=""22"">"
			Response.Write str_action&isUrlStr&str_Href
			Response.Write "</td>"
			obj_news_rs.MoveNext
			icNum = icNum + 1
			if icNum mod 5 = 0 then
				Response.Write("</tr><tr>")
			End if
		loop
		Response.Write "</tr>"
		Response.Write "</table></td>"
		Response.Write "</tr>"
		Response.Write "</table>"
	End If	
%>
<%
strpage=request("page")
if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
if Trim(Request.QueryString("Action")) = "draft" then:tmp_draft = " and isdraft=1 and  isRecyle=0":Else:tmp_draft = "":End if
if Trim(Request.QueryString("Action")) = "reycle" then:tmp_isRecyle = " and isRecyle=1":Else:tmp_isRecyle = "":End if
if Trim(Request.QueryString("ClassID")) <>"" then:str_ClassID = " and ClassID='"& NoSqlHack(Request.QueryString("Classid"))&"'":Else:str_ClassID = "":End if
newslist_sql = "Select ID,NewsID,PopID,ClassID,NewsTitle,IsURL,isPicNews,URLAddress,Editor,Hits,NewsProperty,isLock,isdraft,isRecyle,addtime,author,source from FS_NS_News where Editor='"& session("Admin_Name")&"' "& tmp_draft & tmp_isRecyle & str_ClassID &" Order by PopID desc,addtime desc,ID desc"
Set obj_newslist_rs = Server.CreateObject(G_FS_RS)
obj_newslist_rs.Open newslist_sql,Conn,1,1
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="myForm" method="post" action="News_Manage.asp">
    <tr class="xingmu"> 
      <td colspan="2" class="xingmu"> <div align="center"> </div>
        <div align="center"> 
          <% =  Fs_news.allInfotitle %>
          标题</div></td>
      <td width="12%" class="xingmu"><div align="center">状态</div></td>
      <td width="7%" class="xingmu"><div align="center">审核</div></td>
      <td width="30%" class="xingmu"><div align="center">操作</div></td>
    </tr>
    <%
		if obj_newslist_rs.eof then
			   obj_newslist_rs.close
			   set obj_newslist_rs=nothing
			   Response.Write"<TR  class=""hback""><TD colspan=""7""  class=""hback"" height=""40"">没有"& Fs_news.allInfotitle &"。</TD></TR>"
		else
			   str_showTF = 1
				obj_newslist_rs.PageSize=int_RPP
				cPageNo=NoSqlHack(Request.QueryString("Page"))
				If cPageNo="" Then cPageNo = 1
				If not isnumeric(cPageNo) Then cPageNo = 1
				cPageNo = Clng(cPageNo)
				If cPageNo<=0 Then cPageNo=1
				If cPageNo>obj_newslist_rs.PageCount Then cPageNo=obj_newslist_rs.PageCount 
				obj_newslist_rs.AbsolutePage=cPageNo
				for i=1 to obj_newslist_rs.pagesize
					if obj_newslist_rs.eof Then exit For 
						Str_GetPopID = obj_newslist_rs("PopID")
						if Str_GetPopID = 5 then
							Str_PopID = "<IMG Src=""images/newstype/5.gif"" border=""0"" alt=""总置顶"& Fs_news.allInfotitle &",点击查看简洁内容"">"
							str_Top = "<a href=News_Manage.asp?NewsID="& obj_newslist_rs("NewsID")&"&Action=signUnTop  onClick=""{if(confirm('确定解除总固顶吗？')){return true;}return false;}"">解固</a>"
						Elseif Str_GetPopID = 4 then
							Str_PopID = "<IMG Src=""images/newstype/4.gif"" border=""0"" alt=""栏目置顶"& Fs_news.allInfotitle &",点击查看简洁内容"">"
							str_Top = "<a href=News_Manage.asp?NewsID="& obj_newslist_rs("NewsID")&"&Action=signUnTop  onClick=""{if(confirm('确定解除栏目固顶吗？')){return true;}return false;}"">解固</a>"
						Elseif Str_GetPopID = 3 then
							Str_PopID = "<IMG Src=""images/newstype/3.gif"" border=""0"" alt=""总推荐"& Fs_news.allInfotitle &",点击查看简洁内容"">"
							str_Top = "<a href=News_Manage.asp?NewsID="& obj_newslist_rs("NewsID")&"&Action=signTop  onClick=""{if(confirm('确定固顶吗？')){return true;}return false;}"">固顶</a>"
						Elseif Str_GetPopID = 2 then
							Str_PopID = "<IMG Src=""images/newstype/2.gif"" border=""0"" alt=""栏目推荐"& Fs_news.allInfotitle &",点击查看简洁内容"">"
							str_Top = "<a href=News_Manage.asp?NewsID="& obj_newslist_rs("NewsID")&"&Action=signTop onClick=""{if(confirm('确定固顶吗？')){return true;}return false;}"">固顶</a>"
						Elseif Str_GetPopID = 0 then
							Str_PopID = "<IMG Src=""images/newstype/0.gif"" border=""0"" alt=""一般"& Fs_news.allInfotitle &",点击查看简洁内容"">"
							str_Top = "<a href=News_Manage.asp?NewsID="& obj_newslist_rs("NewsID")&"&Action=signTop onClick=""{if(confirm('确定固顶吗？')){return true;}return false;}"">固顶</a>"
						End if
						if obj_newslist_rs("isUrl") = 1 then
							str_UrlTitle = "<a href="""& obj_newslist_rs("URLAddress") &""" target=""_blank""><Img src=""../images/folder/url.gif"" border=""0"" alt=""标题新闻,点击转到标题地址""></img></a>"
						Else
							str_UrlTitle = ""
							if obj_newslist_rs("isPicNews") = 1 then
								tmp_pictf="<a href=""javascript:m_PicUrl('News_Pic_Modify.asp?NewsID="&obj_newslist_rs("NewsID")&"')""><Img src=""../images/folder/img.gif"" alt=""图片新闻,点击更改图片"" border=""0""></img></a>"
							else
								tmp_pictf="<Img src=""../images/folder/folder_1.gif"" alt=""文字新闻""></img>"
							end if
						end if
	%>
    <tr> 
      <td width="6%" height="18" class="hback" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(M_Newsid<% = obj_newslist_rs("ID")%>);"  language=javascript><% = Str_PopID %></td>
      <td width="45%" class="hback"> <% = str_UrlTitle %> <% = tmp_pictf %> <a href="News_edit.asp?NewsID=<% = obj_newslist_rs("NewsID")%>&ClassID=<% = obj_newslist_rs("ClassID")%>" title="添加日期：<% = obj_newslist_rs("addtime")%>"> 
        <% = GotTopic(obj_newslist_rs("Newstitle"),50)%>
        </a></td>
      <td class="hback"> <div align="center"><%
	  if obj_newslist_rs("isRecyle")=1 then
	  	Response.Write("<span class=""tx"">回收站</span>")
	  Elseif obj_newslist_rs("isRecyle")=0 and obj_newslist_rs("isdraft")=1 then
	  	Response.Write("草稿箱")
	  End if
	  if obj_newslist_rs("isRecyle")=0 and obj_newslist_rs("isdraft")=0 then
	  	response.Write("正常"& Fs_news.allInfotitle &"")
	  End if
	  %></div></td>
      <td class="hback"><div align="center"> 
          <%if obj_newslist_rs("isLock")=1 then response.Write"<a href=""News_Manage.asp?NewsID="& obj_newslist_rs("NewsId") &"&Action=singleCheck"" onClick=""{if(confirm('确定通过审核吗？')){return true;}return false;}""><span class=""tx""><b>×</b></span></a>":else response.Write"<a href=""News_Manage.asp?NewsID="& obj_newslist_rs("NewsId") &"&Action=singleUnCheck"" onClick=""{if(confirm('确定锁定吗？')){return true;}return false;}""><b>√</b></a>"%>
        </div></td>
      <td class="hback"><div align="center"> 
          <%if obj_newslist_rs("isRecyle")=1 then%>
          <a href="News_Recyle.asp">回收站</a>
		<%Elseif obj_newslist_rs("isdraft")=1 then%>
          <a href="News_edit.asp?NewsID=<% = obj_newslist_rs("NewsID")%>&ClassID=<% = obj_newslist_rs("ClassID")%>">修改</a>｜<a href="News_MyFolder.asp?NewsID=<% = obj_newslist_rs("NewsID")%>&Action=draftDel"  onClick="{if(confirm('确定要删除吗？')){return true;}return false;}">删除</a> 
		  <%Else%>
          <a href="News_edit.asp?NewsID=<% = obj_newslist_rs("NewsID")%>&ClassID=<% = obj_newslist_rs("ClassID")%>">修改</a>｜<a href="News_MyFolder.asp?NewsID=<% = obj_newslist_rs("NewsID")%>&Action=signDel"  onClick="{if(confirm('确定要删除吗？\n\n如果你在系统参数设置中设置删除<% =  Fs_news.allInfotitle %>到回收站\n<% =  Fs_news.allInfotitle %>将删除到回收站中!\n必要时候可还原')){return true;}return false;}">删除</a> 
          <%End if%>
        </div></td>
    </tr>
    <tr id="M_Newsid<% = obj_newslist_rs("ID")%>" style="display:none"> 
      <td height="35" colspan="5" class="hback"> <table width="100%" border="0" cellspacing="1" cellpadding="2" class="table">
          <tr class="hback"> 
            <td width="45%" height="20" class="hback"><font style="font-size:12px"> 
              <% =  Fs_news.allInfotitle %>
              类型： 
              <%
		if  split(obj_newslist_rs("NewsProperty"),",")(0) then Response.Write("荐｜")
		if  split(obj_newslist_rs("NewsProperty"),",")(1) then Response.Write("滚｜")
		if  split(obj_newslist_rs("NewsProperty"),",")(2) then Response.Write("评｜")
		if  split(obj_newslist_rs("NewsProperty"),",")(3) then Response.Write("稿｜")
		if  split(obj_newslist_rs("NewsProperty"),",")(4) then Response.Write("远图｜")
		if  split(obj_newslist_rs("NewsProperty"),",")(5) then Response.Write("头｜")
		if  split(obj_newslist_rs("NewsProperty"),",")(6) then Response.Write("热｜")
		if  split(obj_newslist_rs("NewsProperty"),",")(7) then Response.Write("精｜")
		if  split(obj_newslist_rs("NewsProperty"),",")(8) then Response.Write("规｜")
		if  split(obj_newslist_rs("NewsProperty"),",")(9) then Response.Write("公｜")
		if  split(obj_newslist_rs("NewsProperty"),",")(10) then Response.Write("幻｜")
		%>
              </font></td>
            <td width="22%" class="hback"><font style="font-size:12px">日期： 
              <% = obj_newslist_rs("addtime")%>
              </font></td>
            <td width="14%" class="hback"><font style="font-size:12px">作者： 
              <% = obj_newslist_rs("author")%>
              </font></td>
            <td width="19%" class="hback"><font style="font-size:12px">来源： 
              <% = obj_newslist_rs("source")%>
              </font></td>
          </tr>
        </table></td>
    </tr>
    <%
			  obj_newslist_rs.MoveNext
		  Next
	%>
  </form>
  <tr> 
    <td height="18" colspan="5" class="hback"> <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="79%" colspan="2" align="right"> <%
			response.Write "<p>"&  fPageCount(obj_newslist_rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
		End if
	%> </td>
        </tr>
      </table></td>
  </tr>
</table>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
</body>
</html>
<%
set Fs_news = nothing
%>
<script language="JavaScript" type="text/JavaScript" src="js/Public.js"></script>
<script language="JavaScript" type="text/JavaScript">
function opencat(cat)
{
  if(cat.style.display=="none"){
     cat.style.display="";
  } else {
     cat.style.display="none"; 
  }
}
function SelectAllClass(){
  for(var i=0;i<document.form_m.s_Classid.length;i++){
    document.form_m.s_Classid.options[i].selected=true;}
}
function UnSelectAllClass(){
  for(var i=0;i<document.form_m.s_Classid.length;i++){
    document.form_m.s_Classid.options[i].selected=false;}
}

function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = myForm.elements[i];  
    if (e.name != 'chkall')  
       e.checked = myForm.chkall.checked;  
    }  
	}
function m_PicUrl(gotoURL) {
	   var open_url = gotoURL;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=1,width=550,height=480');
}
</script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






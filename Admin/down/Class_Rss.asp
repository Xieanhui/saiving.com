<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
	Dim Conn,User_Conn
	Dim CharIndexStr
	Dim Fs_news,obj_news_rs,obj_news_rs_1,isUrlStr,str_Href,obj_cnews_rs,news_count,str_Href_title,str_action,str_ClassID,news_SQL
	Dim obj_newslist_rs,newslist_sql,strpage,str_showTF,str_ClassID_1,str_Editor,str_Keyword,str_GetKeyword,str_ktype
	Dim select_count,select_pagecount,i,Str_GetPopID,Str_PopID,str_check,str_UrlTitle,icNum,str_addType,str_addType_1
	Dim str_Rec,str_isTop,str_hot,str_pic,str_highlight,str_bignews,str_filt,str_Constr,str_Top,tmp_pictf
	Dim str_s_classIDarray,tmp_splitarrey_id,tmp_splitarrey_Classid,tmp_i,str_Move_type,str_t_classID,C_NewsIDarrey,Tmp_rs,Tmp_TF_Rs
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
	
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
	
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF 
	'权限判断
	'Call MF_Check_Pop_TF("NS_Class_000001") 
	'得到会员组列表 
	set Fs_news = new Cls_News
	Fs_News.GetSysParam()
	If Not Fs_news.IsSelfRefer Then response.write "非法提交数据":Response.end
	str_ClassID = NoSqlHack(Request.QueryString("ClassID"))
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>RSS___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu"> <strong>XML</strong><a href="../../help?Lable=NS_Class_RSS" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a>　　　　　　　　　　　　　　 
      <%
	if Trim(Request.QueryString("ClassID")) <>"" Then
		Response.Write "位置：<a href=""Class_Rss.asp"" class=""sd""><b>XML</b></a>&nbsp;>>&nbsp;"&Fs_news.GetAdd_ClassName(NoSqlHack(Request.QueryString("ClassID")))
	Else
		Response.Write"位置：所有XML"
	End if
	if str_ClassID<>"" then
		news_SQL = "Select Orderid,id,ClassID,ClassName,ClassEName,IsUrl,AddNewsType from FS_DS_Class where Parentid  = '"& NoSqlHack(str_ClassID) &"' and ReycleTF=0 Order by Orderid desc,ID desc"
	Else
		news_SQL = "Select Orderid,id,ClassID,ClassName,ClassEName,IsUrl,AddNewsType from FS_DS_Class where Parentid  = '0'  and ReycleTF=0  Order by Orderid desc,ID desc"
	End if
	Set obj_news_rs = server.CreateObject(G_FS_RS)
	obj_news_rs.Open news_SQL,Conn,1,3
	if fs_news.addNewsType = 1 then str_addType_1 ="News_add.asp":else:str_addType_1 ="News_add_Conc.asp":end if
	%> </td>
  </tr>
  <tr> 
    <form name="form1" method="post" action="">
      <td width="94%" height="18" class="hback"> <div align="left"><a href="Class_rss.asp">首页</a> 
          | <a href="Class_Rss.asp?ClassID=<%=Request.QueryString("ClassID")%>">所有<% =  Fs_news.allInfotitle %>
          </a> |&nbsp; <a href="Class_Rss.asp?ClassID=<%=Request.QueryString("ClassID")%>&isCheck=1&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">已审核</a> 
          &nbsp;|&nbsp; <a href="Class_Rss.asp?ClassID=<%=Request.QueryString("ClassID")%>&isCheck=0&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">未审核</a> 
          | <a href="Class_Rss.asp?ClassID=<%=Request.QueryString("ClassID")%>&NewsTyp=Constr&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">投稿</a> 
          | <a href="Class_Rss.asp?ClassID=<%=Request.QueryString("ClassID")%>&NewsTyp=Constr&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>"></a><a href="Class_Rss.asp?ClassID=<%=Request.QueryString("ClassID")%>&NewsTyp=recTF">推荐 
          </a> | <a href="Class_Rss.asp?ClassID=<%=Request.QueryString("ClassID")%>&NewsTyp=isTop&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">置顶 
          </a> | <a href="Class_Rss.asp?ClassID=<%=Request.QueryString("ClassID")%>&NewsTyp=hot&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">热点 
          </a> | <a href="Class_Rss.asp?ClassID=<%=Request.QueryString("ClassID")%>&NewsTyp=pic&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">图片 
          </a> | <a href="Class_Rss.asp?ClassID=<%=Request.QueryString("ClassID")%>&NewsTyp=highlight&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">精彩 
          </a> | <a href="Class_Rss.asp?ClassID=<%=Request.QueryString("ClassID")%>&NewsTyp=bignews&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">头条 
          </a> | <a href="Class_Rss.asp?ClassID=<%=Request.QueryString("ClassID")%>&NewsTyp=filt&Keyword=<%=Request("keyword")%>&ktype=<%=Request("ktype")%>">幻灯片</a>　　</div></td>
    </form>
  </tr>
</table>
  <%
	  if Not obj_news_rs.eof then
		Response.Write("<table width=""98%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"" class=""table""> <tr class=""hback""><td>")
		Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" >")
		Response.Write("<tr>")
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
			Else
				isUrlStr = ""
				str_Href = "<a href=""Class_Rss.asp?ClassID="&obj_news_rs("ClassID")&"""><img src=""../Images/rss.gif"" border=""0"" alt=""查看RSS""></a>"
				str_Href_title = "<a href=""Class_Rss.asp?ClassID="& obj_news_rs("ClassID") &""" title=""点击进入下一级栏目"">"& obj_news_rs("ClassName") &"</a>"
			End if
			Set obj_news_rs_1 = server.CreateObject(G_FS_RS)
			obj_news_rs_1.Open "Select Count(ID) from FS_DS_Class where ParentID='"& obj_news_rs("ClassID") &"'",Conn,1,1
			if obj_news_rs_1(0)>0 then
				str_action=  "<img src=""images/+.gif""></img>"& str_Href_title &""
			Else
				str_action=  "<img src=""images/-.gif""></img>"& str_Href_title &""
			End if
			obj_news_rs_1.close:set obj_news_rs_1 =nothing
			'得到下载数量
			if obj_news_rs("IsUrl") = 0 then
				Set obj_cnews_rs = server.CreateObject(G_FS_RS)
				obj_cnews_rs.Open "Select ID from FS_DS_List where ClassID='"& obj_news_rs("ClassID") &"'",Conn,1,1
				news_count = "("&obj_cnews_rs.recordcount&"/"&fs_news.GetTodayNewsCount(obj_news_rs("ClassID"))
				obj_cnews_rs.close:set obj_cnews_rs = nothing
			Else
				news_count = ""
			End if
			Response.Write"<td height=""22"">"
			Response.Write str_action&isUrlStr&news_count&str_Href
			Response.Write "</td>"
			obj_news_rs.MoveNext
			icNum = icNum + 1
			if icNum mod 4 = 0 then
				Response.Write("</tr><tr>")
			End if
		loop
		Response.Write("</tr></table></td></tr></table>")
	End if
%>

  
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="table">
  <form name="form2" method="post" action="">
    <tr> 
      <td><div align="center"> 
          <%
			Dim Rss_List,sCrLf
			Rss_List = ""
			sCrLf = chr(13) & chr(10)
			Rss_List = Rss_List &"gb2312"
			Rss_List = Rss_List &"text/xml"
			Rss_List = "<?xml version='1.0' encoding='gb2312'?>" & sCrLf
			Rss_List = Rss_List & "<rss version='2.0'>"&sCrLf
			Rss_List = Rss_List &  "<channel>"&sCrLf
			Rss_List = Rss_List &  "	<title>风讯</title>"&sCrLf
			Rss_List = Rss_List &  "	<description>风讯公司</description>"&sCrLf
			Rss_List = Rss_List &  "	<link>http://www.foosun.cn</link>"&sCrLf
			Rss_List = Rss_List &  "	<language>zh-cn</language>"&sCrLf
			Rss_List = Rss_List &  "	<docs>foosun Article Center</docs>"&sCrLf
			Rss_List = Rss_List &  "	<generator>Rss Generator By Foosun Inc.</generator>"&sCrLf
			Rss_List = Rss_List &  "	<image>"&sCrLf
			Rss_List = Rss_List &  "		<title>风讯</title>"&sCrLf
			Rss_List = Rss_List &  "		<link>http://www.foosun.cn</link>"&sCrLf
			Rss_List = Rss_List &  "		<url>http://vutoo.com/html/images/grzz.gif</url>"&sCrLf
			Rss_List = Rss_List &  "	</image>"&sCrLf
			Call GetFunctionstr
					if Request("NewsTyp") = "recTF" Then:str_Rec=" and "& CharIndexStr &"(NewsProperty,1,1)='1'":Else:str_Rec="":End if
					if Request("NewsTyp") = "isTop" Then:str_isTop=" and PopID=4 or PoPID=5":Else:str_isTop="":End if
					if Request("NewsTyp") = "hot" Then:str_hot=" and "& CharIndexStr &"(NewsProperty,13,1)='1'":Else:str_hot="":End if
					if Request("NewsTyp") = "pic" Then:str_pic=" and  isPicNews=1":Else:str_pic="":End if
					if Request("NewsTyp") = "highlight" Then:str_highlight=" and "& CharIndexStr &"(NewsProperty,15,1)='1'":Else:str_highlight="":End if
					if Request("NewsTyp") = "bignews" Then:str_bignews="  and "& CharIndexStr &"(NewsProperty,11,1)='1'":Else:str_bignews="":End if
					if Request("NewsTyp") = "filt" Then:str_filt=" and "& CharIndexStr &"(NewsProperty,21,1)='1'":Else:str_filt="":End if
					if Request("NewsTyp") = "Constr" Then:str_Constr=" and "& CharIndexStr &"(NewsProperty,7,1)='1'":Else:str_Constr="":End if
					if Trim(Request("Editor")) <>"" then:str_Editor = " and Editor = '"& Request("Editor")&"'":Else:str_Editor = "":End if
					if str_ClassID<>"" and len(str_ClassID)=15 then str_ClassID_1 = " and ClassID='"& str_ClassID &"'":Else:str_ClassID_1 = "":End if
					if Request("isCheck") = "1" then
						str_check = " and islock=0"
					elseif Request("isCheck") = "0" then
						str_check = " and islock=1"
					Else
						str_Check = ""
					End if
					newslist_sql = "Select top 50 ID,NewsID,PopID,ClassID,Content,NewsTitle,IsURL,isPicNews,URLAddress,Editor,Hits,NewsProperty,isLock,isRecyle,addtime,author,source from FS_NS_News where isRecyle=0 and isdraft=0 "& str_Editor & str_Rec & str_isTop & str_hot & str_pic & str_highlight & str_bignews & str_filt & str_Constr & str_ClassID_1 & str_check  &" Order by PopID desc,addtime desc,ID desc"
					Set obj_newslist_rs = Server.CreateObject(G_FS_RS)
					obj_newslist_rs.Open newslist_sql,Conn,1,3
					if not obj_newslist_rs.eof then
						do while not obj_newslist_rs.eof 
						Rss_List = Rss_List & "<item>"&sCrLf 
						Rss_List = Rss_List & "	<title>"& obj_newslist_rs("NewsTitle") &"</title>"&sCrLf 
						if obj_newslist_rs("isUrl")=1 then
							Rss_List = Rss_List & "	<link>"&obj_newslist_rs("URLAddress")&"</link>"&sCrLf 
							'Rss_List = Rss_List & "<description><![CDATA[标题下载,无内容]></description>"&sCrLf 
						Else
							Rss_List = Rss_List & "	<link>""1.html"&"</link>"&sCrLf 
							Rss_List = Rss_List & "	<description><![CDATA["&obj_newslist_rs("Content")&"]]></description>"&sCrLf 
						end if
						if len(trim(obj_newslist_rs("author")))=0 then
							Rss_List = Rss_List & "	<author>"&obj_newslist_rs("Source")&"</author>"&sCrLf 
						Else
							Rss_List = Rss_List &  "	<author>"&obj_newslist_rs("author")&"</author>"&sCrLf 
						End if
						Rss_List = Rss_List &  "	<pubDate>"&obj_newslist_rs("AddTime")&"</pubDate>"&sCrLf 
						Rss_List = Rss_List & "</item>"&sCrLf&sCrLf
							obj_newslist_rs.movenext
						Loop
						obj_newslist_rs.close:set obj_newslist_rs = nothing
					else
						Rss_List = Rss_List &""
					End if
			Rss_List = Rss_List & "</channel>"
			Rss_List = Rss_List & "</rss>"
		  %>
          <textarea name="RssShow" rows="30" style="width:100%"><% = Rss_List%></textarea>
        </div></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
set obj_newslist_rs = nothing
obj_news_rs.close
set obj_news_rs =nothing
set Fs_news = nothing
%>
<script language="JavaScript" type="text/JavaScript">

</script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






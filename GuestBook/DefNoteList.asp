<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp"-->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_InterFace/ns_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
MF_Default_Conn
Dim Conn,User_Conn
Dim Configobj,Topic,isUser,PageS,Style,sql,MSTitle
Set Configobj= server.CreateObject(G_FS_RS)
sql="select ID,Title,IsUser,IsAut,PageSize,Style From FS_WS_Config"
configobj.open sql,Conn,1,1
if not configobj.eof then
	Topic=configobj("Title")
	PageS=configobj("PageSize")
	IsUser=configobj("IsUser")
	MSTitle=configobj("Title")
	Style = configobj("Style")
	if Style<>"" then
		Style = Style
	else
		Style = "3"
	end if
end if
response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = Style
set configobj=nothing
Dim int_Start,int_RPP,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=PageS '设置每页显示数目
toF_="<font face=webdings>9</font>"   			'首页 
str_nonLinkColor_="#999999" '非热链接颜色 
'int_RPP = 30
int_showNumberLink_=10 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
toF_="<font face=webdings>9</font>"   			'首页 
toP10_=" <font face=webdings>7</font>"			'上十
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"
%>
<html>
<HEAD>
<TITLE><%=GetGuestBookTitle%></TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../<% = G_USER_DIR %>/images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="javascript">
function ShowNote(NoteID,ClassName,ClassID)
{
location="ShowNote.asp?NoteID="+NoteID+"&ClassName="+ClassName+"&ClassID="+ClassID;
}
</script>
<% 
dim ClassID,ClassRs,NoteSql,NoteSqlEnd,SelectClassID,NoteRs,NoteAct
if NoSqlHack(request.QueryString("ClassID"))<>"" then
	ClassID=NoSqlHack(trim(request.QueryString("ClassID")))
	Set ClassRs=Server.CreateObject(G_FS_RS)
	ClassRs.open "Select ID,ClassID,ClassName,ClassExp,Pid,Author from FS_WS_Class Where ClassID='"&NoSqlHack(ClassID)&"'",Conn,1,1
%>
<body>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
<%
	if not ClassRs.eof then
%>
  <tr>
   <td height="36" align="left" class="xingmu"><img src="images/Forum_nav.gif"> <a href="index.asp" class="Top_Navi"><strong><%=MSTitle%></strong></a>-><%=ClassRs("ClassName")%>->帖子列表</td>
  </tr>
  <tr>
    <td class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="53%" height="28"><strong>&nbsp;·<%=ClassRs("ClassExp")%></strong></td>
        <td width="47%"><div align="center">
    <%
	if IsUser="0" then 
		response.Write("<a href=AddnewNotes.asp?ClassID="&ClassRs("ClassID")&"&ClassName="&ClassRs("ClassName")&">发表话题</a>｜")
	elseif	session("FS_UserName")<>"" then
		response.Write("<a href=AddnewNotes.asp?ClassID="&ClassRs("ClassID")&"&ClassName="&ClassRs("ClassName")&">发表话题</a>｜")
	end if
	%>
            <a href="?Act=all&ClassID=<%=ClassRs("ClassID")%>">所有贴子</a>｜<a href="?Act=Before&ClassID=<%=ClassRs("ClassID")%>">推荐帖子</a>｜<a href="?Act=Person&ClassID=<%=ClassRs("ClassID")%>">人气贴子</a></div></td>
      </tr>
    </table></td>
  </tr>
  <tr id="Note" style="display:">
    <td >

</table>
	<%
	NoteSql="Select ID,ClassID,[User],Topic,Body,AddDate,IsTop,State,Style,IsAdmin,Answer,Hit,LastUpdateDate,LastUpdateUser,Face from FS_WS_BBS"
	NoteSqlEnd=" order by IsTop DESC,LastUpdateDate desc,AddDate DESC,id desc"
	if NoSqlHack(Request.queryString("Act"))<>"" then
		NoteAct=NoSqlHack(Request.queryString("Act"))
		SelectClassID=NoSqlHack(Request.querystring("ClassID"))
		Select Case NoteAct
		Case "All"
			if SelectClassID<>"" then
				NoteSql=NoteSql&" Where State='0' and IsAdmin<>'1' and ClassID='"&NoSqlHack(SelectClassID)&"' and ClassID='"&ClassRs("ClassID")&"' "&NoteSqlEnd
			else
				NoteSql=NoteSql&" Where State='0' and IsAdmin<>'1' and ClassID='"&NoSqlHack(ClassRs("ClassID"))&"' "&NoteSqlEnd
			end if
		Case "Adime"
			if SelectClassID<>"" then
				NoteSql=NoteSql&" Where State='0' and IsAdmin<>'1' and ClassID='"&NoSqlHack(SelectClassID)&"' and  ClassID='"&ClassRs("ClassID")&"' and IsAdmin='1' and ParentID='0'  order by IsTop DESC,AddDate DESC"
			else
				NoteSql=NoteSql&" Where State='0' and IsAdmin<>'1'  and  ClassID='"&NoSqlHack(ClassRs("ClassID"))&"' and IsAdmin='1' and ParentID='0'  order by IsTop DESC,AddDate DESC"
			end if
		Case "Y"
			if SelectClassID<>"" then
				NoteSql=NoteSql&" Where State='0' and IsAdmin<>'1' and ClassID='"&NoSqlHack(SelectClassID)&"' and ClassID='"&ClassRs("ClassID")&"'and ClassID='"&NoSqlHack(SelectClassID)&"' and ParentID='0'  " &NoteSqlEnd
			else
				NoteSql=NoteSql&" Where State='0' and IsAdmin<>'1'  and ClassID='"&NoSqlHack(ClassRs("ClassID"))&"' and ClassID='"&SelectClassID&"' and ParentID='0'  " &NoteSqlEnd
			end if
		Case "N"
			if SelectClassID<>"" then
				NoteSql=NoteSql&" Where State='0' and IsAdmin<>'1' and ClassID='"&NoSqlHack(SelectClassID)&"' and ClassID='"&ClassRs("ClassID")&"' and State='1' and ParentID='0' " &NoteSqlEnd
			else
				NoteSql=NoteSql&" Where State='0' and IsAdmin<>'1' and  ClassID='"&NoSqlHack(ClassRs("ClassID"))&"' and State='1' and ParentID='0' " &NoteSqlEnd
			end if
		Case "Before"	
			if SelectClassID<>"" then
				NoteSql=NoteSql&" where State='0' and IsAdmin<>'1' and ClassID='"&NoSqlHack(SelectClassID)&"' and ClassID='"&ClassRs("ClassID")&"' and IsTop='1' and ParentID='0' "  &NoteSqlEnd
			else
				NoteSql=NoteSql&" where State='0' and IsAdmin<>'1' and ClassID='"&NoSqlHack(ClassRs("ClassID"))&"' and IsTop='1' and ParentID='0' "  &NoteSqlEnd
			end if
		Case "Person"
			if SelectClassID<>"" then
				NoteSql=NoteSql&" where State='0' and IsAdmin<>'1' and ClassID='"&NoSqlHack(SelectClassID)&"' and ClassID='"&ClassRs("ClassID")&"' and ParentID='0'  order by IsTop DESC,Hit desc,LastUpdateDate DESC,AddDate DESC,id desc"
			else
				NoteSql=NoteSql&" where State='0' and IsAdmin<>'1' and  ClassID='"&NoSqlHack(ClassRs("ClassID"))&"' and ParentID='0'  order by IsTop DESC,Hit DESC,LastUpdateDate desc,AddDate DESC,id desc"
			end if
		Case Else
			if SelectClassID<>"" then
				NoteSql=NoteSql&" Where State='0' and IsAdmin<>'1' and ClassID='"&NoSqlHack(SelectClassID)&"' and ClassID='"&ClassRs("ClassID")&"' and ParentID='0' "&NoteSqlEnd
			else
				NoteSql=NoteSql&" Where State='0' and IsAdmin<>'1' and ClassID='"&NoSqlHack(ClassRs("ClassID"))&"' and ParentID='0' "&NoteSqlEnd
			end if
		End Select
	else
		if SelectClassID<>"" then
			NoteSql=NoteSql&" Where State='0' and IsAdmin<>'1' and ClassID='"&NoSqlHack(SelectClassID)&"' and ClassID='"&ClassRs("ClassID")&"' and ParentID='0'  "&NoteSqlEnd
		else
			NoteSql=NoteSql&" Where State='0' and IsAdmin<>'1' and ClassID='"&NoSqlHack(ClassRs("ClassID"))&"' and ParentID='0'  "&NoteSqlEnd
		end if
	end if
	Set NoteRs=Server.CreateObject(G_FS_RS)
	NoteRs.open NoteSql,Conn,1,1
	%>
	
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
       <form method="post" action="NoteDel.asp?Act=del" id="<%=ClassRs("ClassID")%>" name="<%=ClassRs("ClassID")%>">
	   <%
	   if not NoteRs.eof then	
		   NoteRs.PageSize=int_RPP
			cPageNo=NoSqlHack(Request.QueryString("Page"))
			If cPageNo="" Then 
				cPageNo = 1
			End if
			If not isnumeric(cPageNo) Then 
				cPageNo = 1
			End If
			cPageNo = Clng(cPageNo)
			If cPageNo<=0 Then 
				cPageNo=1
			End If
			If cPageNo>NoteRs.PageCount Then 
				cPageNo=NoteRs.PageCount 
			End IF
			NoteRs.AbsolutePage=cPageNo		
			%>
		<tr class="hback_1">
		 <td width="8%"><div align="center"><strong>状态</strong></div></td>		
		 <td width="39%" ><div align="left"><strong>贴子标题</strong></div></td>
		 <td width="14%" ><div align="center"><strong>作 者</strong></div></td>
		 <td width="10%"  ><div align="center"><strong>回 复</strong></div></td>
		 <td width="10%"  ><div align="center"><strong>人 气</strong></div></td>
		 <td width="19%"  ><div align="center"><strong>最后更新</strong></div></td>
		<%
			FOR int_Start=1 TO int_RPP 	
		 %>
	    <tr class="hback">
		<td height="31">
		  <div align="center">
		    <%
		  	if NoteRs("State")="1" then
			 	Response.write("<img src=""images\lock.gif"" alt=""被锁帖子"">")
			elseif NoteRs("IsAdmin")="1" then
			 		Response.write("<img src=""images\Admin.gif"" alt=""管理员可见帖子"">")
			elseif NoteRs("IsTop")="1" then
			 		Response.write("<img src=""images\top.gif"" alt=""推荐帖子"">")
			else				
			 	Response.write("<img src=""images\gogo.gif"" alt=""普通帖子"">")
			end if
		  %>		
	      </div></td>
		  <td ><a href="#" onClick="ShowNote(<%=NoteRs("ID")%>,'<%=ClassRs("ClassName")%>','<%=ClassRs("ClassID")%>')"><%=left(NoteRs("Topic"),30)%></a></td>
		  <%if NoteRs("User")<>"游客" then%>
		  <td class="tdhback" ><div align="center"><a href="../<%=G_USER_DIR%>/ShowUser.asp?UserName=<%=NoteRs("User")%>" target="_blank"><%=NoteRs("User")%></a></div></td>
		  <%else%>
		  <td class="tdhback" ><div align="center"><%=NoteRs("User")%></div></td>
		  <%end if%>
		  <td><div align="center"><%=NoteRs("Answer")%></div></td>
		  <td><div align="center"><%=NoteRs("Hit")%></div></td>
		  <td class="tdhback"	><div align="right"><%=NoteRS("LastUpdateDate")%>｜<%=NoteRS("LastUpdateUser")%></div></td>
        </tr>
		<%
		 NoteRs.MoveNext
		 if NoteRs.eof or NoteRs.bof then exit for
     	 NEXT		   	  
	    %>
	<%
	 Response.Write("<tr><td class=""hback_1"" colspan=""9"" align=""right"">"&fPageCount(NoteRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo) &"</td></tr>")

	else
	%></form>	 
</table>
  <%
		response.Write("<div align=center>没有帖</div>")
	END IF
	Set NoteRs=nothing
  end if
end if	
Set Conn=nothing
%>
</body>
</html>







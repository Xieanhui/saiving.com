
<%
'定义分页函数的参数
Dim int_RPP,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=15 '设置每页显示数目
toF_="<font face=webdings>9</font>"   			'首页 
str_nonLinkColor_="#999999" '非热链接颜色
int_showNumberLink_=10 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
toF_="<font face=webdings>9</font>"   			'首页 
toP10_=" <font face=webdings>7</font>"			'上十
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"
%>
		<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="table" >
          <%
  Dim obj_news_rs,obj_news_sql
  KeyWord=request.Form("SearchKey")
  	obj_news_sql="Select ID,ClassID,ClassName from FS_NS_NewsClass where Parentid  = '0' and ReycleTF<>1 order by OrderID desc"
	Set obj_news_rs = server.CreateObject(G_FS_RS)
	obj_news_rs.Open obj_news_sql,Conn,1,1
  %>
  	<form name='SearchForm' method="get" target='_self' action="?">
            <tr height="21" bgcolor="#FFFFFF">
              <td width="11%" align="center" class="hback" ><input type="checkbox" name="UnAll" value="UnNews" checked="checked" />
                不规则新闻 </td>
              <td width="11%" height="35" align="center" class="hback" >栏目列表</td>
              <td align="center" width="40%" class="hback" ><select name="ClassID" onChange="listnews(this,'<%=NewsID%>')" style="width:100%">
                  <option value="" selected>所有栏目</option>
                  <%	
		  if Request.QueryString("ClassID")<>"" then
				dim RecentRS
				Set RecentRS= server.CreateObject(G_FS_RS)
				RecentRS.open "Select ID,ClassID,ClassName from FS_NS_NewsClass where ClassID='"&NoSqlHack(Request.QueryString("ClassID"))&"'and ReycleTF<>1",Conn,1,1
				if not RecentRS.eof then
				%>
                  <option value="<%=RecentRS("ClassID")%>" selected><%=RecentRS("ClassName")%></option>
                  <%				
				end if
				Set RecentRS=nothing
		 end if		
		 do while not obj_news_rs.eof
		  %>
                  <option value="<%=obj_news_rs("ClassID")%>"><%=obj_news_rs("ClassName")%></option>
                  <% 
		 Response.write(GetChildNewsList(obj_news_rs("ID"),1))
		obj_news_rs.MoveNext
		loop
		Set obj_news_rs=nothing
		%>
              </select></td>
              <td align="center" width="40%" class="hback" ><input type="text" name="SearchKey" value="<%=KeyWord%>">
                  <input name="submit2" type="button" value="搜 索" onclick="listSearch('<%=Request.QueryString("NewsID")%>')">              </td>
            </tr>
          </form>
		  </table>
		<form name="selectNews">
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table"  >
    <tr class="xingmu">
    <td width="9%" align=center>副新闻</td>
    <td width="46%" align="center">输入不规则标题</td>
    <td width="45%" align="center">新闻标题</td>
  </tr>
  <%
Dim temp
Dim IDList : IDList = "," '列出不重复的新闻ID
Dim ListObj,ListSql
Dim TempClassID,KeyWord
TempClassID = Cstr(Request.QueryString("ClassID"))
KeyWord = NoSqlHack(NoCSSHackAdmin(Request.QueryString("SearchKey"),"关键字"))
set ListObj = server.CreateObject(G_FS_RS)
ListSQL="select ID,NewsID,ClassID,NewsTitle,CurtTitle From FS_NS_News where isLock<>1 and isRecyle=0  and Mid(NewsProperty,17,1)=1 Order By PopId Desc "
'全部新闻	
If Request("UnAll")="" then
	ListSQL="select ID,NewsID,ClassID,NewsTitle,CurtTitle From FS_NS_News where isLock<>1 and isRecyle=0 Order By PopId Desc "
	If KeyWord<>"" Then
		If TempClassID<>"" Then
			ListSQL = "select ID,NewsID,ClassID,NewsTitle,CurtTitle From FS_NS_News  where (KeyWords like '%"&KeyWord&"%' or NewsTitle like '%"&KeyWord&"%') and ClassID='"&TempClassID&"'and isLock<>1 and isRecyle=0  Order By PopId Desc"
		Else
			ListSQL = "select ID,NewsID,ClassID,NewsTitle,CurtTitle From FS_NS_News where (KeyWords like '%"&KeyWord&"%' or NewsTitle like '%"&KeyWord&"%') and isLock<>1 and isRecyle=0  Order By PopId Desc"
		End If
	else if TempClassID<>"" then
			ListSQL = "select ID,NewsID,ClassID,NewsTitle,CurtTitle From FS_NS_News  where  ClassID='"&TempClassID&"'and isLock<>1 and isRecyle=0 Order By PopId Desc"
		 end if	
	End IF

Else
	ListSQL="select ID,NewsID,ClassID,NewsTitle,CurtTitle From FS_NS_News where isLock<>1 and isRecyle=0  and Mid(NewsProperty,17,1)=1 Order By PopId Desc "
	If KeyWord<>"" Then
		If TempClassID<>"" Then
			ListSQL = "select ID,NewsID,ClassID,NewsTitle,CurtTitle From FS_NS_News  where (KeyWords like '%"&KeyWord&"%' or NewsTitle like '%"&KeyWord&"%') and ClassID='"&TempClassID&"'and isLock<>1 and isRecyle=0 and Mid(NewsProperty,17,1)=1 Order By PopId Desc"
		Else
			ListSQL = "select ID,NewsID,ClassID,NewsTitle,CurtTitle From FS_NS_News where (KeyWords like '%"&KeyWord&"%' or NewsTitle like '%"&KeyWord&"%') and isLock<>1 and isRecyle=0 and Mid(NewsProperty,17,1)=1 Order By PopId Desc"
		End If
	else if TempClassID<>"" then
			ListSQL = "select ID,NewsID,ClassID,NewsTitle,CurtTitle From FS_NS_News  where  ClassID='"&TempClassID&"'and isLock<>1 and isRecyle=0 and Mid(NewsProperty,17,1)=1 Order By PopId Desc"
		 end if	
	End IF
ENd IF
	ListObj.open ListSQL,Conn,1,1
	if Not ListObj.eof Then
		ListObj.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("Page"))
		If cPageNo="" Then 
			cPageNo = 1
		End if
		If not isnumeric(cPageNo) Then 
			cPageNo = 1
			cPageNo = Clng(cPageNo)
		End If
		If cPageNo<=0 Then 
			cPageNo=1		
		End If
		If cPageNo>ListObj.PageCount Then 
			cPageNo=ListObj.PageCount 
			ListObj.AbsolutePage=cPageNo
		End IF
		do while not ListObj.eof
			If Instr(IDList,","&ListObj("ID")&",")=0 Then
				'Response.write "<tr><td colspan=5 class=""hback""></td></tr>"
				if trim(ListObj("CurtTitle"))<>"" then
					temp=ListObj("CurtTitle")
				else
					temp=ListObj("NewsTitle")
				end if
				Response.write "<tr><td align=""center"" class=""hback""><input type='checkbox' name='NewsID' value="&ListObj("NewsID")&":"&temp&" onpropertychange=""addrow(this,shebei,'"&ListObj("NewsID")&"','"&temp&"','"&ListObj("NewsTitle")&"')""></td><td align=center class=""hback""><input title='修改不规则调用时使用的标题' type='text' name='File_"&ListObj("NewsID")&"' Style='width:100%' value='"&temp&"'></td><td class=""hback""><a href='"&(ListObj("NewsID"))&"' target=_blank title=点击查看本条新闻>"&ListObj("NewsTitle")&"</a></td></tr>"
			End If
			IDList = IDList & ListObj("ID") & ","
		ListObj.movenext
		Loop
		Response.Write("<tr><td class=""xingmu"" colspan=""4"" align=""right"">"&fPageCount(ListObj,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf&"</td></tr>")

	Else
		Response.write "<tr><td colspan=4 class=""hback""></td></tr><tr><td colspan=5 height=23>没有与关键字相关的新闻</td></tr>"
	End If
	Response.write "<tr><td colspan=4 class=""hback""></td></tr>"
	ListObj.Close	
	Set ListObj=nothing
%>
</table>
</form>	






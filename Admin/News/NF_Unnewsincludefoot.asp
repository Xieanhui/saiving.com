
<%
'�����ҳ�����Ĳ���
Dim int_RPP,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=15 '����ÿҳ��ʾ��Ŀ
toF_="<font face=webdings>9</font>"   			'��ҳ 
str_nonLinkColor_="#999999" '����������ɫ
int_showNumberLink_=10 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
toF_="<font face=webdings>9</font>"   			'��ҳ 
toP10_=" <font face=webdings>7</font>"			'��ʮ
toP1_=" <font face=webdings>3</font>"			'��һ
toN1_=" <font face=webdings>4</font>"			'��һ
toN10_=" <font face=webdings>8</font>"			'��ʮ
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
                ���������� </td>
              <td width="11%" height="35" align="center" class="hback" >��Ŀ�б�</td>
              <td align="center" width="40%" class="hback" ><select name="ClassID" onChange="listnews(this,'<%=NewsID%>')" style="width:100%">
                  <option value="" selected>������Ŀ</option>
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
                  <input name="submit2" type="button" value="�� ��" onclick="listSearch('<%=Request.QueryString("NewsID")%>')">              </td>
            </tr>
          </form>
		  </table>
		<form name="selectNews">
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table"  >
    <tr class="xingmu">
    <td width="9%" align=center>������</td>
    <td width="46%" align="center">���벻�������</td>
    <td width="45%" align="center">���ű���</td>
  </tr>
  <%
Dim temp
Dim IDList : IDList = "," '�г����ظ�������ID
Dim ListObj,ListSql
Dim TempClassID,KeyWord
TempClassID = Cstr(Request.QueryString("ClassID"))
KeyWord = NoSqlHack(NoCSSHackAdmin(Request.QueryString("SearchKey"),"�ؼ���"))
set ListObj = server.CreateObject(G_FS_RS)
ListSQL="select ID,NewsID,ClassID,NewsTitle,CurtTitle From FS_NS_News where isLock<>1 and isRecyle=0  and Mid(NewsProperty,17,1)=1 Order By PopId Desc "
'ȫ������	
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
				Response.write "<tr><td align=""center"" class=""hback""><input type='checkbox' name='NewsID' value="&ListObj("NewsID")&":"&temp&" onpropertychange=""addrow(this,shebei,'"&ListObj("NewsID")&"','"&temp&"','"&ListObj("NewsTitle")&"')""></td><td align=center class=""hback""><input title='�޸Ĳ��������ʱʹ�õı���' type='text' name='File_"&ListObj("NewsID")&"' Style='width:100%' value='"&temp&"'></td><td class=""hback""><a href='"&(ListObj("NewsID"))&"' target=_blank title=����鿴��������>"&ListObj("NewsTitle")&"</a></td></tr>"
			End If
			IDList = IDList & ListObj("ID") & ","
		ListObj.movenext
		Loop
		Response.Write("<tr><td class=""xingmu"" colspan=""4"" align=""right"">"&fPageCount(ListObj,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf&"</td></tr>")

	Else
		Response.write "<tr><td colspan=4 class=""hback""></td></tr><tr><td colspan=5 height=23>û����ؼ�����ص�����</td></tr>"
	End If
	Response.write "<tr><td colspan=4 class=""hback""></td></tr>"
	ListObj.Close	
	Set ListObj=nothing
%>
</table>
</form>	






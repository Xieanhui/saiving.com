<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="NF_News_Function.asp"-->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_User_Conn
'session判断
MF_Session_TF
'权限判断
'Call MF_Check_Pop_TF("NS_Class_000001")
'得到会员组列表
dim Fs_news
set Fs_news = new Cls_News

Dim CharIndexStr
CharIndexStr=all_substring
%>
<html>
<head>
<title>不规则新闻</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="js/Public.js"></script>
<script language="javascript" src="../../Fs_inc/CheckJs.js"></script>
<script language="javascript" src="../../FS_INC/prototype.js"></script>
</head>
<body onselectstart="return false;">
<%
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=10 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings>9</font>"  			'首页 
toP10_=" <font face=webdings>7</font>"			'上十
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"				'尾页
%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
		<td>
			<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
				<tr class="xingmu">
					<td width="50%" align="center" class="xingmu">新闻标题</td>
					<td width="20%" align="center" class="xingmu">新闻栏目</td>
					<td width="*%" align=center class="xingmu">加入不规则</td>
				</tr>
				<%
Dim temp,UnAll
Dim IDList : IDList = "," '列出不重复的新闻ID
Dim ListObj,ListSql
Dim TempClassID,KeyWord
TempClassID = NoSqlHack(Cstr(Request("ClassID")))
KeyWord = NoCSSHackAdmin(Request("SearchKey"),"关键字")
KeyWord = Replace(KeyWord,",","%")
KeyWord = Replace(KeyWord," ","%")
KeyWord = Replace(KeyWord,";","%")
KeyWord = NoSqlHack(KeyWord)
set ListObj = server.CreateObject(G_FS_RS)

'全部新闻
Dim StrSearch
UnAll = Request("UnAll")
StrSearch = ""
If UnAll<>"" then
	StrSearch=StrSearch & " and "&CharIndexStr&"(NewsProperty,17,1)=1"
End If
If KeyWord<>"" Then
	StrSearch=StrSearch & " and (KeyWords like '%"&KeyWord&"%' or NewsTitle like '%"&KeyWord&"%')"
End If
If TempClassID<>"" Then
	StrSearch=StrSearch & " and FS_NS_News.ClassID='"&TempClassID&"'"
End If

ListSQL="select FS_NS_News.ID,NewsID,FS_NS_NewsClass.ClassID,FS_NS_NewsClass.ClassName,NewsTitle,CurtTitle From FS_NS_News,FS_NS_NewsClass where FS_NS_News.ClassID=FS_NS_NewsClass.ClassID and isLock<>1 and isRecyle=0"&StrSearch&" Order By PopId Desc,FS_NS_News.AddTime Desc"
ListObj.open ListSQL,Conn,1,1
if Not ListObj.eof Then		
	ListObj.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>ListObj.PageCount Then cPageNo=ListObj.PageCount 
	ListObj.AbsolutePage=cPageNo
	FOR int_Start=1 TO int_RPP 
		If Instr(IDList,","&ListObj("ID")&",")=0 Then
			'Response.write "<tr><td colspan=5 class=""hback""></td></tr>"
			if trim(ListObj("CurtTitle"))<>"" then
				temp=ListObj("CurtTitle")
			else
				temp=ListObj("NewsTitle")
			end if
			Response.write "<tr><td class=""hback""><a href='News_Review.asp?NewsID="&(ListObj("NewsID"))&"&ClassID="&(ListObj("ClassID"))&"' title=点击查看本条新闻 target=""_blank"">"&left(ListObj("NewsTitle"),36)&"</a></td><td align=center class=""hback"">"&ListObj("ClassName")&"</td><td align=center class=""hback""><button id=""New"&ListObj("NewsID")&""" onclick=""AddUnNewList('"&ListObj("NewsID")&"','"&Replace(Replace(ListObj("NewsTitle"),"""",""),"'","")&"');"">加入不规则新闻</button></td></tr>"&vbcrlf
		End If
		IDList = IDList & ListObj("ID") & ","
		ListObj.MoveNext
	If ListObj.eof or ListObj.bof then exit for
  NEXT
Else
	Response.write "<tr><td colspan=""3"" class=""hback""></td></tr><tr><td colspan=""3"" height=""23"">没有与关键字相关的新闻</td></tr>"
End If
	Response.Write("<tr><td class=""hback"" colspan=""3"" align=""right"">"&fPageCount(ListObj,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf&"</td></tr>")
ListObj.Close	
Set ListObj=nothing
%>
			</table>
		</td>
	</tr>
</table>
<script language="javascript">
<!--

function AddUnNewList(NewId,NewTitle){
	var ListLen=parent.UnNewArray.length;
	var FindFlag=-1;
	for (var i=0;i<ListLen;i++){
		if (parent.UnNewArray[i][0]==NewId){
			FindFlag=i;
			break;
		}
	}
	if (FindFlag>-1){
		if (confirm("确定移除吗？")){
			parent.UnNewArray.remove(FindFlag,1);
			parent.DisplayUnNews();
			CheckUnNews();
			parent.UnNewPreviewCh();
		}
	}else{
		parent.UnNewArray[ListLen]=[NewId,NewTitle,NewTitle,(ListLen+1)];
		parent.DisplayUnNews();
		CheckUnNews();
		parent.UnNewPreviewCh();
	}
}
function CheckUnNews(){
	var ListLen=parent.UnNewArray.length;
	var FindFlag=-1;
	var buttons=document.getElementsByTagName("button");
	for (var j=0;j<buttons.length;j++){
		FindFlag=-1
		for (var i=0;i<ListLen;i++){
			if (("New"+parent.UnNewArray[i][0])==buttons[j].id){
				FindFlag=j;
				break;
			}
		}
		if (FindFlag>-1){
			buttons[j].innerText="从不规则新闻中删除";
		}else{
			buttons[j].innerText="加入不规则新闻";
		}
	}
}
CheckUnNews();
-->
</script>
</body>
<%
Set Conn=nothing
%>
</html>







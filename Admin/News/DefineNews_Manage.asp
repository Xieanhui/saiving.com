<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
'session判断
MF_Session_TF 
if not MF_Check_Pop_TF("NS_UnRl") then Err_Show
dim Fs_news
set Fs_news = new Cls_News
%>
<html>
<head>
<title>不规则新闻 规则管理</title> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="js/Public.js"></script>
<script language="javascript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = MainForm.elements[i];  
    if (e.name != 'chkall')  
       e.checked = MainForm.chkall.checked;  
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
</head>
<body>
<%
Dim Str_PopID 
Str_PopID= "<IMG Src=""images/newstype/0.gif"" border=""0"" alt=""不规则新闻"& Fs_news.allInfotitle &",点击查看全部内容"">"
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu"><a href="#" onMouseOver="this.T_BGCOLOR='#404040';this.T_FONTCOLOR='#FFFFFF';return escape('<div align=\'center\'>FoosunCMS5.0<br> Code by 文永钟 <BR>Copyright (c) 2006 Foosun Inc</div>')" class="sd"><strong>不规则新闻管理</strong></a></td>
  </tr>
  <tr> 
    <td height="26" colspan="0" valign="middle" class="hback">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="5%" class="hback" align="center" vlign="middle" ><a href="UnRegulatenewAdd.asp">添加</a></td>
		  <td width="2%" class="Gray" align="center" vlign="middle">|</td>
		  <td width="5%" align="center" vlign="middle"class="hback"><a href="../../help?Lable=UnRegualNewAdd" target="_blank" style="cursor:help;">帮助</a></td>
		  <td width="644">&nbsp; </td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class="table">
  <form name="MainForm" method="post" action="UnRegulatenewDel.asp?Action=del">
  <tr class="hback"> 
	<td class="xingmu" colspan="3">
	  <div align="center" >主新闻</div></td>       
	<td width="27%" class="xingmu">
		<div align="center">操作</div></td>
	</tr>
 <%
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=15 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings>9</font>"  			'首页 
toP10_=" <font face=webdings>7</font>"			'上十
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"	
dim RsNewsObj,TmpRsObj
Set RsNewsObj = server.createobject(G_FS_RS)
Set TmpRsObj = server.createobject(G_FS_RS)
'取主新闻ID
Dim MainNewsID,TmpSql,i,Str,RowValue
RowValue=1
RsNewsObj.Open "Select DisTinct UnRegulatedMain From [FS_NS_News_Unrgl] order by UnRegulatedMain DESC ",Conn,1,1
IF not RsNewsObj.eof  then
	RsNewsObj.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>RsNewsObj.PageCount Then cPageNo=RsNewsObj.PageCount 
	RsNewsObj.AbsolutePage=cPageNo
FOR int_Start=1 TO int_RPP
TmpRsObj.open "Select UnregNewsName From FS_NS_News_Unrgl where UnregulatedMain='"&RsNewsObj("UnRegulatedMain")&"' order by [Rows],ID ASC",Conn,1,1
if not TmpRsObj.eof then
	Str=TmpRsObj(0)
else 
	Str="已被删除"
end if
TmpRsObj.close
%>
  <tr> 
      <td width="3%" class="hback" align="center"><div align="center"> 
          <input name="C_NewsID" type="checkbox" id="C_NewsID" value="<% = RsNewsObj("UnRegulatedMain")%>">
        </div></td>
      <td width="4%" height="18" align="center" class="hback" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(M_Newsid<% = RsNewsObj("UnRegulatedMain")%>);"  language=javascript><% = Str_PopID %></td>
      <td width="66%" height="18" class="hback" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(M_Newsid<% = RsNewsObj("UnRegulatedMain")%>);"  language=javascript><% = Str %></td>
     <td class="hback"><div align="center"><a href="UnRegulatenewAdd.asp?NewsID=<% = RsNewsObj("UnRegulatedMain")%>&Action=Edit">修改</a>｜<a href="UnRegulatenewDel.asp?NewsID=<% = RsNewsObj("UnRegulatedMain")%>&Action=signDel"  onClick="{if(confirm('确定要删除吗？\n\n如果你在系统参数设置中设置删除<% =  Fs_news.allInfotitle %>到回收站\n<% =  Fs_news.allInfotitle %>将删除到回收站中!\n必要时候可还原')){return true;}return false;}">删除</a></div></td>
   </tr>
	 <tr id="M_Newsid<% = RsNewsObj("UnRegulatedMain")%>" style="display:none">
      <td height="35" colspan="7" class="hback">       <table width="100%" border="0" cellspacing="1" cellpadding="2" class="table">
        <tr class="hback">
          <td width="45%" height="20" class="hback" colspan="4"><font style="font-size:12px">
            <%
				If Not IsNull(MainNewsID) then
					TmpRsObj.open "Select ID,UnregulatedMain,MainUnregNewsID,UnregNewsName,[Rows] from FS_NS_News_Unrgl Where UnregulatedMain='"&RsNewsObj("UnRegulatedMain")&"' order by [Rows] ASC,ID ASC",Conn,1,3
					Do While Not TmpRsObj.eof
						response.Write("<a href='#'>"&TmpRsObj("UnregNewsName")&"</a>")
						RowValue=Cint(TmpRsObj("Rows"))
						TmpRsObj.movenext
						if not TmpRsObj.eof then
							if Cint(TmpRsObj("Rows"))=RowValue then
								Response.Write("&nbsp;")
							else
								Response.Write("<br>")
							end if
						end if						
					Loop
					TmpRsObj.close
				End if
					%>
          </font></td>
        </tr> 
		 </table></td>
	</tr>	   
<%
		RsNewsObj.MoveNext
		if RsNewsObj.eof or RsNewsObj.bof then exit for
      NEXT
%>
<tr> 
      <td height="30" colspan="4" class="hback">
<div align="right"> 
          <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
          选择所有 
          <input name="Submit" type="submit" id="Submit"  onClick="{if(confirm('确实要进行删除吗?')){this.document.MainForm.submit();return true;}return false;}" value=" 删除 ">
        </div></td>
    </tr>
	<%
	 Response.Write("<tr><td class=""hback"" colspan=""4"" align=""right"">"&fPageCount(RsNewsObj,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf&"</td></tr>")
	%></form>	 
</table>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
<iframe name="tempFrame" height=0 width=0 frameborder=0 style="display:none"></iframe>
</body>
<% 
else
end if
Set RsNewsObj=nothing
Set Conn=nothing
%>
</html>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






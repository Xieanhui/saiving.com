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
'session�ж�
MF_Session_TF 
if not MF_Check_Pop_TF("NS_UnRl") then Err_Show
dim Fs_news
set Fs_news = new Cls_News
%>
<html>
<head>
<title>���������� �������</title> 
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
Str_PopID= "<IMG Src=""images/newstype/0.gif"" border=""0"" alt=""����������"& Fs_news.allInfotitle &",����鿴ȫ������"">"
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu"><a href="#" onMouseOver="this.T_BGCOLOR='#404040';this.T_FONTCOLOR='#FFFFFF';return escape('<div align=\'center\'>FoosunCMS5.0<br> Code by ������ <BR>Copyright (c) 2006 Foosun Inc</div>')" class="sd"><strong>���������Ź���</strong></a></td>
  </tr>
  <tr> 
    <td height="26" colspan="0" valign="middle" class="hback">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="5%" class="hback" align="center" vlign="middle" ><a href="UnRegulatenewAdd.asp">���</a></td>
		  <td width="2%" class="Gray" align="center" vlign="middle">|</td>
		  <td width="5%" align="center" vlign="middle"class="hback"><a href="../../help?Lable=UnRegualNewAdd" target="_blank" style="cursor:help;">����</a></td>
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
	  <div align="center" >������</div></td>       
	<td width="27%" class="xingmu">
		<div align="center">����</div></td>
	</tr>
 <%
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=15 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings>9</font>"  			'��ҳ 
toP10_=" <font face=webdings>7</font>"			'��ʮ
toP1_=" <font face=webdings>3</font>"			'��һ
toN1_=" <font face=webdings>4</font>"			'��һ
toN10_=" <font face=webdings>8</font>"			'��ʮ
toL_="<font face=webdings>:</font>"	
dim RsNewsObj,TmpRsObj
Set RsNewsObj = server.createobject(G_FS_RS)
Set TmpRsObj = server.createobject(G_FS_RS)
'ȡ������ID
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
	Str="�ѱ�ɾ��"
end if
TmpRsObj.close
%>
  <tr> 
      <td width="3%" class="hback" align="center"><div align="center"> 
          <input name="C_NewsID" type="checkbox" id="C_NewsID" value="<% = RsNewsObj("UnRegulatedMain")%>">
        </div></td>
      <td width="4%" height="18" align="center" class="hback" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(M_Newsid<% = RsNewsObj("UnRegulatedMain")%>);"  language=javascript><% = Str_PopID %></td>
      <td width="66%" height="18" class="hback" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(M_Newsid<% = RsNewsObj("UnRegulatedMain")%>);"  language=javascript><% = Str %></td>
     <td class="hback"><div align="center"><a href="UnRegulatenewAdd.asp?NewsID=<% = RsNewsObj("UnRegulatedMain")%>&Action=Edit">�޸�</a>��<a href="UnRegulatenewDel.asp?NewsID=<% = RsNewsObj("UnRegulatedMain")%>&Action=signDel"  onClick="{if(confirm('ȷ��Ҫɾ����\n\n�������ϵͳ��������������ɾ��<% =  Fs_news.allInfotitle %>������վ\n<% =  Fs_news.allInfotitle %>��ɾ��������վ��!\n��Ҫʱ��ɻ�ԭ')){return true;}return false;}">ɾ��</a></div></td>
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
          ѡ������ 
          <input name="Submit" type="submit" id="Submit"  onClick="{if(confirm('ȷʵҪ����ɾ����?')){this.document.MainForm.submit();return true;}return false;}" value=" ɾ�� ">
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
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






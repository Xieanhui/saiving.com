<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp" -->
<%
  dim rs,i,picNumber,Imglist,o_class
  if Request.QueryString("classid")<>"" then
  	o_class=" and Classid="&CintStr(Request.QueryString("classid"))&""
  else
  	o_class=""
  end if
  set rs = Server.CreateObject(G_FS_RS)
  rs.open "select top 100 id,title,PicSavePath,UserNumber From FS_ME_Photo where UserNumber='"&Fs_User.UserNumber&"' "& o_class &"",User_Conn,1,3
  if rs.eof then
  	Response.Write("û�з���������ͼƬ")
	rs.close:set rs=nothing
	conn.close
	Set conn=Nothing
	Response.End
  Else
	picNumber=rs.Recordcount
	Imglist="	ImgName[0]=""../sys_images/DefaultPreview.gif"";" & vbcrlf
	For i=1 To picNumber
  		Imglist = Imglist & "	ImgName[" & i & "]=""" &  rs("PicSavePath") & """;" & vbcrlf
  		rs.MoveNext
  	Next
	rs.close:set rs=nothing
  end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-������</title>
<script language="JavaScript">
var mid=0;
var t=0;
  function ImgArray(len)
  {
   this.length=len;
   }
  ImgName=new ImgArray(<%=picNumber%>);
  <%
  Response.write Imglist
  %>
  function play_filt()
  {
    t_end=document.player.intsec.value;
	if (t==<%=picNumber%>){t=0;}else{t++;}
	if (t==<%=picNumber%>){t_end=100;}
	if (mid==0){
	  img.style.filter="blendTrans(Duration=1)";
	  img.filters[0].apply();
	  img.src=ImgName[t];
	  tIndex=t;
	  img.filters[0].play();
	  mytimeout=setTimeout("play_filt()",t_end);
	}
   }
   function go(id){
   	if (id==1){
   		mid=0
   		play_filt();
   		}
   	else if(id==2){
   		mid=1;
   		}
   	else if(id==3){ 
   		mid=1;
   		t=t+1;
   		if (t<=<%=picNumber%>) img.src=ImgName[t+1]; 		
	}
	else if(id==4){
		mid=1;
		t=t-1;
		 if (t>0) img.src=ImgName[t];
	}
	else if(id==5){
		mid=1;
		t=0;
		play_filt();
	}
	else{
		mid=0;
		t=0;
		play_filt();
		}
   	}
</script>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<body   onload="play_filt()">
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td class="hback"><strong>λ�ã�</strong><a href="../">��վ��ҳ</a> &gt;&gt; 
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; <a href="PhotoManage.asp">������</a> 
            &gt;&gt;</td>
        </tr>
      </table> 
		  
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr>
          <td class="hback"><a href="PhotoManage.asp">�����ҳ</a>��<a href="Photo_add.asp">����ͼƬ</a>��<a href="PhotoManage.asp?isRec=1">���Ƽ���ͼƬ</a>��<a href="Photo_filt.asp">�õ�Ƭ����</a>��<a href="Photo_Class.asp">������</a></td>
        </tr>
        <tr> 
          <td class="hback"> 
            <%
		  response.Write("	<table width=""98%"" align=center cellpadding=""2"" cellspacing=""1""><tr>")
		  dim t_k
		  t_k=0
		  set rs = Server.CreateObject(G_FS_RS)
		  rs.open "select id,title,UserNumber From FS_ME_PhotoClass where UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
		  do while not rs.eof 
		  	Response.Write("	<td width=""24%"" valign=bottom><img src=""images/folderopened.gif""></img><a href=Photo_filt.asp?classid="&rs("id")&">"&rs("title")&"</a></td>")
		  rs.movenext
		  t_k = t_k+1
		  if t_k mod 4 =0 then
		  	Response.Write("	</tr>")
		  end if
		  loop
		  response.Write("	</table>")
		  rs.close:set rs=nothing
		  %>
          </td>
        </tr>
      </table> 
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" style="border-collapse: collapse" class="table">
        <tr class="hback">
          <td align=center class="xingmu"> 
            <p align="center"><b><font color="#000000">�õ�Ƭ����</font></b><BR/>
            </td>
    </tr>
  	<form name="player" method="post" action="">
  	<tr class="hback">
  	        <td> <div align="center">Ƶ�ʣ� 
                <select name="intsec">
                  <option value=1000>1��</option>
                  <option value=3000 selected>3��</option>
                  <option value=5000>5��</option>
                  <option value=8000>8��</option>
                  <option value=10000>10��</option>
                </select>
                <input type="button" value="��ʼ" onClick="javascipt:go(1)">
                <input type="button" value="ֹͣ" onClick="javascipt:go(2)">
                <input type="button" value="��һ��" onClick="javascipt:go(3)">
                <input type="button" value="��һ��" onClick="javascipt:go(4)">
                <input type="button" value="�ر�" onClick="javascipt:window.close();">
              </div></td>
  	</tr class="hback">
	</form>
   <tr class="hback">
          <td height="377" align=center valign=top> 
            <div align="center"><BR/>
              <img src="../sys_images/DefaultPreview.gif" name="img" style="cursor:hand" onClick="window.open(img.src);" onload="if(this.width>=400)this.width=400;"> 
              <BR/>
            </div></tr>
  </table>
       </td>
    </tr>
	
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="Copyright.asp" -->
        </div></td>
    </tr>
 
</table>
</body>
</html>
<%
Set Fs_User = Nothing
%>

<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->






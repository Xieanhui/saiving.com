<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
'�������
dim book_rs,new_count
set book_rs= Server.CreateObject(G_FS_RS)
book_rs.open "Select count(BookID) From FS_ME_Book where M_ReadUserNumber='"&Fs_User.UserNumber&"' and M_ReadTF=0 and M_Type=2",User_Conn,1,3
 if book_rs(0)>0 then
	 new_count = "<span class=""tx""><b>����������"& book_rs(0) &"</b></span>"
 else
	 new_count =  book_rs(0)
 end if
book_rs.close:set book_rs = nothing
if Request("Action")="del" then
	if trim(replace(Request("id")," ",""))="" then
		strShowErr = "<li>������ѡ��һ��</li>"
		Response.Redirect(""& s_savepath &"/lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	Conn.execute("Delete from FS_SD_News where id in ("& FormatIntArr(Request("id")) &") and UserNumber='"& Fs_User.UserNumber &"'")
		strShowErr = "<li>ɾ���ɹ�</li>"
		Response.Redirect(""& s_savepath &"/lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="& s_savepath &"/supply/PublicManage.asp")
		Response.end
end if
if NoSqlHack(Request.QueryString("Refresh"))="1" then
	dim u_rs
	set u_rs= Server.CreateObject(G_FS_RS)
	u_rs.open "select EditTime,id from FS_SD_News where id="& CintStr(Request.QueryString("id")) &" and UserNumber='"& Fs_User.UserNumber &"'",Conn,1,3
	u_rs("EditTime")=now
	u_rs.update
	u_rs.close:set u_rs=nothing
	strShowErr = "<li>�ط��ɹ�</li>"
	Response.Redirect(""& s_savepath &"/lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="& s_savepath &"/supply/PublicManage.asp")
	Response.end
end if
if NoSqlHack(Request("UserLock"))="1" then
	if trim(replace(Request("id")," ",""))="" then
		strShowErr = "<li>������ѡ��һ��</li>"
		Response.Redirect(""& s_savepath &"/lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	Conn.execute("Update FS_SD_News set islock=1 where id in ("& FormatIntArr(Request("id")) &") and UserNumber='"& Fs_User.UserNumber &"'")
		strShowErr = "<li>�����ɹ�</li>"
		Response.Redirect(""& s_savepath &"/lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="& s_savepath &"/supply/PublicManage.asp")
		Response.end
end if
if NoSqlHack(Request("UserLock"))="0" then
	if trim(replace(Request("id")," ",""))="" then
		strShowErr = "<li>������ѡ��һ��</li>"
		Response.Redirect(""& s_savepath &"/lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	Conn.execute("Update FS_SD_News set islock=0 where id in ("& FormatIntArr(Request("id")) &") and UserNumber='"& Fs_User.UserNumber &"'")
		strShowErr = "<li>�ύ�ɹ����ȴ����</li>"
		Response.Redirect(""& s_savepath &"/lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="& s_savepath &"/supply/PublicManage.asp")
		Response.end
end if
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,rs,i
int_RPP=20 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"
dim tmp_lock,tmp_valite,tmp_uselock,tmp_valite_1,para_rs,g_tmp
set para_rs=Conn.execute("select top 1 warnTime from FS_SD_Config")
if para_rs.eof then
	response.Write("�Ҳ�������������Ϣ")
	response.end
	set para_rs=nothing
else
	g_tmp=para_rs(0)
	if request.QueryString("usvalite")="1" then
		if G_IS_SQL_DB=1 then
			tmp_valite_1 = " and (dateadd(d,ValidTime,EditTime)-getdate())<="&para_rs(0)&" and (dateadd(d,ValidTime,EditTime)>=getdate())"
		else
			tmp_valite_1 =  " and (datevalue(editTime)+ValidTime)-datevalue(now)<="&para_rs(0)&" and (datevalue(editTime)+ValidTime)>=datevalue(now)"
		end if
	else
		tmp_valite_1=""
	end if
	set para_rs=nothing
end if
strpage=request("page")
'if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
</head>
<body>

<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="../top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="../Top_navi.asp" -->
    </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="../menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback">
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td  valign="top">���λ�ã�<a href="../../">��վ��ҳ</a> &gt;&gt; <a href="../main.asp">��Ա��ҳ</a> 
            &gt;&gt; <a href="PublicManage.asp">����ϵͳ</a> &gt;&gt; ������Ϣ����</td>
        </tr>
        <tr class="hback"> 
          <td  valign="top"><a href="PublicSupply.asp">������Ϣ</a>��<a href="PublicManage.asp">������Ϣ</a>��<a href="PublicManage.asp#top10">�����Ϣ�������(TOP10)</a>��<a href="PublicManage.asp#new">���¹�����Ϣ</a>��<a href="PublicManage.asp#rec">�����Ƽ�</a>��<a href="../Book.asp?M_type=2">�ҵ�������(<%=new_count%>)</a></td>
        </tr>
      </table>
     
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback">
          <td width="10%"><span class="tx">ǿ���Ƽ�</span></td> 
          <td width="90%">
		  <a href=" " id="printWord" ></a>
		  <%
		  	dim i_loop
			set rs= Server.CreateObject(G_FS_RS)
			if G_IS_SQL_DB=0 Then
				rs.open "select  top 6 ID,PubTitle,PubType,Addtime,hits,PubAddress,EditTime,UserNumber From FS_SD_News where isLock=0 and isPass=1 and (datevalue(editTime)+ValidTime)>datevalue(now) and OrderID=2 order by EditTime desc,id desc",Conn,1,3
			Else
				rs.open "select  top 6 ID,PubTitle,PubType,Addtime,hits,PubAddress,EditTime,UserNumber From FS_SD_News where isLock=0 and isPass=1 and dateadd(d,ValidTime,EditTime)>getdate() and OrderID=2 order by EditTime desc,id desc",Conn,1,3
			End if
			if rs.eof then
				response.Write("<span class=""tx"">û�Ƽ���Ϣ</span>")
			else
			i_loop=0
		  %>
		  <script language="JavaScript">
				var PublicTime = 3500; 
				var TextTime = 20; 
				var infoi = 0;
				var txti = 0;
				var txttimer;
				var infotimer;
				var publictitle = new Array(); 
				var publichref = new Array(); 
				<%do while not rs.eof %>
				publictitle[<%=i_loop%>] = "<%=Fs_User.GetFriendName(rs("UserNumber"))%>:<%=rs("PubTitle")%>[<%=rs("EditTime")%>]";
				publichref[<%=i_loop%>] = "../../Supply/Supply.asp?id=<%=rs("id")%>";
				<%rs.movenext%>
				<%i_loop=i_loop+1%>
				<%loop%>
				<%rs.close:set rs=nothing%>
				function showinfo()
				{
				var endstr = "_"
				hwinfotr = publictitle[infoi];
				infolink = publichref[infoi];
				if(txti==(hwinfotr.length-1)){endstr="";}
				if(txti>=hwinfotr.length){
				clearInterval(txttimer);
				clearInterval(infotimer);
				infoi++;
				if(infoi>=publictitle.length){
				infoi = 0
				}
				infotimer = setInterval("showinfo()",PublicTime);
				txti = 0;
				return;
				}
				clearInterval(txttimer);
				document.getElementById("printWord").href=infolink;
				document.getElementById("printWord").innerHTML = hwinfotr.substring(0,txti+1)+endstr;
				txti++;
				txttimer = setInterval("showinfo()",TextTime);
				}
				showinfo();
			</script>
			<%end if%>
		  </td>
        </tr>
      </table>
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="myForm" method="post" action="">
          <tr> 
            <td height="23" colspan="8" class="hback_1"><a href="PublicManage.asp">ȫ����Ϣ</a>��<a href="PublicManage.asp?islock=0">�����</a>��<a href="PublicManage.asp?islock=1">δ���</a>��<a href="PublicManage.asp?User_lock=1">�û�����</a>��<a href="PublicManage.asp?User_lock=0">�û�δ����</a>��<a href="PublicManage.asp?valite=1">������Ϣ</a>��<a href="PublicManage.asp?usvalite=1" title="��Ч��С��<%=g_tmp%>�����Ϣ">�������ڵ���Ϣ</a></td>
          </tr>
          <tr> 
            <td width="4%" class="hback_1"><div align="center"> 
                <input name="chkall" type="checkbox" id="chkall" onClick="CheckAll(this.form);" value="checkbox" title="���ѡ�����л��߳�������ѡ��">
              </div></td>
            <td width="27%" class="hback_1"><div align="center"><strong>����</strong></div></td>
            <td width="13%" class="hback_1"><div align="center"><strong>��¼ר��</strong></div></td>
            <td width="11%" class="hback_1"><div align="center"><strong>����ʱ��</strong></div></td>
            <td width="21%" class="hback_1"><div align="center"><strong>����</strong></div></td>
            <td width="8%" class="hback_1"><div align="center"><strong>��˽��</strong></div></td>
            <td width="8%" class="hback_1"><strong>�û�����</strong></td>
            <td width="8%" class="hback_1"><div align="center"><strong>���</strong></div></td>
          </tr>
          <%
		  	if Request.QueryString("islock")="0" then
				tmp_lock=" and ispass=1"
			elseif Request.QueryString("islock")="1" then
				tmp_lock=" and ispass=0"
			else
				tmp_lock=""
			end if
			if Request.QueryString("User_lock")="1" then
				tmp_uselock = " and islock=1"
			elseif Request.QueryString("User_lock")="0" then
				tmp_uselock = " and islock=0"
			else
				tmp_uselock = " "
			end if
			if Request.QueryString("valite")="1" then
				if G_IS_SQL_DB=1 then
					tmp_valite = " and dateadd(d,ValidTime,EditTime)<getdate()"
				else
					tmp_valite = " and (datevalue(editTime)+ValidTime)<datevalue(now)"
				end if
			else
				tmp_valite = " "
			end if
			set rs = Server.CreateObject(G_FS_RS)
			rs.open "select id,PubTitle,EditTime,ValidTime,ispass,islock,hits,MyClassID from FS_SD_News where UserNumber='"& Fs_User.UserNumber&"' " &tmp_lock & tmp_valite & tmp_uselock& tmp_valite_1 &" order by EditTime desc,OrderID desc,id desc",Conn,1,3
			if rs.eof then
			   rs.close
			   set rs=nothing
			   Response.Write"<tr  class=""hback""><td colspan=""8""  class=""hback"" height=""40"">û�м�¼��</td></tr>"
			else
				rs.PageSize=int_RPP
				cPageNo=NoSqlHack(Request.QueryString("Page"))
				If cPageNo="" Then cPageNo = 1
				If not isnumeric(cPageNo) Then cPageNo = 1
				cPageNo = Clng(cPageNo)
				If cPageNo<=0 Then cPageNo=1
				If cPageNo>rs.PageCount Then cPageNo=rs.PageCount 
				rs.AbsolutePage=cPageNo
				for i=1 to rs.pagesize
					if rs.eof Then exit For 
		  %>
          <tr  onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td class="hback"><input name="id" type="checkbox" id="id" value="<%=rs("id")%>"></td>
            <td class="hback"><a href="<% = Replace("/"&G_VIRTUAL_ROOT_DIR&"/" ,"//","/")%>Supply/Supply.asp?id=<%=rs("id")%>" target="_blank"><%=rs("PubTitle")%></a></td>
            <td class="hback"><div align="center"> 
                <%
			dim rs_z
			set rs_z = User_Conn.execute("select ClassID,ClassCName From FS_ME_InfoClass where ClassID="& rs("MyClassID")&" and UserNumber='"&Fs_User.UserNumber&"'")
			if rs_z.eof then
				response.Write "--"
			else
				Response.Write "<a href=../ShowInfoClass.asp?ClassID="& rs_z("Classid")&"&UserNumber="& Fs_User.UserNumber &">" &rs_z("ClassCName")&"</a>"
			end if
			%>
              </div></td>
            <td class="hback"><div align="center"><%=datevalue(rs("EditTime"))+rs("ValidTime")%></div></td>
            <td class="hback"><div align="center"><a href="PublicSupplyEdit.asp?id=<%=rs("id")%>&Action=edit">�޸�</a> 
                <a href="PublicManage.asp?id=<%=rs("id")%>&Refresh=1">�ط�</a> 
                <a href="PublicSupplyEdit.asp?id=<%=rs("id")%>&Action=Copy">����</a> 
                <%if rs("islock")=0 then%>
                <a href="PublicManage.asp?id=<%=rs("id")%>&UserLock=1"  onClick="{if(confirm('ȷ��Ҫ������\n������ǰ̨��������ʾ')){return true;}return false;}">����</a> 
                <%else%>
                <a href="PublicManage.asp?id=<%=rs("id")%>&UserLock=0"  onClick="{if(confirm('ȷ��Ҫ������\n������ǰ̨��������ʾ')){return true;}return false;}">�ύ</a> 
                <%end if%>
                <a href="PublicManage.asp?id=<%=rs("id")%>&Action=del"  onClick="{if(confirm('ȷ��Ҫɾ����')){return true;}return false;}">ɾ��</a> 
              </div></td>
            <td class="hback"><div align="center"> 
                <%if rs("ispass")=1 then:response.Write"�����":else:Response.Write"<span class=""tx"">δ���</span>":end if%>
              </div></td>
            <td class="hback"><div align="center"> 
                <%if rs("islock")=1 then:response.Write("<span class=""tx"">�ѳ���</span>"):else:Response.Write("δ����"):end if%>
              </div></td>
            <td class="hback"><div align="center"><%=rs("hits")%></div></td>
          </tr>
          <%
			  rs.movenext
		  next
		  %>
          <tr  class="hback"> 
            <td colspan="8"><div align="right"> 
                <%
			response.Write "<p>"&  fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
			%>
                <input name="Action" type="hidden" id="Action">
                <input name="UserLock" type="hidden" id="UserLock">
                <input type="button" name="Submit" value="ɾ��"  onClick="document.myForm.Action.value='del';{if(confirm('ȷ���������ѡ��ļ�¼��')){this.document.myForm.submit();return true;}return false;}">
                <input type="button" name="Submit2" value="������Ϣ"  onClick="document.myForm.UserLock.value='1';{if(confirm('ȷ��������Ϣ��\n�����󽫲�����ʾ')){this.document.myForm.submit();return true;}return false;}">
                <input type="button" name="Submit3" value="�ύ��Ϣ"  onClick="document.myForm.UserLock.value='0';{if(confirm('ȷ�������ύ��Ϣ��')){this.document.myForm.submit();return true;}return false;}">
              </div></td>
          </tr>
          <% 
		  rs.close:set rs=nothing
		  end if
		  %>
        </form>
      </table>
      <a name="top10"></a>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr> 
          <td colspan="5" class="xingmu"> 
            <table width="98%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="2%" valign="top"><strong><img src="../Images/folderclosed.gif" width="16" height="16"></strong></td>
                <td width="98%" valign="bottom" class="xingmu"><strong>�����Ϣ���������ǰ10��</strong> 
                </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td width="9%" class="hback_1"><div align="center"><strong>����</strong></div></td>
          <td width="33%" class="hback_1"><div align="center"><strong>����</strong></div></td>
          <td width="23%" class="hback_1"><div align="center"><strong>��������</strong></div></td>
          <td width="18%" class="hback_1"><div align="center"><strong>��������</strong></div></td>
          <td width="17%" class="hback_1"><div align="center"><strong>���</strong></div></td>
        </tr>
        <%
		set rs= Server.CreateObject(G_FS_RS)
		rs.open "select  top 10 ID,PubTitle,PubType,Addtime,ValidTime,EditTime,hits,UserNumber From FS_SD_News where UserNumber='"& Fs_User.UserNumber &"' order by hits desc,id desc",Conn,1,3
		do while not rs.eof 
		%>
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
          <td class="hback"> <div align="center"> 
              <%
		  select Case rs("PubType")
		  	case 0
				response.Write("��Ӧ")
			case 1
				response.Write("<span class=""tx"">��</span>")
			case 2
				response.Write("����")
			case 3
				response.Write("����")
			case else
				Response.Write("����")
		 end select
		  %>
            </div></td>
          <td class="hback"><div align="left"><a href="<% = Replace("/"&G_VIRTUAL_ROOT_DIR&"/" ,"//","/")%>Supply/Supply.asp?id=<%=rs("id")%>" target="_blank"><%=rs("PubTitle")%></a></div></td>
          <td class="hback"><div align="center"><%=rs("Addtime")%></div></td>
          <td class="hback"><div align="center"><%=datevalue(rs("EditTime"))+rs("ValidTime")%></div></td>
          <td class="hback"><div align="center"><a href=ShowUser.asp?UserNumber=<%=rs("UserNumber")%> target="_blank"><%=rs("hits")%></a></div></td>
        </tr>
        <%
			rs.movenext
		loop
		rs.close:set rs=nothing
		%>
      </table>
      <a name="new"></a> <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr> 
          <td colspan="5" class="xingmu"> 
            <table width="98%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="2%" valign="top"><strong><img src="../Images/folderclosed.gif" width="16" height="16"></strong></td>
                <td width="98%" valign="bottom"  class="xingmu"><strong>���¹�����Ϣ20��</strong></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td width="9%" class="hback_1"><div align="center"><strong>����</strong></div></td>
          <td width="33%" class="hback_1"><div align="center"><strong>����</strong></div></td>
          <td width="23%" class="hback_1"><div align="center"><strong>��������</strong></div></td>
          <td width="18%" class="hback_1"><div align="center"><strong>����</strong></div></td>
          <td width="17%" class="hback_1"><div align="center"><strong>������</strong></div></td>
        </tr>
        <%
		set rs= Server.CreateObject(G_FS_RS)
		if G_IS_SQL_DB=0 Then
			rs.open "select  top 20 ID,PubTitle,PubType,Addtime,hits,PubAddress,UserNumber From FS_SD_News where isLock=0 and isPass=1 and (datevalue(editTime)+ValidTime)>datevalue(now) order by EditTime desc,id desc",Conn,1,3
		Else
			rs.open "select  top 20 ID,PubTitle,PubType,Addtime,hits,PubAddress,UserNumber From FS_SD_News where isLock=0 and isPass=1 and dateadd(d,ValidTime,EditTime)>getdate() order by EditTime desc,id desc",Conn,1,3
		End if
		do while not rs.eof 
		%>
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
          <td class="hback"> <div align="center"> 
              <%
		  select Case rs("PubType")
		  	case 0
				response.Write("��Ӧ")
			case 1
				response.Write("<span class=""tx"">��</span>")
			case 2
				response.Write("����")
			case 3
				response.Write("����")
			case else
				Response.Write("����")
		 end select
		  %>
            </div></td>
          <td class="hback"><div align="left"><a href="<% = Replace("/"&G_VIRTUAL_ROOT_DIR&"/" ,"//","/")%>Supply/Supply.asp?id=<%=rs("id")%>" target="_blank"><%=rs("PubTitle")%></a></div></td>
          <td class="hback"><div align="center"><%=rs("Addtime")%></div></td>
          <td class="hback"><div align="center"><%=rs("PubAddress")%></div></td>
          <%if rs("UserNumber")="0" then%>
          <td class="hback"><div align="center">����Ա</div></td>
          <%else%>
          <td class="hback"><div align="center"><a href=ShowUser.asp?UserNumber=<%=rs("UserNumber")%> target="_blank"><%=Fs_User.GetFriendName(rs("UserNumber"))%></a></div></td>
          <%end if%>
        </tr>
        <%
			rs.movenext
		loop
		rs.close:set rs=nothing
		%>
      </table>
      <a name="rec" id="rec"></a> 
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr> 
          <td colspan="5" class="xingmu"> 
            <table width="98%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="2%" valign="top"><strong><img src="../Images/folderclosed.gif" width="16" height="16"></strong></td>
                <td width="98%" valign="bottom" class="xingmu"><strong>�����Ƽ���Ϣ20��</strong></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td width="9%" class="hback_1"><div align="center"><strong>����</strong></div></td>
          <td width="33%" class="hback_1"><div align="center"><strong>����</strong></div></td>
          <td width="23%" class="hback_1"><div align="center"><strong>��������</strong></div></td>
          <td width="18%" class="hback_1"><div align="center"><strong>����</strong></div></td>
          <td width="17%" class="hback_1"><div align="center"><strong>������</strong></div></td>
        </tr>
        <%
		set rs= Server.CreateObject(G_FS_RS)
		if G_IS_SQL_DB=0 Then
			rs.open "select  top 20 ID,PubTitle,PubType,Addtime,hits,PubAddress,UserNumber From FS_SD_News where isLock=0 and isPass=1 and (datevalue(editTime)+ValidTime)>datevalue(now) and OrderID=1 order by EditTime desc,id desc",Conn,1,3
		Else
			rs.open "select  top 20 ID,PubTitle,PubType,Addtime,hits,PubAddress,UserNumber From FS_SD_News where isLock=0 and isPass=1 and dateadd(d,ValidTime,EditTime)>getdate() and OrderID=1 order by EditTime desc,id desc",Conn,1,3
		End if
		do while not rs.eof 
		%>
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
          <td class="hback"> <div align="center"> 
              <%
		  select Case rs("PubType")
		  	case 0
				response.Write("��Ӧ")
			case 1
				response.Write("<span class=""tx"">��</span>")
			case 2
				response.Write("����")
			case 3
				response.Write("����")
			case else
				Response.Write("����")
		 end select
		  %>
            </div></td>
          <td class="hback"><div align="left"><a href="<% = Replace("/"&G_VIRTUAL_ROOT_DIR&"/" ,"//","/")%>Supply/Supply.asp?id=<%=rs("id")%>" target="_blank"><%=rs("PubTitle")%></a></div></td>
          <td class="hback"><div align="center"><%=rs("Addtime")%></div></td>
          <td class="hback"><div align="center"><%=rs("PubAddress")%></div></td>
          <%if rs("UserNumber")="0" then%>
          <td class="hback"><div align="center">����Ա</div></td>
          <%else%>
          <td class="hback"><div align="center"><a href=ShowUser.asp?UserNumber=<%=rs("UserNumber")%> target="_blank"><%=Fs_User.GetFriendName(rs("UserNumber"))%></a></div></td>
          <%end if%>
        </tr>
        <%
			rs.movenext
		loop
		rs.close:set rs=nothing
		%>
      </table></td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="../Copyright.asp" -->
        </div></td>
    </tr>
</table>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = myForm.elements[i];  
    if (e.name != 'chkall')  
       e.checked = myForm.chkall.checked;  
    }  
	}
</script>

<%
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->






<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
	dim Conn,User_Conn,strShowErr
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF
	if not MF_Check_Pop_TF("ME_Photo") then Err_Show 
	if not MF_Check_Pop_TF("ME040") then Err_Show 

	Function GetFriendName(f_strNumber)
		Dim RsGetFriendName
		Set RsGetFriendName = User_Conn.Execute("Select UserName From FS_ME_Users Where UserNumber = '"& f_strNumber &"'")
		If  Not RsGetFriendName.eof  Then 
			GetFriendName = RsGetFriendName("UserName")
		Else
			GetFriendName = 0
		End If 
		set RsGetFriendName = nothing
	End Function 
	if Request("Action")="del" then
		if trim(Request("Id"))="" then
			strShowErr = "<li>��ѡ������һ��</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		User_Conn.execute("Delete From FS_ME_Photo where id in ("&FormatIntArr(Request("Id"))&")")
		strShowErr = "<li>ɾ���ɹ�</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Photo.asp")
		Response.End
	ElseIf Request("Action")="rec" Then
		User_Conn.execute("Update FS_ME_Photo set isRec=1 where id in ("&FormatIntArr(Request("Id"))&")")
		strShowErr = "<li>�޸ĳɹ�</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Photo.asp")
		Response.End
	Elseif Request("Action")="unrec" Then 
		User_Conn.execute("Update FS_ME_Photo set isRec=0 where id in ("&FormatIntArr(Request("Id"))&")")
		strShowErr = "<li>�޸ĳɹ�</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Photo.asp")
		Response.End
	end If
	if Request("Action")="all" then
		User_Conn.execute("Delete From FS_ME_Photo")
			strShowErr = "<li>ɾ�����гɹ�</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Photo.asp")
			Response.end
	end if
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,i
	int_RPP=15 '����ÿҳ��ʾ��Ŀ
	int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
	showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
	str_nonLinkColor_="#999999" '����������ɫ
	toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
	toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
	toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
	toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
	toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
	toL_="<font face=webdings title=""���һҳ"">:</font>"
	strpage=request("page")
	if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td width="100%" class="xingmu">������</td>
  </tr class="hback">
  <tr class="hback">
    <td class="hback"><a href="UserReport.asp">����</a>��<a href="Photo.asp?Action=all" onClick="{if(confirm('ȷ��ɾ����')){return true;}return false;}">����������</a></td>
  </tr class="hback">
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="myform" method="post" action="">
    <%
			dim rs,isrec
			set rs = Server.CreateObject(G_FS_RS)
			rs.open "select * from FS_ME_Photo  order by id desc",User_Conn,1,1
			if rs.eof then
			   rs.close
			   set rs=nothing
			   Response.Write"<tr  class=""hback""><td colspan=""1""  class=""hback"" height=""40"">û�м�¼��</td></tr>"
			else
				rs.PageSize=int_RPP
				cPageNo=NoSqlHack(Request.QueryString("Page"))
				If cPageNo="" Then cPageNo = 1
				If not isnumeric(cPageNo) Then cPageNo = 1
				cPageNo = Clng(cPageNo)
				If cPageNo>rs.PageCount Then cPageNo=rs.PageCount 
				If cPageNo<=0 Then cPageNo=1
				rs.AbsolutePage=cPageNo
				for i=1 to rs.pagesize
					if rs.eof Then exit For
					isrec=rs("isrec")
			%>
    <tr class="hback"> 
      <td width="21%" rowspan="5" class="hback_1"> <table border="0" align="center" cellpadding="2" cellspacing="1" class="table">
          <tr> 
            <td class="hback">
              <%if isnull(trim(rs("PicSavePath"))) then%>
              <img src="Images/nopic_supply.gif" width="90"  id="pic_p_1">
              <%else%>
              <a href="<%=rs("PicSavePath")%>" target="_blank"><img src="<%=rs("PicSavePath")%>" width="90" border="0" id="pic_p_1"></a>
              <%end if%>
            </td>
          </tr>
        </table></td>
      <td width="12%" class="hback"><div align="center"><strong>��Ƭ���ƣ�</strong></div></td>
      <td width="40%" class="hback"><font style="font-size:14px"><span class="hback_1">
	  <%if isrec=1 Then Response.write("<img src=""../images/award.gif"" alt=""�Ƽ�"" />")%>
	  <strong><%=rs("title")%></strong></span></font><a href="../../<%=G_USER_DIR%>/ShowUser.asp?UserNumber=<% = rs("UserNumber")%>" target="_blank">(<%=GetFriendName(rs("UserNumber"))%>)</a></td>
      <td width="10%"><div align="center">���������</div></td>
      <td width="17%"><%=rs("hits")%></td>
    </tr>
    <tr class="hback"> 
      <td><div align="center">�������ڣ�</div></td>
      <td><%=rs("Addtime")%></td>
      <td><div align="center">ͼƬ��С��</div></td>
      <td><%=rs("PicSize")%>byte</td>
    </tr>
    <tr class="hback"> 
      <td><div align="center">��Ƭ��ַ��</div></td>
      <td colspan="3"><%=rs("PicSavePath")%></td>
    </tr>
    <tr class="hback"> 
      <td><div align="center">��Ƭ������</div></td>
      <td colspan="3"><%=rs("Content")%></td>
    </tr>
    <tr class="hback"> 
      <td><div align="center">��Ƭ���ࣺ</div></td>
      <td>
        <%
			dim c_rs
			if rs("ClassID")=0 then
				response.Write("û����")
			else
				set c_rs=User_Conn.execute("select ID,title From Fs_me_photoclass where id="&rs("classid"))
				Response.Write c_rs("title")
				c_rs.close:set c_rs=nothing
			end if
			%>
      </td>
      <td colspan="2"><div align="center"><a href="Photo.asp?id=<%=rs("id")%>&Action=del" onClick="{if(confirm('ȷ��ͨ��ɾ����')){return true;}return false;}">ɾ��</a>  |  
	  <a href="Photo.asp?id=<%=rs("id")%>&Action=rec">�Ƽ�</a> |
	  <a href="Photo.asp?id=<%=rs("id")%>&Action=unrec">ȡ���Ƽ�</a> |
          <input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>">
        </div></td>
    </tr>
    <tr class="hback"> 
      <td colspan="5" height="3" class="xingmu"></td>
    </tr>
    <%
			rs.movenext
		next
		%>
    <tr class="hback"> 
      <td colspan="5"> 
        <%
			response.Write "<p>"&  fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
		  end if
		  %>
        �� 
        <input name="chkall" type="checkbox" id="chkall" onClick="CheckAll(this.form);" value="checkbox" title="���ѡ�����л��߳�������ѡ��">
        ȫѡ�� 
        <input name="Action" type="hidden" id="Action"> <input type="button" name="Submit" value="ɾ��"  onClick="document.myform.Action.value='del';{if(confirm('ȷ���������ѡ��ļ�¼��')){this.document.myform.submit();return true;}return false;}"></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
Conn.close:set conn=nothing
User_Conn.close:set User_Conn=nothing
%>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<myform.elements.length;i++)  
    {  
    var e = myform.elements[i];  
    if (e.name != 'chkall')  
       e.checked = myform.chkall.checked;  
    }  
  }
</script>







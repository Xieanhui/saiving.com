<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
Dim straction
straction = Request("action")
if straction="Unmessage" then
	User_Conn.execute("update FS_ME_Users set ismessage= 0 where UserNumber='"& Fs_User.UserNumber &"'")
	strShowErr = "<li>���ı�վ����ȡ��</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../main.asp")
	Response.end
Elseif straction = "ismessage" then
	User_Conn.execute("update FS_ME_Users set ismessage= 1 where UserNumber='"& Fs_User.UserNumber &"'")
	strShowErr = "<li>���ı�վ���ϳɹ�</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../main.asp")
	Response.end
Elseif straction = "Open" then
	User_Conn.execute("update FS_ME_Users set isOpen= 1 where UserNumber='"& Fs_User.UserNumber &"'")
	strShowErr = "<li>���⿪�����Ͽ���</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../main.asp")
	Response.end
Elseif straction = "Close" then
	User_Conn.execute("update FS_ME_Users set isOpen= 0 where UserNumber='"& Fs_User.UserNumber &"'")
	strShowErr = "<li>���⿪������ȡ��</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../main.asp")
	Response.end
End if
%>
<%
Dim house_rs,mystatus,order,sqlstatement,order_status,audit
Dim HouseName_order,Class_order,OpenDate_order
if not isCorp() then 
	response.Redirect("HS_Tenancy.asp")
End if
session("audit")=NoSqlHack(request.querystring("audit"))
order=NoSqlHack(request.QueryString("order"))
mystatus=CintStr(request.QueryString("status"))
if mystatus=0 then
	mystatus=1
End if
if session("orderstatus")="" then
	session("orderstatus")="asc"
End if
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,i
'---------------------------------��ҳ����
int_RPP=15 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"			'βҳ
'--------------------------------------------------
writeinfo(session("orderstatus") )
select case order
	case "orderbyname" 
		if session("audit")="1" then 
			sqlstatement="select ID,HouseName,Position,Direction,Class,OpenDate,PreSaleNumber,IssueDate,PreSaleRange,Status,Price,PubDate,Tel,Click,UserNumber,Audited from FS_HS_Quotation where status="&CintStr(mystatus)&" and Audited=1 and Usernumber='"&session("FS_UserNumber")&"' order  by HouseName "&session("orderstatus") 
		else 
			sqlstatement="select ID,HouseName,Position,Direction,Class,OpenDate,PreSaleNumber,IssueDate,PreSaleRange,Status,Price,PubDate,Tel,Click,UserNumber,Audited from FS_HS_Quotation where status="&CintStr(mystatus)&"  and Usernumber='"&session("FS_UserNumber")&"' order by HouseName "&session("orderstatus") 
		end if
	
	case "orderbyclass" 
		if session("audit")="1" then 
			sqlstatement="select ID,HouseName,Position,Direction,Class,OpenDate,PreSaleNumber,IssueDate,PreSaleRange,Status,Price,PubDate,Tel,Click,UserNumber,Audited from FS_HS_Quotation where status="&CintStr(mystatus)&" and Audited=1 and Usernumber='"&session("FS_UserNumber")&"' order  by class "&session("orderstatus") 
		else
			sqlstatement="select ID, HouseName,Position,Direction,Class,OpenDate,PreSaleNumber,IssueDate,PreSaleRange,Status,Price,PubDate,Tel,Click,UserNumber,Audited from FS_HS_Quotation where status="&CintStr(mystatus)&" and Usernumber='"&session("FS_UserNumber")&"' order by Class "&session("orderstatus")
		end if
	
	case "orderbytime" 
		if session("audit")="1" then
			sqlstatement="select ID,HouseName,Position,Direction,Class,OpenDate,PreSaleNumber,IssueDate,PreSaleRange,Status,Price,PubDate,Tel,Click,UserNumber,Audited from FS_HS_Quotation where status="&CintStr(mystatus)&" and Audited=1 and Usernumber='"&session("FS_UserNumber")&"' order  by OpenDate "&session("orderstatus") 
		else
			sqlstatement="select ID ,HouseName,Position,Direction,Class,OpenDate,PreSaleNumber,IssueDate,PreSaleRange,Status,Price,PubDate,Tel,Click,UserNumber,Audited from FS_HS_Quotation where status="&CintStr(mystatus)&" and Usernumber='"&session("FS_UserNumber")&"' order by OpenDate "&session("orderstatus")
		End if
	
	case else 
		if session("audit")="1" then
			sqlstatement="select ID,HouseName,Position,Direction,Class,OpenDate,PreSaleNumber,IssueDate,PreSaleRange,Status,Price,PubDate,Tel,Click,UserNumber,Audited from FS_HS_Quotation where status="&CintStr(mystatus)&" and Audited=1 and Usernumber='"&session("FS_UserNumber")&"' order  by id desc"
		else
			sqlstatement="select ID, HouseName,Position,Direction,Class,OpenDate,PreSaleNumber,IssueDate,PreSaleRange,Status,Price,PubDate,Tel,Click,UserNumber,Audited from FS_HS_Quotation  where status="&CintStr(mystatus)&" and Usernumber='"&session("FS_UserNumber")&"' order by id desc"
		End if
end select
'����״̬
Function changeOrderStatus()
	Dim order_desc
	if session("orderstatus")="asc" then
		session("orderstatus")="desc"
		order_desc="<font color='red'>��</font>"
	else
		session("orderstatus")="asc"
		order_desc="<font color='red'>��</font>"
	End if
	select case order
		case "orderbyname"  HouseName_order=order_desc
							Class_order=""
							OpenDate_order=""
							
		case "orderbyclass" Class_order=order_desc
							HouseName_order=""
							OpenDate_order=""
							
		case "orderbytime" OpenDate_order=order_desc
							HouseName_order=""
							Class_order=""
	end select
End Function
if trim(order)<>"" then
	changeOrderStatus()
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>User Manage Center-��վ���ݹ���ϵͳ</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<script language="javascript" src="../../FS_Inc/CheckJs.js"></script>
<script language="javascript" src="../../FS_Inc/prototype.js"></script>
<script language="javascript" src="../../FS_Inc/PublicJS.js"></script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td><!--#include file="../top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr class="back">
    <td   colspan="2" class="xingmu" height="26"><!--#include file="../Top_navi.asp" -->
    </td>
  </tr>
  <tr class="back">
    <td width="18%" valign="top" class="hback"><div align="left">
        <!--#include file="../menu.asp" -->
      </div></td>
    <td width="82%" valign="top" class="hback">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
        <tr class="xingmu">
          <td align="center"><a href="houseManage.asp" style="color:#FF0000"  class="sd">¥����Ϣ</a></td>
          <td align="center"><a href="HS_Tenancy.asp"  class="sd">������Ϣ</a></td>
          <td align="center"><a href="HS_Second.asp"  class="sd">���ַ���Ϣ</a></td>
        <tr>
          <td colspan="3">
		  <div id="container">
              <!--ajax����-->
              <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
                
                <tr>
                  <td align="center" <%if mystatus=1 or mystatus="" then response.Write("class='xingmu'") else response.Write("class='hback'")%>><a href="houseManage.asp?status=1&audit=<%=session("audit")%>" class="sd">����չʾ¥��</a></td>
                  <td align="center" <%if mystatus=2 then response.Write("class='xingmu'") else response.Write("class='hback'")%>><a href="houseManage.asp?status=2&audit=<%=session("audit")%>" class="sd">�ڷ�¥����Ϣ</a></td>
                  <td align="center" <%if mystatus=3 then response.Write("class='xingmu'") else response.Write("class='hback'")%>><a href="houseManage.asp?status=3&audit=<%=session("audit")%>" class="sd">�ַ�¥����Ϣ��Ϣ</a></td>
                  <td align="center" class="hback"><a href="#" class="sd" title="ѡ�и�ѡ�򣬲鿴�Ѿ������Ϣ" onClick="linkShowaudit();showaudit();">�������Ϣ</a><input type="checkbox" name="myaudit" id="audit" value="audit" onClick="showaudit()" <%if session("audit")="1" then response.Write("checked")%>/></td>
				<tr>
              </table>
              <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
                <tr class="hback_1">
                  <td class="hback_1" align="center" width="30%"><a href="houseManage.asp?order=orderbyname&status=<%=mystatus%>&audit=<%=session("audit")%>" class="sd">¥������<%=HouseName_order%></a></td>
                  <td class="hback_1" align="center" width="25%"><a href="houseManage.asp?order=orderbyclass&status=<%=mystatus%>&audit=<%=session("audit")%>" class="sd">��Ŀ���<%=Class_order%></a></td>
                  <td class="hback_1" align="center" width="15%"><a href="houseManage.asp?order=orderbytime&status=<%=mystatus%>&audit=<%=session("audit")%>" class="sd">��������<%=OpenDate_order%></a></td>
                  <td class="hback_1" align="center" width="20%"><a href="#" class="sd">����</a></td>
                  <td class="hback_1" align="center" width="5%"><input type="checkbox" value="" name="quotationlist"  onclick="selectAll(document.all('quotationlist'))"/></td>
                </tr>
                <%
					Set house_rs=Server.CreateObject(G_FS_RS)
					house_rs.open sqlstatement,Conn,1,1
					If Not house_rs.eof then
					'��ҳʹ��-----------------------------------
						house_rs.PageSize=int_RPP
						cPageNo=NoSqlHack(Request.QueryString("page"))
						If cPageNo="" Then cPageNo = 1
						If not isnumeric(cPageNo) Then cPageNo = 1
						cPageNo = Clng(cPageNo)
						If cPageNo<=0 Then cPageNo=1
						If cPageNo>house_rs.PageCount Then cPageNo=house_rs.PageCount 
						house_rs.AbsolutePage=cPageNo
					End if
					for i=0 to int_RPP
						if house_rs.eof then exit for
						if house_rs("audited")="0" then
							audit="<font color='red'>δ���</font>"
						else
							audit="�����"
						End if
						Response.Write("<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>"&vbcrlf)
						response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('detail"&house_rs("ID")&"','_detail"&house_rs("ID")&"')""><a href='HS_Quotation_Edit.asp?action=edit&id="&house_rs("ID")&"'>"&house_rs("HouseName")&"</a></td>"&vbcrlf)
						response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('detail"&house_rs("ID")&"','_detail"&house_rs("ID")&"')"">"&house_rs("Class")&"</td>"&vbcrlf)
						response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('detail"&house_rs("ID")&"','_detail"&house_rs("ID")&"')"">"&house_rs("OpenDate")&"</td>"&vbcrlf)
						response.Write("<td align='center' class='hback'><span id='audit"&house_rs("ID")&"'>"&audit&"</span>|<span id='edit"&house_rs("ID")&"'><a href='HS_Quotation_Edit.asp?action=edit&id="&house_rs("ID")&"'>�޸�</a></span>|<span id='delete"&house_rs("ID")&"'><a href='#' class='sd' onclick=""deleteAction('delete"&house_rs("ID")&"','"&house_rs("ID")&"')"">ɾ��</a></span></td>"&vbcrlf)
						response.Write("<td align='center' class='hback'><input type='checkbox' name='quotationlist' value='"&house_rs("ID")&"'</td>"&vbcrlf)
						response.Write("</tr>"&vbcrlf)
						Response.Write("<tr id=""detail"&house_rs("ID")&""" style=""display:none"">"&vbcrlf)
						response.Write("<td colspan='7' class='hback'>"&vbcrlf)
						Response.Write("��λ�ã�"&house_rs("Position")&"|����"&house_rs("Direction")&vbcrlf)
						response.Write("</td>"&vbcrlf)
						response.Write("</tr>"&vbcrlf)
						response.Write("<tr id=""_detail"&house_rs("ID")&""" style=""display:none"">"&vbcrlf)
						response.Write("<td colspan='7' class='hback'>"&vbcrlf)
						response.Write("��Ԥ�����֤��["&house_rs("PreSaleNumber")&"]|���ۣ���"&house_rs("Price")&"��|��ϵ�绰��"&house_rs("Tel")&vbcrlf)
						response.Write("</td>"&vbcrlf)
						Response.Write("</tr>"&vbcrlf)
						response.Write("</div>"&vbcrlf)
						house_rs.movenext
					next
	
				%>
                <tr>
                  <td colspan="5" align="right" class="hback"><button onClick="javascript:location='HS_Quotation_Edit.asp?action=add'">������Ϣ</button>
                    <button onClick="deleteBatAction()">����ɾ��</button>
                    &nbsp; </td>
                </tr>
                <%
					Response.Write("<tr>"&vbcrlf)
					Response.Write("<td align='right' colspan='7'  class=""hback"">"&fPageCount(house_rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td>"&vbcrlf)
					Response.Write("</tr>"&vbcrlf)
				%>
              </table>
            </div></tr>
      </table></td>
  </tr>
  <tr class="back">
    <td height="20"  colspan="2" class="xingmu"><div align="left">
        <!--#include file="../Copyright.asp" -->
      </div></td>
  </tr>
</table>
</body>
</html>
<%
Set Fs_User = Nothing
Set User_Conn=nothing
Set Conn=nothing
house_rs.close
Set house_rs=nothing
%>
<script language="javascript">
function linkShowaudit()
{
	if($('audit').checked)
		$('audit').checked=false;
	else
		$('audit').checked=true;
	
}
function showaudit()
{
	if($('audit').checked)
	{
		location="houseManage.asp?audit=1&status=<%=mystatus%>"	
	}else
	{
		location="houseManage.asp?audit=0&status=<%=mystatus%>";
	}
}
//����
//ɾ��
function deleteAction(container,id)
{
	var url="HS_Quotation_Open_Action.asp?rad="+Math.random();
	var param="action=delete"+"&id="+id;
	if(confirm("ȷ��Ҫɾ��������Ϣ��"))
		var myAjax = new Ajax.Request(url,{method: 'get', parameters: param, onComplete: showResponse});
	function showResponse(originalRequest)
	{
		var value= originalRequest.responseText;
		if(value=="ok")
		{
			$(container).parentNode.parentNode.style.display='none';
			$('recordcount').innerHTML=parseInt($('recordcount').innerHTML)-1
			alert('ɾ���ɹ���')
		}
		else
		{
			alert("�����쳣������ͷ���Ա��ϵ��");
		}
	}
}
//����ɾ��
function deleteBatAction()
{
	var url="HS_Quotation_Open_Action.asp?rad="+Math.random();
	var elements=document.all('quotationlist');
	var id="";
	var count=0;
	for(var i=1;i<elements.length;i++)
	{
		if(elements[i].checked)
		{
			if (id==""){
				id=elements[i].value;
			}else{
				id+=","+elements[i].value;
			}
			if(elements[i].parentNode.parentNode.style.display!="none")
				count+=1;
		}
	}
	
	if(id=="")
	{
		alert("��ѡ��Ҫɾ���ļ�¼��");
		return;
	}
	param="action=delete"+"&id="+id;
	if(confirm("ȷ��Ҫɾ��������Ϣ��"))
		var myAjax = new Ajax.Request(url,{method: 'get', parameters: param, onComplete: showResponse});
	function showResponse(originalRequest)
	{
		var value= originalRequest.responseText;
		if(value=="ok")
		{
			for(var i=1;i<elements.length;i++)
			{
				if(elements[i].checked)
				{
					elements[i].parentNode.parentNode.style.display='none';
				}
			}
			$('recordcount').innerHTML=parseInt($('recordcount').innerHTML)-count;
			alert(count+'����Ϣɾ���ɹ���')
			count=0;
		}
		else
		{
			alert("�����쳣������ͷ���Ա��ϵ��");
			alert(originalRequest.responseText);
		}
	}
}
</script>
</html>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->







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
	strShowErr = "<li>订阅本站资料取消</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../main.asp")
	Response.end
Elseif straction = "ismessage" then
	User_Conn.execute("update FS_ME_Users set ismessage= 1 where UserNumber='"& Fs_User.UserNumber &"'")
	strShowErr = "<li>订阅本站资料成功</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../main.asp")
	Response.end
Elseif straction = "Open" then
	User_Conn.execute("update FS_ME_Users set isOpen= 1 where UserNumber='"& Fs_User.UserNumber &"'")
	strShowErr = "<li>对外开放资料开启</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../main.asp")
	Response.end
Elseif straction = "Close" then
	User_Conn.execute("update FS_ME_Users set isOpen= 0 where UserNumber='"& Fs_User.UserNumber &"'")
	strShowErr = "<li>对外开放资料取消</li>"
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
'---------------------------------分页定义
int_RPP=15 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"			'尾页
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
'排序状态
Function changeOrderStatus()
	Dim order_desc
	if session("orderstatus")="asc" then
		session("orderstatus")="desc"
		order_desc="<font color='red'>↑</font>"
	else
		session("orderstatus")="asc"
		order_desc="<font color='red'>↓</font>"
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
<title>User Manage Center-网站内容管理系统</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
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
          <td align="center"><a href="houseManage.asp" style="color:#FF0000"  class="sd">楼盘信息</a></td>
          <td align="center"><a href="HS_Tenancy.asp"  class="sd">租赁信息</a></td>
          <td align="center"><a href="HS_Second.asp"  class="sd">二手房信息</a></td>
        <tr>
          <td colspan="3">
		  <div id="container">
              <!--ajax容器-->
              <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
                
                <tr>
                  <td align="center" <%if mystatus=1 or mystatus="" then response.Write("class='xingmu'") else response.Write("class='hback'")%>><a href="houseManage.asp?status=1&audit=<%=session("audit")%>" class="sd">形象展示楼盘</a></td>
                  <td align="center" <%if mystatus=2 then response.Write("class='xingmu'") else response.Write("class='hback'")%>><a href="houseManage.asp?status=2&audit=<%=session("audit")%>" class="sd">期房楼盘信息</a></td>
                  <td align="center" <%if mystatus=3 then response.Write("class='xingmu'") else response.Write("class='hback'")%>><a href="houseManage.asp?status=3&audit=<%=session("audit")%>" class="sd">现房楼盘信息信息</a></td>
                  <td align="center" class="hback"><a href="#" class="sd" title="选中复选框，查看已经审核信息" onClick="linkShowaudit();showaudit();">已审核信息</a><input type="checkbox" name="myaudit" id="audit" value="audit" onClick="showaudit()" <%if session("audit")="1" then response.Write("checked")%>/></td>
				<tr>
              </table>
              <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
                <tr class="hback_1">
                  <td class="hback_1" align="center" width="30%"><a href="houseManage.asp?order=orderbyname&status=<%=mystatus%>&audit=<%=session("audit")%>" class="sd">楼盘名称<%=HouseName_order%></a></td>
                  <td class="hback_1" align="center" width="25%"><a href="houseManage.asp?order=orderbyclass&status=<%=mystatus%>&audit=<%=session("audit")%>" class="sd">项目类别<%=Class_order%></a></td>
                  <td class="hback_1" align="center" width="15%"><a href="houseManage.asp?order=orderbytime&status=<%=mystatus%>&audit=<%=session("audit")%>" class="sd">开盘日期<%=OpenDate_order%></a></td>
                  <td class="hback_1" align="center" width="20%"><a href="#" class="sd">操作</a></td>
                  <td class="hback_1" align="center" width="5%"><input type="checkbox" value="" name="quotationlist"  onclick="selectAll(document.all('quotationlist'))"/></td>
                </tr>
                <%
					Set house_rs=Server.CreateObject(G_FS_RS)
					house_rs.open sqlstatement,Conn,1,1
					If Not house_rs.eof then
					'分页使用-----------------------------------
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
							audit="<font color='red'>未审核</font>"
						else
							audit="已审核"
						End if
						Response.Write("<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>"&vbcrlf)
						response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('detail"&house_rs("ID")&"','_detail"&house_rs("ID")&"')""><a href='HS_Quotation_Edit.asp?action=edit&id="&house_rs("ID")&"'>"&house_rs("HouseName")&"</a></td>"&vbcrlf)
						response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('detail"&house_rs("ID")&"','_detail"&house_rs("ID")&"')"">"&house_rs("Class")&"</td>"&vbcrlf)
						response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('detail"&house_rs("ID")&"','_detail"&house_rs("ID")&"')"">"&house_rs("OpenDate")&"</td>"&vbcrlf)
						response.Write("<td align='center' class='hback'><span id='audit"&house_rs("ID")&"'>"&audit&"</span>|<span id='edit"&house_rs("ID")&"'><a href='HS_Quotation_Edit.asp?action=edit&id="&house_rs("ID")&"'>修改</a></span>|<span id='delete"&house_rs("ID")&"'><a href='#' class='sd' onclick=""deleteAction('delete"&house_rs("ID")&"','"&house_rs("ID")&"')"">删除</a></span></td>"&vbcrlf)
						response.Write("<td align='center' class='hback'><input type='checkbox' name='quotationlist' value='"&house_rs("ID")&"'</td>"&vbcrlf)
						response.Write("</tr>"&vbcrlf)
						Response.Write("<tr id=""detail"&house_rs("ID")&""" style=""display:none"">"&vbcrlf)
						response.Write("<td colspan='7' class='hback'>"&vbcrlf)
						Response.Write("├位置："&house_rs("Position")&"|方向："&house_rs("Direction")&vbcrlf)
						response.Write("</td>"&vbcrlf)
						response.Write("</tr>"&vbcrlf)
						response.Write("<tr id=""_detail"&house_rs("ID")&""" style=""display:none"">"&vbcrlf)
						response.Write("<td colspan='7' class='hback'>"&vbcrlf)
						response.Write("├预售许可证：["&house_rs("PreSaleNumber")&"]|报价：￥"&house_rs("Price")&"万|联系电话："&house_rs("Tel")&vbcrlf)
						response.Write("</td>"&vbcrlf)
						Response.Write("</tr>"&vbcrlf)
						response.Write("</div>"&vbcrlf)
						house_rs.movenext
					next
	
				%>
                <tr>
                  <td colspan="5" align="right" class="hback"><button onClick="javascript:location='HS_Quotation_Edit.asp?action=add'">发布信息</button>
                    <button onClick="deleteBatAction()">批量删除</button>
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
//操作
//删除
function deleteAction(container,id)
{
	var url="HS_Quotation_Open_Action.asp?rad="+Math.random();
	var param="action=delete"+"&id="+id;
	if(confirm("确认要删除该条信息？"))
		var myAjax = new Ajax.Request(url,{method: 'get', parameters: param, onComplete: showResponse});
	function showResponse(originalRequest)
	{
		var value= originalRequest.responseText;
		if(value=="ok")
		{
			$(container).parentNode.parentNode.style.display='none';
			$('recordcount').innerHTML=parseInt($('recordcount').innerHTML)-1
			alert('删除成功！')
		}
		else
		{
			alert("发生异常，请与客服人员联系！");
		}
	}
}
//批量删除
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
		alert("请选择要删除的记录！");
		return;
	}
	param="action=delete"+"&id="+id;
	if(confirm("确认要删除该条信息？"))
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
			alert(count+'条信息删除成功！')
			count=0;
		}
		else
		{
			alert("发生异常，请与客服人员联系！");
			alert(originalRequest.responseText);
		}
	}
}
</script>
</html>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->







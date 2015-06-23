<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"

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
Dim tenancy_rs,mytype,order,sqlstatement,order_status,tClass,HouseStyleArray,HouseStyle,Decoration,UseFor,Floor,equip
Dim id_order,usefor_order,Class_order,price_order,time_order
session("audit")=request.querystring("audit")
order=request.QueryString("order")
mytype=trim(request.QueryString("type"))
if mytype="" then
	mytype=1
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
select case order
	case "orderbyid" 
		if session("audit")="1" then 
			sqlstatement="select tid,UseFor,class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate,Audited,UserNumber  from FS_HS_Tenancy where UserNumber='"&session("FS_UserNumber")&"' and Audited=1 and order by tid "&session("orderstatus")
		else 
			sqlstatement="select tid,UseFor,class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate,Audited,UserNumber  from FS_HS_Tenancy where UserNumber='"&session("FS_UserNumber")&"' order by tid "&session("orderstatus")
		end if
		
	case "orderbyusefor" 
		if session("audit")="1" then 
			sqlstatement="select tid,UseFor,class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate,Audited,UserNumber  from FS_HS_Tenancy where UserNumber='"&session("FS_UserNumber")&"' and Audited=1 order by UseFor "&session("orderstatus")
		else
			sqlstatement="select tid,UseFor,class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate,Audited,UserNumber  from FS_HS_Tenancy  where UserNumber='"&session("FS_UserNumber")&"' order by UseFor "&session("orderstatus")
		end if
		
	case "orderbyclass" 
		if session("audit")="1" then 
			sqlstatement="select tid,UseFor,class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate,Audited,UserNumber  from FS_HS_Tenancy where UserNumber='"&session("FS_UserNumber")&"' and Audited=1 order by Class "&session("orderstatus")
		else
			sqlstatement="select tid,UseFor,class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate,Audited,UserNumber  from FS_HS_Tenancy where UserNumber='"&session("FS_UserNumber")&"' order by Class "&session("orderstatus")
		end if
		
	case "orderbyprice" 
		if session("audit")="1" then
			sqlstatement="select tid,UseFor,class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate,Audited,UserNumber  from FS_HS_Tenancy where Audited=1 order by Price "&session("orderstatus")
		else
			sqlstatement="select tid,UseFor,class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate,Audited,UserNumber  from FS_HS_Tenancy where UserNumber='"&session("FS_UserNumber")&"' order by Price "&session("orderstatus")
		End if
		
	case "orderbytime" 
		if session("audit")="1" then
			sqlstatement="select tid,UseFor,class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate,Audited,UserNumber  from FS_HS_Tenancy where UserNumber='"&session("FS_UserNumber")&"' and Audited=1 order by PubDate "&session("orderstatus")
		else
			sqlstatement="select tid,UseFor,class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate,Audited,UserNumber  from FS_HS_Tenancy where UserNumber='"&session("FS_UserNumber")&"' order by PubDate "&session("orderstatus")
		End if
	
	case else 
		if session("audit")="1" then
			sqlstatement="select tid,UseFor,class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate,Audited,UserNumber  from FS_HS_Tenancy where UserNumber='"&session("FS_UserNumber")&"' and Audited=1 order by tid desc"
		else
			sqlstatement="select tid,UseFor,class,Position,CityArea,Price,HouseStyle,Area,Floor,BuildDate,equip,Decoration,LinkMan,Contact,Period,Remark,PubDate,Audited,UserNumber  from FS_HS_Tenancy where UserNumber='"&session("FS_UserNumber")&"' order by tid desc"
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
		case "orderbyid"    id_order=order_desc
							usefor_order=""
							Class_order=""
							price_order=""
							time_order=""
		
		case "orderbyusefor" usefor_order=order_desc
							 id_order=""
							 class_order=""
							 price_order=""
							 time_order=""	
											
		case "orderbyclass" Class_order=order_desc
							id_order=""
							time_order=""
							usefor_order=""
							price_order=""
							
		case "orderbyprice" price_order=order_desc
							id_order=""
							Class_order=""
							time_order=""
							usefor_order=""
							
		case "orderbytime"  time_order=order_desc
							id_order=""
							Class_order=""
							usefor_order=""
							price_order=""
	end select
End Function
if trim(order)<>"" then
	changeOrderStatus()
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%></title>
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
    <td width="82%" valign="top" class="hback"><table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
        <tr>
          <td class="xingmu" align="center" <%if not isCorp() then Response.Write("style=""display:none""")%>><a href="houseManage.asp">楼盘信息</a></td>
          <td class="xingmu" align="center"><a href="HS_Tenancy.asp" style="color:#FF0000">租赁信息</a></td>
          <td class="xingmu" align="center"><a href="HS_Second.asp" >二手房信息</a></td>
        <tr>
          <td colspan="3"><table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
              <tr>
                <td class="xingmu" colspan="4">租赁信息 <a href="#" onClick="history.back()">后退</a> <a href="#" class="sd" title="选中复选框，查看已审核信息" onClick="linkShowaudit();showaudit();">已审核信息</a>
                  <input type="checkbox" name="myaudit" id="audit" value="" onClick="showaudit()" <%if session("audit")="1" then response.Write("checked")%>/></td>
              </tr>
            </table>
            <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
              <tr>
                <td class="hback_1" align="center" width="10%"><a href="HS_Tenancy.asp?order=orderbyid&type=<%=mytype%>&audit=<%=session("audit")%>" class="sd">编号<%=id_order%></a></td>
                <td class="hback_1" align="center" width="20%"><a href="HS_Tenancy.asp?order=orderbyusefor&type=<%=mytype%>&audit=<%=session("audit")%>" class="sd">用途<%=usefor_order%></a></td>
                <td class="hback_1" align="center" width="15%"><a href="HS_Tenancy.asp?order=orderbyclass&type=<%=mytype%>&audit=<%=session("audit")%>" class="sd">类型<%=class_order%></a></td>
                <td class="hback_1" align="center" width="15%"><a href="HS_Tenancy.asp?order=orderbyprice&type=<%=mytype%>&audit=<%=session("audit")%>" class="sd">价格<%=price_order%></a></td>
                <td class="hback_1" align="center" width="15%"><a href="HS_Tenancy.asp?order=orderbytime&type=<%=mytype%>&audit=<%=session("audit")%>" class="sd">发布时间<%=time_order%></a></td>
                <td class="hback_1" align="center" width="20%"><a href="#" class="sd">操作</a></td>
                <td class="hback_1" align="center" width="5%"><input type="checkbox" value="" name="tenancylist"  onclick="selectAll(document.all('tenancylist'))"/></td>
              </tr>
              <%
	Set tenancy_rs=Server.CreateObject(G_FS_RS)
	tenancy_rs.open sqlstatement,Conn,1,1
	If Not tenancy_rs.eof then
	'分页使用-----------------------------------
		tenancy_rs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo<=0 Then cPageNo=1
		If cPageNo>tenancy_rs.PageCount Then cPageNo=tenancy_rs.PageCount 
		tenancy_rs.AbsolutePage=cPageNo
	End if
	Dim audit
	for i=0 to int_RPP
		if tenancy_rs.eof then exit for
		select case tenancy_rs("Class")
		case "1"  tClass="出租"
		case "2"  tClass="求租"
		case "3"  tClass="出售"
		case "4"  tClass="求购"
		case "5"  tClass="合租"
		case "6"  tClass="转让" 
		end select
		if tenancy_rs("UseFor")="1" then
			UseFor="住房"
		else
			UseFor="写字间"
		end if
		if tenancy_rs("audited")="1" then
			audit="已审核"
		else
			audit="<font color='red'>未审核</font>"
		End if
		Response.Write("<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('1_detail"&tenancy_rs("TID")&"','2_detail"&tenancy_rs("TID")&"','3_detail"&tenancy_rs("TID")&"','4_detail"&tenancy_rs("TID")&"')"">"&tenancy_rs("TID")&"</td>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('1_detail"&tenancy_rs("TID")&"','2_detail"&tenancy_rs("TID")&"','3_detail"&tenancy_rs("TID")&"','4_detail"&tenancy_rs("TID")&"')"">"&UseFor&"</td>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('1_detail"&tenancy_rs("TID")&"','2_detail"&tenancy_rs("TID")&"','3_detail"&tenancy_rs("TID")&"','4_detail"&tenancy_rs("TID")&"')"">"&tClass&"</td>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('1_detail"&tenancy_rs("TID")&"','2_detail"&tenancy_rs("TID")&"','3_detail"&tenancy_rs("TID")&"','4_detail"&tenancy_rs("TID")&"')"">￥"&tenancy_rs("price")&"元/月</td>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('1_detail"&tenancy_rs("TID")&"','2_detail"&tenancy_rs("TID")&"','3_detail"&tenancy_rs("TID")&"','4_detail"&tenancy_rs("TID")&"')"">"&tenancy_rs("PubDate")&"</td>"&vbcrlf)
		response.Write("<td align='center' class='hback'><span id='audit"&tenancy_rs("TID")&"'>"&audit&"</span>&nbsp;|&nbsp;<span id='edit"&tenancy_rs("TID")&"'><a href='HS_Tenancy_Edit.asp?action=edit&id="&tenancy_rs("TID")&"'>修改</a></span>&nbsp;|&nbsp;<span id='delete"&tenancy_rs("TID")&"'><a href='#' class='sd' onclick=""deleteAction('delete"&tenancy_rs("TID")&"','"&tenancy_rs("TID")&"')"">删除</a></span></td>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown='document.body.onmousedown=null'><input type='checkbox' name='tenancylist' value='"&tenancy_rs("TID")&"'</td>"&vbcrlf)
		response.Write("</tr>"&vbcrlf)
		Response.Write("<tr id=""1_detail"&tenancy_rs("TID")&""" style=""display:none"">"&vbcrlf)
		response.Write("<td colspan='7' class='hback'>"&vbcrlf)
		Response.Write("├位置："&tenancy_rs("Position")&"&nbsp;|&nbsp;区县："&tenancy_rs("CityArea")&vbcrlf)
		response.Write("</td>"&vbcrlf)
		response.Write("</tr>"&vbcrlf)
		'户型处理
		if trim(tenancy_rs("HouseStyle"))<>"" then
			HouseStyleArray=split(tenancy_rs("HouseStyle"),",")
			HouseStyle=HouseStyleArray(0)&"室"&HouseStyleArray(1)&"厅"&HouseStyleArray(2)&"卫"
		else
			HouseStyle="暂无信息"
		End if
		'楼层处理
		if trim(tenancy_rs("Floor"))<>"" then
			Floor="总"&split(tenancy_rs("Floor"),",")(0)&"层，第"&split(tenancy_rs("Floor"),",")(1)&"层"
		else
			Floor="暂无信息"
		end if
		response.Write("<tr id=""2_detail"&tenancy_rs("TID")&""" style=""display:none"">"&vbcrlf)
		response.Write("<td colspan='7' class='hback'>"&vbcrlf)
		Response.Write("├户型："&HouseStyle&"&nbsp;|&nbsp;面积："&tenancy_rs("Area")&"平米&nbsp;|&nbsp;楼层："&Floor&vbcrlf)
		response.Write("</td>"&vbcrlf)
		Response.Write("</tr>"&vbcrlf)
		'处理装修情况
		select case tenancy_rs("Decoration")
			case "1" Decoration="简单装修"
			case "2" Decoration="中档装修"
			case "3" Decoration="高档装修"
			case "4" Decoration="没有装修"
			case else  Decoration="暂无信息"
		end select
		'处理设施
		if tenancy_rs("equip")<>"" then
			equip=replace(replace(replace(replace(replace(replace(tenancy_rs("equip"),"l","通水"),"m","电"),"n","气"),"x","电话"),"y","光纤"),"z","宽带")
		else
			equip="无相关信息"
		End if
		response.Write("<tr id=""3_detail"&tenancy_rs("TID")&""" style=""display:none"">"&vbcrlf)
		response.Write("<td colspan='7' class='hback'>"&vbcrlf)
		Response.Write("├配套设施："&equip&"&nbsp;|&nbsp;装修情况："&Decoration&vbcrlf)
		response.Write("</td>"&vbcrlf)
		Response.Write("</tr>"&vbcrlf)

		response.Write("<tr id=""4_detail"&tenancy_rs("TID")&""" style=""display:none"">"&vbcrlf)
		response.Write("<td colspan='7' class='hback'>"&vbcrlf)
		response.Write("├联系人："&tenancy_rs("LinkMan")&"&nbsp;|&nbsp;联系方式："&tenancy_rs("Contact")&vbcrlf)
		response.Write("</td>"&vbcrlf)
		Response.Write("</tr>"&vbcrlf)
		response.Write("</div>"&vbcrlf)
		tenancy_rs.movenext
	next
%>
              <tr>
                <td colspan="7" align="right" class="hback"><button onClick="javascript:location='HS_Tenancy_Edit.asp?action=add'">发布信息</button>
                  <button onClick="deleteBatAction()">批量删除</button></td>
              </tr>
              <%
	Response.Write("<tr>"&vbcrlf)
	Response.Write("<td align='right' colspan='7'  class=""hback"">"&fPageCount(tenancy_rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td>"&vbcrlf)
	Response.Write("</tr>"&vbcrlf)
%>
            </table></td>
        </tr>
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
tenancy_rs.close
Set tenancy_rs=nothing
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
		location="HS_Tenancy.asp?audit=1";
	}else
	{
		location="HS_Tenancy.asp?audit=0";
	}
}
//审核与取消审核
function auditBatAction(type)
{
	if (type=="audit")
	{
		var url="Hs_Tenancy_Action.asp?rad="+Math.random();
		var elements=document.all('tenancylist');
		var id="";
		var count=0;
		for(var i=1;i<elements.length;i++)
		{
			if(elements[i].checked)
			{
				id=elements[i].value+","+id;
				if(elements[i].parentNode.parentNode.style.display!="none")
					count+=1;
			}
		}
		if(id.length<2)
		{
			alert("请选择审核的记录！");
			return;
		}
		param="action=Audit"+"&id="+id;
		if(confirm("确认要审核该条信息？"))
			var myAjax = new Ajax.Request(url,{method: 'get', parameters: param, onComplete: showResponse1});
		function showResponse1(originalRequest)
		{
			var value= originalRequest.responseText;
			if(value=="ok")
			{
				$('recordcount').innerHTML=parseInt($('recordcount').innerHTML)-count;
				alert(count+'条信息审核成功！')
				count=0;
				location="HS_Tenancy.asp";
			}
			else
			{
				alert("发生异常，请与客服人员联系！");
			}
		}
	}
	else
	{
		var url="Hs_Tenancy_Action.asp?rad="+Math.random();
		var elements=document.all('tenancylist');
		var id="";
		var count=0;
		for(var i=1;i<elements.length;i++)
		{
			if(elements[i].checked)
			{
				id=elements[i].value+","+id;
				if(elements[i].parentNode.parentNode.style.display!="none")
					count+=1;
			}
		}
		if(id.length<2)
		{
			alert("请选择要取消审核的记录！");
			return;
		}
		param="action=UnAudit"+"&id="+id;
		if(confirm("确认要取消审核该条信息？"))
			var myAjax = new Ajax.Request(url,{method: 'get', parameters: param, onComplete: showResponse});
		function showResponse(originalRequest)
		{
			var value= originalRequest.responseText;
			if(value=="ok")
			{
				$('recordcount').innerHTML=parseInt($('recordcount').innerHTML)-count;
				alert(count+'条信息取消审核成功！')
				count=0;
				location="HS_Tenancy.asp";
			}
			else
			{
				alert("发生异常，请与客服人员联系！");
			}
		}
	}
}
//操作
//删除
function deleteAction(container,id)
{
	var url="Hs_Tenancy_Action.asp?rad="+Math.random();
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
	var url="Hs_Tenancy_Action.asp?rad="+Math.random();
	var elements=document.all('tenancylist');
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
		}
	}
}
</script>
</html>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->







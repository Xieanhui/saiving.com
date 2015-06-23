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
Dim Second_rs,mytype,order,sqlstatement,order_status,tClass,HouseStyleArray,HouseStyle,Decoration,UseFor,Floor,FloorType,equip
Dim BelongType,Structure
Dim id_order,usefor_order,Class_order,price_order,time_order
session("audit") = NoSqlHack(request.querystring("audit"))
order = NoSqlHack(request.QueryString("order"))
mytype = NoSqlHack(request.QueryString("type"))
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
		if session("audit")="1"  then 
			sqlstatement="select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark,PubDate,Audited,PicNumber from FS_HS_Second where  Usernumber='"&session("FS_UserNumber")&"' and Audited=1 order by sid "&session("orderstatus")
		else 
			sqlstatement="select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark,PubDate,Audited,PicNumber from FS_HS_Second where Usernumber='"&session("FS_UserNumber")&"' order by sid "&session("orderstatus")
		end if
		
	case "orderbyusefor" 
		if session("audit")="1"  then 
			sqlstatement="select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark,PubDate,Audited,PicNumber from FS_HS_Second where Usernumber='"&session("FS_UserNumber")&"' and Audited=1 order by UseFor "&session("orderstatus")
		else
			sqlstatement="select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark,PubDate,Audited,PicNumber from FS_HS_Second  where Usernumber='"&session("FS_UserNumber")&"' order by UseFor "&session("orderstatus")
		end if
		
	case "orderbyclass" 
		if session("audit")="1"  then 
			sqlstatement="select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark,PubDate,Audited,PicNumber from FS_HS_Second where Usernumber='"&session("FS_UserNumber")&"' and Audited=1 order by Class "&session("orderstatus")
		else
			sqlstatement="select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark,PubDate,Audited,PicNumber from FS_HS_Second where Usernumber='"&session("FS_UserNumber")&"' order by Class "&session("orderstatus")
		end if
		
	case "orderbyprice" 
		if session("audit")="1"  then
			sqlstatement="select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark,PubDate,Audited,PicNumber from FS_HS_Second where Usernumber='"&session("FS_UserNumber")&"' and Audited=1 order by Price "&session("orderstatus")
		else
			sqlstatement="select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark,PubDate,Audited,PicNumber from FS_HS_Second where Usernumber='"&session("FS_UserNumber")&"' order by Price "&session("orderstatus")
		End if
		
	case "orderbytime" 
		if session("audit")="1"  then
			sqlstatement="select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark,PubDate,Audited,PicNumber from FS_HS_Second where Usernumber='"&session("FS_UserNumber")&"' and Audited=1 order by PubDate "&session("orderstatus")
		else
			sqlstatement="select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark,PubDate,Audited,PicNumber from FS_HS_Second where Usernumber='"&session("FS_UserNumber")&"' order by PubDate "&session("orderstatus")
		End if
	
	case else 
		if session("audit")="1"  then
			sqlstatement="select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark,PubDate,Audited,PicNumber from FS_HS_Second where Usernumber='"&session("FS_UserNumber")&"' and Audited=1 order by sid desc"
		else
			sqlstatement="select SID,Class,UserNumber,Label,UseFor,FloorType,BelongType,HouseStyle,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark,PubDate,Audited,PicNumber from FS_HS_Second  where Usernumber='"&session("FS_UserNumber")&"' order by sid desc"
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
End if%>
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
          <td class="xingmu" align="center"><a href="HS_Tenancy.asp">租赁信息</a></td>
          <td class="xingmu" align="center"><a href="HS_Second.asp" style="color:#FF0000">二手房信息</a></td>
        <tr>
          <td colspan="3"><table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
              <tr>
                <td class="xingmu" colspan="4">二手房信息&nbsp;|&nbsp;<a href="#" onClick="history.back()">后退</a>&nbsp;|&nbsp;<a href="#" class="sd" title="选中复选框，查看个人二手房" onClick="linkShowclass();showclass();">已审核</a>
                  <input type="checkbox" name="ifaudit" id="audit" value="" onClick="showclass()" <%if session("audit")="1"  then response.Write("checked")%>/></td>
              </tr>
            </table>
            <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
              <tr>
                <td class="hback_1" align="center" width="10%"><a href="HS_Second.asp?order=orderbyid&type=<%=mytype%>&class=<%=session("class")%>" class="sd">编号<%=id_order%></a></td>
                <td class="hback_1" align="center" width="20%"><a href="HS_Second.asp?order=orderbyusefor&type=<%=mytype%>&audit=<%=session("audit")%>" class="sd">用途<%=usefor_order%></a></td>
                <td class="hback_1" align="center" width="15%"><a href="HS_Second.asp?order=orderbyclass&type=<%=mytype%>&audit=<%=session("audit")%>" class="sd">类型<%=class_order%></a></td>
                <td class="hback_1" align="center" width="15%"><a href="HS_Second.asp?order=orderbyprice&type=<%=mytype%>&audit=<%=session("audit")%>" class="sd">价格<%=price_order%></a></td>
                <td class="hback_1" align="center" width="15%"><a href="HS_Second.asp?order=orderbytime&type=<%=mytype%>&audit=<%=session("audit")%>" class="sd">发布时间<%=time_order%></a></td>
                <td class="hback_1" align="center" width="20%"><a href="#" class="sd">操作</a></td>
                <td class="hback_1" align="center" width="5%"><input type="checkbox" value="" name="Secondlist"  onclick="selectAll(document.all('Secondlist'))"/></td>
              </tr>
              <%
	Set Second_rs=Server.CreateObject(G_FS_RS)
	Second_rs.open sqlstatement,Conn,1,1
	If Not Second_rs.eof then
	'分页使用-----------------------------------
		Second_rs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo<=0 Then cPageNo=1
		If cPageNo>Second_rs.PageCount Then cPageNo=Second_rs.PageCount 
		Second_rs.AbsolutePage=cPageNo
	End if
	Dim audit
	for i=0 to int_RPP
		if Second_rs.eof then exit for
		select case Second_rs("Class")
			case "1"  tClass="个人"
			case "2"  tClass="中介"
		end select
		if Second_rs("UseFor")="1" then
			UseFor="住房"
		elseif Second_rs("UseFor")="2" then
			UseFor="写字间"
		Else
			UseFor="其他"
		end if
		if Second_rs("audited")="1" then
			audit="已审核"
		else
			audit="<font color='red'>未审核</font>"
		End if
		Response.Write("<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('1_detail"&Second_rs("sid")&"','2_detail"&Second_rs("sid")&"','3_detail"&Second_rs("sid")&"','4_detail"&Second_rs("sid")&"')"">"&Second_rs("sid")&"</td>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('1_detail"&Second_rs("sid")&"','2_detail"&Second_rs("sid")&"','3_detail"&Second_rs("sid")&"','4_detail"&Second_rs("sid")&"')"">"&UseFor&"</td>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('1_detail"&Second_rs("sid")&"','2_detail"&Second_rs("sid")&"','3_detail"&Second_rs("sid")&"','4_detail"&Second_rs("sid")&"')"">"&tClass&"</td>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('1_detail"&Second_rs("sid")&"','2_detail"&Second_rs("sid")&"','3_detail"&Second_rs("sid")&"','4_detail"&Second_rs("sid")&"')"">￥"&Second_rs("price")&"万</td>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('1_detail"&Second_rs("sid")&"','2_detail"&Second_rs("sid")&"','3_detail"&Second_rs("sid")&"','4_detail"&Second_rs("sid")&"')"">"&Second_rs("PubDate")&"</td>"&vbcrlf)
		response.Write("<td align='center' class='hback'><span id='audit"&Second_rs("sid")&"'>"&audit&"</span>&nbsp;|&nbsp;<span id='edit"&Second_rs("sid")&"'><a href='HS_Second_Edit.asp?action=edit&id="&Second_rs("sid")&"'>修改</a></span>&nbsp;|&nbsp;<span id='delete"&Second_rs("sid")&"'><a href='#' class='sd' onclick=""deleteAction('delete"&Second_rs("sid")&"','"&Second_rs("sid")&"')"">删除</a></span></td>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown='document.body.onmousedown=null'><input type='checkbox' name='Secondlist' value='"&Second_rs("sid")&"'</td>"&vbcrlf)
		response.Write("</tr>"&vbcrlf)
		Response.Write("<tr id=""1_detail"&Second_rs("sid")&""" style=""display:none"">"&vbcrlf)
		response.Write("<td colspan='7' class='hback'>"&vbcrlf)
		Response.Write("├位置："&Second_rs("Position")&"&nbsp;|&nbsp;区县："&Second_rs("CityArea")&vbcrlf)
		response.Write("</td>"&vbcrlf)
		response.Write("</tr>"&vbcrlf)
		'户型处理
		if trim(Second_rs("HouseStyle"))<>"" then
			HouseStyleArray=split(Second_rs("HouseStyle"),",")
			HouseStyle=HouseStyleArray(0)&"室"&HouseStyleArray(1)&"厅"&HouseStyleArray(2)&"卫"
		else
			HouseStyle="暂无信息"
		End if
		'楼层处理
		if trim(Second_rs("Floor"))<>"" then
			Floor="总"&split(Second_rs("Floor"),",")(0)&"层，第"&split(Second_rs("Floor"),",")(1)&"层"
		else
			Floor="暂无信息"
		end if
		'楼层类型
		if Second_rs("FloorType")="1" then 
			FloorType="多层"
		Else
			FloorType="单层"
		End if
		'处理装修情况
		select case Second_rs("Decoration")
			case "1" Decoration="简单装修"
			case "2" Decoration="中档装修"
			case "3" Decoration="高档装修"
			case "4" Decoration="没有装修"
			case else  Decoration="暂无信息"
		end select
		'处理设施
		if Second_rs("equip")<>"" then
			equip=replace(replace(replace(replace(replace(replace(Second_rs("equip"),"l","通水"),"m","电"),"n","气"),"x","电话"),"y","光纤"),"z","宽带")
		else
			equip="无相关信息"
		End if
		'产权性质
		if Second_rs("BelongType")="1" then
			BelongType="商品房"
		else
			BelongType="集资房"
		End if
		'1钢筋混凝土结构、2砖混结构、3砖木结构、4简易结构
		select case Second_rs("Structure")
			case "1" Structure="钢筋混凝土"
			case "2" Structure="砖混"
			case "3" Structure="砖木"
			case "4" Structure="简易"
		End select
		response.Write("<tr id=""2_detail"&Second_rs("sid")&""" style=""display:none"">"&vbcrlf)
		response.Write("<td colspan='7' class='hback'>"&vbcrlf)
		Response.Write("├户型："&HouseStyle&"&nbsp;|&nbsp;性质："&BelongType&"&nbsp;|&nbsp;结构："&Structure&"&nbsp;|&nbsp;面积："&Second_rs("Area")&"平米&nbsp;|&nbsp;楼层信息：["&FloorType&"]"&Floor&vbcrlf)
		response.Write("</td>"&vbcrlf)
		Response.Write("</tr>"&vbcrlf)
		response.Write("<tr id=""3_detail"&Second_rs("sid")&""" style=""display:none"">"&vbcrlf)
		response.Write("<td colspan='7' class='hback'>"&vbcrlf)
		Response.Write("├配套设施："&equip&"&nbsp;|&nbsp;装修情况："&Decoration&"&nbsp;|&nbsp;建筑年代："&Second_rs("BuildDate")&vbcrlf)
		response.Write("</td>"&vbcrlf)
		Response.Write("</tr>"&vbcrlf)
		response.Write("<tr id=""4_detail"&Second_rs("sid")&""" style=""display:none"">"&vbcrlf)
		response.Write("<td colspan='7' class='hback'>"&vbcrlf)
		response.Write("├联系人："&Second_rs("LinkMan")&"&nbsp;|&nbsp;联系方式："&Second_rs("Contact")&vbcrlf)
		response.Write("</td>"&vbcrlf)
		Response.Write("</tr>"&vbcrlf)
		response.Write("</div>"&vbcrlf)
		Second_rs.movenext
	next
%>
              <tr>
                <td colspan="7" align="right" class="hback"><button onClick="javascript:location='HS_Second_Edit.asp?action=add'">发布信息</button>
                  <button onClick="deleteBatAction()">批量删除</button></td>
              </tr>
              <%
	Response.Write("<tr>"&vbcrlf)
	Response.Write("<td align='right' colspan='7'  class=""hback"">"&fPageCount(Second_rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td>"&vbcrlf)
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
Second_rs.close
Set Second_rs=nothing
%>
<script language="javascript">
function linkShowclass()
{
	if($('audit').checked)
	{
		$('audit').checked=false;
	}else
	{
		$('audit').checked=true;
	}

}
function showclass()
{
	if($('audit').checked)
	{
		location="HS_Second.asp?audit=1&type=<%=mytype%>";
	}
	else
	{
		location="HS_Second.asp?audit=&type=<%=mytype%>";
	}
}
//删除
function deleteAction(container,id)
{
	var url="HS_Second_Action.asp?rad="+Math.random();
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
	var url="HS_Second_Action.asp?rad="+Math.random();
	var elements=document.all('Secondlist');
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
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->







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
    <td width="82%" valign="top" class="hback"><table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
        <tr>
          <td class="xingmu" align="center" <%if not isCorp() then Response.Write("style=""display:none""")%>><a href="houseManage.asp">¥����Ϣ</a></td>
          <td class="xingmu" align="center"><a href="HS_Tenancy.asp">������Ϣ</a></td>
          <td class="xingmu" align="center"><a href="HS_Second.asp" style="color:#FF0000">���ַ���Ϣ</a></td>
        <tr>
          <td colspan="3"><table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
              <tr>
                <td class="xingmu" colspan="4">���ַ���Ϣ&nbsp;|&nbsp;<a href="#" onClick="history.back()">����</a>&nbsp;|&nbsp;<a href="#" class="sd" title="ѡ�и�ѡ�򣬲鿴���˶��ַ�" onClick="linkShowclass();showclass();">�����</a>
                  <input type="checkbox" name="ifaudit" id="audit" value="" onClick="showclass()" <%if session("audit")="1"  then response.Write("checked")%>/></td>
              </tr>
            </table>
            <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
              <tr>
                <td class="hback_1" align="center" width="10%"><a href="HS_Second.asp?order=orderbyid&type=<%=mytype%>&class=<%=session("class")%>" class="sd">���<%=id_order%></a></td>
                <td class="hback_1" align="center" width="20%"><a href="HS_Second.asp?order=orderbyusefor&type=<%=mytype%>&audit=<%=session("audit")%>" class="sd">��;<%=usefor_order%></a></td>
                <td class="hback_1" align="center" width="15%"><a href="HS_Second.asp?order=orderbyclass&type=<%=mytype%>&audit=<%=session("audit")%>" class="sd">����<%=class_order%></a></td>
                <td class="hback_1" align="center" width="15%"><a href="HS_Second.asp?order=orderbyprice&type=<%=mytype%>&audit=<%=session("audit")%>" class="sd">�۸�<%=price_order%></a></td>
                <td class="hback_1" align="center" width="15%"><a href="HS_Second.asp?order=orderbytime&type=<%=mytype%>&audit=<%=session("audit")%>" class="sd">����ʱ��<%=time_order%></a></td>
                <td class="hback_1" align="center" width="20%"><a href="#" class="sd">����</a></td>
                <td class="hback_1" align="center" width="5%"><input type="checkbox" value="" name="Secondlist"  onclick="selectAll(document.all('Secondlist'))"/></td>
              </tr>
              <%
	Set Second_rs=Server.CreateObject(G_FS_RS)
	Second_rs.open sqlstatement,Conn,1,1
	If Not Second_rs.eof then
	'��ҳʹ��-----------------------------------
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
			case "1"  tClass="����"
			case "2"  tClass="�н�"
		end select
		if Second_rs("UseFor")="1" then
			UseFor="ס��"
		elseif Second_rs("UseFor")="2" then
			UseFor="д�ּ�"
		Else
			UseFor="����"
		end if
		if Second_rs("audited")="1" then
			audit="�����"
		else
			audit="<font color='red'>δ���</font>"
		End if
		Response.Write("<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('1_detail"&Second_rs("sid")&"','2_detail"&Second_rs("sid")&"','3_detail"&Second_rs("sid")&"','4_detail"&Second_rs("sid")&"')"">"&Second_rs("sid")&"</td>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('1_detail"&Second_rs("sid")&"','2_detail"&Second_rs("sid")&"','3_detail"&Second_rs("sid")&"','4_detail"&Second_rs("sid")&"')"">"&UseFor&"</td>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('1_detail"&Second_rs("sid")&"','2_detail"&Second_rs("sid")&"','3_detail"&Second_rs("sid")&"','4_detail"&Second_rs("sid")&"')"">"&tClass&"</td>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('1_detail"&Second_rs("sid")&"','2_detail"&Second_rs("sid")&"','3_detail"&Second_rs("sid")&"','4_detail"&Second_rs("sid")&"')"">��"&Second_rs("price")&"��</td>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown=""Element.toggle('1_detail"&Second_rs("sid")&"','2_detail"&Second_rs("sid")&"','3_detail"&Second_rs("sid")&"','4_detail"&Second_rs("sid")&"')"">"&Second_rs("PubDate")&"</td>"&vbcrlf)
		response.Write("<td align='center' class='hback'><span id='audit"&Second_rs("sid")&"'>"&audit&"</span>&nbsp;|&nbsp;<span id='edit"&Second_rs("sid")&"'><a href='HS_Second_Edit.asp?action=edit&id="&Second_rs("sid")&"'>�޸�</a></span>&nbsp;|&nbsp;<span id='delete"&Second_rs("sid")&"'><a href='#' class='sd' onclick=""deleteAction('delete"&Second_rs("sid")&"','"&Second_rs("sid")&"')"">ɾ��</a></span></td>"&vbcrlf)
		response.Write("<td align='center' class='hback' onmousedown='document.body.onmousedown=null'><input type='checkbox' name='Secondlist' value='"&Second_rs("sid")&"'</td>"&vbcrlf)
		response.Write("</tr>"&vbcrlf)
		Response.Write("<tr id=""1_detail"&Second_rs("sid")&""" style=""display:none"">"&vbcrlf)
		response.Write("<td colspan='7' class='hback'>"&vbcrlf)
		Response.Write("��λ�ã�"&Second_rs("Position")&"&nbsp;|&nbsp;���أ�"&Second_rs("CityArea")&vbcrlf)
		response.Write("</td>"&vbcrlf)
		response.Write("</tr>"&vbcrlf)
		'���ʹ���
		if trim(Second_rs("HouseStyle"))<>"" then
			HouseStyleArray=split(Second_rs("HouseStyle"),",")
			HouseStyle=HouseStyleArray(0)&"��"&HouseStyleArray(1)&"��"&HouseStyleArray(2)&"��"
		else
			HouseStyle="������Ϣ"
		End if
		'¥�㴦��
		if trim(Second_rs("Floor"))<>"" then
			Floor="��"&split(Second_rs("Floor"),",")(0)&"�㣬��"&split(Second_rs("Floor"),",")(1)&"��"
		else
			Floor="������Ϣ"
		end if
		'¥������
		if Second_rs("FloorType")="1" then 
			FloorType="���"
		Else
			FloorType="����"
		End if
		'����װ�����
		select case Second_rs("Decoration")
			case "1" Decoration="��װ��"
			case "2" Decoration="�е�װ��"
			case "3" Decoration="�ߵ�װ��"
			case "4" Decoration="û��װ��"
			case else  Decoration="������Ϣ"
		end select
		'������ʩ
		if Second_rs("equip")<>"" then
			equip=replace(replace(replace(replace(replace(replace(Second_rs("equip"),"l","ͨˮ"),"m","��"),"n","��"),"x","�绰"),"y","����"),"z","����")
		else
			equip="�������Ϣ"
		End if
		'��Ȩ����
		if Second_rs("BelongType")="1" then
			BelongType="��Ʒ��"
		else
			BelongType="���ʷ�"
		End if
		'1�ֽ�������ṹ��2ש��ṹ��3שľ�ṹ��4���׽ṹ
		select case Second_rs("Structure")
			case "1" Structure="�ֽ������"
			case "2" Structure="ש��"
			case "3" Structure="שľ"
			case "4" Structure="����"
		End select
		response.Write("<tr id=""2_detail"&Second_rs("sid")&""" style=""display:none"">"&vbcrlf)
		response.Write("<td colspan='7' class='hback'>"&vbcrlf)
		Response.Write("�����ͣ�"&HouseStyle&"&nbsp;|&nbsp;���ʣ�"&BelongType&"&nbsp;|&nbsp;�ṹ��"&Structure&"&nbsp;|&nbsp;�����"&Second_rs("Area")&"ƽ��&nbsp;|&nbsp;¥����Ϣ��["&FloorType&"]"&Floor&vbcrlf)
		response.Write("</td>"&vbcrlf)
		Response.Write("</tr>"&vbcrlf)
		response.Write("<tr id=""3_detail"&Second_rs("sid")&""" style=""display:none"">"&vbcrlf)
		response.Write("<td colspan='7' class='hback'>"&vbcrlf)
		Response.Write("��������ʩ��"&equip&"&nbsp;|&nbsp;װ�������"&Decoration&"&nbsp;|&nbsp;���������"&Second_rs("BuildDate")&vbcrlf)
		response.Write("</td>"&vbcrlf)
		Response.Write("</tr>"&vbcrlf)
		response.Write("<tr id=""4_detail"&Second_rs("sid")&""" style=""display:none"">"&vbcrlf)
		response.Write("<td colspan='7' class='hback'>"&vbcrlf)
		response.Write("����ϵ�ˣ�"&Second_rs("LinkMan")&"&nbsp;|&nbsp;��ϵ��ʽ��"&Second_rs("Contact")&vbcrlf)
		response.Write("</td>"&vbcrlf)
		Response.Write("</tr>"&vbcrlf)
		response.Write("</div>"&vbcrlf)
		Second_rs.movenext
	next
%>
              <tr>
                <td colspan="7" align="right" class="hback"><button onClick="javascript:location='HS_Second_Edit.asp?action=add'">������Ϣ</button>
                  <button onClick="deleteBatAction()">����ɾ��</button></td>
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
//ɾ��
function deleteAction(container,id)
{
	var url="HS_Second_Action.asp?rad="+Math.random();
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
		}
	}
}
</script>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






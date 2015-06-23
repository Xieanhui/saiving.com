<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="lib/cls_house.asp"-->
<%
Dim straction,currenttype
currenttype=request.QueryString("currenttype")
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
Dim house_rs,action,str_CurrPath,houseObj,id,i
set houseObj=New cls_House
action=request.QueryString("action")
id=CintStr(request.QueryString("id"))
if action="edit" then
	houseObj.getQuotationInfo(id)
End if
str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
<head>
<script language="javascript" src="../../FS_Inc/CheckJs.js"></script>
<script language="javascript" src="../../FS_Inc/prototype.js"></script>
<script language="javascript" src="../../FS_Inc/PublicJS.js"></script>
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
		<td   colspan="2" class="xingmu" height="26">
			<!--#include file="../Top_navi.asp" -->
		</td>
	</tr>
	<tr class="back">
		<td width="18%" valign="top" class="hback">
			<div align="left">
				<!--#include file="../menu.asp" -->
			</div>
		</td>
		<td width="82%" valign="top" class="hback">
			<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
				<tr>
					<td class="xingmu">发布楼盘信息</td>
				</tr>
				<tr>
					<td class="hback"><a href="houseManage.asp">楼盘信息</a>&nbsp;|&nbsp;<a href='#' onClick="history.back()">后退</a>&nbsp;</td>
				</tr>
			</table>
			<form action="HS_Quotation_Open_Action.asp?action=<%=action%>" method="post" id="QuotationForm">
				<input type="hidden" value="<%=id%>" name="id" />
				<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table" id="inputpanel">
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td width="17%" align="right" class="hback">名称：</td>
						<td width="83%" align="left" class="hback">
							<input name="txt_HouseName" type="text" id="txt_HouseName" style="width:60%" onFocus="Do.these('txt_HouseName',function(){ return isEmpty('txt_HouseName','span_HouseName');})" onBlur="Do.these('txt_HouseName',function(){ return isEmpty('txt_HouseName','span_HouseName');})" onKeyUp="Do.these('txt_HouseName',function(){ return isEmpty('txt_HouseName','span_HouseName');})" value="<%=houseObj.HouseName%>" maxlength="50"/>
							<span class="STYLE1">*</span> <span id="span_HouseName"></span></td>
					</tr>
					<!--------------------2/13-----chen------------->
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">开发商：</td>
						<td align="left" class="hback">
							<input name="txt_KaiFaShang" type="text" id="txt_KaiFaShang" style="width:60%" onFocus="this.className='RightInput'" value="<%=houseObj.KaiFaShang%>" maxlength="200" />
						</td>
					</tr>
					<!--------------------------------------------->
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">楼盘位置：</td>
						<td align="left" class="hback">
							<input name="txt_Position" type="text" id="txt_Position" style="width:60%" onFocus="this.className='RightInput'" value="<%=houseObj.Position%>" maxlength="200" />
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">楼盘方位：</td>
						<td align="left" class="hback">
							<input name="txt_Direction" type="text" id="txt_Direction" style="width:60%" onFocus="this.className='RightInput'" value="<%=houseObj.direction%>" maxlength="200" />
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">项目类别：</td>
						<td align="left" class="hback">
							<input name="txt_class" type="text" id="txt_class" style="width:60%"onfocus="this.className='RightInput'" value="<%=houseObj.hclass%>" maxlength="30" />
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">开盘日期：</td>
						<td align="left" class="hback">
							<input name="txt_OpenDate" type="text" id="txt_OpenDate" style="width:60%" onFocus="this.className='RightInput'" value="<%=houseObj.OpenDate%>" maxlength="20" readonly="true" />
							<input type="button" name="chooseEndTime" value="日 期" onClick="OpenWindowAndSetValue('../../admin/CommPages/SelectDate.asp',280,110,window,document.all.QuotationForm.txt_OpenDate);document.all.QuotationForm.txt_OpenDate.focus();">
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">预售许可证：</td>
						<td align="left" class="hback">
							<input name="txt_PreSaleNumber" type="text" id="txt_PreSaleNumber" style="width:60%" onFocus="this.className='RightInput'" value="<%=houseObj.preSaleNumber%>" maxlength="50"/>
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">发证日期：</td>
						<td align="left" class="hback">
							<input name="txt_IssueDate" type="text" id="txt_IssueDate" style="width:60%" onFocus="this.className='RightInput'" value="<%=houseObj.IssueDate%>" maxlength="20" readonly="true" />
							<input type="button" name="chooseEndTime" value="日 期" onClick="OpenWindowAndSetValue('../../admin/CommPages/SelectDate.asp',280,110,window,document.all.QuotationForm.txt_IssueDate);document.all.QuotationForm.txt_IssueDate.focus();">
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">预售范围：</td>
						<td align="left" class="hback">
							<input name="txt_PreSaleRange" type="text" id="txt_PreSaleRange" style="width:60%" onFocus="this.className='RightInput'" value="<%=houseObj.PreSaleRange%>" maxlength="250" />
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">楼盘介绍：</td>
						<td align="left" class="hback">
							<textarea name="txt_introduction" cols="80" rows="15" id="txt_introduction" onFocus="this.className='RightInput'" ><%=houseObj.introduction%></textarea>
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">房屋状况：</td>
						<td align="left" class="hback">
							<select name="sel_Status" id="sel_Status">
								<option value="1" <%if houseObj.status="1" then Response.Write("selected")%>>形象展示</option>
								<option value="2" <%if houseObj.status="2" then Response.Write("selected")%>>期房</option>
								<option value="3" <%if houseObj.status="3" then Response.Write("selected")%>>现房</option>
							</select>
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">报价：</td>
						<td align="left" class="hback">
							<input name="txt_Price" type="text" id="txt_Price" style="width:60%" onFocus="Do.these('txt_Price',function(){return isNumber('txt_Price','span_price','请填写正确的格式',false);})" onKeyUp="Do.these('txt_Price',function(){return isNumber('txt_Price','span_price','请填写正确的格式',false);})" value="<%=houseObj.Price%>" maxlength="14"/>
							万<span id="span_price"></span> </td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">联系电话：</td>
						<td align="left" class="hback">
							<input name="txt_tel" type="text" id="txt_tel" style="width:60%" onFocus="this.className='RightInput'" value="<%=houseObj.tel%>" maxlength="20"  />
						</td>
					</tr>
					<tr id="tr_pic_num" onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">图片数量：</td>
						<td align="left" class="hback">
							<input name="txt_PicNum" type="text" id="txt_PicNum" style="width:60%" onFocus="if($('txt_PicNum').value==''){$('pictable').style.display='none';return;};Do.these('txt_PicNum',function(){return isNumber('txt_PicNum','span_picNum','请填写正确的格式',false);});" onBlur="if($('txt_PicNum').value==''){$('txt_PicNum').value='0'}Do.these('txt_PicNum',function(){return isNumber('txt_PicNum','span_picNum','请填写正确的格式',false);});" onKeyUp="Do.these('txt_PicNum',function(){return isNumber('txt_PicNum','span_picNum','请填写正确的格式',false);});;if(!isNaN(this.value)){addRow(this.value);$('pictable').style.display='';};" value="<%=houseObj.picNumber%>" maxlength="4"/>
							&nbsp;<span id="span_picNum"></span></td>
					</tr>
					<tr class="hback" onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td colspan="2">
							<table width="100%" border="0" cellpadding="0" cellspacing="1" id="pictable" style="display:none">
								<%
	if isnumeric(houseObj.picNumber) then
		if Cint(houseObj.picNumber)>0 and id<>"" then
			Dim pic_rs,picArray
			Set pic_rs=Conn.execute("select Pic from FS_HS_Picture where HS_type=1 and id="&CintStr(id))
			for i=0 to houseObj.picNumber-1
				Response.Write("<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>"&vbcrlf)
				Response.Write("<td width='15%' align='right'>图片"&(i+1)&"</td>"&vbcrlf)
				response.Write("<td width='85%' align='left'>"&vbcrlf)
				if not pic_rs.eof then
					Response.Write("<input name='txt_PicNum_"&(i+1)&"' type='text' id='txt_PicNum' style='width:75%' value='"&pic_rs("pic")&"' onfocus=""this.className='RightInput'"" onblur=""this.className=''"" /><input type='button' name='PPPChoose"&(i+1)&"' value='选择图片' onClick=""OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath="&str_CurrPath&"',500,300,window,this.parentNode.firstChild);"">"&vbcrlf)
				else
					Response.Write("<input name='txt_PicNum_"&(i+1)&"' type='text' id='txt_PicNum' style='width:75%'/><input type='button' name='PPPChoose"&(i+1)&"' value='选择图片' onClick=""OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath="&str_CurrPath&"',500,300,window,this.parentNode.firstChild);"">"&vbcrlf)
				End if
				Response.Write("</td>"&vbcrlf)
				Response.Write("</tr>"&vbcrlf)
				pic_rs.movenext
			next
		End if
	End if
%>
							</table>
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">&nbsp;</td>
						<td align="left" class="hback">
							<input type="Button" name="mysubmit" value="提交" onClick="submitForm()" />
							<input type="Button" name="myrest" value="重置" onClick="javascript:if(confirm('是否重置所有表单项？'))$('QuotationForm').reset()" />
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
	<tr class="back">
		<td height="20"  colspan="2" class="xingmu">
			<div align="left">
				<!--#include file="../Copyright.asp" -->
			</div>
		</td>
	</tr>
</table>
</body>
</html>
<%
Set Fs_User = Nothing
Set Conn=nothing
Set house_rs=nothing
%>
<script language="javascript">
<!--
if(!isNaN($F('txt_PicNum')))
{
	$('pictable').style.display="";
}else
{
	$('pictable').style.display="none";
}
function addRow(number)
{
	var oBody=$("pictable");
	oBody.firstChild.innerText="";
	for(var i=0;i<number;i++)
	{
		var myTR =oBody.insertRow();
		for(var j=0;j<2;j++) 
		{
			var myTD=myTR.insertCell();
			myTD.className="hback"
			if(j==0)
			{
				myTD.style.width="15%"
				myTD.align="right"
				myTD.innerHTML="图片"+(i+1).toString()+"：";
			}
			else
			{
				myTD.style.width="80%"
				myTD.align="left"
				myTD.innerHTML="<input name='txt_PicNum_"+(i+1)+"' type='text' id='txt_PicNum' style='width:75%'/><input type='button' name='PPPChoose"+(i+1)+"' value='选择图片' onClick=\"OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,this.parentNode.firstChild);\">";
			}
		}
	}
}
function submitForm()
{
	var flag1=isEmpty('txt_HouseName','span_HouseName');
	var falg2=isNumber('txt_Price','span_price','请填写正确的数字格式！',false);
	if(flag1&&falg2)
		$('QuotationForm').submit();
}
-->
</script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->







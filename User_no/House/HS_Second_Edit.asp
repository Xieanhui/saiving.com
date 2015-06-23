<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="lib/cls_second.asp"-->
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
'on error resume next
Dim second_rs,action,str_CurrPath,SecondObj,sid,i,tmpArray
set SecondObj=New cls_Second
action=request.QueryString("action")
sid=CintStr(request.QueryString("id"))
if action="edit" then
	SecondObj.getSecondInfo(sid)
End if
str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
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
					<td class="xingmu">发布二手房信息</td>
				</tr>
				<tr>
					<td class="xingmu"><a href="HS_Second.asp">二手房信息</a>&nbsp;|&nbsp;<a href='#' onClick="history.back()">后退</a></td>
				</tr>
			</table>
			<form action="HS_Second_Action.asp?action=<%=action%>" method="post" id="TenancyForm">
				<input type="hidden" value="<%=sid%>" name="id"/>
				<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table" id="inputpanel">
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td width="17%" align="right" class="hback">用途：</td>
						<td width="83%" align="left" class="hback">
							<select name="sel_usefor" id="sel_usefor">
								<option value="1" <%if SecondObj.UseFor="1" then response.Write("selected")%>>住房</option>
								<option value="2" <%if SecondObj.UseFor="2" then response.Write("selected")%>>写字间</option>
							</select>
							<select name="sel_class" id="sel_class">
								<option value="1" <%if SecondObj.sclass="1" then response.Write("selected")%>>个人</option>
								<option value="2" <%if SecondObj.sclass="2" then response.Write("selected")%>>中介</option>
							</select>
							&nbsp;产权性质：
							<select name="sel_BelongType" id="sel_BelongType">
								<option value="1" <%if SecondObj.BelongType="1" then response.Write("selected")%>>商品房</option>
								<option value="2" <%if SecondObj.BelongType="2" then response.Write("selected")%>>集资房</option>
							</select>
							</span></td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">房屋编号：</td>
						<td align="left" class="hback">
							<input name="txt_Label" maxlength="50" type="text" id="txt_Label" style="width:60%" value="<%=secondObj.label%>"  onFocus="this.className='RightInput'"/>
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">房源地址：</td>
						<td align="left" class="hback">
							<input name="txt_address" maxlength="100" type="text" id="txt_address" style="width:60%" value="<%=secondObj.address%>"  onFocus="this.className='RightInput'"/>
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">地段：</td>
						<td align="left" class="hback">
							<input name="txt_Position" maxlength="100" type="text" id="txt_Position" style="width:60%" value="<%=secondObj.Position%>" onFocus="this.className='RightInput'"/>
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">区县：</td>
						<td align="left" class="hback">
							<input name="txt_cityarea" maxlength="20" type="text" id="txt_cityarea" style="width:60%" value="<%=secondObj.cityarea%>" onFocus="this.className='RightInput'"/>
						</td>
					</tr>
					<%
'户型处理
Dim HouseStyleArray(3)
HouseStyleArray(0)="1"
HouseStyleArray(1)="1"
HouseStyleArray(2)="1"
if SecondObj.HouseStyle<>"" then
	tmpArray=split(SecondObj.HouseStyle,",")
	for i=0 to Ubound(tmpArray)
		if i>2 then exit for
		HouseStyleArray(i)=tmpArray(i)
	next
End if
%>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">户型：</td>
						<td align="left" class="hback">
							<input name="txt_HouseStyle_1" type="text" id="txt_HouseStyle_1" size="15"  maxlength="6" onFocus="Do.these('txt_HouseStyle_1',function(){return isNumber('txt_HouseStyle_1','span_housestyle','请填写正确的格式',true);})"
  onkeyup="Do.these('txt_HouseStyle_1',function(){return isNumber('txt_HouseStyle_1','span_housestyle','请填写正确的格式',true);})" value="<%=HouseStyleArray(0)%>"/>
							室
							<input name="txt_HouseStyle_2" type="text" id="txt_HouseStyle_2" size="15" maxlength="6" onFocus="Do.these('txt_HouseStyle_2',function(){return isNumber('txt_HouseStyle_2','span_housestyle','请填写正确的格式',true);})" onKeyUp="Do.these('txt_HouseStyle_2',function(){return isNumber('txt_HouseStyle_2','span_housestyle','请填写正确的格式',true);})" value="<%=HouseStyleArray(1)%>"/>
							厅
							<input name="txt_HouseStyle_3" type="text" id="txt_HouseStyle_3" size="15" maxlength="6" onFocus="Do.these('txt_HouseStyle_3',function(){return isNumber('txt_HouseStyle_3','span_housestyle','请填写正确的格式',true);})" onKeyUp="Do.these('txt_HouseStyle_3',function(){return isNumber('txt_HouseStyle_3','span_housestyle','请填写正确的格式',true);})" value="<%=HouseStyleArray(2)%>"/>
							卫 <span id="span_housestyle"></span></td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">建筑年代：</td>
						<td align="left" class="hback">
							<input name="txt_BuildDate" maxlength="4" type="text" id="txt_BuildDate" style="width:60%"  onFocus="Do.these('txt_BuildDate',function(){ return isNumber('txt_BuildDate','span_BuildDate','请填写正确的格式',true);})" onKeyUp="Do.these('txt_BuildDate',function(){ return isNumber('txt_BuildDate','span_BuildDate','请填写正确的格式',true);})" value="<%if secondObj.BuildDate<>"" then Response.Write(secondObj.BuildDate) else Response.Write("0")%>"/>
							<span id="span_BuildDate"></span> 
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">面积：</td>
						<td align="left" class="hback">
							<input name="txt_Area" type="text" id="txt_Area" style="width:60%" maxlength="10" onFocus="Do.these('txt_Area',function(){ return isNumber('txt_Area','span_area','请填写正确的格式',false);})" onKeyUp="Do.these('txt_Area',function(){ return isNumber('txt_Area','span_area','请填写正确的格式',false);})" value="<%if secondObj.Area<>"" then Response.Write(secondObj.Area) else Response.Write("0")%>"/>
							<span id="span_area"></span></td>
					</tr>
					<%
'楼层处理
Dim FloorArray(2)
FloorArray(0)="0"
FloorArray(1)="0"
if SecondObj.Floor<>"" then
	tmpArray=split(secondObj.Floor,",")
	for i=0 to Ubound(tmpArray)
		if i>1 then exit for
		FloorArray(i)=tmpArray(i)
	next
End if

%>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">楼层：</td>
						<td align="left" class="hback">
							<select name="sel_FloorType" id="sel_FloorType">
								<option value="1" >多 层</option>
								<option value="2">单 层</option>
								<option value="3">小高层</option>
							</select>
							&nbsp;总
							<input name="txt_Floor_1" type="text" id="txt_Floor_1" maxlength="6" size="15" onFocus="Do.these('txt_Floor_1',function(){ return isNumber('txt_Floor_1','span_floor','请填写正确的格式',true);})" onKeyUp="Do.these('txt_Floor_1',function(){ return isNumber('txt_Floor_1','span_floor','请填写正确的格式',true);})" value="<%=FloorArray(0)%>"/>
							层，位于
							<input name="txt_Floor_2" type="text" id="txt_Floor_2" maxlength="6" size="15" onFocus="Do.these('txt_Floor_2',function(){ return isNumber('txt_Floor_2','span_floor','请填写正确的格式',true);})" onKeyUp="Do.these('txt_Floor_2',function(){ return isNumber('txt_Floor_2','span_floor','请填写正确的格式',true);})" value="<%=FloorArray(1)%>"/>
							层<span id="span_floor"></span></td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">结构：</td>
						<td align="left" class="hback">
							<select name="sel_Structure" id="sel_Structure">
								<option value="1" <%if secondObj.Structure="l" then Response.Write("selected")%>>钢筋混凝土</option>
								<option value="2" <%if secondObj.Structure="2" then Response.Write("selected")%>>砖混结构</option>
								<option value="3" <%if secondObj.Structure="3" then Response.Write("selected")%>>砖木结构</option>
								<option value="4" <%if secondObj.Structure="4" then Response.Write("selected")%>>简易结构</option>
							</select>
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">配套设施：</td>
						<td align="left" class="hback">
							<input name="chk_equip" type="checkbox" id="chk_equip" value="l" <%if instr(secondObj.equip,"l")>0 then Response.Write("checked")%> />
							水
							<input name="chk_equip" type="checkbox" id="chk_equip" value="m" <%if instr(secondObj.equip,"m")>0 then Response.Write("checked")%> />
							电
							<input name="chk_equip" type="checkbox" id="chk_equip" value="n" <%if instr(secondObj.equip,"n")>0 then Response.Write("checked")%> />
							气
							<input name="chk_equip_x" type="checkbox" id="chk_equip_x" value="x" <%if instr(secondObj.equip,"x")>0 then Response.Write("checked")%> />
							电话
							<input name="chk_equip" type="checkbox" id="chk_equip" value="y" <%if instr(secondObj.equip,"y")>0 then Response.Write("checked")%> />
							光纤
							<input name="chk_equip" type="checkbox" id="chk_equip" value="z" <%if instr(secondObj.equip,"z")>0 then Response.Write("checked")%>  />
							宽带</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">装修情况：</td>
						<td align="left" class="hback">
							<select name="sel_Decoration" id="sel_Decoration">
								<option value="1" <%if secondObj.Decoration="1" then Response.Write("selected")%>>简单装修</option>
								<option value="2" <%if secondObj.Decoration="2" then Response.Write("selected")%>>中档装修</option>
								<option value="3" <%if secondObj.Decoration="3" then Response.Write("selected")%>>高档装修</option>
								<option value="4" <%if secondObj.Decoration="4" then Response.Write("selected")%>>没有装修</option>
							</select>
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">报价：</td>
						<td align="left" class="hback">
							<input name="txt_Price" type="text" id="txt_Price" style="width:60%" maxlength="24" onFocus="Do.these('txt_Price',function(){ return isNumber('txt_Price','span_price','请填写正确的格式',false);})" onKeyUp="Do.these('txt_Price',function(){ return isNumber('txt_Price','span_price','请填写正确的格式',false);})" value="<%if secondObj.price<>"" then response.Write(secondObj.price) else Response.Write("0")%>"/>
							万<span id="span_price"></span></td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">联系人：</td>
						<td align="left" class="hback">
							<input name="txt_LinkMan" maxlength="100" type="text" id="txt_LinkMan" style="width:60%" onFocus="this.className='RightInput'" value="<%=secondObj.linkMan%>"/>
						</td>
					</tr>
					<tr onMouseOver="overColor(this)" onMouseOut="outColor(this)">
						<td align="right" class="hback">联系方式：</td>
						<td align="left" class="hback">
							<input name="txt_Contact" maxlength="100" type="text" id="txt_Contact" style="width:60%" onFocus="this.className='RightInput'" value="<%=secondObj.Contact%>"/>
						</td>
					</tr>
					<tr onMouseOver="overColor(this)" onMouseOut="outColor(this)">
						<td align="right" class="hback">备注：</td>
						<td align="left" class="hback">
							<textarea name="txt_Remark" maxlength="250" cols="80" rows="15" id="txt_Remark" onFocus="this.className='RightInput'"><%=secondObj.remark%></textarea>
						</td>
					</tr>
					<tr id="tr_pic_num" onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">图片数量：</td>
						<td align="left" class="hback">
							<input name="txt_PicNum" type="text" maxlength="4" id="txt_PicNum" style="width:60%" value="<%=secondObj.picNumber%>" onFocus="if(this.value==''){$('pictable').style.display='none';return;};Do.these('txt_PicNum',function(){ return isNumber('txt_PicNum','span_picNum','请填写正确的格式',true);});"    onKeyUp="Do.these('txt_PicNum',function(){ return isNumber('txt_PicNum','span_picNum','请填写正确的格式',true);});if(!isNaN(this.value)){addRow(this.value);$('pictable').style.display=''};" onBlur="if(this.value==''){$('txt_PicNum').value='0'};Do.these('txt_PicNum',function(){ return isNumber('txt_PicNum','span_picNum','请填写正确的格式',true);})"/>
							<span id="span_picNum"></span></td>
					</tr>
					<tr class="hback" onMouseOver=overColor(this) onMouseOut=outColor(this) >
						<td colspan="2">
							<table width="100%" border="0" cellpadding="0" cellspacing="1" id="pictable" style="display:none">
								<%
	if isnumeric(secondObj.picNumber) then
		if Cint(secondObj.picNumber)>0 and sid<>"" then
			Dim pic_rs,picArray
			Set pic_rs=Conn.execute("select Pic from FS_HS_Picture where HS_type=3 and id="&CintStr(sid))
			for i=0 to secondObj.picNumber-1
				if pic_rs.eof then exit for
				Response.Write("<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>"&vbcrlf)
				Response.Write("<td width='15%' align='right'>图片"&(i+1)&"</td>"&vbcrlf)
				response.Write("<td width='85%' align='left'>"&vbcrlf)
				if not pic_rs.eof then
					Response.Write("<input name='txt_PicNum_"&(i+1)&"' type='text' id='txt_PicNum' style='width:75%' value='"&pic_rs("pic")&"'/><input type='button' name='PPPChoose"&(i+1)&"' value='选择图片' onClick=""OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath="&str_CurrPath&"',500,300,window,this.parentNode.firstChild);"">"&vbcrlf)
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
							<input type="Button" name="mysubmit" onClick="CheckThisForm(this.form)" value="提交" />
							<input type="Button" name="myrest" value="重置" onClick="javascript:if(confirm('是否重置所有表单项？'))$('TenancyForm').reset()" />
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
Set SecondObj = Nothing
%>
<script language="javascript">
<!--
function CheckThisForm(formobj){
	var inputcheck;
	inputcheck=isNumber('txt_HouseStyle_1','span_housestyle','请填写正确的格式',true);
	inputcheck=inputcheck&&isNumber('txt_HouseStyle_2','span_housestyle','请填写正确的格式',true);
	inputcheck=inputcheck&&isNumber('txt_HouseStyle_3','span_housestyle','请填写正确的格式',true);
	inputcheck=inputcheck&&isNumber('txt_BuildDate','span_BuildDate','请填写正确的格式',true);
	inputcheck=inputcheck&&isNumber('txt_Area','span_area','请填写正确的格式',false);
	inputcheck=inputcheck&&isNumber('txt_Floor_1','span_floor','请填写正确的格式',true);
	inputcheck=inputcheck&&isNumber('txt_Floor_2','span_floor','请填写正确的格式',true);
	inputcheck=inputcheck&&isNumber('txt_Price','span_price','请填写正确的格式',false);
	inputcheck=inputcheck&&isNumber('txt_PicNum','span_picNum','请填写正确的格式',true);
	if (inputcheck){
		formobj.submit();
	}else{
		alert("请填写正确的格式");
	}
}
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
-->
</script>
</html>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->







<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="./lib/cls_tenancy.asp"-->
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
Dim tenancy_rs,action,str_CurrPath,tenancyObj,tid,i,tmpArray
set TenancyObj=New cls_tenancy
action=request.QueryString("action")
tid=CintStr(request.QueryString("id"))
if action="edit" then
	TenancyObj.getTenancyInfo(tid)
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
    <td width="82%" valign="top" class="hback"><table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
        <tr>
          <td class="xingmu">发布租赁信息</td>
        </tr>
        <tr>
          <td class="xingmu">租赁信息|<a href='#' onClick="history.back()">后退</a>&nbsp;</td>
        </tr>
      </table>
      <form action="HS_Tenancy_Action.asp?action=<%=action%>" method="post" id="TenancyForm">
        <input type="hidden" value="<%=tid%>" name="id"/>
        <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table" id="inputpanel">
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td width="17%" align="right" class="hback">用途：</td>
            <td width="83%" align="left" class="hback"><select name="sel_usefor" id="sel_usefor">
                <option value="1" <%if TenancyObj.UseFor="1" then response.Write("selected")%>>住房</option>
                <option value="2" <%if TenancyObj.UseFor="2" then response.Write("selected")%>>写字间</option>
              </select>
              <span id="span_HouseName"></span> &nbsp;&nbsp;&nbsp;&nbsp;房屋性质
              <!-----------------2/13 chen--------------------------------------->
              <select name="sel_XingZhi" id="sel_XingZhi">
                <option value="1" <%if TenancyObj.XingZhi="1" then response.Write("selected")%>>商品房</option>
                <option value="2" <%if TenancyObj.XingZhi="2" then response.Write("selected")%>>集资房</option>
                <option value="3" <%if TenancyObj.XingZhi="3" then response.Write("selected")%>>其他房</option>
              </select>
              <span id="span_HouseName"></span> </td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td align="right" class="hback">类型：</td>
            <td align="left" class="hback"><select name="sel_class" id="sel_class">
                <option value="1" <%if TenancyObj.tclass="1" then response.Write("selected")%>>出租</option>
                <option value="2" <%if TenancyObj.tclass="2" then response.Write("selected")%>>求租</option>
                <option value="3" <%if TenancyObj.tclass="3" then response.Write("selected")%>>出售</option>
                <option value="4" <%if TenancyObj.tclass="4" then response.Write("selected")%>>求购</option>
                <option value="5" <%if TenancyObj.tclass="5" then response.Write("selected")%>>合租</option>
                <option value="6" <%if TenancyObj.tclass="6" then response.Write("selected")%>>转让</option>
              </select>
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;杂物间
              <select name="ZaWuJian" id="ZaWuJian">
                <option value="0" <%if TenancyObj.ZaWuJian="0" then response.Write("selected")%>>有</option>
                <option value="1" <%if TenancyObj.ZaWuJian="1" then response.Write("selected")%>>无</option>
              </select>
            </td>
          </tr>
          <!---------------------2/13 chen------------------------------------->
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td align="right" class="hback">小区名称：</td>
            <td align="left" class="hback"><input name="txt_XiaoQuName" type="text" id="txt_XiaoQuName" style="width:60%" onFocus="this.className='RightInput'" value="<%=TenancyObj.XiaoQuName%>" maxlength="200"/>
            </td>
          </tr>
          <!----------------------------------------------------------->
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td align="right" class="hback">房源地址：</td>
            <td align="left" class="hback"><input name="txt_Position" type="text" id="txt_Position" style="width:60%" onFocus="this.className='RightInput'" value="<%=TenancyObj.Position%>" maxlength="200"/>
            </td>
          </tr>
          <!-------------------------2/13 chen-------------------------------->
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td align="right" class="hback">交通状况描述：</td>
            <td align="left" class="hback"><input name="txt_JiaoTong" type="text" id="txt_JiaoTong" style="width:60%" onFocus="this.className='RightInput'" value="<%=TenancyObj.JiaoTong%>" maxlength="200"/>
            </td>
          </tr>
          <!------------------------------------------------------------------>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td align="right" class="hback">区县：</td>
            <td align="left" class="hback"><input name="txt_cityarea" type="text" id="txt_cityarea" style="width:60%" onFocus="this.className='RightInput'" value="<%=TenancyObj.cityarea%>" maxlength="20"/>
            </td>
          </tr>
          <%
'户型处理
Dim HouseStyleArray(3)
HouseStyleArray(0)="1"
HouseStyleArray(1)="1"
HouseStyleArray(2)="1"
if TenancyObj.HouseStyle<>"" then
	tmpArray=split(TenancyObj.HouseStyle,",")
	for i=0 to Ubound(tmpArray)
		if i>2 then exit for
		HouseStyleArray(i)=tmpArray(i)
	next
End if
%>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td align="right" class="hback">户型：</td>
            <td align="left" class="hback"><input name="txt_HouseStyle_1" type="text" id="txt_HouseStyle_1"  onFocus="Do.these('txt_HouseStyle_1',function(){return isNumber('txt_HouseStyle_1','span_housestyle','请填写正确的格式',true);})"
  onkeyup="Do.these('txt_HouseStyle_1',function(){return isNumber('txt_HouseStyle_1','span_housestyle','请填写正确的格式',true);})" value="<%=HouseStyleArray(0)%>" size="15" maxlength="6"/>
              室
              <input name="txt_HouseStyle_2" type="text" id="txt_HouseStyle_2"  onFocus="Do.these('txt_HouseStyle_2',function(){return isNumber('txt_HouseStyle_2','span_housestyle','请填写正确的格式',true);})" onKeyUp="Do.these('txt_HouseStyle_2',function(){return isNumber('txt_HouseStyle_2','span_housestyle','请填写正确的格式',true);})" value="<%=HouseStyleArray(1)%>" size="15" maxlength="6"/>
              厅
              <input name="txt_HouseStyle_3" type="text" id="txt_HouseStyle_3"  onFocus="Do.these('txt_HouseStyle_3',function(){return isNumber('txt_HouseStyle_3','span_housestyle','请填写正确的格式',true);})" onKeyUp="Do.these('txt_HouseStyle_3',function(){return isNumber('txt_HouseStyle_3','span_housestyle','请填写正确的格式',true);})" value="<%=HouseStyleArray(2)%>" size="15" maxlength="6"/>
              卫 <span id="span_housestyle"></span></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td align="right" class="hback">建筑年代：</td>
            <td align="left" class="hback"><input name="txt_BuildDate" type="text" id="txt_BuildDate" style="width:60%" onFocus="Do.these('txt_BuildDate',function(){ return isNumber('txt_BuildDate','span_BuildDate','请填写正确的格式',true);})" onKeyUp="Do.these('txt_BuildDate',function(){ return isNumber('txt_BuildDate','span_BuildDate','请填写正确的格式',true);})" value="<%if tenancyObj.builddate<>"" then Response.Write(tenancyObj.builddate) else Response.Write("0")%>" maxlength="4"/>
              <span id="span_BuildDate"></span> 
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td align="right" class="hback">面积：</td>
            <td align="left" class="hback"><input name="txt_Area" type="text" id="txt_Area" style="width:60%" onFocus="Do.these('txt_Area',function(){ return isNumber('txt_Area','span_area','请填写正确的格式',false);})" onKeyUp="Do.these('txt_Area',function(){ return isNumber('txt_Area','span_area','请填写正确的格式',false);})" value="<%if tenancyObj.Area<>"" then Response.Write(tenancyObj.Area) else Response.Write("0")%>" maxlength="10"/>
              平米 <span id="span_area"></span></td>
          </tr>
          <%
'楼层处理
Dim FloorArray(2)
FloorArray(0)="0"
FloorArray(1)="0"
if tenancyObj.Floor<>"" then
	tmpArray=split(tenancyObj.Floor,",")
	for i=0 to Ubound(tmpArray)
		if i>1 then exit for
		FloorArray(i)=tmpArray(i)
	next
End if

%>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td align="right" class="hback">楼层：</td>
            <td align="left" class="hback"> 总
              <input name="txt_Floor_1" type="text" id="txt_Floor_1" onFocus="Do.these('txt_Floor_1',function(){ return isNumber('txt_Floor_1','span_floor','请填写正确的格式',true);})" onKeyUp="Do.these('txt_Floor_1',function(){ return isNumber('txt_Floor_1','span_floor','请填写正确的格式',true);})" value="<%=FloorArray(0)%>" size="15" maxlength="30"/>
              层，位于
              <input name="txt_Floor_2" type="text" id="txt_Floor_2" onFocus="Do.these('txt_Floor_2',function(){ return isNumber('txt_Floor_2','span_floor','请填写正确的格式',true);})" onKeyUp="Do.these('txt_Floor_2',function(){ return isNumber('txt_Floor_2','span_floor','请填写正确的格式',true);})" value="<%=FloorArray(1)%>" size="15" maxlength="30"/>
              层<span id="span_floor"></span></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td align="right" class="hback">配套设施：</td>
            <td align="left" class="hback"><input name="chk_equip" type="checkbox" id="chk_equip" value="l" <%if instr(tenancyObj.equip,"l")>0 then Response.Write("checked")%> />
              水
              <input name="chk_equip" type="checkbox" id="chk_equip" value="m" <%if instr(tenancyObj.equip,"m")>0 then Response.Write("checked")%> />
              电
              <input name="chk_equip" type="checkbox" id="chk_equip" value="n" <%if instr(tenancyObj.equip,"n")>0 then Response.Write("checked")%> />
              气
              <input name="chk_equip_x" type="checkbox" id="chk_equip_x" value="x" <%if instr(tenancyObj.equip,"x")>0 then Response.Write("checked")%> />
              电话
              <input name="chk_equip" type="checkbox" id="chk_equip" value="y" <%if instr(tenancyObj.equip,"y")>0 then Response.Write("checked")%> />
              光纤
              <input name="chk_equip" type="checkbox" id="chk_equip" value="z" <%if instr(tenancyObj.equip,"z")>0 then Response.Write("checked")%>  />
              宽带</td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td align="right" class="hback">装修情况：</td>
            <td align="left" class="hback"><select name="sel_Decoration" id="sel_Decoration">
                <option value="1" <%if tenancyObj.Decoration="1" then Response.Write("selected")%>>简单装修</option>
                <option value="2" <%if tenancyObj.Decoration="2" then Response.Write("selected")%>>中档装修</option>
                <option value="3" <%if tenancyObj.Decoration="3" then Response.Write("selected")%>>高档装修</option>
                <option value="4" <%if tenancyObj.Decoration="4" then Response.Write("selected")%>>没有装修</option>
              </select>
            </td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td align="right" class="hback">报价：</td>
            <td align="left" class="hback"><input name="txt_Price" type="text" id="txt_Price" style="width:60%" onFocus="Do.these('txt_Price',function(){ return isNumber('txt_Price','span_price','请填写正确的格式',false);})" onKeyUp="Do.these('txt_Price',function(){ return isNumber('txt_Price','span_price','请填写正确的格式',false);})" value="<%if tenancyObj.price<>"" then response.Write(tenancyObj.price) else Response.Write("0")%>" maxlength="14"/>
              元<span id="span_price"></span></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td align="right" class="hback">联系人：</td>
            <td align="left" class="hback"><input name="txt_LinkMan" type="text" id="txt_LinkMan" style="width:60%" onFocus="this.className='RightInput'" value="<%=tenancyObj.linkMan%>" maxlength="100"/>
            </td>
          </tr>
          <tr onMouseOver="overColor(this)" onMouseOut="outColor(this)">
            <td align="right" class="hback">联系方式：</td>
            <td align="left" class="hback"><input name="txt_Contact" type="text" id="txt_Contact" style="width:60%" onFocus="this.className='RightInput'" value="<%=tenancyObj.Contact%>" maxlength="100"/>
            </td>
          </tr>
          <tr onMouseOver="overColor(this)" onMouseOut="outColor(this)">
            <td align="right" class="hback">有效期：</td>
            <td align="left" class="hback"><input name="txt_Period" type="text" id="txt_Period" style="width:60%" onFocus="this.className='RightInput'" value="<%if tenancyObj.Period<>"" then response.Write(tenancyObj.Period) else Response.Write("三个月")%>" maxlength="10"/>
              <span id="span_period"></span></td>
          </tr>
          <tr onMouseOver="overColor(this)" onMouseOut="outColor(this)">
            <td align="right" class="hback">备注：</td>
            <td align="left" class="hback"><textarea name="txt_Remark" cols="80" rows="15" id="txt_Remark" onFocus="this.className='RightInput'"><%=TenancyObj.remark%></textarea>
            </td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
          <tr id="tr_pic_num" onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td align="right" class="hback">图片数量：</td>
            <td align="left" class="hback"><input name="txt_PicNum" type="text" id="txt_PicNum" style="width:60%" onFocus="if(this.value==''){$('pictable').style.display='none';return;};Do.these('txt_PicNum',function(){ return isNumber('txt_PicNum','span_picNum','请填写正确的格式',true);});" onBlur="if(this.value==''){$('txt_PicNum').value='0'};Do.these('txt_PicNum',function(){ return isNumber('txt_PicNum','span_picNum','请填写正确的格式',true);})"    onKeyUp="Do.these('txt_PicNum',function(){ return isNumber('txt_PicNum','span_picNum','请填写正确的格式',true);});if(!isNaN(this.value)){addRow(this.value);$('pictable').style.display=''};" value="<%=tenancyObj.picNumber%>" maxlength="4"/>
              <span id="span_picNum"></span></td>
          </tr>
          <tr class="hback" onMouseOver=overColor(this) onMouseOut=outColor(this) >
            <td colspan="2"><table width="100%" border="0" cellpadding="0" cellspacing="1" id="pictable" style="display:none">
                <%
	if isnumeric(TenancyObj.picNumber) then
		if Cint(TenancyObj.picNumber)>0 and tid<>"" then
			Dim pic_rs,picArray
			Set pic_rs=Conn.execute("select Pic from FS_HS_Picture where HS_type=2 and id="&tid)
			for i=0 to TenancyObj.picNumber-1
				if pic_rs.eof then exit for
				Response.Write("<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>"&vbcrlf)
				Response.Write("<td width='15%' align='right'>图片"&(i+1)&"</td>"&vbcrlf)
				response.Write("<td width='85%' align='left'>"&vbcrlf)
				if not pic_rs.eof then
					Response.Write("<input name='txt_PicNum_"&(i+1)&"' type='text' id='txt_PicNum' style='width:75%' value='"&pic_rs("pic")&"'/><input type='button' name='PPPChoose"&(i+1)&"' value='选择图片' onClick=""OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath="&str_CurrPath&"&f_UserNumber="&session("FS_UserNumber")&"',500,300,window,this.parentNode.firstChild);"">"&vbcrlf)
				else
					Response.Write("<input name='txt_PicNum_"&(i+1)&"' type='text' id='txt_PicNum' style='width:75%'/><input type='button' name='PPPChoose"&(i+1)&"' value='选择图片' onClick=""OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath="&str_CurrPath&"&f_UserNumber="&session("FS_UserNumber")&"',500,300,window,this.parentNode.firstChild);"">"&vbcrlf)
				End if
				Response.Write("</td>"&vbcrlf)
				Response.Write("</tr>"&vbcrlf)
				pic_rs.movenext
			next
		End if
	End if
%>
              </table></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
            <td align="right" class="hback">&nbsp;</td>
            <td align="left" class="hback"><input type="Button" name="mysubmit" onClick="CheckThisForm(this.form)" value="提交" />
              <input type="Button" name="myrest" value="重置" onClick="javascript:if(confirm('是否重置所有表单项？'))$('TenancyForm').reset()" />
            </td>
          </tr>
        </table>
      </form></td>
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
Set tenancy_rs=nothing
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
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->







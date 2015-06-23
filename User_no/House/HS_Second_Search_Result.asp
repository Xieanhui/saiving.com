<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%'Copyright (c) 2006 Foosun Inc. 
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim Hs_Rs
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=15 '设置每页显示数目
int_showNumberLink_=10 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings>9</font>"   			'首页 
toP10_=" <font face=webdings>7</font>"			'上十
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"				'尾页

''得到相关表的值。
Function Get_OtherTable_Value(This_Fun_Sql)
	Dim This_Fun_Rs
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
	else
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	end if			
	if not This_Fun_Rs.eof then 
		Get_OtherTable_Value = This_Fun_Rs(0)
	else
		Get_OtherTable_Value = ""
	end if
	if Err.Number>0 then 
		Err.Clear
		response.Redirect("../lib/error.asp?ErrCodes=<li>Get_OtherTable_Value未能得到相关数据。错误描述："&Err.Description&"</li>") : response.End()
	end if
	set This_Fun_Rs=nothing 
End Function
  
Function Get_While_Info(Add_Sql)
	Add_Sql = Decrypt(Add_Sql)
	Dim Get_Html,This_Fun_Sql,ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp
	Str_Tmp = "SID,HouseStyle,PubDate,Audited,Class,UserNumber,Label,UseFor,FloorType,BelongType,Structure,Area,BuildDate,Price,CityArea,Address,Floor,Position,Decoration,LinkMan,Contact,equip,Remark"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_HS_Second"
	Arr_Tmp = split(Str_Tmp,",")
	for each Str_Tmp in Arr_Tmp
		Req_Str = NoSqlHack(Trim(request(Str_Tmp)))
		if Req_Str<>"" then 				
			select case Str_Tmp
			case "SID","Class","UseFor","FloorType","BelongType","Structure","Area","BuildDate","Price","Decoration","PubDate","Audited","PicNumber","isRecyle"
			''数字,日期
				regxp = "|<|>|=|<=|>=|<>|"
				if instr(regxp,"|"&left(Req_Str,1)&"|")>0 or instr(regxp,"|"&left(Req_Str,2)&"|")>0 then 
					New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" "& Req_Str
				elseif instr(Req_Str,"*")>0 then 
					if left(Req_Str,1)="*" then Req_Str = "%"&mid(Req_Str,2)
					if right(Req_Str,1)="*" then Req_Str = mid(Req_Str,1,len(Req_Str) - 1) & "%"							
				New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" like '"& Req_Str &"'"							
				else	
					New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" = "& Req_Str
				end if		
			case else
			''字符
				New_Search_Str = and_where(New_Search_Str) & Search_TextArr(Req_Str,Str_Tmp,"")
			end select 		
		end if
	next
	if New_Search_Str<>"" then This_Fun_Sql = and_where(This_Fun_Sql) & replace(New_Search_Str," where ","")
	if instr(Add_Sql,"filterorderby")>0 then 
		This_Fun_Sql = This_Fun_Sql &"  "& replace(replace(Add_Sql,"csed"," Desc"),"filterorderby","Order By ")
	elseif Add_Sql<>"" then 
		This_Fun_Sql = and_where(This_Fun_Sql) &" "& Add_Sql	
	end if
	Str_Tmp = ""
	'response.Write(This_Fun_Sql)
	On Error Resume Next
	Set Hs_Rs = CreateObject(G_FS_RS)
	Hs_Rs.Open This_Fun_Sql,Conn,1,1	
	if Err<>0 then 
		Err.Clear
		response.Redirect("../lib/error.asp?ErrCodes=<li>查询出错："&Err.Description&"</li><li>请检查字段类型是否匹配.</li>")
		response.End()
	end if
	IF not Hs_Rs.eof THEN

	Hs_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>Hs_Rs.PageCount Then cPageNo=Hs_Rs.PageCount 
	Hs_Rs.AbsolutePage=cPageNo
	
	  FOR int_Start=1 TO int_RPP 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" onmouseover=""this.style.cursor = 'hand'"" onclick=""javascript:if(TD_U_"&Hs_Rs("SID")&".style.display=='') TD_U_"&Hs_Rs("SID")&".style.display='none'; else TD_U_"&Hs_Rs("SID")&".style.display='';"" title='点击查看更多信息'>"&Hs_Rs("SID")&"</td>" & vbcrlf
		Str_Tmp = Hs_Rs("HouseStyle")
		Str_Tmp = left(Str_Tmp,instr(Str_Tmp,",") - 1)&"室"& mid(Str_Tmp,instr(Str_Tmp,",") + 1,len(Str_Tmp)) &"厅" & right(Str_Tmp,len(Str_Tmp) - instrrev(Str_Tmp,","))&"室"
		Get_Html = Get_Html & "<td align=""center"">"&Str_Tmp&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Hs_Rs("PubDate")&"</td>" & vbcrlf
		if Hs_Rs("Audited")=0 then 
			Str_Tmp = "未审核"
		elseif Hs_Rs("Audited")=1 then 
			Str_Tmp = "已审核"
		else
			Str_Tmp = "未定义的状态?"	
		end if
		Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		if Hs_Rs("UserNumber")=0 then 
			Str_Tmp = "系统管理员"
		elseif Hs_Rs("UserNumber")<>"" then 
			Str_Tmp = Get_OtherTable_Value("select C_Name+' [地址:'+C_Address+' 电话:'+C_Tel+' 传真:'+C_Fax+']' as tmpfild from FS_ME_CorpUser where UserNumber='"&Hs_Rs("UserNumber")&"'")
			if Str_Tmp<>"" then 
				Str_Tmp = "<span  onmouseover=""this.style.cursor = 'hand'"" id="""&Hs_Rs("UserNumber")&""" title="""&Str_Tmp&""">中介公司:"&Hs_Rs("UserNumber")&"</span>"
			else
				Str_Tmp = "未知的用户"
			end if	
		else
			Str_Tmp = "中介公司"	
		end if
		Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		''++++++++++++++++++++++++++++++++++++++点开时显示详细信息。
		Str_Tmp = Hs_Rs("Floor")
		Str_Tmp = "总层" & left(Str_Tmp,instr(Str_Tmp,",") - 1)&",第"& mid(Str_Tmp,instr(Str_Tmp,",") + 1,len(Str_Tmp)) &"层"
		Get_Html = Get_Html & "<tr class=""hback"" id=""TD_U_"& Hs_Rs("SID") &""" style=""display:none""><td colspan=20>" & vbcrlf
		Get_Html = Get_Html & "<table width=""100%"" height=""30"" border=""0"" cellspacing=""1"" cellpadding=""2"" class=""table"">" & vbcrlf 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>交易类型:"&Replacestr(Hs_Rs("Class"),"1:个人,2:中介")& "</td><td>会员编号:"&Replacestr(Hs_Rs("UserNumber"),"0:个人")&"</td><td>房屋编号:"&Replacestr(Hs_Rs("Label"),":无")& "</td><td>使用性质:"&Replacestr(Hs_Rs("UseFor"),"1:住宅,2:办公,3:其它")& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>住宅类别:"&Replacestr(Hs_Rs("FloorType"),"1:多层,2:单层")& "</td><td>产权性质:"&Replacestr(Hs_Rs("BelongType"),"1:商品房,2:集资房")&"</td><td>建筑结构:"&Replacestr(Hs_Rs("Structure"),"1:砖混,else")& "</td><td>建筑面积:"&Hs_Rs("Area")& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>建筑年代:"&Hs_Rs("BuildDate") & "</td><td>售价:"&Hs_Rs("Price")&"万元</td><td>区县:"&Hs_Rs("CityArea") & "</td><td>地址:"&Hs_Rs("Address")& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>楼层:"& Str_Tmp & "</td><td>地段:"&Hs_Rs("Position")&"</td><td>装修情况:"&Replacestr(Hs_Rs("Decoration"),"1:简单装修2:中档装修3:高档装修") & "</td><td>联系人:"&Hs_Rs("LinkMan")& "</td></tr>" & vbcrlf
		Str_Tmp = Hs_Rs("equip")
		Str_Tmp = Replace(Str_Tmp,"l","水")
		Str_Tmp = Replace(Str_Tmp,"m","电")
		Str_Tmp = Replace(Str_Tmp,"n","气")
		Str_Tmp = Replace(Str_Tmp,"x","电话")
		Str_Tmp = Replace(Str_Tmp,"y","光纤")
		Str_Tmp = Replace(Str_Tmp,"z","宽带")
		Str_Tmp = Replace(Str_Tmp,"0","无")
		Str_Tmp = Replace(Str_Tmp,"m","有")
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>联系方式:"& Hs_Rs("Contact") & "</td><td>配套设施:"& Str_Tmp &"</td><td>备注:"&Hs_Rs("Remark")& "</td><td>&nbsp;</td></tr>" & vbcrlf
		Get_Html = Get_Html & "</table>" & vbcrlf
		Get_Html = Get_Html &"</td></tr>" & vbcrlf
		''+++++++++++++++++++++++++++++++++++++++		
		Hs_Rs.MoveNext
 		if Hs_Rs.eof or Hs_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(Hs_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	Get_Html = Get_Html &"</td></tr>"
	Hs_Rs.close
	Get_While_Info = Get_Html
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-楼盘搜索</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
//点击标题排序
/////////////////////////////////////////////////////////
var Old_Sql = document.URL;
function OrderByName(FildName)
{
	var New_Sql='';
	var oldFildName="";
	if (Old_Sql.indexOf("&Add_Sql=")==-1&&Old_Sql.indexOf("?Add_Sql=")==-1)
	{
		if (Old_Sql.indexOf("=")>-1)
			New_Sql = Old_Sql+"&Add_Sql=filterorderby" + FildName + "csed";
		else
			New_Sql = Old_Sql+"?Add_Sql=filterorderby" + FildName + "csed";
	}
	else
	{	
		var tmp_arr_ = Old_Sql.split('?')[1].split('&');
		for(var ii=0;ii<tmp_arr_.length;ii++)
		{
			if (tmp_arr_[ii].indexOf("filterorderby")>-1)
			{
				oldFildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby") + "filterorderby".length , tmp_arr_[ii].length);
				break;	
			}
		}
		oldFildName.indexOf("csed")>-1?New_Sql = Old_Sql.replace(oldFildName,FildName):New_Sql = Old_Sql.replace(oldFildName,FildName+"csed");
	}	
	//alert(New_Sql);
	location = New_Sql;
}
/////////////////////////////////////////////////////////
-->
</script>
<head>
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
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="../Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="../menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr class="hback"> 
            
          <td class="hback"><strong>位置：</strong><a href="../../">网站首页</a> &gt;&gt; 
            <a href="../main.asp">会员首页</a> &gt;&gt; <a href="default.asp">房产</a>－二手房搜索</td>
          </tr>
        </table>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr  class="hback"> 
    <td class="xingmu" >二手房搜索</td>
  </tr>
  <tr  class="hback"> 
    <td><a href="HS_Quotation_Search.asp">首页</a>
      | <a href="HS_Quotation_Search.asp">楼盘信息</a> | 
      <a href="HS_Second_Search.asp">二手房信息</a> | 
	  <a href="HS_Tenancy_Search.asp">租赁信息</a></td>
  </tr>
</table>
<%
'******************************************************************
Call View
'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" method="post" action="?Act=Del&TableName=<%=request.QueryString("TableName")%>">
    <tr  class="hback"> 
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('SID')" class="sd"><b>〖ID〗</b></a> 
        <span id="Show_Oder_SID"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('HouseStyle')" class="sd"><b>房源类型</b></a> 
        <span id="Show_Oder_HouseStyle"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('PubDate')" class="sd"><b>发布日期</b></a> 
        <span id="Show_Oder_PubDate"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('Audited')" class="sd"><b>是否审核</b></a> 
        <span id="Show_Oder_Audited"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('UserNumber')" class="sd"><b>会员编号</b></a> 
        <span id="Show_Oder_UserNumber"></span></td>
    </tr>
    <%
		response.Write( Get_While_Info( Encrypt(request.QueryString("Add_Sql")) ) )
	%>
  </form>
</table>
<%End Sub%>
       </td>
    </tr>
    <tr class="back"> 
      <td height="20" colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="../Copyright.asp" -->
        </div></td>
    </tr>
 
</table>
<%
Set Hs_Rs=nothing
Set Fs_User = Nothing
%>
<script language="JavaScript">
<!--//判断后将排序完善.字段名后面显示指示
//打开后根据规则显示箭头
var Req_FildName;
if (Old_Sql.indexOf("filterorderby")>-1)
{
	var tmp_arr_ = Old_Sql.split('?')[1].split('&');
	for(var ii=0;ii<tmp_arr_.length;ii++)
	{
		if (tmp_arr_[ii].indexOf("filterorderby")>-1)
		{
			if(Old_Sql.indexOf("csed")>-1)
				{Req_FildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby") + "filterorderby".length , tmp_arr_[ii].indexOf("csed"));break;}
			else
				{Req_FildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby") + "filterorderby".length , tmp_arr_[ii].length);break;}	
		}
	}	
	if (document.getElementById('Show_Oder_'+Req_FildName)!=null)  
	{
		if(Old_Sql.indexOf(Req_FildName + "csed")>-1)
		{
			eval('Show_Oder_'+Req_FildName).innerText = '↓';
		}
		else
		{
			eval('Show_Oder_'+Req_FildName).innerText = '↑';
		}
	}	
}
/////////////////////////////////////////////////////////
-->
</script>

<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-楼盘搜索</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
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
            <a href="../main.asp">会员首页</a> &gt;&gt; <a href="default.asp">房产</a>－楼盘搜索</td>
          </tr>
        </table>
        
      <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" method="post" action="HS_Quotation_Search_Result.asp">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" >查询房源 楼盘信息</td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">ID</td>
      <td>
	   <input type="text" name="ID" size="40" value="">
      </td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">楼盘名称</td>
      <td> <input type="text" name="HouseName" size="40" value="">
              <span class="tx">支持A* *B A B模式字符类型同理</span></td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">楼盘位置</td>
      <td> <input type="text" name="Position" size="40" value="">
      </td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">楼盘方位</td>
      <td> <input type="text" name="Direction" size="40" value="">
      </td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">项目类别（如高层，低层）</td>
      <td> <input type="text" name="Class" size="40" value="">
      </td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">开盘日期</td>
      <td> <input type="text" name="OpenDate" size="40" value="">
      </td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">预售许可证</td>
      <td> <input type="text" name="PreSaleNumber" size="40" value="">
      </td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">发证日期</td>
      <td> <input type="text" name="IssueDate" size="40" value="">
      </td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">本次预售范围</td>
      <td> <input type="text" name="PreSaleRange" size="40" value="">
      </td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">房屋状况</td>
      <td> <input type="text" name="Status" size="40" value="">
      </td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">报价</td>
      <td> <input type="text" name="Price" size="40" value="">
      </td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">发布日期</td>
      <td> <input type="text" name="PubDate" size="40" value="">
              <span class="tx">支持&gt;,&gt;=,&lt;,&lt;=,=2006-08数字日期型同理</span></td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">联系电话</td>
      <td> <input type="text" name="Tel" size="40" value="">
      </td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">点击</td>
      <td> <input type="text" name="Click" size="40" value="">
              <span class="tx">支持&gt;,&gt;=,&lt;,&lt;=,=888数字日期型同理</span></td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">会员编号</td>
      <td> <input type="text" name="UserNumber" size="40" value="">
      <span class="tx">0表示系统管理员，否则是房产公司</span></td>
    </tr>
    <tr  class="hback"> 
      <td align="right">是否通过审核</td>
      <td> <input type="radio" name="Audited"  value="1">
        已审核 
        <input type="radio" name="Audited"  value="0">
        未审核</td>
    </tr>
	
    <tr  class="hback"> 
      <td colspan="4"> <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit" name="submit" value=" 执行查询 " /> 
              &nbsp; <input type="reset" name="ReSet" id="ReSet" value=" 重置 " /> 
            </td>
          </tr>
        </table></td>
    </tr>
  </form>
</table>

       </td>
    </tr>
    <tr class="back"> 
      <td height="20" colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="../Copyright.asp" -->
        </div></td>
    </tr>
 
</table>
<%
Set Fs_User = Nothing
%>

<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






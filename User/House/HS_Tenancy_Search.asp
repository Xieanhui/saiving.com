<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-租赁搜索</title>
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
            <a href="../main.asp">会员首页</a> &gt;&gt; <a href="default.asp">房产</a>－租赁搜索</td>
          </tr>
        </table>

        <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
          <form name="form3" method="post" action="HS_Tenancy_Search_Result.asp">
            <tr  class="hback"> 
              <td colspan="3" align="left" class="xingmu" >查询房源 租赁信息</td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">ID</td>
              <td> <input type="text" name="TID" size="40" value=""> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">用途</td>
              <td> <select name="UseFor">
                  <option value="">所有</option>
                  <option value="1">住房</option>
                  <option value="2">写字间</option>
                </select> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">类型</td>
              <td> <select name="Class">
                  <option value="">所有</option>
                  <option value="1">出租</option>
                  <option value="2">求租</option>
                  <option value="3">出售</option>
                  <option value="4">求购</option>
                  <option value="5">合租</option>
                  <option value="6">转让</option>
                </select> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">房源地址</td>
              <td> <input type="text" name="Position" size="40" value=""> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">区县</td>
              <td> <input type="text" name="CityArea" size="40" value=""> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">价格</td>
              <td> <input type="text" name="Price" size="40" value=""> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">户型</td>
              <td> <input type="text" name="Hou愀正??? seStyle" size="40" value=""> <span class="tx">户型,存储格式:l,m,nl:室m:厅n:卫</span></td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">面积</td>
              <td> <input type="text" name="Area" size="40" value=""> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">楼层</td>
              <td> <input type="text" name="Floor" size="40" value=""> <span class="tx">楼层,存储格式:m,nm:总层n:第几层</span></td>
            </tr>
            <tr  class="hback"> 
              <td align="right">配套设施</td>
              <td> <input type="text" name="equip" size="40" value=""> <span class="tx">保存格式:、l,m,n,x,y,zl:通水m:电n:气x:电话y:光纤z:表示宽带1表示有,0表示无</span></td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">装修情况</td>
              <td> <select name="Decoration">
                  <option value="">所有</option>
                  <option value="1">简单装修</option>
                  <option value="2">中档装修</option>
                  <option value="3">高档装修</option>
                </select> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">联系人</td>
              <td> <input type="text" name="LinkMan" size="40" value=""> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">有效期</td>
              <td> <select name="Period">
                  <option value="">所有</option>
                  <option value="一周">一周</option>
                  <option value="两周">两周</option>
                  <option value="三周">三周</option>
                  <option value="一月">一月</option>
                  <option value="两月">两月</option>
                </select> <span class="tx">有效期:一周,两周,三周,一月,两月),所有只保留三月</span> 
              </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">发布日期</td>
              <td> <input type="text" name="PubDate" size="40" value=""> <span class="tx">可以用“>2006-08”的形式</span> 
              </td>
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






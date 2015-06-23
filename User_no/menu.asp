<script language="javascript" src="../FS_INC/prototype.js"></script>
<script language=javascript>
function status(cat)
{
	switch(cat)
	{
		case "menuMsg":
			if ('<%=Request.Cookies("FoosunUserCookies")("FoosunMenuMsg")%>'=='1')
			{
				 $(cat).style.display="";
			}else
			{
				 $(cat).style.display="none";
			};break;
		case "menuFriend":
			if ('<%=Request.Cookies("FoosunUserCookies")("FoosunMenuFriend")%>'=='1')
			{
				 $(cat).style.display="";
			}else
			{
				 $(cat).style.display="none";
			};break;
		case "menuSpecial":
			if ('<%=Request.Cookies("FoosunUserCookies")("FoosunMenuSpecial")%>'=='1')
			{
				 $(cat).style.display="";
			}else
			{
				 $(cat).style.display="none";
			};break;
	}
}
function opencat(cat)
{
  if($(cat).style.display=="none")
  {
     $(cat).style.display="";
	 switch(cat)
	 {
		case "menuMsg":new Ajax.Request("MenuAction.asp",{method:'get',parameters:"type=msg&value=1&rnd=+"+Math.random()});break
		case "menuFriend":new Ajax.Request("MenuAction.asp",{method:'get',parameters:"type=friend&value=1&rnd=+"+Math.random()});break
		case "menuSpecial":new Ajax.Request("MenuAction.asp",{method:'get',parameters:"type=special&value=1&rnd=+"+Math.random()});break
	 }
  } else 
  {
     $(cat).style.display="none"; 
	 switch(cat)
	 {
		case "menuMsg":new 		Ajax.Request("MenuAction.asp",{method:'get',parameters:"type=msg&value=0&rnd=+"+Math.random()});break
		case "menuFriend":new Ajax.Request("MenuAction.asp",{method:'get',parameters:"type=friend&value=0&rnd=+"+Math.random()});break
		case "menuSpecial":new Ajax.Request("MenuAction.asp",{method:'get',parameters:"type=special&value=0&rnd=+"+Math.random()});break
	 }
  }
}

</script> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<table width="170" border="0" align="center" cellpadding="3" cellspacing="0"  class="table">
  <tr class="hback"> 
    <td valign="top"  id="menuMsg_main" style="CURSOR: hand"  onmouseup="opencat('menuMsg');" onMouseOver="this.className='bg'" onMouseOut="this.className='bg1'" language=javascript><strong><img src="<%=s_savepath%>/images/folderclosed.gif"></strong><font style="font-size:14px;font-weight: bold">站内消息</font></td>
  </tr>
  <tr  id="menuMsg" style="DISPLAY:none " class="hback"> 
    <td> <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
          <td width="16%"><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
          <td width="84%"><a href="<%=s_savepath%>/message_write.asp">撰写消息</a></td>
        </tr>
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
          <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
          <td><a href="<%=s_savepath%>/Message_box.asp?type=rebox">收件箱</a></td>
        </tr>
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
          <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
          <td><a href="<%=s_savepath%>/Message_box.asp?type=drabox">草稿箱</a><a href="award.asp"></a></td>
        </tr>
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
          <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
          <td><a href="<%=s_savepath%>/Message_box.asp?type=sendbox">发件箱</a></td>
        </tr>
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
          <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
          <td><a href="<%=s_savepath%>/book.asp?M_Type=0">我的留言</a></td>
        </tr>
      </table></td>
  </tr>
  <tr class="hback"> 
    <td valign="top"   id="menuFriend_main" style="CURSOR: hand"  onmouseup="opencat('menuFriend');" onMouseOver="this.className='bg'" onMouseOut="this.className='bg1'" language=javascript><strong><img src="<%=s_savepath%>/images/folderclosed.gif"></strong><font style="font-size:14px;font-weight: bold">朋友圈</font></td>
  </tr>
  <tr id="menuFriend" style="display:none" class="hback"> 
    <td> <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>  
          <td width="16%"><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
          <td width="84%"><a href="<%=s_savepath%>/Friend.asp?type=0">好朋友</a>｜<a href="<%=s_savepath%>/Friend_add.asp?type=0">添加</a></td>
        </tr>
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>  
          <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
          <td><a href="<%=s_savepath%>/Friend.asp?type=1">陌生人</a>｜<a href="<%=s_savepath%>/Friend_add.asp?type=1">添加</a></td>
        </tr>
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>  
          <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
          <td><a href="<%=s_savepath%>/Friend.asp?type=2">黑名单</a>｜<a href="<%=s_savepath%>/Friend_add.asp?type=2">添加</a></td>
        </tr>
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
          <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
          <td><a href="<%=s_savepath%>/Corp_card.asp">名片管理</a></td>
        </tr>
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>  
          <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
          <td><a href="<%=s_savepath%>/Group.asp">社群讨论</a></td>
        </tr>
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>  
          <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
          <td><a href="<%=s_savepath%>/Groupclass.asp">社群分类</a></td>
        </tr>
      </table></td>
  </tr>
  <tr class="hback"> 
    <td valign="top"  id="menuSpecail_main" style="CURSOR: hand"  onmouseup="opencat('menuSpecial');" onMouseOver="this.className='bg'" onMouseOut="this.className='bg1'" language=javascript><strong><img src="<%=s_savepath%>/images/folderclosed.gif"></strong><font style="font-size:14px;font-weight: bold">我的专栏</font></td>
  </tr>
  <tr  id="menuSpecial" style="DISPLAY:none" class="hback"> 
    <td> <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
          <td width="15%"><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
          <td width="85%" valign="bottom"><a href="<%=s_savepath%>/favorite.asp">我的收藏夹</a></td>
        </tr>
		<%if menu_RssFeed=1 then%>
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
          <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
          <td valign="bottom"><a href="<%=s_savepath%>/RssFeed.asp">RSS聚合订阅</a></td>
        </tr>
		<%end if%>
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
          <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
          <td valign="bottom"><a href="<%=s_savepath%>/PhotoManage.asp">相册管理</a></td>
        </tr>
        <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
          <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
          <td valign="bottom"><a href="<%=s_savepath%>/Filemanage.asp">文件管理</a></td>
        </tr>
        <tr> 
          <td  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(id10);" onMouseOver="this.className='bg'" onMouseOut="this.className='bg1'" language=javascript><div align="right"><strong><img src="<%=s_savepath%>/images/folderclosed.gif"></strong></div></td>
          <td valign="bottom"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(id10);" onMouseOver="this.className='bg'" onMouseOut="this.className='bg1'" language=javascript>我的专栏</td>
        </tr>
        <tr> 
          <td colspan="2" valign="top"   id="id10"> <div align="right"> 
              <table width="100%" border="0" cellspacing="1" cellpadding="2">
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
                  <td width="18%"><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif"></div></td>
                  <td width="82%" valign="bottom"><a href="<%=s_savepath%>/infomanage.asp">专栏分类</a></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
                  <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif"></div></td>
                  <td valign="bottom"><a href="<%=s_savepath%>/contr/contrManage.asp">稿件管理</a></td>
                </tr>
				<%if IsExist_SubSys("HS") Then%>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
                  <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif"></div></td>
                  <td valign="bottom"><a href="<%=s_savepath%>/house/houseManage.asp">房产管理</a></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
                  <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif"></div></td>
                  <td valign="bottom"><a href="<%=s_savepath%>/house/default.asp">房产搜索</a></td>
                </tr>
				<%end if%>
				<%if IsExist_SubSys("SD") Then%>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
                  <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
                  <td valign="bottom"><a href="<%=s_savepath%>/Supply/PublicManage.asp">供求信息</a></td>
                </tr>
				<%End if%>
				<%if IsExist_SubSys("AP") Then%>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
                  <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
                  <td valign="bottom"><a href="<%=s_savepath%>/job/job_resume.asp">求职</a></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
                  <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
                  <td valign="bottom"><a href="<%=s_savepath%>/job/job_applications.asp">招聘</a></td>
                </tr>
				<%end if%>
                <!--<tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
                  <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
                  <td valign="bottom"><a href="<%=s_savepath%>/download/">下载</a></td>
                </tr>-->
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
                  <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
                  <td valign="bottom"><a href="<%=s_savepath%>/i_Blog/index.asp">计划(日志)</a></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
                  <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" ></div></td>
                  <td valign="bottom"><a href="<%=s_savepath%>/review.asp">我的评论</a></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td><div align="right"><img src="<%=s_savepath%>/images/folderopened.gif" alt="award"></div></td>
                  <td valign="bottom">
				  <a href="<%=s_savepath%>/award/award.asp">抽奖<span id="img_container"></span></a>
				  </td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
                  <td colspan="2"><div align="left"><img src="<%=s_savepath%>/images/folderopened.gif" ><span><a href="javascript:OpenEditerWindow('<%=s_savepath%>/Flink/Flink_add.asp','window','520','420');" style="cursor:hand;">申请友情连接</a></span></div></td>
                </tr>
              </table>
            </div></td>
        </tr>
      </table></td>
  </tr>
</table>
<div align="center">
  <script language="JavaScript" src="<%=s_savepath_1%>FS_Inc/PublicJS.js" type="text/JavaScript"></script>
  <script language="javascript" src="<%=s_savepath_1%>FS_Inc/prototype.js"></script>
  <script language="JavaScript" type="text/JavaScript">
	/***************************************/
	var  myAjax = new Ajax.Updater('img_container', "<%="http://"&request.ServerVariables("HTTP_HOST")&s_savepath&"/award/awardAction.asp"%>", {method: 'get', parameters: "action=menu"});//是否有激活的抽奖
	/*****************************************/		
		var menuOffX=0	//菜单距连接文字最左端距离
		var menuOffY=18	//菜单距连接文字顶端距离
		var fo_shadows=new Array()
		var linkset=new Array()
		
		var ie4=document.all&&navigator.userAgent.indexOf("Opera")==-1
		var ns6=document.getElementById&&!document.all
		var ns4=document.layers
		
		function showmenu(e,index,p,paging){
			if (!document.all&&!document.getElementById&&!document.layers)
				return
			which=linkset[index]
			var pSize=10	//每页连接数
			var pNum=Math.floor((which.length-1)/pSize)+1		//页数
			
			clearhidemenu()
			ie_clearshadow()
			if (pNum==1){
				which=which.join("")
			}else{
				which=which.slice( (p-1)*pSize, p*pSize )
				which=which.join("")
				which+="<div class=\"menuitems\" >"
				if (p==1)
				{
					which+="&nbsp;&nbsp;&nbsp;&nbsp;<font face=webdings color=gray>7</font>"
				}else{
					which+="&nbsp;&nbsp;&nbsp;&nbsp;<font face=webdings style=cursor:hand onclick=showmenu(event,"+ index +","+ (p-1) +",true) >7</font>"
				}
				if (p==pNum)
				{
					which+="<font face=webdings color=gray>8</font>"
				}else{
					which+="<font face=webdings style=cursor:hand onclick=showmenu(event,"+ index +","+ (p+1) +",true) >8</font>"
				}
				which+="</div>"
			}
			
			menuobj=ie4? document.all.popmenu : ns6? document.getElementById("popmenu") : ns4? document.popmenu : ""
			menuobj.thestyle=(ie4||ns6)? menuobj.style : menuobj
			
			if (ie4||ns6)
				menuobj.innerHTML=which
			else{
				menuobj.document.write('<layer name=gui bgColor=#E6E6E6 width=165 onmouseover="clearhidemenu()" onmouseout="hidemenu()">'+which+'</layer>')
				menuobj.document.close()
			}
			menuobj.contentwidth=(ie4||ns6)? menuobj.offsetWidth : menuobj.document.gui.document.width
			menuobj.contentheight=(ie4||ns6)? menuobj.offsetHeight : menuobj.document.gui.document.height
			
			eventX=ie4? event.clientX : ns6? e.clientX : e.x
			eventY=ie4? event.clientY : ns6? e.clientY : e.y
			
			var rightedge=ie4? document.body.clientWidth-eventX : window.innerWidth-eventX
			var bottomedge=ie4? document.body.clientHeight-eventY : window.innerHeight-eventY
			
			
			if (!paging)
			{	
				if (rightedge<menuobj.contentwidth)
					menuobj.thestyle.left=ie4? document.body.scrollLeft+eventX-menuobj.contentwidth+menuOffX : ns6? window.pageXOffset+eventX-menuobj.contentwidth : eventX-menuobj.contentwidth
				else
					menuobj.thestyle.left=ie4? ie_x(event.srcElement)+menuOffX : ns6? window.pageXOffset+eventX : eventX
				
				if (bottomedge<menuobj.contentheight)
					menuobj.thestyle.top=ie4? document.body.scrollTop+eventY-menuobj.contentheight-event.offsetY+menuOffY : ns6? window.pageYOffset+eventY-menuobj.contentheight : eventY-menuobj.contentheight
				else
					menuobj.thestyle.top=ie4? ie_y(event.srcElement)+menuOffY : ns6? window.pageYOffset+eventY : eventY
			}
				
			menuobj.thestyle.visibility="visible"
			ie_dropshadow(menuobj,"#CCCCCC",3)
			return false
		}
		function ie_x(e){  
			var l=e.offsetLeft;  
			while(e=e.offsetParent){  
				l+=e.offsetLeft;  
			}  
			return l;  
		}  
		
		
		function ie_y(e){  
			var t=e.offsetTop;  
			while(e=e.offsetParent){  
				t+=e.offsetTop;  
			}  
			return t;  
		}  
		
		function ie_dropshadow(el, color, size)
		{
			var i;
			for (i=size; i>0; i--)
			{
				var rect = document.createElement('div');
				var rs = rect.style
				rs.position = 'absolute';
				rs.left = (el.style.posLeft + i) + 'px';
				rs.top = (el.style.posTop + i) + 'px';
				rs.width = el.offsetWidth + 'px';
				rs.height = el.offsetHeight + 'px';
				rs.zIndex = el.style.zIndex - i;
				rs.backgroundColor = color;
				var opacity = 1 - i / (i + 1);
				rs.filter = 'alpha(opacity=' + (100 * opacity) + ')';
				//el.insertAdjacentElement('afterEnd', rect);		// 输出阴影，若不想要投影，可以注释此句
				fo_shadows[fo_shadows.length] = rect;
			}
		}
		function ie_clearshadow()
		{
			for(var i=0;i<fo_shadows.length;i++)
			{
				if (fo_shadows[i])
					fo_shadows[i].style.display="none"
			}
			fo_shadows=new Array();
		}
		
		
		function contains_ns6(a, b) {
			while (b.parentNode)
				if ((b = b.parentNode) == a)
					return true;
			return false;
		}
		
		function hidemenu(){
			if (window.menuobj)
				menuobj.thestyle.visibility=(ie4||ns6)? "hidden" : "hide"
			ie_clearshadow()
		}
		
		function dynamichide(e){
			if (ie4&&!menuobj.contains(e.toElement))
				hidemenu()
			else if (ns6&&e.currentTarget!= e.relatedTarget&& !contains_ns6(e.currentTarget, e.relatedTarget))
				hidemenu()
		}
		
		function delayhidemenu(){
			if (ie4||ns6||ns4)
				delayhide=setTimeout("hidemenu()",500)
		}
		
		function clearhidemenu(){
			if (window.delayhide)
				clearTimeout(delayhide)
		}
		
		function highlightmenu(e,state){
			if (document.all)
				source_el=event.srcElement
			else if (document.getElementById)
				source_el=e.target
			if (source_el.className=="menuitems"){
				source_el.id=(state=="on")? "mouseoverstyle" : ""
			}
			else{
				while(source_el.id!="popmenu"){
					source_el=document.getElementById? source_el.parentNode : source_el.parentElement
					if (source_el.className=="menuitems"){
						source_el.id=(state=="on")? "mouseoverstyle" : ""
					}
				}
			}
		}
		
		if (ie4||ns6)
		document.onclick=hidemenu
		
		linkset[0]=new Array()
		linkset[0][0]='<div class=menuitems><a class="menu_ctr" href="<%=s_savepath%>/UserPara.asp">站点参数</a></div>'
		linkset[0][1]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/UserList.asp>用户统计</a></div>'
		linkset[0][2]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/callboard.asp>会员公告</a></div>'
		linkset[0][3]='<div class=menuitems><a class="</div>'
		linkset[1]=new Array()
		linkset[1][0]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/myinfo.asp>基本信息</a></div>'
		linkset[1][1]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/MyContact.asp>联系方式</a></div>'
		linkset[1][2]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/Myvalidcode.asp>安全资料</a></div>'
		linkset[1][3]='<div class=menuitems><a class="menu_ctr" >----------</a></div>'
		linkset[1][4]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/Myaccount.asp>我的保险箱</a></div>'
		linkset[1][5]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/award/award.asp>抽奖/问答</a></div>'
		//linkset[1][7]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/award.asp?action=Change>积分兑奖</a></div>'
		//linkset[1][8]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/Filemanage.asp>文件管理</a></div>'
		
		linkset[2]=new Array()
		linkset[2][0]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/corp_info.asp>企业资料</a></div>'
		linkset[2][1]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/corp_templet.asp>模板管理</a></div>'
		linkset[2][2]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/corp_logo.asp>Logo管理</a></div>'
		linkset[2][3]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/corp_bank.asp>银行资料</a></div>'
		//linkset[2][4]='<div class=menuitems><a class="menu_ctr" href=corp_Out.asp>注销企业</a></div>'
		linkset[3]=new Array()
		linkset[3][0]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/buybag.asp>购物车</a></div>'
		linkset[3][1]='<div class=menuitems><a class="menu_ctr" >----------</a></div>'
		linkset[3][2]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/card.asp>点卡冲值</a></div>'
		linkset[3][3]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/pay.asp>在线冲值</a></div>'
		//linkset[3][4]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/buypop.asp>购买权限</a></div>'
		linkset[3][5]='<div class=menuitems><a class="menu_ctr" >----------</a></div>'
		linkset[3][6]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/order.asp>定单管理</a></div>'
		linkset[3][7]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/history.asp>交易明晰</a></div>'
		linkset[3][8]='<div class=menuitems><a class="menu_ctr" href=<%=s_savepath%>/get_Thing.asp>获取商品</a></div>'
		function imgzoom(img,maxsize){
			var a=new Image();
			a.src=img.src
			if(a.width > maxsize * 4)
			{
				img.style.width=maxsize;
			}
			else if(a.width >= maxsize)
			{
				img.style.width=Math.round(a.width * Math.floor(4 * maxsize / a.width) / 4);
			}
			return false;
		}
		function zoom_img(e, o)
		{
		var zoom = parseInt(o.style.zoom, 10) || 100;
		zoom += event.wheelDelta / 12;
		if (zoom > 0) o.style.zoom = zoom + '%';
		return false;
		}  
/*0-------------------------------------*/
status('menuMsg');
status('menuFriend');
status('menuSpecial');
</script>
  <script language="JavaScript" src="<%=s_savepath_1%>FS_Inc/CheckJs.js" type="text/JavaScript"></script>
  <script language="JavaScript" src="<%=s_savepath_1%>FS_Inc/Prototype.js" type="text/JavaScript"></script>
  <script language="JavaScript" src="<%=s_savepath_1%>FS_Inc/PublicJS.js" type="text/JavaScript"></script>
</div>







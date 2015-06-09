define("eventUtil",[],function(){function e(){var e=event||window.event;e&&e.preventDefault?e.preventDefault():e.returnValue=!1}function t(){var e=event||window.event,t=e.srcElement||e.target;e&&e.stopPropagation?e.stopPropagation():e.cancelBubble=!0}function n(e,t,n){e.addEventListener?e.addEventListener(t,n,!1):e.attachEvent?e.attachEvent("on"+t,n):e["on"+t]=n}function r(e,t,n){e.removeEventListener?e.removeEventListener(t,n,!1):e.detachEvent&&e.detachEvent("on"+t,n)}function i(e){var t=window.onload;typeof window.onload!="function"?window.onload=e:window.onload=function(){t(),e()}}function s(e){var t=window.onscroll;typeof window.onscroll!="function"?window.onscroll=e:window.onscroll=function(){t(),e()}}function o(e){var t=window.onresize;typeof window.onresize!="function"?window.onresize=e:window.onresize=function(){t(),e()}}return{preventDefault:e,stopPropagation:t,addHandler:n,removeHandler:r,addLoadEvent:i,addScrollEvent:s,addResizeEvent:o}}),define("ScrollNav",[],function(){function i(){e=document.documentElement.scrollTop||document.body.scrollTop,t=document.documentElement.clientWidth||document.body.clientWidth,n=e-r.offsetTop,n>10?(r.style.position="fixed",r.style.top="0px",t<940?r.style.left="0px":r.style.left=(t-r.clientWidth)/2+"px"):n<=0&&(r.style.position="relative",r.style.left="0px")}function s(){t=document.documentElement.clientWidth||document.body.clientWidth,t>940?r.style.width="100%":r.style.width="940px"}var e=null,t=null,n=null,r=document.getElementById("navigate");return{fixNav:i,resizeNav:s}}),define("AskOnline",["eventUtil"],function(e){function t(){var t=document.getElementsByName("askOnline")||null;if(t!=null){var n="http://qiao.baidu.com/v3/?module=default&controller=im&action=index&ucid=7225288&type=n&siteid=4628760",r="在线咨询",i=800,s=550,o=(window.screen.height-s)/2,u=(window.screen.width-i)/2,a="height="+s+",width="+i+",top="+o+",left="+u+",toolbar=no,menubar=no,scrollbars=no, resizable=no,location=no,status=no";for(var f=0;f<t.length;f++)e.addHandler(t.item(f),"click",function(t){e.preventDefault(),window.open(n,r,a)})}}return{askOnline:t}}),define("utilities",[],function(){function e(e,t){try{if(e instanceof Array){if(e.length>0)for(var n=0;n<e.length;n++)e[n].style.display=t}else e!=null&&(e.style.display=t)}catch(r){console("argument must Input a document element:"+r.name+":"+r.message)}}function t(e,t){var n=document.getElementsByTagName("head")[0],r=document.createElement("script");r.src=e,r.onload=r.onreadystatechange=function(){(!this.readyState||this.readyState=="loaded"||this.readyState=="complete")&&t()},n.appendChild(r)}return{displayElts:e,proxyResponse:t}}),define("cookie",[],function(){function e(t,n){var r=e.arguments,i=e.arguments.length,s=i>2?r[2]:null;if(s!=null){var o=new Date;o.setTime(o.getTime()+s*1e3*3600*24)}document.cookie=t+"="+escape(n)+(s==null?"":"; expires="+o)}function t(e){var t=e+"=";if(document.cookie.length>0){var n=document.cookie.indexOf(t);if(n!=-1){n+=t.length;var r=document.cookie.indexOf(";",n);return r==-1&&(r=document.cookie.length),unescape(document.cookie.substring(n,r))}return""}}function n(e){var n=t(e);if(n!=null){var r=new Date;r.setTime(r.getTime()-1),document.cookie=e+"="+n+";expires="+r.toGMTString()}}return{setCookie:e,getCookie:t,delCookie:n}}),define("ajax",[],function(){function e(e,t){var n=null;window.XMLHttpRequest?n=new XMLHttpRequest:window.ActiveXObject&&(n=new ActiveXObject("Microsoft.XMLHTTP")),n!=null?(n.onreadystatechange=function(){n.readyState===4&&(n.status===200?t.innerHTML=n.responseText:console.log("The data request failed:"+n.statusText))},n.open("GET",e,!0),n.send(null)):console.log("Your browser does not support this function")}return{loadXMLDoc:e}}),define("modal",["eventUtil","utilities"],function(e,t){function s(e){t.displayElts(n,e)}var n=document.getElementById("modalWindow"),r=document.getElementById("modalContent"),i=document.getElementById("modalCloser");return e.addHandler(i,"click",function(e){t.displayElts(n,"none")}),{oModalContent:r,displayModalWindow:s}}),define("agencyAnchor",["eventUtil","ajax","modal"],function(e,t,n){function i(e){for(var t=0;t<r.length;t++)if(r[t][0]==e)return r[t][1]}function s(r){var s=i(r),o=document.getElementsByName("contact"),u;for(u=0;u<o.length;u++)o[u].setAttribute("href","http://www.saiving.com/about/contact/index.html#"+s),e.addHandler(o[u],"click",function(){e.preventDefault(),t.loadXMLDoc("http://www.saiving.com/contactHtml/"+s+".txt",n.oModalContent),n.displayModalWindow("block")})}var r=[["北京","beijing"],["北京市","beijing"],["河北","beijing"],["河北省","beijing"],["天津","beijing"],["天津市","beijing"],["内蒙古","ruimong"],["内蒙古自治区","ruimong"],["新疆","xinjiang"],["新疆自治区","xinjiang"],["甘肃","ganshu"],["甘肃省","ganshu"],["陕西","xianxi"],["陕西省","xianxi"],["青海","qinghai"],["青海省","qinghai"],["辽宁","dongshanxin"],["辽宁省","dongshanxin"],["吉林","dongshanxin"],["吉林省","dongshanxin"],["黑龙江","dongshanxin"],["黑龙江省","dongshanxin"],["宁夏","ganshu"],["宁夏自治区","ganshu"],["河南","henan"],["河南省","henan"],["山东","shandong"],["山东省","shandong"],["山西","shanxi"],["山西省","shanxi"],["四川","sichung"],["四川省","sichung"],["重庆","sichung"],["重庆市","sichung"],["江苏","jiangshu"],["江苏省","jiangshu"],["湖北","hubei"],["湖北省","hubei"]["安徽","anhui"],["安徽省","anhui"],["上海","shanghai"],["上海市","shanghai"]["浙江","jiejiang"],["浙江省","jiejiang"]["湖南","hunan"],["湖南省","hunan"],["福建","fujian"],["福建省","fujian"],["江西","jiangxi"],["江西省","jiangxi"],["广西","guangxi"],["广西省","guangxi"],["广东","guangdong"],["广东省","guangdong"],["贵州","guizhou"],["贵州省","guizhou"],["云南","yunnan"],["云南省","yunnan"],["海南","guangdong"],["海南省","guangdong"],["西藏","qinghai"],["西藏自治区","qinghai"],["台湾","beijing"],["台湾省","beijing"],["香港","beijing"],["香港特别行政区","beijing"],["澳门","beijing"],["澳门特别行政区","beijing"]];return{setAnchor:s}}),define("agencyDetail",[],function(){function n(n){switch(n){case"北京":case"北京市":e="北京总部";break;case"河北":case"河北省":case"天津":case"天津市":case"内蒙古":case"内蒙古自治区":case"新疆":case"新疆自治区":case"甘肃":case"甘肃省":case"陕西":case"陕西省":case"青海":case"青海省":case"辽宁":case"辽宁省":case"吉林":case"吉林省":case"黑龙江":case"黑龙江省":case"宁夏":case"宁夏自治区":case"河南":case"河南省":case"山东":case"山东省":case"山西":case"山西省":case"四川":case"四川省":case"重庆":case"重庆市":case"江苏":case"江苏省":case"湖北":case"湖北省":case"安徽":case"安徽省":case"上海":case"上海市":case"浙江":case"浙江省":case"湖南":case"湖南省":case"福建":case"福建省":case"江西":case"江西省":case"广西":case"广西省":case"广东":case"广东省":case"贵州":case"贵州省":case"云南":case"云南省":case"海南":case"海南省":case"西藏":case"西藏自治区":e=n+"办事处";break;case"台湾":case"台湾省":case"香港":case"香港特别行政区":case"澳门":case"澳门特别行政区":e="港澳台及海外办事处";break;default:e="北京总部"}t.innerHTML=e}var e="",t=document.getElementById("agency");return{showAgency:n}}),define("agency",["utilities","cookie","agencyAnchor","agencyDetail"],function(e,t,n,r){var i=t.getCookie("agency");if(!i||i=="")try{e.proxyResponse("http://int.dpool.sina.com.cn/iplookup/iplookup.php?format=js",function(){var e="";remote_ip_info["province"]==null||remote_ip_info["province"]=="保留地址"||remote_ip_info["province"]==""?e="北京":e=remote_ip_info.province,n.setAnchor(e),r.showAgency(e),t.setCookie("agency",e,10)})}catch(s){e.proxyResponse("http://pv.sohu.com/cityjson",function(){var e=returnCitySN.cname;n.setAnchor(e),r.showAgency(e),t.setCookie("agency",e,10)})}else n.setAnchor(i),r.showAgency(i)}),define("backToTop",["eventUtil"],function(e){function u(){e.addScrollEvent(function(){r=document.documentElement.clientHeight||document.body.clientHeight,n=document.documentElement.scrollTop||document.body.scrollTop,n>=r?t.style.display="block":t.style.display="none",o||clearInterval(s),o=!1})}function a(){e.addHandler(t,"click",function(){clearInterval(s),s=setInterval(function(){n=document.documentElement.scrollTop||document.body.scrollTop,i=Math.ceil(n/5),document.body.scrollTop-=i,document.documentElement.scrollTop-=i,o=!0,n<=0&&(clearInterval(s),document.body.scrollTop=0,document.documentElement.scrollTop=0)},2)})}var t=document.getElementById("backToTop"),n,r,i=0,s=null,o=!0;return{stopScroll:u,backToTop:a}}),require.config({paths:{ajax:"ajax",eventUtil:"eventUtil",cookie:"cookie",utilities:"utilities",ScrollNav:"ScrollNav",AskOnline:"AskOnline",modal:"modal",agencyAnchor:"agencyAnchor",agencyDetail:"agencyDetail",agency:"agency",backToTop:"backToTop"}}),require(["eventUtil","ScrollNav","AskOnline","agency","backToTop"],function(e,t,n,r,i){e.addScrollEvent(t.fixNav),e.addResizeEvent(t.resizeNav),n.askOnline(),i.stopScroll(),i.backToTop()}),define("main",function(){});
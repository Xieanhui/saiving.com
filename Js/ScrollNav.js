define(function(){//固定顶部导航
    var scrollTop = null;
    var clientWidth = null;
    var pos = null;
    var nav = document.getElementById("navigate");

    function fixNav() {
        scrollTop = document.documentElement.scrollTop || document.body.scrollTop;
        clientWidth = document.documentElement.clientWidth || document.body.clientWidth;
        pos = scrollTop - nav.offsetTop;

        if (pos > 10) {
            nav.style.position = "fixed";
            nav.style.top = "0px";
            if (clientWidth < 940)
                nav.style.left = "0px";
            else
                nav.style.left = (clientWidth - nav.clientWidth)/2 + "px";
        }
        else if (pos <= 0) {
            nav.style.position = "relative";
            nav.style.left = "0px";
        }
    }

    function resizeNav() {//缩放窗口调整网页宽度
        clientWidth = document.documentElement.clientWidth || document.body.clientWidth;
        if (clientWidth > 940) {
            nav.style.width = "100%";
        } else {
            nav.style.width = "940px";
        }
    }

    return {
        fixNav : fixNav,
        resizeNav: resizeNav
    };
});
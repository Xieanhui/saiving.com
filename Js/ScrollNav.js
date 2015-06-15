define(['utilities'],function(utilities) { //固定顶部导航
    var scrollTop,
        pos,
        nav = utilities.g("navigate");

    function fixNav() {
        scrollTop = document.documentElement.scrollTop || document.body.scrollTop;
       // clientWidth = document.documentElement.clientWidth || document.body.clientWidth;
        pos = scrollTop - nav.offsetTop;

        if (pos > 10) {
            nav.style.position = "fixed";
            nav.style.top = "0px";
        } else if (pos <= 10) {
            nav.style.position = "relative";
        }
    }

    return {
        fixNav: fixNav
    };
});

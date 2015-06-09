require.config({

    paths: {
        'ajax': 'ajax',
        'eventUtil': 'eventUtil',
        'cookie': 'cookie',
        'utilities': 'utilities',
        'ScrollNav': 'ScrollNav',
        'AskOnline': 'AskOnline',
        'modal': 'modal',
        'agencyAnchor': 'agencyAnchor',
        'agencyDetail': 'agencyDetail',
        'agency': 'agency',
        'backToTop': 'backToTop'
    }

});

require(['eventUtil', 'ScrollNav', 'AskOnline', 'agency', 'backToTop'], function(eventUtil, ScrollNav, AskOnline, agency, backToTop) {

    eventUtil.addScrollEvent(ScrollNav.fixNav); //固定导航
    eventUtil.addResizeEvent(ScrollNav.resizeNav); //调整理窗口时缩放导航尺寸

    AskOnline.askOnline(); //点击打开在线咨询

    backToTop.stopScroll(); //停止返回页面顶部
    backToTop.backToTop(); //返回页面顶部

});

require.config({
    baseUrl: '../js',
    paths: {
        'ajax': '../../js/ajax',
        'eventUtil': '../../js/eventUtil',
        'cookie': '../../js/cookie',
        'Browser': '../../js/Browser',
        'utilities': '../../js/utilities',
        'AskOnline': '../../js/AskOnline',
        'modal': '../../js/modal',
        'agencyAnchor': '../../js/agencyAnchor',
        'agencyDetail': '../../js/agencyDetail',
        'agency': '../../js/agency',
        'condensate_slide': 'condensate_slide'
    }

});

require(['eventUtil', 'AskOnline', 'agency', 'condensate_slide'], function(eventUtil, AskOnline, agency, condensate_slide) {

    AskOnline.askOnline(); //点击打开在线咨询

    eventUtil.onDOMContentLoaded(condensate_slide.slide); //冷凝水回收设备页内滚动图片

});

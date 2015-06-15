require.config({

    paths: {
        'ajax': 'ajax',
        'eventUtil': 'eventUtil',
        'cookie': 'cookie',
        'Browser': 'Browser',
        'utilities': 'utilities',
        'ScrollNav': 'ScrollNav',
        'AskOnline': 'AskOnline',
        'modal': 'modal',
        'agencyAnchor': 'agencyAnchor',
        'agencyDetail': 'agencyDetail',
        'agency': 'agency',
        'condensate_slide': 'condensate_slide'
    }

});

require(['eventUtil', 'ScrollNav', 'AskOnline', 'agency', 'condensate_slide'], function(eventUtil, ScrollNav, AskOnline, agency, condensate_slide) {

    eventUtil.addScrollEvent(ScrollNav.fixNav); //固定导航

    AskOnline.askOnline(); //点击打开在线咨询

    eventUtil.addLoadEvent(condensate_slide.slide);

});

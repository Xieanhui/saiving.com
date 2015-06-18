require.config({
    baseUrl: '../js',
    paths: {
        'ajax': '../../js/ajax',
        'eventUtil': '../../js/eventUtil',
        'cookie': '../../js/cookie',
        'Browser': '../../js/Browser',
        'utilities': '../../js/utilities',
        'AskOnline': './AskOnline',
        'modal': '../../js/modal',
        'agencyAnchor': '../../js/agencyAnchor',
        'agencyDetail': '../../js/agencyDetail',
        'agency': '../../js/agency',
        'indexSlide': './indexSlide'
    }

});

require(['eventUtil', 'AskOnline', 'agency', 'indexSlide'], function(eventUtil, AskOnline, agency, indexSlide) {

    AskOnline.askOnline(); //点击打开在线咨询

    indexSlide.init();
    indexSlide.slide();
});

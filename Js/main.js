require.config({

    paths: {
        'ajax': 'ajax',
        'eventUtil': 'eventUtil',
        'cookie': 'cookie',
        'Browser': 'Browser',
        'utilities': 'utilities',
        'AskOnline': 'AskOnline',
        'modal': 'modal',
        'agencyAnchor': 'agencyAnchor',
        'agencyDetail': 'agencyDetail',
        'agency': 'agency'
    }

});

require(['eventUtil', 'AskOnline', 'agency'], function(eventUtil, AskOnline, agency) {

    AskOnline.askOnline(); //点击打开在线咨询

});

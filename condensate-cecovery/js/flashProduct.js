define(['utilities', 'eventUtil', 'animal'], function(utilities, eventUtil, animal) {

    var ofirstPage = utilities.getElementsByClassName(document, 'div', 'firstPage')[0];
    var arrMainProduct1 = utilities.getElementsByClassName(ofirstPage, 'ul', 'mainProduct')[0];
    var oseventhPage = utilities.getElementsByClassName(document, 'div', 'seventhPage')[0];
    var arrMainProduct2 = utilities.getElementsByClassName(oseventhPage, 'ul', 'mainProduct')[0];

    var scrollTop = document.documentElement.scrollTop || document.body.scrollTop,
        clientHeight = document.documentElement.clientHeight || document.body.clientHeight;

    if (ofirstPage.offsetTop - scrollTop <= clientHeight) {
        animal.move(arrMainProduct1, {
            'top': 280,
            'opacity': 100
        });
    }

    eventUtil.addScrollEvent(function() {
        scrollTop = document.documentElement.scrollTop || document.body.scrollTop;
        clientHeight = document.documentElement.clientHeight || document.body.clientHeight;

        if ((oseventhPage.offsetTop - scrollTop + 400) <= clientHeight) {
            animal.move(arrMainProduct2, {
                'top': 280,
                'opacity': 100
            });
        }

    });
});

define(['utilities', 'eventUtil', 'animal', 'onView'], function(utilities, eventUtil, animal, onView) {

    var ofirstPage = utilities.getElementsByClassName(document, 'div', 'firstPage')[0];
    var arrMainProduct1 = utilities.getElementsByClassName(ofirstPage, 'ul', 'mainProduct')[0];
    var arrProductLi1 = arrMainProduct1.getElementsByTagName('li');
    var oseventhPage = utilities.getElementsByClassName(document, 'div', 'seventhPage')[0];
    var arrMainProduct2 = utilities.getElementsByClassName(oseventhPage, 'ul', 'mainProduct')[0];
    var arrProductLi2 = arrMainProduct2.getElementsByTagName('li');

    var oSecondPage = utilities.getElementsByClassName(document, 'div', 'secondPage')[0];
    var oAssitProduct = oSecondPage.getElementsByTagName('ul')[0];

    var oSixthPage = utilities.getElementsByClassName(document, 'div', 'sixthPage')[0];
    var oCopartner = utilities.getElementsByClassName(oSixthPage, 'div', 'copartner')[0];

    var timer,
        k = 0,
        startMove = false;
    var scrollTop = document.documentElement.scrollTop || document.body.scrollTop,
        clientHeight = document.documentElement.clientHeight || document.body.clientHeight;

    if (ofirstPage.offsetTop - scrollTop <= clientHeight) {
        animal.move(arrProductLi1[0], {
            'top': 0,
            'opacity': 100
        }, 6, function() {
            animal.move(arrProductLi1[1], {
                'top': 0,
                'opacity': 100
            }, 6, function() {
                animal.move(arrProductLi1[2], {
                    'top': 0,
                    'opacity': 100
                }, 6, function() {
                    animal.move(arrProductLi1[3], {
                        'top': 0,
                        'opacity': 100
                    }, 6);
                });
            });
        });
    }

    onView.onView(oSecondPage, 500, function() {
        animal.move(oAssitProduct, {
            'top': 0,
            'opacity': 100
        }, 20);
    });

    onView.onView(oSixthPage, 600, function() {
        animal.move(oCopartner, {opacity:100, top: 0}, 20);
    });

    onView.onView(oseventhPage, 410, function() {
        animal.move(arrProductLi2[0], {
            'top': 0,
            'opacity': 100
        }, 6, function() {
            animal.move(arrProductLi2[1], {
                'top': 0,
                'opacity': 100
            }, 6, function() {
                animal.move(arrProductLi2[2], {
                    'top': 0,
                    'opacity': 100
                }, 6, function() {
                    animal.move(arrProductLi2[3], {
                        'top': 0,
                        'opacity': 100
                    }, 6);
                });
            });
        });
    });


});

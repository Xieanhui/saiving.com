define(['utilities', 'eventUtil', 'onView'], function(utilities, eventUtil, onView) {

    var oindustryNum = utilities.g('industryNum');
    var oenterpriseNum = utilities.g('enterpriseNum');
    var othirdPage = utilities.getElementsByClassName(document, "div", "thirdPage")[0];
    var i = 0,
        j = 0;

    onView.onView(othirdPage, 180, function() {
        othirdPage.timer1 = setInterval(function() {
            i += 1;
            if (i >= 20) {
                i = 20;
                clearInterval(othirdPage.timer1);
            }
            oindustryNum.innerHTML = i;
        }, 500);
    });

    onView.onView(othirdPage, 180, function() {
        othirdPage.timer2 = setInterval(function() {
            j += 2;
            if (j >= 500) {
                j = 500;
                clearInterval(othirdPage.timer2);
            }
            oenterpriseNum.innerHTML = j;
        }, 40);

    });
});

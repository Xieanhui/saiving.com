define(['utilities', 'eventUtil'], function(utilities, eventUtil) {

    var oindustryNum = utilities.g('industryNum');
    var oenterpriseNum = utilities.g('enterpriseNum');
    var othirdPage = utilities.getElementsByClassName(document, "div", "thirdPage")[0];
    var timer1,
        timer2,
        i = 0,
        j = 0,
        k = 0,
        scrollTop = 0,
        clientHeight = 0,
        startCount = false;

    eventUtil.addHandler(window, 'scroll', isScroll);

    function isScroll() {
         scrollTop = document.documentElement.scrollTop || document.body.scrollTop;
        clientHeight = document.documentElement.clientHeight || document.body.clientHeight;
        k = othirdPage.offsetTop - clientHeight;
        if (k < scrollTop) {
            startCount = true;
            eventUtil.removeHandler(window, 'scroll', isScroll);
        }
    }

    timer1 = setInterval(function() {
        if (startCount == true) {
            i += 1;
            if (i >= 20) {
                i = 20;
                clearInterval(timer1);
            }
            oindustryNum.innerHTML = i;
        }
    }, 500);

    timer2 = setInterval(function() {
        if (startCount == true) {
            j += 1;
            if (j >= 500) {
                j = 500;
                clearInterval(timer2);
            }
            oenterpriseNum.innerHTML = j;
        }
    }, 30);

});

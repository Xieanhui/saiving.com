define(['eventUtil'], function(eventUtil) {

    function onView(el, k, fn) {
        var startMove = false,
            scrollTop = 0,
            clientHeight = 0,
            i = 0,
            timer;

        eventUtil.addHandler(window, 'scroll', isScroll);

        function isScroll() {
            scrollTop = document.documentElement.scrollTop || document.body.scrollTop;
            clientHeight = document.documentElement.clientHeight || document.body.clientHeight;
            i = el.offsetTop - scrollTop + k;
            if (i <= clientHeight) {
                startMove = true;                
                eventUtil.removeHandler(window, 'scroll', isScroll);
            }
        }

        timer = setInterval(function() {
            if (startMove == true) {
            	clearInterval(timer);
            	startMove = false;
                if (fn) fn();
            }
        }, 1);

    }

    return {
        onView: onView
    }

});

define(['utilities'], function(utilities) {

    function move(el, json, duration, fn) {
        clearInterval(el.timer);
        var flag;
        el.timer = setInterval(function() {
            flag = true;
            for (var attr in json) {
                var iCurr = 0;
                if (attr == 'opacity') {
                    iCurr = Math.round(parseFloat(utilities.getStyle(el, attr)) * 100);
                } else {
                    iCurr = parseInt(utilities.getStyle(el, attr));
                }
                var speed = (json[attr] - iCurr) / 10;
                speed = speed > 0 ? Math.ceil(speed) : Math.floor(speed);
                if (iCurr != json[attr]) {
                    flag = false;
                }
                if (attr == 'opacity') {
                    el.style.opacity = (iCurr + speed) / 100;
                    el.style.filter = 'alpha(opacity=' + (iCurr + speed) + ')';
                } else {
                    el.style[attr] = iCurr + speed + 'px';
                }
            }

            if (flag) {
                clearInterval(el.timer);
                if(fn) fn();
            }

        }, duration);
    }

    return {
        move: move
    }

});

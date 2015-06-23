define(['utilities', 'eventUtil'], function(utilities, eventUtil) {

    var oFullSlide = utilities.getElementsByClassName(document, "div", "fullSlide")[0],
        timer = null,
        step = 0,
        current = 0;
    var obd = utilities.getElementsByClassName(oFullSlide, 'div', 'bd')[0];
    var obdUl = oFullSlide.getElementsByTagName("ul")[0];
    var ohd = utilities.getElementsByClassName(oFullSlide, 'div', 'hd')[0];
    var ohdLi = ohd.getElementsByTagName("li");
    var prev = utilities.getElementsByClassName(oFullSlide, 'span', 'prev')[0];
    var next = utilities.getElementsByClassName(oFullSlide, 'span', 'next')[0];

    function slideMove(cur) {
        if (timer) {
            clearInterval(timer);
            clearTimeout(timer);
        }
        step = -(cur * 1920);
        obdUl.style.left = step + "px";
        for (var j = ohdLi.length - 1; j >= 0; j--) {
            ohdLi[j].className = "";
        }
        ohdLi[cur].className = "on";
        timer = setTimeout(slide, 3000);
    }

    function init() {

        obd.style.marginLeft = -(obd.offsetWidth / 2) + "px";
        obd.style.left = "50%";

        for (var i = ohdLi.length - 1; i >= 0; i--) {
            ohdLi[i].index = i;
        }

        eventUtil.addHandler(oFullSlide, 'mouseover', function() {
            prev.style.display = "block";
            next.style.display = "block";
        });

        eventUtil.addHandler(oFullSlide, 'mouseout', function() {
            prev.style.display = "none";
            next.style.display = "none";
        });

        eventUtil.addHandler(prev, 'click', function() {
            current -= 1;
            if (current < 0) current = 2;
            slideMove(current);
        });

        eventUtil.addHandler(next, 'click', function() {
            current += 1;
            if (current > 2) current = 0;
            slideMove(current);
        });

        eventUtil.addHandler(ohd, 'click', function(event) {
            var ev = event || window.event;
            var target = ev.target || ev.srcElement;
            if (!isNaN(target.index)) {
                current = target.index;
                slideMove(current);
            }

        });        
    }

    eventUtil.addResizeEvent(function() {
        obd.style.left = "50%";
        if (oFullSlide.offsetWidth > 940) {
            obd.style.width = "120%";
            obd.style.marginLeft = -(obd.offsetWidth / 2) + "px";
        } else if (oFullSlide.offsetWidth == 940) {
            obd.style.width = "1900px";
            obd.style.marginLeft = "-960px";
        }
    });

    function slide() {
        clearTimeout(timer);
        timer = setInterval(function() {
            step -= 192;
            current = Math.abs(Math.floor(step / 1920));
            if (current == 3) {
                step = 0;
                current = 0
            };
            if (step % 1920 == 0) {
                clearInterval(timer);
                timer = setTimeout(slide, 3000);
            }
            obdUl.style.left = step + "px";

        }, 100);
    }

    return {
        init: init,
        slide: slide
    }

});

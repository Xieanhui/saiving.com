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

    obd.style.marginLeft = -(obd.offsetWidth / 2) + "px";

    for (var i = ohdLi.length - 1; i >= 0; i--) {
        ohdLi[i].index = i;
    }

    function init() {

        eventUtil.addHandler(oFullSlide, 'mouseover', function(event) {

            prev.style.display = "block";
            next.style.display = "block";

            eventUtil.addHandler(prev, 'click', function() {
                if (timer) {
                    clearInterval(timer);
                    clearTimeout(timer);
                }
                current -= 1;
                if (current < 0) current = 3;
                step = -(current * 1920);
                obdUl.style.left = step + "px";
                for (var j = ohdLi.length - 1; j >= 0; j--) {
                    ohdLi[j].className = "";
                }
                if (current == 3) {
                    ohdLi[2].className = "on";
                } else {
                    ohdLi[current].className = "on";
                }
                timer = setTimeout(slide, 3000);
            });

            eventUtil.addHandler(next, 'click', function() {
                if (timer) {
                    clearInterval(timer);
                    clearTimeout(timer);
                }
                current += 1;
                if (current > 3) current = 0;
                step = -(current * 1920);
                obdUl.style.left = step + "px";
                for (var j = ohdLi.length - 1; j >= 0; j--) {
                    ohdLi[j].className = "";
                }
                if (current == 3) {
                    ohdLi[2].className = "on";
                } else {
                    ohdLi[current].className = "on";
                }
                timer = setTimeout(slide, 3000);
            });

            eventUtil.addHandler(ohdLi, 'click', function() {
                if (timer) {
                    clearInterval(timer);
                    clearTimeout(timer);
                }
                current = this.index;
                step = -(current * 1920);
                obdUl.style.left = step + "px";
                for (var j = ohdLi.length - 1; j >= 0; j--) {
                    ohdLi[j].className = "";
                }
                this.className = "on";
                setTimeout(slide, 3000);
            });


        });


        eventUtil.addHandler(oFullSlide, 'mouseout', function(event) {
            prev.style.display = "none";
            next.style.display = "none";

            eventUtil.removeHandler(prev, 'click', function() {
                if (timer) {
                    clearInterval(timer);
                    clearTimeout(timer);
                }
                current -= 1;
                if (current < 0) current = 3;
                step = -(current * 1920);
                obdUl.style.left = step + "px";
                for (var j = ohdLi.length - 1; j >= 0; j--) {
                    ohdLi[j].className = "";
                }
                if (current == 3) {
                    ohdLi[2].className = "on";
                } else {
                    ohdLi[current].className = "on";
                }
                timer = setTimeout(slide, 3000);
            });

            eventUtil.removeHandler(next, 'click', function() {
                if (timer) {
                    clearInterval(timer);
                    clearTimeout(timer);
                }
                current += 1;
                if (current > 3) current = 0;
                step = -(current * 1920);
                obdUl.style.left = step + "px";
                for (var j = ohdLi.length - 1; j >= 0; j--) {
                    ohdLi[j].className = "";
                }
                if (current == 3) {
                    ohdLi[2].className = "on";
                } else {
                    ohdLi[current].className = "on";
                }
                timer = setTimeout(slide, 3000);
            });

            eventUtil.removeHandler(ohdLi, 'click', function() {
                if (timer) {
                    clearInterval(timer);
                    clearTimeout(timer);
                }
                current = this.index;
                step = -(current * 1920);
                obdUl.style.left = step + "px";
                for (var j = ohdLi.length - 1; j >= 0; j--) {
                    ohdLi[j].className = "";
                }
                this.className = "on";
                setTimeout(slide, 3000);
            });


        });

    }


    eventUtil.addResizeEvent(function() {
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
            if (current == 4) {
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

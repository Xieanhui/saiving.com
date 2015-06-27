define(['utilities', 'eventUtil'], function(utilities, eventUtil) {

    var olive = utilities.g("live"),
        step = 0,
        timer = null,
        current = 0;
    var oLiveCtrl = (utilities.g("liveCtrl") != null) ? utilities.g("liveCtrl").getElementsByTagName('li') : null;

if (olive != null && oLiveCtrl != null) {
        oLiveCtrl[0].style.backgroundColor = "#ffffff";
        for (var j = 0; j < oLiveCtrl.length; j++) {
            oLiveCtrl[j].index = j;
        }
        eventUtil.addHandler(oLiveCtrl, 'click', function(event) {
            for (var j = 0; j < oLiveCtrl.length; j++) {
                oLiveCtrl[j].style.backgroundColor = "#555555";
            }
            step = -(this.index * 760);
            olive.style.left = step + "px";
            this.style.backgroundColor = "#ffffff";
            if (timer) clearInterval(timer);
            timer = setTimeout(slide, 2500);
        });
    }

    function slide() {
        if (olive != null) {
            if (timer) {
                clearInterval(timer);
                clearTimeout(timer);
            }
            timer = setInterval(function() {
                step -= 76;
                if (olive.offsetLeft <= -1520) step = 0;
                if (step % 760 == 0) {
                    current = Math.abs(Math.floor(step / 760));
                    if (current == oLiveCtrl.length) current = 0;
                    for (var j = 0; j < oLiveCtrl.length; j++) {
                        oLiveCtrl[j].style.backgroundColor = "#555555";
                    }
                    oLiveCtrl[current].style.backgroundColor = "#ffffff";
                    if (timer) clearInterval(timer);
                    timer = setTimeout(slide, 2500);
                }        
                
                olive.style.left = step + "px";
            }, 20);
        }
    }

    return {
        slide: slide
    }
});

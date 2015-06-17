define(['utilities', 'eventUtil'], function(utilities, eventUtil) {

    var olive = utilities.g("live"),
        step = 0,
        timer = null,
        current = 0;
    var oLiveCtrl = (utilities.g("liveCtrl") != null) ? utilities.g("liveCtrl").getElementsByTagName('li') : null;

    function slide() {
        if (olive != null) {
            clearTimeout(timer);
            timer = setInterval(function() {
                if (olive.offsetLeft <= -4000) step = 0;
                if (step % 2000 == 0) {
                    current = Math.abs(Math.floor(step / 2000));
                    if (current == oLiveCtrl.length) current = 0;
                    for (var j = 0; j < oLiveCtrl.length; j++) {
                        oLiveCtrl[j].style.backgroundColor = "#555555";
                    }
                    oLiveCtrl[current].style.backgroundColor = "#ffffff";
                    if (timer) clearInterval(timer);
                    timer = setTimeout(slide, 3000);
                }

                step -= 100;
                olive.style.left = step + "px";

            }, 10);
        }
    }


    if (olive != null && oLiveCtrl != null) {

        oLiveCtrl[0].style.backgroundColor = "#ffffff";

        for (var j = 0; j < oLiveCtrl.length; j++) {
            oLiveCtrl[j].index = j;
        }

        eventUtil.addHandler(oLiveCtrl, 'click', function(event) {
            if (timer) {
                clearInterval(timer);
                clearTimeout(timer);
            }
            for (var j = 0; j < oLiveCtrl.length; j++) {
                oLiveCtrl[j].style.backgroundColor = "#555555";
            }
            step = -(this.index * 2000);
            olive.style.left = step + "px";
            this.style.backgroundColor = "#ffffff";
            timer = setTimeout(slide, 3000);
        });
    }

    return {
        slide: slide
    }
});

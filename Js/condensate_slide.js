define(['utilities'],function(utilities) {
    var olive = utilities.g("live");
    var step = 0;
    var timer = null;

    function slide() {
        if (olive != null) {
            if(timer){clearInterval(timer);}
            timer = setInterval(function() {
                step -= 100;
                if (olive.offsetLeft <= -4000) step = 0;
                if (step % 2000 == 0) {
                    clearInterval(timer);
                    timer = setTimeout(slide, 3000);
                }
                olive.style.left = step + "px";

            }, 10);
        }

    }

    return {
        slide: slide
    }
})

define(['eventUtil', 'utilities'], function(eventUtil, utilities) {
    var oseventhPage = utilities.getElementsByClassName('seventhPage')[0];
    var arrImg = oseventhPage.getElementsByTagName('img');

    for (var i = arrImg.length - 1; i >= 0; i--) {
        arrImg[i].flag = 'no';
        arrImg[i].onload = function() {
            arrImg[i].flag = 'yes';
        }
    }

    clearInterval(timer);
    timer = setInterval(function() {
        if (arrImg[0].flag == 'yes' && arrImg[1].flag == 'yes' && arrImg[2].flag == 'yes' && arrImg[3].flag == 'yes') {
            clearInterval(timer);

            
        }
    }, 0);
});

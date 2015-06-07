define(function () {

    function loadXMLDoc(url, element) {//发送异步HTTP请求

        var xmlhttp = null;

        if (window.XMLHttpRequest) {
            xmlhttp = new XMLHttpRequest();
        } else if (window.ActiveXObject) {
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
        }

        if (xmlhttp != null) {
            xmlhttp.onreadystatechange = function () {
                if (xmlhttp.readyState === 4) {
                    if (xmlhttp.status === 200) {
                        element.innerHTML = xmlhttp.responseText;
                    } else {
                        console.log("The data request failed:" + xmlhttp.statusText);
                    }
                }
            };
            xmlhttp.open("GET", url, true);
            xmlhttp.send(null);

        } else {
            console.log("Your browser does not support this function");
        }
    }

    return {
        loadXMLDoc:loadXMLDoc
    }
    
});

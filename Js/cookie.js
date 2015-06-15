define(function() {

    function setCookie(name, value) {

        var argv = setCookie.arguments,
            argc = setCookie.arguments.length,
            expires = (argc > 2) ? argv[2] : null;

        if (expires != null) {

            var d = new Date();

            d.setTime(d.getTime() + (expires * 1000 * 3600 * 24));
        }

        document.cookie = name + "=" + escape(value) + ((expires == null) ? "" : "; expires=" + d);

    }

    function getCookie(name) {

        var search = name + "=";
        
        if (document.cookie.length > 0) {
            var offset = document.cookie.indexOf(search);
            if (offset != -1) {
                offset += search.length;
                var end = document.cookie.indexOf(";", offset);
                if (end == -1) end = document.cookie.length;
                return unescape(document.cookie.substring(offset, end));
            } else {
                return "";
            }
        }

    }

    function delCookie(name) {

        var cval = getCookie(name);

        if (cval != null) {
            var exp = new Date();
            exp.setTime(exp.getTime() - 1);
            document.cookie = name + "=" + cval + ";expires=" + exp.toGMTString();
        }

    }

    return {
        setCookie: setCookie,
        getCookie: getCookie,
        delCookie: delCookie
    };

});

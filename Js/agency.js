define(['utilities', 'cookie', 'agencyAnchor', 'agencyDetail'], function(utilities, cookie, agencyAnchor, agencyDetail) {

    var agencyCookie = cookie.getCookie("agency");
    
    if (!agencyCookie || agencyCookie == "") {
        try {
            utilities.proxyResponse("http://int.dpool.sina.com.cn/iplookup/iplookup.php?format=js", function() {
                var agencyText = "";
                if (remote_ip_info["province"] == null || remote_ip_info["province"] == "\u4fdd\u7559\u5730\u5740" || remote_ip_info["province"] == "") {
                    agencyText = "\u5317\u4eac";
                } else {
                    agencyText = remote_ip_info["province"];
                }                
                agencyAnchor.setAnchor(agencyText);
                agencyDetail.showAgency(agencyText);
                cookie.setCookie("agency", agencyText, 10);
            });
        } catch (err) {
            utilities.proxyResponse("http://pv.sohu.com/cityjson", function() {
                var agencyText = returnCitySN["cname"];
                agencyAnchor.setAnchor(agencyText);
                agencyDetail.showAgency(agencyText);
                cookie.setCookie("agency", agencyText, 10);
            });
        }
    } else {
        agencyAnchor.setAnchor(agencyCookie);
        agencyDetail.showAgency(agencyCookie);
    }

});

shpUtility.factory('shpPropertyBag', ['$log', '$http', 'shpCommon', function($log, $http, shpCommon) {

    //Get URLs
    var hostUrl = shpCommon.SPHostUrl;
    var appweburl = shpCommon.SPAppWebUrl;

    return {
        SaveObjToCurrentWeb: saveObjToCurrentWebPropertyBag,
        SaveObjToRootWeb: saveObjToRootWebPropertyBag,
        GetValue: getPropertyBagValue
    }

    function saveObjToCurrentWebPropertyBag(jsonObject, callback) {
        var oWebsite = clientCtx.get_web();
        savePropertyBag(oWebsite, jsonObject, callback);
    }

    function saveObjToRootWebPropertyBag(jsonObject, callback) {
        var oWebsite = clientCtx.get_site().get_rootWeb();
        savePropertyBag(oWebsite, jsonObject, callback);
    }

    function getPropertyBagValue(key, callback, optionalWebUrl) {
        var rootPath = optionalWebUrl || appweburl;
        var url = rootPath + '/_api/web/AllProperties?$select=' + key;

        $http.get(url).
        success(function(result) {
            var value = "";
            if (result[key])
                value = result[key];
            callback(value);
        }).
        error(function(data, status) {
            $log.log(status);
            $log.log(data);
        });
    }

}]);
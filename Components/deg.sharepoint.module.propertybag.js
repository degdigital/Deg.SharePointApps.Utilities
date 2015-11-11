DegSharepointUtility.PropertyBag = {


    SaveObjToCurrentWeb: function(jsonObject, callback) {
        var oWebsite = clientCtx.get_web();
        savePropertyBag(oWebsite, jsonObject, callback)

    },

    SaveObjToRootWeb: function(jsonObject, callback) {
        var oWebsite = clientCtx.get_site().get_rootWeb();
        savePropertyBag(oWebsite, jsonObject, callback);
    },

    GetValue: function(key, callback, optionalWebUrl) {
        var rootPath = optionalWebUrl || getAppWebUrl();
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

}
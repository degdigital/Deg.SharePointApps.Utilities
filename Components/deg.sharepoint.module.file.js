DegSharepointUtility.File = {

    CreateAtHost: function(appWebFileUrl, serverRelativeUrl, fileName, isPublishRequired, callback) {
        $http.get(appWebFileUrl).
        success(function(fileContents) {
            if (fileContents !== undefined && fileContents.length > 0) {
                if (!isPublishRequired) {
                    uploadFileToHostWeb(serverRelativeUrl, fileName, fileContents, isPublishRequired, callback);
                } else {
                    var fileRelativeUrl = serverRelativeUrl + '/' + fileName;
                    loadFileAtHostWeb(fileRelativeUrl, function(result) {
                        var fileExists = result.Success;
                        if (!fileExists) {
                            uploadFileToHostWeb(serverRelativeUrl, fileName, fileContents, isPublishRequired, callback);
                        } else {
                            checkOutFileAtHostWeb(fileRelativeUrl, function(result) {
                                if (result.Success) {
                                    uploadFileToHostWeb(serverRelativeUrl, fileName, fileContents, isPublishRequired, callback);
                                } else {
                                    if (callback) callback(result);
                                }
                            });
                        }
                    });
                }
            } else {
                if (callback) callback({
                    Success: false,
                    Message: 'Failed to read file from app web.'
                });
            }
        }).error(function(data, status) {
            if (callback) callback({
                Success: false,
                Message: 'Request for file in app web failed: ' + status
            });
        });
    },

    LoadAtHost: function(fileRelativeUrl, callback) {
        var hostWebUrl = getHostWebUrl();
        var serverRelativeUrl = getRelativeUrlFromAbsolute(hostWebUrl);
        var hostWebContext = new SP.ClientContext(getRelativeUrlFromAbsolute(hostWebUrl));

        var fileUrl = serverRelativeUrl + fileRelativeUrl;
        var file = hostWebContext.get_web().getFileByServerRelativeUrl(fileUrl);
        hostWebContext.load(file);

        hostWebContext.executeQueryAsync(onLoadFileSuccess, onLoadFileFail);

        function onLoadFileSuccess() {
            if (callback) callback({
                Success: true,
                Message: 'File loaded from host web successfully: ' + fileUrl
            });
        }

        function onLoadFileFail(sender, args) {
            if (callback) callback({
                Success: false,
                Message: 'Failed to read file from host web. Error: ' + sender.statusCode
            });
        }
    },

    CheckOutAtHost: function(fileRelativeUrl, callback) {
        var hostWebUrl = getHostWebUrl();
        var serverRelativeUrl = getRelativeUrlFromAbsolute(hostWebUrl);
        var hostWebContext = new SP.ClientContext(getRelativeUrlFromAbsolute(hostWebUrl));

        var fileUrl = serverRelativeUrl + fileRelativeUrl;
        var file = hostWebContext.get_web().getFileByServerRelativeUrl(fileUrl);
        hostWebContext.load(file);

        hostWebContext.executeQueryAsync(onLoadFileSuccess, onLoadFileFail);

        function onLoadFileSuccess() {
            var isCheckedOut = file.get_checkOutType() == 0;
            if (!isCheckedOut) {
                file.checkOut();
                hostWebContext.executeQueryAsync(onCheckoutFileSuccess, onCheckoutFileFail);
            } else {
                if (callback) callback({
                    Success: true,
                    Message: 'File checked out in host web successfully: ' + fileRelativeUrl
                });
            }
        }

        function onLoadFileFail(sender, args) {
            if (callback) callback({
                Success: false,
                Message: 'Failed to read file from host web. Error: ' + sender.statusCode
            });
        }

        function onCheckoutFileSuccess() {
            if (callback) callback({
                Success: true,
                Message: 'File checked out in host web successfully: ' + fileRelativeUrl
            });
        }

        function onCheckoutFileFail(sender, args) {
            if (callback) callback({
                Success: false,
                Message: 'Failed to checkout file at host web. Error: ' + sender.statusCode
            });
        }
    },

    PublishFileToHost: function(fileRelativeUrl, callback) {
        var hostWebUrl = getHostWebUrl();
        var serverRelativeUrl = getRelativeUrlFromAbsolute(hostWebUrl);
        var hostWebContext = new SP.ClientContext(getRelativeUrlFromAbsolute(hostWebUrl));

        var fileUrl = serverRelativeUrl + fileRelativeUrl;
        var file = hostWebContext.get_web().getFileByServerRelativeUrl(fileUrl);
        hostWebContext.load(file);

        hostWebContext.executeQueryAsync(onLoadFileSuccess, onLoadFileFail);

        function onLoadFileSuccess() {
            var isCheckedOut = file.get_checkOutType() == 0;
            if (!isCheckedOut) {
                file.checkIn();
                file.publish();
                hostWebContext.executeQueryAsync(onPublishFileSuccess, onPublishFileFail);
            } else {
                if (callback) callback({
                    Success: true,
                    Message: 'File published in host web successfully: ' + fileRelativeUrl
                });
            }
        }

        function onLoadFileFail(sender, args) {
            if (callback) callback({
                Success: false,
                Message: 'Failed to read file from host web. Error: ' + sender.statusCode
            });
        }

        function onPublishFileSuccess() {
            if (callback) callback({
                Success: true,
                Message: 'File published in host web successfully: ' + fileRelativeUrl
            });
        }

        function onPublishFileFail(sender, args) {
            if (callback) callback({
                Success: false,
                Message: 'Failed to publish file into host web. Error: ' + sender.statusCode
            });
        }
    };

};
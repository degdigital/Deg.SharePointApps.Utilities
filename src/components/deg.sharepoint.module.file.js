shpUtility.factory('shpFile', ['$http', function($http) {


    return {
        CreateAtHost: readFromAppWebAndProvisionToHost,
        LoadAtHost: loadFileAtHostWeb,
        CheckOutAtHost: checkOutFileAtHostWeb,
        PublishFileToHost: publishFileToHostWeb,
        UploadFileToHostWeb: uploadFileToHostWeb
    }

    function readFromAppWebAndProvisionToHost(appWebFileUrl, serverRelativeUrl, fileName, isPublishRequired, callback) {
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
    }

    function uploadFileToHostWeb(serverRelativeUrl, fileName, contents, isPublishRequired, callback) {
        var hostWebUrl = getHostWebUrl();
        var hostWebContext = new SP.ClientContext(getRelativeUrlFromAbsolute(hostWebUrl));

        var createInfo = new SP.FileCreationInformation();
        createInfo.set_content(new SP.Base64EncodedByteArray());
        for (var i = 0; i < contents.length; i++) {
            createInfo.get_content().append(contents.charCodeAt(i));
        }
        createInfo.set_overwrite(true);
        createInfo.set_url(fileName);
        var files = hostWebContext.get_web().getFolderByServerRelativeUrl(serverRelativeUrl).get_files();
        hostWebContext.load(files);
        files.add(createInfo);

        hostWebContext.executeQueryAsync(onProvisionFileSuccess, onProvisionFileFail);

        function onProvisionFileSuccess() {
            var fileRelativeUrl = serverRelativeUrl + '/' + fileName;
            if (isPublishRequired) {
                publishFileToHostWeb(fileRelativeUrl, function(result) {
                    if (callback) callback(result);
                });
            } else {
                if (callback) callback({
                    Success: true,
                    Message: 'File published in host web successfully: ' + fileRelativeUrl
                });
            }
        }

        function onProvisionFileFail(sender, args) {
            if (callback) callback({
                Success: false,
                Message: 'Failed to provision file into host web. Error: ' + sender.statusCode
            });
        }
    }

    function loadFileAtHostWeb(fileRelativeUrl, callback) {
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
    }

    function checkOutFileAtHostWeb(fileRelativeUrl, callback) {
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
    }

    function publishFileToHostWeb(fileRelativeUrl, callback) {
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
    }

}]);
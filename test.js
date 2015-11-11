var DegSharepointUtility = {};

DegSharepointUtility.User = {

    GetCurrentUserName: function getUserName(callback) {
        try {
            var clientCtx = SP.ClientContext.get_current();
            var user = clientCtx.get_web().get_currentUser();
            clientCtx.load(user);
            clientCtx.executeQueryAsync(onGetUserNameSuccess, onGetUserNameFail);
        } catch (e) {
            callback(null);
        }

        function onGetUserNameSuccess() {
            var userName = user.get_title();
            callback(userName);
        }

        function onGetUserNameFail(sender, args) {
            $log.log('Failed to get user name. Error:' + args.get_message());
        }
    },

    GetCurrent: function getCurrentUser(callback) {
        var context = SP.ClientContext.get_current();
        var user = context.get_web().get_currentUser();

        context.load(user);
        context.executeQueryAsync(onGetUserNameSuccess, onGetUserNameFail);

        function onGetUserNameSuccess() {
            if (callback) callback(user);
        }

        function onGetUserNameFail(sender, args) {
            $log.log('Failed to get user. Error: ' + args.get_message());
        }
    },

    GetId: function getUserId(loginName, callback) {
        var context = SP.ClientContext.get_current();
        var user = context.get_web().ensureUser(loginName);
        context.load(user);
        context.executeQueryAsync(onEnsureUserSuccess, onEnsureUserFail);

        function onEnsureUserSuccess() {
            if (callback) callback(user.get_id());
        }

        function onEnsureUserFail(sender, args) {
            $log.log('Failed to ensure user. Error: ' + args.get_message());
        }
    },
    
    GetCurrentUserProfileProperties: function getCurrentUserProperties() {

        var deferred = $q.defer();

        var clientContext = SP.ClientContext.get_current();
        var peopleManager = new SP.UserProfiles.PeopleManager(clientContext);
        var userProfileProperties = peopleManager.getMyProperties();

        clientContext.load(userProfileProperties);
        clientContext.executeQueryAsync(
            function() {
                deferred.resolve(userProfileProperties.get_userProfileProperties());
            },
            function(data) {
                console.log(data);
                deferred.reject(data);
                console.log("Error");
            }
        );

        return deferred.promise;

    }

};

DegSharepointUtility.File = {

    CreateAtHost: function readFromAppWebAndProvisionToHost(appWebFileUrl, serverRelativeUrl, fileName, isPublishRequired, callback) {
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

    LoadAtHost: function loadFileAtHostWeb(fileRelativeUrl, callback) {
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

    CheckOutAtHost: function checkOutFileAtHostWeb(fileRelativeUrl, callback) {
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
    
    PublishFileToHost: function publishFileToHostWeb(fileRelativeUrl, callback) {
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
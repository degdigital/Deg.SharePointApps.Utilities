/** 
* @version 0.2
* @license MIT
* by Nick Aranzamendi
*/

'use strict';
angular.module("Deg.SharePoint", []).service('spService', ['$http', '$log', '$q', function ($http, $log, $q) {
    return {
        User: {
            GetCurrentUserName: getUserName
        },
        CtxInfo: {
            SPAppWebUrl: getAppWebUrl(),
            //SPContextInfo: _spPageContextInfo,
            SPHostUrl: getHostWebUrl(),
        },
        PropBag: {
            SaveObjToCurrentWeb: saveObjToCurrentWebPropertyBag,
            SaveObjToRootWeb: saveObjToRootWebPropertyBag,
            GetValue: getPropertyBagValue,
        },
        Utilities: {
            GetQsParam: getUrlParam
        },
        Lists: {
            CreateAtHost: createListInHost,
            AddFieldToListAtHost: addFieldToRootList,
            Exist: existRootList
        },
        Columns: {
            CreateAtHost: createRootField
        },
        CTypes: {
            CreateAtHost: createContentTypeInHost
        },
        Files: {
            CreateAtHost: readFromAppWebAndProvisionToHost,
            LoadAtHost: loadFileAtHostWeb,
            CheckOutAtHost: checkOutFileAtHostWeb,
            PublishFileToHost: publishFileToHostWeb        
        }
    };
    function createCtype(cTypeInfo, callback, sPCtx) {
        var ctx = sPCtx || clientCtx;
        var web = ctx.get_web();
        var ctypes = web.get_contentTypes();
        var creationInfo = new SP.ContentTypeCreationInformation();
        creationInfo.set_name(cTypeInfo.Name);
        creationInfo.set_description(cTypeInfo.Description);
        creationInfo.set_group(cTypeInfo.Group);
        if (cTypeInfo.ParentContentType && cTypeInfo.ParentContentType != '') {
            var parent = ctypes.getById(cTypeInfo.ParentContentType);
            creationInfo.set_parentContentType(parent);
        }
        ctypes.add(creationInfo);
        ctx.load(ctypes);
        ctx.executeQueryAsync(onProvisionContentTypeSuccess, onProvisionContentTypeFail);
        function onProvisionContentTypeSuccess() {
            if (callback)
                callback(creationInfo);
        }
        function onProvisionContentTypeFail(sender, args) {
            $log.log("Error: " + args.get_message());
        }
    }
    // ctypeInfo { Name :'', Description : '', Group: '', ParentContentType: 'optional'}
    function createContentTypeInHost(cTypeInfo, callback) {
        var hostWebContext = getHostWebContext();
        createCtype(cTypeInfo, callback, hostWebContext);
    }

    function createListInHost(listName, callback) {
        //Get URLs
        var hostUrl = getHostWebUrl();
        var appweburl = getAppWebUrl();
        //Get Contexts
        var appContext = new SP.ClientContext(appweburl);
        var hostContext = new SP.AppContextSite(appContext, hostUrl);
        //Get root web
        var oWebsite = hostContext.get_web();

        //Create list
        var listCreationInfo = new SP.ListCreationInformation();
        listCreationInfo.set_title(listName);
        listCreationInfo.set_templateType(SP.ListTemplateType.genericList);

        var oList = oWebsite.get_lists().add(listCreationInfo);

        appContext.load(oList);
        appContext.executeQueryAsync(
            function (sender, args) {
                checkCallback();
            },
            function (sender, args) {
                $log.log('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
            }
        );

        function checkCallback() {
            if (callback) {
                callback();
            }
        }
    }

    function existRootList(listName, callback) {
        
        //Get URLs
        var hostUrl = getHostWebUrl();
        var appweburl = getAppWebUrl();
        //Get Contexts
        var appContext = new SP.ClientContext(appweburl);
        var hostContext = new SP.AppContextSite(appContext, hostUrl);
        //Get root web
        var oList = hostContext.get_web().get_lists().getByTitle(listName);
        
        appContext.load(oList);
        appContext.executeQueryAsync(
            function (sender, args) {
                checkCallback();
            },
            function (sender, args) {
                $log.log('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
                
            }
        );

        function checkCallback() {
            if (callback) {
                callback();
            }
            
        }
        

    }

    function addFieldToRootList(listName, fieldDisplayName, fieldName, fieldType, fieldExtra, callback) {
        //var deferred = $.Deferred();
        var deferred = $q.defer();
        //Get URLs
        var hostUrl = getHostWebUrl();
        var appweburl = getAppWebUrl();
        //Get Contexts
        var appContext = new SP.ClientContext(appweburl);
        var hostContext = new SP.AppContextSite(appContext, hostUrl);
        //Get root web
        var oList = hostContext.get_web().get_lists().getByTitle(listName);

        var extraTypeDefinition = "";
        if (fieldType == "URL") {
            extraTypeDefinition = "Format=\'" + fieldExtra + "\'";
        }

        var fieldDefinition = "<Field DisplayName=\'" + fieldName + "\' Type=\'" + fieldType + "\' " + extraTypeDefinition + "/>";
        var oField = oList.get_fields().addFieldAsXml(
            fieldDefinition,
            true,
            SP.AddFieldOptions.defaultValue
        );

        var fieldNumber;
        switch (fieldType) {
            case "Number":
                fieldNumber = appContext.castTo(oField, SP.FieldNumber);
                break;
            case "URL":
                fieldNumber = appContext.castTo(oField, SP.FieldUrl);
                break;
            case "User":
                fieldNumber = appContext.castTo(oField, SP.FieldUser);
                break;
            case "TaxonomyFieldType":
                fieldNumber = appContext.castTo(oField, SP.Field);
                break;
            default:
                fieldNumber = appContext.castTo(oField, SP.FieldText);
                break;
        }

        fieldNumber.set_title(fieldDisplayName);
        fieldNumber.update();

        appContext.load(oField);
        appContext.executeQueryAsync(
            function (sender, args) {
                checkCallback();
            },
            function (sender, args) {
                $log.log('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
                deferred.reject();
            }
        );

        function checkCallback() {
            if (callback) {
                callback();
            }
            deferred.resolve();
        }
        return deferred.promise;
    }


    function createRootField(fieldsXml, callback) {
        var hostWebContext = getHostWebContext();
        createFields(fieldsXml, callback, hostWebContext);
    }
    function getHostWebContext() {
        var hostUrl = getHostWebUrl();
        var hostWebContext = new SP.ClientContext(getRelativeUrlFromAbsolute(hostUrl));
        return hostWebContext;
    }

    function createFields(fieldsXml, callback, sPCtx) {
        var clientCtx = SP.ClientContext.get_current();
        // field XML format :: <Field DisplayName='Field Name' Name='NoSpaceName' ID='{2d9c2efe-58f2-4003-85ce-0251eb174096}' Group='Group Name' Type='Text' />
        var ctx = sPCtx || clientCtx;
        var fieldsExecuted = 0
        var fields = ctx.get_web().get_fields();
        var response = {
            ErrorMessages: [],
            Fields: []
        };
        for (var x = 0; x < fieldsXml.length; x++) {
            var newField = fields.addFieldAsXml(fieldsXml[x], false, SP.AddFieldOptions.AddToNoContentType);
            ctx.load(newField);
            response.Fields.push(newField);
            // executing one by one because process stops if one field errors out when queuing all
            ctx.executeQueryAsync(function () {
                $log.log("Field provisioned in host web successfully");
                checkCallback();
            }, function (sender, args) {
                $log.log('Failed to provision field into host web.');
                response.ErrorMessages.push(args.get_message())
                checkCallback();
            });
        }
        function checkCallback() {
            fieldsExecuted++;
            if (callback && fieldsXml.length === fieldsExecuted) {
                callback(response);
            }
        }
    }
    function getUserName(callback) {
        try {
            var clientCtx = SP.ClientContext.get_current();
            var user = clientCtx.get_web().get_currentUser();
            clientCtx.load(user);
            clientCtx.executeQueryAsync(onGetUserNameSuccess, onGetUserNameFail);
        }
        catch (e) {
            callback(null);
        }
        function onGetUserNameSuccess() {
            var userName = user.get_title();
            callback(userName);
        }
        function onGetUserNameFail(sender, args) {
            $log.log('Failed to get user name. Error:' + args.get_message());
        }

    }

    function getAppWebUrl() {
        return decodeURIComponent(getUrlParam("SPAppWebUrl"));
    }
    function getHostWebUrl() {
        return decodeURIComponent(getUrlParam("SPHostUrl"));
    }

    function getPropertyBagValue(key, callback, optionalWebUrl) {
        var rootPath = optionalWebUrl || getAppWebUrl();
        var url = rootPath + '/_api/web/AllProperties?$select=' + key;

        $http.get(url).
            success(function (result) {
                var value = "";
                if (result[key])
                    value = result[key];
                callback(value);
            }).
            error(function (data, status) {
                $log.log(status);
                $log.log(data);
            });
    }

    function saveObjToCurrentWebPropertyBag(jsonObject, callback) {
        var oWebsite = clientCtx.get_web();
        savePropertyBag(oWebsite, jsonObject, callback)

    }
    function saveObjToRootWebPropertyBag(jsonObject, callback) {
        var oWebsite = clientCtx.get_site().get_rootWeb();
        savePropertyBag(oWebsite, jsonObject, callback);
    }
    function savePropertyBag(oWebsite, obj, callback) {
        clientCtx.load(oWebsite);
        var props = oWebsite.get_allProperties();
        for (var property in obj) {
            if (obj.hasOwnProperty(property)) {
                props.set_item(property, obj[property]);
            }
        }
        oWebsite.update();
        clientCtx.load(oWebsite);
        clientCtx.executeQueryAsync(onQuerySucceeded, onQueryFailed);

        function onQuerySucceeded() {

            if (callback)
                callback();
            else
                $log.log("Properties saved");
        }
        function onQueryFailed(sender, args) {
            $log.log('Request failed. ' + args.get_message() +
                '\n' + args.get_stackTrace());
        }
    }

    /** File **/
    function readFromAppWebAndProvisionToHost(appWebFileUrl, serverRelativeUrl, fileName, isPublishRequired, callback) {
        $http.get(appWebFileUrl).
            success(function (fileContents) {
                if (fileContents !== undefined && fileContents.length > 0) {
                    if (!isPublishRequired) {
                        uploadFileToHostWeb(serverRelativeUrl, fileName, fileContents, isPublishRequired, callback);
                    }
                    else {
                        var fileRelativeUrl = serverRelativeUrl + '/' + fileName;
                        loadFileAtHostWeb(fileRelativeUrl, function (result) {
                            var fileExists = result.Success;
                            if (!fileExists) {
                                uploadFileToHostWeb(serverRelativeUrl, fileName, fileContents, isPublishRequired, callback);
                            }
                            else {
                                checkOutFileAtHostWeb(fileRelativeUrl, function (result) {
                                    if (result.Success) {
                                        uploadFileToHostWeb(serverRelativeUrl, fileName, fileContents, isPublishRequired, callback);
                                    }
                                    else {
                                        if (callback) callback(result);
                                    }
                                });
                            }
                        });
                    }
                }
                else {
                    if (callback) callback({ Success: false, Message: 'Failed to read file from app web.' });
                }
            }).error(function (data, status) {
                if (callback) callback({ Success: false, Message: 'Request for file in app web failed: ' + status });
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
                publishFileToHostWeb(fileRelativeUrl, function (result) {
                    if (callback) callback(result);
                });
            }
            else {
                if (callback) callback({ Success: true, Message: 'File published in host web successfully: ' + fileRelativeUrl });
            }
        }
        function onProvisionFileFail(sender, args) {
            if (callback) callback({ Success: false, Message: 'Failed to provision file into host web. Error: ' + sender.statusCode });
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
            }
            else {
                if (callback) callback({ Success: true, Message: 'File checked out in host web successfully: ' + fileRelativeUrl });
            }
        }
        function onLoadFileFail(sender, args) {
            if (callback) callback({ Success: false, Message: 'Failed to read file from host web. Error: ' + sender.statusCode });
        }
        function onCheckoutFileSuccess() {
            if (callback) callback({ Success: true, Message: 'File checked out in host web successfully: ' + fileRelativeUrl });
        }
        function onCheckoutFileFail(sender, args) {
            if (callback) callback({ Success: false, Message: 'Failed to checkout file at host web. Error: ' + sender.statusCode });
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
            }
            else {
                if (callback) callback({ Success: true, Message: 'File published in host web successfully: ' + fileRelativeUrl });
            }
        }
        function onLoadFileFail(sender, args) {
            if (callback) callback({ Success: false, Message: 'Failed to read file from host web. Error: ' + sender.statusCode });
        }
        function onPublishFileSuccess() {
            if (callback) callback({ Success: true, Message: 'File published in host web successfully: ' + fileRelativeUrl });
        }
        function onPublishFileFail(sender, args) {
            if (callback) callback({ Success: false, Message: 'Failed to publish file into host web. Error: ' + sender.statusCode });
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
            if (callback) callback({ Success: true, Message: 'File loaded from host web successfully: ' + fileUrl });
        }
        function onLoadFileFail(sender, args) {
            if (callback) callback({ Success: false, Message: 'Failed to read file from host web. Error: ' + sender.statusCode });
        }
    }

    // Helpers
    function getUrlParam(key) {
        var vars = [], hash;
        var hashes = window.location.href.slice(window.location.href.indexOf('?') + 1).split('&');
        for (var i = 0; i < hashes.length; i++) {
            hash = hashes[i].split('=');
            vars.push(hash[0]);
            vars[hash[0]] = hash[1];
        }
        return vars[key];
    }
    function getRelativeUrlFromAbsolute(absoluteUrl) {
        absoluteUrl = absoluteUrl.replace('http://', '');
        absoluteUrl = absoluteUrl.replace('https://', '');
        var parts = absoluteUrl.split('/');
        var relativeUrl = '/';
        for (var i = 1; i < parts.length; i++) {
            relativeUrl += parts[i] + '/';
        }
        return relativeUrl;
    }

}])
.directive('ngAppFrame', ['$timeout', '$window', function ($timeout, $window) {

    return {
        restrict: 'E',
        link: function (scope, element, attrs) {
            element.css("display", "block");
            scope.$watch
            (
                function () {
                    return element[0].offsetHeight;
                },
                function (newHeight, oldHeight) {

                    if (newHeight != oldHeight) {
                        $timeout(function () {
                            var height = attrs.minheight ? newHeight + parseInt(attrs.minheight) : newHeight;
                            var id = getQsParam("SenderId");
                            var message = "<message senderId=" + id + ">resize(100%," + height + ")</message>";
                            $window.parent.postMessage(message, "*");
                        }, 0);// timeout needed to wait for DOM to update
                    }
                }
            );
        }
    };
    function getQsParam(name) {
        var match = RegExp('[?&]' + name + '=([^&]*)').exec($window.location.search);
        return match && decodeURIComponent(match[1].replace(/\+/g, ' '));
    }
}])

.directive('ngChromeControl', ['$window', function ($window) {

    return {
        restrict: 'E',
        link: function (scope, element, attrs) {
            element.attr("id", "chrome_ctrl_placeholder");
            //host url and title set automatically by SP
            var options = {
                appIconUrl: attrs.appiconurl,
                appTitle: attrs.apptitle,
            };
            if (attrs.apphelppageurl)
                options.appHelpPageUrl = attrs.apphelppageurl;

            var nav = new SP.UI.Controls.Navigation(element[0].id, options);
            nav.setVisible(true);

            scope.$watch
            (
                function () {
                    return attrs.apptitle;
                },
                function (newTitle, oldTitle) {
                    if (newTitle != oldTitle)
                        element.find('.ms-core-pageTitle').text(newTitle);
                }
            );
        }
    };

}])
.directive('ngPeoplePicker', [function () {
    return {
        restrict: 'A',
        require: "ngModel",
        link: link
    }

    function link(scope, element, attrs, ngModel) {

        var elementId = attrs.id;

        var returnUsername = (attrs.ngPeoplePicker == "username");

        var schema = {
            SearchPrincipalSource: 15,
            ResolvePrincipalSource: 15,
            MaximumEntitySuggestions: 50,
            Width: "100%",
            OnUserResolvedClientScript: onUserResolve
        };

        schema.Required = true;

        schema.PrincipalAccountType = "User,DL";

        if (attrs.allowmultiple) {
            schema.AllowMultipleValues = true;
        }
        else {
            schema.AllowMultipleValues = false;
        }

        SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
            SPClientPeoplePicker_InitStandaloneControlWrapper(elementId, null, schema);
        });

        function onUserResolve() {
            var peoplePickerDictKey = elementId + "_TopSpan";
            var peoplePicker = SPClientPeoplePicker.SPClientPeoplePickerDict[peoplePickerDictKey];
            var people = peoplePicker.GetAllUserInfo();
            var returnValues = [];
            if (returnUsername) {
                angular.forEach(people, function (person) {
                    returnValues.push(person.AutoFillKey);
                });
            }
            else {
                angular.forEach(people, function (person) {
                    returnValues.push(person.EntityData.Email);
                });
            }
            if (schema.AllowMultipleValues) {
                ngModel.$setViewValue(returnValues);
            }
            else if (returnValues.length > 0) {
                ngModel.$setViewValue(returnValues[0]);
            }
        }
    }
}]);
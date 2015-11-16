shpUtility.service('spService', ['$http', '$log', '$q',
    'shpCommon',
    'shpUser',
    'shpPropertyBag',
    'shpItem',
    'shpList',
    'shpContentType',
    'shpFile',
    'shpColumn',
    'shpGroup',
    'shpTaxonomy',

    function($http, $log, $q, shpCommon, shpUser, shpPropertyBag, shpItem, shpList, shpContentType, shpFile, shpColumn, shpGroup, shpTaxonomy) {

        return {
            User: {
                GetCurrentUserName: shpUser.GetCurrentUserName,
                GetCurrent: shpUser.GetCurrent,
                GetId: shpUser.GetId,
                GetCurrentUserProfileProperties: shpUser.GetId
            },
            CtxInfo: {
                SPAppWebUrl: shpCommon.SPAppWebUrl,
                SPHostUrl: shpCommon.SPHostUrl
            },
            PropBag: {
                SaveObjToCurrentWeb: shpPropertyBag.SaveObjToCurrentWeb,
                SaveObjToRootWeb: shpPropertyBag.SaveObjToRootWeb,
                GetValue: shpPropertyBag.GetValue
            },
            Utilities: {
                GetFormDigest: shpCommon.GetFormDigest,
                HostWebContext: shpCommon.HostWebContext,
                GetQsParam: shpCommon.GetQsParam
            },
            Lists: {
                CreateAtHost: shpList.CreateAtHost,
                AddFieldToListAtHost: shpList.AddFieldToListAtHost,
                Exist: shpList.Exist
            },
            Items: {
                Create: shpItem.Create,
                GetAll: shpItem.GetAll,
                Update: shpItem.Update
            },
            Columns: {
                CreateAtHost: shpColumn.CreateAtHost
            },
            CTypes: {
                CreateAtHost: shpContentType.CreateAtHost
            },
            Files: {
                CreateAtHost: shpFile.CreateAtHost,
                LoadAtHost: shpFile.LoadAtHost,
                CheckOutAtHost: shpFile.CheckOutAtHost,
                PublishFileToHost: shpFile.PublishFileToHost,
                UploadFileToHostWeb: shpFile.UploadFileToHostWeb
            },
            Groups: {
                LoadAtHost: shpGroup.LoadAtHost,
                CreateAtHost: shpGroup.CreateAtHost,
                IsCurrentUserMember: shpGroup.IsCurrentUserMember
            },
            Taxonomy: {
                GetTermSetValues: shpTaxonomy.GetTermSetValues
            }
        };
    }
]);


shpUtility.directive('ngAppFrame', ['$timeout', '$window', function($timeout, $window) {

    return {
        restrict: 'E',
        link: function(scope, element, attrs) {
            element.css("display", "block");
            scope.$watch(
                function() {
                    return element[0].offsetHeight;
                },
                function(newHeight, oldHeight) {

                    //if (newHeight != oldHeight) {
                    $timeout(function() {
                        if (typeof attrs.minheight == 'undefined') {
                            attrs.minheight = 50;
                        }
                        var height = attrs.minheight ? newHeight + parseInt(attrs.minheight) : newHeight;
                        var id = getQsParam("SenderId");
                        var message = "<message senderId=" + id + ">resize(100%," + height + ")</message>";
                        $window.parent.postMessage(message, "*");
                    }, 0); // timeout needed to wait for DOM to update
                    //}
                }
            );
        }
    };

    function getQsParam(name) {
        var match = RegExp('[?&]' + name + '=([^&]*)').exec($window.location.search);
        return match && decodeURIComponent(match[1].replace(/\+/g, ' '));
    }
}]);


shpUtility.directive('ngChromeControl', ['$window', function($window) {

    return {
        restrict: 'E',
        link: function(scope, element, attrs) {
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

            scope.$watch(
                function() {
                    return attrs.apptitle;
                },
                function(newTitle, oldTitle) {
                    if (newTitle != oldTitle)
                        element.find('.ms-core-pageTitle').text(newTitle);
                }
            );
        }
    };
}]);

shpUtility.directive('ngPeoplePicker', function() {
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
        } else {
            schema.AllowMultipleValues = false;
        }

        SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function() {
            SPClientPeoplePicker_InitStandaloneControlWrapper(elementId, null, schema);
        });

        function onUserResolve() {
            var peoplePickerDictKey = elementId + "_TopSpan";
            var peoplePicker = SPClientPeoplePicker.SPClientPeoplePickerDict[peoplePickerDictKey];
            var people = peoplePicker.GetAllUserInfo();
            var returnValues = [];
            if (returnUsername) {
                angular.forEach(people, function(person) {
                    returnValues.push(person.AutoFillKey);
                });
            } else {
                angular.forEach(people, function(person) {
                    returnValues.push(person.EntityData.Email);
                });
            }
            if (schema.AllowMultipleValues) {
                ngModel.$setViewValue(returnValues);
            } else if (returnValues.length > 0) {
                ngModel.$setViewValue(returnValues[0]);
            }
        }
    }
});
DegSharepointUtility.List = {

    CreateAtHost: function(listName, callback, listTemplate) {
        var deferred = $q.defer();
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
        listCreationInfo.set_templateType(listTemplate);

        var oList = oWebsite.get_lists().add(listCreationInfo);

        appContext.load(oList);
        appContext.executeQueryAsync(
            function(sender, args) {
                checkCallback();
            },
            function(sender, args) {
                $log.log('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
                deferred.reject("Error creating " + listName + " at host.");
            }
        );

        function checkCallback() {
            if (callback) {
                callback();
            }
            deferred.resolve();
        }
    },
    AddFieldToListAtHost: function(listName, fieldDisplayName, fieldName, required, fieldType, fieldExtra, callback) {

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
        switch (fieldType) {
            case "URL":
            case "DateTime":
                if (fieldExtra != "")
                    extraTypeDefinition = "Format=\'" + fieldExtra + "\'";
                break;
            case "Note":
                if (fieldExtra != "")
                    extraTypeDefinition = "RichText=\'" + ((fieldExtra) ? "TRUE" : "FALSE") + "\'";
                break;
            case "User":
                if (fieldExtra != "")
                    extraTypeDefinition = "UserSelectionMode=\'" + fieldExtra + "\'";
                break;
        }


        var requiredTxt = (required) ? "TRUE" : "FALSE";
        var fieldDefinition = "<Field DisplayName=\'" + fieldDisplayName + "\' Name=\'" + fieldName + "\' Required=\'" + requiredTxt + "\' Type=\'" + fieldType + "\' " + extraTypeDefinition + "/>";
        var oField = oList.get_fields().addFieldAsXml(
            fieldDefinition,
            true,
            SP.AddFieldOptions.addFieldInternalNameHint
        );

        appContext.load(oField);

        switch (fieldType) {
            case "Choice":
                var fieldConverted = appContext.castTo(oField, SP.FieldChoice);
                fieldConverted.set_choices(fieldExtra);
                fieldConverted.update();
                appContext.executeQueryAsync(
                    function(sender, args) {
                        checkCallback();
                    },
                    function(sender, args) {
                        $log.log('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
                        deferred.reject();
                    }
                );
                break;
            case "TaxonomyFieldType":
            case "TaxonomyFieldTypeMulti":
                if (fieldExtra) {
                    var session = SP.Taxonomy.TaxonomySession.getTaxonomySession(appContext);
                    var store = session.getDefaultSiteCollectionTermStore();
                    var group = store.get_groups().getByName(fieldExtra.TaxonomyGroup);
                    var set = group.get_termSets().getByName(fieldExtra.TaxonomySet);

                    appContext.load(store, "Id");
                    appContext.load(set, "Id");
                    appContext.executeQueryAsync(
                        function(sender, args) {
                            var fieldConverted = appContext.castTo(oField, SP.Taxonomy.TaxonomyField);
                            fieldConverted.set_sspId(store.get_id());
                            fieldConverted.set_termSetId(set.get_id());
                            fieldConverted.set_createValuesInEditForm(fieldExtra.CreateValuesInEditForm);
                            fieldConverted.set_allowMultipleValues(fieldExtra.AllowMultipleValues);
                            fieldConverted.update();
                            appContext.executeQueryAsync(
                                function(sender, args) {
                                    checkCallback();
                                },
                                function(sender, args) {
                                    $log.log('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
                                    deferred.reject();
                                }
                            );
                        },
                        function(sender, args) {
                            alert("Error creating TaxonomyFieldType");
                        }
                    );
                }

                break;
            default:
                appContext.executeQueryAsync(
                    function(sender, args) {
                        checkCallback();
                    },
                    function(sender, args) {
                        $log.log('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
                        deferred.reject();
                    }
                );
                break;
        }

        function checkCallback() {
            if (callback) {
                callback();
            }
            deferred.resolve();
        }
        return deferred.promise;

    },

    Exist: function(listName, callback) {
        var deferred = $q.defer();
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
            function(sender, args) {
                checkCallback();
                deferred.resolve();
            },
            function(sender, args) {
                $log.log('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
                deferred.reject();
            }
        );

        function checkCallback() {
            if (callback) {
                callback();
            }

        }

        return deferred.promise;
    }
}
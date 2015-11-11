var DegSharepointUtility = {};

DegSharepointUtility.User = {

    GetCurrentUserName: function(callback) {
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

    GetCurrent: function(callback) {
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

    GetId: function(loginName, callback) {
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

    GetCurrentUserProfileProperties: function() {

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
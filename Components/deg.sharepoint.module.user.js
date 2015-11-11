shpUtility.factory('shpUser', ['$log', '$q', 'shpCommon', function($log, $q, shpCommon) {

    return {
        GetCurrentUserName: getUserName,
        GetCurrent: getCurrentUser,
        GetId: getUserId,
        GetCurrentUserProfileProperties: getCurrentUserProperties
    }

    function getUserName(callback) {
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
    }

    function getCurrentUser(callback) {
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

    function getUserId(loginName, callback) {
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

    function getCurrentUserProperties() {

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

}]);
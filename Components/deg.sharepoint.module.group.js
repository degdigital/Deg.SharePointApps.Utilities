shpUtility.factory('shpGroup', ['$http', 'shpCommon', function($http, shpCommon) {

    //Get URLs
    var hostUrl = shpCommon.SPHostUrl;
    var appweburl = shpCommon.SPAppWebUrl;

    return {
        LoadAtHost: loadGroupAtHostWeb,
        CreateAtHost: createGroupAtHostWeb,
        IsCurrentUserMember: isCurrentUserMemberOfGroup
    }

    /** Groups **/
    function loadGroupAtHostWeb(groupName, callback) {
        var hostWebUrl = hostUrl;
        var serverRelativeUrl = getRelativeUrlFromAbsolute(hostWebUrl);
        var hostWebContext = new SP.ClientContext(getRelativeUrlFromAbsolute(hostWebUrl));

        var groupCollection = hostWebContext.get_web().get_siteGroups();
        var group = groupCollection.getByName(groupName);
        hostWebContext.load(group);

        hostWebContext.executeQueryAsync(onGetGroupSuccess, onGetGroupFail);

        function onGetGroupSuccess() {
            if (callback) callback({
                Success: true,
                Message: "Group loaded from host web successfully"
            });
        }

        function onGetGroupFail(data, args) {
            if (callback) callback({
                Success: false,
                Message: "Failed to load group from host web. Error: " + args.get_message()
            });
        }
    }

    function createGroupAtHostWeb(groupName, callback) {
        var hostWebUrl = hostUrl;
        var serverRelativeUrl = getRelativeUrlFromAbsolute(hostWebUrl);
        var hostWebContext = new SP.ClientContext(getRelativeUrlFromAbsolute(hostWebUrl));

        var group = new SP.GroupCreationInformation();
        group.set_title(groupName);

        var oGroup = hostWebContext.get_web().get_siteGroups().add(group);
        hostWebContext.load(oGroup);

        hostWebContext.executeQueryAsync(onCreateGroupSuccess, onCreateGroupFail);

        function onCreateGroupSuccess() {
            if (callback) callback({ Success: true, Message: "Group created at host web successfully" });
        }
        function onCreateGroupFail(data, args) {
            if (callback) callback({ Success: false, Message: "Failed to create group at host web. Error: " + args.get_message() });
        }
    }

    function isCurrentUserMemberOfGroup(groupName, callback) {
        var hostWebUrl = hostUrl;
        var serverRelativeUrl = shpCommon.GetRelativeUrlFromAbsolute(hostWebUrl);
        var hostWebContext = new SP.ClientContext(shpCommon.GetRelativeUrlFromAbsolute(hostWebUrl));

        var groupCollection = hostWebContext.get_web().get_siteGroups();
        var group = groupCollection.getByName(groupName);
        hostWebContext.load(group);

        var currentUser = hostWebContext.get_web().get_currentUser();
        hostWebContext.load(currentUser);

        var groupUsers = group.get_users();
        hostWebContext.load(groupUsers);

        hostWebContext.executeQueryAsync(onGetGroupsSuccess, onGetGroupsFailure);

        function onGetGroupsSuccess(sender, args) {
            var isUserInGroup = false;
            var groupUserEnumerator = groupUsers.getEnumerator();
            while (groupUserEnumerator.moveNext()) {
                var groupUser = groupUserEnumerator.get_current();
                if (groupUser.get_id() == currentUser.get_id()) {
                    isUserInGroup = true;
                    break;
                }
            }
            if (callback) callback({ Success: true, IsUserInGroup: isUserInGroup });
        }

        function onGetGroupsFailure(sender, args) {
            if (callback) callback({ Success: false, Message: "Failed to create group at host web. Error: " + args.get_message() });
        }
    }
}]);
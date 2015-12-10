shpUtility.factory('shpItem', ['$log', '$q', 'shpCommon', function($log, $q, shpCommon) {

	//Get URLs
	var hostUrl = shpCommon.SPHostUrl;
	var appweburl = shpCommon.SPAppWebUrl;
	
	return {
		Create: createListItem,
		GetAll: getAllItems,
		Update: updateListItem
	}

	function createListItem(listName, listProperties, callback) {

		//Get Contexts
		var appContext = new SP.ClientContext(appweburl);
		var hostContext = new SP.AppContextSite(appContext, hostUrl);
		//Get root web
		var oList = hostContext.get_web().get_lists().getByTitle(listName);

		var itemCreateInfo = new SP.ListItemCreationInformation();
		var oListItem = oList.addItem(itemCreateInfo);

		angular.forEach(listProperties, function(value, key) {
			if (value == "currentuser") {
				var current = hostContext.get_web().get_currentUser();
				oListItem.set_item(key, current);
			} else {
				oListItem.set_item(key, value);
			}

		});
		oListItem.update();

		appContext.load(oListItem);
		appContext.executeQueryAsync(
			function(sender, args) {
				checkCallback();
			},
			function(sender, args) {
				$log.log('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
			}
		);

		function checkCallback() {
			if (callback) {
				callback();
			}
		}
	}

	function getAllItems(listName, _query, extend) {

		var deferred = $q.defer();

		var query = (_query) ? _query : '<View><Query></Query></View>';
		//Get Contexts
		var appContext = new SP.ClientContext(appweburl);
		var hostContext = new SP.AppContextSite(appContext, hostUrl);

		var oList = hostContext.get_web().get_lists().getByTitle(listName);

		var camlQuery = new SP.CamlQuery();
		camlQuery.set_viewXml(query);

		var oListItems = oList.getItems(camlQuery);

		appContext.load(oListItems);
		appContext.executeQueryAsync(
			Function.createDelegate(this, function() {
				var entries = [];
				var itemsCount = oListItems.get_count();
				for (var i = 0; i < itemsCount; i++) {
					var item = oListItems.itemAt(i);
					entries.push(item.get_fieldValues());
				}
				deferred.resolve(entries);
			}),
			Function.createDelegate(this, function() {
				deferred.reject('An error has occurred when retrieving items');
			})
		);

		return deferred.promise;
	}

	function updateListItem(listName, listItemId, listProperties) {
		var deferred = $q.defer();

		/* TODO: Use variables for app and host url's */
		var appContext = new SP.ClientContext(appweburl);
		var hostContext = new SP.AppContextSite(appContext, hostUrl);

		var oList = hostContext.get_web().get_lists().getByTitle(listName);
		var oListItem = oList.getItemById(listItemId);

		angular.forEach(listProperties, function(value, key) {
			oListItem.set_item(key, value);
		});
		oListItem.update();

		appContext.executeQueryAsync(
			Function.createDelegate(this, function() {
				deferred.resolve();
			}),
			Function.createDelegate(this, function() {
				deferred.reject('An error has occurred when updating the item.');
			})
		);

		return deferred.promise;
	}

}]);
DegSharepointUtility.Item = {


	Create: function(listName, listProperties, callback) {

		//Get URLs
		var hostUrl = getHostWebUrl();
		var appweburl = getAppWebUrl();
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
	},

	GetAll: function(listName, _query, extend) {

		var deferred = $q.defer();

		var query = (_query) ? _query : '<View><Query></Query></View>';
		//Get URLs
		var hostUrl = getHostWebUrl();
		var appweburl = getAppWebUrl();
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
	},
	,
	Update: function(listName, listItemId, listProperties) {
		var deferred = $q.defer();

		var appContext = new SP.ClientContext(getAppWebUrl());
		var hostContext = new SP.AppContextSite(appContext, getHostWebUrl());

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




}
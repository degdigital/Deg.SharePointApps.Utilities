shpUtility.factory('shpContentType', ['$q', function($q) {

	return {
		CreateAtHost = function(cTypeInfo, callback) {
			var hostWebContext = getHostWebContext();
			createCtype(cTypeInfo, callback, hostWebContext);
		},
		Create = createCtype
	}


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


}]);
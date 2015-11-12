shpUtility.factory('shpColumn', ['$http', 'shpCommon', function($http, shpCommon) {

	var hostWebContext = shpCommon.HostWebContext;

	return {
		CreateAtHost: createFields
	}

	function createFields(fieldsXml, callback) {

		var clientCtx = SP.ClientContext.get_current();
		// field XML format :: <Field DisplayName='Field Name' Name='NoSpaceName' ID='{2d9c2efe-58f2-4003-85ce-0251eb174096}' Group='Group Name' Type='Text' />
		var ctx = hostWebContext || clientCtx;
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
			ctx.executeQueryAsync(function() {
				$log.log("Field provisioned in host web successfully");
				checkCallback();
			}, function(sender, args) {
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

}]);
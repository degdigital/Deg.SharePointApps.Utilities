shpUtility.factory('shpCommon', function() {

	return {
		GetFormDigest: getFormDigest,
		SPAppWebUrl: getAppWebUrl(),
		SPHostUrl: getHostWebUrl(),
		HostWebContext: getHostWebContext,
		GetQsParam: getUrlParam
	}

	function getUrlParam(key) {
		var vars = [],
			hash;
		var hashes = window.location.href.slice(window.location.href.indexOf('?') + 1).split('&');
		for (var i = 0; i < hashes.length; i++) {
			hash = hashes[i].split('=');
			vars.push(hash[0]);
			vars[hash[0]] = hash[1];
		}
		return vars[key];
	}

	function getFormDigest(callback) {
		$.ajax({
			url: getAppWebUrl() + "/_api/contextinfo",
			type: "POST",
			headers: {
				"accept": "application/json;odata=verbose",
				"contentType": "text/xml"
			},
			success: function(data) {
				var requestDigest = data.d.GetContextWebInformation.FormDigestValue;
				if (callback) callback({
					error: false,
					requestDigest: requestDigest
				});
			},
			error: function(error) {
				if (callback) callback({
					error: true,
					errorMessage: JSON.stringify(error)
				});
			}
		});
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

	function getAppWebUrl() {
		return decodeURIComponent(getUrlParam("SPAppWebUrl"));
	}

	function getHostWebUrl() {
		return decodeURIComponent(getUrlParam("SPHostUrl"));
	}

	function getHostWebContext() {
		var hostUrl = getHostWebUrl();
		var hostWebContext = new SP.ClientContext(getRelativeUrlFromAbsolute(hostUrl));
		return hostWebContext;
	}

});
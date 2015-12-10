## Setup environment

```bash
npm install --save-dev gulp
npm install --save-dev gulp-concat gulp-uglify gulp-rename
```

## Download it using bower

```bash
* bower install DegSharepointUtilities
```

# Deg.SharePointApps.Utilities
* Automatically resize app part iframes
* Directive for client ribbon bar
* Utilities for SharePoint property bag management
* CRUD List Operations
* Field and Content Types provisioning
* Helpers for creating and publishing Files
* App Context Helpers (AppUrl, HostUrl, currentUser, etc)


# Common (shpCommon)
* GetFormDigest
* SPAppWebUrl
* SPHostUrl
* HostWebContext
* GetQsParam
* GetRelativeUrlFromAbsolute

## Usage

```js
var resultsPerPage = spService.Utilities.GetQsParam("ResultsPerPage");
```


# ContentType (shpContentType)
* CreateAtHost: Creates a content type in root site.

# Item (shpItem)
* Create
* GetAll
* Update

# List (shpList)
* CreateAtHost
```js
var contactList = "Contacts";
spService.Lists.CreateAtHost(contactList, createFields);
```
* AddFieldToListAtHost
```js
//Generic example
spService.Lists.AddFieldToListAtHost(LISTNAME, DISPLAYNAME, INTERNALNAME, bool:REQUIRED, TYPE, FIELD EXTRA, function:CALLBACK)

var createFields = function () {
	//Text
    spService.Lists.AddFieldToListAtHost(contactList, "Email", "fuseEmail", false, "Text", "", null)
    .then(function () {
    	//User
        return spService.Lists.AddFieldToListAtHost(contactList, "User", "fuseUser", false, "User", "", null);
    })
    .then(function () {
    	//Image
        return spService.Lists.AddFieldToListAtHost(contactList, "Picture", "fusePic", false, "URL", "Image", null);
    })
    .then(function () {
    	//Hyperlink
        return spService.Lists.AddFieldToListAtHost(contactList, "Profile Page", "fuseLink", false, "URL", "Hyperlink", null);
    })
    .then(function () {
    	//TaxonomyFieldType
        return spService.Lists.AddFieldToListAtHost(contactList, "Enterprise Keywords", "TaxKeyword", false, "TaxonomyFieldType", null, null);
    });
};
```
* Exist
```js
spService.Lists.Exist(contactList, function () {
    ..
});
```

# Column (shpColumn)
* CreateAtHost

# File (shpFile)
* CreateAtHost
* LoadAtHost
* CheckOutAtHost
* PublishFileToHost
* UploadFileToHostWeb

# Group (shpGroup)
* LoadAtHost
* CreateAtHost
* IsCurrentUserMember

# PropertyBag (shpPropertyBag)
* SaveObjToCurrentWeb
* SaveObjToRootWeb
* GetValue

# Taxonomy (shpTaxonomy)
* GetTermSetValues

# User (shpUser)
* GetCurrentUserName
* GetCurrent
* GetId
* GetCurrentUserProfileProperties

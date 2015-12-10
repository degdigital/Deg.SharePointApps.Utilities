## Setup environment

```bash
npm install --save-dev gulp
npm install --save-dev gulp-concat gulp-uglify gulp-rename
```

## Download it using bower

```bash
bower install DegSharepointUtilities
```

# Directives
* Automatically resize app part iframes
```html
<ng-app-frame>
</ng-app-frame>
```
* People Picker 
```html
<div ng-people-picker accounttype="SPGroup" ></div>

<div ng-people-picker accounttype="User" ></div>
```
* Directive for client ribbon bar
* Utilities for SharePoint property bag management
* CRUD List Operations
* Field and Content Types provisioning
* Helpers for creating and publishing Files
* App Context Helpers (AppUrl, HostUrl, currentUser, etc)


# Common (shpCommon)
* GetFormDigest
```js
spService.Utilities.GetFormDigest(function (result) {
	..
});
```
* SPAppWebUrl
* SPHostUrl
* HostWebContext
* GetQsParam
```js
var resultsPerPage = spService.Utilities.GetQsParam("ResultsPerPage");
```
* GetRelativeUrlFromAbsolute
```js
var relativeUrl = spService.Utilities.GetRelativeUrlFromAbsolute("url");
```



# ContentType (shpContentType)
* CreateAtHost: Creates a content type in root site.
```js
```

# Item (shpItem)
* Create
```js
```
* GetAll
```js
```
* Update
```js
```

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
```js
spService.Groups.CreateAtHost("FUSE Picture Gallery Admins");
```
* IsCurrentUserMember
```js
spService.Groups.IsCurrentUserMember('FUSE Picture Gallery Admins',
    function (result) {
        ..
    }
);
```
# PropertyBag (shpPropertyBag)
* SaveObjToCurrentWeb
* SaveObjToRootWeb
* GetValue

# Taxonomy (shpTaxonomy)
* GetTermSetValues
```js
spService.Taxonomy.GetTermSetValues("FUSE", "FUSE Business Units", function (termStoreValues) { 
..
});
```

# User (shpUser)
* GetCurrentUserName
* GetCurrent
* GetId
* GetCurrentUserProfileProperties

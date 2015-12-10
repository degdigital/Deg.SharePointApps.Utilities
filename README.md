# Setup environment

* npm install --save-dev gulp
* npm install --save-dev gulp-concat gulp-uglify gulp-rename

# Download it using bower

* bower install DegSharepointUtilities


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

# ContentType (shpContentType)
* CreateAtHost: Creates a content type in root site.

# Item (shpItem)
* Create
* GetAll
* Update

# List (shpList)
* CreateAtHost
* AddFieldToListAtHost
* Exist

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

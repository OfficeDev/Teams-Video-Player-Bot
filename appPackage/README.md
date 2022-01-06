# App Package

Your Teams "app package" is used to publish your app configuration to Teams by manually using the "Upload a custom app" from the app gallery in Teams.

This package can also be used to manually publish your app to the Teams store or to your organization's store.

You need to provide and remplace the following information in the manifest.json file:
- **`<<YOUR-MICROSOFT-APP-ID>>`** - Is the Application ID (aka client ID) from Azure AD
- **`<<VIDEO-SITE-DOMAIN>>`** - Is the domain name of where your videos are hosted - You can remove this property if you use SharePoint only
- **`<<YOUR-SPO-DOMAIN>>`** - Is the domain name of your SharePoint Online site (if videos are hosted on SharePoint)

To create the ZIP package, selection the files 'outline.png', 'color.png' and 'manifest.json' to create a ZIP file. Please make sure the 'manifest.json' is updated first based on the above instrusctions and that you select the 3 files direclty to create the ZIP file (don't ZIP the directory that contains the files)

**NOTE: Values in the manifest.json file will need to be updated before creating the zip package to use for publishing.**
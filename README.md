# App Insights Web Part

## Summary

Add Azure Application Insights tracking to a page using this SPFx web part.

![Screenshot of edit panel](/images/webpart-edit-panel.png)

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.4.1-green.svg)

**Compatibility:** SharePoint Online, SharePoint 2019

## Setup Instructions
### Pre-requisites
- App Catalog: Ensure the [App Catalog](https://docs.microsoft.com/en-us/sharepoint/use-app-catalog) is setup in your SharePoint Online tenant.

### Tenant Installation
1. Download the latest SPFx package file from [releases](https://github.com/bschlintz/spfx-appinsights-webpart/releases/latest) or clone the repo and build the package yourself.
1. Upload sppkg file to the 'Apps for SharePoint' library in your Tenant App Catalog.
1. Click Deploy

### Site Installation
1. Go to the target site, then **Add an App**
1. Click the `Azure App Insights Web Part` app to install it in your site
1. Edit the page where you want to add the web part, then add the `Azure App Insights` web part and configure it
1. Save and publish the page

### Updates
Follow the same steps as installation. Overwrite the existing package in the 'Apps for SharePoint' library when uploading the new package. 

> __Tip__: Be sure to check-in the sppkg file after the deployment if it is left checked-out.

## Version history

Version|Date|Comments
-------|----|--------
1.0.0|December 18, 2020| Initial release.

## Disclaimer
Microsoft provides programming examples for illustration only, without warranty either expressed or implied, including, but not limited to, the implied warranties of merchantability and/or fitness for a particular purpose. We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys' fees, that arise or result from the use or distribution of the Sample Code.

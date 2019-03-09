/*
 * @name SPCustomFormUpdater
 * Version 1.0.0
 * No dependencies!
 * @description
 * @documentation
 * @author Sam Perrow sam.perrow399@gmail.com
 *
 * Copyright 2019 Sam Perrow (email : sam.perrow399@gmail.com)
 * Licensed under the MIT license:
 * http://www.opensource.org/licenses/mit-license.php
 */

// bugs:
// 1. 

// Part 1.
// get all URL's of subsites in a site collection
// SP.SOD.executeFunc('sp.js', 'SP.ClientContext', GetSubSiteUrls);

var GetSubSiteUrls = function () {
    function enumWebs(propertiesToRetrieve, success, error) {
        var ctx = new SP.ClientContext.get_current();
        var rootWeb = ctx.get_site().get_rootWeb();
        var result = [];
        var level = 0;
        ctx.load(rootWeb, propertiesToRetrieve);
        result.push(rootWeb);
        var colPropertiesToRetrieve = String.format('Include({0})', propertiesToRetrieve.join(','));
        var enumWebsInner = function (web, result, success, error) {
            level++;
            var ctx = web.get_context();
            var webs = web.get_webs();
            ctx.load(webs, colPropertiesToRetrieve);
            ctx.executeQueryAsync(
                function () {
                    for (var i = 0; i < webs.get_count(); i++) {
                        var web = webs.getItemAtIndex(i);
                        result.push(web);
                        enumWebsInner(web, result, success, error);
                    }
                    level--;
                    if (level == 0 && success)
                        success(result);
                },
                fail);
        };
        enumWebsInner(rootWeb, result, success, error);
    }

    function success(sites) {
        var urls = [];
        for (var i = 1; i < sites.length; i++) {
            urls.push(sites[i].get_url());
        }
        console.log(urls);
        return getListGuidsForSubsites(urls); // uncomment this to activate the script and update forms!
    }

    function fail(sender, args) {
        console.log(args.get_message());
    }
    enumWebs(['Url', 'Fields'], success, fail);
}

//part 2: filter down list of provided subsites and return the ones that have the targeted list
getListGuidsForSubsites( ["https://carepoint.health.mil/sites/VHCCA/RHCAtlantic/Dev"]);

function getListGuidsForSubsites(sites) {
    var targetSites = [];
    var siteIndex = 0;
    var listName = 'Appointments';
    controller();

    function controller() {
        if (siteIndex < sites.length) {
            getListGuid(sites[siteIndex]);
        } else {
            console.log('done getting list guids');
            return GetListAndFormData(targetSites);
        }
    }

    function getListGuid(siteURL) {
        var currentcontext = new SP.ClientContext(siteURL);
        var list = currentcontext.get_web().get_lists().getByTitle(listName);
        currentcontext.load(list, 'Id');
        currentcontext.executeQueryAsync(
            function (sender, args) {
                var listGuid = list.get_id().toString();
                targetSites.push({
                    subsiteURL: siteURL,
                    listGUID: listGuid
                });
                console.log('list found at ' + siteURL);
                siteIndex++;
                controller();
            },
            function (sender, args) {
                console.warn(args.get_errorDetails);
                console.log('No list at the subsite: ' + siteURL);
                siteIndex++;
                controller();
            });
    }
}

/*
 * Part 3.
 * create new target forms for each target site.
 * a) get source file content (from .txt files)
 * b) get web part ID for each form
 * c) create new form with the web part ID's, listGuids, and siteURL
 * SP.SOD.executeFunc('sp.js', 'SP.ClientContext', GetListAndFormData);
*/

function GetListAndFormData(targetSites) {

    var listName = 'Appointments';
    var sourceFormData = [];
    var siteIndex = 0;
    var formIndex = 0;
    var path = "https://carepoint.health.mil/sites/VHCCA/assets/formUpdater/appts/";
    var title = '';

    var sourceFileUrls = [{
            filePath: path + 'NewForm.txt',
            title: 'NewForm'
        },
        {
            filePath: path + 'EditForm.txt',
            title: 'EditForm'
        },
        {
            filePath: path + 'DispForm.txt',
            title: 'DispForm'
        }
    ];

    function appendToSiteObj(siteObj) {
        siteObj.listName = listName;

        siteObj.targetForms = [{
                filePath: siteObj.subsiteURL + '/Lists/' + listName + '/NewForm.aspx',
                title: 'NewForm'
            },
            {
                filePath: siteObj.subsiteURL + '/Lists/' + listName + '/EditForm.aspx',
                title: 'EditForm'
            },
            {
                filePath: siteObj.subsiteURL + '/Lists/' + listName + '/DispForm.aspx',
                title: 'DispForm'
            }
        ];
        return siteObj;
    }

    if (siteIndex === 0 && formIndex === 0) {
        getSourceFormData(formIndex);
    }

    function getSourceFormData(formIndex) {
        if (sourceFormData.length < sourceFileUrls.length) {
            getFileContent(sourceFileUrls[formIndex].filePath, createSourceFile);
        } else {
            console.log('Finished collecting all source forms. Proceeding to prepare form data for: ' + targetSites[siteIndex].subsiteURL);
            init();
        }
    }

    function createSourceFile(sourceFileUrl, responseText) {
        var _title = sourceFileUrls[formIndex].title;

        sourceFormData.push({
            url: sourceFileUrl,
            title: _title,
            fileContent: responseText
        });

        formIndex++;
        getSourceFormData(formIndex);
    }

    function init() {
        if (siteIndex < targetSites.length) {
            formIndex = 0;
            appendToSiteObj(targetSites[siteIndex]);
            controller();
        } else {
            console.log( targetSites );
            console.log( 'Finished preparing all form data. Ready to begin updating...');
            UpdateForms(targetSites);
        }
    }

    function controller() {

        if (formIndex < sourceFileUrls.length) {

            if (formIndex < targetSites[siteIndex].targetForms.length) {
                getFileContent(targetSites[siteIndex].targetForms[formIndex].filePath, getWebPartId);
            }

        } else {
            updateSiteObjsWithNewFile();
        }

    }

    function getWebPartId(index, response) {
        var webPartId = '';
        var regex = new RegExp(/(?<=\<div\sWebPartID=")(.*?)(?=\")/, 'ig');
        if (regex.test(response)) {
            webPartId = response.match(regex)[1].toUpperCase(); // the first match is a guid with all zero's. need to improve regex search.
        }
        targetSites[siteIndex].targetForms[formIndex].webPartId = webPartId;
        formIndex++;
        return controller();
    }

    function getFileContent(siteURL, callback) {
        var xhttp = new XMLHttpRequest();
        xhttp.onreadystatechange = function () {
            if (this.readyState == 4 && this.status == 200) {
                return callback(siteURL, this.responseText);
            }
        }
        xhttp.open("GET", siteURL, true);
        xhttp.send();
    }

    function updateSiteObjsWithNewFile() {

        for (var i = 0; i < targetSites[siteIndex].targetForms.length; i++) {
            for (var j = 0; j < sourceFormData.length; j++) {
                if (targetSites[siteIndex].targetForms[i].title === sourceFormData[j].title) {
                    targetSites[siteIndex].targetForms[i].newContent = sourceFormData[j].fileContent;
                    break;
                }
            }
        }
        parseSourceFileContent(targetSites[siteIndex]);

        console.log('finished preparing the form data for: ' + targetSites[siteIndex].subsiteURL);
        siteIndex++;
        init();
    }

    function parseSourceFileContent(targetSite) {
        for (var i = 0; i < targetSite.targetForms.length; i++) {
            var webPartElem = targetSite.targetForms[i].newContent.match(/<WebPartPages:DataFormWebPart(.*?)>/i)[1];
            var oldGuid = webPartElem.match(/ListName="{(.*?)}"/i)[1];
            var oldWebPartId = webPartElem.match(/__WebPartId="{(.*?)}"/i)[1];
            var newWebPartId = targetSite.targetForms[i].webPartId;
            var oldGuidRegex = new RegExp(oldGuid, 'ig');
            var oldWebPartIdRegex = new RegExp(oldWebPartId, 'ig');
            var newFile = targetSite.targetForms[i].newContent.replace(oldGuidRegex, targetSite.listGuid).replace(oldWebPartIdRegex, newWebPartId);
            var originalSiteName = targetSite.targetForms[i].newContent.match(/<ParameterBinding Name="weburl" Location="None" DefaultValue="(.*?)"\/>/i)[1];
            var origSiteNameRegex = new RegExp(originalSiteName, 'ig');
            if (newWebPartId && targetSite.siteUrl !== originalSiteName) {
                newFile = newFile.replace(origSiteNameRegex, targetSite.siteUrl);
            }
            targetSite.targetForms[i].newContent = newFile;
        }
    }

}

/*
 * Part 4.
 * update the forms for each target site, one at a time.
*/

var UpdateForms = function (targetSites) {

    var siteIndex = 0;
    var formIndex = 0;

    controller();
    function controller() {

        if (siteIndex < targetSites.length) {

            if (formIndex < targetSites[siteIndex].targetForms.length) {
                console.log( 'updating: ' + targetSites[siteIndex].targetForms[formIndex].title );
                return updateFile(targetSites[siteIndex]);
            } else {
                formIndex = 0;
                siteIndex++;
                return controller();
            }

        } else {
            console.log( 'Finished update!');
        }
    }

    function updateFile(targetSite) {
                // debugger;

        var clientContext = new SP.ClientContext(targetSite.subsiteURL);
        var list = clientContext.get_web().get_lists().getByTitle(targetSite.listName);
        clientContext.load(list);
        
        var newFileUrl = targetSite.targetForms[formIndex].filePath;
        var fileCreateInfo = new SP.FileCreationInformation();
        fileCreateInfo.set_content(new SP.Base64EncodedByteArray());
        fileCreateInfo.set_url(newFileUrl);
        fileCreateInfo.set_overwrite(true);

        for (var i = 0; i < targetSite.targetForms[formIndex].newContent.length; i++) {
            fileCreateInfo.get_content().append(targetSite.targetForms[formIndex].newContent.charCodeAt(i));
        }

        var newFile = list.get_rootFolder().get_files().add(fileCreateInfo);
        clientContext.load(newFile);
        clientContext.executeQueryAsync(Function.createDelegate(this, successHandler), Function.createDelegate(this, errorHandler));
    }

    function successHandler(sender, args) {
        console.log('Successfully updated: ' + targetSite.targetForms[formIndex].title);
        formIndex++;
        controller();
    }

    function errorHandler(sender, args) {
        formIndex++;
        controller();
        console.log(args.get_message());
    }
    
}
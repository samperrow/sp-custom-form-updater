/*
* @name SPCustomFormUpdater
* Version 1.0.0
* No dependencies!
* @description 
* @documentation 
* @author Sam Perrow sam.perrow399@gmail.com
*
* Copyright 2019  Sam Perrow  (email : sam.perrow399@gmail.com)
* Licensed under the MIT license:
* http://www.opensource.org/licenses/mit-license.php
*/

// bugs:
// 1. errors out when there is not a list at target site.


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

        return getListGuidsForSubsites(urls);                // uncomment this to activate the script and update forms!
    }

    function fail(sender, args) {
        console.log(args.get_message());
    }

    enumWebs(['Url', 'Fields'], success, fail);
}





 //part 2: filter down list of provided subsites and return the ones that have the targeted list

function getListGuidsForSubsites(sites) {
    var targetSiteFormData = [];
    var siteIndex = 0;
    var listName = 'Appointments';

    controller();
    function controller() {

        if (siteIndex < sites.length) {
            getListGuid(sites[siteIndex]);
        } else {
            console.log('done getting list guids');
            console.log(targetSiteFormData);
            return GetListAndFormData(targetSiteFormData);
        }

    }

    function getListGuid(siteURL) {
        var currentcontext = new SP.ClientContext(siteURL);
        var list = currentcontext.get_web().get_lists().getByTitle(listName);
        currentcontext.load(list, 'Id');
        currentcontext.executeQueryAsync(
            function (sender, args) {
                var listGuid = list.get_id().toString();
                targetSiteFormData.push({ subsiteURL: siteURL, listGUID: listGuid });
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
 * a) get web part ID for each form 
 * b) get source file content (from .txt files)
 * c) create new form with the web part ID's, listGuids, and siteURL
* SP.SOD.executeFunc('sp.js', 'SP.ClientContext', GetListAndFormData);
*/

// GetListAndFormData([]);

function GetListAndFormData(targetSites) {

    var listName = 'Appointments';
    var targetSiteFormData = [];
    var sourceFormData = [];
    var siteIndex = 0;
    var formIndex = 0;
    var domain = "https://sarcastasaur.sharepoint.com";
    var title = '';

    var sourceFileUrls = [
        { filePath: domain + '/formUpdater/NewForm.txt', title: 'NewForm' },
        { filePath: domain + '/formUpdater/EditForm.txt', title: 'EditForm' },
        { filePath: domain + '/formUpdater/DispForm.txt', title: 'DispForm' }
    ];

    function TargetSiteFormData(url, listGuid) {
        this.siteUrl = url;
        this.listName = listName;

        if (listGuid) {
            this.targetForms = [
                { filePath: url + '/Lists/' + listName + '/NewForm.aspx', title: 'NewForm' },
                { filePath: url + '/Lists/' + listName + '/EditForm.aspx', title: 'EditForm' },
                { filePath: url + '/Lists/' + listName + '/DispForm.aspx', title: 'DispForm' }
            ];
        }

        this.listGuid = listGuid;
    }

    if (targetSites.length > 0) {
        getSourceFormData(formIndex);
    } else {
        console.log('There are no sites to update.');
    }

    function getSourceFormData(formIndex) {

        if (sourceFormData.length < sourceFileUrls.length) {
            getSourceFileContent(sourceFileUrls[formIndex].filePath, createSourceFile);
        } else {
            console.log('Finished collecting all source form data. Proceeding to collect list guids and web part ids.');
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


    function controller() {

        if (formIndex < sourceFileUrls.length) {

            // get the web part ID for each of the forms.
            if (formIndex < targetSiteFormData[siteIndex].targetForms.length) {
                getSourceFileContent(targetSiteFormData[siteIndex].targetForms[formIndex].filePath, getWebPartId);
            }

        } else if (siteIndex === (targetSiteFormData.length - 1) && formIndex === sourceFileUrls.length) {
            siteIndex++;
            formIndex = 0;
            console.log('Now getting data for: ' + sites[siteIndex]);
            controller();
        } else if (targetSiteFormData.length === sites.length) {
            console.log('Done getting list guid\'s');
            updateSiteObjsWithNewFile();
        }

    }


    function getWebPartId(index, response) {
        var webPartId = '';
        var regex = new RegExp(/(?<=\<div\sWebPartID=")(.*?)(?=\")/, 'ig');

        if (regex.test(response)) {
            webPartId = response.match(regex)[1].toUpperCase();                 // the first match is a guid with all zero's. need to improve regex search.
        }

        targetSiteFormData[siteIndex].targetForms[formIndex].webPartId = webPartId;
        formIndex++;
        return controller();
    }


    function getSourceFileContent(siteURL, callback) {
        var xhttp = new XMLHttpRequest();
        xhttp.onreadystatechange = function () {
            if (this.readyState == 4 && this.status == 200) {
                return callback(siteURL, this.responseText);
            }
        }
        xhttp.open("GET", siteURL, true);
        xhttp.send();
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

    function updateSiteObjsWithNewFile() {

        targetSiteFormData = targetSiteFormData.filter(function (item) {
            return item.targetForms;
        });
        // console.log( targetSiteFormData );

        for (var i = 0; i < targetSiteFormData.length; i++) {
            for (var j = 0; j < targetSiteFormData[i].targetForms.length; j++) {
                for (var k = 0; k < sourceFormData.length; k++) {
                    if (targetSiteFormData[i].targetForms[j].title === sourceFormData[k].title) {
                        targetSiteFormData[i].targetForms[j].newContent = sourceFormData[k].fileContent;
                        break;
                    }
                }
            }
            parseSourceFileContent(targetSiteFormData[i]);
        }
        console.log(targetSiteFormData);
        console.log('Ready to begin the updating of target forms...');
        // UpdateForms(targetSiteFormData);             
    }

}










/*
* Part 4.
* update the forms for each target site, one at a time.
*/
var UpdateForms = function (targets) {

    var siteIndex = 0;
    var formIndex = 0;

    startFileUpdating(siteIndex, formIndex);

    function startFileUpdating(siteIndex, formIndex) {
        return updateFile(targets[siteIndex].siteUrl, targets[siteIndex].targetForms[formIndex]);
    }


    function updateFile(targetSiteUrl, sourceFile) {
        var thisSite = targets[siteIndex];
        var clientContext = new SP.ClientContext(targetSiteUrl);
        var list = clientContext.get_web().get_lists().getByTitle(thisSite.listName);

        var newUrl = targetSiteUrl + '/Lists/' + thisSite.listName + '/' + thisSite.targetForms[formIndex].title + '.aspx';

        clientContext.load(list);

        var fileCreateInfo = new SP.FileCreationInformation();
        fileCreateInfo.set_content(new SP.Base64EncodedByteArray());
        fileCreateInfo.set_url(newUrl);
        fileCreateInfo.set_overwrite(true);

        for (var i = 0; i < thisSite.targetForms[formIndex].newContent.length; i++) {
            fileCreateInfo.get_content().append(thisSite.targetForms[formIndex].newContent.charCodeAt(i));
        }

        var newFile = list.get_rootFolder().get_files().add(fileCreateInfo);
        clientContext.load(newFile);
        clientContext.executeQueryAsync(Function.createDelegate(this, successHandler), Function.createDelegate(this, errorHandler));

        function successHandler(sender, args) {

            console.log('Successfully updated: ' + thisSite.targetForms[formIndex].title + ' on: ' + targetSiteUrl);

            if ((formIndex + 1) < thisSite.targetForms.length) {
                formIndex++;
                startFileUpdating(siteIndex, formIndex);
            } else if ((formIndex + 1) >= thisSite.targetForms.length) {

                if ((siteIndex + 1) < targets.length) {
                    siteIndex++;
                    formIndex = 0;

                    startFileUpdating(siteIndex, formIndex);
                } else {
                    console.log('Completed the update!');
                }
            }
        }

        function errorHandler(sender, args) {
            console.log(args.get_message());
        }
    }
}
/*
 * @name SP Custom Form Updater
 * Version 1.0.0
 * No dependencies!
 * @description A way to utilize client side JSOM to make updates to custom .aspx SharePoint forms.
 * @documentation https://github.com/samperrow/sp-custom-form-updater
 * @author Sam Perrow https:sharepointhacks.com
 *
 * Copyright 2019 Sam Perrow (email : sam.perrow399@gmail.com)
 * Licensed under the MIT license:
 * http://www.opensource.org/licenses/mit-license.php
*/

// Example of how to initialize the script.
var init = function() {

    var listName = 'Appointments';
    var sourceFilePath = "https://example.sharepoint.com/assets";

    var sourceFiles = [
        { filePath: sourceFilePath + '/NewForm.txt',  title: 'NewForm'  },
        { filePath: sourceFilePath + '/EditForm.txt', title: 'EditForm' },
        { filePath: sourceFilePath + '/DispForm.txt', title: 'DispForm' }
    ];

    return CustomFormUpdater(listName, sourceFiles);

}();





function CustomFormUpdater(listName, sourceFiles, sites) {

    // if the sites parameter is available, use that, otherwise collect all subsites URL's.
    (sites) ? GetListGuidsForSubsites(sites) : GetSubSiteUrls();


    // Part 1.
    // Collect all URL's of subsites in a site collection
    function GetSubSiteUrls() {
        
        // borrowed from Vadim Gremyachev: https://sharepoint.stackexchange.com/questions/137996/get-all-subsites-and-rootweb-of-a-site-collection-using-jsom#answer-137998
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
                ctx.executeQueryAsync(function () {
                    for (var i = 0; i < webs.get_count(); i++) {
                        var web = webs.getItemAtIndex(i);
                        result.push(web);
                        enumWebsInner(web, result, success, error);
                    }
                    level--;
                    if (level == 0 && success)
                        success(result);
                }, fail);
            };
            enumWebsInner(rootWeb, result, success, error);
        }

        function success(sites) {
            var urls = [];
            for (var i = 1; i < sites.length; i++) {
                urls.push(sites[i].get_url());
            }

            return GetListGuidsForSubsites(urls);
        }

        function fail(sender, args) {
            console.log(args.get_message());
        }
        enumWebs(['Url', 'Fields'], success, fail);
    }


    /*
    * Part 2.
    * Filter down array of subsite URL's and return the ones that have the specified list.
    * @param is an array of strings, each of which is a subsite URL.
    */
    function GetListGuidsForSubsites(sites) {
        var targetSites = [];
        var siteIndex = 0;
        
        function controller() {
            if (siteIndex < sites.length) {
                getListGuid(sites[siteIndex]);
            } else if (targetSites.length > 0) {
                console.log('Done collecting list guids.');
                return GetListAndFormData(targetSites);
            }
        }
        controller();

        function createTargetFiles(siteURL) {
            var arr = [];

            for (var i = 0; i < sourceFiles.length; i++) {
                arr.push({ filePath: siteURL + '/Lists/' + listName + '/' + sourceFiles[i].title + '.aspx', title: sourceFiles[i].title });
            }
            return arr;
        }

        function getListGuid(siteURL) {
            var currentcontext = new SP.ClientContext(siteURL);
            var list = currentcontext.get_web().get_lists().getByTitle(listName);
            currentcontext.load(list, 'Id');
            currentcontext.executeQueryAsync(
                function() {
                    var listGuid = list.get_id().toString();
                    console.log('List found at ' + siteURL);

                    targetSites.push({
                        subsiteURL: siteURL,
                        listGUID: listGuid,
                        listName: listName,
                        targetForms: createTargetFiles(siteURL)
                    });
                    siteIndex++
                    return controller();
                },
                function (sender, args) {
                    console.warn(args.get_message());
                    siteIndex++
                    return controller();
                });
        }
    }



    /*
    * Part 3.
    * create new target forms for each target site.
    * a) get source file content (from .txt files)
    * b) get web part ID for each form
    * c) create new form with the web part ID's, listGuids, and siteURL
    * @param is an array of objects, each containing the subsite URL, list name, and the list GUID specified to that site.
    */
    function GetListAndFormData(targetSites) {

        var sourceFormData = [];
        var siteIndex = 0;
        var formIndex = 0;

        if (siteIndex === 0 && formIndex === 0) {
            getSourceFormData();
        }

        function getSourceFormData() {
            return (sourceFormData.length < sourceFiles.length) ? getFileContent(sourceFiles[formIndex].filePath, createSourceFile) : init();
        }

        function createSourceFile(sourceFileUrl, responseText) {
            var _title = sourceFiles[formIndex].title;

            sourceFormData.push({
                url: sourceFileUrl,
                title: _title,
                fileContent: responseText
            });

            formIndex++;
            return getSourceFormData();
        }

        function init() {
            if (siteIndex < targetSites.length) {
                console.log('Finished collecting all source forms. Proceeding to prepare form data for: ' + targetSites[siteIndex].subsiteURL);
                formIndex = 0;
                controller();
            } else {
                console.log( targetSites );
                console.log( 'Finished preparing all form data. Ready to begin updating...');
                UpdateForms(targetSites);
            }
        }

        function controller() {
            return (formIndex < sourceFiles.length && formIndex < targetSites[siteIndex].targetForms.length) ? getFileContent(targetSites[siteIndex].targetForms[formIndex].filePath, getWebPartId) : updateSiteObjsWithNewFile();
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
            return xhttp.send();
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

            console.log('Finished preparing form data for: ' + targetSites[siteIndex].subsiteURL);
            siteIndex++;
            return init();
        }

        function parseSourceFileContent(targetSite) {

            for (var i = 0; i < targetSite.targetForms.length; i++) {

                var targetFile = targetSite.targetForms[i].newContent;

                var webPartElem = targetFile.match(/<WebPartPages:DataFormWebPart(.*?)>/i)[1];
                var oldGuid = webPartElem.match(/ListName="{(.*?)}"/i)[1];
                var oldWebPartId = webPartElem.match(/__WebPartId="{(.*?)}"/i)[1];

                var newWebPartId = targetSite.targetForms[i].webPartId;
                var oldGuidRegex = new RegExp(oldGuid, 'ig');
                var oldWebPartIdRegex = new RegExp(oldWebPartId, 'ig');
                var originalSiteName = targetFile.match(/<ParameterBinding Name="weburl" Location="None" DefaultValue="(.*?)"\/>/i)[1];
                var origSiteNameRegex = new RegExp(originalSiteName, 'ig');

                if (newWebPartId && targetSite.subsiteURL !== originalSiteName) {
                    targetSite.targetForms[i].newContent = targetFile.replace(origSiteNameRegex, targetSite.subsiteURL).replace(oldGuidRegex, targetSite.listGUID).replace(oldWebPartIdRegex, newWebPartId);
                }
            }
        }

    }




    /*
    * Part 4.
    * update the forms for each target site, one at a time.
    * @param is an array of objects, each of which
    */
    var UpdateForms = function(targetSites) {

        var siteIndex = 0;
        var formIndex = 0;

        function init() {
            return (siteIndex < targetSites.length) ? controller() : console.log('Finished the update!');
        }
        init();

        function controller() {
            if (formIndex < targetSites[siteIndex].targetForms.length) {
                console.log( 'Updating: ' + targetSites[siteIndex].targetForms[formIndex].title );
                return updateFile(targetSites[siteIndex]);
            } else {
                formIndex = 0;
                siteIndex++;
                return init();
            }
        }

        function updateFile(targetSite) {
            var clientContext = new SP.ClientContext(targetSite.subsiteURL);
            var list = clientContext.get_web().get_lists().getByTitle(targetSite.listName);
            clientContext.load(list);

            var fileCreateInfo = new SP.FileCreationInformation();
            fileCreateInfo.set_content(new SP.Base64EncodedByteArray());
            fileCreateInfo.set_url(targetSite.targetForms[formIndex].filePath);
            fileCreateInfo.set_overwrite(true);

            for (var i = 0; i < targetSite.targetForms[formIndex].newContent.length; i++) {
                fileCreateInfo.get_content().append(targetSite.targetForms[formIndex].newContent.charCodeAt(i));
            }

            var newFile = list.get_rootFolder().get_files().add(fileCreateInfo);
            clientContext.load(newFile);
            clientContext.executeQueryAsync(
                function() {
                    console.log('Successfully updated: ' + targetSites[siteIndex].targetForms[formIndex].title + ' on: ' + targetSites[siteIndex].subsiteURL);
                    formIndex++;
                    controller();
                }, 
                Function.createDelegate(this, errorHandler));
        }

        function errorHandler(sender, args) {
            formIndex++;
            controller();
            console.log(args.get_message());
        }
        
    }

}
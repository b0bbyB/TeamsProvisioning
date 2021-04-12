var getSiteId = require('./getSiteId');
var getDriveId = require('./getDriveId');
var getDownloadUrl = require('./getDownloadUrl');
var getUserId = require('./getUserId');
var downloadDriveItem = require('./downloadDriveItem');
var settings = require('../Settings/settings');

module.exports = function getTemplate(context, token, jsonTemplate,
    displayName, description, owner) {
    
    context.log('Running getTemplate.js');
    context.log('Getting template ' + jsonTemplate);

    return new Promise((resolve, reject) => {

        var template;   // The template as a JavaScript object

        // 1. Get ID of the SharePoint site where template files are stored
        context.log('Debug - getTemplate.js: Connecting to  template URL ' + settings().TEMPLATE_SITE_URL);
        getSiteId(context, token, settings().TENANT_NAME, 
            settings().TEMPLATE_SITE_URL)
        .then((siteId) => {
        // 2. Get the Graph API drive ID for the doc library where template files are stored
        context.log('Debug - getTemplate.js: Trying to get Drive ID');
             return getDriveId(context, token, siteId, settings().TEMPLATE_LIB_NAME);
        })
        .then((driveId) => {
        // 3. Get the download URL for the template file
        context.log('Debug - getTemplate.js: Drive ID ' + driveId);
        context.log('Debug - getTemplate.js: Trying to download the file');
            return getDownloadUrl(context, token, driveId,
                `${jsonTemplate}${settings().TEMPLATE_FILE_EXTENSION}`);
        })
        .then((downloadUrl) => {
        // 4. Get the contents of the template file
        context.log('Debug - getTemplate.js: Get content of the file');
            return downloadDriveItem(context, token, downloadUrl);
        })
        .then((templateString) => {

        // 5. Parse the template; get owner's user ID
        context.log('Debug - getTemplate.js: Trying to parse it');
        context.log(templateString);
            template = JSON.parse(templateString.trimLeft());
        context.log('Debug - getTemplate.js: Parsed OK - Get owner user ID');
            return getUserId (context, token, owner);
        context.log('Debug - getTemplate.js: Got user ID');
        })
        .then((ownerId) => {
        // 6. Add the per-team properties to the template
        context.log('Debug - getTemplate.js: Trying to add the per-team properties to the template');

            template['displayName'] = displayName;
            template['description'] = description;
            template['owners@odata.bind'] = [
                `https://graph.microsoft.com/beta/users('${ownerId}')`
            ];

        // 7. Return the finished template as a string
        context.log('Debug - getTemplate.js: Return the finished template as a string');
            resolve(JSON.stringify(template));
        })
        .catch((ex) => {
            reject(`Error in getTemplate(): ${ex}`);
        });


    });
}

var getSiteId = require('./getSiteId');
var getDriveId = require('./getDriveId');
var getDownloadUrl = require('./getDownloadUrl');
var getUserId = require('./getUserId');
var downloadDriveItem = require('./downloadDriveItem');
var pictureTemplate = "GDPRIcon.png"
var picbase64
var settings = require('../Settings/settings');

module.exports = function getPicture(context, token, pictureTemplate) {

    context.log('Getting picture ' + pictureTemplate);

    return new Promise((resolve, reject) => {

        var picTemplate;   // The picture as a JavaScript object

        // 1. Get ID of the SharePoint site where template files are stored
        getSiteId(context, token, settings().TENANT_NAME, 
            settings().TEMPLATE_SITE_URL)
        .then((siteId) => {
        // 2. Get the Graph API drive ID for the doc library where template files are stored
             return getDriveId(context, token, siteId, settings().TEMPLATE_LIB_NAME);
        })
        .then((driveId) => {
        // 3. Get the download URL for the template file
            return getDownloadUrl(context, token, driveId,
                `${pictureTemplate}`);
        })
        .then((downloadUrl) => {
        // 4. Get the contents of the picture
            return downloadDriveItem(context, token, downloadUrl);
            // 5. Convert to base64 https://www.mavention.nl/blogs-cat/microsoft-graph-api-how-to-change-images/?cn-reloaded=1
         //  picbase64 =  getBase64String(downloadDriveItem).then(base64Image => {
           //     const groupId = 'The group ID here';
             //   const request = {
               //   method: 'PUT',
                 // url:  'https://graph.microsoft.com/v1.0//groups/' + groupID + '/photo/$value',
                  //responseType: 'application/json',
                  // data: base64Image
       

        // 7. Return the finished template as a string
            resolve(JSON.stringify(template));
        })
        .catch((ex) => {
            reject(`Error in getPicture(): ${ex}`);
        });


    });
}
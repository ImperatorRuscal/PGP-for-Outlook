



import * as openpgp from './Scripts/openpgp.min.mjs';
import WKD from './Scripts/wkd.js';
//import * from './Scripts/UtilityFunctions.js';
const { stripHtml, defaults, version, Attribute, CbObj, Opts, Res, Tag } = stringStripHtml;
import { smalltalk } from './Scripts/smalltalk.min.js';

'use strict';


const privateKeyArmored = `
-----BEGIN PGP PRIVATE KEY BLOCK-----

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
rhn9
=3C0S
-----END PGP PRIVATE KEY BLOCK-----
`;


var privateKey = await openpgp.readPrivateKey({ armoredKey: privateKeyArmored });
var item = null;

var recipientList = null;

async function getBodyContentCallback(AsyncResult) {
    
    const messageBody = AsyncResult.value;
    //console.log('Original message body was :: ' + messageBody)
    console.log('Now attempting to encrypt');
    const encrypted = await encryptMessage(messageBody);
    console.log('Encryption result is as follows ::--: ' + encrypted);
    $('#item-body-encrypted-msg').html(encrypted);
    console.log('Now attempting to decrypt');
    const decrypted = await decryptMessage(encrypted);
    console.log('The decryption reads as follows ::--:: ' + decrypted);
    $('#item-body-decrypted-msg').html(decrypted);
}

async function getRecipientsCallback(AsyncResults) {
    console.log('I\'ve been given a list of recipients, let\'s populate the global');
    recipientList = AsyncResults.value;
    console.log(recipientList);
}

function waitForRecipients() {
    if (recipientList==null) {
        setTimeout(waitForRecipients, 500);
    }
}

async function encryptMessage(messageText) {
    var encryptTo = [];
    encryptTo.push(privateKey.toPublic());
    //encryptTo.push(await fetchPublicPGPKeyByEmailUsingVKS('rrockett@acmewidget.com'));
    encryptTo.push(await fetchPublicPGPKeyByEmailUsingWKD('rrockett@acmewidget.com'));
    if (!privateKey.isDecrypted()) {
        console.log('Decrypting the private key for local use');
        privateKey = await openpgp.decryptKey({ privateKey: privateKey, passphrase: (await smalltalk.prompt('PGP Passphrase', 'Please enter your passphrase to continue.', '', {type: 'password'})) })
        Office.context.roamingSettings.set('PGP_Private_Key_Encrypted', privateKey); // As long as we don't call save, we can keep the key in local available in local session and still protected in roaming storage'
    }
    console.log('Cleaning the message');  
    messageText = stripHtml(messageText, {
        stripTogetherWithTheirContents: ["embed","iframe","img","script","xml"],
        onlyStripTags: ["embed", "iframe", "img", "script", "xml"]
    });
    console.log(messageText);
    messageText = messageText.result;
    console.log('Encrypting the message');
    const encrypted = await openpgp.encrypt({
        message: await openpgp.createMessage({ text: messageText }),
        encryptionKeys: encryptTo,
        signingKeys: privateKey
    });
    console.log('Message is encrypted');
    return encrypted;
}

async function decryptMessage(cypherText) {
    console.log('Converting ASCII message to binary');
    const message = await openpgp.readMessage({ armoredMessage: cypherText });
    if (!privateKey.isDecrypted()) {
        console.log('Decrypting the private key for local use');
        privateKey = await openpgp.decryptKey({ privateKey: privateKey, passphrase: (await smalltalk.prompt('PGP Passphrase', 'Please enter your passphrase to continue.', '', { type: 'password' })) })
        Office.context.roamingSettings.set('PGP_Private_Key_Encrypted', privateKey); // As long as we don't call save, we can keep the key in local available in local session and still protected in roaming storage'
    }
    console.log('Decrypting the message');
    var { data: decrypted, signatures } = await openpgp.decrypt({
        message,
        decryptionKeys: privateKey
    });
    console.log('Message is decrypted');
    return decrypted;
}

async function fetchPublicPGPKeyByEmailUsingVKS(emailAddress, keyserver = 'keys.openpgp.org') {
    console.log('I\'m setting up to download the public key from a Hagrid VKS');
    const url = 'https://' + keyserver + '/vks/v1/by-email/' + encodeURIComponent(emailAddress);
    console.log('I\'ll be downloading it from:  ' + url);
    const response = await fetch(url);
    console.log('Download complete');
    if (response.status != 200) {
        console.log('No key found -- uh oh');
        return null;
    }
    console.log('Public key downloaded');
    const armoredKey = await response.text();
    console.log('Found a public key in the directory ::--::  ' + armoredKey);
    console.log('Setting up the binary public key');
    const key = await openpgp.readKey({ armoredKey: armoredKey });
    return key;
}

async function fetchPublicPGPKeyByEmailUsingWKD(emailAddress) {
    console.log('Setting up to fetch the key for ' + emailAddress + ' using WKD.');
    const wkd = new WKD();
    console.log('Performing the WKD lookup.');
    const publicKeyBytes = await wkd.lookup({ email: emailAddress });
    console.log('Public key found.  Converting to local key object.');
    const publicKey = await openpgp.readKey({ binaryKey: publicKeyBytes });
    console.log(publicKey);
    return publicKey;
}


(function () {
    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            console.log('The document is ready.');
            item = Office.context.mailbox.item;

            console.log('Setup with personal PGP keys');
            Office.context.roamingSettings.set('PGP_Private_Key_Encrypted', privateKey);
            Office.context.roamingSettings.saveAsync((saveSettings) => { console.log('PGP Private Key value saved.'); });
            if (!privateKey.isPrivate) {
                console.error('The listed PRIVATE KEY is not actually a private key!!!');
            }
            privateKey = null;
            privateKey = Office.context.roamingSettings.get('PGP_Private_Key_Encrypted');
            console.log('PGP private key fetched');
            console.log(privateKey);

            // Write message property values to the task pane
            $('#item-id').text(item.itemId);
            $('#item-subject').text(item.subject);
            $('#item-internetMessageId').text(item.internetMessageId);
            $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");

            //console.log('Getting message recipients');
            //item.recipients.getAsync(getRecipientsCallback);
            //waitForRecipients;

            console.log('Getting message body');
            item.body.getAsync(Office.CoercionType.Html, getBodyContentCallback);
        });
    });   
})();

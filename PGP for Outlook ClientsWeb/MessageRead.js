



import * as openpgp from './Scripts/openpgp.min.mjs';
import WKD from './Scripts/wkd.js';
//import * from './Scripts/UtilityFunctions.js';
const { stripHtml, defaults, version, Attribute, CbObj, Opts, Res, Tag } = stringStripHtml;
import { smalltalk } from './Scripts/smalltalk.min.js';

'use strict';


const privateKeyArmored = `
-----BEGIN PGP PRIVATE KEY BLOCK-----
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
7CgmmyMjt2iUz1kw1bCjhZ6tBW/8yI8E9w==
=fKB3
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

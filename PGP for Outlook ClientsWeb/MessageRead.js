



import * as openpgp from './Scripts/openpgp.min.mjs';
import WKD from './Scripts/wkd.js';
//import * from './Scripts/UtilityFunctions.js';
const { stripHtml, defaults, version, Attribute, CbObj, Opts, Res, Tag } = stringStripHtml;
import { smalltalk } from './Scripts/smalltalk.min.js';

'use strict';


const privateKeyArmored = `
-----BEGIN PGP PRIVATE KEY BLOCK-----

lQPGBFiA/vUBCADEOWzzvoqeY133MMgYaz9tTLipqki5iztKCQcjqoYupjoGEbRg
gZmxO9b2YStOwh4WUuavnvbCiJxaagUBe9ObTSVau3ZS4UI6ajybrNtxDO3gatGO
zurd5k9KijYrk9csK6tk8vdZcIAdoh81NUytanI4MMB0FQSUs7yoZqCcLTxYJAOz
SNdx2LrcFHxSF/rI3cJGHLIVQqVGXLhzRxBbm4Z2XikW//wXaTuxNb6zgficNnlC
KDTvX7tthlUHydYSJ058WFaVSoT0roQADF6TkFsqYJHqj0nwdLYocuaSFyu1LoK4
/lyPiYUifxBhniBlLa930ct17iSwUQczzO7LABEBAAH+BwMC5CnZ9TOJFAvTVyDx
qkX36d4H+ToeW9yKD1sarwIBlNCjJJaAKhyEshzLCDLGwZnIGiQAZcsx2MrHa0gX
nM1WmSsvuh7QSOUEcx2f0Z/1KqPloR91YZrpO8Mh5dtnUqDRoQ1SYlrqx+pBu9wT
0SjyUH/wtSl6vzRNb2rKm+3Ee+KvQhPyd+bwOzZuqa+Iy+UQLI/eIIfXb8B0Itmz
L5hMFlu9OXRHIR/ikk291VY0Xc+UHSk6WeVODjiIDrKQmWTTODKvimDr+/mgGvvm
mIAGwSuEx6GAIr6Wpe3RmjN58L3VqXxLn3eJuJHVPykldHD43qNx4TeZ7XoqFdKU
JiU5VCmgh5KdkJr5b0c8mvpqNa5I3+zuqqc7AmwpMdabcGgYmsWsobnxy3HwcMOc
srzm1fuKEKMyUrAmbchtNB5DY10Rg1J5Z4koZM+1Q/AZPXdKQDp5JQfUSxqeli/q
jkS0qjP53eF+8dReNb0TG1ri6nv/wjtGfDC75ZEgHFZI3RkGzKTbgdzIwC/CnJik
vXb6B0zf20IvFf/ovbzKhIQPCTiL3rBBFtuogYdQ4gEI8tnzYRCBqEkBkAvtlxIs
L37aSMZcBoJPas4/9VWU3nb+x4/WV31iPHqGxe4/WRkgj70CBlRZ/CZ0D4TZJhEm
UO7zdSVaNfddGmqaOL4iQyH1BVhd4DRQ/SUHcqIM0nS7BhULrzycK13JUtGZTSeL
BHOCYpnt/SBsnzjEXoiJlJ09/5lGi3+xMR053pGF5POq/2Tu+SxyjpBvxriRYcGn
DA9esRA09Z6zvcDd9lmTOgp1D4nXZxdkmCuoGfYA99eXMjhyts8MoVLuXgm1vFps
mtGcyJO72DHoMn7LNF9cx098dhzcezYfJ4paxI3F+PS1dKuG+wE/GtfQM3m9rkwj
SNdTqe8VYsWNtEdSZXRlbnRpb24gRmlsZSBEZWNyeXB0ZXIgPHJldGVudGlvbl9k
ZWNyeXB0ZXJAc2VydmljZXMuYWNtZW5ldHdvcmsuY29tPokBOQQTAQgAIwUCWID+
9QIbAwcLCQgHAwIBBhUIAgkKCwQWAgMBAh4BAheAAAoJELPs3Er+bVfM3FQH/i/e
B8fs6J/WOmBEkFroTQB50yAqcHWtzaAjOwXjBy8IBQnx2U5BAzhZ+7amUhaRl62y
KcRMSSjyLIyLiRqPLqfu9htxJ4CXUkfIhP9H89evCfA8++tB37clAE/hJSgl4WZW
1js9KQrvFPpVHJcrRYWgP2SYkASB2DRrQRc+8a3BYZQTHDyM30UrCa1ZCrabi/Y8
ciSMFM85idnghbZ98lT/B5dXQQ/8fHeuEWxINaCLFT4++++BYVY+QH1YdF5J2gcJ
HmAEeFgmjaN5Qw2VdYzJhbBqMxs/LivQ3uaH5oPxtD4+/6DX+tjWCRko7+T2sl20
bkiIkMfzxMZlMtTrI9aJATMEEAEIAB0WIQRR/qOxwswmotkYJphnQvVPJ7GawgUC
Y+ptwQAKCRBnQvVPJ7Gawj7UB/9f0+hXOC2Y+CB1bm4MaA22ZPPo4id9sCdWOEaT
nc+3XZOEslprGAHBzHP3nQRjzdJlffIKrb46d65L2bTj0ChPLgq3lY9gzA3HtZy0
xPXqIyJxgBofhFY6vzn5QcEDLh1GAl/V3jy5Tu/0CDP35SMijb9YeDpU/HLiLs/b
rECXr86b/5ZVqnmhg9pZAwf9RAvchPGM/BgwDbtnVnN9DcCfG7mdu+p7ETwTYY13
r2fZtoa+hZbEBIZ9qbIRDjO4JDEiaxO8JUrOV1Af268cMXceIntf+nyYOpZXYJO6
QS+AmESA+5XeVsO+5Kg0lrV3xIrR7QSjz99Hx918q6RvxwaBtERDb21wbGlhbmNl
IFByb2Nlc3NpbmcgRW5kcG9pbnQgPENvbXBsaWFuY2VAc2VydmljZXMuYWNtZW5l
dHdvcmsuY29tPokBUQQTAQgAOxYhBNb5pr6mjRwBuytpJLPs3Er+bVfMBQJj7sNE
AhsDBQsJCAcCAiICBhUKCQgLAgQWAgMBAh4HAheAAAoJELPs3Er+bVfMpWkH/R8I
d55wNXEIIwhq7/eSBAghByozif5G4QXfixhB+EFymGiibkg/qh8Cp5AD9Zk9MW4s
hQSFVf9NndCQ5efdhAQDRBVh18z8WMRV7yXPxgzvPUTjYScyW8MKJtNengcMYA95
jwdeWWe89aXb8X0RXx/rliTj/+43nynXuYDbE1bCBSsnxo70uWIyJ8x6h8Hmjw9n
0ROmSoXGpbCVo+kwKaJ9zaG6YlupVxEcfSJI1h/CphgHl7CxohgnKEHpb2jvqJP+
1WM57cr5Z3stslST4jckXRIemNZ0FhX6uzjvetl/24cMx5pPeGHgXqjAH5r8dOj4
fsIQFyXC8kGdmlWYSYCdA8YEWID+9QEIAKanFEfGbHpv3fhqPxD8WIZU7wCmV8N8
XQX0ixhqcY466L8HukqBnpDIun6StbjEsw6nwateUZK7ppt7equI6SNJEjMXbh1N
WyxAhbjTywQUAl5sDmrXEBG7Z977XythXOG6eOABdkk7D6TLaa+OzXMCKbbEe1Wj
3kD/k++asicIjYhqP/mvkaTRdtoPRKmWZXmYqiZfJqqAKRWhTleF7BoZ3MY03HGD
QgQJ+3JqRR2Qv7h4lg9OrCMe+jnYxIUwUphWJXyFO2aXEbIGtdHrJ7GNqUXAKYeT
8yKvynX6lquDusXQolNL6ixvUmAuXWdpJtOcRW0IOSyZJU9G6nIJuO8AEQEAAf4H
AwL3Ujwr36l2ONOakKcpH6HqVTQbQjyqX6QmVszyoDYTKPl0rxsBX71we/iJ4QC8
jUXyAwrkC8YEdXacnpZFYJ0F1MqQjiVuHF9vCZBUuYNrUZx1acmFhqbz8JmK0baX
sJI61p2Lk1GYZrDBLie5nR0fFE/icy0X6+MhtVlW3cGVU9BusQ2+Kma5mOrH7S0T
Tl1bifDv5bhdqF3UFrsfRaSdHyywNHeKlTfwfWeZCKsNkD2DsmvqiCmg4+HQ7v0G
Z0N+bz8G1SuN8h9A7u+Mppz/7zwF54Ze3TFprkhCxc23VLK92rqZ/mHCDwQ7/lAI
c+ET9Uc+skjAJ9JXGCsFd2L2g0x618/PMDECWMrdHcBjypA6OPX9smj66gXs2fxX
AV+wBnNDsjt4mgE63prLAGcBJvAxlZVCA4UGmEwbDXh15SO3/08itdMclgmTYB4T
rQI1TicKk+k4x5ZYyG/wUWkpaVuTKkbAPqkub8oBi3w0ZvWoufEcIDwNp0jYL1g8
UFBtxzefL/Kg/Hsx9Dp/I/aFIZ3nv+OYjtAMJFVpC3sRltfAHLP4qcg+LqBnLdpL
GwVS3zuD9fAMFRzUlZod8sM77KdlsPI2xrzLPcQheChBG6fiYZI0aNL6u+2KEP+e
IYp87O7RvaaGGJ6kTsPWq5NLV6A2niMaTV8FrAiQa3TstExpeOdt9lDvUdspyTCy
GDMSVd5OSfMx2Dd4koonfiXHIYDxX01SNc+sRsiJAgx124ioeCE4mUzLNfelYWUI
Cdr8U/efojLn0omsIEdi1ivSObetsvYn4JAhEHXV9mEacZbBmgQ2/1lzEWp63nWr
yTC1YVzkmLHZCjn4LwEWehdIyHqWvs/tEGLF2wusGOixCS0IdNfy2EFtcyR7Xrgi
JH4UMavGHUCoJODw7spvXvwAMrxMJ9eJAR8EGAEIAAkFAliA/vUCGwwACgkQs+zc
Sv5tV8zM4AgAgphuInWAW+loYKI5LdlbMyuDtf2rVYg9srA6mA95KEW/giYvkQZ1
rCNtU5vWdo1oV40gOAk7sGx1bezCJNvOWeNCmWPA+bjMojoHaLdbVKh8Z2tvFB2Y
tTS3v6ZMNDTQbmkqV2KLsbMXQCk2vnWwPdnlK3jFaq65SG4AAVnuyQYenHSStrPU
y3yJs648vErifcM5vVIlkP62JM7b/sSjZuIU9k3xIeLVK5aoQrS+rU6ol4C2vyw5
YxT2MXQtQar5w/rb5PmFWFpkA9HlMzt/cdJXRirE74IrGAGnnz1vr1Wpw1JXZ13Z
yC6k8ohePhJ2QX+ubJkKN+53NWXKdEGzDg==
=lT6V
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
    //encryptTo.push(await fetchPublicPGPKeyByEmailUsingVKS('emailaddress'));
    //encryptTo.push(await fetchPublicPGPKeyByEmailUsingWKD('emailaddress'));
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
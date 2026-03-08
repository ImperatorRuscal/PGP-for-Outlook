/**
 * pgp-core.js
 * Core PGP cryptographic operations built on OpenPGP.js v5.
 */

import * as openpgp from '../openpgp.min.mjs';

// ── Key generation ────────────────────────────────────────────────────────────

/**
 * Generate a new ECC key pair (Ed25519 signing / X25519 encryption).
 * The private key is returned armored and encrypted with the given passphrase.
 *
 * @param {string} name       - User's full name
 * @param {string} email      - User's email address
 * @param {string} passphrase - Passphrase to protect the private key
 * @returns {{ privateKey: string, publicKey: string }} Armored key strings
 */
export async function generateKeyPair(name, email, passphrase) {
  const { privateKey, publicKey } = await openpgp.generateKey({
    type: 'ecc',
    curve: 'curve25519',
    userIDs: [{ name, email }],
    passphrase,
    format: 'armored',
  });
  return { privateKey, publicKey };
}

// ── Key reading / parsing ─────────────────────────────────────────────────────

/**
 * Parse an armored public key string into a key object.
 */
export async function readPublicKey(armoredKey) {
  return await openpgp.readKey({ armoredKey });
}

/**
 * Parse a binary public key (e.g. from WKD) into a key object.
 */
export async function readPublicKeyFromBinary(binaryKey) {
  return await openpgp.readKey({ binaryKey });
}

/**
 * Decrypt an armored private key using its passphrase.
 * Returns the decrypted private key object, ready for signing/decryption.
 * This does NOT persist anything — caller is responsible for scoping.
 *
 * @param {string} armoredPrivateKey - Armored, passphrase-encrypted private key
 * @param {string} passphrase
 * @returns {openpgp.PrivateKey}
 */
export async function unlockPrivateKey(armoredPrivateKey, passphrase) {
  const privateKey = await openpgp.readPrivateKey({ armoredKey: armoredPrivateKey });
  return await openpgp.decryptKey({ privateKey, passphrase });
}

// ── Key inspection ────────────────────────────────────────────────────────────

/**
 * Extract human-readable metadata from an armored key (public or private).
 *
 * @param {string} armoredKey
 * @returns {{ fingerprint, shortId, keyId, userIds, name, email, created, expires, isPrivate, algorithm }}
 */
export async function getKeyInfo(armoredKey) {
  const key = await openpgp.readKey({ armoredKey });
  const primaryUser = await key.getPrimaryUser();
  const expirationTime = await key.getExpirationTime();

  const fp = key.getFingerprint().toUpperCase();
  return {
    fingerprint: fp,
    // Spaced groups of 4 for readability: ABCD EFGH ...
    fingerprintFormatted: fp.match(/.{1,4}/g).join(' '),
    shortId: fp.slice(-8),
    keyId: key.getKeyID().toHex().toUpperCase(),
    userIds: key.getUserIDs(),
    name: primaryUser?.user?.userID?.name || '',
    email: primaryUser?.user?.userID?.email || '',
    created: key.getCreationTime(),
    expires: expirationTime === Infinity ? null : expirationTime,
    isPrivate: key.isPrivate(),
    algorithm: key.getAlgorithmInfo().algorithm,
  };
}

// ── Message encryption / decryption ──────────────────────────────────────────

/**
 * Encrypt a plaintext message to one or more recipient public key objects.
 * Optionally sign with the sender's unlocked private key.
 *
 * @param {string}            text               - Plaintext message body
 * @param {openpgp.Key[]}     recipientPublicKeys - Array of parsed public key objects
 * @param {openpgp.PrivateKey} [signingKey]       - Unlocked private key for signing
 * @returns {string} Armored PGP message
 */
export async function encryptMessage(text, recipientPublicKeys, signingKey = null) {
  const options = {
    message: await openpgp.createMessage({ text }),
    encryptionKeys: recipientPublicKeys,
  };
  if (signingKey) {
    options.signingKeys = signingKey;
  }
  return await openpgp.encrypt(options);
}

/**
 * Decrypt an armored PGP message.
 *
 * @param {string}            armoredMessage   - PGP-armored ciphertext
 * @param {openpgp.PrivateKey} decryptionKey   - Unlocked private key
 * @param {openpgp.Key[]}     [verificationKeys] - Public keys to check signatures against
 * @returns {{ data: string, signatureResult: { valid: boolean|null, signedByKeyId: string|null } }}
 */
export async function decryptMessage(armoredMessage, decryptionKey, verificationKeys = []) {
  const message = await openpgp.readMessage({ armoredMessage });
  const result = await openpgp.decrypt({
    message,
    decryptionKeys: decryptionKey,
    verificationKeys: verificationKeys.length > 0 ? verificationKeys : undefined,
    expectSigned: false,
  });

  let signatureResult = { valid: null, signedByKeyId: null };
  if (result.signatures && result.signatures.length > 0) {
    const sig = result.signatures[0];
    try {
      await sig.verified;
      signatureResult.valid = true;
      signatureResult.signedByKeyId = sig.keyID?.toHex()?.toUpperCase() || null;
    } catch {
      signatureResult.valid = false;
      signatureResult.signedByKeyId = sig.keyID?.toHex()?.toUpperCase() || null;
    }
  }

  return { data: result.data, signatureResult };
}

// ── Attachment encryption / decryption ───────────────────────────────────────

/**
 * Encrypt binary attachment data as an armored PGP message.
 * The original filename is embedded in the PGP literal data packet.
 *
 * @param {Uint8Array}        data               - Raw file bytes
 * @param {string}            filename           - Original filename
 * @param {openpgp.Key[]}     recipientPublicKeys
 * @param {openpgp.PrivateKey} [signingKey]
 * @returns {string} Armored PGP message (save as "<filename>.pgp")
 */
export async function encryptAttachment(data, filename, recipientPublicKeys, signingKey = null) {
  const options = {
    message: await openpgp.createMessage({ binary: data, filename }),
    encryptionKeys: recipientPublicKeys,
    format: 'armored',
  };
  if (signingKey) {
    options.signingKeys = signingKey;
  }
  return await openpgp.encrypt(options);
}

/**
 * Decrypt an armored PGP attachment message.
 *
 * @param {string}            armoredMessage
 * @param {openpgp.PrivateKey} decryptionKey
 * @returns {{ data: Uint8Array, filename: string }}
 */
export async function decryptAttachment(armoredMessage, decryptionKey) {
  const message = await openpgp.readMessage({ armoredMessage });
  const result = await openpgp.decrypt({
    message,
    decryptionKeys: decryptionKey,
    format: 'binary',
  });

  // The filename lives in the first literal data packet
  const filename = message.packets?.[0]?.filename || 'decrypted_file';
  return { data: result.data, filename };
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/**
 * Detect whether a string contains a PGP armored block.
 * Returns the type found, or null.
 *
 * @param {string} text
 * @returns {'encrypted'|'signed'|'public-key'|'private-key'|null}
 */
export function detectPgpContent(text) {
  if (!text) return null;
  if (text.includes('-----BEGIN PGP MESSAGE-----'))        return 'encrypted';
  if (text.includes('-----BEGIN PGP SIGNED MESSAGE-----')) return 'signed';
  if (text.includes('-----BEGIN PGP PUBLIC KEY BLOCK-----')) return 'public-key';
  if (text.includes('-----BEGIN PGP PRIVATE KEY BLOCK-----')) return 'private-key';
  return null;
}

/**
 * Convert a base64 string to Uint8Array (for handling Office.js attachment content).
 */
export function base64ToUint8Array(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/**
 * Convert a Uint8Array to a base64 string.
 */
export function uint8ArrayToBase64(bytes) {
  let binary = '';
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

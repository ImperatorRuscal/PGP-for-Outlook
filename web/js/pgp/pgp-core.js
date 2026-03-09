/**
 * pgp-core.js
 * Core PGP cryptographic operations built on OpenPGP.js v5.
 *
 * This module is the only place in the add-in that directly calls the
 * OpenPGP.js library.  All other modules go through these wrappers, which
 * keeps the crypto surface small and easy to audit.
 *
 * Supported key types (passed as the `keyType` argument to generateKeyPair):
 *
 *   'ecc'  — Ed25519 (signing) + X25519 (encryption), curve25519.  Default.
 *     Rationale:
 *       - Compact keys and signatures (~200 bytes vs ~512 bytes for RSA-2048)
 *       - Widely supported by modern OpenPGP clients (GnuPG ≥ 2.1, Thunderbird, etc.)
 *       - Deterministic signing — no random-number dependency during sign operations
 *       - Strongly recommended by the OpenPGP RFC 9580 "crypto-refresh" update
 *
 *   'rsa4096' — RSA-4096 (sign + encrypt, legacy).
 *     Rationale:
 *       - Maximum interoperability with older PGP clients (GnuPG < 2.1, PGP 2.x, etc.)
 *       - Larger keys (~6–10 KB stored) and slower generation (~5–15 s in-browser)
 *       - Choose this only when you need to exchange keys with clients that do not
 *         support modern ECC algorithms (pre-2015 software or constrained appliances)
 */

import * as openpgp from '../openpgp.min.mjs';

// ── Key generation ────────────────────────────────────────────────────────────

/**
 * Generate a new PGP key pair protected by a passphrase.
 *
 * The passphrase is applied using AES-256 symmetric encryption inside the
 * OpenPGP packet structure (S2K + CFB).  The private key material is never
 * accessible without the passphrase.
 *
 * @param {string} name              - User's full name (embedded in the key UID)
 * @param {string} email             - User's email address (embedded in the key UID)
 * @param {string} passphrase        - Passphrase to protect the private key
 * @param {'ecc'|'rsa4096'} [keyType='ecc']
 *   'ecc'     — Ed25519 / X25519 (modern, compact, fast).  Recommended default.
 *   'rsa4096' — RSA-4096 (legacy interoperability with older PGP clients such as
 *               GnuPG < 2.1).  Key generation takes several seconds in-browser
 *               and produces larger keys (~6–10 KB vs ~3–6 KB for ECC).
 * @returns {Promise<{ privateKey: string, publicKey: string }>} Armored key strings
 */
export async function generateKeyPair(name, email, passphrase, keyType = 'ecc') {
  let genOptions;

  if (keyType === 'rsa4096') {
    genOptions = {
      type: 'rsa',
      rsaBits: 4096,
      userIDs: [{ name, email }],
      passphrase,
      format: 'armored',
    };
  } else {
    // Default: ECC (Ed25519 + X25519 / curve25519)
    genOptions = {
      type: 'ecc',
      curve: 'curve25519',
      userIDs: [{ name, email }],
      passphrase,
      format: 'armored',
    };
  }

  const { privateKey, publicKey } = await openpgp.generateKey(genOptions);
  return { privateKey, publicKey };
}

// ── Key reading / parsing ─────────────────────────────────────────────────────

/**
 * Parse an armored public key string into a key object.
 * Throws if the armor is malformed or the key type is unexpected.
 */
export async function readPublicKey(armoredKey) {
  return await openpgp.readKey({ armoredKey });
}

/**
 * Parse a binary public key (e.g. the raw response from a WKD lookup)
 * into a key object.
 */
export async function readPublicKeyFromBinary(binaryKey) {
  return await openpgp.readKey({ binaryKey });
}

/**
 * Decrypt an armored private key using its passphrase.
 * Returns an in-memory unlocked key object ready for signing or decryption.
 *
 * SECURITY NOTE: The returned object contains the raw private key material in
 * memory for the duration of the operation.  It is NOT persisted anywhere by
 * this function; callers should discard it as soon as the crypto operation
 * is complete (i.e. do not store it in module-level state).
 *
 * @param {string} armoredPrivateKey - Armored, passphrase-encrypted private key
 * @param {string} passphrase        - The key's passphrase
 * @returns {Promise<openpgp.PrivateKey>} Unlocked private key object
 * @throws If the passphrase is wrong or the armored text is corrupt
 */
export async function unlockPrivateKey(armoredPrivateKey, passphrase) {
  const privateKey = await openpgp.readPrivateKey({ armoredKey: armoredPrivateKey });
  return await openpgp.decryptKey({ privateKey, passphrase });
}

// ── Key inspection ────────────────────────────────────────────────────────────

/**
 * Extract human-readable metadata from an armored key (public or private).
 * Safe to call with a private key — no sensitive data is returned.
 *
 * @param {string} armoredKey
 * @returns {{
 *   fingerprint: string,         // 40-char hex, uppercase
 *   fingerprintFormatted: string, // spaced groups of 4: "ABCD EFGH …"
 *   shortId: string,             // last 8 chars of fingerprint
 *   keyId: string,               // hex key ID, uppercase
 *   userIds: string[],
 *   name: string,
 *   email: string,
 *   created: Date,
 *   expires: Date|null,          // null = no expiration
 *   isPrivate: boolean,
 *   algorithm: string
 * }}
 */
export async function getKeyInfo(armoredKey) {
  const key = await openpgp.readKey({ armoredKey });
  const primaryUser = await key.getPrimaryUser();

  // getExpirationTime() returns the JS Date of the earliest binding-signature
  // expiry, or Infinity if no expiry is set on any valid signature.
  const expirationTime = await key.getExpirationTime();

  const fp = key.getFingerprint().toUpperCase();
  return {
    fingerprint: fp,
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
 * The message is encrypted once for every recipient key (OpenPGP PKESK
 * packets), meaning any one recipient can decrypt it independently.
 *
 * @param {string}             text               - Plaintext message body
 * @param {openpgp.Key[]}      recipientPublicKeys - Parsed public key objects
 * @param {openpgp.PrivateKey} [signingKey]        - Unlocked private key for signing (optional)
 * @returns {Promise<string>} Armored PGP message ("-----BEGIN PGP MESSAGE-----")
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
 * Signature verification is opportunistic: if verificationKeys are provided
 * and the message was signed, the result includes a validity flag.  If no
 * verification keys are provided (or the message is unsigned), signatureResult
 * will have valid === null (not false — that would imply a failed verification).
 *
 * @param {string}             armoredMessage    - PGP-armored ciphertext
 * @param {openpgp.PrivateKey} decryptionKey     - Unlocked private key
 * @param {openpgp.Key[]}      [verificationKeys] - Public keys for signature checking
 * @returns {Promise<{
 *   data: string,
 *   signatureResult: { valid: boolean|null, signedByKeyId: string|null }
 * }>}
 */
export async function decryptMessage(armoredMessage, decryptionKey, verificationKeys = []) {
  const message = await openpgp.readMessage({ armoredMessage });
  const result = await openpgp.decrypt({
    message,
    decryptionKeys: decryptionKey,
    verificationKeys: verificationKeys.length > 0 ? verificationKeys : undefined,
    // expectSigned: false means we don't throw if there's no signature.
    // This is correct behavior for encrypted-only (no signature) messages.
    expectSigned: false,
  });

  let signatureResult = { valid: null, signedByKeyId: null };
  if (result.signatures && result.signatures.length > 0) {
    const sig = result.signatures[0];
    try {
      // sig.verified is a Promise that resolves on valid and rejects on invalid.
      // We await it inside a try/catch to avoid unhandled rejections.
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
 *
 * The original filename is stored inside the PGP Literal Data packet so that
 * the recipient can restore the correct filename when they decrypt.
 * The encrypted output should be saved with a ".pgp" suffix (e.g. "report.pdf.pgp").
 *
 * @param {Uint8Array}         data               - Raw file bytes
 * @param {string}             filename           - Original filename (stored in PGP packet)
 * @param {openpgp.Key[]}      recipientPublicKeys
 * @param {openpgp.PrivateKey} [signingKey]       - Optional signing key
 * @returns {Promise<string>} Armored PGP message
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
 * @param {string}             armoredMessage
 * @param {openpgp.PrivateKey} decryptionKey
 * @returns {Promise<{ data: Uint8Array, filename: string }>}
 */
export async function decryptAttachment(armoredMessage, decryptionKey) {
  const message = await openpgp.readMessage({ armoredMessage });
  const result = await openpgp.decrypt({
    message,
    decryptionKeys: decryptionKey,
    // format: 'binary' returns a Uint8Array instead of a string, which is
    // required for arbitrary binary files (not just text).
    format: 'binary',
  });

  // The embedded filename lives in the first Literal Data packet.
  // Fall back to a generic name if it wasn't stored (shouldn't happen for
  // files encrypted by this add-in, but handles keys from third-party clients).
  const filename = message.packets?.[0]?.filename || 'decrypted_file';
  return { data: result.data, filename };
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/**
 * Detect whether a string contains a PGP armored block and return its type.
 * Used to decide which UI panel to show when reading a message.
 *
 * @param {string} text
 * @returns {'encrypted'|'signed'|'public-key'|'private-key'|null}
 */
export function detectPgpContent(text) {
  if (!text) return null;
  if (text.includes('-----BEGIN PGP MESSAGE-----'))          return 'encrypted';
  if (text.includes('-----BEGIN PGP SIGNED MESSAGE-----'))   return 'signed';
  if (text.includes('-----BEGIN PGP PUBLIC KEY BLOCK-----')) return 'public-key';
  if (text.includes('-----BEGIN PGP PRIVATE KEY BLOCK-----')) return 'private-key';
  return null;
}

/**
 * Convert a base64 string to Uint8Array.
 * Used to convert the base64 attachment content returned by Office.js
 * getAttachmentContentAsync() into raw bytes for encryption.
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
 * Used to convert encrypted bytes into the format expected by
 * Office.js addFileAttachmentFromBase64Async().
 */
export function uint8ArrayToBase64(bytes) {
  let binary = '';
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

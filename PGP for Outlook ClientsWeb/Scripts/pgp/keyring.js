/**
 * keyring.js
 * Manages the local keyring — the collection of trusted contacts' public keys
 * stored in Office roaming settings.
 */

import { getKeyring, saveKeyring, estimateStorageUsage, STORAGE_LIMIT_BYTES } from './key-storage.js';
import { getKeyInfo, readPublicKey } from './pgp-core.js';

// Warn when storage is above this fraction of the limit
const STORAGE_WARN_THRESHOLD = 0.8;

// ── CRUD operations ───────────────────────────────────────────────────────────

/**
 * Add or replace a contact's public key in the keyring.
 * Validates the armored key before saving.
 *
 * @param {string} email        - Contact's email address (used as the lookup key)
 * @param {string} armoredKey   - Armored PGP public key string
 * @returns {{ info: object, storageWarning: boolean }}
 */
export async function addContactKey(email, armoredKey) {
  // Validate the key parses correctly before persisting
  const info = await getKeyInfo(armoredKey);
  if (info.isPrivate) {
    throw new Error('Refusing to store a private key in the shared keyring.');
  }

  const keyring = getKeyring();
  keyring[email.toLowerCase()] = armoredKey;
  await saveKeyring(keyring);

  const usage = estimateStorageUsage();
  return {
    info,
    storageWarning: usage > STORAGE_LIMIT_BYTES * STORAGE_WARN_THRESHOLD,
    storageUsed: usage,
  };
}

/**
 * Retrieve a contact's armored public key from the keyring.
 * Returns null if not found.
 *
 * @param {string} email
 * @returns {string|null}
 */
export function getContactKey(email) {
  const keyring = getKeyring();
  return keyring[email.toLowerCase()] || null;
}

/**
 * Remove a contact's key from the keyring.
 *
 * @param {string} email
 */
export async function removeContactKey(email) {
  const keyring = getKeyring();
  delete keyring[email.toLowerCase()];
  await saveKeyring(keyring);
}

/**
 * Check whether a contact's key is in the keyring.
 *
 * @param {string} email
 * @returns {boolean}
 */
export function hasContactKey(email) {
  const keyring = getKeyring();
  return email.toLowerCase() in keyring;
}

// ── Listing ───────────────────────────────────────────────────────────────────

/**
 * List all contacts in the keyring with their parsed key metadata.
 * Keys that fail to parse are included with an `error` field.
 *
 * @returns {Array<{ email: string, armoredKey: string, info?: object, error?: string }>}
 */
export async function listContactKeys() {
  const keyring = getKeyring();
  const results = [];
  for (const [email, armoredKey] of Object.entries(keyring)) {
    try {
      const info = await getKeyInfo(armoredKey);
      results.push({ email, armoredKey, info });
    } catch (e) {
      results.push({ email, armoredKey, error: `Could not parse key: ${e.message}` });
    }
  }
  // Sort by email for consistent display
  results.sort((a, b) => a.email.localeCompare(b.email));
  return results;
}

// ── Key object retrieval ──────────────────────────────────────────────────────

/**
 * Get a parsed public key object for a contact.
 * Returns null if not found or if the stored key is unparseable.
 *
 * @param {string} email
 * @returns {openpgp.Key|null}
 */
export async function getContactKeyObject(email) {
  const armoredKey = getContactKey(email);
  if (!armoredKey) return null;
  try {
    return await readPublicKey(armoredKey);
  } catch {
    return null;
  }
}

// ── Storage diagnostics ───────────────────────────────────────────────────────

/**
 * Return an estimate of keyring-only storage usage and remaining capacity.
 */
export function getKeyringStorageInfo() {
  const keyring = getKeyring();
  const count = Object.keys(keyring).length;
  const keyringBytes = JSON.stringify(keyring).length;
  const totalBytes = estimateStorageUsage();
  const remainingBytes = STORAGE_LIMIT_BYTES - totalBytes;

  return {
    count,
    keyringBytes,
    totalBytes,
    remainingBytes,
    nearLimit: totalBytes > STORAGE_LIMIT_BYTES * STORAGE_WARN_THRESHOLD,
  };
}

/**
 * key-discovery.js
 * Resolves email addresses to PGP public keys using multiple sources:
 *   1. Local keyring (roaming settings)
 *   2. Web Key Directory (WKD)
 *   3. Verified Key Server (VKS) — keys.openpgp.org
 *
 * Import a key found via WKD/VKS into the local keyring using keyring.addContactKey()
 * after user confirmation.
 */

import WKD from '../wkd.js';
import { getContactKeyObject } from './keyring.js';
import { readPublicKey, readPublicKeyFromBinary } from './pgp-core.js';

// ── Status codes ──────────────────────────────────────────────────────────────

export const KeyStatus = Object.freeze({
  FOUND_LOCAL: 'found_local',   // Key in the local keyring
  FOUND_WKD:   'found_wkd',     // Key discovered via WKD
  FOUND_VKS:   'found_vks',     // Key discovered via keys.openpgp.org
  NOT_FOUND:   'not_found',     // No key found anywhere
  ERROR:       'error',         // Lookup failed with an exception
});

// ── Single-key discovery ──────────────────────────────────────────────────────

/**
 * Attempt to find a public key for an email address.
 * Sources tried in order: local keyring → WKD → VKS.
 *
 * @param {string} email
 * @returns {{ key: openpgp.Key|null, status: KeyStatus, source: string|null, armoredKey: string|null }}
 */
export async function discoverKey(email) {
  const normalised = email.trim().toLowerCase();

  // 1. Local keyring
  try {
    const localKey = await getContactKeyObject(normalised);
    if (localKey) {
      return {
        key: localKey,
        status: KeyStatus.FOUND_LOCAL,
        source: 'Local keyring',
        armoredKey: null, // already stored locally
      };
    }
  } catch (e) {
    console.warn('Keyring lookup error for', normalised, e);
  }

  // 2. WKD
  try {
    const result = await fetchFromWKD(normalised);
    if (result) {
      return { key: result.key, status: KeyStatus.FOUND_WKD, source: 'WKD', armoredKey: result.armoredKey };
    }
  } catch (e) {
    console.info('WKD lookup failed for', normalised, '—', e.message);
  }

  // 3. VKS (keys.openpgp.org)
  try {
    const result = await fetchFromVKS(normalised);
    if (result) {
      return { key: result.key, status: KeyStatus.FOUND_VKS, source: 'keys.openpgp.org', armoredKey: result.armoredKey };
    }
  } catch (e) {
    console.info('VKS lookup failed for', normalised, '—', e.message);
  }

  return { key: null, status: KeyStatus.NOT_FOUND, source: null, armoredKey: null };
}

// ── Individual source fetchers ────────────────────────────────────────────────

/**
 * Fetch a public key via WKD.
 * Returns null if no key found for this address.
 *
 * @param {string} email
 * @returns {{ key: openpgp.Key, armoredKey: string }|null}
 */
export async function fetchFromWKD(email) {
  const wkd = new WKD();
  const binaryKey = await wkd.lookup({ email });
  if (!binaryKey || binaryKey.length === 0) return null;
  const key = await readPublicKeyFromBinary(binaryKey);
  // Convert back to armored for storage
  const armoredKey = key.armor();
  return { key, armoredKey };
}

/**
 * Fetch a public key from keys.openpgp.org (VKS /vks/v1/by-email).
 * VKS only returns keys whose email addresses have been verified by the owner.
 *
 * @param {string} email
 * @param {string} [keyserver='keys.openpgp.org']
 * @returns {{ key: openpgp.Key, armoredKey: string }|null}
 */
export async function fetchFromVKS(email, keyserver = 'keys.openpgp.org') {
  const url = `https://${keyserver}/vks/v1/by-email/${encodeURIComponent(email)}`;
  const response = await fetch(url);
  if (response.status !== 200) return null;
  const armoredKey = await response.text();
  if (!armoredKey || !armoredKey.includes('BEGIN PGP')) return null;
  const key = await readPublicKey(armoredKey);
  return { key, armoredKey };
}

// ── Bulk resolution ───────────────────────────────────────────────────────────

/**
 * Resolve an array of email addresses to their best-available public keys.
 * Runs all lookups concurrently.
 *
 * @param {string[]} emails
 * @returns {Array<{ email: string, key: openpgp.Key|null, status: KeyStatus, source: string|null, armoredKey: string|null }>}
 */
export async function resolveRecipients(emails) {
  return Promise.all(
    emails.map(async (email) => {
      try {
        const result = await discoverKey(email);
        return { email, ...result };
      } catch (e) {
        return { email, key: null, status: KeyStatus.ERROR, source: null, armoredKey: null, error: e.message };
      }
    })
  );
}

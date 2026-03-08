/**
 * key-storage.js
 * Wrapper around Office.context.roamingSettings for persisting PGP keys.
 *
 * Storage keys:
 *   pgp_private_key  - Armored, passphrase-encrypted private key string
 *   pgp_public_key   - Armored public key string
 *   pgp_key_meta     - Object: { name, email, fingerprint, keyId, created, expires }
 *   pgp_keyring      - Object: { "email@example.com": "-----BEGIN PGP PUBLIC KEY BLOCK-----..." }
 *   pgp_org_override - Object: manual org config override
 *
 * Note: Office roaming settings have a ~32KB total limit.
 * A passphrase-encrypted private key is typically 3-6KB armored.
 * Each contact public key is typically 1-3KB armored.
 * Monitor estimateStorageUsage() to avoid hitting the ceiling.
 */

const KEYS = {
  PRIVATE:      'pgp_private_key',
  PUBLIC:       'pgp_public_key',
  META:         'pgp_key_meta',
  KEYRING:      'pgp_keyring',
  ORG_OVERRIDE: 'pgp_org_override',
};

function settings() {
  return Office.context.roamingSettings;
}

function saveAsync() {
  return new Promise((resolve, reject) => {
    settings().saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
      } else {
        resolve();
      }
    });
  });
}

// ── Own key pair ──────────────────────────────────────────────────────────────

export function hasKeyPair() {
  return !!settings().get(KEYS.PRIVATE);
}

export function getPrivateKey() {
  return settings().get(KEYS.PRIVATE) || null;
}

export function getPublicKey() {
  return settings().get(KEYS.PUBLIC) || null;
}

export function getKeyMetadata() {
  return settings().get(KEYS.META) || null;
}

export async function saveKeyPair(armoredPrivateKey, armoredPublicKey, metadata) {
  settings().set(KEYS.PRIVATE, armoredPrivateKey);
  settings().set(KEYS.PUBLIC, armoredPublicKey);
  settings().set(KEYS.META, metadata);
  await saveAsync();
}

export async function clearKeyPair() {
  settings().remove(KEYS.PRIVATE);
  settings().remove(KEYS.PUBLIC);
  settings().remove(KEYS.META);
  await saveAsync();
}

// ── Keyring (contacts' public keys) ──────────────────────────────────────────

export function getKeyring() {
  return settings().get(KEYS.KEYRING) || {};
}

export async function saveKeyring(keyring) {
  settings().set(KEYS.KEYRING, keyring);
  await saveAsync();
}

// ── Org config override ───────────────────────────────────────────────────────

export function getOrgOverride() {
  return settings().get(KEYS.ORG_OVERRIDE) || null;
}

export async function saveOrgOverride(config) {
  settings().set(KEYS.ORG_OVERRIDE, config);
  await saveAsync();
}

export async function clearOrgOverride() {
  settings().remove(KEYS.ORG_OVERRIDE);
  await saveAsync();
}

// ── Storage diagnostics ───────────────────────────────────────────────────────

/**
 * Estimate current roaming settings storage usage in bytes.
 * Roaming settings are serialized as JSON; this gives an approximation.
 */
export function estimateStorageUsage() {
  const data = {
    [KEYS.PRIVATE]:      settings().get(KEYS.PRIVATE) || '',
    [KEYS.PUBLIC]:       settings().get(KEYS.PUBLIC) || '',
    [KEYS.META]:         settings().get(KEYS.META) || {},
    [KEYS.KEYRING]:      settings().get(KEYS.KEYRING) || {},
    [KEYS.ORG_OVERRIDE]: settings().get(KEYS.ORG_OVERRIDE) || {},
  };
  return JSON.stringify(data).length;
}

export const STORAGE_LIMIT_BYTES = 32768; // 32KB Office roaming settings ceiling

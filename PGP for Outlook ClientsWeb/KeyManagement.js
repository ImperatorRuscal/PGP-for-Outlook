'use strict';
/**
 * KeyManagement.js
 * Task pane for managing the user's PGP identity and keyring.
 *
 * Sections:
 *
 *  MY KEY PAIR
 *    Generate a new Ed25519/X25519 key pair protected by a passphrase.
 *    The passphrase-encrypted private key and the public key are stored in
 *    Office roaming settings — they follow the user across devices via their
 *    Microsoft 365 account, but never leave the Office ecosystem unencrypted.
 *    The user can copy or email their public key to contacts, and can delete
 *    or regenerate the key pair at any time.
 *
 *  CONTACTS' KEYRING
 *    A local store of trusted contacts' public keys, also in roaming settings.
 *    Keys can be added by searching (WKD → VKS auto-discovery) or by pasting
 *    armored text directly.  A storage-usage warning appears when the 32 KB
 *    roaming settings ceiling is within 20% of being reached.
 *
 *  ORGANISATION SETTINGS
 *    The add-in tries to load org-level config from:
 *      https://<user-email-domain>/.well-known/pgp-outlook-config.json
 *    IT admins can enable/configure the company key feature by publishing that
 *    file.  A manual override (stored in roaming settings) takes precedence
 *    and is intended for orgs that cannot host a well-known file.
 */

import { generateKeyPair, getKeyInfo } from './Scripts/pgp/pgp-core.js';
import {
  hasKeyPair, getPrivateKey, getPublicKey, getKeyMetadata,
  saveKeyPair, clearKeyPair,
  getOrgOverride, saveOrgOverride, clearOrgOverride,
} from './Scripts/pgp/key-storage.js';
import {
  addContactKey, removeContactKey, listContactKeys, getKeyringStorageInfo,
} from './Scripts/pgp/keyring.js';
import { discoverKey, KeyStatus } from './Scripts/pgp/key-discovery.js';
import {
  loadOrgConfig, getOrgConfig, isCompanyKeyEnabled, isCompanyKeyRequired,
  getCompanyKeyEmails, fetchCompanyKeys,
} from './Scripts/pgp/org-config.js';

// ── Helpers ───────────────────────────────────────────────────────────────────

function el(id) { return document.getElementById(id); }

function showStatus(containerId, message, type = 'info') {
  const container = el(containerId);
  container.className = `pgp-alert pgp-alert--${type}`;
  container.textContent = message;
  container.classList.remove('pgp-hidden');
}

function hideStatus(containerId) {
  el(containerId).classList.add('pgp-hidden');
}

function formatDate(date) {
  if (!date) return 'Never';
  return date.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
}

// ── My key pair panel ─────────────────────────────────────────────────────────

async function refreshMyKeyPanel() {
  if (!hasKeyPair()) {
    el('panel-no-key').classList.remove('pgp-hidden');
    el('panel-has-key').classList.add('pgp-hidden');
    return;
  }

  el('panel-no-key').classList.add('pgp-hidden');
  el('panel-has-key').classList.remove('pgp-hidden');

  const meta = getKeyMetadata();
  if (meta) {
    el('key-uid').textContent = `${meta.name} <${meta.email}>`;
    el('key-created').textContent = `Created: ${formatDate(meta.created ? new Date(meta.created) : null)}`;
    el('key-expires').textContent = meta.expires
      ? `Expires: ${formatDate(new Date(meta.expires))}`
      : 'No expiration';
    el('key-fingerprint').textContent = meta.fingerprintFormatted || meta.fingerprint;

    // Warn if expired
    if (meta.expires && new Date(meta.expires) < new Date()) {
      el('key-status-badge').className = 'pgp-badge pgp-badge--error';
      el('key-status-badge').textContent = 'Expired';
    }
  }
}

function showGenerateForm() {
  el('panel-generate-form').classList.remove('pgp-hidden');
  hideStatus('gen-status');
  el('gen-name').focus();
}

function hideGenerateForm() {
  el('panel-generate-form').classList.add('pgp-hidden');
  el('gen-name').value = '';
  el('gen-email').value = '';
  el('gen-passphrase').value = '';
  el('gen-passphrase-confirm').value = '';
  hideStatus('gen-status');
}

async function handleGenerate() {
  const name       = el('gen-name').value.trim();
  const email      = el('gen-email').value.trim();
  const passphrase = el('gen-passphrase').value;
  const confirm    = el('gen-passphrase-confirm').value;

  if (!name)       return showStatus('gen-status', 'Full name is required.', 'error');
  if (!email)      return showStatus('gen-status', 'Email address is required.', 'error');
  if (!passphrase) return showStatus('gen-status', 'A passphrase is required.', 'error');
  if (passphrase !== confirm) return showStatus('gen-status', 'Passphrases do not match.', 'error');
  if (passphrase.length < 8) return showStatus('gen-status', 'Passphrase must be at least 8 characters.', 'warning');

  const btn = el('btn-generate-confirm');
  const spinner = el('gen-spinner');
  btn.disabled = true;
  spinner.classList.remove('pgp-hidden');
  hideStatus('gen-status');

  try {
    const { privateKey: armoredPrivate, publicKey: armoredPublic } = await generateKeyPair(name, email, passphrase);
    const info = await getKeyInfo(armoredPublic);

    await saveKeyPair(armoredPrivate, armoredPublic, {
      name:                info.name,
      email:               info.email,
      fingerprint:         info.fingerprint,
      fingerprintFormatted: info.fingerprintFormatted,
      keyId:               info.keyId,
      created:             info.created?.toISOString(),
      expires:             info.expires?.toISOString() ?? null,
    });

    hideGenerateForm();
    await refreshMyKeyPanel();
    showStatus('status-bar', 'Key pair generated and saved successfully.', 'success');
  } catch (e) {
    showStatus('gen-status', `Key generation failed: ${e.message}`, 'error');
  } finally {
    btn.disabled = false;
    spinner.classList.add('pgp-hidden');
  }
}

async function handleCopyPublicKey() {
  const armoredKey = getPublicKey();
  if (!armoredKey) return;
  try {
    await navigator.clipboard.writeText(armoredKey);
    showStatus('status-bar', 'Public key copied to clipboard.', 'success');
  } catch {
    // Fallback — show it in a prompt
    window.prompt('Copy the public key below:', armoredKey);
  }
}

function handleSendPublicKey() {
  const armoredKey = getPublicKey();
  const meta = getKeyMetadata();
  if (!armoredKey || !meta) return;

  // displayNewMessageForm opens a new compose window pre-filled with the
  // public key.  The recipient is left blank so the user can address it.
  // The armored key is wrapped in <pre> to preserve its line structure in
  // the HTML body; the recipient can copy it out and import it in any
  // OpenPGP-compatible client.
  Office.context.mailbox.displayNewMessageForm({
    toRecipients: [],
    subject: `PGP Public Key for ${meta.name}`,
    htmlBody:
      `<p>Hi,</p>` +
      `<p>Please find my PGP public key below. ` +
      `You can use this to send me encrypted messages.</p>` +
      `<p><strong>Fingerprint:</strong> ${meta.fingerprintFormatted || meta.fingerprint}</p>` +
      `<pre>${armoredKey}</pre>`,
  });
}

/**
 * Download the user's passphrase-encrypted private key as a .asc file.
 *
 * The exported file is the same armored text stored in roaming settings —
 * it is already encrypted with the user's passphrase (AES-256 via OpenPGP
 * S2K), so it is safe to store on disk, in cloud storage, etc.  It cannot
 * be decrypted or used without the passphrase.
 *
 * The suggested filename includes the key's short fingerprint ID so the user
 * can tell multiple backups apart, e.g. "pgp-private-key-ABCD1234.asc".
 */
function handleExportPrivateKey() {
  const armoredPrivate = getPrivateKey();
  if (!armoredPrivate) return;

  const meta     = getKeyMetadata();
  const shortId  = meta?.fingerprint?.slice(-8) || 'key';
  const filename = `pgp-private-key-${shortId}.asc`;

  const blob = new Blob([armoredPrivate], { type: 'text/plain;charset=utf-8' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);

  showStatus('status-bar',
    `Private key backup downloaded as "${filename}". ` +
    'Keep this file secure — it is protected by your passphrase.',
    'warning'
  );
}

async function handleDeleteKey() {
  if (!confirm('Are you sure you want to delete your key pair? This cannot be undone. You will need to generate a new key pair and share your new public key with all your contacts.')) return;
  await clearKeyPair();
  await refreshMyKeyPanel();
  showStatus('status-bar', 'Key pair deleted.', 'warning');
}

// ── Keyring panel ─────────────────────────────────────────────────────────────

async function refreshKeyringPanel() {
  const list = el('keyring-list');
  const empty = el('keyring-empty');
  const countBadge = el('keyring-count');
  const storageWarning = el('storage-warning');

  const contacts = await listContactKeys();
  countBadge.textContent = `${contacts.length} key${contacts.length !== 1 ? 's' : ''}`;

  // Remove all existing items except the empty placeholder
  Array.from(list.querySelectorAll('.pgp-key-item-wrapper')).forEach(el => el.remove());

  if (contacts.length === 0) {
    empty.classList.remove('pgp-hidden');
  } else {
    empty.classList.add('pgp-hidden');
    contacts.forEach(contact => {
      const li = document.createElement('li');
      li.className = 'pgp-key-item pgp-key-item-wrapper';
      li.dataset.email = contact.email;

      if (contact.error) {
        li.innerHTML = `
          <div class="pgp-key-item__header">
            <div class="pgp-key-item__identity">
              <div class="pgp-key-item__email">${escHtml(contact.email)}</div>
              <div class="pgp-key-item__meta" style="color:#a80000;">${escHtml(contact.error)}</div>
            </div>
            <button class="pgp-btn pgp-btn--danger pgp-btn--sm btn-remove-key" data-email="${escHtml(contact.email)}">Remove</button>
          </div>`;
      } else {
        const info = contact.info;
        li.innerHTML = `
          <div class="pgp-key-item__header">
            <div class="pgp-key-item__identity">
              <div class="pgp-key-item__email">${escHtml(contact.email)}</div>
              ${info.name ? `<div class="pgp-key-item__name">${escHtml(info.name)}</div>` : ''}
              <div class="pgp-key-item__meta">
                Algorithm: ${escHtml(info.algorithm)} &nbsp;·&nbsp;
                ${info.expires ? `Expires: ${formatDate(info.expires)}` : 'No expiration'}
              </div>
            </div>
            <button class="pgp-btn pgp-btn--danger pgp-btn--sm btn-remove-key" data-email="${escHtml(contact.email)}">Remove</button>
          </div>
          <span class="pgp-fingerprint">${escHtml(info.fingerprintFormatted)}</span>`;
      }
      list.appendChild(li);
    });
  }

  // Storage warning
  const storageInfo = getKeyringStorageInfo();
  storageWarning.classList.toggle('pgp-hidden', !storageInfo.nearLimit);
}

async function handleFindKey() {
  const email = el('keyring-search').value.trim();
  if (!email) return;

  const container = el('find-result');
  container.innerHTML = `<div class="pgp-alert pgp-alert--info"><span class="pgp-spinner"></span> Looking up key for ${escHtml(email)}…</div>`;
  container.classList.remove('pgp-hidden');

  try {
    const result = await discoverKey(email);

    if (result.status === KeyStatus.NOT_FOUND) {
      container.innerHTML = `<div class="pgp-alert pgp-alert--warning">No key found for <strong>${escHtml(email)}</strong> via WKD or keyserver. You can import one manually.</div>`;
      return;
    }

    const info = await getKeyInfo(result.key.armor());
    const sourceLabel = result.source;

    let html = `
      <div class="pgp-alert pgp-alert--success">
        <div>
          <strong>Key found</strong> via ${escHtml(sourceLabel)}<br/>
          ${info.name ? `${escHtml(info.name)}<br/>` : ''}
          <span class="pgp-fingerprint" style="margin-top:4px;">${escHtml(info.fingerprintFormatted)}</span>
        </div>
      </div>`;

    if (result.status !== KeyStatus.FOUND_LOCAL && result.armoredKey) {
      html += `<button class="pgp-btn pgp-btn--primary pgp-btn--sm pgp-mt-sm" id="btn-save-found-key" data-email="${escHtml(email)}" data-key="${escHtml(result.armoredKey)}">Save to Keyring</button>`;
    } else {
      html += `<div class="pgp-badge pgp-badge--success pgp-mt-sm">Already in keyring</div>`;
    }

    container.innerHTML = html;

    el('btn-save-found-key')?.addEventListener('click', async (e) => {
      const btn = e.currentTarget;
      try {
        await addContactKey(btn.dataset.email, btn.dataset.key);
        await refreshKeyringPanel();
        container.innerHTML = `<div class="pgp-alert pgp-alert--success">Key for ${escHtml(email)} saved to keyring.</div>`;
      } catch (err) {
        container.innerHTML = `<div class="pgp-alert pgp-alert--error">Could not save key: ${escHtml(err.message)}</div>`;
      }
    });

  } catch (e) {
    container.innerHTML = `<div class="pgp-alert pgp-alert--error">Lookup failed: ${escHtml(e.message)}</div>`;
  }
}

async function handleImportKey() {
  const email      = el('import-email').value.trim();
  const armoredKey = el('import-key-text').value.trim();

  if (!email)      return showStatus('import-status', 'Email address is required.', 'error');
  if (!armoredKey) return showStatus('import-status', 'Paste an armored public key.', 'error');

  try {
    const { info, storageWarning } = await addContactKey(email, armoredKey);
    el('import-email').value = '';
    el('import-key-text').value = '';
    el('panel-import-form').classList.add('pgp-hidden');
    await refreshKeyringPanel();
    showStatus('status-bar',
      `Key for ${email} saved. Fingerprint: ${info.fingerprintFormatted}`,
      storageWarning ? 'warning' : 'success'
    );
  } catch (e) {
    showStatus('import-status', `Could not save key: ${e.message}`, 'error');
  }
}

// ── Org settings panel ────────────────────────────────────────────────────────

function refreshOrgPanel() {
  const override = getOrgOverride();
  const config = getOrgConfig();

  if (override) {
    showStatus('org-status',
      'Using manual override. Auto-discovery from domain is skipped.',
      'warning'
    );
  } else if (isCompanyKeyEnabled()) {
    showStatus('org-status',
      `Org config loaded from domain. Company key: ${getCompanyKeyEmails().join(', ')}`,
      'info'
    );
  } else {
    showStatus('org-status',
      'No org config found. Company key feature is disabled.',
      'neutral'
    );
  }

  el('org-key-enabled').checked  = isCompanyKeyEnabled();
  el('org-key-required').checked = isCompanyKeyRequired();
  el('org-key-emails').value     = getCompanyKeyEmails().join(', ');
}

async function handleSaveOrgOverride() {
  const enabled  = el('org-key-enabled').checked;
  const required = el('org-key-required').checked;
  const emailsRaw = el('org-key-emails').value;
  const emails = emailsRaw.split(',').map(e => e.trim()).filter(Boolean);

  await saveOrgOverride({ companyKeyEnabled: enabled, companyKeyRequired: required, companyKeyEmails: emails });
  showStatus('org-save-status', 'Override saved.', 'success');
  refreshOrgPanel();
}

async function handleClearOrgOverride() {
  await clearOrgOverride();
  showStatus('org-save-status', 'Override cleared. Auto-discovery will be used.', 'info');
  refreshOrgPanel();
}

// ── XSS-safe HTML escaping ────────────────────────────────────────────────────

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── Bootstrap ─────────────────────────────────────────────────────────────────

Office.onReady(async () => {
  const userEmail = Office.context.mailbox.userProfile?.emailAddress || '';

  // Load org config from domain or override
  await loadOrgConfig(userEmail);

  // Wire generate-form buttons (there are two "show-generate" buttons)
  document.querySelectorAll('#btn-show-generate').forEach(btn => {
    btn.addEventListener('click', showGenerateForm);
  });
  el('btn-generate-cancel').addEventListener('click', hideGenerateForm);
  el('btn-generate-confirm').addEventListener('click', handleGenerate);

  el('btn-copy-pubkey')?.addEventListener('click', handleCopyPublicKey);
  el('btn-send-pubkey')?.addEventListener('click', handleSendPublicKey);
  el('btn-export-key')?.addEventListener('click', handleExportPrivateKey);
  el('btn-delete-key')?.addEventListener('click', handleDeleteKey);

  // Keyring
  el('btn-find-key').addEventListener('click', handleFindKey);
  el('keyring-search').addEventListener('keydown', e => { if (e.key === 'Enter') handleFindKey(); });

  el('btn-show-import').addEventListener('click', () => {
    el('panel-import-form').classList.toggle('pgp-hidden');
  });
  el('btn-import-cancel').addEventListener('click', () => {
    el('panel-import-form').classList.add('pgp-hidden');
  });
  el('btn-import-confirm').addEventListener('click', handleImportKey);

  // Keyring list — delegate remove button clicks
  el('keyring-list').addEventListener('click', async (e) => {
    const btn = e.target.closest('.btn-remove-key');
    if (!btn) return;
    const email = btn.dataset.email;
    if (confirm(`Remove key for ${email}?`)) {
      await removeContactKey(email);
      await refreshKeyringPanel();
      showStatus('status-bar', `Key for ${email} removed.`, 'info');
    }
  });

  // Org settings
  el('btn-save-org').addEventListener('click', handleSaveOrgOverride);
  el('btn-clear-org').addEventListener('click', handleClearOrgOverride);

  // Initial render
  await refreshMyKeyPanel();
  await refreshKeyringPanel();
  refreshOrgPanel();
});

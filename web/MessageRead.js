'use strict';
/**
 * MessageRead.js
 * Task pane for reading PGP-encrypted or PGP-signed messages.
 *
 * Capabilities:
 *  - Detect PGP content in the message body (encrypted / signed)
 *  - Decrypt the body using the user's private key
 *  - Verify signatures against keys in the local keyring or via WKD/VKS
 *  - List .pgp attachments and allow individual decryption + download
 */

import * as openpgp from './js/openpgp.min.mjs';
import {
  unlockPrivateKey,
  decryptMessage, decryptAttachment,
  detectPgpContent,
} from './js/pgp/pgp-core.js';
import { hasKeyPair, getPrivateKey } from './js/pgp/key-storage.js';
import { getContactKeyObject } from './js/pgp/keyring.js';
import { discoverKey, KeyStatus } from './js/pgp/key-discovery.js';
import {
  cacheSessionKey, getSessionKey, clearSessionKey,
  getSessionEmail, getSessionShortId, onSessionCleared,
} from './js/pgp/session-cache.js';

// ── Helpers ───────────────────────────────────────────────────────────────────

function el(id) { return document.getElementById(id); }

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function showSection(id) { el(id).classList.remove('pgp-hidden'); }
function hideSection(id) { el(id).classList.add('pgp-hidden'); }

function showStatus(message, type = 'info') {
  const bar = el('status-bar');
  bar.className = `pgp-alert pgp-alert--${type}`;
  bar.innerHTML = message;
  bar.classList.remove('pgp-hidden');
}

// ── Session status ────────────────────────────────────────────────────────────

function updateSessionStatus() {
  const bar   = el('session-status');
  const label = el('session-status-text');
  const email   = getSessionEmail();
  const shortId = getSessionShortId();
  if (email) {
    label.textContent = `Key unlocked: ${email}${shortId ? ' · …' + shortId : ''}`;
    bar.classList.remove('pgp-hidden');
  } else {
    bar.classList.add('pgp-hidden');
  }
}

// ── Passphrase modal ──────────────────────────────────────────────────────────

function promptPassphrase(message = 'Enter your passphrase to decrypt.') {
  return new Promise((resolve, reject) => {
    const modal = el('passphrase-modal');
    const input = el('passphrase-input');
    const errEl = el('passphrase-error');
    const msgEl = el('passphrase-modal-msg');

    msgEl.textContent = message;
    input.value = '';
    errEl.classList.add('pgp-hidden');
    modal.style.display = 'flex';
    modal.classList.remove('pgp-hidden');
    input.focus();

    function cleanup() {
      modal.style.display = '';
      modal.classList.add('pgp-hidden');
      el('btn-passphrase-ok').removeEventListener('click', onOk);
      el('btn-passphrase-cancel').removeEventListener('click', onCancel);
      input.removeEventListener('keydown', onKeydown);
    }

    function onOk() {
      const val = input.value;
      if (!val) {
        errEl.textContent = 'Passphrase is required.';
        errEl.classList.remove('pgp-hidden');
        return;
      }
      cleanup();
      resolve(val);
    }

    function onCancel() { cleanup(); reject(new Error('Cancelled.')); }
    function onKeydown(e) {
      if (e.key === 'Enter')  onOk();
      if (e.key === 'Escape') onCancel();
    }

    el('btn-passphrase-ok').addEventListener('click', onOk);
    el('btn-passphrase-cancel').addEventListener('click', onCancel);
    input.addEventListener('keydown', onKeydown);
  });
}

// ── Message body detection ────────────────────────────────────────────────────

/**
 * Extract the plain text from an HTML string by rendering it into a temporary
 * element and reading innerText.  This decodes HTML entities and inserts line
 * breaks at block-level elements, faithfully recreating what a reader would see.
 * Used as a fallback when the plain-text body from Office.js has been mangled
 * by Outlook's HTML round-trip (missing blank lines, entity-encoded dashes, etc.).
 */
function extractTextFromHtml(html) {
  // Normalize block-level closing tags to newlines BEFORE assigning innerHTML.
  // Outlook often wraps PGP armor lines in <p> elements; without this step,
  // innerText sees a single run of text with no line breaks between the armor
  // lines, producing a one-line string that looks nothing like valid PGP armor.
  const normalized = html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n');
  const div = document.createElement('div');
  div.innerHTML = normalized;
  // innerText respects CSS display, inserting newlines at block boundaries.
  // This preserves the blank-line separator that PGP armor requires.
  return div.innerText ?? div.textContent ?? '';
}

async function detectAndRenderBody() {
  // Primary: read as plain text.
  let body = await getBodyAsync(Office.CoercionType.Text);
  let pgpType = detectPgpContent(body);

  // Fallback: if no PGP markers found in the text body, try the HTML body.
  // When Outlook sends an HTML-format email the plain-text rendition can lose
  // blank lines, entity-encode dashes, or otherwise mangle the PGP armor so
  // that the -----BEGIN marker is no longer a literal substring.
  if (!pgpType) {
    try {
      const htmlBody = await getBodyAsync(Office.CoercionType.Html);
      const textFromHtml = extractTextFromHtml(htmlBody);
      const pgpTypeFromHtml = detectPgpContent(textFromHtml);
      if (pgpTypeFromHtml) {
        body = textFromHtml;
        pgpType = pgpTypeFromHtml;
      }
    } catch {
      // HTML body unavailable — stick with what we have.
    }
  }

  el('detection-loading').classList.add('pgp-hidden');

  const result = el('detection-result');
  result.classList.remove('pgp-hidden');

  if (!pgpType) {
    result.innerHTML = `<div class="pgp-alert pgp-alert--info">
      This message does not appear to contain PGP content.
    </div>`;
    renderPgpAttachments(); // still look for .pgp attachments
    return;
  }

  if (pgpType === 'encrypted') {
    result.innerHTML = `<div class="pgp-alert pgp-alert--info">
      <strong>Encrypted message</strong> — PGP-encrypted content detected.
    </div>`;
    showSection('section-decrypt');
    el('btn-decrypt').addEventListener('click', () => handleDecryptBody(body));
  }

  if (pgpType === 'signed') {
    result.innerHTML = `<div class="pgp-alert pgp-alert--info">
      <strong>Signed message</strong> — PGP-signed content detected.
    </div>`;
    showSection('section-signed-only');
    await handleVerifySignedBody(body);
  }

  if (pgpType === 'public-key') {
    result.innerHTML = `<div class="pgp-alert pgp-alert--warning">
      This message contains a <strong>PGP public key</strong>.
      You can copy it and import it via <em>Manage Keys</em>.
    </div>`;
  }

  if (pgpType === 'private-key') {
    result.innerHTML = `<div class="pgp-alert pgp-alert--error">
      ⚠ This message contains what appears to be a <strong>private key</strong>.
      Do not share or import private keys.
    </div>`;
  }

  renderPgpAttachments();
}

// ── Decrypt body ──────────────────────────────────────────────────────────────

async function handleDecryptBody(encryptedBody) {
  const btn = el('btn-decrypt');
  const spinner = el('decrypt-spinner');
  btn.disabled = true;
  spinner.classList.remove('pgp-hidden');

  try {
    // Check the session cache before prompting — avoids repeated passphrase
    // entry when the user decrypts several messages or attachments in one session.
    let privateKey = getSessionKey();

    if (!privateKey) {
      const passphrase = await promptPassphrase('Enter your passphrase to decrypt this message.');
      privateKey = await unlockPrivateKey(getPrivateKey(), passphrase);

      // Cache for 15 minutes of inactivity.
      const userEmail = Office.context.mailbox.userProfile?.emailAddress || '';
      cacheSessionKey(privateKey, userEmail, '');
      updateSessionStatus();
    }

    // Attempt to get the sender's public key for signature verification
    const senderEmail = Office.context.mailbox.item.from?.emailAddress;
    const verificationKeys = await resolveVerificationKeys(senderEmail);

    const { data, signatureResult } = await decryptMessage(
      encryptedBody, privateKey, verificationKeys
    );

    renderDecryptedBody(data, signatureResult, senderEmail);
    hideSection('section-decrypt');

  } catch (e) {
    if (e.message === 'Cancelled.') {
      /* user cancelled — silently re-enable */
    } else if (e.message?.includes('Error decrypting') || e.message?.includes('Decryption error')) {
      showStatus('Decryption failed — wrong passphrase or key?', 'error');
    } else {
      showStatus(`Decryption failed: ${escHtml(e.message)}`, 'error');
    }
  } finally {
    btn.disabled = false;
    spinner.classList.add('pgp-hidden');
  }
}

function renderDecryptedBody(text, signatureResult, senderEmail) {
  showSection('section-decrypted');

  // Render signature badge
  const sigBadge = el('signature-badge');
  const sigDetails = el('signature-details');

  if (signatureResult.valid === null) {
    sigBadge.innerHTML = `<span class="pgp-badge pgp-badge--neutral">No signature</span>`;
  } else if (signatureResult.valid) {
    sigBadge.innerHTML = `<span class="pgp-badge pgp-badge--success">✓ Valid signature</span>`;
    if (senderEmail) {
      sigDetails.textContent = `Signed by ${senderEmail} · Key ID: ${signatureResult.signedByKeyId || 'unknown'}`;
      sigDetails.classList.remove('pgp-hidden');
    }
  } else {
    sigBadge.innerHTML = `<span class="pgp-badge pgp-badge--error">✗ Invalid signature</span>`;
    sigDetails.textContent = `Signature could not be verified. The message may have been tampered with.`;
    sigDetails.classList.remove('pgp-hidden');
  }

  // Detect whether the decrypted payload is HTML (messages encrypted by this
  // add-in always are, since we encrypt CoercionType.Html).  Fall back to
  // plain-text rendering for messages encrypted by third-party clients.
  const trimmed = text.trimStart();
  const isHtml  = trimmed.startsWith('<!DOCTYPE') ||
                  trimmed.startsWith('<html')     ||
                  trimmed.startsWith('<HTML');

  if (isHtml) {
    // Render in a sandboxed iframe.  'allow-same-origin' lets the iframe read
    // its srcdoc content but scripts are NOT allowed (no 'allow-scripts').
    // This prevents any JavaScript inside the decrypted HTML from running.
    const frame = el('decrypted-html-frame');
    frame.srcdoc = text;
    el('decrypted-html-wrapper').classList.remove('pgp-hidden');

    // Resize iframe to fit content once it loads
    frame.addEventListener('load', () => {
      try {
        frame.style.height = frame.contentDocument.body.scrollHeight + 32 + 'px';
      } catch { /* cross-origin guard — shouldn't fire with srcdoc + allow-same-origin */ }
    }, { once: true });
  } else {
    el('decrypted-body').textContent = text;
    el('decrypted-body').classList.remove('pgp-hidden');
  }

  el('btn-copy-decrypted').addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(text);
      showStatus('Decrypted content copied to clipboard.', 'success');
    } catch {
      window.prompt('Copy the decrypted content:', text);
    }
  });
}

// ── Verify signed-only body ───────────────────────────────────────────────────

async function handleVerifySignedBody(signedBody) {
  const statusEl = el('signed-body-status');
  const bodyEl = el('signed-body');

  // Extract the text between "-----BEGIN PGP SIGNED MESSAGE-----" and the signature
  const textMatch = signedBody.match(
    /-----BEGIN PGP SIGNED MESSAGE-----[\s\S]*?\n\n([\s\S]*?)\n-----BEGIN PGP SIGNATURE-----/
  );
  const plainText = textMatch ? textMatch[1] : '';

  if (plainText) {
    bodyEl.textContent = plainText;
    bodyEl.classList.remove('pgp-hidden');
  }

  statusEl.innerHTML = `<div class="pgp-alert pgp-alert--info"><span class="pgp-spinner"></span> Verifying signature…</div>`;

  try {
    const senderEmail = Office.context.mailbox.item.from?.emailAddress;
    const verificationKeys = await resolveVerificationKeys(senderEmail);

    if (verificationKeys.length === 0) {
      statusEl.innerHTML = `<div class="pgp-alert pgp-alert--warning">
        Cannot verify signature — no public key found for <strong>${escHtml(senderEmail || 'sender')}</strong>.
        Import their key via Manage Keys to verify future messages.
      </div>`;
      return;
    }

    // Use openpgp.js directly for clearsigned message verification
    const cleartextMessage = await openpgp.readCleartextMessage({ cleartextMessage: signedBody });
    const verifyResult = await openpgp.verify({ message: cleartextMessage, verificationKeys });
    const sig = verifyResult.signatures[0];

    try {
      await sig.verified;
      statusEl.innerHTML = `<div class="pgp-alert pgp-alert--success">
        ✓ Valid signature from <strong>${escHtml(senderEmail || 'sender')}</strong>
      </div>`;
    } catch {
      statusEl.innerHTML = `<div class="pgp-alert pgp-alert--error">
        ✗ Invalid signature — this message may have been modified after signing.
      </div>`;
    }
  } catch (e) {
    statusEl.innerHTML = `<div class="pgp-alert pgp-alert--warning">Signature verification failed: ${escHtml(e.message)}</div>`;
  }
}

// ── Resolve sender's verification key ────────────────────────────────────────

async function resolveVerificationKeys(senderEmail) {
  if (!senderEmail) return [];
  try {
    // Try local keyring first (fast), then skip network lookup for read pane UX
    const localKey = await getContactKeyObject(senderEmail);
    if (localKey) return [localKey];

    // Try WKD/VKS silently
    const result = await discoverKey(senderEmail);
    return result.key ? [result.key] : [];
  } catch {
    return [];
  }
}

// ── Encrypted attachments ─────────────────────────────────────────────────────

function renderPgpAttachments() {
  const item = Office.context.mailbox.item;
  const attachments = item.attachments || [];
  const pgpAttachments = attachments.filter(a => !a.isInline && a.name.endsWith('.pgp'));

  if (pgpAttachments.length === 0) return;

  showSection('section-attachments');
  const list = el('attachment-list');
  list.innerHTML = '';

  pgpAttachments.forEach(att => {
    const li = document.createElement('li');
    li.className = 'pgp-attachment-item';
    li.innerHTML = `
      <span class="pgp-attachment-item__name" title="${escHtml(att.name)}">${escHtml(att.name)}</span>
      <button class="pgp-btn pgp-btn--secondary pgp-btn--sm btn-decrypt-att" data-id="${escHtml(att.id)}" data-name="${escHtml(att.name)}">
        Decrypt &amp; Download
      </button>`;
    list.appendChild(li);
  });

  list.addEventListener('click', async (e) => {
    const btn = e.target.closest('.btn-decrypt-att');
    if (!btn) return;

    btn.disabled = true;
    btn.textContent = '…';

    const attachmentId = btn.dataset.id;
    const attachmentName = btn.dataset.name;

    try {
      let privateKey = getSessionKey();
      if (!privateKey) {
        const passphrase = await promptPassphrase(`Enter your passphrase to decrypt ${attachmentName}.`);
        privateKey = await unlockPrivateKey(getPrivateKey(), passphrase);
        const userEmail = Office.context.mailbox.userProfile?.emailAddress || '';
        cacheSessionKey(privateKey, userEmail, '');
        updateSessionStatus();
      }

      const contentResult = await getAttachmentContentAsync(item, attachmentId);
      const armoredMessage = atob(contentResult.content);

      const { data: decryptedBytes, filename } = await decryptAttachment(armoredMessage, privateKey);

      // Trigger browser download
      downloadBytes(decryptedBytes, filename || attachmentName.replace(/\.pgp$/i, ''));
      showStatus(`"${filename || attachmentName}" decrypted and downloaded.`, 'success');

    } catch (e) {
      if (e.message !== 'Cancelled.') {
        showStatus(`Could not decrypt ${attachmentName}: ${escHtml(e.message)}`, 'error');
      }
    } finally {
      btn.disabled = false;
      btn.textContent = 'Decrypt & Download';
    }
  });
}

function getAttachmentContentAsync(item, attachmentId) {
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error.message));
    });
  });
}

function downloadBytes(bytes, filename) {
  const blob = new Blob([bytes]);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

// ── Office.js wrappers ────────────────────────────────────────────────────────

function getBodyAsync(coercionType) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(coercionType, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error.message));
    });
  });
}

// ── Bootstrap ─────────────────────────────────────────────────────────────────

Office.onReady(async () => {
  if (!hasKeyPair()) {
    el('panel-no-key').classList.remove('pgp-hidden');
    el('detection-loading').classList.add('pgp-hidden');
    el('detection-result').innerHTML = `<div class="pgp-alert pgp-alert--warning">Generate a key pair first to use decryption.</div>`;
    el('detection-result').classList.remove('pgp-hidden');
    return;
  }

  // Reflect current session cache state and keep it in sync
  updateSessionStatus();
  onSessionCleared(updateSessionStatus);

  el('btn-lock-session').addEventListener('click', () => {
    clearSessionKey(); // triggers onSessionCleared → updateSessionStatus
  });

  await detectAndRenderBody();
});

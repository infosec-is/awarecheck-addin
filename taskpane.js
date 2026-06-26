/**
 * taskpane.js
 * -----------
 * Runs inside the Outlook task pane.
 *
 * Flow:
 *  1. Office.onReady() — wait for Outlook to finish loading.
 *  2. User clicks "Check this email".
 *  3. We read From, Reply-To, Subject, and the first 500 chars of the body
 *     using the Office.js Mailbox API.
 *  4. POST that data to https://localhost:3000/analyse.
 *  5. Render the verdict.
 */

'use strict';

// ─── View helpers ──────────────────────────────────────────────────────────────

/** Show exactly one view, hide all others. */
function showView(id) {
  document.querySelectorAll('.view').forEach(el => el.classList.remove('active'));
  document.getElementById(id).classList.add('active');
}

// ─── Office initialisation ────────────────────────────────────────────────────

// Office.onReady fires once Outlook has fully loaded the task pane.
// We must not touch Office.context.mailbox before this.
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById('check-btn').addEventListener('click', runCheck);
    document.getElementById('again-btn').addEventListener('click', () => showView('idle-view'));
    document.getElementById('retry-btn').addEventListener('click', runCheck);

    // Remember the panel was open so it reopens automatically next time Outlook starts
    if (Office.addin && typeof Office.addin.setStartupBehavior === 'function') {
      Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    }
  }
});

// ─── Main check flow ──────────────────────────────────────────────────────────

async function runCheck() {
  showView('loading-view');

  try {
    const emailData = await readEmailData();
    const startMs   = Date.now();
    const verdict   = await postToServer(emailData);
    verdict.ms      = verdict.ms ?? (Date.now() - startMs); // use server time if available
    renderVerdict(verdict);
  } catch (err) {
    showError(err.message);
  }
}

// ─── Read email data from Outlook ─────────────────────────────────────────────

/**
 * Returns a promise that resolves to { from, replyTo, subject, body_snippet }.
 * We use Promise-based wrappers around Office.js callbacks so the code stays
 * clean and readable.
 */
function readEmailData() {
  const item = Office.context.mailbox.item;

  // subject is a plain string property — no callback needed
  const subject = item.subject || '';

  // sender address — include display name so spoofing checks work
  // Format: "Display Name <email@domain.com>" or just "email@domain.com"
  const from = item.from
    ? (item.from.displayName
        ? `${item.from.displayName} <${item.from.emailAddress}>`
        : (item.from.emailAddress || ''))
    : '';

  // reply-to (may be an array in some Outlook versions)
  let replyTo = '';
  if (item.replyTo) {
    if (Array.isArray(item.replyTo) && item.replyTo.length > 0) {
      replyTo = item.replyTo[0].emailAddress || '';
    } else if (typeof item.replyTo === 'object') {
      replyTo = item.replyTo.emailAddress || '';
    }
  }

  // Read body and internet headers in parallel, then combine
  const bodyPromise = new Promise((resolve, reject) => {
    item.body.getAsync(
      Office.CoercionType.Text,
      {},
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve((result.value || '').substring(0, 500));
        } else {
          resolve(''); // non-fatal
        }
      }
    );
  });

  // getAllInternetHeadersAsync requires Mailbox 1.6 — gracefully skip if unavailable
  const headersPromise = new Promise((resolve) => {
    if (typeof item.getAllInternetHeadersAsync !== 'function') {
      resolve('');
      return;
    }
    item.getAllInternetHeadersAsync({}, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || '');
      } else {
        resolve(''); // non-fatal
      }
    });
  });

  return Promise.all([bodyPromise, headersPromise]).then(([body_snippet, internet_headers]) => ({
    from, replyTo, subject, body_snippet, internet_headers
  }));
}

// ─── POST to the local server ─────────────────────────────────────────────────

async function postToServer(emailData) {
  let response;
  try {
    // HTTP on 3001 avoids certificate trust issues in Outlook's WebView2
    response = await fetch('http://localhost:3001/analyse', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify(emailData),
    });
  } catch (networkErr) {
    throw new Error(
      'Could not reach the AwareCHECK server.\n' +
      'Make sure you started the server with: node server.js\n\n' +
      'Technical detail: ' + networkErr.name + ' — ' + networkErr.message
    );
  }

  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`Server returned an error (${response.status}): ${text}`);
  }

  return response.json();
}

// ─── Render verdict ───────────────────────────────────────────────────────────

const ICONS = {
  fraud:      '✗',
  suspicious: '⚠',
  safe:       '✓',
};

function renderVerdict(verdict) {
  const { level, headline, explanation, findings, action, ms } = verdict;

  // Colour the header band
  const header = document.getElementById('verdict-header');
  header.className = `verdict-header ${level}`;
  document.getElementById('verdict-icon').textContent  = ICONS[level] ?? '?';
  document.getElementById('verdict-level').textContent = level.toUpperCase();

  // Apply level class to the result view so CSS can colour sub-elements
  const resultView = document.getElementById('result-view');
  resultView.className = `view ${level}`;

  // Text content
  document.getElementById('result-headline').textContent    = headline;
  document.getElementById('result-explanation').textContent = explanation;
  document.getElementById('result-action').textContent      = action;

  // Findings list
  const list = document.getElementById('findings-list');
  list.innerHTML = '';

  if (findings && findings.length > 0) {
    findings.forEach(({ icon, text }) => {
      const li = document.createElement('li');
      li.innerHTML =
        `<span class="fi-icon">${icon}</span>` +
        `<span>${escapeHtml(text)}</span>`;
      list.appendChild(li);
    });
  }

  // Privacy note at the bottom
  const timeLabel = ms != null ? ` Took ${ms} ms.` : '';
  document.getElementById('privacy-note').textContent =
    `Checked on your device. Nothing was sent externally.${timeLabel}`;

  showView('result-view');
}

// ─── Error display ────────────────────────────────────────────────────────────

function showError(message) {
  document.getElementById('error-message').textContent = message;
  showView('error-view');
}

// ─── Utility ──────────────────────────────────────────────────────────────────

/** Prevent accidental XSS from email content appearing in the UI */
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

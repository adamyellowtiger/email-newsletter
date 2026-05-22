const NEWSLETTER_VERSION = '1.0.0';

const SHEET_NAMES = {
  SUBSCRIBERS: 'Subscribers',
  SOURCES: 'Sources',
  CANDIDATES: 'Candidates',
  ISSUES: 'Issues',
  SEND_LOG: 'SendLog',
  CONFIG: 'Config'
};

const HEADERS = {};
HEADERS[SHEET_NAMES.SUBSCRIBERS] = [
  'email',
  'name',
  'status',
  'source',
  'consent_date',
  'confirmed',
  'unsubscribe_token',
  'last_sent_at'
];
HEADERS[SHEET_NAMES.SOURCES] = [
  'name',
  'url',
  'type',
  'category',
  'enabled',
  'weight'
];
HEADERS[SHEET_NAMES.CANDIDATES] = [
  'issue_date',
  'source_name',
  'category',
  'title',
  'url',
  'published_at',
  'summary',
  'why_it_matters',
  'score',
  'selected'
];
HEADERS[SHEET_NAMES.ISSUES] = [
  'issue_id',
  'issue_date',
  'status',
  'subject',
  'html_body',
  'text_body',
  'preview_sent_at',
  'approved_at',
  'sent_at',
  'sent_count',
  'error_count',
  'notes'
];
HEADERS[SHEET_NAMES.SEND_LOG] = [
  'timestamp',
  'issue_id',
  'batch',
  'recipient_count',
  'recipients',
  'status',
  'error'
];
HEADERS[SHEET_NAMES.CONFIG] = [
  'key',
  'value',
  'notes'
];

const DEFAULT_CONFIG = [
  ['newsletter_name', 'AI Daily', 'Public name shown in emails.'],
  ['admin_email', '', 'Preview and alert recipient. Defaults to the active Google account when available.'],
  ['reply_to', '', 'Reply-to address. Defaults to admin_email.'],
  ['from_name', 'AI Daily', 'Display name for GmailApp.sendEmail.'],
  ['timezone', 'America/Toronto', 'Used for issue dates and trigger expectations.'],
  ['lookback_days', '7', 'Only include items published within this many days.'],
  ['max_candidates', '60', 'Maximum candidates written to the Candidates tab.'],
  ['max_items_in_issue', '8', 'Maximum linked items included in the newsletter.'],
  ['unsubscribe_base_url', '', 'Apps Script web app URL used for unsubscribe links.'],
  ['alert_on_held_send', 'TRUE', 'Email admin when a scheduled send is held because the issue is not approved.'],
  ['footer_reason', 'You are receiving this because you joined the invite-only AI Daily beta.', 'Footer consent reminder.']
];

const DEFAULT_SOURCES = [
  ['OpenAI News', 'https://openai.com/news/rss.xml', 'rss', 'Product News', true, 10],
  ['Google AI', 'https://blog.google/technology/ai/rss/', 'rss', 'Product News', true, 9],
  ['Microsoft AI Blog', 'https://blogs.microsoft.com/ai/feed/', 'rss', 'Business', true, 8],
  ['Anthropic via Google News', 'https://news.google.com/rss/search?q=site%3Aanthropic.com%2Fnews%20Anthropic&hl=en-US&gl=US&ceid=US%3Aen', 'rss', 'Product News', true, 8],
  ['MIT Technology Review AI', 'https://www.technologyreview.com/topic/artificial-intelligence/feed/', 'rss', 'Big Picture', true, 8],
  ['The Verge AI', 'https://www.theverge.com/rss/ai-artificial-intelligence/index.xml', 'rss', 'Consumer AI', true, 8],
  ['VentureBeat AI', 'https://venturebeat.com/category/ai/feed', 'rss', 'Business', true, 7],
  ['Hugging Face Blog', 'https://huggingface.co/blog/feed.xml', 'rss', 'Tools', true, 6],
  ['Google DeepMind Blog', 'https://deepmind.google/blog/rss.xml', 'rss', 'Research, Translated', true, 4],
  ['Google Research', 'https://research.google/blog/rss/', 'rss', 'Research, Translated', false, 3],
  ['Meta Newsroom', 'https://about.fb.com/feed/', 'rss', 'Business', true, 5],
  ['Meta Engineering', 'https://engineering.fb.com/feed/', 'rss', 'Technical', false, 2],
  ['arXiv cs.AI', 'https://export.arxiv.org/rss/cs.AI', 'rss', 'Research, Translated', false, 1],
  ['arXiv cs.LG', 'https://export.arxiv.org/rss/cs.LG', 'rss', 'Research, Translated', false, 1]
];

const AI_FACTS = [
  'A chatbot can sound confident even when it is wrong, so important answers still need a source you can check.',
  'When a company says a model is multimodal, it usually means it can work with more than text, such as images, audio, or video.',
  'Open-weight AI means people can download model files, but the license still decides what they are allowed to do with them.',
  'Benchmarks are like test scores: useful clues, but not proof that a tool will be best for your exact task.',
  'Fast response time often matters as much as raw intelligence when AI is used in customer support, search, or voice products.',
  'Fine-tuning teaches a model a pattern or style; retrieval gives it fresh information to look up while answering.',
  'AI safety tests reduce risk, but they do not guarantee a model will behave perfectly in every real-world situation.',
  'Smaller AI models can be cheaper, faster, and good enough for focused jobs like sorting emails or summarizing notes.',
  'The most useful AI launches explain price, limits, who can use it, and what problem it actually solves.',
  'If an AI feature is free today, check whether there are usage limits, data-sharing terms, or paid plans later.'
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Newsletter')
    .addItem('1. Setup / repair sheets', 'setupNewsletter')
    .addItem('Use general-reader source mix', 'refreshGeneralReaderSources')
    .addSeparator()
    .addItem('Fetch sources + draft preview', 'fetchAndDraftIssue')
    .addItem('Email preview for latest draft', 'sendPreviewForLatestDraft')
    .addItem('Send approved issue now', 'sendApprovedIssue')
    .addSeparator()
    .addItem('Install daily triggers', 'installDailyTriggers')
    .addItem('Remove daily triggers', 'removeNewsletterTriggers')
    .addToUi();
}

function setupNewsletter() {
  const ss = getSpreadsheet_();
  Object.keys(SHEET_NAMES).forEach(function (key) {
    ensureSheet_(ss, SHEET_NAMES[key], HEADERS[SHEET_NAMES[key]]);
  });

  seedConfigDefaults_();
  seedSourceDefaults_();
  assignSubscriberDefaults_();
  SpreadsheetApp.flush();
}

function installDailyTriggers() {
  removeNewsletterTriggers();
  ScriptApp.newTrigger('runMorningDraft')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .nearMinute(0)
    .create();

  ScriptApp.newTrigger('sendApprovedIssue')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .nearMinute(30)
    .create();
}

function removeNewsletterTriggers() {
  const managedHandlers = {
    runMorningDraft: true,
    sendApprovedIssue: true
  };
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (managedHandlers[trigger.getHandlerFunction()]) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function runMorningDraft() {
  setupNewsletter();
  const issue = buildDailyDraft_({ force: false });
  if (issue) {
    sendPreview_(issue);
  }
}

function fetchAndDraftIssue() {
  setupNewsletter();
  const issue = buildDailyDraft_({ force: true });
  sendPreview_(issue);
}

function sendPreviewForLatestDraft() {
  setupNewsletter();
  const issue = getLatestIssue_();
  if (!issue) {
    throw new Error('No issue exists yet. Run "Fetch sources + draft preview" first.');
  }
  sendPreview_(issue);
}

function sendApprovedIssue() {
  setupNewsletter();
  assignSubscriberDefaults_();

  const issue = getLatestIssue_();
  if (!issue) {
    logSend_(new Date(), '', 'approval-check', 0, '', 'no_issue', 'No issue exists.');
    return;
  }

  if (String(issue.status).toLowerCase() === 'sent') {
    logSend_(new Date(), issue.issue_id, 'approval-check', 0, '', 'already_sent', 'Latest issue is already sent.');
    return;
  }

  if (String(issue.status).toLowerCase() !== 'approved') {
    logSend_(new Date(), issue.issue_id, 'approval-check', 0, '', 'held_not_approved', 'Issue status is "' + issue.status + '".');
    maybeAlertHeldSend_(issue);
    return;
  }

  const subscribers = getActiveSubscribers_();
  const alreadySent = getSentEmailSet_(issue.issue_id);
  const pendingSubscribers = subscribers.filter(function (subscriber) {
    return !alreadySent[String(subscriber.email).toLowerCase()];
  });

  if (pendingSubscribers.length === 0) {
    updateIssue_(issue.rowNumber, {
      status: 'sent',
      sent_at: new Date(),
      sent_count: subscribers.length,
      error_count: 0,
      notes: 'No pending subscribers. All active subscribers were already sent or no active subscribers exist.'
    });
    logSend_(new Date(), issue.issue_id, 'personalized-complete', 0, '', 'sent', 'No pending subscribers.');
    return;
  }

  const remainingQuota = MailApp.getRemainingDailyQuota();
  if (pendingSubscribers.length > remainingQuota) {
    const message = 'Needed ' + pendingSubscribers.length + ' recipient quota but only ' + remainingQuota + ' remains today.';
    updateIssue_(issue.rowNumber, {
      error_count: pendingSubscribers.length,
      notes: message
    });
    logSend_(new Date(), issue.issue_id, 'quota-check', pendingSubscribers.length, '', 'quota_blocked', message);
    sendAdminAlert_('AI Daily send blocked by Gmail quota', message);
    return;
  }

  if (!issue.approved_at) {
    updateIssue_(issue.rowNumber, { approved_at: new Date() });
  }

  let sentThisRun = 0;
  let errorsThisRun = 0;
  pendingSubscribers.forEach(function (subscriber, index) {
    const batchName = 'personalized-' + padNumber_(index + 1, 3);
    try {
      sendIssueToSubscriber_(issue, subscriber);
      sentThisRun += 1;
      setSubscriberLastSent_(subscriber.rowNumber, new Date());
      logSend_(new Date(), issue.issue_id, batchName, 1, subscriber.email, 'sent', '');
      Utilities.sleep(200);
    } catch (error) {
      errorsThisRun += 1;
      logSend_(new Date(), issue.issue_id, batchName, 1, subscriber.email, 'error', errorToString_(error));
    }
  });

  const sentAfterRun = Object.keys(getSentEmailSet_(issue.issue_id)).length;
  if (errorsThisRun === 0 && sentAfterRun >= subscribers.length) {
    updateIssue_(issue.rowNumber, {
      status: 'sent',
      sent_at: new Date(),
      sent_count: sentAfterRun,
      error_count: 0,
      notes: 'Sent successfully.'
    });
  } else {
    updateIssue_(issue.rowNumber, {
      sent_count: sentAfterRun,
      error_count: errorsThisRun,
      notes: 'Some sends failed or remain pending. Re-run "Send approved issue now" to retry unsent subscribers.'
    });
    sendAdminAlert_(
      'AI Daily send finished with errors',
      'Issue ' + issue.issue_id + ' sent to ' + sentAfterRun + ' subscriber(s), with ' + errorsThisRun + ' error(s). Check SendLog.'
    );
  }
}

function doGet(e) {
  setupNewsletter();
  const token = e && e.parameter ? e.parameter.token : '';
  if (token) {
    return unsubscribeByToken_(token);
  }
  return renderUnsubscribeForm_('');
}

function doPost(e) {
  setupNewsletter();
  const email = e && e.parameter ? normalizeEmail_(e.parameter.email) : '';
  if (!email) {
    return renderUnsubscribeForm_('Enter the email address you want to unsubscribe.');
  }
  const result = unsubscribeByEmail_(email);
  return HtmlService.createHtmlOutput(result.message);
}

function buildDailyDraft_(options) {
  const force = options && options.force === true;
  const timezone = getConfigValue_('timezone', 'America/Toronto');
  const today = formatDate_(new Date(), 'yyyy-MM-dd', timezone);
  const existingToday = getLatestIssueForDate_(today);

  if (!force && existingToday) {
    if (String(existingToday.status).toLowerCase() === 'sent') {
      logSend_(new Date(), existingToday.issue_id, 'draft-check', 0, '', 'already_sent', 'Today has already been sent.');
      return null;
    }
    return existingToday;
  }

  const candidates = fetchCandidateItems_();
  writeCandidates_(today, candidates);
  const issue = composeIssue_(today, candidates);
  appendIssue_(issue);
  return getIssueById_(issue.issue_id);
}

function fetchCandidateItems_() {
  const config = getConfigMap_();
  const maxCandidates = numberFromConfig_(config, 'max_candidates', 60);
  const lookbackDays = numberFromConfig_(config, 'lookback_days', 7);
  const sources = getRowsAsObjects_(SHEET_NAMES.SOURCES)
    .filter(function (source) {
      return truthy_(source.enabled) && String(source.type).toLowerCase() === 'rss' && source.url;
    });

  const now = new Date();
  const cutoff = new Date(now.getTime() - lookbackDays * 24 * 60 * 60 * 1000);
  let allItems = [];

  sources.forEach(function (source) {
    try {
      const items = fetchFeedItems_(source, cutoff, now);
      allItems = allItems.concat(items);
    } catch (error) {
      logSend_(new Date(), '', 'fetch:' + source.name, 0, '', 'source_error', errorToString_(error));
    }
  });

  const seen = {};
  const deduped = [];
  allItems
    .sort(function (a, b) {
      return b.score - a.score;
    })
    .forEach(function (item) {
      const key = dedupeKey_(item);
      if (!seen[key]) {
        seen[key] = true;
        deduped.push(item);
      }
    });

  return deduped.slice(0, maxCandidates);
}

function fetchFeedItems_(source, cutoff, now) {
  const response = UrlFetchApp.fetch(source.url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      'User-Agent': 'AI Daily Newsletter Apps Script/' + NEWSLETTER_VERSION
    }
  });
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('HTTP ' + code + ' from ' + source.url);
  }

  const xml = response.getContentText();
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const feedItems = parseFeedRoot_(root, source);
  return feedItems
    .filter(function (item) {
      return item.title && item.url && item.published_at && item.published_at >= cutoff && item.published_at <= now;
    })
    .map(function (item) {
      item.score = scoreItem_(item, source, now);
      item.summary = summarizeText_(item.summary || item.title, 240);
      item.why_it_matters = whyItMatters_(item);
      return item;
    });
}

function parseFeedRoot_(root, source) {
  const rootName = String(root.getName()).toLowerCase();
  if (rootName === 'feed') {
    return childrenByName_(root, 'entry').map(function (entry) {
      return parseAtomEntry_(entry, source);
    });
  }

  let itemParent = firstChildByName_(root, 'channel') || root;
  return childrenByName_(itemParent, 'item').map(function (item) {
    return parseRssItem_(item, source);
  });
}

function parseRssItem_(item, source) {
  const title = cleanText_(firstChildText_(item, ['title']));
  const url = cleanText_(firstChildText_(item, ['link'])) || cleanText_(firstChildText_(item, ['guid']));
  const dateText = firstChildText_(item, ['pubDate', 'published', 'updated', 'date']);
  const summary = firstChildText_(item, ['description', 'summary', 'encoded', 'content']);
  return {
    source_name: source.name,
    category: source.category || 'General AI',
    title: title,
    url: normalizeUrl_(url),
    published_at: parseDate_(dateText),
    summary: htmlToText_(summary),
    score: 0
  };
}

function parseAtomEntry_(entry, source) {
  const title = cleanText_(firstChildText_(entry, ['title']));
  const url = atomLink_(entry) || cleanText_(firstChildText_(entry, ['id']));
  const dateText = firstChildText_(entry, ['published', 'updated', 'date']);
  const summary = firstChildText_(entry, ['summary', 'content']);
  return {
    source_name: source.name,
    category: source.category || 'General AI',
    title: title,
    url: normalizeUrl_(url),
    published_at: parseDate_(dateText),
    summary: htmlToText_(summary),
    score: 0
  };
}

function composeIssue_(issueDate, candidates) {
  const config = getConfigMap_();
  const timezone = config.timezone || 'America/Toronto';
  const maxItems = numberFromConfig_(config, 'max_items_in_issue', 8);
  const accessible = candidates.filter(function (item) {
    return !isResearchItem_(item) && !isDeepResearchItem_(item);
  });
  const researchPool = candidates.filter(function (item) {
    return isResearchItem_(item) && !isDeepResearchItem_(item);
  });
  const selected = accessible.concat(researchPool).slice(0, maxItems);
  const usedKeys = {};
  const topStories = selected.filter(function (item) {
    return !isResearchItem_(item);
  }).slice(0, 3);
  markItemsUsed_(topStories, usedKeys);
  const tools = selected.filter(function (item) {
    return !usedKeys[dedupeKey_(item)] && itemMatches_(item, ['tool', 'app', 'feature', 'assistant', 'copilot', 'launch', 'release', 'agent', 'model']);
  }).slice(0, 3);
  markItemsUsed_(tools, usedKeys);
  const research = selected.filter(function (item) {
    return !usedKeys[dedupeKey_(item)] && isResearchItem_(item);
  }).slice(0, 1);
  markItemsUsed_(research, usedKeys);
  const links = selected.filter(function (item) {
    return !usedKeys[dedupeKey_(item)];
  });
  const fact = aiFactForDate_(issueDate, timezone);
  const subject = buildSubject_(topStories);
  const issueId = formatDate_(new Date(), 'yyyyMMdd-HHmmss', timezone);
  const htmlBody = renderIssueHtml_(issueDate, topStories, tools, research, links, fact);
  const textBody = renderIssueText_(issueDate, topStories, tools, research, links, fact);

  return {
    issue_id: issueId,
    issue_date: issueDate,
    status: 'draft',
    subject: subject,
    html_body: htmlBody,
    text_body: textBody,
    preview_sent_at: '',
    approved_at: '',
    sent_at: '',
    sent_count: 0,
    error_count: 0,
    notes: 'Review this row. Change status to approved when ready to send.'
  };
}

function renderIssueHtml_(issueDate, topStories, tools, research, selected, fact) {
  const newsletterName = getConfigValue_('newsletter_name', 'AI Daily');
  const footerReason = getConfigValue_('footer_reason', 'You are receiving this because you joined the invite-only AI Daily beta.');
  const html = [
    '<div style="margin:0;padding:0;background:#f6f7f9;color:#1f2933;font-family:Arial,Helvetica,sans-serif;">',
    '<div style="max-width:680px;margin:0 auto;padding:24px 16px;">',
    '<div style="background:#ffffff;border:1px solid #e5e7eb;border-radius:8px;padding:24px;">',
    '<p style="margin:0 0 6px 0;color:#6b7280;font-size:13px;">' + escapeHtml_(issueDate) + '</p>',
    '<h1 style="margin:0 0 18px 0;font-size:26px;line-height:1.2;color:#111827;">' + escapeHtml_(newsletterName) + '</h1>',
    '<p style="margin:0 0 22px 0;font-size:15px;line-height:1.5;color:#374151;">A plain-English daily scan of the AI news, tools, and changes worth knowing.</p>',
    renderSectionHtml_('What Happened', topStories),
    renderSectionHtml_('Useful Tools & Features', tools),
    renderSectionHtml_('Research, Translated', research),
    '<h2 style="margin:26px 0 10px 0;font-size:18px;color:#111827;">One Useful AI Fact</h2>',
    '<p style="margin:0 0 18px 0;font-size:15px;line-height:1.6;color:#374151;">' + escapeHtml_(fact) + '</p>',
    renderSectionHtml_('More To Skim', selected),
    '<hr style="border:none;border-top:1px solid #e5e7eb;margin:26px 0 16px 0;">',
    '<p style="margin:0 0 8px 0;font-size:12px;line-height:1.5;color:#6b7280;">' + escapeHtml_(footerReason) + '</p>',
    '<p style="margin:0;font-size:12px;line-height:1.5;color:#6b7280;">Reply with corrections, feedback, or tips. <a href="{{UNSUBSCRIBE_URL}}" style="color:#2563eb;">Unsubscribe</a></p>',
    '</div>',
    '</div>',
    '</div>'
  ].join('');
  return html;
}

function renderSectionHtml_(title, items) {
  if (!items || items.length === 0) {
    return '';
  }
  const blocks = items.map(function (item) {
    return [
      '<div style="margin:0 0 16px 0;padding:0 0 14px 0;border-bottom:1px solid #eef2f7;">',
      '<p style="margin:0 0 4px 0;font-size:12px;color:#6b7280;">' + escapeHtml_(item.source_name) + ' &middot; ' + escapeHtml_(item.category) + '</p>',
      '<h3 style="margin:0 0 6px 0;font-size:16px;line-height:1.35;color:#111827;"><a href="' + escapeAttribute_(item.url) + '" style="color:#111827;text-decoration:underline;">' + escapeHtml_(item.title) + '</a></h3>',
      '<p style="margin:0 0 6px 0;font-size:14px;line-height:1.55;color:#374151;">' + escapeHtml_(plainEnglishSummary_(item)) + '</p>',
      '<p style="margin:0;font-size:13px;line-height:1.5;color:#4b5563;"><strong>Why it matters:</strong> ' + escapeHtml_(item.why_it_matters) + '</p>',
      '</div>'
    ].join('');
  }).join('');
  return '<h2 style="margin:26px 0 12px 0;font-size:18px;color:#111827;">' + escapeHtml_(title) + '</h2>' + blocks;
}

function renderIssueText_(issueDate, topStories, tools, research, selected, fact) {
  const newsletterName = getConfigValue_('newsletter_name', 'AI Daily');
  const footerReason = getConfigValue_('footer_reason', 'You are receiving this because you joined the invite-only AI Daily beta.');
  return [
    newsletterName + ' - ' + issueDate,
    '',
    'A plain-English daily scan of the AI news, tools, and changes worth knowing.',
    '',
    renderSectionText_('What Happened', topStories),
    renderSectionText_('Useful Tools & Features', tools),
    renderSectionText_('Research, Translated', research),
    'One Useful AI Fact',
    fact,
    '',
    renderSectionText_('More To Skim', selected),
    footerReason,
    'Unsubscribe: {{UNSUBSCRIBE_URL}}',
    'Reply with corrections, feedback, or tips.'
  ].join('\n');
}

function renderSectionText_(title, items) {
  if (!items || items.length === 0) {
    return '';
  }
  const lines = [title];
  items.forEach(function (item, index) {
    lines.push((index + 1) + '. ' + item.title);
    lines.push('Source: ' + item.source_name + ' / ' + item.category);
    lines.push(plainEnglishSummary_(item));
    lines.push('Why it matters: ' + item.why_it_matters);
    lines.push(item.url);
    lines.push('');
  });
  return lines.join('\n');
}

function buildSubject_(topStories) {
  if (!topStories || topStories.length === 0) {
    return 'AI Daily: Today in AI';
  }
  const first = shorten_(topStories[0].title, 58);
  const second = topStories.length > 1 ? shorten_(topStories[1].title, 46) : 'new AI updates';
  return 'AI Daily: ' + first + ' + ' + second;
}

function sendPreview_(issue) {
  const adminEmail = getAdminEmail_();
  const html = personalizeBody_(issue.html_body, {
    unsubscribeUrl: '#preview-unsubscribe-link',
    email: adminEmail
  });
  const text = personalizeBody_(issue.text_body, {
    unsubscribeUrl: '#preview-unsubscribe-link',
    email: adminEmail
  });

  GmailApp.sendEmail(adminEmail, '[PREVIEW] ' + issue.subject, text, {
    htmlBody: [
      '<div style="font-family:Arial,Helvetica,sans-serif;background:#fff7ed;border:1px solid #fed7aa;padding:12px;margin-bottom:12px;">',
      '<strong>Preview only.</strong> To send, change this issue row status to <code>approved</code>, then run Send approved issue now or wait for the morning send trigger.',
      '</div>',
      html
    ].join(''),
    name: getConfigValue_('from_name', getConfigValue_('newsletter_name', 'AI Daily')),
    replyTo: getReplyTo_()
  });

  updateIssue_(issue.rowNumber, { preview_sent_at: new Date() });
}

function sendIssueToSubscriber_(issue, subscriber) {
  const unsubscribeUrl = buildUnsubscribeUrl_(subscriber.unsubscribe_token, subscriber.email);
  const html = personalizeBody_(issue.html_body, {
    unsubscribeUrl: unsubscribeUrl,
    email: subscriber.email
  });
  const text = personalizeBody_(issue.text_body, {
    unsubscribeUrl: unsubscribeUrl,
    email: subscriber.email
  });

  GmailApp.sendEmail(subscriber.email, issue.subject, text, {
    htmlBody: html,
    name: getConfigValue_('from_name', getConfigValue_('newsletter_name', 'AI Daily')),
    replyTo: getReplyTo_()
  });
}

function personalizeBody_(body, values) {
  return String(body || '')
    .replace(/\{\{UNSUBSCRIBE_URL\}\}/g, values.unsubscribeUrl)
    .replace(/\{\{RECIPIENT_EMAIL\}\}/g, values.email);
}

function writeCandidates_(issueDate, candidates) {
  const sheet = getSheet_(SHEET_NAMES.CANDIDATES);
  clearDataRows_(sheet);
  const rows = candidates.map(function (item, index) {
    return [
      issueDate,
      item.source_name,
      item.category,
      item.title,
      item.url,
      item.published_at,
      item.summary,
      item.why_it_matters,
      item.score,
      index < 8
    ];
  });
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, HEADERS[SHEET_NAMES.CANDIDATES].length).setValues(rows);
  }
}

function appendIssue_(issue) {
  appendObject_(SHEET_NAMES.ISSUES, issue);
}

function updateIssue_(rowNumber, patch) {
  updateRowByHeader_(SHEET_NAMES.ISSUES, rowNumber, patch);
}

function maybeAlertHeldSend_(issue) {
  if (!truthy_(getConfigValue_('alert_on_held_send', 'TRUE'))) {
    return;
  }
  const today = formatDate_(new Date(), 'yyyy-MM-dd', getConfigValue_('timezone', 'America/Toronto'));
  const noteKey = 'held_alert_' + issue.issue_id + '_' + today;
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty(noteKey)) {
    return;
  }
  properties.setProperty(noteKey, 'sent');
  sendAdminAlert_(
    'AI Daily held: issue is not approved',
    'Issue ' + issue.issue_id + ' is still "' + issue.status + '". No subscriber email was sent.'
  );
}

function sendAdminAlert_(subject, message) {
  const admin = getAdminEmail_();
  GmailApp.sendEmail(admin, subject, message, {
    name: getConfigValue_('from_name', getConfigValue_('newsletter_name', 'AI Daily')),
    replyTo: getReplyTo_()
  });
}

function unsubscribeByToken_(token) {
  const subscribers = getRowsAsObjects_(SHEET_NAMES.SUBSCRIBERS);
  const match = subscribers.find(function (subscriber) {
    return String(subscriber.unsubscribe_token) === String(token);
  });

  if (!match) {
    return HtmlService.createHtmlOutput('<p>Unsubscribe link not found. Reply to the newsletter and we will remove you manually.</p>');
  }

  updateRowByHeader_(SHEET_NAMES.SUBSCRIBERS, match.rowNumber, {
    status: 'unsubscribed'
  });
  return HtmlService.createHtmlOutput('<p>You have been unsubscribed from AI Daily.</p>');
}

function unsubscribeByEmail_(email) {
  const subscribers = getRowsAsObjects_(SHEET_NAMES.SUBSCRIBERS);
  const match = subscribers.find(function (subscriber) {
    return normalizeEmail_(subscriber.email) === email;
  });

  if (!match) {
    return {
      ok: false,
      message: '<p>That email was not found. Reply to the newsletter and we will help.</p>'
    };
  }

  updateRowByHeader_(SHEET_NAMES.SUBSCRIBERS, match.rowNumber, {
    status: 'unsubscribed'
  });
  return {
    ok: true,
    message: '<p>You have been unsubscribed from AI Daily.</p>'
  };
}

function renderUnsubscribeForm_(message) {
  const html = [
    '<!doctype html><html><head><base target="_top"><meta name="viewport" content="width=device-width, initial-scale=1"></head>',
    '<body style="font-family:Arial,Helvetica,sans-serif;max-width:520px;margin:40px auto;padding:0 16px;">',
    '<h1 style="font-size:24px;">Unsubscribe from AI Daily</h1>',
    message ? '<p style="color:#b91c1c;">' + escapeHtml_(message) + '</p>' : '',
    '<form method="post">',
    '<label for="email">Email address</label><br>',
    '<input id="email" name="email" type="email" required style="width:100%;padding:10px;margin:8px 0 16px 0;box-sizing:border-box;">',
    '<button type="submit" style="padding:10px 14px;">Unsubscribe</button>',
    '</form>',
    '</body></html>'
  ].join('');
  return HtmlService.createHtmlOutput(html);
}

function getActiveSubscribers_() {
  return getRowsAsObjects_(SHEET_NAMES.SUBSCRIBERS)
    .filter(function (subscriber) {
      return isValidEmail_(subscriber.email) &&
        String(subscriber.status || '').toLowerCase() === 'active' &&
        truthy_(subscriber.confirmed);
    });
}

function assignSubscriberDefaults_() {
  const sheet = getSheet_(SHEET_NAMES.SUBSCRIBERS);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return;
  }
  const headers = values[0];
  const indexes = headerIndexes_(headers);
  let changed = false;

  for (let row = 1; row < values.length; row += 1) {
    const email = normalizeEmail_(values[row][indexes.email]);
    if (!email) {
      continue;
    }
    values[row][indexes.email] = email;
    if (!values[row][indexes.status]) {
      values[row][indexes.status] = 'active';
      changed = true;
    }
    if (!values[row][indexes.source]) {
      values[row][indexes.source] = 'invite';
      changed = true;
    }
    if (!values[row][indexes.consent_date]) {
      values[row][indexes.consent_date] = new Date();
      changed = true;
    }
    if (values[row][indexes.confirmed] === '') {
      values[row][indexes.confirmed] = true;
      changed = true;
    }
    if (!values[row][indexes.unsubscribe_token]) {
      values[row][indexes.unsubscribe_token] = Utilities.getUuid();
      changed = true;
    }
  }

  if (changed) {
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  }
}

function setSubscriberLastSent_(rowNumber, date) {
  updateRowByHeader_(SHEET_NAMES.SUBSCRIBERS, rowNumber, {
    last_sent_at: date
  });
}

function getSentEmailSet_(issueId) {
  const sent = {};
  getRowsAsObjects_(SHEET_NAMES.SEND_LOG).forEach(function (log) {
    if (String(log.issue_id) === String(issueId) && String(log.status).toLowerCase() === 'sent') {
      String(log.recipients || '').split(',').forEach(function (email) {
        const normalized = normalizeEmail_(email);
        if (normalized) {
          sent[normalized] = true;
        }
      });
    }
  });
  return sent;
}

function logSend_(timestamp, issueId, batch, recipientCount, recipients, status, error) {
  appendObject_(SHEET_NAMES.SEND_LOG, {
    timestamp: timestamp,
    issue_id: issueId,
    batch: batch,
    recipient_count: recipientCount,
    recipients: recipients,
    status: status,
    error: error
  });
}

function seedConfigDefaults_() {
  const existing = {};
  getRowsAsObjects_(SHEET_NAMES.CONFIG).forEach(function (row) {
    if (row.key) {
      existing[String(row.key)] = true;
    }
  });

  const activeEmail = Session.getActiveUser().getEmail();
  DEFAULT_CONFIG.forEach(function (row) {
    if (!existing[row[0]]) {
      const value = row[0] === 'admin_email' || row[0] === 'reply_to'
        ? activeEmail
        : row[1];
      getSheet_(SHEET_NAMES.CONFIG).appendRow([row[0], value, row[2]]);
    }
  });
}

function seedSourceDefaults_() {
  const existingUrls = {};
  getRowsAsObjects_(SHEET_NAMES.SOURCES).forEach(function (row) {
    if (row.url) {
      existingUrls[String(row.url)] = true;
    }
  });

  DEFAULT_SOURCES.forEach(function (source) {
    if (!existingUrls[source[1]]) {
      getSheet_(SHEET_NAMES.SOURCES).appendRow(source);
    }
  });
}

function refreshGeneralReaderSources() {
  const sheet = getSheet_(SHEET_NAMES.SOURCES);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 1) {
    ensureHeaders_(sheet, HEADERS[SHEET_NAMES.SOURCES]);
  }

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = {};
  header.forEach(function (name, i) {
    index[String(name)] = i + 1;
  });

  const rowsByUrl = {};
  getRowsAsObjects_(SHEET_NAMES.SOURCES).forEach(function (row) {
    if (row.url) {
      rowsByUrl[String(row.url)] = row.rowNumber;
    }
  });

  DEFAULT_SOURCES.forEach(function (source) {
    const rowNumber = rowsByUrl[source[1]];
    if (rowNumber) {
      sheet.getRange(rowNumber, index.name).setValue(source[0]);
      sheet.getRange(rowNumber, index.category).setValue(source[3]);
      sheet.getRange(rowNumber, index.enabled).setValue(source[4]);
      sheet.getRange(rowNumber, index.weight).setValue(source[5]);
    } else {
      sheet.appendRow(source);
    }
  });

  SpreadsheetApp.flush();
}

function getSpreadsheet_() {
  const configuredId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (configuredId) {
    return SpreadsheetApp.openById(configuredId);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet. Bind this script to the newsletter Google Sheet or set Script Property SPREADSHEET_ID.');
  }
  return ss;
}

function ensureSheet_(ss, name, requiredHeaders) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  ensureHeaders_(sheet, requiredHeaders);
  sheet.setFrozenRows(1);
}

function ensureHeaders_(sheet, requiredHeaders) {
  const lastColumn = Math.max(sheet.getLastColumn(), requiredHeaders.length);
  const firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const existing = firstRow.filter(function (value) {
    return value !== '';
  });

  if (existing.length === 0) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return;
  }

  const merged = existing.slice();
  requiredHeaders.forEach(function (header) {
    if (merged.indexOf(header) === -1) {
      merged.push(header);
    }
  });
  sheet.getRange(1, 1, 1, merged.length).setValues([merged]);
}

function getSheet_(name) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) {
    throw new Error('Missing sheet "' + name + '". Run setupNewsletter first.');
  }
  return sheet;
}

function getRowsAsObjects_(sheetName) {
  const sheet = getSheet_(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }
  const headers = values[0].map(String);
  const rows = [];
  for (let row = 1; row < values.length; row += 1) {
    const obj = { rowNumber: row + 1 };
    let hasValue = false;
    headers.forEach(function (header, col) {
      obj[header] = values[row][col];
      if (values[row][col] !== '') {
        hasValue = true;
      }
    });
    if (hasValue) {
      rows.push(obj);
    }
  }
  return rows;
}

function appendObject_(sheetName, object) {
  const sheet = getSheet_(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const row = headers.map(function (header) {
    return Object.prototype.hasOwnProperty.call(object, header) ? object[header] : '';
  });
  sheet.appendRow(row);
}

function updateRowByHeader_(sheetName, rowNumber, patch) {
  const sheet = getSheet_(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  Object.keys(patch).forEach(function (key) {
    const index = headers.indexOf(key);
    if (index === -1) {
      throw new Error('Missing header "' + key + '" in ' + sheetName);
    }
    sheet.getRange(rowNumber, index + 1).setValue(patch[key]);
  });
}

function clearDataRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}

function getConfigMap_() {
  const map = {};
  getRowsAsObjects_(SHEET_NAMES.CONFIG).forEach(function (row) {
    if (row.key) {
      map[String(row.key)] = row.value;
    }
  });
  return map;
}

function getConfigValue_(key, fallback) {
  const map = getConfigMap_();
  return map[key] !== undefined && map[key] !== '' ? map[key] : fallback;
}

function numberFromConfig_(config, key, fallback) {
  const value = Number(config[key]);
  return Number.isFinite(value) && value > 0 ? value : fallback;
}

function getAdminEmail_() {
  const configured = getConfigValue_('admin_email', '');
  const active = Session.getActiveUser().getEmail();
  const email = normalizeEmail_(configured || active);
  if (!email) {
    throw new Error('Set admin_email in the Config tab.');
  }
  return email;
}

function getReplyTo_() {
  return normalizeEmail_(getConfigValue_('reply_to', '')) || getAdminEmail_();
}

function getLatestIssue_() {
  const issues = getRowsAsObjects_(SHEET_NAMES.ISSUES);
  if (!issues.length) {
    return null;
  }
  return issues[issues.length - 1];
}

function getLatestIssueForDate_(issueDate) {
  const issues = getRowsAsObjects_(SHEET_NAMES.ISSUES)
    .filter(function (issue) {
      return String(issue.issue_date) === String(issueDate);
    });
  return issues.length ? issues[issues.length - 1] : null;
}

function getIssueById_(issueId) {
  const matches = getRowsAsObjects_(SHEET_NAMES.ISSUES)
    .filter(function (issue) {
      return String(issue.issue_id) === String(issueId);
    });
  return matches.length ? matches[0] : null;
}

function scoreItem_(item, source, now) {
  let score = Number(source.weight) || 5;
  const ageHours = Math.max(0, (now.getTime() - item.published_at.getTime()) / (60 * 60 * 1000));
  if (ageHours <= 24) {
    score += 8;
  } else if (ageHours <= 72) {
    score += 5;
  } else if (ageHours <= 168) {
    score += 2;
  } else {
    score -= 3;
  }

  const text = (item.title + ' ' + item.summary + ' ' + item.category).toLowerCase();
  const keywordWeights = {
    'launch': 4,
    'release': 4,
    'introducing': 4,
    'new feature': 4,
    'available': 3,
    'app': 3,
    'assistant': 3,
    'copilot': 3,
    'agent': 3,
    'voice': 3,
    'video': 3,
    'image': 2,
    'consumer': 3,
    'school': 2,
    'workplace': 2,
    'business': 2,
    'enterprise': 2,
    'pricing': 2,
    'safety': 2,
    'policy': 2,
    'regulation': 2,
    'open source': 1,
    'api': 1,
    'model': 1,
    'benchmark': -2,
    'research': -2,
    'paper': -3,
    'arxiv': -4,
    'eval': -2
  };

  Object.keys(keywordWeights).forEach(function (keyword) {
    if (text.indexOf(keyword) !== -1) {
      score += keywordWeights[keyword];
    }
  });

  if (isDeepResearchItem_(item)) {
    score -= 8;
  } else if (isResearchItem_(item)) {
    score -= 3;
  }

  if (text.indexOf('sponsored') !== -1 || text.indexOf('webinar') !== -1) {
    score -= 4;
  }
  return score;
}

function whyItMatters_(item) {
  const text = (item.title + ' ' + item.summary + ' ' + item.category).toLowerCase();
  if (text.indexOf('api') !== -1 || text.indexOf('developer') !== -1) {
    return 'This may soon show up inside apps, websites, or workflows people already use.';
  }
  if (text.indexOf('agent') !== -1) {
    return 'Agent features matter because they can handle multi-step chores, not just answer one question at a time.';
  }
  if (text.indexOf('open source') !== -1 || text.indexOf('open-weight') !== -1) {
    return 'More open releases can make powerful AI cheaper and easier for smaller teams to try.';
  }
  if (isResearchItem_(item)) {
    return 'The useful part is the direction of travel: what AI may become better at next, in plain terms.';
  }
  if (text.indexOf('safety') !== -1 || text.indexOf('policy') !== -1 || text.indexOf('regulation') !== -1) {
    return 'Rules and safety choices shape which AI tools people can trust and companies can actually ship.';
  }
  if (text.indexOf('funding') !== -1 || text.indexOf('acquisition') !== -1 || text.indexOf('enterprise') !== -1) {
    return 'Money and partnerships show which AI ideas are moving from experiments into real products.';
  }
  return 'This is worth knowing because it may affect the AI tools, apps, or companies people hear about next.';
}

function aiFactForDate_(issueDate, timezone) {
  const date = parseDate_(issueDate) || new Date();
  const day = Number(formatDate_(date, 'D', timezone)) || 1;
  return AI_FACTS[(day - 1) % AI_FACTS.length];
}

function itemMatches_(item, keywords) {
  const text = (item.title + ' ' + item.summary + ' ' + item.category + ' ' + item.source_name).toLowerCase();
  return keywords.some(function (keyword) {
    return text.indexOf(keyword.toLowerCase()) !== -1;
  });
}

function isResearchItem_(item) {
  return itemMatches_(item, ['research', 'paper', 'arxiv', 'benchmark', 'eval', 'technical report']) ||
    String(item.category).toLowerCase().indexOf('research') !== -1;
}

function isDeepResearchItem_(item) {
  return itemMatches_(item, ['arxiv', 'cs.ai', 'cs.lg', 'benchmark', 'evaluation suite', 'technical report', 'proceedings']);
}

function plainEnglishSummary_(item) {
  const text = summarizeText_(item.summary || item.title, 220)
    .replace(/^abstract\s*[:.-]\s*/i, '')
    .replace(/\bwe (propose|present|introduce|study|evaluate)\b/gi, 'the authors $1')
    .replace(/\bstate-of-the-art\b/gi, 'top-performing')
    .replace(/\bLLMs?\b/g, 'AI models')
    .replace(/\bevals?\b/gi, 'tests');

  if (!isResearchItem_(item)) {
    return text;
  }

  if (itemMatches_(item, ['robot', 'video', 'image', 'voice', 'audio'])) {
    return 'Researchers are showing progress in how AI understands or creates media and physical-world actions. ' + text;
  }
  if (itemMatches_(item, ['safety', 'alignment', 'risk'])) {
    return 'This is about making AI systems more predictable and safer to use. ' + text;
  }
  return 'This is a research update, so treat it as an early signal rather than a finished product. ' + text;
}

function markItemsUsed_(items, usedKeys) {
  items.forEach(function (item) {
    usedKeys[dedupeKey_(item)] = true;
  });
}

function atomLink_(entry) {
  const links = childrenByName_(entry, 'link');
  for (let i = 0; i < links.length; i += 1) {
    const rel = attributeValue_(links[i], 'rel');
    const href = attributeValue_(links[i], 'href');
    if (href && (!rel || rel === 'alternate')) {
      return normalizeUrl_(href);
    }
  }
  return links.length ? normalizeUrl_(attributeValue_(links[0], 'href') || links[0].getText()) : '';
}

function firstChildText_(element, names) {
  for (let i = 0; i < names.length; i += 1) {
    const child = firstChildByName_(element, names[i]);
    if (child) {
      return child.getText();
    }
  }
  return '';
}

function firstChildByName_(element, name) {
  const lower = String(name).toLowerCase();
  const children = element.getChildren();
  for (let i = 0; i < children.length; i += 1) {
    if (String(children[i].getName()).toLowerCase() === lower) {
      return children[i];
    }
  }
  return null;
}

function childrenByName_(element, name) {
  const lower = String(name).toLowerCase();
  return element.getChildren().filter(function (child) {
    return String(child.getName()).toLowerCase() === lower;
  });
}

function attributeValue_(element, name) {
  const attribute = element.getAttribute(name);
  return attribute ? attribute.getValue() : '';
}

function parseDate_(value) {
  if (value instanceof Date) {
    return value;
  }
  if (!value) {
    return new Date();
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return new Date();
  }
  return date;
}

function formatDate_(date, pattern, timezone) {
  return Utilities.formatDate(parseDate_(date), timezone || 'America/Toronto', pattern);
}

function cleanText_(text) {
  return String(text || '').replace(/\s+/g, ' ').trim();
}

function htmlToText_(html) {
  return cleanText_(String(html || '')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'"));
}

function summarizeText_(text, maxLength) {
  const clean = cleanText_(text);
  if (clean.length <= maxLength) {
    return clean;
  }
  const truncated = clean.slice(0, maxLength + 1);
  const sentenceEnd = Math.max(truncated.lastIndexOf('. '), truncated.lastIndexOf('! '), truncated.lastIndexOf('? '));
  if (sentenceEnd > 80) {
    return truncated.slice(0, sentenceEnd + 1);
  }
  const lastSpace = truncated.lastIndexOf(' ');
  return truncated.slice(0, lastSpace > 80 ? lastSpace : maxLength).trim() + '...';
}

function normalizeUrl_(url) {
  const clean = cleanText_(url);
  if (!clean) {
    return '';
  }
  return clean
    .replace(/([?&])utm_[^=&]+=[^&]+/gi, '$1')
    .replace(/([?&])fbclid=[^&]+/gi, '$1')
    .replace(/[?&]$/, '')
    .replace(/#.*$/, '');
}

function dedupeKey_(item) {
  if (item.url) {
    return normalizeUrl_(item.url).toLowerCase();
  }
  return cleanText_(item.title).toLowerCase();
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeAttribute_(value) {
  return escapeHtml_(value).replace(/`/g, '&#96;');
}

function shorten_(text, maxLength) {
  const clean = cleanText_(text);
  if (clean.length <= maxLength) {
    return clean;
  }
  const clipped = clean.slice(0, maxLength + 1);
  const lastSpace = clipped.lastIndexOf(' ');
  return clipped.slice(0, lastSpace > 20 ? lastSpace : maxLength).trim() + '...';
}

function truthy_(value) {
  if (value === true) {
    return true;
  }
  const normalized = String(value || '').trim().toLowerCase();
  return normalized === 'true' || normalized === 'yes' || normalized === '1' || normalized === 'active';
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizeEmail_(email));
}

function headerIndexes_(headers) {
  const indexes = {};
  headers.forEach(function (header, index) {
    indexes[String(header)] = index;
  });
  return indexes;
}

function padNumber_(value, length) {
  let text = String(value);
  while (text.length < length) {
    text = '0' + text;
  }
  return text;
}

function buildUnsubscribeUrl_(token, email) {
  const configured = getConfigValue_('unsubscribe_base_url', '');
  let base = configured;
  if (!base) {
    try {
      base = ScriptApp.getService().getUrl();
    } catch (error) {
      base = '';
    }
  }
  if (base) {
    return base + (base.indexOf('?') === -1 ? '?' : '&') + 'token=' + encodeURIComponent(token);
  }
  return 'mailto:' + encodeURIComponent(getReplyTo_()) + '?subject=' + encodeURIComponent('Unsubscribe ' + email);
}

function errorToString_(error) {
  if (!error) {
    return '';
  }
  return error.stack || error.message || String(error);
}

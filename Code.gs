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
  ['OpenAI News', 'https://openai.com/news/rss.xml', 'rss', 'General AI', true, 10],
  ['Google DeepMind Blog', 'https://deepmind.google/blog/rss.xml', 'rss', 'Research', true, 10],
  ['Google AI', 'https://blog.google/technology/ai/rss/', 'rss', 'General AI', true, 8],
  ['Google Research', 'https://research.google/blog/rss/', 'rss', 'Research', true, 8],
  ['Microsoft AI Blog', 'https://blogs.microsoft.com/ai/feed/', 'rss', 'Business', true, 7],
  ['Hugging Face Blog', 'https://huggingface.co/blog/feed.xml', 'rss', 'Tools', true, 8],
  ['Meta Newsroom', 'https://about.fb.com/feed/', 'rss', 'Business', true, 5],
  ['Meta Engineering', 'https://engineering.fb.com/feed/', 'rss', 'Research', true, 6],
  ['arXiv cs.AI', 'https://export.arxiv.org/rss/cs.AI', 'rss', 'Research', true, 6],
  ['arXiv cs.LG', 'https://export.arxiv.org/rss/cs.LG', 'rss', 'Research', true, 6],
  ['MIT Technology Review AI', 'https://www.technologyreview.com/topic/artificial-intelligence/feed/', 'rss', 'Business', true, 7],
  ['The Verge AI', 'https://www.theverge.com/rss/ai-artificial-intelligence/index.xml', 'rss', 'General AI', true, 6],
  ['VentureBeat AI', 'https://venturebeat.com/category/ai/feed', 'rss', 'Business', true, 6],
  ['Anthropic via Google News', 'https://news.google.com/rss/search?q=site%3Aanthropic.com%2Fnews%20Anthropic&hl=en-US&gl=US&ceid=US%3Aen', 'rss', 'General AI', true, 7]

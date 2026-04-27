/* ═══════════════════════════════════════════════════════════════════════════
   MBA Domain & Role Discovery Tool — Application Logic
   Scoring logic mirrors the original Excel workbook exactly.
   ═══════════════════════════════════════════════════════════════════════════ */

// ─── GOOGLE SHEETS CONFIG ─────────────────────────────────────────────────────
// Paste your Google Apps Script Web App URL here after deployment.
// Instructions: see google-apps-script.js in this project.
const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzlfEzS80nSCmzXy7JkbIOWMRcArdzwPYaalm-CwiwAP3zEp-A8KzrAiuAoAYxXGYaQ/exec';

// ─── QUESTIONS ───────────────────────────────────────────────────────────────

const PARTS = [
  { label: 'Part 1: What Energizes You?', qs: [0, 1, 2, 3, 4] },
  { label: 'Part 2: What Are You Good At?', qs: [5, 6, 7, 8, 9] },
  { label: 'Part 3: How Do You Want to Work?', qs: [10, 11, 12, 13, 14] },
  { label: 'Part 4: Real-World Signals', qs: [15, 16, 17, 18, 19] },
];

const QUESTIONS = [
  // ── Part 1 ──
  {
    text: 'When you have a free Saturday, you\'re most likely to:',
    options: {
      A: 'Research a random topic until you become an expert',
      B: 'Meet friends, network, or talk to new people',
      C: 'Work on creative stuff — writing, designing, content',
      D: 'Organize your space, plan your week, optimize something',
    }
  },
  {
    text: 'In a group project, you naturally gravitate towards:',
    options: {
      A: 'Analyzing the problem, building frameworks, strategy',
      B: 'Presenting, handling external communication',
      C: 'Creating the design, writing content, making it look good',
      D: 'Making timelines, tracking progress, ensuring delivery',
    }
  },
  {
    text: 'What type of wins give you the MOST satisfaction?',
    options: {
      A: 'Cracking a complex problem others couldn\'t solve',
      B: 'Convincing someone skeptical to say \'yes\'',
      C: 'Creating something original that people appreciate',
      D: 'Seeing a messy process become smooth and efficient',
    }
  },
  {
    text: 'When learning something new, you prefer to:',
    options: {
      A: 'Understand deeply — research, analyze first principles',
      B: 'Learn by doing — jump in, figure it out through action',
      C: 'Learn through examples — study great work, then create',
      D: 'Follow structure — step-by-step guides, proven methods',
    }
  },
  {
    text: 'What frustrates you MOST at work or college?',
    options: {
      A: 'Superficial analysis — people don\'t think deeply enough',
      B: 'Slow decisions — people overthink instead of acting',
      C: 'Boring template work — no room for creativity',
      D: 'Chaos and disorganization — no system or process',
    }
  },
  // ── Part 2 ──
  {
    text: 'When given an Excel sheet with messy data, you:',
    options: {
      A: 'Get excited — love cleaning data, finding patterns',
      B: 'Feel okay — can do it but not your favourite thing',
      C: 'Feel dread — rather write or design than crunch numbers',
      D: 'Delegate — tell someone what analysis is needed',
    }
  },
  {
    text: 'In conversations with strangers, you typically:',
    options: {
      A: 'Feel drained — prefer deep talks with people you know',
      B: 'Feel okay — can do it when required',
      C: 'Feel energized — genuinely enjoy meeting new people',
      D: 'Feel strategic — see it as an opportunity to network',
    }
  },
  {
    text: 'When writing (emails, social posts), people often say:',
    options: {
      A: 'Your writing is clear and gets the point across',
      B: 'Your writing is persuasive and makes people act',
      C: 'Your writing is creative with a unique voice',
      D: 'You don\'t write much — you prefer numbers or talking',
    }
  },
  {
    text: 'When plans change suddenly, you typically:',
    options: {
      A: 'Feel stressed — you prefer having a clear plan',
      B: 'Adapt easily — chaos doesn\'t really bother you',
      C: 'Get excited — changes often lead to better outcomes',
      D: 'Analyze first — assess the new situation before acting',
    }
  },
  {
    text: 'Which subject did you find EASIEST (not just fun)?',
    options: {
      A: 'Math / Science — logic and problem-solving',
      B: 'Languages / Literature — reading and writing',
      C: 'Art / Creative subjects — expression and creativity',
      D: 'Social Studies / Economics — systems and society',
    }
  },
  // ── Part 3 ──
  {
    text: 'Your ideal work intensity level:',
    options: {
      A: 'High — 60–80+ hrs/week is fine if I\'m learning fast',
      B: 'Moderate-high — 50–60 hrs, intense but sustainable',
      C: 'Moderate — 45–50 hrs, good work-life balance',
      D: 'Balanced — 40–45 hrs, clear boundaries',
    }
  },
  {
    text: 'You prefer to measure your success by:',
    options: {
      A: 'Revenue / numbers — direct measurable impact',
      B: 'Quality of work — excellence in analysis or craft',
      C: 'Things built — products, content, systems created',
      D: 'Relationships — people helped, connections made',
    }
  },
  {
    text: 'In terms of role definition, you prefer:',
    options: {
      A: 'Clear specialization — become an expert in one area',
      B: 'Structured variety — defined role with cross-exposure',
      C: 'Fluid generalist — do whatever needs to be done',
      D: 'Project-based — different projects, different challenges',
    }
  },
  {
    text: 'Your relationship with rejection and failure:',
    options: {
      A: 'It bothers me — I prefer success I can control',
      B: 'Handle occasionally — constant rejection drains me',
      C: 'Resilient — rejection is feedback, I bounce back quick',
      D: 'Thrive on it — converting \'no\' to \'yes\' excites me',
    }
  },
  {
    text: 'Your preferred type of impact:',
    options: {
      A: 'Immediate — see results in days or weeks',
      B: 'Medium-term — projects show impact over months',
      C: 'Long-term strategic — shapes direction over years',
      D: 'Cumulative — small wins that compound over time',
    }
  },
  // ── Part 4 ──
  {
    text: 'What do you do in free time that could be a job skill?',
    options: {
      A: 'Analyze things — stocks, stats, trends, data',
      B: 'Create content — writing, videos, design, social media',
      C: 'Sell or negotiate — deals, convincing friends',
      D: 'Organize / plan — events, trips, projects',
    }
  },
  {
    text: 'When scrolling social media, you mostly engage with:',
    options: {
      A: 'Educational / informative — learning something new',
      B: 'Business / finance — startups, markets, money',
      C: 'Creative content — memes, trends, aesthetic posts',
      D: 'Self-improvement — productivity, growth content',
    }
  },
  {
    text: 'Which of these have you actually done (or genuinely tried)?',
    options: {
      A: 'Started a business, sold something, or freelanced',
      B: 'Created online content — blog, YouTube, Instagram',
      C: 'Analyzed data or built models for fun / college projects',
      D: 'Organized events, led clubs, or managed projects',
    }
  },
  {
    text: 'People come to you for advice on:',
    options: {
      A: 'Career / money — practical, strategic thinking',
      B: 'Creative work — writing, design, content feedback',
      C: 'People problems — handling situations, convincing others',
      D: 'Getting things done — planning, organizing, executing',
    }
  },
  {
    text: 'Your honest biggest weakness is:',
    options: {
      A: 'Impatience — want results fast, frustrated by slow pace',
      B: 'Overthinking — analyze too much, struggle to act',
      C: 'Inconsistency — great in bursts, struggle with routine',
      D: 'Risk aversion — prefer safe, proven paths',
    }
  },
];

// ─── SCORING RULES ───────────────────────────────────────────────────────────
// Each rule: { q: questionIndex (0-based), a: answer letter }
// When answers[q] === a, domain scores +10 points.
// Note: Finance has a deliberate duplicate {q:5, a:'A'} matching the Excel source.

const SCORING_RULES = {
  'Consulting & Strategy': [
    { q: 0, a: 'A' }, { q: 1, a: 'A' }, { q: 2, a: 'A' }, { q: 3, a: 'A' }, { q: 4, a: 'A' },
    { q: 6, a: 'D' }, { q: 9, a: 'A' }, { q: 10, a: 'A' }, { q: 14, a: 'C' }, { q: 18, a: 'A' },
  ],
  'Sales & Business Dev': [
    { q: 0, a: 'B' }, { q: 1, a: 'B' }, { q: 2, a: 'B' }, { q: 3, a: 'B' }, { q: 4, a: 'B' },
    { q: 6, a: 'C' }, { q: 11, a: 'A' }, { q: 13, a: 'D' }, { q: 15, a: 'C' }, { q: 18, a: 'C' },
  ],
  'Marketing & Branding': [
    { q: 0, a: 'C' }, { q: 1, a: 'C' }, { q: 2, a: 'C' }, { q: 3, a: 'C' }, { q: 4, a: 'C' },
    { q: 7, a: 'B' }, { q: 7, a: 'C' }, { q: 11, a: 'C' }, { q: 16, a: 'C' }, { q: 19, a: 'C' },
  ],
  'Content & Creative': [
    { q: 0, a: 'C' }, { q: 1, a: 'C' }, { q: 2, a: 'C' }, { q: 4, a: 'C' }, { q: 7, a: 'C' },
    { q: 9, a: 'B' }, { q: 9, a: 'C' }, { q: 11, a: 'C' }, { q: 15, a: 'B' }, { q: 18, a: 'B' },
  ],
  'Finance & Investment': [
    { q: 3, a: 'A' }, { q: 5, a: 'A' }, { q: 9, a: 'A' }, { q: 10, a: 'A' }, { q: 12, a: 'A' },
    { q: 5, a: 'A' }, // intentional duplicate mirroring Excel formula
    { q: 15, a: 'A' }, { q: 16, a: 'B' }, { q: 17, a: 'C' }, { q: 19, a: 'B' },
  ],
  'Product & Growth': [
    { q: 1, a: 'A' }, { q: 2, a: 'A' }, { q: 8, a: 'C' }, { q: 10, a: 'B' }, { q: 12, a: 'D' },
    { q: 3, a: 'B' }, { q: 6, a: 'B' }, { q: 11, a: 'C' }, { q: 14, a: 'B' }, { q: 17, a: 'D' },
  ],
  'Operations & Process': [
    { q: 0, a: 'D' }, { q: 1, a: 'D' }, { q: 2, a: 'D' }, { q: 3, a: 'D' }, { q: 4, a: 'D' },
    { q: 8, a: 'A' }, { q: 10, a: 'D' }, { q: 13, a: 'A' }, { q: 15, a: 'D' }, { q: 19, a: 'D' },
  ],
  "Founder's Office": [
    { q: 0, a: 'D' }, { q: 3, a: 'B' }, { q: 4, a: 'B' }, { q: 8, a: 'B' }, { q: 8, a: 'C' },
    { q: 10, a: 'C' }, { q: 12, a: 'C' }, { q: 15, a: 'D' }, { q: 17, a: 'A' }, { q: 19, a: 'C' },
  ],
  'Data & Analytics': [
    { q: 5, a: 'A' }, { q: 9, a: 'A' }, { q: 9, a: 'D' }, { q: 10, a: 'D' }, { q: 12, a: 'A' },
    { q: 15, a: 'A' }, { q: 16, a: 'A' }, { q: 17, a: 'C' }, { q: 13, a: 'A' }, { q: 19, a: 'D' },
  ],
  'Human Resources': [
    { q: 0, a: 'B' }, { q: 6, a: 'C' }, { q: 7, a: 'B' }, { q: 10, a: 'D' }, { q: 11, a: 'D' },
    { q: 5, a: 'C' }, { q: 13, a: 'B' }, { q: 16, a: 'D' }, { q: 17, a: 'D' }, { q: 18, a: 'C' },
  ],
};

// ─── DOMAIN META ─────────────────────────────────────────────────────────────

const DOMAIN_META = {
  'Consulting & Strategy': {
    colorClass: 'dh-consulting',
    emoji: '🔍',
    tagline: 'Problem-solvers who love frameworks & rigorous analysis',
    bestFor: 'You thrive analyzing ambiguous problems, building structured recommendations, and working across multiple industries. You\'re okay with long hours in exchange for accelerated learning.',
    roles: [
      { name: 'Management Consulting Intern', stipend: '₹10K–85K/month', hours: '50–70 hrs/week', access: 'Highly Competitive', about: 'Research, model building, client presentations, structured problem-solving.' },
      { name: 'Research Analyst Intern', stipend: '₹10K–45K/month', hours: '45–55 hrs/week', access: 'Moderate', about: 'Deep dives into industries, market sizing models, research reports.' },
      { name: 'Strategy Associate Intern', stipend: '₹15K–50K/month', hours: '50–60 hrs/week', access: 'Competitive', about: 'Business cases, competitive intelligence, board-level presentations.' },
    ],
    topSkills: ['Excel Modelling', 'PowerPoint Storytelling', 'Case Frameworks', 'Structured Thinking', 'Market Research', 'Data Synthesis'],
  },
  'Sales & Business Dev': {
    colorClass: 'dh-sales',
    emoji: '🤝',
    tagline: 'Relationship builders who thrive on targets & rejection',
    bestFor: 'You love meeting new people, are energized by saying \'yes\' and converting \'no\', and want a role where your effort directly equals results. Thick skin and competitive spirit are your superpowers.',
    roles: [
      { name: 'B2B Sales / SDR Intern', stipend: '₹8K–35K + incentives', hours: '45–55 hrs/week', access: 'Accessible', about: '30–50 cold calls/day, booking demo meetings, CRM updates.' },
      { name: 'Business Development Intern', stipend: '₹8K–40K + incentives', hours: '45–55 hrs/week', access: 'Accessible', about: 'Lead research, outreach, pipeline management, proposals.' },
      { name: 'Partnerships & Alliances Intern', stipend: '₹10K–35K/month', hours: '40–50 hrs/week', access: 'Moderate', about: 'Strategic partnership identification, deal structuring, relationship management.' },
    ],
    topSkills: ['CRM (Salesforce/HubSpot)', 'Cold Calling', 'Email Outreach', 'LinkedIn Sales Navigator', 'Negotiation', 'Pipeline Management'],
  },
  'Marketing & Branding': {
    colorClass: 'dh-marketing',
    emoji: '📣',
    tagline: 'Creatives with business sense who build brand equity',
    bestFor: 'You\'re always online, understand what goes viral, love creating content, and can balance creative instincts with data. You want to build brands — not just run campaigns.',
    roles: [
      { name: 'Social Media Marketing Intern', stipend: '₹8K–25K/month', hours: '40–50 hrs/week', access: 'Accessible', about: 'Daily posts, engagement, content calendars, basic paid promotions.' },
      { name: 'Performance Marketing Intern', stipend: '₹15K–30K/month', hours: '45–55 hrs/week', access: 'Moderate', about: 'Google & Meta Ads, A/B testing, ROI tracking, campaign optimization.' },
      { name: 'Influencer Marketing Intern', stipend: '₹10K–30K/month', hours: '40–50 hrs/week', access: 'Accessible', about: 'Influencer discovery, outreach, negotiation, campaign reporting.' },
      { name: 'Brand Marketing Intern', stipend: '₹10K–40K/month', hours: '45–55 hrs/week', access: 'Moderate', about: 'Brand campaigns, consumer research, creative briefs, agency coordination.' },
    ],
    topSkills: ['Content Writing', 'Canva / Design Tools', 'Meta & Google Ads', 'Analytics (GA4)', 'Brand Strategy', 'Consumer Research'],
  },
  'Content & Creative': {
    colorClass: 'dh-content',
    emoji: '✍️',
    tagline: 'Storytellers and visual creators who craft brand narratives',
    bestFor: 'You think in stories, have a distinctive voice, and create content people actually want to consume. You enjoy the craft of writing, design, or video — and you\'re okay with subjective feedback.',
    roles: [
      { name: 'Content Writer Intern', stipend: '₹10K–30K/month', hours: '40–50 hrs/week', access: 'Accessible', about: 'Blogs, articles, SEO content, email newsletters.' },
      { name: 'Copywriter Intern', stipend: '₹10K–30K/month', hours: '40–50 hrs/week', access: 'Accessible', about: 'Ad copy, landing pages, product descriptions, brand voice.' },
      { name: 'Content Strategist Intern', stipend: '₹12K–35K/month', hours: '40–50 hrs/week', access: 'Moderate', about: 'Content calendars, SEO strategy, performance analysis.' },
      { name: 'Graphic Design Intern', stipend: '₹8K–25K/month', hours: '40–50 hrs/week', access: 'Accessible', about: 'Visuals for social, marketing collateral, brand assets.' },
    ],
    topSkills: ['Creative Writing', 'SEO Basics', 'Canva / Figma', 'Content Strategy', 'Copywriting Frameworks', 'Analytics'],
  },
  'Finance & Investment': {
    colorClass: 'dh-finance',
    emoji: '📊',
    tagline: 'Numbers-first thinkers who thrive in high-pressure finance roles',
    bestFor: 'You love financial models, are energized by quantitative analysis, and want direct exposure to capital allocation or corporate finance. You\'re okay with grueling hours in exchange for top-tier exit opportunities.',
    roles: [
      { name: 'FP&A Intern', stipend: '₹15K–45K/month', hours: '45–55 hrs/week', access: 'Moderate', about: 'Budgeting, forecasting, variance analysis, financial reporting.' },
      { name: 'Investment Banking Analyst Intern', stipend: '₹25K–80K/month', hours: '70–100 hrs/week', access: 'Highly Competitive', about: 'Deal modelling, pitchbooks, M&A / IPO support, due diligence.' },
      { name: 'Equity Research Intern', stipend: '₹10K–65K/month', hours: '50–65 hrs/week', access: 'Competitive', about: 'Company models, industry reports, earnings analysis, investment theses.' },
      { name: 'Venture Capital Analyst Intern', stipend: '₹20K–50K/month', hours: '50–60 hrs/week', access: 'Competitive', about: 'Deal sourcing, startup evaluation, portfolio monitoring, market research.' },
    ],
    topSkills: ['Excel Financial Modelling', 'Accounting Fundamentals', 'Valuation Techniques', 'Bloomberg / FactSet', 'PowerPoint Pitchbooks', 'Financial Statement Analysis'],
  },
  'Product & Growth': {
    colorClass: 'dh-product',
    emoji: '🚀',
    tagline: 'Data-driven builders obsessed with user problems and metrics',
    bestFor: 'You obsess over why products succeed, think in terms of user problems not features, and want to work at the intersection of engineering, design, and business. Metrics and experimentation excite you.',
    roles: [
      { name: 'Product Management Intern', stipend: '₹25K–50K/month', hours: '45–55 hrs/week', access: 'Highly Competitive', about: 'User research, PRD writing, sprint planning, A/B testing, roadmap prioritization.' },
      { name: 'Growth Marketing / Hacker Intern', stipend: '₹15K–45K/month', hours: '45–55 hrs/week', access: 'Moderate', about: 'AARRR funnel analysis, growth experiments, dashboards, referral mechanics.' },
      { name: 'User Research Intern', stipend: '₹10K–35K/month', hours: '40–50 hrs/week', access: 'Moderate', about: 'User interviews, usability tests, persona creation, insight synthesis.' },
    ],
    topSkills: ['Product Analytics (Mixpanel/Amplitude)', 'SQL', 'User Research Methods', 'A/B Testing', 'Wireframing (Figma)', 'Growth Frameworks (AARRR)'],
  },
  'Operations & Process': {
    colorClass: 'dh-operations',
    emoji: '⚙️',
    tagline: 'System builders who turn chaos into efficient, scalable processes',
    bestFor: 'You\'re energized by making messy processes clean and efficient. You love checklists, timelines, and SOPs. You want to be the person who makes the machine run smoothly.',
    roles: [
      { name: 'Business Operations Intern', stipend: '₹10K–35K/month', hours: '45–55 hrs/week', access: 'Accessible', about: 'Process mapping, reporting dashboards, cross-functional coordination.' },
      { name: 'Supply Chain Intern', stipend: '₹10K–40K/month', hours: '45–55 hrs/week', access: 'Moderate', about: 'Vendor management, logistics tracking, inventory analysis, cost optimization.' },
      { name: 'Program Manager Intern', stipend: '₹15K–40K/month', hours: '45–55 hrs/week', access: 'Moderate', about: 'Project planning, stakeholder coordination, timeline management, status reporting.' },
    ],
    topSkills: ['Excel / Google Sheets', 'Process Mapping', 'Project Management Tools', 'Data Analysis', 'Stakeholder Communication', 'SOP Documentation'],
  },
  "Founder's Office": {
    colorClass: 'dh-founders',
    emoji: '🏢',
    tagline: 'Generalist operators who want to learn everything at a startup',
    bestFor: 'You\'re adaptable, execution-oriented, and want exposure to all parts of a business. You\'re comfortable with ambiguity, don\'t need a defined role, and want to work closely with leadership from day one.',
    roles: [
      { name: "Founder's Office Intern", stipend: '₹10K–75K/month', hours: '50–70 hrs/week', access: 'Competitive', about: 'Cross-functional execution, special projects, leadership support, zero-to-one work.' },
      { name: 'Chief of Staff / EA Intern', stipend: '₹10K–50K/month', hours: '50–60 hrs/week', access: 'Moderate', about: 'Meeting prep, stakeholder management, key project execution, reporting.' },
    ],
    topSkills: ['Execution Bias', 'Communication', 'Excel / Sheets', 'Project Coordination', 'Problem-Solving', 'Adaptability'],
  },
  'Data & Analytics': {
    colorClass: 'dh-data',
    emoji: '📈',
    tagline: 'Quantitative thinkers who find patterns where others see noise',
    bestFor: 'You love working with numbers, finding patterns in data, and turning raw information into business decisions. You\'re comfortable with SQL, Excel models, and data visualization.',
    roles: [
      { name: 'Business Analyst Intern', stipend: '₹15K–45K/month', hours: '45–55 hrs/week', access: 'Competitive', about: 'Requirements gathering, process analysis, dashboards, stakeholder reporting.' },
      { name: 'Data Analyst / BI Intern', stipend: '₹12K–35K/month', hours: '40–50 hrs/week', access: 'Competitive', about: 'SQL queries, dashboards (Tableau/Power BI), cohort analysis, data cleaning.' },
    ],
    topSkills: ['SQL', 'Excel (Pivot Tables, VLOOKUP)', 'Tableau / Power BI', 'Python (Pandas)', 'Statistical Analysis', 'Data Storytelling'],
  },
  'Human Resources': {
    colorClass: 'dh-hr',
    emoji: '👥',
    tagline: 'People-first professionals who build and nurture great teams',
    bestFor: 'You\'re genuinely curious about people, have high emotional intelligence, and want to shape company culture. You\'re energized by helping others and find satisfaction in building communities.',
    roles: [
      { name: 'Talent Acquisition Intern', stipend: '₹10K–30K/month', hours: '40–50 hrs/week', access: 'Accessible', about: 'Sourcing, screening, interview coordination, candidate experience.' },
      { name: 'Employer Branding Intern', stipend: '₹10K–25K/month', hours: '40–45 hrs/week', access: 'Accessible', about: 'EVP content, LinkedIn presence, campus branding, culture communication.' },
      { name: 'HR Generalist Intern', stipend: '₹10K–25K/month', hours: '40–50 hrs/week', access: 'Accessible', about: 'Onboarding, policy support, employee engagement, HRIS management.' },
    ],
    topSkills: ['Communication & Empathy', 'Excel / HRIS Tools', 'Talent Sourcing (LinkedIn)', 'Interviewing Techniques', 'Content Writing (EVP)', 'Relationship Management'],
  },
};

// ─── APP STATE ────────────────────────────────────────────────────────────────

const App = (() => {
  let currentQ = 0;
  const answers = new Array(20).fill(null);

  // User details captured from the form
  let userData = { name: '', email: '', contact: '', course: '', specialisation: '' };

  // ── Helpers ──────────────────────────────────────────────────────────────

  function getPartFor(qIdx) {
    return PARTS.find(p => p.qs.includes(qIdx)) || PARTS[0];
  }

  function calcScores() {
    const scores = {};
    for (const [domain, rules] of Object.entries(SCORING_RULES)) {
      let s = 0;
      for (const rule of rules) {
        if (answers[rule.q] === rule.a) s += 10;
      }
      scores[domain] = s;
    }
    return scores;
  }

  function getRating(score) {
    if (score >= 70) return 'excellent';
    if (score >= 50) return 'good';
    if (score >= 30) return 'moderate';
    return 'low';
  }

  function getRatingLabel(score) {
    if (score >= 70) return 'Excellent';
    if (score >= 50) return 'Good';
    if (score >= 30) return 'Moderate';
    return 'Low';
  }

  function showScreen(id) {
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    document.getElementById(id).classList.add('active');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  // ── Render Question ───────────────────────────────────────────────────────

  function renderQuestion() {
    const q = QUESTIONS[currentQ];
    const part = getPartFor(currentQ);

    document.getElementById('part-label').textContent = part.label;
    document.getElementById('q-counter').textContent = `${currentQ + 1} / 20`;
    document.getElementById('q-number').textContent = `Q${currentQ + 1}`;
    document.getElementById('q-text').textContent = q.text;

    const pct = Math.max(5, Math.round(((currentQ + 1) / 20) * 100));
    document.getElementById('progress-fill').style.width = pct + '%';

    // Options
    const grid = document.getElementById('options-grid');
    grid.innerHTML = '';
    for (const [letter, text] of Object.entries(q.options)) {
      const btn = document.createElement('button');
      btn.className = 'option-btn' + (answers[currentQ] === letter ? ' selected' : '');
      btn.innerHTML = `<span class="option-letter">${letter}</span><span class="option-text">${text}</span>`;
      btn.addEventListener('click', () => selectAnswer(letter));
      grid.appendChild(btn);
    }

    // Nav buttons
    document.getElementById('btn-prev').disabled = currentQ === 0;
    const nextBtn = document.getElementById('btn-next');
    if (currentQ === 19) {
      nextBtn.textContent = 'See My Results →';
      nextBtn.disabled = answers[currentQ] === null;
    } else {
      nextBtn.textContent = 'Next →';
      nextBtn.disabled = answers[currentQ] === null;
    }

    renderDots();
  }

  function renderDots() {
    const container = document.getElementById('nav-dots');
    container.innerHTML = '';
    for (let i = 0; i < 20; i++) {
      const dot = document.createElement('div');
      dot.className = 'nav-dot' +
        (i === currentQ ? ' current' : '') +
        (answers[i] !== null && i !== currentQ ? ' answered' : '');
      dot.title = `Q${i + 1}`;
      dot.addEventListener('click', () => { currentQ = i; renderQuestion(); });
      container.appendChild(dot);
    }
  }

  function selectAnswer(letter) {
    answers[currentQ] = letter;
    renderQuestion();

    // Auto-advance after a short delay
    setTimeout(() => {
      if (currentQ < 19) {
        currentQ++;
        renderQuestion();
      }
    }, 420);
  }

  // ── Render Results ────────────────────────────────────────────────────────

  function renderResults() {
    const scores = calcScores();
    const sorted = Object.entries(scores).sort((a, b) => b[1] - a[1]);
    const top3 = sorted.slice(0, 3);

    // Podium
    const medals = ['🥇', '🥈', '🥉'];
    const rankLabels = ['#1 Best Fit', '#2 Strong Fit', '#3 Good Fit'];
    const podium = document.getElementById('podium-grid');
    podium.innerHTML = top3.map(([domain, score], i) => `
      <div class="podium-item rank-${i + 1}">
        <span class="podium-medal">${medals[i]}</span>
        <div class="podium-rank-label">${rankLabels[i]}</div>
        <div class="podium-domain">${domain}</div>
        <div class="podium-score-badge">${score}/100</div>
        <div class="podium-match-label">${getRatingLabel(score)} Match</div>
      </div>
    `).join('');

    // Scores list
    const list = document.getElementById('scores-list');
    list.innerHTML = sorted.map(([domain, score]) => {
      const rating = getRating(score);
      return `
        <div class="score-row">
          <span class="score-domain">${domain}</span>
          <div class="score-bar-wrap">
            <div class="score-bar ${rating}" data-width="${score}" style="width:0%"></div>
          </div>
          <div class="score-value">
            <span class="sv-num">${score}</span>
            <span class="score-rating ${rating}">${getRatingLabel(score)}</span>
          </div>
        </div>
      `;
    }).join('');

    // Domain detail cards for top 3
    const detailsContainer = document.getElementById('domain-details');
    detailsContainer.innerHTML = '<h2 style="font-size:20px;font-weight:800;color:#0F172A;margin:0 0 16px 0;">Explore Your Top 3 Domains</h2>';
    top3.forEach(([domain, score], i) => {
      detailsContainer.appendChild(buildDomainCard(domain, score, i));
    });

    showScreen('screen-results');

    // Send data to Google Sheets
    sendToSheets(top3);

    // Animate bars after paint
    requestAnimationFrame(() => {
      setTimeout(() => {
        document.querySelectorAll('.score-bar').forEach(bar => {
          bar.style.width = bar.dataset.width + '%';
        });
      }, 200);
    });

    // Auto-open first domain card
    setTimeout(() => {
      const first = detailsContainer.querySelector('.domain-detail-card');
      if (first) first.classList.add('open');
    }, 400);
  }

  function buildDomainCard(domain, score, rank) {
    const meta = DOMAIN_META[domain];
    if (!meta) return document.createElement('div');

    const card = document.createElement('div');
    card.className = 'domain-detail-card';

    const medals = ['🥇', '🥈', '🥉'];
    const rankLabel = ['Best Fit', 'Strong Fit', 'Good Fit'][rank];

    card.innerHTML = `
      <div class="domain-detail-header ${meta.colorClass}" onclick="this.parentElement.classList.toggle('open')">
        <div class="domain-detail-rank">${medals[rank]}</div>
        <div class="domain-detail-title">
          <div class="domain-detail-name">${meta.emoji} ${domain}</div>
          <div class="domain-detail-meta">${rankLabel} · ${score}/100 · ${getRatingLabel(score)} Match</div>
        </div>
        <div class="domain-detail-toggle">▾</div>
      </div>
      <div class="domain-detail-body">
        <div class="best-for-box" style="margin-top:20px">
          <div class="best-for-label">This Domain Is For You If…</div>
          <div class="best-for-text">${meta.bestFor}</div>
        </div>

        <div class="skills-section">
          <div class="skills-title">🛠 Top Skills to Build</div>
          <div class="skills-list">
            ${meta.topSkills.map(s => `<span class="skill-chip">${s}</span>`).join('')}
          </div>
        </div>

        <div class="skills-section">
          <div class="skills-title">💼 Common Internship Roles</div>
          <div class="roles-grid">
            ${meta.roles.map(r => `
              <div class="role-card">
                <div class="role-name">${r.name}</div>
                <div class="role-meta">
                  <span class="role-tag stipend">💰 ${r.stipend}</span>
                  <span class="role-tag hours">🕐 ${r.hours}</span>
                  <span class="role-tag access">📌 ${r.access}</span>
                </div>
                <div class="role-about">${r.about}</div>
              </div>
            `).join('')}
          </div>
        </div>
      </div>
    `;

    return card;
  }

  // ── Google Sheets Submission ──────────────────────────────────────────────

  function sendToSheets(top3) {
    if (!GOOGLE_SCRIPT_URL || GOOGLE_SCRIPT_URL === 'YOUR_GOOGLE_APPS_SCRIPT_URL_HERE') return;

    const banner = document.getElementById('submit-banner');
    if (banner) { banner.className = 'submit-banner sending'; banner.textContent = '⏳ Saving your results…'; banner.style.display = 'flex'; }

    const payload = {
      name: userData.name,
      email: userData.email,
      contact: userData.contact,
      course: userData.course,
      specialisation: userData.specialisation,
      domain1: top3[0][0], score1: top3[0][1],
      domain2: top3[1][0], score2: top3[1][1],
      domain3: top3[2][0], score3: top3[2][1],
      timestamp: new Date().toISOString(),
    };

    fetch(GOOGLE_SCRIPT_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    })
      .then(() => {
        if (banner) { banner.className = 'submit-banner success'; banner.textContent = '✅ Results saved successfully!'; }
      })
      .catch(() => {
        if (banner) { banner.className = 'submit-banner failed'; banner.textContent = '⚠️ Could not save results — check your network.'; }
      });
  }

  // ── Public API ────────────────────────────────────────────────────────────

  function showForm() {
    showScreen('screen-userform');
    // Clear any previous error
    const err = document.getElementById('form-error');
    if (err) err.style.display = 'none';
    // Clear inputs
    ['f-name', 'f-email', 'f-contact', 'f-course', 'f-spec'].forEach(id => {
      const el = document.getElementById(id);
      if (el) { el.value = ''; el.classList.remove('input-error'); }
    });
  }

  function submitUserForm() {
    const fields = [
      { id: 'f-name', key: 'name', label: 'Full Name' },
      { id: 'f-email', key: 'email', label: 'Email Address' },
      { id: 'f-contact', key: 'contact', label: 'Contact Number' },
      { id: 'f-course', key: 'course', label: 'Course' },
      { id: 'f-spec', key: 'specialisation', label: 'Specialisation' },
    ];

    const err = document.getElementById('form-error');
    let missing = [];

    fields.forEach(f => {
      const el = document.getElementById(f.id);
      const val = el ? el.value.trim() : '';
      if (!val) {
        missing.push(f.label);
        if (el) el.classList.add('input-error');
      } else {
        if (el) el.classList.remove('input-error');
        userData[f.key] = val;
      }
    });

    // Basic email check
    const emailEl = document.getElementById('f-email');
    if (emailEl && emailEl.value.trim() && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailEl.value.trim())) {
      missing.push('valid Email Address');
      emailEl.classList.add('input-error');
    }

    if (missing.length) {
      err.textContent = 'Please fill in: ' + missing.join(', ');
      err.style.display = 'block';
      return;
    }

    err.style.display = 'none';
    startQuiz();
  }

  function startQuiz() {
    currentQ = 0;
    answers.fill(null);
    showScreen('screen-quiz');
    renderQuestion();
  }

  function nextQ() {
    if (answers[currentQ] === null) return;
    if (currentQ === 19) {
      renderResults();
    } else {
      currentQ++;
      renderQuestion();
    }
  }

  function prevQ() {
    if (currentQ > 0) {
      currentQ--;
      renderQuestion();
    }
  }

  function restart() {
    showForm();
  }

  return { showForm, submitUserForm, startQuiz, nextQ, prevQ, restart };
})();

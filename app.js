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
    colorClass: 'dh-consulting', emoji: '🔍',
    tagline: 'Problem-solvers who love frameworks & rigorous analysis',
    bestFor: 'You thrive analyzing ambiguous problems, building structured recommendations, and working across multiple industries. You\'re okay with long hours in exchange for accelerated learning.',
    roles: [
      {
        name: 'Management Consulting Intern', stipend: '₹10K–85K/month', duration: '2–3 months (summer)', hours: '50–70 hrs/week', access: 'Highly Competitive',
        dayToDay: ['Conduct secondary research on market size, competitors, industry trends (2–3 hrs/day)', 'Build Excel models for financial analysis and scenario planning', 'Create PowerPoint presentations with structured storylines for client delivery', 'Synthesize findings from multiple data sources into actionable insights', 'Attend client meetings, take detailed notes, track action items', 'Support senior consultants on workstream deliverables', 'Conduct primary research: expert interviews, surveys when needed', 'Iterate on analyses based on manager and partner feedback'],
        hardSkills: [{rank:1,skill:'Microsoft Excel (Pivot tables, VLOOKUP, INDEX-MATCH, modeling)',level:'Advanced'},{rank:2,skill:'Microsoft PowerPoint (Story-driven decks, clean design)',level:'Advanced'},{rank:3,skill:'Market Research (Primary + secondary methods)',level:'Intermediate'},{rank:4,skill:'Case Problem-Solving (Frameworks: profitability, market entry)',level:'Intermediate'},{rank:5,skill:'Data Visualization',level:'Basic'},{rank:6,skill:'Financial Statement Analysis',level:'Basic'},{rank:7,skill:'Basic Statistics',level:'Basic'}],
        softSkills: ['Structured thinking (breaking down ambiguity into components)', 'Intellectual curiosity (learning new industries every month)', 'Grace under pressure (tight deadlines are the norm)', 'Perfectionism (every number, every slide must be flawless)', 'Client communication (professional, concise, confident)'],
        degreeAlignmentPrimary: 'MBA, Economics, Engineering (from top institutes), Commerce',
        degreeAlignmentSecondary: 'HARD TRUTH: MBB and Big 4 Strategy heavily recruit from IIMs, IITs, SRCC, LSR. BUT boutique firms (Praxis, Avalon, Redseer) are more accessible. Any degree can work if you: crack case interviews, demonstrate structured thinking, have strong academics. Liberal arts students: your research and writing skills are valuable. Start with smaller firms, build experience.',
        isThisForYou: 'This role is for you if: You enjoy breaking down complex, ambiguous problems. You want exposure to multiple industries quickly. You thrive under pressure and tight deadlines. You\'re okay with 60–70 hour weeks for accelerated learning. You want strong exit opportunities (strategy roles, startups, MBA).',
      },
      {
        name: 'Research Analyst Intern', stipend: '₹10K–45K/month', duration: '2–4 months', hours: '45–55 hrs/week', access: 'Moderate',
        dayToDay: ['Conduct deep secondary research on assigned industries/topics', 'Analyze market data, competitive landscape, and industry trends', 'Build Excel models for market sizing and forecasting', 'Create research reports and presentations', 'Maintain databases of companies, market data, and research sources', 'Conduct expert interviews to validate hypotheses', 'Support client deliverables with data and analysis', 'Stay updated on industry news and developments'],
        hardSkills: [{rank:1,skill:'Research Methods (Secondary research, database usage)',level:'Advanced'},{rank:2,skill:'Excel (Analysis, modeling, data manipulation)',level:'Intermediate'},{rank:3,skill:'PowerPoint (Research presentations)',level:'Intermediate'},{rank:4,skill:'Industry Databases (Statista, IBISWorld, Bloomberg)',level:'Developing'},{rank:5,skill:'Data Visualization',level:'Basic'},{rank:6,skill:'Basic Statistical Analysis',level:'Basic'},{rank:7,skill:'Report Writing',level:'Intermediate'}],
        softSkills: ['Intellectual curiosity (genuine interest in understanding industries)', 'Attention to detail (research must be accurate and sourced)', 'Synthesis ability (turning data into insights)', 'Written communication (clear, concise reports)', 'Self-motivation (research often requires independent work)'],
        degreeAlignmentPrimary: 'Economics, Statistics, Mathematics, Engineering',
        degreeAlignmentSecondary: 'Research roles value analytical thinking. Liberal arts students: your research and writing skills from papers/theses are directly applicable. B.Com students: financial analysis is helpful. Show any research experience — academic papers, thesis work, even well-researched blog posts.',
        isThisForYou: 'This role is for you if: You enjoy deep-diving into topics and becoming an expert. You like finding patterns in data. You can synthesize multiple sources into clear narratives. You prefer depth over breadth. You want to influence decisions through rigorous analysis.',
      },
      {
        name: 'Strategy Associate Intern', stipend: '₹15K–50K/month', duration: '3–6 months', hours: '50–60 hrs/week', access: 'Competitive',
        dayToDay: ['Analyze market trends and competitive landscape for strategic planning', 'Build business cases for new initiatives, market entry, or investments', 'Support M&A due diligence with research and analysis', 'Create board-level presentations and strategic memos', 'Conduct financial modeling for strategic decisions', 'Coordinate cross-functional inputs for strategy projects', 'Track competitive intelligence and market developments', 'Present findings to senior leadership'],
        hardSkills: [{rank:1,skill:'Financial Modeling (DCF basics, scenario analysis)',level:'Intermediate'},{rank:2,skill:'Advanced Excel (Modeling, sensitivity tables)',level:'Advanced'},{rank:3,skill:'Strategic Frameworks (Porter\'s 5 Forces, SWOT, etc.)',level:'Intermediate'},{rank:4,skill:'PowerPoint (Executive presentations)',level:'Advanced'},{rank:5,skill:'Market Sizing',level:'Intermediate'},{rank:6,skill:'M&A Concepts (basic)',level:'Basic'},{rank:7,skill:'SQL (basic data extraction)',level:'Basic'}],
        softSkills: ['Strategic thinking (connecting dots to see big picture)', 'Executive communication (presenting to C-suite)', 'Analytical rigor (recommendations backed by data)', 'Cross-functional collaboration (working with multiple teams)', 'Comfort with ambiguity (strategy is never black and white)'],
        degreeAlignmentPrimary: 'MBA, Economics, Engineering, CA/CFA',
        degreeAlignmentSecondary: 'Strategy roles prefer analytical backgrounds. B.Com/BBA students: strengthen your financial modeling skills. Liberal arts: possible but need to demonstrate quantitative ability. Prior consulting or finance internships help significantly.',
        isThisForYou: 'This role is for you if: You think about the \'why\' behind business decisions. You enjoy connecting industry trends to company strategy. You want to influence high-level decisions. You can communicate complex analysis simply. You see yourself in corporate strategy or private equity long-term.',
      },
    ],
    topSkills: ['Excel Modelling', 'PowerPoint Storytelling', 'Case Frameworks', 'Structured Thinking', 'Market Research', 'Data Synthesis'],
  },

  'Sales & Business Dev': {
    colorClass: 'dh-sales', emoji: '🤝',
    tagline: 'Relationship builders who thrive on targets & rejection',
    bestFor: 'You love meeting new people, are energized by saying \'yes\' and converting \'no\', and want a role where your effort directly equals results. Thick skin and competitive spirit are your superpowers.',
    roles: [
      {
        name: 'B2B Sales / SDR Intern', stipend: '₹8K–35K/month + incentives', duration: '2–4 months', hours: '45–55 hrs/week', access: 'Accessible',
        dayToDay: ['Make 30–50 cold calls daily to potential business clients', 'Send 20–30 personalized cold emails to decision makers', 'Research and qualify leads using LinkedIn, databases, company websites', 'Book 5–10 demo/discovery meetings per week for senior sales team', 'Update CRM (Salesforce/HubSpot) with every interaction and status', 'Follow up with prospects who showed interest but didn\'t convert', 'Create personalized pitch decks for specific accounts', 'Track your metrics: calls made, emails sent, meetings booked, conversion rates'],
        hardSkills: [{rank:1,skill:'CRM Tools (Salesforce, HubSpot, Zoho)',level:'Intermediate'},{rank:2,skill:'Cold Calling Techniques',level:'Must develop on job'},{rank:3,skill:'Email Outreach (Personalization, sequences)',level:'Intermediate'},{rank:4,skill:'LinkedIn Sales Navigator',level:'Basic'},{rank:5,skill:'Lead Research & Qualification',level:'Intermediate'},{rank:6,skill:'PowerPoint (Pitch decks)',level:'Basic'},{rank:7,skill:'Basic Product/Industry Knowledge',level:'Ongoing'}],
        softSkills: ['Thick skin for rejection (95% of calls/emails get rejected — this is normal)', 'Persistence without being annoying', 'Active listening (understanding prospect pain points)', 'Energy and enthusiasm (people buy from people they like)', 'Competitiveness (you\'ll be measured against targets constantly)', 'Resilience (bad days are frequent — bounce back quickly)'],
        degreeAlignmentPrimary: 'Any degree — results matter, not credentials',
        degreeAlignmentSecondary: 'This is the great equalizer. BBA, B.Com, B.Tech, BA, B.Sc — nobody cares. What matters: Can you communicate? Can you handle rejection? Can you hit targets? Show any sales experience: college fest sponsorships, freelance sales, even selling products for family business. Attitude > Academics.',
        isThisForYou: 'This role is for you if: You\'re not afraid of rejection — you see it as a numbers game. You\'re energized by talking to strangers. You\'re competitive and target-driven. You want direct visibility into your impact (revenue). You\'re okay with variable pay tied to performance.',
      },
      {
        name: 'Business Development Intern', stipend: '₹8K–40K/month + incentives', duration: '2–4 months', hours: '45–55 hrs/week', access: 'Accessible',
        dayToDay: ['Research and identify 15–20 potential clients/partners weekly', 'Conduct outreach via email, LinkedIn, and warm introductions', 'Qualify leads: understand needs, budget, timeline, decision-makers', 'Schedule and prepare for discovery/demo meetings', 'Maintain CRM with accurate pipeline updates and notes', 'Support proposal and pitch deck preparation', 'Track BD metrics: outreach volume, response rate, meetings booked', 'Attend networking events and conferences when possible'],
        hardSkills: [{rank:1,skill:'Lead Research (LinkedIn, Crunchbase, databases)',level:'Intermediate'},{rank:2,skill:'CRM Management (Salesforce, HubSpot)',level:'Intermediate'},{rank:3,skill:'Cold Outreach (Email writing, LinkedIn messages)',level:'Intermediate'},{rank:4,skill:'PowerPoint (Proposals, pitch decks)',level:'Basic'},{rank:5,skill:'Excel (Pipeline tracking, analysis)',level:'Basic'},{rank:6,skill:'Video Call Tools (Zoom, Google Meet)',level:'Basic'},{rank:7,skill:'Basic Market Research',level:'Basic'}],
        softSkills: ['Communication (clear, concise, professional)', 'Relationship building (BD is about trust, not transactions)', 'Persistence (most outreach gets ignored)', 'Curiosity (understanding prospect\'s business deeply)', 'Organization (juggling multiple prospects at different stages)'],
        degreeAlignmentPrimary: 'BBA, B.Com, MBA, Any business degree',
        degreeAlignmentSecondary: 'ALL degrees welcome. This role rewards hustle over pedigree. B.Tech students: your structured thinking helps in qualification. Arts students: your communication skills are valuable. Show any entrepreneurial or networking experience. College sponsorships, event organizing, freelancing — all count.',
        isThisForYou: 'This role is for you if: You enjoy meeting new people and building relationships. You\'re good at understanding what others need. You can articulate value propositions clearly. Rejection doesn\'t demotivate you. You want a role where effort directly equals results.',
      },
      {
        name: 'Partnerships & Alliances Intern', stipend: '₹10K–35K/month', duration: '3–6 months', hours: '40–50 hrs/week', access: 'Moderate',
        dayToDay: ['Research potential partnership opportunities (co-marketing, integrations, channel)', 'Reach out to potential partners via email and LinkedIn', 'Coordinate meetings between internal teams and partner organizations', 'Support partnership deal structuring and documentation', 'Track partnership performance metrics: revenue share, leads generated', 'Maintain partner relationship through regular check-ins', 'Create partnership proposals and pitch materials', 'Document partnership processes and playbooks'],
        hardSkills: [{rank:1,skill:'Partner Research (Identifying strategic fit)',level:'Intermediate'},{rank:2,skill:'PowerPoint (Partnership proposals)',level:'Intermediate'},{rank:3,skill:'Excel (Partnership modeling, tracking)',level:'Basic'},{rank:4,skill:'Contract Basics (Understanding term sheets)',level:'Basic'},{rank:5,skill:'CRM/Partner Tracking Tools',level:'Basic'},{rank:6,skill:'Market Research',level:'Basic'},{rank:7,skill:'Basic Financial Understanding',level:'Basic'}],
        softSkills: ['Strategic thinking (identifying win-win opportunities)', 'Relationship management (long-term partnerships over quick wins)', 'Negotiation (finding mutually beneficial terms)', 'Cross-functional coordination (working with legal, product, marketing)', 'Patience (partnership deals take time)'],
        degreeAlignmentPrimary: 'BBA, B.Com, MBA, Economics',
        degreeAlignmentSecondary: 'Any degree with business exposure works. B.Tech students: your analytical skills help in partnership modeling. Show any experience managing external stakeholders or coordinating between organizations. More strategic than pure sales.',
        isThisForYou: 'This role is for you if: You think about win-win scenarios, not just closing deals. You\'re good at building long-term relationships. You can navigate complex stakeholder dynamics. You enjoy strategic planning over tactical execution. You want to work on deals that transform businesses.',
      },
    ],
    topSkills: ['CRM (Salesforce/HubSpot)', 'Cold Calling', 'Email Outreach', 'LinkedIn Sales Navigator', 'Negotiation', 'Pipeline Management'],
  },

  'Marketing & Branding': {
    colorClass: 'dh-marketing', emoji: '📣',
    tagline: 'Creatives with business sense who build brand equity',
    bestFor: 'You\'re always online, understand what goes viral, love creating content, and can balance creative instincts with data. You want to build brands — not just run campaigns.',
    roles: [
      {
        name: 'Social Media Marketing Intern', stipend: '₹8K–25K/month', duration: '2–3 months', hours: '40–50 hrs/week', access: 'Accessible',
        dayToDay: ['Create and schedule 3–5 social media posts daily across Instagram, LinkedIn, Twitter', 'Write engaging captions with relevant hashtags and CTAs', 'Respond to comments, DMs, and mentions within 2–4 hours', 'Track daily/weekly metrics: follower growth, engagement rate, reach', 'Research trending topics, hashtags, and competitor content every morning', 'Coordinate with designers for visual assets and creatives', 'Run basic paid promotions and boost posts', 'Create monthly content calendars and performance reports'],
        hardSkills: [{rank:1,skill:'Social Media Platforms (Instagram, LinkedIn, Twitter)',level:'Expert level'},{rank:2,skill:'Content Writing (Captions, hooks, CTAs)',level:'Advanced'},{rank:3,skill:'Canva/Basic Design Tools',level:'Intermediate'},{rank:4,skill:'Analytics Tools (Meta Business Suite, native analytics)',level:'Intermediate'},{rank:5,skill:'Hashtag Research & Trend Analysis',level:'Intermediate'},{rank:6,skill:'Basic Paid Promotion (Boosting posts)',level:'Basic'},{rank:7,skill:'Content Calendar Management',level:'Basic'}],
        softSkills: ['Creativity and trend-awareness (you need to spot viral trends early)', 'Quick wit for real-time engagement and replies', 'Consistency (posting daily is non-negotiable)', 'Thick skin for negative comments and trolls', 'Adaptability to platform algorithm changes'],
        degreeAlignmentPrimary: 'Mass Communication, Journalism, BBA (Marketing)',
        degreeAlignmentSecondary: 'ANY degree accepted — this is one of the most accessible roles. B.Tech? Show your personal social media growth. Arts/Humanities? Your writing skills are an advantage. Commerce? Leverage your business understanding. Build a portfolio of 2–3 accounts you\'ve managed (even personal).',
        isThisForYou: 'This role is for you if: You\'re always on Instagram/Twitter and understand what goes viral. You enjoy creating content and don\'t mind the grind of daily posting. You can write catchy one-liners. You\'re comfortable being the \'brand voice\' online. You check your engagement metrics like others check their messages.',
      },
      {
        name: 'Performance Marketing Intern', stipend: '₹15K–30K/month', duration: '2–4 months', hours: '45–55 hrs/week', access: 'Moderate',
        dayToDay: ['Set up and manage paid ad campaigns on Google Ads and Meta Ads Manager', 'Monitor campaign performance metrics daily: CPC, CTR, ROAS, CPA, CPL', 'A/B test ad creatives, copy, audiences, and landing pages (run 3–5 tests/week)', 'Adjust bids, budgets, and targeting based on real-time data', 'Build weekly performance reports with insights and optimization recommendations', 'Research keywords, audience segments, and competitor ad strategies', 'Coordinate with designers for ad creatives and landing page updates', 'Track attribution and conversion paths across channels'],
        hardSkills: [{rank:1,skill:'Google Ads (Search, Display, YouTube)',level:'Advanced'},{rank:2,skill:'Meta Ads Manager (Facebook, Instagram)',level:'Advanced'},{rank:3,skill:'Google Analytics 4',level:'Intermediate'},{rank:4,skill:'Excel/Sheets (Pivot tables, VLOOKUP, reporting)',level:'Intermediate'},{rank:5,skill:'A/B Testing Methodology',level:'Intermediate'},{rank:6,skill:'UTM Tracking & Attribution',level:'Basic'},{rank:7,skill:'Landing Page Optimization (CRO basics)',level:'Basic'}],
        softSkills: ['Analytical mindset (you\'ll live in spreadsheets and dashboards)', 'Attention to detail (one wrong decimal in budget = disaster)', 'Quick decision-making (campaigns need real-time optimization)', 'Comfort with spending other people\'s money', 'Experimentation mindset (most tests will fail, and that\'s okay)'],
        degreeAlignmentPrimary: 'BBA, B.Com, Economics, Engineering (any branch)',
        degreeAlignmentSecondary: 'ANY degree works here — skills matter most. Get Google Ads and Meta Blueprint certifications (both free). Run a small campaign for any local business or personal project. Even ₹500 of ad spend experience beats no experience. Arts/Humanities students: your consumer psychology understanding is valuable.',
        isThisForYou: 'This role is for you if: You love seeing direct, measurable results (ROI, conversions). Numbers excite you more than creative brainstorming. You enjoy the thrill of optimizing and beating benchmarks. You can make quick decisions with incomplete data. You want marketing skills that are in-demand across every industry.',
      },
      {
        name: 'Influencer Marketing Intern', stipend: '₹10K–30K/month', duration: '2–3 months', hours: '40–50 hrs/week', access: 'Accessible',
        dayToDay: ['Research and identify 20–30 relevant influencers weekly across Instagram, YouTube, LinkedIn', 'Reach out to influencers via DM/email with collaboration proposals', 'Negotiate deliverables, timelines, and compensation with influencers', 'Coordinate content briefs, product shipments, and posting schedules', 'Track influencer content performance: reach, engagement, conversions', 'Maintain influencer database with contact info, rates, and past performance', 'Handle influencer queries and relationship management', 'Create campaign reports with ROI analysis'],
        hardSkills: [{rank:1,skill:'Influencer Discovery Tools (Upfluence, HypeAuditor, manual research)',level:'Intermediate'},{rank:2,skill:'Outreach & Negotiation',level:'Intermediate'},{rank:3,skill:'Excel/Sheets (Database management, tracking)',level:'Intermediate'},{rank:4,skill:'Social Media Analytics',level:'Basic'},{rank:5,skill:'Contract/Brief Writing',level:'Basic'},{rank:6,skill:'Project Coordination',level:'Basic'},{rank:7,skill:'Basic ROI Calculation',level:'Basic'}],
        softSkills: ['Relationship building (influencers are people, not billboards)', 'Negotiation skills (rates are always negotiable)', 'Organization (you\'ll juggle 15–20 influencer relationships simultaneously)', 'Persistence (most outreach gets ignored — keep going)', 'Cultural awareness (understanding what resonates with different audiences)'],
        degreeAlignmentPrimary: 'Mass Communication, BBA, Journalism',
        degreeAlignmentSecondary: 'ANY degree accepted. This role values hustle over academics. If you understand influencer culture and can negotiate, you\'re qualified. B.Tech students: your project management skills transfer well. Show any experience coordinating with multiple stakeholders.',
        isThisForYou: 'This role is for you if: You follow influencers and understand the creator economy. You\'re good at building relationships and don\'t mind cold outreach. You can negotiate without being pushy. You\'re organized enough to manage multiple timelines. You enjoy the mix of creative (content) and operational (coordination) work.',
      },
      {
        name: 'Brand Marketing Intern', stipend: '₹10K–40K/month', duration: '3–6 months', hours: '45–55 hrs/week', access: 'Moderate',
        dayToDay: ['Assist in developing brand campaigns, messaging, and creative briefs', 'Conduct consumer research: surveys, focus groups, social listening', 'Coordinate with creative agencies on ad production and deliverables', 'Analyze brand health metrics: awareness, perception, NPS', 'Monitor competitor brand activities and market trends weekly', 'Support packaging design reviews and brand guideline compliance', 'Help plan brand activations, events, and sponsorships', 'Create campaign performance reports and post-mortems'],
        hardSkills: [{rank:1,skill:'Brand Strategy Concepts (Positioning, equity, architecture)',level:'Intermediate'},{rank:2,skill:'Consumer Research Methods',level:'Intermediate'},{rank:3,skill:'PowerPoint (Brand decks, campaign presentations)',level:'Advanced'},{rank:4,skill:'Market Research Tools (SurveyMonkey, Google Forms)',level:'Basic'},{rank:5,skill:'Competitive Analysis',level:'Intermediate'},{rank:6,skill:'Basic Design Sense (Understanding aesthetics)',level:'Basic'},{rank:7,skill:'Project Management',level:'Basic'}],
        softSkills: ['Consumer empathy (understanding what drives behavior)', 'Creativity with business sense (ideas must serve objectives)', 'Storytelling ability (brands are built through narratives)', 'Attention to brand consistency (every detail matters)', 'Agency collaboration (working with external creative partners)'],
        degreeAlignmentPrimary: 'BBA (Marketing), Mass Communication, Economics',
        degreeAlignmentSecondary: 'Psychology and Sociology students have an edge in consumer understanding. Commerce students: leverage your business acumen. B.Tech: possible but need to demonstrate marketing interest through projects or certifications. FMCG companies (HUL, P&G) prefer top B-school candidates, but startups and D2C brands are more accessible.',
        isThisForYou: 'This role is for you if: You think about brands beyond logos — their personality and meaning. You notice advertising everywhere and analyze what makes it work. You care about consumer psychology. You want to build memorable brands, not just run campaigns. You\'re okay with longer-term impact over instant metrics.',
      },
    ],
    topSkills: ['Content Writing', 'Canva / Design Tools', 'Meta & Google Ads', 'Analytics (GA4)', 'Brand Strategy', 'Consumer Research'],
  },

  'Content & Creative': {
    colorClass: 'dh-content', emoji: '✍️',
    tagline: 'Storytellers and visual creators who craft brand narratives',
    bestFor: 'You think in stories, have a distinctive voice, and create content people actually want to consume. You enjoy the craft of writing, design, or video — and you\'re okay with subjective feedback.',
    roles: [
      {
        name: 'Content Writer Intern', stipend: '₹10K–30K/month', duration: '2–3 months', hours: '40–50 hrs/week', access: 'Accessible',
        dayToDay: ['Write 2–4 blog posts/articles per week based on briefs', 'Research topics thoroughly before writing (1–2 hours per piece)', 'Edit and proofread content for grammar, clarity, and flow', 'Optimize content for SEO: keywords, meta descriptions, headers', 'Maintain brand voice and tone consistency across all content', 'Track content performance: views, time on page, engagement', 'Incorporate feedback from editors quickly', 'Manage multiple pieces at different stages simultaneously'],
        hardSkills: [{rank:1,skill:'Writing (Clear, engaging, error-free)',level:'Advanced'},{rank:2,skill:'SEO Basics (Keywords, on-page optimization)',level:'Intermediate'},{rank:3,skill:'Research Skills',level:'Intermediate'},{rank:4,skill:'CMS (WordPress, Webflow)',level:'Basic'},{rank:5,skill:'Editing & Proofreading',level:'Intermediate'},{rank:6,skill:'Basic Analytics (GA, content metrics)',level:'Basic'},{rank:7,skill:'Content Calendar Management',level:'Basic'}],
        softSkills: ['Creativity (fresh angles on topics)', 'Attention to detail (typos destroy credibility)', 'Adaptability (writing in different tones and styles)', 'Receptivity to feedback (your draft will be edited)', 'Deadline discipline (content calendars are strict)', 'Curiosity (researching new topics constantly)'],
        degreeAlignmentPrimary: 'English, Journalism, Mass Communication, Literature',
        degreeAlignmentSecondary: 'Writing portfolio matters more than degree. B.Tech students: you can write technical content, start a tech blog. Commerce students: finance/business content is in demand. ANY degree works with strong writing samples. Start a blog, write on Medium, build portfolio.',
        isThisForYou: 'This role is for you if: You love writing and notice good/bad prose everywhere. You can adapt your style for different audiences. You enjoy researching and learning new topics. You accept edits without being defensive. You want to be known for your writing craft.',
      },
      {
        name: 'Copywriter Intern', stipend: '₹10K–30K/month', duration: '2–3 months', hours: '40–50 hrs/week', access: 'Accessible',
        dayToDay: ['Write ad copy, taglines, and short-form marketing content', 'Create email subject lines and CTAs that convert', 'Write social media captions and posts daily', 'Develop product descriptions and landing page copy', 'A/B test different copy versions', 'Study competitor messaging and positioning', 'Maintain brand voice across all communications', 'Work with designers on visual-copy alignment'],
        hardSkills: [{rank:1,skill:'Copywriting (Persuasive, punchy writing)',level:'Advanced'},{rank:2,skill:'Understanding of Marketing Funnels',level:'Intermediate'},{rank:3,skill:'A/B Testing Concepts',level:'Basic'},{rank:4,skill:'Basic Design Sense',level:'Basic'},{rank:5,skill:'Social Media Platform Knowledge',level:'Intermediate'},{rank:6,skill:'Email Marketing Basics',level:'Basic'},{rank:7,skill:'Brand Voice Documentation',level:'Developing'}],
        softSkills: ['Creativity (generating many options quickly)', 'Persuasion (understanding what makes people act)', 'Brevity (saying more with less)', 'Consumer psychology awareness', 'Thick skin (copy gets revised constantly)'],
        degreeAlignmentPrimary: 'English, Advertising, Mass Communication, Journalism',
        degreeAlignmentSecondary: 'Copywriting is a skill, not a degree. Any background works with a strong copy portfolio. Create spec ads for brands you like. Write mock email sequences. Show any copy that drove results. Advertising agencies value raw talent over degrees.',
        isThisForYou: 'This role is for you if: You can say a lot in very few words. You understand what makes people click, buy, or act. You notice and analyze ads everywhere. You enjoy the challenge of brevity. You want your words to drive action, not just inform.',
      },
      {
        name: 'Content Strategist Intern', stipend: '₹12K–35K/month', duration: '2–4 months', hours: '40–50 hrs/week', access: 'Moderate',
        dayToDay: ['Develop content calendars aligned with business objectives', 'Conduct content audits and gap analysis', 'Research audience needs and content preferences', 'Define content pillars and topic clusters', 'Work with writers to brief content requirements', 'Track content performance and optimize strategy', 'Analyze competitor content strategies', 'Create content guidelines and style documentation'],
        hardSkills: [{rank:1,skill:'Content Strategy Frameworks',level:'Developing'},{rank:2,skill:'SEO (Topic clustering, keyword strategy)',level:'Intermediate'},{rank:3,skill:'Analytics (GA, content metrics)',level:'Intermediate'},{rank:4,skill:'Editorial Calendar Management',level:'Intermediate'},{rank:5,skill:'Research & Analysis',level:'Intermediate'},{rank:6,skill:'PowerPoint (Strategy presentations)',level:'Basic'},{rank:7,skill:'Writing',level:'Intermediate'}],
        softSkills: ['Strategic thinking (content aligned to business goals)', 'Organization (managing editorial operations)', 'Analytical mindset (performance-driven decisions)', 'Communication (briefing writers effectively)', 'Long-term thinking (content compounds over time)'],
        degreeAlignmentPrimary: 'Mass Communication, Marketing, English, Journalism',
        degreeAlignmentSecondary: 'Strategy roles require a mix of creative and analytical. Commerce/BBA students: your business understanding helps with alignment to goals. Any background with demonstrated content work can transition. Show content projects with measurable results.',
        isThisForYou: 'This role is for you if: You think about content as a system, not just individual pieces. You can connect content to business outcomes. You\'re organized enough to manage editorial calendars. You enjoy the planning side of content. You want to grow into content leadership.',
      },
      {
        name: 'Graphic Design Intern', stipend: '₹8K–25K/month', duration: '2–3 months', hours: '40–50 hrs/week', access: 'Accessible',
        dayToDay: ['Create social media graphics and marketing visuals daily', 'Design presentations, infographics, and reports', 'Support brand consistency across all visual materials', 'Resize and adapt designs for different platforms and formats', 'Collaborate with content and marketing teams on creative briefs', 'Maintain design asset library and templates', 'Iterate on designs based on feedback', 'Stay updated on design trends'],
        hardSkills: [{rank:1,skill:'Adobe Creative Suite (Photoshop, Illustrator)',level:'Intermediate to Advanced'},{rank:2,skill:'Figma/Sketch',level:'Intermediate'},{rank:3,skill:'Canva',level:'Advanced'},{rank:4,skill:'Typography & Color Theory',level:'Intermediate'},{rank:5,skill:'Social Media Design Specs',level:'Intermediate'},{rank:6,skill:'Basic Video Editing',level:'Basic'},{rank:7,skill:'Design Principles',level:'Intermediate'}],
        softSkills: ['Creativity (fresh visual approaches)', 'Attention to detail (pixel-perfect execution)', 'Receptivity to feedback (designs will be revised)', 'Time management (multiple design requests daily)', 'Collaboration (working with non-designers)'],
        degreeAlignmentPrimary: 'Design (Graphic, Visual, Communication), Fine Arts',
        degreeAlignmentSecondary: 'Design is portfolio-driven. Non-design degrees can break in with strong portfolios. Self-taught designers are common. Build portfolio: redesign existing brands, create spec work, do freelance projects. Tools like Canva make design accessible. Show 10–15 strong pieces.',
        isThisForYou: 'This role is for you if: You think visually and notice design everywhere. You enjoy creating graphics and visual content. You can take feedback without being defensive. You like the variety of daily design requests. You want to build a strong portfolio quickly.',
      },
    ],
    topSkills: ['Creative Writing', 'SEO Basics', 'Canva / Figma', 'Content Strategy', 'Copywriting Frameworks', 'Analytics'],
  },

  'Finance & Investment': {
    colorClass: 'dh-finance', emoji: '📊',
    tagline: 'Numbers-first thinkers who thrive in high-pressure finance roles',
    bestFor: 'You love financial models, are energized by quantitative analysis, and want direct exposure to capital allocation or corporate finance. You\'re okay with grueling hours in exchange for top-tier exit opportunities.',
    roles: [
      {
        name: 'FP&A (Financial Planning & Analysis) Intern', stipend: '₹15K–45K/month', duration: '2–4 months', hours: '45–55 hrs/week', access: 'Moderate',
        dayToDay: ['Assist in monthly/quarterly budgeting and forecasting processes', 'Build and maintain financial reports and variance analysis', 'Support month-end close activities and reconciliations', 'Create dashboards for financial KPIs', 'Analyze business unit performance vs. budget', 'Help with capital allocation and investment analysis', 'Prepare materials for finance reviews with leadership', 'Maintain financial models and update assumptions'],
        hardSkills: [{rank:1,skill:'Excel (Financial reporting, modeling, pivot tables)',level:'Intermediate'},{rank:2,skill:'Financial Statement Understanding (P&L, BS, CF)',level:'Intermediate'},{rank:3,skill:'Budgeting & Forecasting',level:'Developing'},{rank:4,skill:'Variance Analysis',level:'Developing'},{rank:5,skill:'Data Visualization (Charts, dashboards)',level:'Intermediate'},{rank:6,skill:'Basic Accounting Principles',level:'Intermediate'},{rank:7,skill:'PowerPoint (Finance presentations)',level:'Basic'}],
        softSkills: ['Attention to detail (one wrong number breaks trust)', 'Deadline management (month-end is non-negotiable)', 'Cross-functional communication (explaining finance to non-finance)', 'Comfort with repetition (some tasks are cyclical)', 'Analytical mindset (understanding what numbers mean)'],
        degreeAlignmentPrimary: 'B.Com, CA (pursuing), MBA Finance, Economics',
        degreeAlignmentSecondary: 'FP&A is more accessible than IB or ER. Engineering students: your analytical skills transfer, but learn accounting basics. BBA students: this is a natural fit. Focus on Excel skills and understanding financial statements. Internships in audit or accounting also help.',
        isThisForYou: 'This role is for you if: You enjoy planning and forecasting. You want work-life balance within finance. You like explaining numbers to non-finance people. You want to understand how a company\'s finances actually work. You prefer predictability over deal-based intensity.',
      },
      {
        name: 'Investment Banking Analyst Intern', stipend: '₹25K–80K/month', duration: '2–3 months (summer)', hours: '70–100 hrs/week', access: 'Highly Competitive',
        dayToDay: ['Build complex financial models (DCF, LBO, merger models)', 'Create pitch books and client presentations', 'Conduct company and industry research for deal origination', 'Prepare comparable company and precedent transaction analyses', 'Support live deal execution: due diligence, documentation', 'Update models and materials based on senior banker feedback', 'Attend client calls and meetings when appropriate', 'Work on multiple deals simultaneously'],
        hardSkills: [{rank:1,skill:'Excel (Advanced modeling, shortcuts, macros)',level:'Expert'},{rank:2,skill:'Financial Statement Analysis (Deep linkages between statements)',level:'Advanced'},{rank:3,skill:'Valuation (DCF, Comps, Precedents, LBO)',level:'Advanced'},{rank:4,skill:'PowerPoint (Pitch books, client-ready quality)',level:'Advanced'},{rank:5,skill:'Accounting (GAAP/IFRS, adjustments)',level:'Intermediate'},{rank:6,skill:'M&A Concepts (Deal structures, synergies)',level:'Intermediate'},{rank:7,skill:'Industry Knowledge',level:'Ongoing'}],
        softSkills: ['Extreme attention to detail (one error = destroyed credibility)', 'Stamina (80–100 hour weeks are real)', 'Grace under pressure (last-minute changes are constant)', 'Perfectionism (good enough doesn\'t exist)', 'Client management (professional at all times)'],
        degreeAlignmentPrimary: 'CA, CFA (pursuing), MBA Finance (top schools), Economics (top colleges)',
        degreeAlignmentSecondary: 'HARD TRUTH: IB is extremely competitive. IIT/IIM pedigree matters at bulge bracket banks. BUT boutique banks (Avendus, o3 Capital, RBSA) are more accessible. Any degree can work if you: master financial modeling, network aggressively, and ace technicals. Start with smaller banks, build experience.',
        isThisForYou: 'This role is for you if: You\'re obsessed with accuracy — errors are unacceptable. You\'re genuinely willing to sacrifice work-life balance for intense learning. Financial statements and deal structures excite you. You want the fastest path to high compensation. You thrive under extreme pressure.',
      },
      {
        name: 'Equity Research Intern', stipend: '₹10K–65K/month', duration: '2–4 months', hours: '50–65 hrs/week', access: 'Competitive',
        dayToDay: ['Build and maintain financial models (3-statement, valuation)', 'Track company news, earnings calls, and industry developments daily', 'Write sections of research reports with investment thesis', 'Conduct industry analysis and competitive positioning', 'Maintain databases and update estimates', 'Attend management calls and summarize key points', 'Present stock ideas to senior analysts', 'Monitor portfolio companies continuously'],
        hardSkills: [{rank:1,skill:'Financial Modeling (3-statement, forecasting)',level:'Advanced'},{rank:2,skill:'Equity Valuation (P/E, EV/EBITDA, DCF)',level:'Advanced'},{rank:3,skill:'Financial Statement Analysis',level:'Advanced'},{rank:4,skill:'Research Writing (Clear investment thesis)',level:'Intermediate'},{rank:5,skill:'Industry Analysis',level:'Intermediate'},{rank:6,skill:'Excel (Efficient modeling)',level:'Advanced'},{rank:7,skill:'Bloomberg/Databases',level:'Basic exposure'}],
        softSkills: ['Intellectual curiosity (genuine interest in companies and industries)', 'Writing ability (clear, concise research reports)', 'Independent thinking (forming your own views, not just following consensus)', 'Conviction (defending your thesis when challenged)', 'Continuous learning (markets evolve constantly)'],
        degreeAlignmentPrimary: 'CA, CFA (pursuing), MBA Finance, Economics',
        degreeAlignmentSecondary: 'ER values demonstrated passion for markets. Any degree can work if you: actively follow stocks, have investment experience (even personal portfolio), and can articulate investment theses. Write mock research reports to showcase skills. Join investment clubs.',
        isThisForYou: 'This role is for you if: You love following markets and have opinions on stocks. Reading annual reports excites you. You can form and defend an investment thesis. You want to become an expert in specific sectors. You prefer research over deal execution.',
      },
      {
        name: 'Venture Capital Analyst Intern', stipend: '₹20K–50K/month', duration: '3–6 months', hours: '50–60 hrs/week', access: 'Competitive',
        dayToDay: ['Source and screen startup investment opportunities', 'Conduct market research on emerging sectors and trends', 'Build financial models and valuation for startups', 'Prepare investment memos and due diligence reports', 'Attend startup pitches and take notes', 'Track portfolio company performance', 'Research competitive landscape for sectors of interest', 'Support partners with deal execution'],
        hardSkills: [{rank:1,skill:'Startup Valuation Methods',level:'Intermediate'},{rank:2,skill:'Market Research & Analysis',level:'Intermediate'},{rank:3,skill:'Financial Modeling (startup-specific)',level:'Intermediate'},{rank:4,skill:'Excel (Analysis, cap tables)',level:'Intermediate'},{rank:5,skill:'PowerPoint (Investment memos)',level:'Intermediate'},{rank:6,skill:'Due Diligence Process',level:'Developing'},{rank:7,skill:'Understanding of Startup Ecosystem',level:'Intermediate'}],
        softSkills: ['Pattern recognition (identifying promising startups)', 'Networking (connecting with founders, other VCs)', 'Skeptical optimism (balancing enthusiasm with rigor)', 'Intellectual curiosity (learning new industries constantly)', 'Long-term thinking (VC returns take years)'],
        degreeAlignmentPrimary: 'MBA, Engineering, Economics, CA/CFA',
        degreeAlignmentSecondary: 'VC is relationship-driven. Any degree works if you: understand startups deeply, have worked at startups, or can demonstrate strong analytical skills. Having a network in the startup ecosystem helps significantly. Start by writing about startups, attending events.',
        isThisForYou: 'This role is for you if: You\'re fascinated by startups and founders. You can quickly assess business models and market potential. You enjoy meeting new people and building relationships. You think about investing with a 10-year horizon. You want to be at the intersection of tech and finance.',
      },
    ],
    topSkills: ['Excel Financial Modelling', 'Accounting Fundamentals', 'Valuation Techniques', 'Bloomberg / FactSet', 'PowerPoint Pitchbooks', 'Financial Statement Analysis'],
  },

  'Product & Growth': {
    colorClass: 'dh-product', emoji: '🚀',
    tagline: 'Data-driven builders obsessed with user problems and metrics',
    bestFor: 'You obsess over why products succeed, think in terms of user problems not features, and want to work at the intersection of engineering, design, and business. Metrics and experimentation excite you.',
    roles: [
      {
        name: 'Product Management Intern', stipend: '₹25K–50K/month', duration: '3–6 months', hours: '45–55 hrs/week', access: 'Highly Competitive',
        dayToDay: ['Conduct user research: 3–5 user interviews per week, surveys, usability tests', 'Analyze product metrics using Mixpanel/Amplitude: DAU, retention, funnel conversion', 'Write PRDs (Product Requirement Documents) for new features', 'Prioritize features using frameworks (RICE, MoSCoW, Kano)', 'Attend and contribute to sprint planning and standups with engineering', 'Run A/B tests and analyze results for feature iterations', 'Create wireframes and collaborate with designers on user flows', 'Present findings and recommendations to product leadership'],
        hardSkills: [{rank:1,skill:'Product Analytics (Mixpanel, Amplitude, GA4)',level:'Intermediate'},{rank:2,skill:'SQL (Data extraction, basic queries)',level:'Intermediate'},{rank:3,skill:'User Research Methods (Interviews, surveys)',level:'Intermediate'},{rank:4,skill:'Wireframing Basics (Figma, Balsamiq)',level:'Basic'},{rank:5,skill:'A/B Testing & Experimentation',level:'Basic'},{rank:6,skill:'PRD/Spec Writing',level:'Developing'},{rank:7,skill:'Basic Understanding of Software Development',level:'Basic'}],
        softSkills: ['User empathy (obsessing over user problems, not features)', 'Data-driven decision making (gut feel isn\'t enough)', 'Cross-functional communication (translating between eng, design, business)', 'Prioritization (saying no to 90% of ideas)', 'Comfort with ambiguity (requirements change constantly)'],
        degreeAlignmentPrimary: 'Engineering (any branch), Design, MBA',
        degreeAlignmentSecondary: 'PM is one of the hardest roles to break into without tech background. BUT it\'s possible: Economics/Psychology students — leverage user behavior understanding. BBA students — focus on business-side PM roles. CRITICAL: Build a side project, write product teardowns, show SQL skills. APM programs (Google, Flipkart) are very competitive. Startup PM roles are more accessible.',
        isThisForYou: 'This role is for you if: You obsess over why products succeed or fail. You think in terms of user problems, not features. You can work with engineers without being one. You\'re comfortable with constantly changing priorities. You want to build things millions of people use.',
      },
      {
        name: 'Growth Marketing / Hacker Intern', stipend: '₹15K–45K/month', duration: '2–4 months', hours: '45–55 hrs/week', access: 'Moderate',
        dayToDay: ['Analyze growth funnels daily: acquisition, activation, retention, revenue, referral', 'Identify drop-off points and hypothesize improvement experiments', 'Design and run 5–10 experiments weekly (small, fast, measurable)', 'Build growth dashboards tracking key metrics', 'Research growth tactics from other companies and industries', 'Implement quick wins: email sequences, onboarding flows, referral mechanics', 'Collaborate with product and marketing on growth initiatives', 'Document learnings and create growth playbooks'],
        hardSkills: [{rank:1,skill:'Growth Frameworks (AARRR, Growth Loops)',level:'Intermediate'},{rank:2,skill:'Analytics Tools (Mixpanel, Amplitude, GA4)',level:'Intermediate'},{rank:3,skill:'SQL (Funnel analysis, cohort analysis)',level:'Intermediate'},{rank:4,skill:'A/B Testing (Design, analysis, significance)',level:'Intermediate'},{rank:5,skill:'Basic Performance Marketing',level:'Basic'},{rank:6,skill:'Email Marketing Tools (Mailchimp, Customer.io)',level:'Basic'},{rank:7,skill:'Excel (Growth modeling, projections)',level:'Intermediate'}],
        softSkills: ['Experimentation obsession (run tests constantly, most will fail)', 'Speed over perfection (80% solution shipped beats 100% solution planned)', 'Creative problem-solving (unconventional tactics often work best)', 'Data intuition (spotting patterns in metrics)', 'Comfort with failure (95% of experiments don\'t work)'],
        degreeAlignmentPrimary: 'Engineering, Economics, BBA, MBA',
        degreeAlignmentSecondary: 'Growth is more about mindset than degree. Any background works if you demonstrate: analytical thinking, experimentation bias, and hustle. Show personal growth projects: growing a newsletter, building audience on social media, any metric you\'ve improved. Numbers speak louder than degrees.',
        isThisForYou: 'This role is for you if: You\'re obsessed with metrics and moving numbers. You love running experiments and learning from failures. You think in terms of leverage — small efforts, big impact. You\'re impatient and want to see results quickly. You don\'t need permission to try things.',
      },
      {
        name: 'User Research Intern', stipend: '₹10K–35K/month', duration: '2–4 months', hours: '40–50 hrs/week', access: 'Moderate',
        dayToDay: ['Recruit and schedule 5–10 user interviews per week', 'Conduct user interviews and usability tests (in-person and remote)', 'Synthesize interview notes into actionable insights', 'Design and deploy surveys to validate hypotheses', 'Create user personas and journey maps', 'Present research findings to product and design teams', 'Maintain user research repository and documentation', 'Identify patterns across research to inform product roadmap'],
        hardSkills: [{rank:1,skill:'User Interview Techniques',level:'Intermediate (develop on job)'},{rank:2,skill:'Survey Design (Typeform, Google Forms)',level:'Intermediate'},{rank:3,skill:'Usability Testing Methods',level:'Basic'},{rank:4,skill:'Qualitative Data Analysis',level:'Intermediate'},{rank:5,skill:'Presentation Skills (Research readouts)',level:'Intermediate'},{rank:6,skill:'Documentation (Notion, Confluence)',level:'Basic'},{rank:7,skill:'Basic Statistics',level:'Basic'}],
        softSkills: ['Active listening (hearing what users don\'t explicitly say)', 'Empathy (understanding user context and emotions)', 'Objectivity (reporting findings without bias)', 'Curiosity (digging deeper with follow-up questions)', 'Synthesis ability (turning 20 interviews into 3 insights)'],
        degreeAlignmentPrimary: 'Psychology, Sociology, Design, Anthropology',
        degreeAlignmentSecondary: 'Social science backgrounds have an advantage here. B.Tech students: your structured thinking helps in synthesis, but you need to develop empathy skills. BBA/Commerce: possible with demonstrated interest. Show any interview or research experience — even college projects count.',
        isThisForYou: 'This role is for you if: You\'re genuinely curious about why people do what they do. You\'re a good listener and can make people comfortable. You can synthesize qualitative data into clear insights. You care about user problems over business metrics. You enjoy the \'detective work\' of understanding behavior.',
      },
    ],
    topSkills: ['Product Analytics (Mixpanel/Amplitude)', 'SQL', 'User Research Methods', 'A/B Testing', 'Wireframing (Figma)', 'Growth Frameworks (AARRR)'],
  },

  'Operations & Process': {
    colorClass: 'dh-operations', emoji: '⚙️',
    tagline: 'System builders who turn chaos into efficient, scalable processes',
    bestFor: 'You\'re energized by making messy processes clean and efficient. You love checklists, timelines, and SOPs. You want to be the person who makes the machine run smoothly.',
    roles: [
      {
        name: 'Business Operations Intern', stipend: '₹10K–35K/month', duration: '2–4 months', hours: '45–55 hrs/week', access: 'Accessible',
        dayToDay: ['Analyze operational metrics to identify inefficiencies and improvement areas', 'Document processes and create SOPs (Standard Operating Procedures)', 'Support process improvement and automation initiatives', 'Coordinate between teams for operational execution', 'Build operational dashboards and reports', 'Handle escalations and problem-solve real-time issues', 'Track vendor and partner performance', 'Implement quick fixes for operational bottlenecks'],
        hardSkills: [{rank:1,skill:'Excel (Data analysis, pivot tables, dashboards)',level:'Advanced'},{rank:2,skill:'Process Mapping (Flowcharts, swimlane diagrams)',level:'Intermediate'},{rank:3,skill:'Data Analysis',level:'Intermediate'},{rank:4,skill:'Project Management Basics',level:'Developing'},{rank:5,skill:'SQL (Basic queries for operational data)',level:'Basic'},{rank:6,skill:'Documentation & SOPs',level:'Intermediate'},{rank:7,skill:'Basic Automation (Zapier, simple scripts)',level:'Basic'}],
        softSkills: ['Problem-solving (ops is about fixing things that break)', 'Attention to detail (small errors compound in operations)', 'Process orientation (thinking in systems and workflows)', 'Calm under pressure (things will go wrong constantly)', 'Cross-functional coordination (ops touches every team)'],
        degreeAlignmentPrimary: 'Engineering (Industrial, Mechanical), BBA, B.Com',
        degreeAlignmentSecondary: 'Ops values practical problem-solving over degrees. Any background works if you demonstrate: ability to fix problems, create systems, and stay organized. Show any experience organizing events, managing logistics, or improving processes — even for college fests.',
        isThisForYou: 'This role is for you if: Inefficiency genuinely frustrates you. You love creating systems and processes. You\'re detail-oriented and catch errors others miss. You want immediate, tangible impact. You stay calm when things go wrong.',
      },
      {
        name: 'Supply Chain Intern', stipend: '₹10K–40K/month', duration: '3–6 months', hours: '45–55 hrs/week', access: 'Moderate',
        dayToDay: ['Analyze inventory levels and demand patterns', 'Support vendor management and procurement processes', 'Track logistics and delivery performance metrics', 'Help with demand forecasting models', 'Coordinate with suppliers and logistics partners', 'Build reports on supply chain KPIs: fill rate, lead time, costs', 'Identify cost optimization opportunities', 'Support new vendor onboarding'],
        hardSkills: [{rank:1,skill:'Excel (Inventory analysis, forecasting)',level:'Advanced'},{rank:2,skill:'Supply Chain Fundamentals (Procurement, logistics)',level:'Developing'},{rank:3,skill:'Data Analysis',level:'Intermediate'},{rank:4,skill:'ERP Systems Basics (SAP concepts)',level:'Basic awareness'},{rank:5,skill:'Inventory Management Concepts',level:'Developing'},{rank:6,skill:'Vendor Management',level:'Basic'},{rank:7,skill:'Basic SQL',level:'Basic'}],
        softSkills: ['Negotiation (working with vendors on terms)', 'Relationship management (maintaining supplier relationships)', 'Crisis management (supply disruptions happen)', 'Systematic thinking (end-to-end process view)', 'Attention to numbers (margins are tight)'],
        degreeAlignmentPrimary: 'Engineering (Industrial, Mechanical), Operations Management',
        degreeAlignmentSecondary: 'Commerce and BBA students can break in with demonstrated interest. B.Tech students have an advantage in process understanding. Show any experience with logistics, procurement, or managing physical goods movement.',
        isThisForYou: 'This role is for you if: You find logistics and movement of goods fascinating. You enjoy optimizing costs and efficiency. You\'re good at coordinating between multiple parties. You can think about end-to-end systems. You want problems with real-world physical impact.',
      },
      {
        name: 'Program Manager Intern', stipend: '₹15K–40K/month', duration: '2–4 months', hours: '45–55 hrs/week', access: 'Moderate',
        dayToDay: ['Track project timelines, milestones, and deliverables across teams', 'Create and maintain project documentation and status reports', 'Coordinate meetings between multiple stakeholders', 'Identify risks and dependencies, escalate blockers', 'Support resource planning and allocation', 'Build project dashboards for leadership visibility', 'Document lessons learned and process improvements', 'Manage communication between teams'],
        hardSkills: [{rank:1,skill:'Project Management Tools (Asana, Jira, Monday)',level:'Intermediate'},{rank:2,skill:'Excel (Tracking, Gantt charts, dashboards)',level:'Intermediate'},{rank:3,skill:'Documentation (Confluence, Notion)',level:'Intermediate'},{rank:4,skill:'Meeting Facilitation',level:'Developing'},{rank:5,skill:'Risk Management Basics',level:'Basic'},{rank:6,skill:'PowerPoint (Status presentations)',level:'Basic'},{rank:7,skill:'Basic Understanding of Software Development',level:'Helpful'}],
        softSkills: ['Organization (managing multiple workstreams)', 'Communication (clear status updates to stakeholders)', 'Influence without authority (getting things done across teams)', 'Problem anticipation (seeing blockers before they hit)', 'Patience (herding cats is part of the job)'],
        degreeAlignmentPrimary: 'Engineering, BBA, MBA',
        degreeAlignmentSecondary: 'Any degree works if you demonstrate strong organization and communication. Show any experience coordinating complex projects — college events, technical projects, group assignments managed well. PMP/Agile certifications help but aren\'t required.',
        isThisForYou: 'This role is for you if: You\'re the person who keeps group projects on track. You love creating checklists and tracking progress. You can coordinate between people who don\'t naturally communicate. You stay calm when deadlines approach. You get satisfaction from shipping projects, not just starting them.',
      },
    ],
    topSkills: ['Excel / Google Sheets', 'Process Mapping', 'Project Management Tools', 'Data Analysis', 'Stakeholder Communication', 'SOP Documentation'],
  },

  "Founder's Office": {
    colorClass: 'dh-founders', emoji: '🏢',
    tagline: 'Generalist operators who want to learn everything at a startup',
    bestFor: 'You\'re adaptable, execution-oriented, and want exposure to all parts of a business. You\'re comfortable with ambiguity, don\'t need a defined role, and want to work closely with leadership from day one.',
    roles: [
      {
        name: "Founder's Office Intern", stipend: '₹10K–75K/month', duration: '3–6 months', hours: '50–70 hrs/week', access: 'Competitive',
        dayToDay: ['Work on high-priority strategic projects assigned directly by founders', 'Prepare presentations for investors, board meetings, and partners', 'Conduct research on new business opportunities, markets, competitors', 'Support cross-functional initiatives that don\'t fit any specific team', 'Handle ad-hoc tasks across functions: ops, marketing, BD, product', 'Attend leadership meetings and track action items to completion', 'Create reports and dashboards for founder visibility', 'Coordinate between teams to unblock bottlenecks'],
        hardSkills: [{rank:1,skill:'PowerPoint (Executive presentations, investor decks)',level:'Advanced'},{rank:2,skill:'Excel (Analysis, modeling, dashboards)',level:'Intermediate'},{rank:3,skill:'Research (Market, competitive, industry)',level:'Intermediate'},{rank:4,skill:'Project Management',level:'Developing'},{rank:5,skill:'Writing (Emails, memos, reports)',level:'Intermediate'},{rank:6,skill:'Basic Financial Understanding',level:'Basic'},{rank:7,skill:'Data Analysis',level:'Basic'}],
        softSkills: ['Extreme adaptability (priorities change daily)', 'Ownership mindset (no one will hand-hold you)', 'Discretion (you\'ll hear confidential information)', 'Speed of execution (startups move fast)', 'Comfort with chaos (nothing is structured)', 'Proactive communication (don\'t wait to be asked)'],
        degreeAlignmentPrimary: 'Any degree — attitude matters most',
        degreeAlignmentSecondary: 'Founder\'s Office is the great equalizer. BBA, B.Tech, BA, B.Sc — doesn\'t matter. What matters: Can you execute fast? Can you handle ambiguity? Are you a self-starter? Show any entrepreneurial experience: college startups, side projects, freelancing. Hunger > Credentials.',
        isThisForYou: 'This role is for you if: You want to see how a startup operates from the inside. You thrive in chaos and constantly changing priorities. You want maximum learning over a defined role. You can pick up new skills quickly. Startup intensity excites rather than stresses you.',
      },
      {
        name: 'Chief of Staff / Executive Assistant Intern', stipend: '₹10K–50K/month', duration: '3–6 months', hours: '50–70 hrs/week', access: 'Moderate',
        dayToDay: ['Manage founder\'s calendar, scheduling, and meeting preparation', 'Prepare briefing documents for meetings and calls', 'Draft communications on behalf of leadership', 'Track key initiatives and ensure follow-through', 'Coordinate board meeting logistics and materials', 'Handle confidential projects and sensitive information', 'Serve as liaison between founder and rest of organization', 'Anticipate needs before being asked'],
        hardSkills: [{rank:1,skill:'Calendar & Email Management',level:'Advanced'},{rank:2,skill:'PowerPoint (Meeting prep, decks)',level:'Intermediate'},{rank:3,skill:'Excel (Tracking, dashboards)',level:'Basic'},{rank:4,skill:'Written Communication',level:'Advanced'},{rank:5,skill:'Project Coordination',level:'Intermediate'},{rank:6,skill:'Meeting Facilitation',level:'Developing'},{rank:7,skill:'Basic Research',level:'Intermediate'}],
        softSkills: ['Anticipation (knowing what\'s needed before being asked)', 'Discretion (absolute confidentiality)', 'Organization (managing multiple priorities flawlessly)', 'Diplomatic communication (representing leadership)', 'Calm under pressure (urgent requests are constant)'],
        degreeAlignmentPrimary: 'Any degree',
        degreeAlignmentSecondary: 'This role values organization and communication over academic background. Show any experience managing complex schedules, coordinating events, or handling multiple stakeholders. Executive presence matters — practice professional communication.',
        isThisForYou: 'This role is for you if: You\'re ultra-organized and detail-oriented. You anticipate needs before they\'re expressed. You can handle confidential information with discretion. You want to learn how executives operate. You\'re comfortable being the \'behind the scenes\' enabler.',
      },
    ],
    topSkills: ['Execution Bias', 'Communication', 'Excel / Sheets', 'Project Coordination', 'Problem-Solving', 'Adaptability'],
  },

  'Data & Analytics': {
    colorClass: 'dh-data', emoji: '📈',
    tagline: 'Quantitative thinkers who find patterns where others see noise',
    bestFor: 'You love working with numbers, finding patterns in data, and turning raw information into business decisions. You\'re comfortable with SQL, Excel models, and data visualization.',
    roles: [
      {
        name: 'Business Analyst Intern', stipend: '₹15K–45K/month', duration: '2–4 months', hours: '45–55 hrs/week', access: 'Competitive',
        dayToDay: ['Extract and analyze data using SQL queries', 'Build dashboards and automated reports for business teams', 'Identify trends and patterns in business data', 'Support decision-making with data-driven insights', 'Create data visualizations for stakeholders', 'Document data definitions and metrics', 'Work with product/ops teams to define requirements', 'Present findings and recommendations to leadership'],
        hardSkills: [{rank:1,skill:'SQL (Complex queries, joins, aggregations)',level:'Advanced'},{rank:2,skill:'Excel (Advanced analysis, pivot tables)',level:'Advanced'},{rank:3,skill:'Data Visualization (Tableau, Power BI, Looker)',level:'Intermediate'},{rank:4,skill:'Python/R Basics (pandas, basic analysis)',level:'Basic'},{rank:5,skill:'Statistics (Descriptive, basic inferential)',level:'Intermediate'},{rank:6,skill:'Business Understanding (KPIs, metrics)',level:'Developing'},{rank:7,skill:'Presentation Skills',level:'Intermediate'}],
        softSkills: ['Analytical thinking (seeing patterns in data)', 'Communication (translating data to insights)', 'Curiosity (asking why the numbers are what they are)', 'Attention to detail (data must be accurate)', 'Stakeholder management (understanding what business needs)'],
        degreeAlignmentPrimary: 'Engineering, Statistics, Mathematics, Economics',
        degreeAlignmentSecondary: 'BA is technical-leaning. Commerce students: possible with strong SQL and Excel. BBA: need to build technical skills. Key differentiator: SQL proficiency. Take online courses, practice on real datasets. Show any data analysis projects.',
        isThisForYou: 'This role is for you if: You love finding patterns in data. You can translate analysis into simple insights. You\'re comfortable with SQL and spreadsheets. You want to influence decisions through data. You enjoy the intersection of tech and business.',
      },
      {
        name: 'Data Visualization / BI Intern', stipend: '₹12K–35K/month', duration: '2–4 months', hours: '40–50 hrs/week', access: 'Competitive',
        dayToDay: ['Build and maintain dashboards in Tableau/Power BI/Looker', 'Extract data using SQL for visualization needs', 'Design visually clear and insightful charts and graphs', 'Automate reporting workflows', 'Document dashboard logic and data sources', 'Work with stakeholders to understand visualization requirements', 'Ensure data accuracy in all visualizations', 'Create self-serve analytics capabilities for teams'],
        hardSkills: [{rank:1,skill:'Tableau/Power BI/Looker',level:'Advanced'},{rank:2,skill:'SQL',level:'Intermediate'},{rank:3,skill:'Data Visualization Best Practices',level:'Intermediate'},{rank:4,skill:'Excel',level:'Intermediate'},{rank:5,skill:'Basic Statistics',level:'Basic'},{rank:6,skill:'Data Modeling Concepts',level:'Basic'},{rank:7,skill:'Design Principles',level:'Basic'}],
        softSkills: ['Visual design sense (what makes data clear)', 'Attention to detail (visualizations must be accurate)', 'User empathy (understanding what stakeholders need to see)', 'Communication (explaining visualization choices)', 'Patience (stakeholder requirements often unclear)'],
        degreeAlignmentPrimary: 'Engineering, Statistics, Mathematics, Design',
        degreeAlignmentSecondary: 'BI roles are accessible to multiple backgrounds. Commerce students: learn Tableau/Power BI. Design students: your visual sense is valuable, learn SQL. Any background works with demonstrated portfolio of dashboards. Free Tableau Public account to showcase work.',
        isThisForYou: 'This role is for you if: You love turning complex data into clear visuals. You have an eye for design and aesthetics. You enjoy making information accessible to non-technical people. You\'re patient with stakeholder feedback. You like the technical + creative combination.',
      },
    ],
    topSkills: ['SQL', 'Excel (Pivot Tables, VLOOKUP)', 'Tableau / Power BI', 'Python (Pandas)', 'Statistical Analysis', 'Data Storytelling'],
  },

  'Human Resources': {
    colorClass: 'dh-hr', emoji: '👥',
    tagline: 'People-first professionals who build and nurture great teams',
    bestFor: 'You\'re genuinely curious about people, have high emotional intelligence, and want to shape company culture. You\'re energized by helping others and find satisfaction in building communities.',
    roles: [
      {
        name: 'Talent Acquisition Intern', stipend: '₹10K–30K/month', duration: '2–4 months', hours: '40–50 hrs/week', access: 'Accessible',
        dayToDay: ['Source candidates through LinkedIn, job boards, and referrals', 'Screen resumes and conduct initial phone screens (20–30/week)', 'Coordinate interviews between candidates and hiring managers', 'Maintain ATS (Applicant Tracking System) with updated candidate status', 'Draft and post job descriptions on various platforms', 'Support employer branding initiatives', 'Track recruitment metrics: time-to-fill, source effectiveness', 'Provide candidate experience feedback to improve process'],
        hardSkills: [{rank:1,skill:'LinkedIn Recruiter/Sourcing',level:'Intermediate'},{rank:2,skill:'ATS Management (Greenhouse, Lever, Zoho)',level:'Developing'},{rank:3,skill:'Resume Screening',level:'Developing'},{rank:4,skill:'Interview Scheduling & Coordination',level:'Basic'},{rank:5,skill:'Job Description Writing',level:'Basic'},{rank:6,skill:'Excel (Recruitment tracking)',level:'Basic'},{rank:7,skill:'Employer Branding Basics',level:'Basic'}],
        softSkills: ['Communication (representing company to candidates)', 'Judgment of people (quick assessment of fit)', 'Organization (managing multiple open roles)', 'Persistence (talent market is competitive)', 'Sales mindset (selling the company to candidates)'],
        degreeAlignmentPrimary: 'Any degree — people skills matter most',
        degreeAlignmentSecondary: 'TA is one of the most accessible corporate roles. Psychology students: your understanding of behavior helps. BBA/B.Com: natural fit. B.Tech: you can recruit for tech roles with your domain understanding. Show any experience evaluating or interviewing people.',
        isThisForYou: 'This role is for you if: You\'re good at reading people quickly. You enjoy the hunt for finding great candidates. You can sell the company and role enthusiastically. You\'re highly organized with multiple priorities. You want to shape who joins the company.',
      },
      {
        name: 'Employer Branding Intern', stipend: '₹10K–25K/month', duration: '2–3 months', hours: '40–45 hrs/week', access: 'Accessible',
        dayToDay: ['Create content showcasing company culture (posts, videos, blogs)', 'Manage employer branding social media accounts (LinkedIn, Glassdoor)', 'Coordinate employee stories and testimonials', 'Support campus recruitment marketing efforts', 'Track employer brand metrics: Glassdoor ratings, social engagement', 'Research competitor employer branding strategies', 'Help with career site content and updates', 'Support internal communications for culture initiatives'],
        hardSkills: [{rank:1,skill:'Content Creation (Writing, basic video)',level:'Intermediate'},{rank:2,skill:'Social Media Management',level:'Intermediate'},{rank:3,skill:'Canva/Basic Design',level:'Basic'},{rank:4,skill:'Glassdoor/LinkedIn Management',level:'Basic'},{rank:5,skill:'Copywriting',level:'Intermediate'},{rank:6,skill:'Analytics (Social media metrics)',level:'Basic'},{rank:7,skill:'Photography/Video Basics',level:'Helpful'}],
        softSkills: ['Storytelling (making culture compelling)', 'Creativity (standing out in crowded employer market)', 'Collaboration (working with employees across company)', 'Empathy (understanding what candidates care about)', 'Attention to brand voice (consistency matters)'],
        degreeAlignmentPrimary: 'Mass Communication, Journalism, BBA',
        degreeAlignmentSecondary: 'Mix of marketing and HR. Any degree works if you have content creation skills. English/Literature students: your writing helps. Show any content portfolio — even personal social media presence counts.',
        isThisForYou: 'This role is for you if: You understand what makes companies attractive to work at. You enjoy creating content and storytelling. You care about culture and employee experience. You want a creative role within HR. You\'re comfortable interviewing employees for their stories.',
      },
      {
        name: 'HR Generalist Intern', stipend: '₹10K–25K/month', duration: '2–4 months', hours: '40–50 hrs/week', access: 'Accessible',
        dayToDay: ['Support employee onboarding and documentation', 'Assist with HR administration and record keeping', 'Help organize employee engagement activities and events', 'Support performance management process administration', 'Handle employee queries on policies and benefits', 'Maintain HRIS data accuracy', 'Assist with training coordination', 'Support exit process and documentation'],
        hardSkills: [{rank:1,skill:'HRIS/HR Systems',level:'Developing'},{rank:2,skill:'Excel (HR data, tracking)',level:'Basic'},{rank:3,skill:'Documentation & Filing',level:'Basic'},{rank:4,skill:'Employee Onboarding Process',level:'Developing'},{rank:5,skill:'Basic Employment Law Awareness',level:'Basic'},{rank:6,skill:'Training Coordination',level:'Basic'},{rank:7,skill:'Event Planning',level:'Basic'}],
        softSkills: ['Confidentiality (HR deals with sensitive information)', 'Empathy (employees come with problems)', 'Organization (multiple administrative tasks)', 'Patience (repetitive queries are common)', 'Approachability (people should feel comfortable asking)'],
        degreeAlignmentPrimary: 'Any degree',
        degreeAlignmentSecondary: 'HR Generalist is very accessible. Psychology students: employee relations is a good fit. BBA: natural path. Any background works if you demonstrate: organization, empathy, and discretion. Show any experience handling confidential matters or coordinating with groups.',
        isThisForYou: 'This role is for you if: You genuinely care about employee experience. You\'re organized and detail-oriented. You can maintain confidentiality. You enjoy helping people with their queries. You want exposure to all aspects of HR.',
      },
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

    // Domain detail cards for top 3 — tab-switched
    const detailsContainer = document.getElementById('domain-details');
    const tabsContainer = document.getElementById('domain-tabs');
    const tabLabels = ['🥇 #1 Best Fit', '🥈 #2 Strong Fit', '🥉 #3 Good Fit'];

    // Build cards array
    const cards = top3.map(([domain, score], i) => buildDomainCard(domain, score, i));

    // Build tab buttons
    tabsContainer.innerHTML = tabLabels.map((label, i) => `
      <button class="domain-tab-btn ${i === 0 ? 'active' : ''}" data-idx="${i}" onclick="App.switchDomainTab(${i})">
        ${label}<br><span class="domain-tab-name">${top3[i][0]}</span>
      </button>
    `).join('');
    tabsContainer.style.display = 'flex';

    // Render single active card
    detailsContainer.innerHTML = '';
    const heading = document.createElement('h2');
    heading.style.cssText = 'font-size:20px;font-weight:800;color:#0F172A;margin:0 0 16px 0;';
    heading.textContent = 'Deep Dive: Your Top Domain';
    detailsContainer.appendChild(heading);
    detailsContainer.appendChild(cards[0]);

    // Store cards for tab switching
    window._domainCards = cards;
    window._domainHeading = heading;

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

    // Auto-open active domain card
    setTimeout(() => {
      const first = detailsContainer.querySelector('.domain-detail-card');
      if (first) first.classList.add('open');
    }, 400);
  }

  function switchDomainTab(idx) {
    const detailsContainer = document.getElementById('domain-details');
    // Update tab active state
    document.querySelectorAll('.domain-tab-btn').forEach((btn, i) => {
      btn.classList.toggle('active', i === idx);
    });
    // Swap card
    detailsContainer.innerHTML = '';
    detailsContainer.appendChild(window._domainHeading);
    detailsContainer.appendChild(window._domainCards[idx]);
    // Auto-open
    setTimeout(() => {
      const card = detailsContainer.querySelector('.domain-detail-card');
      if (card) card.classList.add('open');
    }, 100);
  }

  function buildDomainCard(domain, score, rank) {
    const meta = DOMAIN_META[domain];
    if (!meta) return document.createElement('div');

    const card = document.createElement('div');
    card.className = 'domain-detail-card';

    const medals = ['🥇', '🥈', '🥉'];
    const rankLabel = ['Best Fit', 'Strong Fit', 'Good Fit'][rank];

    const rolesHTML = meta.roles.map((r, idx) => `
      <div class="role-card">
        <div class="role-card-header" onclick="this.nextElementSibling.style.display = this.nextElementSibling.style.display === 'block' ? 'none' : 'block'; this.querySelector('.role-toggle').textContent = this.nextElementSibling.style.display === 'block' ? '▲ Hide Details' : '▼ View Details'">
          <div>
            <div class="role-name">${r.name}</div>
            <div class="role-meta">
              <span class="role-tag stipend">💰 ${r.stipend}</span>
              <span class="role-tag duration">📅 ${r.duration}</span>
              <span class="role-tag hours">🕐 ${r.hours}</span>
              <span class="role-tag access">📌 ${r.access}</span>
            </div>
          </div>
          <div class="role-toggle">▼ View Details</div>
        </div>
        <div class="role-detail" style="display:none">
          <div class="role-section">
            <div class="role-section-title">📋 Day-to-Day Reality</div>
            <ul class="role-bullets">${r.dayToDay.map(d => `<li>${d}</li>`).join('')}</ul>
          </div>
          <div class="role-two-col">
            <div class="role-section">
              <div class="role-section-title">🛠 Hard Skills (Priority Order)</div>
              <ol class="role-skills-list">${r.hardSkills.map(s => `<li><strong>${s.skill}</strong> <span class="skill-level">${s.level}</span></li>`).join('')}</ol>
            </div>
            <div class="role-section">
              <div class="role-section-title">🧠 Soft Skills</div>
              <ul class="role-bullets">${r.softSkills.map(s => `<li>${s}</li>`).join('')}</ul>
            </div>
          </div>
          <div class="role-section">
            <div class="role-section-title">🎓 Degree Alignment</div>
            <div class="degree-primary"><strong>Primary:</strong> ${r.degreeAlignmentPrimary}</div>
            <div class="degree-secondary">${r.degreeAlignmentSecondary}</div>
          </div>
          <div class="role-section is-for-you-section">
            <div class="role-section-title">✅ Is This Role For You?</div>
            <p class="is-for-you-text">${r.isThisForYou}</p>
          </div>
        </div>
      </div>
    `).join('');

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
          <div class="skills-title">💼 Internship Roles (click each to expand full details)</div>
          <div class="roles-grid-full">${rolesHTML}</div>
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

  return { showForm, submitUserForm, startQuiz, nextQ, prevQ, restart, switchDomainTab };
})();

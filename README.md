# MBA Domain & Role Discovery Tool

**Live tool:** Answer 20 questions → get your top 3 MBA internship domain recommendations with match scores (0–100).

---

## Deploy to GitHub Pages

```bash
# 1. Create a new GitHub repo (e.g. "mba-discovery-tool")
# 2. Push these files:
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/mba-discovery-tool.git
git push -u origin main

# 3. In GitHub → Settings → Pages → Source: "Deploy from branch" → main → / (root)
# Your app will be live at: https://YOUR_USERNAME.github.io/mba-discovery-tool/
```

---

## File Structure

```
mba-discovery-tool/
├── index.html   — App shell, all screens (Welcome / Quiz / Results)
├── style.css    — All styling (responsive, animated)
├── app.js       — Questions, scoring logic, rendering
└── README.md    — This file
```

---

## How the Scoring Works

Scoring mirrors the original Excel workbook exactly.

Each domain has **10 scoring rules**. A rule is `[questionIndex, answerLetter]`. When the user's answer matches a rule, that domain scores **+10 points**. Maximum score is normally **100** per domain.

**Exception:** Finance & Investment has a deliberate duplicate rule (`Q6=A` appears twice in the Excel formula), so its theoretical max is **110**.

### Score Ratings

| Score | Rating |
|-------|--------|
| 70–100 | Excellent Match |
| 50–69 | Good Match |
| 30–49 | Moderate Match |
| 0–29 | Low Match |

---

## Updating Questions or Scoring

All data lives in `app.js`:

- **Questions:** Edit the `QUESTIONS` array (lines ~20–160). Each entry has `text` and `options: {A, B, C, D}`.
- **Scoring rules:** Edit the `SCORING_RULES` object. Each domain has an array of `{q: questionIndex, a: answerLetter}` objects. `q` is 0-based (Q1 = 0, Q20 = 19).
- **Domain info (roles, skills, description):** Edit the `DOMAIN_META` object.

No build step required — it's plain HTML/CSS/JS.

---

## Credits

Original Excel framework designed by **Tanmay** · [tanmay.7264@gmail.com](mailto:tanmay.7264@gmail.com)

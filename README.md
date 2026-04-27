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

### Domain → Scoring Rules Mapping

| Domain | Key Signals |
|--------|-------------|
| Consulting & Strategy | Q1–5=A, Q7=D, Q10=A, Q11=A, Q15=C, Q19=A |
| Sales & Business Dev | Q1–5=B, Q7=C, Q12=A, Q14=D, Q16=C, Q19=C |
| Marketing & Branding | Q1–5=C, Q8=B or C, Q12=C, Q17=C, Q20=C |
| Content & Creative | Q1–3,5=C, Q8=C, Q10=B or C, Q12=C, Q16=B, Q19=B |
| Finance & Investment | Q4,6(×2),10,11,13=A, Q17=B, Q18=C, Q20=B |
| Product & Growth | Q2,3=A, Q4=B, Q7=B, Q9=C, Q11=B, Q12=C, Q13=D, Q15=B, Q18=D |
| Operations & Process | Q1–5=D, Q9=A, Q11=D, Q14=A, Q16=D, Q20=D |
| Founder's Office | Q1=D, Q4,5=B, Q9=B/C, Q11=C, Q13=C, Q16=D, Q18=A, Q20=C |
| Data & Analytics | Q6=A, Q10=A/D, Q11=D, Q13=A, Q14=A, Q16=A, Q17=A, Q18=C, Q20=D |
| Human Resources | Q1=B, Q6=C, Q7=C, Q8=B, Q11=D, Q12=D, Q14=B, Q17=D, Q18=D, Q19=C |

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

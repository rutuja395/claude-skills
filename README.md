# 🧠 claude-skills

A collection of Claude Cowork skills built for real-world personal and professional use.

Install any skill by double-clicking the `.skill` file in Claude Cowork, or via the Claude Code CLI.

---

## 📦 Skills

### 💰 [indian-finance-analyser](./skills/indian-finance-analyser/)

> Your personal financial analyst — drop in any Indian bank or credit card statement and get a full spending report.

**Works with:** HDFC, Kotak, ICICI, SBI, Axis, Yes Bank, IndusInd, Federal Bank, AU, RBL, Paytm, and more.

**What it does:**
- Auto-detects PDF format for any Indian bank or credit card
- Extracts all transactions and categorises spending intelligently
- Generates a rich Excel report with 6 sheets: Dashboard, All Transactions, Category Breakdown, Monthly Trends, Top Merchants, Insights & Budget
- Gives you a personalised analyst brief in chat — not generic advice, actual insights based on your data
- Handles password-protected PDFs, multi-month statements, and duplicate detection

**How to trigger it:**
> "Analyse my statements", "here are my bank PDFs", "where is my money going this month", "give me a spending report"

**Output preview:**

| Sheet | Contents |
|---|---|
| 📊 Dashboard | KPI tiles, monthly breakdown, payment method split |
| 📁 All Transactions | Filterable table, colour-coded by category |
| 🗂️ Category Breakdown | Totals, % share, avg ticket, pie chart |
| 📅 Monthly Trends | Month-by-month, weekday vs weekend, bar chart |
| 🏪 Top Merchants | Top 20 by spend, recurring payments |
| 💡 Insights & Budget | Overspending alerts, savings tips, budget plan |

---

## 🚀 Installation

1. Download the `.skill` file from the skill's folder
2. Double-click it — Claude Cowork will install it automatically
3. Start a new Cowork session and drop in your PDF statements

---

## 🤝 Contributing

Have a skill idea? PRs welcome! Each skill lives in its own folder under `skills/` with:
```
skills/
└── your-skill-name/
    ├── SKILL.md          # Required — instructions + frontmatter
    ├── scripts/          # Optional — Python/shell scripts
    └── references/       # Optional — reference docs
```

---

## 📄 Licence

MIT — use freely, build on it, share it.

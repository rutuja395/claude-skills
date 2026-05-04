# 💰 Finance Analyser

A Claude Cowork skill that acts as your personal financial analyst — works with bank and credit card statements from any bank, anywhere in the world.

## Supported Banks & Cards

| Region | Banks |
|--------|-------|
| 🇮🇳 India | HDFC Bank, Kotak Mahindra, ICICI Bank, SBI, Axis Bank |
| 🇬🇧 UK | Barclays, HSBC, Monzo, Revolut, Lloyds, NatWest |
| 🇺🇸 US | Chase, Bank of America, Wells Fargo, Citi, Amex |
| 🇦🇺 Australia | CommBank, ANZ, Westpac |
| 🇨🇦 Canada | RBC, TD Bank |
| 🇦🇪 UAE | Emirates NBD, Mashreq |
| 🌍 Any other bank | Generic table parser as fallback |

## How to Use

1. Install the skill by double-clicking `finance-analyser.skill`
2. Open Claude Cowork and start a session
3. Drop your PDF bank/credit card statements into your workspace folder
4. Say: **"Analyse my statements"**

Claude will:
- Auto-detect each PDF's bank and statement format
- Extract all transactions
- Detect the currency automatically (₹, £, €, A$, AED, C$, and more)
- Categorise spending across 18 categories
- Generate a formatted Excel report with charts and insights
- Give you a personalised analyst brief with budget recommendations

## Categories

Food & Dining · Groceries · Shopping · Travel & Transport · Subscriptions · Telecom · Health & Fitness · Entertainment · Electronics · Fuel · Education · Insurance · Luxury · Personal Care · Rent & Housing · Utilities · Investments · Transfers

## Tips

- **Multiple months?** Drop all PDFs in at once — the skill handles them together and shows month-over-month trends
- **Password-protected PDF?** Tell Claude the password when asked
- **Multiple banks?** Mix PDFs from different banks — the skill auto-detects each one
- **Unknown merchants?** They land in Miscellaneous — tell Claude what they are and it'll learn for next time
- **Any currency?** Currency is auto-detected from the bank format, or pass `--currency £` to override

## Files

```
finance-analyser/
├── SKILL.md                          # Skill instructions
├── scripts/
│   ├── extract_transactions.py       # Universal PDF extractor (auto-detects bank & currency)
│   └── generate_report.py            # Excel report generator with charts
└── references/
    └── category_map.md               # Full categorisation rules (global merchants)
```

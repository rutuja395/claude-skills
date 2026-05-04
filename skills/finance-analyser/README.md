# 💰 Indian Finance Analyser

A Claude Cowork skill that acts as your personal financial analyst for Indian bank and credit card statements.

## Supported Banks & Cards

| Bank | Credit Card | Savings Account |
|---|---|---|
| HDFC Bank | ✅ | ✅ |
| Kotak Mahindra | ✅ | ✅ |
| ICICI Bank | ✅ | ✅ |
| SBI | — | ✅ |
| Axis Bank | ✅ | ✅ |
| Yes Bank | — | ✅ |
| IndusInd Bank | ✅ | — |
| Federal Bank | — | ✅ |
| AU Small Finance | — | ✅ |
| RBL Bank | ✅ | — |
| Paytm Payments Bank | ✅ | — |
| *Any other bank* | 🔄 Generic fallback | 🔄 Generic fallback |

## How to Use

1. Install the skill by double-clicking `indian-finance-analyser.skill`
2. Open Claude Cowork and start a session
3. Drop your PDF bank/credit card statements into your workspace folder
4. Say: **"Analyse my statements"**

Claude will:
- Auto-detect each PDF's bank and format
- Extract all transactions
- Categorise spending across 18 categories
- Generate an Excel report
- Give you a personalised analyst brief

## Categories

Food & Dining · Groceries · Shopping · Travel & Transport · Subscriptions · Telecom · Health & Fitness · Entertainment · Electronics · Fuel · Education · Insurance · Luxury & Jewelry · Personal Care · Rent & Housing · Utilities · Investments · Transfers

## Tips

- **Multiple months?** Drop all PDFs in at once — the skill handles them together and shows month-over-month trends
- **Password-protected PDF?** Tell Claude the password when asked
- **Unknown merchants?** They land in Miscellaneous — tell Claude what they are and it'll remember for next month
- **Multiple banks?** Point Claude at a folder containing PDFs from all your banks — it figures out the rest

## Files

```
indian-finance-analyser/
├── SKILL.md                          # Skill instructions
├── scripts/
│   ├── extract_transactions.py       # Universal PDF extractor (auto-detects format)
│   └── generate_report.py            # Excel report generator
└── references/
    └── category_map.md               # Full categorisation rules
```

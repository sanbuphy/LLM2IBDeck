# IBDeck JSON Spec Schema

The generator accepts a JSON file with this structure.

```json
{
  "title": "Deck title",
  "subtitle": "Optional subtitle",
  "theme": "ib-classic",
  "footer": "Confidential | Draft",
  "slides": [
    {
      "type": "title",
      "title": "Deck title",
      "subtitle": "Subtitle"
    },
    {
      "type": "summary",
      "title": "Action title",
      "bullets": ["Message 1", "Message 2"],
      "source": "Source: ..."
    },
    {
      "type": "two_column",
      "title": "Action title",
      "columns": [
        {"heading": "Column A", "bullets": ["Point"]},
        {"heading": "Column B", "bullets": ["Point"]}
      ],
      "source": "Source: ..."
    },
    {
      "type": "table",
      "title": "Action title",
      "headers": ["Metric", "2024", "2025E"],
      "rows": [["Revenue", "$100m", "$130m"]],
      "source": "Source: ..."
    },
    {
      "type": "chart",
      "title": "Action title",
      "chart_type": "bar",
      "categories": ["2024", "2025E"],
      "series": [{"name": "Revenue", "values": [100, 130]}],
      "source": "Source: ..."
    },
    {
      "type": "matrix",
      "title": "Action title",
      "headers": ["Driver", "Evidence", "Impact"],
      "rows": [["Demand", "High growth", "High"]],
      "source": "Source: ..."
    },
    {
      "type": "roadmap",
      "title": "Action title",
      "phases": [
        {"name": "Phase 1", "period": "0-3 months", "items": ["Validate demand"]}
      ],
      "source": "Source: ..."
    },
    {
      "type": "section",
      "title": "Appendix"
    }
  ]
}
```

Unknown slide types fall back to a generic content slide with title and bullets.

Bundled example specs live in `ibdeck/assets/`:

- `sample_market_scan.json`
- `finance_mna_target_profile.json`
- `finance_ipo_investor_update.json`
- `finance_industry_research.json`

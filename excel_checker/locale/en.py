"""English translations for Excel-Reifecheck."""

STRINGS: dict[str, str] = {

    # =====================================================================
    # Models – Enum display names
    # =====================================================================
    "cat.structure": "Structure & Normal Form",
    "cat.formula": "Formulas & References",
    "cat.volume": "Volume & Limits",

    "rectype.db_migration": "Database Migration",
    "rectype.power_bi": "Power BI / Reporting Tool",
    "rectype.data_warehouse": "Data Warehouse",
    "rectype.sharepoint_list": "SharePoint List",
    "rectype.normalization": "Data Normalization",
    "rectype.split_workbook": "Split Workbook",
    "rectype.cleanup": "Data Cleanup",
    "rectype.ok": "Excel Usage Acceptable",

    # =====================================================================
    # CLI
    # =====================================================================
    "cli.description": (
        "Excel-Reifecheck – Data maturity check for Excel, "
        "the first step toward AI readiness."
    ),
    "cli.files_help": "One or more Excel files (.xlsx) to check.",
    "cli.html_help": "Write the HTML report to this file.",
    "cli.open_help": "Open the HTML report in a browser after creation.",
    "cli.quiet_help": "Output only the score.",
    "cli.lang_help": "Output language (de/en).",
    "cli.file_not_found": "File not found: {path}",
    "cli.unsupported": "Only .xlsx/.xlsm files are supported: {path}",
    "cli.health_check": "📊 Excel Health Check: {filename}",
    "cli.score": "{icon} Score: {score}/100",
    "cli.file_size": "   File size: {size} MB",
    "cli.sheets": "   Sheets: {count}",
    "cli.analysis_time": "   Analysis time: {elapsed}s",
    "cli.findings_header": "--- Findings ({count}) ---",
    "cli.rec_header": "--- Strategic Recommendations ---",
    "cli.rec_priority": "  📌 Priority {priority}: {title}",
    "cli.rec_reason": "     Reason: {reason}",
    "cli.rec_action": "     Action: {action}",
    "cli.html_report": "📄 HTML report: {path}",

    # =====================================================================
    # Engine – Progress messages
    # =====================================================================
    "engine.file_not_found": "File not found: {path}",
    "engine.loading": "📂 Loading file{mode}",
    "engine.loading_detail": "Reading {size}…",
    "engine.stats_queue": "📊 Collecting sheet statistics",
    "engine.analyzing": "📊 Analyzing sheets",
    "engine.analyzing_detail": "{count} sheets: {names}",
    "engine.sheet_progress": "📊 Sheet {idx}/{count}: {title}",
    "engine.sheet_detail": "Analyzing…",
    "engine.volume_light_queue": "📋 Volume analysis (light mode)",
    "engine.volume_light": "📋 Volume analysis",
    "engine.volume_light_detail": "Light mode – metadata-based",
    "engine.rule_queue": "🔍 {name}",
    "engine.rule_checking": "🔍 {name}",
    "engine.rule_detail": "Rule {idx}/{count} · {rule_id}",
    "engine.rule_error": "⚠️ {name}",
    "engine.rule_error_detail": "Error: {error}",
    "engine.recs_queue": "🎯 Generating recommendations",
    "engine.recs_progress": "🎯 Generating recommendations",
    "engine.done": "✅ Analysis complete",
    "engine.done_detail": "Health score: {score}/100 · {findings} findings",
    "engine.light_hint": " (light mode for large files)",
    # Light mode findings
    "engine.vol_extreme_msg": "File size of {size:.0f} MB is extremely large.",
    "engine.vol_extreme_detail": "Files over 50 MB cause massive performance problems.",
    "engine.vol_extreme_tip": "Urgently migrate to a database or Power BI.",
    "engine.vol_large_msg": "File size of {size:.0f} MB is very large.",
    "engine.vol_large_detail": "At this size, Excel becomes unstable and slow.",
    "engine.vol_large_tip": "Migration to a database or splitting the file is recommended.",
    "engine.vol_rows_abused": "{rows} rows – Excel is being misused as a database.",
    "engine.vol_rows_abused_tip": "For this amount of data, a real database is the better choice.",
    "engine.vol_rows_high": "{rows} rows – high data volume for an Excel file.",
    "engine.vol_rows_high_tip": "Consider a SharePoint list or database as an alternative.",
    "engine.vol_many_sheets": "{count} worksheets – very complex file.",
    "engine.vol_many_sheets_tip": "Split into multiple files or migrate to a structured solution.",
    "engine.light_msg": (
        "Light analysis: File is {size:.1f} MB – "
        "detailed checks of formulas, colors, and structure were skipped."
    ),
    "engine.light_detail": (
        "For files over 15 MB, only the volume analysis is performed "
        "to avoid long waiting times."
    ),
    "engine.light_tip": "For a full analysis, reduce the file size to under 15 MB.",

    # =====================================================================
    # Rules – STR (Structure)
    # =====================================================================
    "STR-001.name": "Merged Cells",
    "STR-001.msg": (
        "Automation potential detected: {count} merged cell ranges "
        "prevent automated data processing."
    ),
    "STR-001.detail": "Ranges: {ranges}",
    "STR-001.tip": (
        "Tip: Replace merged cells with 'Center Across Selection' – "
        "it looks the same but keeps the data machine-readable "
        "and ready for dashboards or Power BI."
    ),

    "STR-002.name": "Data Type Homogeneity",
    "STR-002.msg": (
        "Column {col} contains mixed data types – "
        "mainly {dominant} ({pct}), "
        "but also {minority}. This makes automated analysis harder."
    ),
    "STR-002.tip": (
        "Tip: If column {col} is homogeneous, it can be "
        "automatically imported into dashboards, pivot tables, or databases. "
        "Special values like 'N/A' are best moved to a separate status column."
    ),

    "STR-003.name": "Header Detection",
    "STR-003.msg.none": "No clear header row detected – this makes automated processing harder.",
    "STR-003.tip.none": (
        "Tip: A clear header in row 1 makes the data immediately "
        "usable for filters, pivot tables, and automated analysis."
    ),
    "STR-003.msg.late": "Header row appears to start at row {row}.",
    "STR-003.tip.late": (
        "Tip: Move meta information (title, date, author) to a separate "
        "info sheet. Then the data can start directly from row 1 and is "
        "immediately processable."
    ),
    "STR-003.msg.dupes": "Duplicate column headers: {dupes}",
    "STR-003.tip.dupes": (
        "Tip: Unique column names are essential for automated analysis. "
        "Duplicate names cause confusion in formulas and tools."
    ),

    "STR-004.name": "Empty Separator Rows/Columns",
    "STR-004.msg": "{count} empty rows within the data range.",
    "STR-004.tip": (
        "Tip: Visual grouping works without empty rows – "
        "Excel grouping, conditional formatting, or separate sheets "
        "keep the data contiguous and analyzable."
    ),

    "STR-005.name": "Identifier Consistency",
    "STR-005.msg.inconsistent": (
        "Column {col}: Inconsistent ID numbering "
        "for prefix '{prefix}{sep}'."
    ),
    "STR-005.detail.inconsistent": (
        "Different digit widths: {widths}. "
        "Examples: {examples}"
    ),
    "STR-005.tip.inconsistent": (
        "Tip: A consistent format like '{prefix}{sep}001', "
        "'{prefix}{sep}002' (same digit count) makes sorting easier "
        "and prevents mix-ups in references."
    ),
    "STR-005.msg.gaps": "Column {col}: Gaps in the ID sequence '{prefix}{sep}...'.",
    "STR-005.detail.gaps": "Missing numbers: {missing}",
    "STR-005.tip.gaps": (
        "Note: Gaps may indicate deleted entries. "
        "It's worth checking if data is missing here."
    ),
    "STR-005.msg.separators": (
        "Column {col}: Prefix '{prefix}' uses "
        "different separators: {seps}."
    ),
    "STR-005.tip.separators": "Use a consistent separator, e.g. always '{prefix}-'.",
    "STR-005.msg.dupes": "Column {col}: Duplicate identifiers found.",
    "STR-005.detail.dupes": "Duplicates: {dupes}",
    "STR-005.tip.dupes": (
        "Tip: Duplicate IDs can cause assignment errors. "
        "Check which entry is the correct one."
    ),

    "STR-006.name": "Missing Primary Key",
    "STR-006.msg": (
        "No column with unique values found – "
        "this table has no reliable key "
        "for identifying individual entries."
    ),
    "STR-006.detail": (
        "{rows} data rows in {cols} columns, "
        "but no value is unique per row."
    ),
    "STR-006.tip": (
        "This is a real risk: Without a unique key, "
        "duplicates cannot be detected, data cannot be safely linked, and "
        "changes cannot be tracked. Recommendation: Add a sequential "
        "number or structured code (e.g. PRJ-001) as the first column."
    ),

    "STR-007.name": "Free-text IDs (not sortable)",
    "STR-007.msg": (
        "Column {col} ('{header}') is used as an ID, "
        "but contains free text instead of structured codes."
    ),
    "STR-007.detail": (
        "Examples: {examples}. "
        "{ratio} of values are long free text."
    ),
    "STR-007.tip": (
        "Free-text IDs (e.g. 'Project Vienna Center Renovation Q4') "
        "are not sortable, can easily diverge, and make links unreliable. "
        "Better: Use short, structured codes like 'PRJ-0042' "
        "and put the free text in a separate description column."
    ),

    # =====================================================================
    # Rules – FRM (Formulas)
    # =====================================================================
    "FRM-001.name": "Absolute vs. Relative References",
    "FRM-001.msg": (
        "High proportion of fixed references ($): {pct} of "
        "{total} references are absolute."
    ),
    "FRM-001.detail": "Absolute: {absolute}, Relative: {relative}, Mixed: {mixed}",
    "FRM-001.tip": (
        "Tip: Many fixed references indicate formulas that are hard "
        "to copy and maintain. Named ranges (e.g. 'VAT' instead of "
        "'$B$1') make formulas readable and flexible."
    ),

    "FRM-002.name": "Volatile Functions (Performance)",
    "FRM-002.msg": (
        "Performance bottleneck detected: Function {func}() is used "
        "in {count} cells and forces a complete recalculation "
        "with every change."
    ),
    "FRM-002.detail": "Cells (sample): {cells}",
    "FRM-002.tip": (
        "Tip: {func}() is a 'volatile' function – Excel "
        "recalculates it with EVERY change, even when nothing "
        "relevant has changed. Alternatives: INDEX/MATCH instead of "
        "OFFSET, fixed date values instead of TODAY()."
    ),

    "FRM-003.name": "Lookup Intensity",
    "FRM-003.msg": (
        "This sheet simulates a database: {count} "
        "lookup functions link data like SQL JOINs, only slower."
    ),
    "FRM-003.detail": "Functions: {summary}",
    "FRM-003.tip": (
        "Tip: This many lookups indicate that the data actually "
        "belongs together. A relational data structure (Power Query, "
        "SharePoint list, or database) would automate these links "
        "and make the file significantly faster."
    ),

    "FRM-004.name": "Circular Reference Hints",
    "FRM-004.msg": "Possible circular references in {count} cells detected.",
    "FRM-004.detail": "Suspicious cells: {cells}",
    "FRM-004.tip": (
        "Note: Circular references can lead to unpredictable calculations. "
        "They usually arise from copy-paste errors. "
        "Please check the highlighted cells."
    ),

    "FRM-005.name": "Cross-Sheet & Cross-File References",
    "FRM-005.msg.external": (
        "External file references detected: {count} "
        "formulas reference other Excel files."
    ),
    "FRM-005.detail.external": "Referenced files: {files}",
    "FRM-005.tip.external": (
        "Note: External references make the file dependent on other "
        "files – if they are moved or renamed, the formulas break. "
        "Better: Manage data centrally (e.g. Power Query or SharePoint) "
        "and query from there."
    ),
    "FRM-005.msg.cross_sheet": (
        "Complex networking: This sheet references "
        "{count} other sheets."
    ),
    "FRM-005.detail.cross_sheet": "Referenced sheets: {sheets}",
    "FRM-005.tip.cross_sheet": (
        "FYI: Heavily networked sheets can become confusing. "
        "A central data sheet as a 'single source of truth' can "
        "simplify the structure."
    ),

    # =====================================================================
    # Rules – VOL (Volume)
    # =====================================================================
    "VOL-001.name": "Data Volume",
    "VOL-001.threshold.100k": "has outgrown Excel",
    "VOL-001.threshold.50k": "is ready for a database upgrade",
    "VOL-001.threshold.10k": "is approaching the range where a database would be more efficient",
    "VOL-001.msg": (
        "Congratulations – this dataset has grown! "
        "{rows} rows: This data {label}."
    ),
    "VOL-001.detail": "Rows: {rows}, Columns: {cols}",
    "VOL-001.tip": (
        "Tip: At this size, databases (SharePoint list, "
        "Dataverse, SQL) offer clear advantages: faster search, "
        "concurrent access, and automatic backups. "
        "We're happy to help with the transition!"
    ),

    "VOL-002.name": "Formula Density",
    "VOL-002.msg.high": (
        "High computational load: {count} formulas "
        "({pct} of all cells). This can "
        "noticeably slow down the file."
    ),
    "VOL-002.detail": "Formulas: {formulas}, Static values: {statics}",
    "VOL-002.tip.high": (
        "Tip: Formulas that no longer change can be pasted as values "
        "(Copy → 'Paste Values'). Even better: Move calculations to "
        "Power Query or a database view."
    ),
    "VOL-002.msg.medium": (
        "{count} formulas in this sheet – "
        "still OK, but keep an eye on it."
    ),
    "VOL-002.tip.medium": (
        "Tip: As data grows, regularly reconsider the calculation logic. "
        "Power Query can replace many formulas."
    ),

    "VOL-003.name": "Number of Worksheets",
    "VOL-003.msg": (
        "This workbook contains {count} worksheets – "
        "indicating a complex data model that has "
        "outgrown Excel."
    ),
    "VOL-003.detail": "Sheets: {sheets}",
    "VOL-003.tip": (
        "Tip: Many linked sheets essentially model a relational "
        "database – just without its benefits. Migration to a real "
        "data structure (Dataverse, SQL) with Power BI as frontend "
        "offers the same overview, but more stable and faster."
    ),

    "VOL-004.name": "File Size",
    "VOL-004.msg.critical": (
        "The file is {size} MB – this causes long "
        "load times and hinders collaboration."
    ),
    "VOL-004.tip.critical": (
        "Urgent action needed: Files this size belong in a "
        "professional system. We're happy to advise on the best solution "
        "(database, data warehouse, SharePoint)."
    ),
    "VOL-004.msg.warning": (
        "The file is {size} MB – this can "
        "noticeably affect loading time."
    ),
    "VOL-004.tip.warning": (
        "Tip: Archive unused data, compress images, "
        "or paste calculations as values to reduce file size."
    ),

    # =====================================================================
    # Rules – IMP (Implicit Knowledge)
    # =====================================================================
    "IMP-001.name": "Undocumented Color Codes",
    "IMP-001.msg.many": (
        "Implicit knowledge: {unique} different background colors "
        "in {total} cells – but no documented legend found."
    ),
    "IMP-001.detail": "Colors: {colors}",
    "IMP-001.tip.many": (
        "Tip: Color codes often contain important business logic "
        "(e.g. traffic light status, priorities) that only the creator knows. "
        "A legend or better: a dedicated status column makes this "
        "knowledge accessible to everyone and automatically analyzable."
    ),
    "IMP-001.msg.some": (
        "{unique} different background colors in use – "
        "if they have meaning, a legend would be helpful."
    ),
    "IMP-001.tip.some": (
        "Tip: If colors represent status or categories, "
        "create a dedicated column – this makes the data "
        "filterable and analyzable."
    ),
    "IMP-001.msg.font": (
        "{unique} different font colors in use – "
        "they could be implicit categorizations."
    ),
    "IMP-001.tip.font": (
        "Tip: Font colors as meaning carriers are invisible "
        "in print and in analysis. Better: Create a dedicated column "
        "for the classification."
    ),

    "IMP-002.name": "Hidden Sheets",
    "IMP-002.msg.very_hidden": (
        "{count} sheet(s) are 'very hidden' "
        "(visible only via VBA) – this poses risks."
    ),
    "IMP-002.detail": "Sheets: {sheets}",
    "IMP-002.tip.very_hidden": (
        "Note: 'Very hidden' sheets often contain critical "
        "calculations or master data that only the creator knows. "
        "This is a knowledge risk – please review and document."
    ),
    "IMP-002.msg.hidden": "{count} hidden sheet(s) found.",
    "IMP-002.tip.hidden": (
        "Tip: Hidden sheets often contain helper tables or "
        "master data. Please check if the content is still current "
        "and whether it should be managed centrally."
    ),

    "IMP-003.name": "Hidden Rows/Columns",
    "IMP-003.msg": (
        "Hidden areas: {rows} rows and "
        "{cols} columns are hidden."
    ),
    "IMP-003.tip": (
        "Tip: Hidden areas often contain outdated or "
        "sensitive data. Better: Delete unneeded data or "
        "move it to an archive sheet. Sensitive data belongs in "
        "a protected system."
    ),

    "IMP-004.name": "Conditional Formatting (Implicit Logic)",
    "IMP-004.msg": (
        "Business logic hidden in formatting: "
        "{count} conditional formatting rules – these "
        "rules are hard to maintain and invisible to others."
    ),
    "IMP-004.tip": (
        "Tip: Conditional formatting often encodes important "
        "business rules (e.g. 'red when overdue'). "
        "Better: Create a dedicated status column with formulas – "
        "then the logic is visible, documentable, and analyzable."
    ),

    "IMP-005.name": "Hardcoded Values in Formulas",
    "IMP-005.msg": (
        "Implicit knowledge in formulas: {count} "
        "hardcoded values repeat across formulas."
    ),
    "IMP-005.detail": "Values: {values}",
    "IMP-005.tip": (
        "Tip: Recurring numbers in formulas (e.g. tax rates, "
        "factors, limits) should be defined as named ranges or in "
        "a parameter table. This way, only ONE place needs updating."
    ),

    "IMP-006.name": "Hardcoded Validation Lists",
    "IMP-006.msg": (
        "{count} dropdown lists with manually typed values – "
        "changes must be manually applied to each list."
    ),
    "IMP-006.tip": (
        "Tip: Maintain dropdown values in a separate master data table "
        "and link via range reference. Then only ONE place needs "
        "updating for changes."
    ),

    "IMP-007.name": "Non-descriptive Sheet Names",
    "IMP-007.msg": "Non-descriptive sheet names: {names}",
    "IMP-007.tip": (
        "Tip: Descriptive sheet names like 'Revenue_2024' or "
        "'Customers_Master' help everyone navigate quickly. "
        "'Sheet1' tells nobody anything."
    ),

    "IMP-008.name": "Business Logic in Comments",
    "IMP-008.msg": (
        "Business logic in cell comments detected: "
        "{count} comments contain "
        "cues like 'Attention', 'Do not change' etc."
    ),
    "IMP-008.detail": "Cells: {cells}",
    "IMP-008.tip": (
        "Note: When critical rules exist only in comments, "
        "they are easily lost or overlooked. "
        "Better: Record business rules as data validation, conditional "
        "formatting, or in a documentation sheet."
    ),

    "IMP-009.name": "Misleading Number Formats",
    "IMP-009.msg": (
        "Special number formats detected that can make the displayed "
        "value differ from the actual value."
    ),
    "IMP-009.detail": "Formats: {formats}",
    "IMP-009.tip": (
        "FYI: Number formats that embed text or visually alter values "
        "can cause misunderstandings "
        "(e.g. '1' is displayed, but '1000' is stored). "
        "Please check the critical cells."
    ),

    "IMP-010.name": "Sheet Protection & Locked Areas",
    "IMP-010.msg": "{count} sheet(s) are protected: {sheets}",
    "IMP-010.tip": (
        "FYI: Sheet protection indicates that certain areas are critical. "
        "Good! But: Document WHY and WHAT is protected. "
        "If the creator leaves the company, sheet protection "
        "can become a problem."
    ),

    # =====================================================================
    # Recommendations
    # =====================================================================
    "rec.db.title": "This data is ideal for a database solution",
    "rec.db.action": (
        "Recommended next step: Schedule a meeting with the Digital Workplace team. "
        "We'll analyze together which database solution fits best "
        "(SharePoint list, Dataverse, SQL Server) and guide the migration. "
        "Existing Excel reports can continue as Power BI dashboards."
    ),
    "rec.db.reason.rows_high": "{rows} rows exceed the practical Excel range",
    "rec.db.reason.rows_medium": "{rows} rows are approaching the range where databases are more efficient",
    "rec.db.reason.lookups": "Excessive lookup functions simulate database JOINs",
    "rec.db.reason.size": "File size ({size} MB) causes performance issues",
    "rec.db.reason.sheets": "{count} sheets model a relational data model",
    "rec.db.reason.external": "External file references indicate distributed, related data",

    "rec.powerbi.title": "Move calculations and visualizations to a reporting tool",
    "rec.powerbi.reason": (
        "{formulas} formulas and/or complex conditional formatting "
        "indicate elaborate analysis logic that can be more stably and "
        "flexibly implemented in Power BI or a comparable tool."
    ),
    "rec.powerbi.action": (
        "Power BI can access the Excel data directly and display "
        "analyses dynamically. Advantage: The calculation logic "
        "is documented, traceable by others, and can be adjusted "
        "without Excel expertise."
    ),

    "rec.sharepoint.title": "SharePoint list as a low-threshold alternative",
    "rec.sharepoint.reason": (
        "The data structure ({rows} rows, {sheets} sheets) "
        "is well suited for a SharePoint list – easy to create, "
        "with version history and access control."
    ),
    "rec.sharepoint.action": (
        "A SharePoint list can be created from an Excel table in minutes "
        "(Excel → 'Export as Table'). "
        "Benefits: Concurrent access, change history, automatic "
        "notifications, and Power Automate integration."
    ),

    "rec.normalization.title": "Clean up data structure – foundation for all next steps",
    "rec.normalization.reason": (
        "{count} structural issues found (merged cells, "
        "mixed data types, missing headers, empty row separators). "
        "These prevent any form of automated processing."
    ),
    "rec.normalization.action": (
        "First step: Convert the data to tabular format "
        "(one header row, no merged cells, no empty rows). "
        "The Digital Workplace team can assist with cleanup – "
        "often a single workshop session is enough."
    ),

    "rec.split.title": "Split workbook",
    "rec.split.reason": (
        "{count} sheets in one file are hard to oversee. "
        "Often different sheets contain different topics that "
        "are better managed separately."
    ),
    "rec.split.action": (
        "Check which sheets belong together thematically and which "
        "are independent datasets. Move independent areas to "
        "separate files or SharePoint lists."
    ),

    "rec.cleanup.title": "Document implicit knowledge and make it explicit",
    "rec.cleanup.reason": (
        "{count} cases of undocumented knowledge found "
        "(color codes without legend, hidden areas, logic in comments). "
        "This knowledge is lost when the creator is unavailable."
    ),
    "rec.cleanup.action": (
        "Strongly recommended: Schedule a 30-minute meeting with the creator "
        "to document the implicit rules. "
        "Color codes → status column, comment rules → data validation, "
        "hidden sheets → documentation or deletion."
    ),

    "rec.ok.title": "Good work – Excel usage is acceptable here",
    "rec.ok.reason": (
        "Score {score}/100. "
        "The file is well structured and within a reasonable scope for Excel."
    ),
    "rec.ok.action": (
        "No urgent actions needed. "
        "For the future: Regularly check whether the data volume is growing."
    ),
    "rec.minor.title": "Minor optimizations recommended",
    "rec.minor.reason": "Score {score}/100.",
    "rec.minor.action": (
        "Follow the individual hints in the report – "
        "most can be implemented with little effort."
    ),

    # =====================================================================
    # Report – Labels & grades
    # =====================================================================
    "report.grade.mega": "Critical Issue",
    "report.grade.real": "Real Problem",
    "report.grade.cosmetic": "Cosmetic Issue",

    "report.score.good": "Good",
    "report.score.needs_improvement": "Needs Improvement",
    "report.score.action_needed": "Action Needed",
    "report.score.urgent": "Urgent Action Needed",

    "report.dim.volume": "Data\nVolume",
    "report.dim.formulas": "Formula\nComplexity",
    "report.dim.networking": "Networking",
    "report.dim.implicit": "Implicit\nKnowledge",
    "report.dim.structure": "Structural\nIssues",
    "report.dim.filesize": "File\nSize",

    # Anti-pattern cards
    "ap.STR-005.name": "ID Sprawl",
    "ap.STR-005.desc": "Inconsistent identifiers (formats, gaps, duplicates).",
    "ap.STR-006.name": "Phantom Key",
    "ap.STR-006.desc": "Data table without a unique, sortable key column.",
    "ap.STR-007.name": "Free-text IDs",
    "ap.STR-007.desc": "IDs from free text instead of sortable codes – not machine-processable.",
    "ap.FRM-004.name": "Circular Reference",
    "ap.FRM-004.desc": "Formulas reference themselves – unpredictable results.",
    "ap.VOL-001.name": "Excel Breaker",
    "ap.VOL-001.desc": "Data volume has exceeded Excel's practical capacity.",
    "ap.VOL-002.name": "Formula Overload",
    "ap.VOL-002.desc": "Too many formulas – instability and slowness.",
    "ap.STR-001.name": "Merged Cells Chaos",
    "ap.STR-001.desc": "Merged cells destroy machine readability.",
    "ap.STR-002.name": "Type Salad",
    "ap.STR-002.desc": "Mixed data types in a column – 1st normal form violated.",
    "ap.STR-003.name": "Headless Table",
    "ap.STR-003.desc": "Missing or unclear headers.",
    "ap.FRM-002.name": "Volatile Trap",
    "ap.FRM-002.desc": "Volatile functions force permanent recalculation.",
    "ap.FRM-003.name": "VLOOKUP Database",
    "ap.FRM-003.desc": "Lookup chains simulate SQL JOINs – much slower.",
    "ap.FRM-005.name": "File Spider Web",
    "ap.FRM-005.desc": "External links – fragile dependencies.",
    "ap.IMP-001.name": "Secret Traffic Light",
    "ap.IMP-001.desc": "Color codes carry undocumented meaning.",
    "ap.IMP-002.name": "Hide & Seek",
    "ap.IMP-002.desc": "Hidden sheets with undocumented knowledge.",
    "ap.IMP-004.name": "Format Logic",
    "ap.IMP-004.desc": "Business logic hidden in conditional formatting.",
    "ap.IMP-005.name": "Magic Numbers",
    "ap.IMP-005.desc": "Hardcoded numbers in formulas without explanation.",
    "ap.VOL-003.name": "Sheet Hydra",
    "ap.VOL-003.desc": "Too many sheets model a relational DB – without its benefits.",
    "ap.VOL-004.name": "Heavyweight",
    "ap.VOL-004.desc": "Oversized file size.",
    "ap.STR-004.name": "Blank Row Layout",
    "ap.STR-004.desc": "Empty rows as visual separators fragment the data range.",
    "ap.FRM-001.name": "Dollar Fixation",
    "ap.FRM-001.desc": "Too many rigid $-references – hard to maintain.",
    "ap.IMP-003.name": "Invisible Data",
    "ap.IMP-003.desc": "Hidden rows/columns.",
    "ap.IMP-006.name": "Typed Dropdowns",
    "ap.IMP-006.desc": "Validation lists with hardcoded values.",
    "ap.IMP-007.name": "Nameless Sheets",
    "ap.IMP-007.desc": "Generic sheet names without meaning.",
    "ap.IMP-008.name": "Comment Knowledge",
    "ap.IMP-008.desc": "Business rules live only in cell comments.",
    "ap.IMP-009.name": "Number Masquerade",
    "ap.IMP-009.desc": "Misleading number formats.",
    "ap.IMP-010.name": "Sheet Protection Mystery",
    "ap.IMP-010.desc": "Protected areas without documentation.",

    # AI-Readiness blockers
    "report.ai.merged": "{count} merged cell ranges",
    "report.ai.merged_desc": "AI cannot recognize tabular structure",
    "report.ai.implicit": "{count}× implicit knowledge",
    "report.ai.implicit_desc": "Color codes, hidden sheets, magic numbers – invisible to AI",
    "report.ai.structure": "{count}× structural issue",
    "report.ai.structure_desc": "Mixed types, missing IDs – AI cannot associate records",
    "report.ai.formulas": "{count}× formula chaining",
    "report.ai.formulas_desc": "VLOOKUP chains & circular references block automated analysis",
    "report.ai.weak": "Structural weaknesses",
    "report.ai.weak_desc": "Data is not in machine-readable form",

    # =====================================================================
    # LLM Analysis
    # =====================================================================
    "llm.no_key": "No API key provided.",
    "llm.no_endpoint": "No endpoint provided. Please configure the Azure AI endpoint URL.",
    "llm.key_invalid": "API key invalid (401 Unauthorized).",
    "llm.access_denied": "Access denied (403 Forbidden).",
    "llm.timeout": "Timeout – server not reachable.",
    "llm.connection_error": "Connection error – endpoint not reachable.",
    "llm.error": "Error: {error}",
    "llm.key_works": "API key works! ✅",
    "llm.parse_error": "The AI analysis could not be fully parsed.",

    # =====================================================================
    # Webapp – Server-side messages
    # =====================================================================
    "web.no_file": "No file selected.",
    "web.unsupported": "Only .xlsx and .xlsm files are supported.",
    "web.no_url": "Please enter a URL.",
    "web.invalid_url": "Please enter a valid https:// URL.",
    "web.session_not_found": "Session not found.",
    "web.report_not_found": "Report not found.",
    "web.report_expired": "The report has expired or was already displayed. Please start a new analysis.",
    "web.error_title": "Error",
    "web.no_file_title": "No File",
    "web.no_file_msg": "Please select an Excel file.",
    "web.wrong_format_title": "Wrong Format",
    "web.analysis_failed_title": "Analysis Failed",
    "web.analysis_failed_msg": "The file could not be analyzed: {error}",
    "web.download_failed_title": "Download Failed",
    "web.invalid_url_title": "Invalid URL",
    "web.no_url_title": "No URL",
    "web.no_url_msg": "Please enter a URL.",
    "web.connection_lost": "Connection to server lost.",
    "web.report_not_found_llm": "Report not found. Please run a standard analysis first.",
    "web.no_key_settings": "No API key provided. Please configure under ⚙️ AI Settings.",
    "web.no_endpoint_settings": "No endpoint configured. Please enter the Azure AI endpoint URL under ⚙️ AI Settings.",
    "web.llm_failed": "LLM analysis failed: {error}",
    "web.login_required": (
        "The link requires authentication. Please create the link via "
        "'Share → Copy Link' with 'Anyone with the link', "
        "or upload the file directly."
    ),
    "web.file_too_large": "The file is larger than 100 MB.",
    "web.not_excel": (
        "The downloaded file is not a valid Excel file. "
        "The link may point to a login page. "
        "Please upload the file directly instead."
    ),
    "web.timeout": "Timeout: The server did not respond in time.",
    "web.connect_error": "Connection error: The server is not reachable.",
    "web.access_denied": (
        "Access denied. Please create a sharing link with "
        "'Anyone with the link', or upload the file directly."
    ),
    "web.http_error": "HTTP error {status} while downloading.",
    "web.download_error": "Error while downloading: {error}",
    "web.back_to_upload": "← Back to Upload",
}

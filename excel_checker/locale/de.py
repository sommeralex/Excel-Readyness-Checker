"""Deutsche Übersetzungen für den Excel-Reifecheck (Standardsprache)."""

# Alle nutzer-sichtbaren Strings, organisiert nach Modul.
# Schlüssel verwenden Punkt-Notation: modul.kontext oder REGEL-ID.kontext

STRINGS: dict[str, str] = {

    # =====================================================================
    # Models – Enum-Anzeigenamen
    # =====================================================================
    "cat.structure": "Struktur & Normalform",
    "cat.formula": "Formeln & Bezüge",
    "cat.volume": "Volumen & Limits",

    "rectype.db_migration": "Datenbank-Migration",
    "rectype.power_bi": "Power BI / Reporting-Tool",
    "rectype.data_warehouse": "Data Warehouse",
    "rectype.sharepoint_list": "SharePoint-Liste",
    "rectype.normalization": "Daten-Normalisierung",
    "rectype.split_workbook": "Workbook aufteilen",
    "rectype.cleanup": "Datenbereinigung",
    "rectype.ok": "Excel-Nutzung vertretbar",

    # =====================================================================
    # CLI
    # =====================================================================
    "cli.description": (
        "Excel-Reifecheck – Datenreife-Check für Excel, "
        "der erste Schritt Richtung AI-Readiness."
    ),
    "cli.files_help": "Eine oder mehrere Excel-Dateien (.xlsx) zum Prüfen.",
    "cli.html_help": "HTML-Bericht in diese Datei schreiben.",
    "cli.open_help": "HTML-Bericht nach Erstellung im Browser öffnen.",
    "cli.quiet_help": "Nur die Bewertung ausgeben.",
    "cli.lang_help": "Sprache für Ausgaben (de/en).",
    "cli.file_not_found": "Datei nicht gefunden: {path}",
    "cli.unsupported": "Nur .xlsx/.xlsm Dateien werden unterstützt: {path}",
    "cli.health_check": "📊 Excel Health Check: {filename}",
    "cli.score": "{icon} Bewertung: {score}/100",
    "cli.file_size": "   Dateigröße: {size} MB",
    "cli.sheets": "   Blätter: {count}",
    "cli.analysis_time": "   Analyse-Zeit: {elapsed}s",
    "cli.findings_header": "--- Ergebnisse ({count}) ---",
    "cli.rec_header": "--- Strategische Empfehlungen ---",
    "cli.rec_priority": "  📌 Priorität {priority}: {title}",
    "cli.rec_reason": "     Grund: {reason}",
    "cli.rec_action": "     Maßnahme: {action}",
    "cli.html_report": "📄 HTML-Bericht: {path}",

    # =====================================================================
    # Engine – Fortschrittsmeldungen
    # =====================================================================
    "engine.file_not_found": "Datei nicht gefunden: {path}",
    "engine.loading": "📂 Datei wird geladen{mode}",
    "engine.loading_detail": "{size} wird eingelesen…",
    "engine.stats_queue": "📊 Sheet-Statistiken sammeln",
    "engine.analyzing": "📊 Sheets werden analysiert",
    "engine.analyzing_detail": "{count} Sheets: {names}",
    "engine.sheet_progress": "📊 Sheet {idx}/{count}: {title}",
    "engine.sheet_detail": "Wird analysiert…",
    "engine.volume_light_queue": "📋 Volumen-Analyse (Light-Modus)",
    "engine.volume_light": "📋 Volumen-Analyse",
    "engine.volume_light_detail": "Light-Modus – Metadaten-basiert",
    "engine.rule_queue": "🔍 {name}",
    "engine.rule_checking": "🔍 {name}",
    "engine.rule_detail": "Regel {idx}/{count} · {rule_id}",
    "engine.rule_error": "⚠️ {name}",
    "engine.rule_error_detail": "Fehler: {error}",
    "engine.recs_queue": "🎯 Empfehlungen generieren",
    "engine.recs_progress": "🎯 Empfehlungen werden generiert",
    "engine.done": "✅ Analyse abgeschlossen",
    "engine.done_detail": "Health-Score: {score}/100 · {findings} Findings",
    "engine.light_hint": " (Light-Modus für große Dateien)",
    # Light-Modus Findings
    "engine.vol_extreme_msg": "Dateigröße von {size:.0f} MB ist extrem groß.",
    "engine.vol_extreme_detail": "Dateien über 50 MB verursachen massive Performance-Probleme.",
    "engine.vol_extreme_tip": "Dringend in eine Datenbank oder Power BI migrieren.",
    "engine.vol_large_msg": "Dateigröße von {size:.0f} MB ist sehr groß.",
    "engine.vol_large_detail": "Bei dieser Größe wird Excel instabil und langsam.",
    "engine.vol_large_tip": "Migration in eine Datenbank oder Aufteilen der Datei empfohlen.",
    "engine.vol_rows_abused": "{rows} Zeilen – Excel wird als Datenbank missbraucht.",
    "engine.vol_rows_abused_tip": "Bei dieser Datenmenge ist eine echte Datenbank die bessere Wahl.",
    "engine.vol_rows_high": "{rows} Zeilen – hohe Datenmenge für eine Excel-Datei.",
    "engine.vol_rows_high_tip": "SharePoint-Liste oder Datenbank als Alternative prüfen.",
    "engine.vol_many_sheets": "{count} Tabellenblätter – sehr komplexe Datei.",
    "engine.vol_many_sheets_tip": "Aufteilen in mehrere Dateien oder in eine strukturierte Lösung migrieren.",
    "engine.light_msg": (
        "Light-Analyse: Datei ist {size:.1f} MB groß – "
        "Detailprüfung von Formeln, Farben und Struktur wurde übersprungen."
    ),
    "engine.light_detail": (
        "Bei Dateien über 15 MB wird nur die Volumen-Analyse durchgeführt, "
        "um lange Wartezeiten zu vermeiden."
    ),
    "engine.light_tip": "Für eine vollständige Analyse die Datei auf unter 15 MB verkleinern.",

    # =====================================================================
    # Regeln – STR (Struktur)
    # =====================================================================
    "STR-001.name": "Verbundene Zellen",
    "STR-001.msg": (
        "Automatisierungs-Potenzial entdeckt: {count} verbundene Zellbereiche "
        "verhindern, dass diese Daten automatisiert ausgewertet werden können."
    ),
    "STR-001.detail": "Bereiche: {ranges}",
    "STR-001.tip": (
        "Tipp: Verbundene Zellen durch 'Über Auswahl zentrieren' ersetzen – "
        "dann sieht es gleich aus, aber die Daten werden maschinenlesbar "
        "und können in Dashboards oder Power BI fließen."
    ),

    "STR-002.name": "Datentyp-Homogenität",
    "STR-002.msg": (
        "Spalte {col} enthält unterschiedliche Datentypen – "
        "hauptsächlich {dominant} ({pct}), "
        "aber auch {minority}. Das erschwert automatische Auswertungen."
    ),
    "STR-002.tip": (
        "Tipp: Wenn Spalte {col} einheitlich ist, kann sie "
        "automatisch in Dashboards, Pivot-Tabellen oder Datenbanken "
        "übernommen werden. Sonderwerte wie 'N/A' am besten in eine "
        "eigene Status-Spalte auslagern."
    ),

    "STR-003.name": "Kopfzeilen-Erkennung",
    "STR-003.msg.none": "Keine eindeutige Kopfzeile erkannt – das erschwert automatische Verarbeitung.",
    "STR-003.tip.none": (
        "Tipp: Eine klare Kopfzeile in Zeile 1 macht die Daten sofort "
        "nutzbar für Filter, Pivot-Tabellen und automatisierte Auswertungen."
    ),
    "STR-003.msg.late": "Kopfzeile scheint erst in Zeile {row} zu beginnen.",
    "STR-003.tip.late": (
        "Tipp: Meta-Informationen (Titel, Datum, Autor) am besten "
        "in ein eigenes Info-Sheet auslagern. Dann können die Daten "
        "direkt ab Zeile 1 starten und sind sofort verarbeitbar."
    ),
    "STR-003.msg.dupes": "Doppelte Spaltenüberschriften: {dupes}",
    "STR-003.tip.dupes": (
        "Tipp: Eindeutige Spaltennamen sind die Basis für automatische "
        "Auswertungen. Doppelte Namen führen zu Verwechslungen in Formeln und Tools."
    ),

    "STR-004.name": "Leere Trennzeilen/-spalten",
    "STR-004.msg": "{count} leere Zeilen innerhalb des Datenbereichs.",
    "STR-004.tip": (
        "Tipp: Visuelle Gliederung geht auch ohne leere Zeilen – "
        "Excel-Gruppierung, bedingte Formatierung oder separate Sheets "
        "halten die Daten zusammenhängend und auswertbar."
    ),

    "STR-005.name": "Identifier-Konsistenz",
    "STR-005.msg.inconsistent": (
        "Spalte {col}: Inkonsistente ID-Nummerierung "
        "für Prefix '{prefix}{sep}'."
    ),
    "STR-005.detail.inconsistent": (
        "Unterschiedliche Ziffernbreiten: {widths}. "
        "Beispiele: {examples}"
    ),
    "STR-005.tip.inconsistent": (
        "Tipp: Ein einheitliches Format wie '{prefix}{sep}001', "
        "'{prefix}{sep}002' (gleiche Stellenanzahl) erleichtert Sortierung "
        "und verhindert Verwechslungen bei der Zuordnung."
    ),
    "STR-005.msg.gaps": "Spalte {col}: Lücken in der ID-Sequenz '{prefix}{sep}...'.",
    "STR-005.detail.gaps": "Fehlende Nummern: {missing}",
    "STR-005.tip.gaps": (
        "Hinweis: Lücken können auf gelöschte Einträge hindeuten. "
        "Es lohnt sich, kurz zu prüfen, ob hier Daten fehlen."
    ),
    "STR-005.msg.separators": (
        "Spalte {col}: Prefix '{prefix}' verwendet "
        "unterschiedliche Trennzeichen: {seps}."
    ),
    "STR-005.tip.separators": "Einheitliches Trennzeichen verwenden, z.B. immer '{prefix}-'.",
    "STR-005.msg.dupes": "Spalte {col}: Doppelte Identifier gefunden.",
    "STR-005.detail.dupes": "Duplikate: {dupes}",
    "STR-005.tip.dupes": (
        "Tipp: Doppelte IDs können zu Zuordnungsfehlern führen. "
        "Am besten prüfen, welcher Eintrag der richtige ist."
    ),

    "STR-006.name": "Fehlender Primärschlüssel",
    "STR-006.msg": (
        "Keine Spalte mit eindeutigen Werten gefunden – "
        "diese Tabelle hat keinen verlässlichen Schlüssel "
        "zur Identifikation einzelner Einträge."
    ),
    "STR-006.detail": (
        "{rows} Datenzeilen in {cols} Spalten, "
        "aber kein Wert ist je Zeile einzigartig."
    ),
    "STR-006.tip": (
        "Das ist ein echtes Risiko: Ohne eindeutigen Schlüssel können "
        "Duplikate nicht erkannt, Daten nicht sicher verknüpft und "
        "Änderungen nicht nachverfolgt werden. Empfehlung: Laufende "
        "Nummer oder strukturierten Code (z.B. PRJ-001) als erste "
        "Spalte einführen."
    ),

    "STR-007.name": "Freitext-IDs (nicht sortierbar)",
    "STR-007.msg": (
        "Spalte {col} ('{header}') wird als ID verwendet, "
        "enthält aber Freitext statt strukturierter Codes."
    ),
    "STR-007.detail": (
        "Beispiele: {examples}. "
        "{ratio} der Werte sind langer Freitext."
    ),
    "STR-007.tip": (
        "IDs aus Freitext (z.B. 'Projekt Wien Mitte Sanierung Q4') "
        "sind nicht sortierbar, können leicht abweichen und machen "
        "Verknüpfungen unzuverlässig. Besser: Kurze, strukturierte "
        "Codes wie 'PRJ-0042' verwenden und den Freitext in eine "
        "eigene Beschreibungsspalte auslagern."
    ),

    # =====================================================================
    # Regeln – FRM (Formeln)
    # =====================================================================
    "FRM-001.name": "Absolute vs. Relative Bezüge",
    "FRM-001.msg": (
        "Hoher Anteil fixer Bezüge ($): {pct} von "
        "{total} Bezügen sind absolut fixiert."
    ),
    "FRM-001.detail": "Absolut: {absolute}, Relativ: {relative}, Gemischt: {mixed}",
    "FRM-001.tip": (
        "Tipp: Viele fixe Bezüge deuten darauf hin, dass Formeln schwer "
        "kopierbar und wartbar sind. Benannte Bereiche (z.B. 'MWSt' statt "
        "'$B$1') machen Formeln lesbar und flexibel."
    ),

    "FRM-002.name": "Volatile Funktionen (Performance)",
    "FRM-002.msg": (
        "Performance-Bremse entdeckt: Funktion {func}() wird "
        "in {count} Zellen verwendet und erzwingt bei jeder "
        "Änderung eine komplette Neuberechnung."
    ),
    "FRM-002.detail": "Zellen (Auswahl): {cells}",
    "FRM-002.tip": (
        "Tipp: {func}() ist eine 'volatile' Funktion – Excel "
        "berechnet sie bei JEDER Änderung neu, auch wenn sich nichts "
        "Relevantes geändert hat. Alternativen: INDEX/MATCH statt "
        "BEREICH.VERSCHIEBEN, feste Datumswerte statt HEUTE()."
    ),

    "FRM-003.name": "Lookup-Intensität",
    "FRM-003.msg": (
        "Dieses Sheet simuliert eine Datenbank: {count} "
        "Lookup-Funktionen verknüpfen Daten wie SQL-JOINs, nur langsamer."
    ),
    "FRM-003.detail": "Funktionen: {summary}",
    "FRM-003.tip": (
        "Tipp: So viele Lookups deuten darauf hin, dass die Daten eigentlich "
        "zusammengehören. Eine relationale Datenstruktur (Power Query, "
        "SharePoint-Liste oder Datenbank) würde diese Verknüpfungen "
        "automatisieren und die Datei deutlich schneller machen."
    ),

    "FRM-004.name": "Zirkelbezug-Hinweise",
    "FRM-004.msg": "Mögliche Zirkelbezüge in {count} Zellen entdeckt.",
    "FRM-004.detail": "Verdächtige Zellen: {cells}",
    "FRM-004.tip": (
        "Hinweis: Zirkelbezüge können zu unvorhersehbaren Berechnungen "
        "führen. Meistens entstehen sie durch Copy-Paste-Fehler. "
        "Bitte die markierten Zellen prüfen."
    ),

    "FRM-005.name": "Sheet- und Datei-Übergreifende Bezüge",
    "FRM-005.msg.external": (
        "Externe Dateibezüge entdeckt: {count} "
        "Formeln verweisen auf andere Excel-Dateien."
    ),
    "FRM-005.detail.external": "Referenzierte Dateien: {files}",
    "FRM-005.tip.external": (
        "Hinweis: Externe Bezüge machen die Datei abhängig von anderen "
        "Dateien – wenn diese verschoben oder umbenannt werden, brechen "
        "die Formeln. Besser: Daten zentral verwalten (z.B. Power Query "
        "oder SharePoint) und von dort abfragen."
    ),
    "FRM-005.msg.cross_sheet": (
        "Komplexe Vernetzung: Dieses Sheet referenziert "
        "{count} andere Sheets."
    ),
    "FRM-005.detail.cross_sheet": "Referenzierte Sheets: {sheets}",
    "FRM-005.tip.cross_sheet": (
        "Zur Info: Stark vernetzte Sheets können unübersichtlich werden. "
        "Ein zentrales Daten–Sheet als 'Single Source of Truth' kann "
        "die Struktur vereinfachen."
    ),

    "FRM-006.name": "Fehlerwerte in Formeln",
    "FRM-006.msg": "Fehlerwert in Formel entdeckt: {error}",
    "FRM-006.detail": "Fehlerwert: {error}",
    "FRM-006.tip": "Tipp: Fehlerwerte wie #DIV/0! oder #WERT! deuten auf fehlerhafte Berechnungen oder ungültige Eingaben hin. Bitte die betroffenen Zellen prüfen und korrigieren.",

    # =====================================================================
    # Regeln – VOL (Volumen)
    # =====================================================================
    "VOL-001.name": "Datenvolumen",
    "VOL-001.threshold.100k": "hat das Excel-Level durchgespielt",
    "VOL-001.threshold.50k": "ist für ein Datenbank-Upgrade reif",
    "VOL-001.threshold.10k": "nähert sich dem Bereich, wo eine Datenbank effizienter wäre",
    "VOL-001.msg": (
        "Gratulation – dieser Datensatz ist gewachsen! "
        "{rows} Zeilen: Dieser Bestand {label}."
    ),
    "VOL-001.detail": "Zeilen: {rows}, Spalten: {cols}",
    "VOL-001.tip": (
        "Tipp: Ab dieser Größe bieten Datenbanken (SharePoint-Liste, "
        "Dataverse, SQL) deutliche Vorteile: schnellere Suche, "
        "gleichzeitiger Zugriff und automatische Backups. "
        "Wir helfen gerne beim Umstieg!"
    ),

    "VOL-002.name": "Formel-Dichte",
    "VOL-002.msg.high": (
        "Hohe Rechenlast: {count} Formeln "
        "({pct} aller Zellen). Das kann die "
        "Datei merklich verlangsamen."
    ),
    "VOL-002.detail": "Formeln: {formulas}, Statische Werte: {statics}",
    "VOL-002.tip.high": (
        "Tipp: Formeln, die sich nicht mehr ändern, können als Werte "
        "eingefügt werden (Kopieren → 'Werte einfügen'). Noch besser: "
        "Berechnungen in Power Query oder eine Datenbank-View auslagern."
    ),
    "VOL-002.msg.medium": (
        "{count} Formeln in diesem Sheet – "
        "noch im grünen Bereich, aber im Auge behalten."
    ),
    "VOL-002.tip.medium": (
        "Tipp: Bei wachsenden Daten die Berechnungslogik regelmäßig "
        "hinterfragen. Power Query kann viele Formeln ersetzen."
    ),

    "VOL-003.name": "Anzahl Tabellenblätter",
    "VOL-003.msg": (
        "Diese Arbeitsmappe enthält {count} Tabellenblätter – "
        "das deutet auf ein komplexes Datenmodell hin, das über "
        "Excel hinausgewachsen ist."
    ),
    "VOL-003.detail": "Sheets: {sheets}",
    "VOL-003.tip": (
        "Tipp: Viele verknüpfte Sheets bilden im Grunde eine relationale "
        "Datenbank ab – nur ohne deren Vorteile. Eine Migration in eine "
        "echte Datenstruktur (Dataverse, SQL) mit Power BI als Frontend "
        "bietet die gleiche Übersicht, aber stabiler und schneller."
    ),

    "VOL-004.name": "Dateigröße",
    "VOL-004.msg.critical": (
        "Die Datei ist {size} MB groß – das führt zu langen "
        "Ladezeiten und erschwert die Zusammenarbeit."
    ),
    "VOL-004.tip.critical": (
        "Dringender Handlungsbedarf: Dateien dieser Größe gehören in ein "
        "professionelles System. Wir beraten gerne, welche Lösung "
        "(Datenbank, Data Warehouse, SharePoint) am besten passt."
    ),
    "VOL-004.msg.warning": (
        "Die Datei ist {size} MB groß – das kann die "
        "Ladezeit spürbar beeinflussen."
    ),
    "VOL-004.tip.warning": (
        "Tipp: Nicht mehr benötigte Daten archivieren, Bilder komprimieren "
        "oder Berechnungen als Werte einfügen, um die Größe zu reduzieren."
    ),

    # =====================================================================
    # Regeln – IMP (Implizites Wissen)
    # =====================================================================
    "IMP-001.name": "Undokumentierte Farbcodes",
    "IMP-001.msg.many": (
        "Implizites Wissen: {unique} verschiedene Hintergrundfarben "
        "in {total} Zellen – aber keine dokumentierte Legende gefunden."
    ),
    "IMP-001.detail": "Farben: {colors}",
    "IMP-001.tip.many": (
        "Tipp: Farbcodes enthalten oft wichtige Geschäftslogik "
        "(z.B. Ampel-Status, Prioritäten), die nur der Ersteller kennt. "
        "Eine Legende oder besser: eine eigene Status-Spalte macht dieses "
        "Wissen für alle zugänglich und automatisiert auswertbar."
    ),
    "IMP-001.msg.some": (
        "{unique} verschiedene Hintergrundfarben im Einsatz – "
        "falls diese eine Bedeutung haben, wäre eine Legende hilfreich."
    ),
    "IMP-001.tip.some": (
        "Tipp: Wenn Farben Status oder Kategorien darstellen, "
        "eine eigene Spalte dafür anlegen – das macht die Daten "
        "filtbar und auswertbar."
    ),
    "IMP-001.msg.font": (
        "{unique} verschiedene Schriftfarben im Einsatz – "
        "könnten implizite Kategorisierungen sein."
    ),
    "IMP-001.tip.font": (
        "Tipp: Schriftfarben als Bedeutungsträger sind unsichtbar "
        "beim Drucken und in Auswertungen. Besser: Eine eigene Spalte "
        "für die Klassifizierung."
    ),

    "IMP-002.name": "Versteckte Blätter",
    "IMP-002.msg.very_hidden": (
        "{count} Sheet(s) sind 'sehr versteckt' "
        "(nur über VBA sichtbar) – das birgt Risiken."
    ),
    "IMP-002.detail": "Sheets: {sheets}",
    "IMP-002.tip.very_hidden": (
        "Hinweis: 'Sehr versteckte' Sheets enthalten oft kritische "
        "Berechnungen oder Stammdaten, die niemand außer dem Ersteller kennt. "
        "Das ist ein Wissens-Risiko – bitte prüfen und dokumentieren."
    ),
    "IMP-002.msg.hidden": "{count} versteckte(s) Sheet(s) gefunden.",
    "IMP-002.tip.hidden": (
        "Tipp: Versteckte Blätter enthalten oft Hilfstabellen oder "
        "Stammdaten. Bitte prüfen, ob der Inhalt noch aktuell ist "
        "und ob er besser zentral verwaltet werden sollte."
    ),

    "IMP-003.name": "Versteckte Zeilen/Spalten",
    "IMP-003.msg": (
        "Versteckte Bereiche: {rows} Zeilen und "
        "{cols} Spalten sind ausgeblendet."
    ),
    "IMP-003.tip": (
        "Tipp: Ausgeblendete Bereiche enthalten oft veraltete oder "
        "sensible Daten. Besser: Nicht benötigte Daten löschen oder "
        "in ein Archiv-Sheet verschieben. Sensible Daten gehören in "
        "ein geschütztes System."
    ),

    "IMP-004.name": "Bedingte Formatierungen (Implizite Logik)",
    "IMP-004.msg": (
        "Geschäftslogik in Formatierung versteckt: "
        "{count} bedingte Formatierungsregeln – diese "
        "Regeln sind schwer wartbar und für andere unsichtbar."
    ),
    "IMP-004.tip": (
        "Tipp: Bedingte Formatierungen bilden oft wichtige "
        "Geschäftsregeln ab (z.B. 'rot wenn überfällig'). "
        "Besser: Eine eigene Status-Spalte mit Formeln anlegen – "
        "dann ist die Logik sichtbar, dokumentierbar und auswertbar."
    ),

    "IMP-005.name": "Hartcodierte Werte in Formeln",
    "IMP-005.msg": (
        "Implizites Wissen in Formeln: {count} "
        "hartcodierte Werte wiederholen sich in Formeln."
    ),
    "IMP-005.detail": "Werte: {values}",
    "IMP-005.tip": (
        "Tipp: Wiederkehrende Zahlen in Formeln (z.B. Steuersätze, "
        "Faktoren, Limits) sollten als 'Benannter Bereich' oder in "
        "einer Parametertabelle definiert werden. So wird beim Ändern "
        "nur EINE Stelle angepasst statt vieler."
    ),

    "IMP-006.name": "Hardcodierte Validierungslisten",
    "IMP-006.msg": (
        "{count} Dropdown-Listen mit eingetippten Werten – "
        "Änderungen müssen manuell in jeder Liste nachgezogen werden."
    ),
    "IMP-006.tip": (
        "Tipp: Dropdown-Werte in einer separaten Stammdaten-Tabelle "
        "pflegen und per Bereichsreferenz einbinden. Dann muss bei "
        "Änderungen nur EINE Stelle aktualisiert werden."
    ),

    "IMP-007.name": "Nicht-sprechende Sheet-Namen",
    "IMP-007.msg": "Nicht-sprechende Sheet-Namen: {names}",
    "IMP-007.tip": (
        "Tipp: Aussagekräftige Sheet-Namen wie 'Umsatz_2024' oder "
        "'Stammdaten_Kunden' helfen allen Beteiligten, sich schnell "
        "zurechtzufinden. 'Tabelle1' sagt niemandem etwas."
    ),

    "IMP-008.name": "Geschäftslogik in Kommentaren",
    "IMP-008.msg": (
        "Geschäftslogik in Zell-Kommentaren entdeckt: "
        "{count} Kommentare enthalten "
        "Hinweise wie 'Achtung', 'Nicht ändern' etc."
    ),
    "IMP-008.detail": "Zellen: {cells}",
    "IMP-008.tip": (
        "Hinweis: Wenn kritische Regeln nur in Kommentaren stehen, "
        "gehen sie leicht verloren oder werden übersehen. "
        "Besser: Geschäftsregeln als Datenvalidierung, bedingte "
        "Formatierung oder in einem Dokumentations-Sheet festhalten."
    ),

    "IMP-009.name": "Irreführende Zahlenformate",
    "IMP-009.msg": (
        "Spezielle Zahlenformate entdeckt, die den angezeigten "
        "Wert vom tatsächlichen Wert abweichen lassen können."
    ),
    "IMP-009.detail": "Formate: {formats}",
    "IMP-009.tip": (
        "Zur Info: Zahlenformate, die Text einblenden oder Werte "
        "optisch verändern, können zu Missverständnissen führen "
        "(z.B. '1' wird angezeigt, aber '1000' ist gespeichert). "
        "Bitte die kritischen Zellen prüfen."
    ),

    "IMP-010.name": "Blattschutz & gesperrte Bereiche",
    "IMP-010.msg": "{count} Sheet(s) sind geschützt: {sheets}",
    "IMP-010.tip": (
        "Zur Info: Blattschutz zeigt, dass bestimmte Bereiche kritisch "
        "sind. Gut so! Aber: Dokumentieren, WARUM und WAS geschützt ist. "
        "Wenn der Ersteller das Unternehmen verlässt, kann der Blattschutz "
        "zum Problem werden."
    ),

    # =====================================================================
    # Empfehlungen
    # =====================================================================
    "rec.db.title": "Diese Daten sind für eine Datenbank-Lösung prädestiniert",
    "rec.db.action": (
        "Empfohlener nächster Schritt: Einen Termin mit dem Digital Workplace Team "
        "vereinbaren. Wir analysieren gemeinsam, welche Datenbank-Lösung am besten "
        "passt (SharePoint-Liste, Dataverse, SQL Server) und begleiten die Migration. "
        "Die bestehenden Excel-Auswertungen können als Power BI Dashboard "
        "weitergeführt werden."
    ),
    "rec.db.reason.rows_high": "{rows} Zeilen übersteigen den sinnvollen Excel-Bereich",
    "rec.db.reason.rows_medium": "{rows} Zeilen nähern sich dem Bereich, wo Datenbanken effizienter sind",
    "rec.db.reason.lookups": "Exzessive Lookup-Funktionen simulieren Datenbank-JOINs",
    "rec.db.reason.size": "Dateigröße ({size} MB) verursacht Leistungsprobleme",
    "rec.db.reason.sheets": "{count} Blätter bilden ein relationales Datenmodell ab",
    "rec.db.reason.external": "Externe Dateibezüge zeigen verteilte, zusammengehörige Daten",

    "rec.powerbi.title": "Berechnungen und Visualisierungen in ein Reporting-Tool auslagern",
    "rec.powerbi.reason": (
        "{formulas} Formeln und/oder komplexe bedingte Formatierungen "
        "deuten auf aufwändige Auswertungslogik hin, die in Power BI oder "
        "einem vergleichbaren Tool stabiler und flexibler abgebildet werden kann."
    ),
    "rec.powerbi.action": (
        "Power BI kann direkt auf die Excel-Daten zugreifen und die "
        "Auswertungen dynamisch darstellen. Vorteil: Die Berechnungslogik "
        "wird dokumentiert, ist für andere nachvollziehbar und kann ohne "
        "Excel-Expertenwissen angepasst werden."
    ),

    "rec.sharepoint.title": "SharePoint-Liste als niederschwellige Alternative",
    "rec.sharepoint.reason": (
        "Die Datenstruktur ({rows} Zeilen, {sheets} Blätter) "
        "eignet sich gut für eine SharePoint-Liste – einfach zu erstellen, "
        "mit Versionshistorie und Berechtigungssteuerung."
    ),
    "rec.sharepoint.action": (
        "Eine SharePoint-Liste kann in wenigen Minuten aus der Excel-Tabelle "
        "erstellt werden (Excel → 'Als Tabelle exportieren'). "
        "Vorteile: Gleichzeitiger Zugriff, Änderungshistorie, automatische "
        "Benachrichtigungen und Integration mit Power Automate."
    ),

    "rec.normalization.title": "Datenstruktur bereinigen – Grundlage für alle weiteren Schritte",
    "rec.normalization.reason": (
        "{count} Strukturprobleme erkannt (verbundene Zellen, "
        "gemischte Datentypen, fehlende Kopfzeilen, Leerzeilen als Trenner). "
        "Diese verhindern jede Form der automatisierten Verarbeitung."
    ),
    "rec.normalization.action": (
        "Erster Schritt: Die Daten in ein tabellarisches Format bringen "
        "(eine Kopfzeile, keine verbundenen Zellen, keine Leerzeilen). "
        "Das Digital Workplace Team kann bei der Bereinigung unterstützen – "
        "oft reicht ein gemeinsamer Workshop-Termin."
    ),

    "rec.split.title": "Arbeitsmappe aufteilen",
    "rec.split.reason": (
        "{count} Blätter in einer Datei sind schwer zu überblicken. "
        "Oft enthalten verschiedene Blätter verschiedene Themen, die "
        "besser getrennt verwaltet werden."
    ),
    "rec.split.action": (
        "Prüfen, welche Sheets thematisch zusammengehören und welche "
        "eigenständige Datensätze sind. Eigenständige Bereiche in "
        "separate Dateien oder SharePoint-Listen auslagern."
    ),

    "rec.cleanup.title": "Implizites Wissen dokumentieren und explizit machen",
    "rec.cleanup.reason": (
        "{count} Fälle von undokumentiertem Wissen gefunden "
        "(Farbcodes ohne Legende, versteckte Bereiche, Logik in Kommentaren). "
        "Dieses Wissen geht verloren, wenn der Ersteller nicht verfügbar ist."
    ),
    "rec.cleanup.action": (
        "Dringend empfohlen: Einen 30-minütigen Termin mit dem Ersteller "
        "vereinbaren, um die impliziten Regeln zu dokumentieren. "
        "Farbcodes → Status-Spalte, Kommentar-Regeln → Datenvalidierung, "
        "Versteckte Blätter → Dokumentation oder Löschung."
    ),

    "rec.ok.title": "Gute Arbeit – die Excel-Nutzung ist hier vertretbar",
    "rec.ok.reason": (
        "Bewertung {score}/100. "
        "Die Datei ist gut strukturiert und im sinnvollen Rahmen für Excel."
    ),
    "rec.ok.action": (
        "Keine dringenden Maßnahmen nötig. "
        "Für die Zukunft: Regelmäßig prüfen, ob die Datenmenge wächst."
    ),
    "rec.minor.title": "Kleinere Optimierungen empfohlen",
    "rec.minor.reason": "Bewertung {score}/100.",
    "rec.minor.action": (
        "Die einzelnen Hinweise im Bericht beachten – "
        "die meisten lassen sich mit wenig Aufwand umsetzen."
    ),

    # =====================================================================
    # Report – Labels & Stufen
    # =====================================================================
    "report.grade.mega": "Mega-Problem",
    "report.grade.real": "Echtes Problem",
    "report.grade.cosmetic": "Schönheitsfehler",

    "report.score.good": "Gut",
    "report.score.needs_improvement": "Verbesserungsbedarf",
    "report.score.action_needed": "Handlungsbedarf",
    "report.score.urgent": "Dringender Handlungsbedarf",

    "report.dim.volume": "Daten-\nvolumen",
    "report.dim.formulas": "Formel-\nKomplexität",
    "report.dim.networking": "Vernetzung",
    "report.dim.implicit": "Implizites\nWissen",
    "report.dim.structure": "Strukturelle\nProbleme",
    "report.dim.filesize": "Datei-\ngröße",

    # Anti-Pattern Karten (Name, Icon, Beschreibung, Schweregrad)
    "ap.STR-005.name": "ID-Wildwuchs",
    "ap.STR-005.desc": "Inkonsistente Identifier (Formate, Lücken, Duplikate).",
    "ap.STR-006.name": "Phantom-Schlüssel",
    "ap.STR-006.desc": "Datentabelle ohne eindeutige, sortierbare Schlüsselspalte.",
    "ap.STR-007.name": "Freitext-IDs",
    "ap.STR-007.desc": "IDs aus Freitext statt sortierbarer Codes – nicht maschinell verarbeitbar.",
    "ap.FRM-004.name": "Zirkelbezug",
    "ap.FRM-004.desc": "Formeln referenzieren sich selbst – unberechenbare Ergebnisse.",
    "ap.VOL-001.name": "Excel-Sprenger",
    "ap.VOL-001.desc": "Datenvolumen hat die sinnvolle Excel-Kapazität überschritten.",
    "ap.VOL-002.name": "Formel-Overload",
    "ap.VOL-002.desc": "Zu viele Formeln – Instabilität und Langsamkeit.",
    "ap.STR-001.name": "Verbundene-Zellen-Chaos",
    "ap.STR-001.desc": "Verbundene Zellen zerstören maschinelle Lesbarkeit.",
    "ap.STR-002.name": "Typ-Salat",
    "ap.STR-002.desc": "Gemischte Datentypen in einer Spalte – 1. Normalform verletzt.",
    "ap.STR-003.name": "Kopflos-Tabelle",
    "ap.STR-003.desc": "Fehlende oder unklare Kopfzeilen.",
    "ap.FRM-002.name": "Volatile-Falle",
    "ap.FRM-002.desc": "Volatile Funktionen erzwingen permanente Neuberechnung.",
    "ap.FRM-003.name": "SVERWEIS-Datenbank",
    "ap.FRM-003.desc": "Lookup-Ketten simulieren SQL-JOINs – viel langsamer.",
    "ap.FRM-005.name": "Datei-Spinnennetz",
    "ap.FRM-005.desc": "Externe Verknüpfungen – fragile Abhängigkeiten.",
    "ap.IMP-001.name": "Geheim-Ampel",
    "ap.IMP-001.desc": "Farbcodes tragen undokumentierte Bedeutung.",
    "ap.IMP-002.name": "Versteckspiel",
    "ap.IMP-002.desc": "Versteckte Blätter mit undokumentiertem Wissen.",
    "ap.IMP-004.name": "Format-Logik",
    "ap.IMP-004.desc": "Geschäftslogik in bedingten Formatierungen versteckt.",
    "ap.IMP-005.name": "Zauberzahlen",
    "ap.IMP-005.desc": "Hartcodierte Zahlen in Formeln ohne Erklärung.",
    "ap.VOL-003.name": "Blatt-Hydra",
    "ap.VOL-003.desc": "Zu viele Blätter bilden eine relationale DB ab – ohne deren Vorteile.",
    "ap.VOL-004.name": "Schwergewicht",
    "ap.VOL-004.desc": "Überdimensionierte Dateigröße.",
    "ap.STR-004.name": "Leerzeilen-Layout",
    "ap.STR-004.desc": "Leere Zeilen als visuelle Trenner fragmentieren den Datenbereich.",
    "ap.FRM-001.name": "Dollar-Fixierung",
    "ap.FRM-001.desc": "Zu viele starre $-Bezüge – schwer wartbar.",
    "ap.IMP-003.name": "Unsichtbare Daten",
    "ap.IMP-003.desc": "Ausgeblendete Zeilen/Spalten.",
    "ap.IMP-006.name": "Eingetippte Dropdowns",
    "ap.IMP-006.desc": "Validierungslisten mit hardcodierten Werten.",
    "ap.IMP-007.name": "Namenlose Blätter",
    "ap.IMP-007.desc": "Generische Blatt-Namen ohne Aussagekraft.",
    "ap.IMP-008.name": "Kommentar-Wissen",
    "ap.IMP-008.desc": "Geschäftsregeln leben nur in Zell-Kommentaren.",
    "ap.IMP-009.name": "Zahlen-Maskerade",
    "ap.IMP-009.desc": "Irreführende Zahlenformate.",
    "ap.IMP-010.name": "Blattschutz-Rätsel",
    "ap.IMP-010.desc": "Geschützte Bereiche ohne Dokumentation.",

    # AI-Readiness Blocker
    "report.ai.merged": "{count} verbundene Zellbereiche",
    "report.ai.merged_desc": "KI kann tabellarische Struktur nicht erkennen",
    "report.ai.implicit": "{count}× implizites Wissen",
    "report.ai.implicit_desc": "Farbcodes, versteckte Blätter, Zauberzahlen – für KI unsichtbar",
    "report.ai.structure": "{count}× Strukturproblem",
    "report.ai.structure_desc": "Gemischte Typen, fehlende IDs – KI kann Datensätze nicht zuordnen",
    "report.ai.formulas": "{count}× Formel-Verkettung",
    "report.ai.formulas_desc": "SVERWEIS-Ketten & Zirkelbezüge blockieren automatische Auswertung",
    "report.ai.weak": "Strukturelle Schwächen",
    "report.ai.weak_desc": "Daten sind nicht in maschinenlesbarer Form",

    # =====================================================================
    # LLM-Analyse
    # =====================================================================
    "llm.no_key": "Kein API-Key angegeben.",
    "llm.no_endpoint": "Kein Endpoint angegeben. Bitte die Azure AI Endpoint-URL konfigurieren.",
    "llm.key_invalid": "API-Key ungültig (401 Unauthorized).",
    "llm.access_denied": "Zugriff verweigert (403 Forbidden).",
    "llm.timeout": "Timeout – Server nicht erreichbar.",
    "llm.connection_error": "Verbindungsfehler – Endpoint nicht erreichbar.",
    "llm.error": "Fehler: {error}",
    "llm.key_works": "API-Key funktioniert! ✅",
    "llm.parse_error": "Die KI-Analyse konnte nicht vollständig geparst werden.",

    # =====================================================================
    # Webapp – Server-seitige Meldungen
    # =====================================================================
    "web.no_file": "Keine Datei ausgewählt.",
    "web.unsupported": "Nur .xlsx und .xlsm Dateien werden unterstützt.",
    "web.no_url": "Bitte eine URL eingeben.",
    "web.invalid_url": "Bitte eine gültige https:// URL eingeben.",
    "web.session_not_found": "Session nicht gefunden.",
    "web.report_not_found": "Report nicht gefunden.",
    "web.report_expired": "Der Report ist abgelaufen oder wurde bereits angezeigt. Bitte eine neue Analyse starten.",
    "web.error_title": "Fehler",
    "web.no_file_title": "Keine Datei",
    "web.no_file_msg": "Bitte eine Excel-Datei auswählen.",
    "web.wrong_format_title": "Falsches Format",
    "web.analysis_failed_title": "Analyse fehlgeschlagen",
    "web.analysis_failed_msg": "Die Datei konnte nicht analysiert werden: {error}",
    "web.download_failed_title": "Download fehlgeschlagen",
    "web.invalid_url_title": "Ungültige URL",
    "web.no_url_title": "Keine URL",
    "web.no_url_msg": "Bitte eine URL eingeben.",
    "web.connection_lost": "Verbindung zum Server verloren.",
    "web.report_not_found_llm": "Report nicht gefunden. Bitte zuerst eine klassische Analyse durchführen.",
    "web.no_key_settings": "Kein API-Key angegeben. Bitte unter ⚙️ KI-Einstellungen konfigurieren.",
    "web.no_endpoint_settings": "Kein Endpoint konfiguriert. Bitte unter ⚙️ KI-Einstellungen die Azure AI Endpoint-URL eingeben.",
    "web.llm_failed": "LLM-Analyse fehlgeschlagen: {error}",
    "web.login_required": (
        "Der Link erfordert eine Anmeldung. Bitte den Link über "
        "'Freigeben → Link kopieren' mit 'Jeder mit dem Link' "
        "erstellen, oder die Datei direkt hochladen."
    ),
    "web.file_too_large": "Die Datei ist größer als 100 MB.",
    "web.not_excel": (
        "Die heruntergeladene Datei ist keine gültige Excel-Datei. "
        "Möglicherweise verweist der Link auf eine Login-Seite. "
        "Bitte die Datei stattdessen direkt hochladen."
    ),
    "web.timeout": "Timeout: Der Server hat nicht rechtzeitig geantwortet.",
    "web.connect_error": "Verbindungsfehler: Der Server ist nicht erreichbar.",
    "web.access_denied": (
        "Zugriff verweigert. Bitte einen Freigabe-Link mit "
        "'Jeder mit dem Link' erstellen, oder die Datei direkt hochladen."
    ),
    "web.http_error": "HTTP-Fehler {status} beim Herunterladen.",
    "web.download_error": "Fehler beim Herunterladen: {error}",
    "web.back_to_upload": "← Zurück zum Upload",
}

"""CLI für den Excel-Reifecheck."""

from __future__ import annotations

import argparse
import os
import sys
import time

from excel_checker.i18n import _, set_language
from excel_checker.engine import analyze
from excel_checker.models import Severity
from excel_checker.report import generate_html


SEVERITY_SYMBOLS = {
    Severity.CRITICAL: "🔴",
    Severity.ERROR: "🟠",
    Severity.WARNING: "🟡",
    Severity.INFO: "🔵",
}


def main():
    parser = argparse.ArgumentParser(
        prog="excel-reifecheck",
        description=_("cli.description"),
    )
    parser.add_argument(
        "files", nargs="+",
        help=_("cli.files_help"),
    )
    parser.add_argument(
        "--html", "-o",
        help=_("cli.html_help"),
    )
    parser.add_argument(
        "--open", action="store_true",
        help=_("cli.open_help"),
    )
    parser.add_argument(
        "--quiet", "-q", action="store_true",
        help=_("cli.quiet_help"),
    )
    parser.add_argument(
        "--lang", choices=["de", "en"], default="de",
        help=_("cli.lang_help"),
    )
    args = parser.parse_args()

    set_language(args.lang)

    for file_path in args.files:
        if not os.path.exists(file_path):
            print(_("cli.file_not_found", path=file_path), file=sys.stderr)
            sys.exit(1)
        if not file_path.lower().endswith(('.xlsx', '.xlsm')):
            print(_("cli.unsupported", path=file_path),
                  file=sys.stderr)
            sys.exit(1)

    for file_path in args.files:
        filename = os.path.basename(file_path)

        print(f"\n{'='*60}")
        print(_("cli.health_check", filename=filename))
        print(f"{'='*60}")

        start = time.time()
        report = analyze(file_path)
        elapsed = time.time() - start

        if args.quiet:
            print(f"{report.health_score}")
            continue

        # Score anzeigen
        score = report.health_score
        if score >= 80:
            score_icon = "🟢"
        elif score >= 60:
            score_icon = "🟡"
        elif score >= 40:
            score_icon = "🟠"
        else:
            score_icon = "🔴"

        print(f"\n{_('cli.score', icon=score_icon, score=score)}")
        print(_("cli.file_size", size=f"{report.file_size_mb:.1f}"))
        print(_("cli.sheets", count=report.sheet_count))
        print(_("cli.analysis_time", elapsed=f"{elapsed:.1f}"))

        # Findings
        if report.findings:
            print(f"\n{_('cli.findings_header', count=len(report.findings))}")
            for f in sorted(report.findings,
                            key=lambda x: list(Severity).index(x.severity)):
                sym = SEVERITY_SYMBOLS.get(f.severity, "⚪")
                sheet_info = f" [{f.sheet}]" if f.sheet else ""
                print(f"  {sym} {f.message}{sheet_info}")
                if f.suggestion:
                    print(f"     💡 {f.suggestion}")

        # Empfehlungen
        if report.recommendations:
            print(f"\n{_('cli.rec_header')}")
            for rec in report.recommendations:
                print(f"\n{_('cli.rec_priority', priority=rec.priority, title=rec.title)}")
                print(_("cli.rec_reason", reason=rec.reason))
                print(_("cli.rec_action", action=rec.action))

        # HTML-Report
        if args.html:
            html_path = args.html
        else:
            html_path = os.path.splitext(file_path)[0] + "_health_report.html"

        html_content = generate_html(report)
        with open(html_path, "w", encoding="utf-8") as fh:
            fh.write(html_content)
        print(f"\n{_('cli.html_report', path=html_path)}")

        if args.open:
            import webbrowser
            webbrowser.open(f"file://{os.path.abspath(html_path)}")


if __name__ == "__main__":
    main()

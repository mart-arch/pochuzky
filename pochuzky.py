"""Lokální přehled času, po který jednotlivé body držela daná strana."""

import argparse
import csv
import sys
import tempfile
import unicodedata
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timezone
from html import escape
from io import BytesIO
from itertools import groupby
from pathlib import Path
from zipfile import BadZipFile
from zoneinfo import ZoneInfo

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


SHEET = "Pochůzky BODY"
POINTS = (
    "Vyšehrad, u štábu NH", "Palmovka", "Pankrác", "Skalka",
    "CENTRAL PARK  - Hůrka", "Strahov Silicon hill", "Park u Kolbenovy",
    "Bod v Háji", "Džbán", "Riegrovy sady Hipstertown", "Nusle/Grebovka",
    "Braník", "Stromovka", "Park Ladronka", "Štáb DH", "Zličín", "Ládví",
    "Kavčí hory", "Parukářka", "Štvanice", "Radlická",
)
PRAGUE = ZoneInfo("Europe/Prague")
UTC = timezone.utc
LABELS = {
    "NH": "NH", "DH": "DH", "Civil": "Civil", "???": "???",
    "": "Prázdná", None: "Nezjištěno",
}
COLORS = ("#356ace", "#ad3e53", "#2d806a", "#bf7b19", "#c5cbd6", "#666277")


def key(value):
    return unicodedata.normalize("NFC", " ".join(str(value).split())).casefold()


POINT_NAMES = {key(name): name for name in POINTS}
SIDES = {key(value): value for value in LABELS if value is not None}


@dataclass(frozen=True)
class Revision:
    time: datetime
    source: str


def parse_time(value):
    """Interpret bare dates and times in Prague. Compare elapsed time in UTC."""
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError:
        raise ValueError(f"Neplatný čas: {value!r}. Použijte datum nebo čas ve formátu ISO.") from None
    if parsed.tzinfo is None:
        candidates = {
            parsed.replace(tzinfo=PRAGUE, fold=fold).astimezone(UTC)
            for fold in (0, 1)
        }
        candidates = {
            time for time in candidates
            if time.astimezone(PRAGUE).replace(tzinfo=None) == parsed
        }
        if len(candidates) != 1:
            raise ValueError(f"Čas {value!r} v Praze neexistuje nebo je dvojznačný. Uveďte posun vůči UTC.")
        return candidates.pop()
    return parsed.astimezone(UTC)


def read_snapshot(content):
    """Read only the two required fields. Missing rows and blank cells differ."""
    book = load_workbook(BytesIO(content), read_only=True, data_only=False)
    try:
        if SHEET not in book.sheetnames:
            raise ValueError(f"Chybí list {SHEET!r}.")
        rows = book[SHEET].iter_rows()
        for row in rows:
            values = [key(cell.value) for cell in row]
            if "název bodu" in values and "strana" in values:
                name_col, side_col = values.index("název bodu"), values.index("strana")
                break
        else:
            raise ValueError("Chybí záhlaví NÁZEV BODU a STRANA na stejném řádku.")

        states = {}
        for row in rows:
            if name_col >= len(row):
                continue
            point = POINT_NAMES.get(key(row[name_col].value))
            if point is None:
                continue
            if point in states:
                raise ValueError(f"Bod {point!r} je v listu vícekrát.")
            cell = row[side_col] if side_col < len(row) else None
            value = cell.value if cell is not None else None
            if value is None:
                states[point] = ""
            elif key(value) in SIDES:
                states[point] = SIDES[key(value)]
            else:
                raise ValueError(f"Bod {point!r} má neplatnou hodnotu ve sloupci STRANA ({cell.coordinate}).")
        return states
    finally:
        book.close()


def read_manifest(path):
    with path.open(encoding="utf-8-sig", newline="") as stream:
        rows = csv.DictReader(stream, delimiter=";")
        if not rows.fieldnames or not {"cas", "soubor"}.issubset(rows.fieldnames):
            raise ValueError("V seznamu snímků chybí záhlaví cas;soubor.")
        for number, row in enumerate(rows, 2):
            if not row["cas"] or not row["soubor"]:
                raise ValueError(f"V seznamu snímků chybí čas nebo soubor na řádku {number}.")
            yield Revision(parse_time(row["cas"].strip()), str(path.parent / row["soubor"].strip()))


def intervals(revisions, read, start, end):
    """Use the last snapshot at/before start and the snapshots inside [start, end)."""
    earlier = max((item.time for item in revisions if item.time <= start), default=None)
    selected = sorted(
        (item for item in revisions if item.time == earlier or start < item.time < end),
        key=lambda item: item.time,
    )
    timeline = {point: [] for point in POINTS}
    states = {}
    cursor = start

    def add_until(boundary):
        if boundary <= cursor:
            return
        for point, spans in timeline.items():
            side = states.get(point)
            if spans and spans[-1][2] == side:
                spans[-1] = (spans[-1][0], boundary, side)
            else:
                spans.append((cursor, boundary, side))

    for time, group in groupby(selected, key=lambda item: item.time):
        add_until(time)
        alternatives = []
        for rev in group:
            try:
                alternatives.append(read_snapshot(read(rev.source)))
            except (BadZipFile, InvalidFileException, OSError, ValueError, RuntimeError) as error:
                raise ValueError(f"Snímek {rev.source} v {local_time(time)}: {error}") from error
        states = {}
        for point in POINTS:
            observed = {snapshot.get(point) for snapshot in alternatives}
            states[point] = observed.pop() if len(observed) == 1 else None
        cursor = max(time, start)
    add_until(end)
    return timeline, len(selected), earlier


def local_time(time):
    return time.astimezone(PRAGUE).isoformat(sep=" ", timespec="auto")


def write_report(directory, timeline, start, end, count, baseline):
    colors = dict(zip(LABELS, COLORS, strict=True))
    duration = (end - start).total_seconds()
    totals = []
    bars = []
    with (directory / "intervaly.csv").open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream, delimiter=";")
        writer.writerow(("Název bodu", "Od", "Do", "Strana", "Hodiny"))
        for point, spans in timeline.items():
            seconds_by_side = Counter()
            segments = []
            for left, right, side in spans:
                seconds = (right - left).total_seconds()
                seconds_by_side[side] += seconds
                writer.writerow((point, local_time(left), local_time(right), LABELS[side], hours(seconds)))
                tip = escape(f"{LABELS[side]} · {local_time(left)} až {local_time(right)} · {hours(seconds)} h")
                width = 100 * seconds / duration
                segments.append(f'<span style="width:{width:.12g}%;background:{colors[side]}" title="{tip}"></span>')
            totals.append((point, *(hours(seconds_by_side[side]) for side in LABELS)))
            bars.append(f'<div class="point"><span>{escape(point)}</span><div class="bar">{"".join(segments)}</div></div>')

    headers = ("Název bodu", *LABELS.values())
    with (directory / "hodiny.csv").open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream, delimiter=";")
        writer.writerow(headers)
        writer.writerows(totals)

    legend = "".join(f'<span><i style="background:{colors[side]}"></i>{escape(label)}</span>'
                     for side, label in LABELS.items())
    head = "".join(f'<th scope="col">{escape(value)}</th>' for value in headers)
    body = "".join(
        f'<tr><th scope="row">{escape(point)}</th>'
        + "".join(f"<td>{value}</td>" for value in values) + "</tr>"
        for point, *values in totals
    )
    initial = local_time(baseline) if baseline is not None else "není k dispozici"
    (directory / "prehled.html").write_text(f"""<!doctype html>
<html lang="cs"><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>Pochůzky · historie stran</title>
<style>
body {{ font: 16px/1.5 system-ui, sans-serif; color: #152238; background: #f5f7fb;
       max-width: 1050px; margin: 0 auto; padding: 1rem 1.5rem 3rem }}
h1, h2 {{ line-height: 1.15 }}
section {{ background: white; padding: 1rem 1.2rem; margin-top: 1rem; border-radius: .75rem }}
.legend {{ display: flex; flex-wrap: wrap; gap: .4rem 1rem; margin: 1rem 0 }}
.legend i {{ display: inline-block; width: .8rem; height: .8rem; margin-right: .3rem }}
.point {{ display: grid; grid-template-columns: minmax(8rem, 30%) 1fr; gap: .8rem;
          align-items: center; padding: .28rem 0 }}
.bar {{ display: flex; height: 1.2rem; overflow: hidden; border-radius: .18rem }}
.bar span {{ flex-shrink: 0 }}
.scroll {{ overflow-x: auto }}
table {{ border-collapse: collapse; width: 100%; font-variant-numeric: tabular-nums }}
th, td {{ padding: .4rem .55rem; border-bottom: 1px solid #e3e8ef; text-align: right }}
th:first-child {{ text-align: left }}
a {{ color: #1d56ac }}
</style>
<h1>Historie stran na bodech</h1>
<p>{local_time(start)} až {local_time(end)} (Europe/Prague) · {hours(duration)} h</p>
<p>Použité snímky: {count} · Výchozí snímek: {initial}</p>
<section><h2>Časová osa</h2><div class="legend">{legend}</div>{"".join(bars)}</section>
<section><h2>Hodiny podle bodu a strany</h2>
<div class="scroll"><table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>
<p><a href="hodiny.csv">Stáhnout součty CSV</a> · <a href="intervaly.csv">Stáhnout intervaly CSV</a></p>
</section>
<p>Údaje vycházejí ze snímků. Změny mezi snímky mohou v přehledu chybět.
Nezjištěno znamená, že pro daný interval není známá jednoznačná hodnota.</p>
</html>
""", encoding="utf-8")


def hours(seconds):
    return f"{seconds / 3600:.6f}".rstrip("0").rstrip(".").replace(".", ",")


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--od", required=True, help="Začátek období včetně, datum nebo čas ISO.")
    parser.add_argument("--do", required=True, help="Konec období bez tohoto okamžiku.")
    source = parser.add_mutually_exclusive_group()
    source.add_argument("--oauth", type=Path, help="JSON s OAuth klientem. Výchozí: credentials.json.")
    source.add_argument("--snimky", type=Path, help="Místní seznam XLSX snímků v CSV (cas;soubor).")
    args = parser.parse_args(argv)
    try:
        start, end = parse_time(args.od), parse_time(args.do)
        if start >= end:
            raise ValueError("Začátek období musí předcházet konci.")
        if args.snimky:
            result = intervals(list(read_manifest(args.snimky)), lambda file: Path(file).read_bytes(), start, end)
        else:
            from drive import connect

            # Do not extrapolate into the future, including time spent authorizing.
            if end > datetime.now(UTC):
                raise ValueError("Konec období pro historii z Google Drive nesmí být v budoucnosti.")
            with connect(args.oauth or Path("credentials.json")) as drive:
                revisions = [
                    Revision(parse_time(time), revision_id)
                    for revision_id, time in drive.revisions().items()
                ]
                result = intervals(revisions, drive.download, start, end)

        timeline, count, baseline = result
        root = Path("vystup")
        root.mkdir(exist_ok=True)
        with tempfile.TemporaryDirectory(prefix=".vypocet-", dir=root) as temporary:
            staged = Path(temporary) / "report"
            staged.mkdir()
            write_report(staged, timeline, start, end, count, baseline)
            target = root / Path(temporary).name.removeprefix(".vypocet-")
            staged.rename(target)
        print(f"Přehled: {(target / 'prehled.html').resolve()}")
        if baseline is None:
            print("Pro začátek období chybí výchozí snímek.", file=sys.stderr)
        return 0
    except ModuleNotFoundError as error:
        print(f"Chybí závislost {error.name}. Spusťte python -m pip install -r requirements.txt.", file=sys.stderr)
        return 1
    except (ValueError, RuntimeError, OSError) as error:
        print(f"Chyba: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())

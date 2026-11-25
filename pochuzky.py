"""Lokální přehled času, po který jednotlivé body držela daná strana."""

import argparse
import csv
import json
import sys
import tempfile
import unicodedata
from collections import Counter
from contextlib import ExitStack, contextmanager
from dataclasses import dataclass
from datetime import datetime, timezone
from html import escape
from io import BytesIO
from itertools import groupby
from pathlib import Path
from random import uniform
from time import sleep
from zipfile import BadZipFile
from zoneinfo import ZoneInfo

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


SHEET = "Pochůzky BODY"
SNAPSHOT_COLUMNS = ("NÁZEV BODU", "STRANA")
OUTPUT = Path("vystup")
MANIFEST = "snimky.csv"
PERIOD = "obdobi.json"
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
        for header_row, row in enumerate(rows, 1):
            values = [key(cell.value) for cell in row]
            if "název bodu" in values and "strana" in values:
                name_col, side_col = values.index("název bodu"), values.index("strana")
                break
        else:
            raise ValueError("Chybí záhlaví NÁZEV BODU a STRANA na stejném řádku.")

        return read_states(
            (row[name_col].value, row[side_col].value if side_col < len(row) else None, f"řádek {number}")
            for number, row in enumerate(rows, header_row + 1)
            if name_col < len(row)
        )
    finally:
        book.close()


def read_states(rows):
    states = {}
    for name, value, location in rows:
        point = POINT_NAMES.get(key(name))
        if point is None:
            continue
        if point in states:
            raise ValueError(f"Bod {point!r} je ve snímku vícekrát ({location}).")
        side = "" if value is None else key(value)
        if side not in SIDES:
            raise ValueError(f"Bod {point!r} má neplatnou hodnotu ve sloupci STRANA ({location}).")
        states[point] = SIDES[side]
    return states


def read_local_snapshot(source):
    path = Path(source)
    if path.suffix.lower() == ".xlsx":
        return read_snapshot(path.read_bytes())
    if path.suffix.lower() != ".csv":
        raise ValueError("Snímek musí být ve formátu CSV nebo XLSX.")
    with path.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.reader(stream, delimiter=";", strict=True)
        if tuple(next(reader, ())) != SNAPSHOT_COLUMNS:
            raise ValueError("Chybí záhlaví NÁZEV BODU;STRANA.")

        def entries():
            for number, row in enumerate(reader, 2):
                if len(row) != 2:
                    raise ValueError(f"Snímek má na řádku {number} jiný počet sloupců než dva.")
                yield (*row, f"řádek {number}")
        return read_states(entries())


def read_manifest(path):
    with path.open(encoding="utf-8-sig", newline="") as stream:
        rows = csv.DictReader(stream, delimiter=";")
        if not rows.fieldnames or not {"cas", "soubor"}.issubset(rows.fieldnames):
            raise ValueError("V seznamu snímků chybí záhlaví cas;soubor.")
        for number, row in enumerate(rows, 2):
            if not row["cas"] or not row["soubor"]:
                raise ValueError(f"V seznamu snímků chybí čas nebo soubor na řádku {number}.")
            yield Revision(parse_time(row["cas"].strip()), str(path.parent / row["soubor"].strip()))


def select_revisions(revisions, start, end):
    """Use the last snapshot at/before start and the snapshots inside [start, end)."""
    earlier = max((item.time for item in revisions if item.time <= start), default=None)
    selected = sorted(
        (item for item in revisions if item.time == earlier or start < item.time < end),
        key=lambda item: item.time,
    )
    return selected, earlier


def read_revision(rev, read):
    try:
        return read(rev.source)
    except (BadZipFile, InvalidFileException, OSError, ValueError, RuntimeError, csv.Error) as error:
        raise ValueError(f"Snímek {rev.source} v {local_time(rev.time)}: {error}") from error


def intervals(revisions, read, start, end):
    selected, earlier = select_revisions(revisions, start, end)
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
        alternatives = [read_revision(rev, read) for rev in group]
        states = {}
        for point in POINTS:
            observed = {snapshot.get(point) for snapshot in alternatives}
            states[point] = observed.pop() if len(observed) == 1 else None
        cursor = max(time, start)
    add_until(end)
    return timeline, len(selected), earlier


def local_time(time):
    return time.astimezone(PRAGUE).isoformat(sep=" ", timespec="auto")


@contextmanager
def stage_output(root, prefix=""):
    root.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".pochuzky-", dir=root) as temporary:
        directory = Path(temporary)
        staged = directory / "obsah"
        staged.mkdir(mode=0o700)
        target = root / (prefix + directory.name.removeprefix(".pochuzky-"))
        yield staged, target
        staged.rename(target)


@contextmanager
def atomic_output(path):
    with tempfile.TemporaryDirectory(prefix=".pochuzky-", dir=path.parent) as temporary:
        staged = Path(temporary) / path.name
        yield staged
        staged.replace(path)


def start_download(revisions, start, end, snapshot_format):
    selected, _ = select_revisions(revisions, start, end)
    full_workbooks = snapshot_format == "xlsx"
    root = Path(tempfile.gettempdir()) if full_workbooks else OUTPUT
    prefix = "pochuzky-snimky-" if full_workbooks else "snimky-"
    root.mkdir(exist_ok=True)
    target = Path(tempfile.mkdtemp(prefix=prefix, dir=root))
    plan = {
        "od": start.isoformat(), "do": end.isoformat(), "format": snapshot_format,
        "revize": [(rev.time.isoformat(), rev.source) for rev in selected],
    }
    with atomic_output(target / PERIOD) as path:
        path.write_text(json.dumps(plan) + "\n", encoding="utf-8")
    return target, selected


def read_download_plan(target, start, end, snapshot_format):
    path = target / PERIOD
    try:
        plan = json.loads(path.read_text(encoding="utf-8"))
        saved_start, saved_end = parse_time(plan["od"]), parse_time(plan["do"])
        saved_format = plan["format"]
        selected = [Revision(parse_time(time), source) for time, source in plan["revize"]]
        if not all(isinstance(rev.source, str) and rev.source for rev in selected):
            raise ValueError("Chybí identifikátor revize.")
    except (KeyError, OSError, TypeError, ValueError) as error:
        raise ValueError(f"Nelze načíst údaje pro navázání z {path}: {error}") from error
    if (saved_start, saved_end, saved_format) != (start, end, snapshot_format):
        raise ValueError("Pro navázání zadejte stejné období a formát snímků jako při prvním spuštění.")
    return selected


def save_snapshots(start, end, snapshot_format, oauth, resume):
    from drive import connect

    with ExitStack() as stack:
        drive = None
        if resume is None:
            drive = stack.enter_context(connect(oauth))
            revisions = [Revision(parse_time(time), source) for source, time in drive.revisions().items()]
            target, selected = start_download(revisions, start, end, snapshot_format)
        else:
            target = resume
            selected = read_download_plan(target, start, end, snapshot_format)

        print(f'Pro případné navázání použijte --navazat "{target.resolve()}".', file=sys.stderr, flush=True)
        snapshots = []
        downloaded = False
        for number, rev in enumerate(selected, 1):
            name = f"revize-{number:05d}.{snapshot_format}"
            path = target / name
            snapshots.append((rev.time.isoformat(), name))
            if path.is_file():
                continue
            if drive is None:
                drive = stack.enter_context(connect(oauth))
            if downloaded:
                sleep(uniform(1, 10))
            with atomic_output(path) as staged:
                if snapshot_format == "xlsx":
                    staged.write_bytes(read_revision(rev, drive.download))
                else:
                    states = read_revision(rev, lambda source: read_snapshot(drive.download(source)))
                    with staged.open("w", encoding="utf-8-sig", newline="") as stream:
                        rows = csv.writer(stream, delimiter=";")
                        rows.writerow(SNAPSHOT_COLUMNS)
                        rows.writerows(states.items())
            downloaded = True

    manifest = target / MANIFEST
    with atomic_output(manifest) as staged:
        with staged.open("w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.writer(stream, delimiter=";")
            writer.writerow(("cas", "soubor"))
            writer.writerows(snapshots)
    return manifest, len(selected)


def check_saved_period(manifest, start, end):
    if manifest.name != MANIFEST:
        return
    path = manifest.with_name(PERIOD)
    if not path.is_file():
        return
    try:
        period = json.loads(path.read_text(encoding="utf-8"))
        saved_start, saved_end = parse_time(period["od"]), parse_time(period["do"])
    except (KeyError, TypeError, ValueError) as error:
        raise ValueError(f"Neplatné období uložené v {path}.") from error
    if not saved_start <= start < end <= saved_end:
        raise ValueError(f"Stažené snímky pokrývají období {local_time(saved_start)} až {local_time(saved_end)}.")


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
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--stahnout", nargs="?", const="csv", choices=("csv", "xlsx"),
        help="Pouze stáhne snímky. Výchozí formát je CSV, volba xlsx uloží celý sešit.",
    )
    mode.add_argument("--snimky", type=Path, help="Zpracuje místní seznam CSV nebo XLSX snímků (cas;soubor).")
    parser.add_argument("--navazat", type=Path, help="Pokračuje ve stahování ve složce vypsané při prvním spuštění.")
    parser.add_argument("--oauth", type=Path, help="JSON s OAuth klientem. Výchozí: credentials.json.")
    args = parser.parse_args(argv)
    if args.snimky and args.oauth:
        parser.error("--oauth nelze použít s místními snímky.")
    if args.snimky and args.navazat:
        parser.error("--navazat nelze použít s --snimky.")
    try:
        start, end = parse_time(args.od), parse_time(args.do)
        if start >= end:
            raise ValueError("Začátek období musí předcházet konci.")
        if args.snimky:
            check_saved_period(args.snimky, start, end)
            result = intervals(list(read_manifest(args.snimky)), read_local_snapshot, start, end)
        else:
            # Do not extrapolate into the future, including time spent authorizing.
            if end > datetime.now(UTC):
                raise ValueError("Konec období pro historii z Google Drive nesmí být v budoucnosti.")
            manifest, count = save_snapshots(
                start, end, args.stahnout or "csv", args.oauth or Path("credentials.json"), args.navazat,
            )
            if args.stahnout:
                print(f"Stažené snímky: {count}")
                print(f"Seznam snímků: {manifest.resolve()}")
                if count == 0:
                    print("Drive nevrátil žádné dostupné revize pro zadané období.", file=sys.stderr)
                return 0
            result = intervals(list(read_manifest(manifest)), read_local_snapshot, start, end)

        timeline, count, baseline = result
        with stage_output(OUTPUT) as (staged, target):
            write_report(staged, timeline, start, end, count, baseline)
        print(f"Přehled: {(target / 'prehled.html').resolve()}")
        if baseline is None:
            print("Pro začátek období chybí výchozí snímek.", file=sys.stderr)
        return 0
    except KeyboardInterrupt:
        print("Běh byl přerušen.", file=sys.stderr)
        return 130
    except ModuleNotFoundError as error:
        print(f"Chybí závislost {error.name}. Spusťte python -m pip install -r requirements.txt.", file=sys.stderr)
        return 1
    except (ValueError, RuntimeError, OSError, csv.Error) as error:
        print(f"Chyba: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())

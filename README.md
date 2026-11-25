# Pochůzky

Lokální program spočítá, kolik hodin měl každý z 21 bodů v listu **Pochůzky BODY** hodnotu **NH**, **DH**, **Civil**, **???** nebo prázdnou buňku. Zdrojová Google tabulka, názvy bodů a povolené hodnoty jsou pevně dané v kódu.

## Instalace

Nainstalujte Python 3.11 nebo novější. V adresáři repozitáře spusťte:

```sh
python -m venv .venv
```

Aktivujte prostředí příkazem `source .venv/bin/activate` na macOS a Linuxu nebo `.venv\Scripts\Activate.ps1` v PowerShellu. Pak nainstalujte závislosti:

```sh
python -m pip install -r requirements.txt
```

## Přístup k historii Google Drive

1. V [Google Cloud Console](https://console.cloud.google.com/) vytvořte nebo vyberte projekt a zapněte [Google Drive API](https://console.cloud.google.com/apis/library/drive.googleapis.com).
2. V Google Auth Platform nastavte obrazovku souhlasu. Pokud aplikaci používáte v režimu External / Testing, přidejte svůj účet mezi testovací uživatele. Program používá oprávnění `https://www.googleapis.com/auth/drive.readonly`.
3. Vytvořte [OAuth klienta typu Desktop app](https://developers.google.com/workspace/drive/api/quickstart/python#authorize_credentials_for_a_desktop_application) a stažený JSON uložte vedle skriptu jako `credentials.json`. Pokud jej máte jinde, předejte cestu pomocí `--oauth`.
4. Spusťte výpočet. V prohlížeči se přihlaste účtem, který může zdrojovou tabulku upravovat a stahovat. OAuth token zůstane v paměti programu.

```sh
python pochuzky.py --od 2026-11-15 --do 2026-11-22
```

Počítá se od začátku včetně do konce bez tohoto okamžiku. Příklad zahrnuje 15.–21. listopadu. Můžete zadat i konkrétní čas, například `2026-11-15T09:30`. Časy bez posunu se berou v pásmu Europe/Prague. Při dvojznačném čase použijte posun, například `2026-10-25T02:30:00+01:00`.

Program běží na vašem počítači. Internet potřebuje pro přihlášení, výpis revizí a export z Google Drive. Přihlašovací JSON i výstupy ponechte mimo git.

## Výstup

Každý výpočet vytvoří vlastní adresář pod `vystup/` a vypíše cestu k `prehled.html`. Přehled otevřete v prohlížeči. Obsahuje barevnou časovou osu a součty v hodinách. Vedle něj jsou `hodiny.csv` se součty a `intervaly.csv` s intervaly, jejich hodnotami a délkou. CSV používají UTF‑8, středník a desetinnou čárku.

Výpočet vezme poslední dostupnou revizi nejpozději na začátku období a projde dostupné revize uvnitř období. Hodnota ze snímku se počítá od času jeho revize do další zaznamenané změny nebo do konce období. Řádky se párují podle názvu bodu. U názvů a hodnot se ignoruje velikost písmen a nadbytečné mezery.

Sloupec **Nezjištěno** pokrývá dobu před prvním dostupným snímkem, chybějící bod a případ, kdy stejný čas má více snímků s různými hodnotami. Existující bod s prázdnou buňkou se počítá do sloupce **Prázdná**. Pokud některou potřebnou revizi nelze stáhnout nebo její formát neodpovídá očekávání, běh skončí chybou.

Google může revize sloučit nebo některé přes API nevrátit. Součty jsou rekonstrukce ze snímků, a proto se mohou lišit od času skutečných změn. Program používá [Drive API pro výpis revizí](https://developers.google.com/workspace/drive/api/reference/rest/v3/revisions/list) a [export konkrétní revize do XLSX](https://developers.google.com/workspace/drive/api/guides/manage-downloads#export_google_workspace_document_content_using_long-running_operations).

## Stažení a výpočet ve dvou krocích

Přepínač `--stahnout` uloží každou dostupnou revizi pro zadané období a výchozí revizi. Výchozím formátem je CSV se sloupci **NÁZEV BODU** a **STRANA** z listu **Pochůzky BODY**. Program stáhne XLSX z Drive a převede jej na CSV lokálně, protože export CSV přímo z Drive obsahuje jen první list.

```powershell
python pochuzky.py --od 2026-11-15 --do 2026-11-22 --stahnout
```

Chcete-li uložit všechny listy v každé revizi, použijte `--stahnout xlsx`. Soubory budou v nové složce uvnitř systémové dočasné složky (na Windows pod `%TEMP%`). Pokud je chcete mít dlouhodobě, přesuňte celou vytvořenou složku jinam.

```powershell
python pochuzky.py --od 2026-11-15 --do 2026-11-22 --stahnout xlsx
```

Obě varianty vypíšou cestu k `snimky.csv`. Z ní pak vytvoříte report bez připojení ke Googlu. Zadejte původní období nebo jeho část a použijte vypsanou cestu.

```powershell
python pochuzky.py --od 2026-11-15 --do 2026-11-22 --snimky "C:\cesta\k\snimky.csv"
```

Bez přepínačů `--stahnout` a `--snimky` se stažení a výpočet provedou v jednom běhu.

## Pokračování po přerušení stahování

Mezi stahováním jednotlivých revizí program čeká náhodně 1–10 sekund. Každý hotový snímek ukládá ihned, a to i při přímém výpočtu reportu. Před stahováním vypíše cestu ke složce se snímky.

Po přerušení nebo chybě HTTP 429 spusťte původní příkaz se stejným obdobím a formátem a přidejte `--navazat` s vypsanou cestou ke složce. Program přeskočí uložené snímky a stáhne zbývající revize z původního seznamu.

```powershell
python pochuzky.py --od 2026-11-15 --do 2026-11-22 --stahnout --navazat "vystup/snimky-abc123"
```

## Výpočet z vlastních XLSX souborů

Připravte seznam v CSV v UTF‑8 s časem pořízení a cestou ke každému XLSX snímku. Relativní cesty se vyhodnocují vůči umístění CSV.

```csv
cas;soubor
2026-11-14T22:00:00+01:00;sobota.xlsx
2026-11-17T09:30:00+01:00;utery.xlsx
```

```sh
python pochuzky.py --od 2026-11-15 --do 2026-11-22 --snimky snimky.csv
```

Jeden XLSX export obsahuje stav z jediného okamžiku. Pro výpočet změn z místních souborů potřebujete snímky z různých časů.

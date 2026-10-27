# IED → OMICRON Auto Config  (IED2OMICRON)

Desktop tool that turns the data of an IEC 61850 substation project (**General_PB** Excel workbook + device files + CID files) into ready-to-load **OMICRON Test Universe** configurations:

- **Hardware configuration** (`.ohc`) — binary input / output names, Sampled Values (SV) analog channels, SV streams and dataset mapping
- **GOOSE configuration** (`.xml`) — GOOSE subscriptions and simulations mapped to the CMC inputs / outputs
- **Update** of existing GOOSE configurations from new CID files, keeping the mappings made in Test Universe
- **Test objects** (`.xrio`) — nameplate, location and nominal values of the device (Device Settings)

> **Status: MVP (v0.0.3).** This is an attempt to automate part of the work on preparing and carrying out tests at a substation. It is not a finished product — always check the generated files in OMICRON Test Universe before using them on site.

What General_PB is: [IEC61850_text_configurator](https://github.com/IevgenGerman/IEC61850_text_configurator)

---

## Download

The Windows executable is attached to every release:
**[Releases](../../releases)** → `IED2OMICRON_<version>.exe` — no Python needed.

## Run from source

```bash
git clone https://github.com/IevgenGerman/OMICRON_CONFIGURATION_FOR_IED.git
cd OMICRON_CONFIGURATION_FOR_IED
pip install -r requirements.txt
python ied2omicron.py
```

Requires Python 3.10+ with tkinter (included in the standard Windows installer).

## Project folder

```text
<Project>\
  General_PB.xlsm          main workbook (sheets Settings_IED, Device_Type, GEN, SV; project name / address)
  DOC\                     device files <IED>.xlsm (sheets Omicron Binary Inputs / Outputs, Omicron AI <test>)
  Library\
    default_01.ohc         master hardware configuration (saved in Test Universe)
    default_goose.xml      master GOOSE configuration (saved in Test Universe)
    test_object_default.xrio   master test object (saved in Test Universe)
  CID\                     CID / ICD / SCD files of all devices taking part in GOOSE
  Omicron\                 created by the program, one sub-folder per IED
```

The master files are empty configurations saved once in Test Universe (CMC type, PTP, network port, GOOSE virtual I/O 7×24, three SV streams). The program fills them; general settings come from the master.

## Usage

| Button | Result |
|---|---|
| **Create Omicron_configurations** | `Omicron\<IED>\<IED>_Omicron_AI_Test_<n>.ohc` (one per AI test sheet) |
| **Create GOOSE_configurations** | `Omicron\<IED>\<IED>_GOOSE.xml` — stops and lists missing CID files before writing anything |
| **Update GOOSE_configurations** | choose `*_GOOSE.xml` file(s) and a folder with new CID files → `<name>_updated_<date>_<time>.xml` next to the original |
| **Create XRIO_test_objects** | `Omicron\<IED>\test_object_<IED>_AI_Test_<n>.xrio` (one per AI test sheet with SV subscriptions); a device without SV subscriptions → `Omicron\<IED>\test_object_<IED>.xrio` (nominal values from the master); an existing file is kept, the new one gets `_<date>_<time>` |
| **Open Omicron folder** | opens `<Project>\Omicron` |
| **i** | built-in user manual (English / Ukrainian) |

Before creating, a list of devices with check boxes is shown. Every action writes a report (`_omicron_report.txt`, `_goose_report.txt`, `_goose_update_report_<date>_<time>.txt`, `_xrio_report.txt`).

The generators also work from the command line:

```bash
python omicron_configuration.py "D:\Projects\Bucha" --all
python omicron_goose_config.py  "D:\Projects\Bucha" --device W1G_D60
python omicron_goose_config.py  --update W1G_D60_GOOSE.xml --cid "D:\NEW_CID"
python omicron_xrio.py           "D:\Projects\Bucha" --device W1G_D60
```

## How it works (short)

- **`.ohc`** — binary MFC archive; the program patches names and values in a master file in place (no structure is created from scratch). SV stream parameters come from General_PB sheet **SV** (svID, MAC, APPID, VLAN, priority, dataset); the CMC channel names from the Omicron sheets of the device file.
- **GOOSE `.xml`** — GCBs are found in the CID files **by MAC address** (so CID IED names may differ from project names). MAC, APPID, GoID, references, VLAN, ConfRev, MinTime/MaxTime and the full data set come from the CID. CMC I/O numbers: `Bin n → n`, `vBin g-k → 24 + (g-1)·24 + k`.
- **Update** — blocks matched by GoCB reference, fallback by MAC; mappings of signals that still exist are kept, added signals come without mapping, removed ones are reported.
- **`.xrio`** — only the `<Value>` of the agreed parameters is replaced in the master:

  | Test Universe | Source in General_PB |
  |---|---|
  | Name/description | IEDName |
  | Manufacturer | Device_Type → Manufacturer of the device type |
  | Device type | GEN → Order code (else Device_Type → order_code) |
  | Device address | GEN → IP SB (else IP PB) |
  | Substation name / address | Editor → Project → Project name / Project address |
  | Bay name | Settings_IED → BAY |
  | Nominal values, residual factors | SV columns N..W of the subscribed streams |

  All SV streams of one test must have the same nominal values; if they differ or are empty, no file is written for that test and the program says why.

## Files

| File | Purpose |
|---|---|
| `ied2omicron.py` | GUI (Tkinter) — start here |
| `omicron_configuration.py` | `.ohc` generator (also CLI) |
| `omicron_goose_config.py` | GOOSE `.xml` generator and updater (also CLI) |
| `omicron_xrio.py` | test object `.xrio` generator (also CLI) |
| `ied_doc_manual.py` | user manual texts (EN / UK) shown by the **i** button |

## Build the exe

```bash
pip install pyinstaller
pyinstaller --clean --onefile --windowed --name IED2OMICRON_v0.0.4 ied2omicron.py
```

Result: `dist\IED2OMICRON_v0.0.4.exe` (or just run `build_exe.bat` - it takes the version from `VERSION` in ied2omicron.py). Pushing a tag `v*` builds the exe automatically with GitHub Actions and attaches it to the release (`.github/workflows/build.yml`).

## Limitations

- GOOSE: only BOOLEAN signals and quality (`.q`) are supported; GCBs with other data types (Dbpos, INT, FLOAT, Timestamp) are skipped and reported.
- Tested with GE devices (UR / MU320e / 850) and OMICRON CMC 356 configurations.

## Disclaimer

Independent tool, not affiliated with or endorsed by OMICRON electronics. OMICRON and Test Universe are trademarks of OMICRON electronics GmbH.

## Author

[Ievgen German](https://www.linkedin.com/in/ievgen-german-945b16b6)

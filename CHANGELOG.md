# Changelog

## v0.0.2
- Renamed to **IED → OMICRON Auto Config** (`IED2OMICRON.exe`, main file `ied2omicron.py`).
- **Update GOOSE_configurations**: update existing `*_GOOSE.xml` files from new CID files.
  - GCB blocks matched by GoCB reference (IED + LD + GCB name), fallback by MAC.
  - MAC, APPID, GoID, references, VLAN, ConfRev, repetition and data set taken from the CID.
  - Enable, SimulateFlag, Multiplier and every CMC mapping of a signal that still exists are kept.
  - GCB gone from its device CID → block removed; no CID for the device → block unchanged.
  - Output `<name>_updated_<date>_<time>.xml` next to the original, plus an update report.
- Built-in user manual (button **i**), English / Ukrainian.

## v0.0.1
- Tkinter GUI (first name: IED_DOC_CONFIGURATOR).
- **Create Omicron_configurations**: OMICRON Test Universe hardware configurations (`.ohc`):
  binary input / output names, SV analog channel names, SV streams and dataset mapping.
- **Create GOOSE_configurations**: GOOSE subscriptions and simulations (`.xml`) built from CID files,
  with a check for missing CID files before anything is written.
- Output in `<project>\Omicron\<IED>\`, reports `_omicron_report.txt` / `_goose_report.txt`.

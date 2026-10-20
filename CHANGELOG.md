# Changelog

## v0.0.3
- **Create XRIO_test_objects**: OMICRON Test Universe test objects (`.xrio`) from `Library\test_object_default.xrio`.
  - One test object per `Omicron AI <test>` sheet with SV subscriptions: `Omicron\<IED>\test_object_<IED>_AI_Test_<n>.xrio`.
  - Name/description, Manufacturer, Device type (order code), Device address (IP SB / IP PB),
    Substation name / address, Bay name from General_PB.
  - Nominal values (f, V, I, primary values) and residual factors (VLN/VN, IN/Inom, VN sec, IN sec,
    instrument transformer) from the SV sheet of General_PB.
  - All streams of a test must have equal, filled nominal values - otherwise no file, a message and a report line.
  - A device without any SV subscription: one `test_object_<IED>.xrio` with nameplate / location only,
    nominal values left as in the master.
  - An existing file is never overwritten: the new file gets the date and time in its name.
  - Everything else of the master stays byte for byte.
- Needs General_PB with the nominal values on Device_Type (SVCB blocks) / SV and the Project address field.
- "Open Omicron folder" moved next to "Browse...".

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

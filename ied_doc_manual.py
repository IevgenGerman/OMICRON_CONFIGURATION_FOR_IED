# -*- coding: utf-8 -*-
"""
ied_doc_manual.py - user manual of "IED → OMICRON Auto Config" (Ukrainian / English),
shown by the button "i".

Markup (one paragraph per line):
  "# text"            title
  "## text"           section heading
  "    text"          (4 spaces) fixed-width line: folder trees, tables
  [[label|url]]       clickable link
"""

URL_SOURCE = "https://github.com/IevgenGerman/OMICRON_CONFIGURATION_FOR_IED"
URL_GENERAL_PB = "https://github.com/IevgenGerman/IEC61850_text_configurator"
URL_AUTHOR = "https://www.linkedin.com/in/ievgen-german-945b16b6"

LANGS = [("en", "English"), ("uk", "Українська")]
DEFAULT_LANG = "en"

MANUAL = {}

# =====================================================================================
MANUAL["uk"] = """# IED → OMICRON Auto Config  {version} — посібник користувача

Вихідний код програми: [[OMICRON_CONFIGURATION_FOR_IED|%(src)s]]

## 1. Про програму
Програма створена як спроба автоматизувати частину робіт з підготовки та безпосереднього проведення випробувань на підстанції.
По суті це MVP (мінімально життєздатний продукт) для майбутнього проєкту. Це не готовий продукт, тому в ньому можуть бути неточності. Перед використанням на об'єкті завжди перевіряйте створені файли в OMICRON Test Universe.
Програма бере дані проєкту General_PB (Excel) і файли CID та створює файли для OMICRON Test Universe. Що таке General_PB: [[IEC61850_text_configurator|%(gpb)s]]
    • конфігурацію обладнання (*.ohc) — назви бінарних входів/виходів, аналогових каналів і потоки Sampled Values (SV);
    • конфігурацію GOOSE (*.xml) — підписки і симуляції GOOSE з прив'язкою до входів/виходів CMC;
    • оновлення вже готових конфігурацій GOOSE за новими CID.
Файли Excel і CID програма тільки читає й ніколи не змінює.
Незалежний інструмент, не пов'язаний з OMICRON electronics і не схвалений нею. OMICRON і Test Universe — торговельні марки OMICRON electronics GmbH.

## 2. Що потрібно
Запуск: exe-файл IED2OMICRON (нічого встановлювати не треба) або з PyCharm / Python: ied2omicron.py. Поруч мають лежати omicron_configuration.py, omicron_goose_config.py та ied_doc_manual.py; потрібен пакет openpyxl.
Структура папки проєкту:
    <Проєкт>\\
      General_PB.xlsm          головна книга проєкту
      DOC\\                     файли пристроїв <IED>.xlsm
      Library\\                 майстер-файли:
         default_01.ohc        майстер конфігурації обладнання
         default_goose.xml     майстер конфігурації GOOSE
      CID\\                     CID (*.cid, *.icd, *.scd, *.iid) усіх пристроїв, що беруть участь у GOOSE
      Omicron\\                 СТВОРЮЄТЬСЯ ПРОГРАМОЮ — результати
         <IED>\\                окрема папка для кожного пристрою
Майстер-файли — це порожні конфігурації, збережені в Test Universe. Програма не створює їх з нуля, а заповнює:
    • default_01.ohc — з увімкненими віртуальними входами/виходами GOOSE (7×24) і трьома потоками SV (6 груп аналогових каналів по 4);
    • default_goose.xml — порожня конфігурація модуля GOOSE.
Загальні налаштування (CMC, PTP, IRIG-B, мережевий порт тощо) задаються один раз у майстрі й переходять у всі створені файли.

## 3. Головне вікно
    Заголовок         назва та версія програми
    Кнопка  i         цей посібник (мову можна обрати вгорі панелі)
    Project folder    шлях до папки проєкту
    Browse...         вибір папки проєкту; програма запам'ятовує останню папку
    Create Omicron_configurations   — розділ 4
    Create GOOSE_configurations     — розділ 5
    Update GOOSE_configurations     — розділ 6
    Open Omicron folder             — розділ 7
    Журнал            хід роботи і підсумковий звіт кожної дії
    Рядок стану       Ready / Working… / результат / Stopped
Поки триває дія, кнопки неактивні, а вікно лишається відкритим (можна гортати журнал). Після завершення з'являється повідомлення.

## 4. Create Omicron_configurations (файли *.ohc)
Що робить: для кожного вибраного пристрою створює конфігурацію обладнання Test Universe — назви бінарних входів і виходів, назви аналогових каналів SV і параметри потоків SV.
Порядок дій:
    1. Вкажіть папку проєкту.
    2. Натисніть кнопку — з'явиться список пристроїв (файли DOC\\*.xlsm) з галочками; кнопки Select all / None.
    3. Якщо в Library кілька файлів *.ohc, над списком з'явиться вибір майстра.
    4. Відмітьте пристрої і натисніть Create (Cancel — скасувати).
Звідки беруться дані (файл пристрою DOC\\<IED>.xlsm):
    • «Omicron Binary Inputs»: A — канал (Bin. in 1…12, vBin. in 1-1 … 7-24), D — Omicron ID → назва бінарного входу CMC;
    • «Omicron Binary Outputs»: A — канал (Bin. out 1…8, vBin. out 1-1 … 7-24), D — Omicron ID → назва бінарного виходу CMC;
    • «Omicron AI <тест>» (Test 1, Test 2, …): A — канал, D — Omicron ID → назва аналогового каналу, F — SV ID → потік групи:
        V L1, V L2, V L3, V E      → SV Voltage 1      I L1 … I E      → SV Current 1
        V(2)-1 … V(2)-E           → SV Voltage 2      I(2)-1 … I(2)-E → SV Current 2
        V(3)-1 … V(3)-E           → SV Voltage 3      I(3)-1 … I(3)-E → SV Current 3
      (група n = потік SV n)
Порожній Omicron ID → канал зберігає назву з майстра («vBin. in 1-2», «V(2)-1» …).
Потоки SV (з General_PB, аркуш SV, за SV ID) — у вікні Test Universe «Configure Sampled Values», вкладка Stream n:
    svID, Destination MAC, APPID, VLAN ID, VLAN priority     — з аркуша SV
    Dataset reference        = IEDName + Dataset   (MU01_QW1GMU01/LLN0$PhsMeas1)
    Control block reference  = IEDName + LD/LLN0$MS$MSVCB0<номер LD>
    ConfRev = 1, порт ETH1
Також заповнюється Dataset Mapping — який канал CMC іде в який атрибут набору даних. Підписи атрибутів (MU01.TCTR1.AmpSv…) і рядок «Select IED / Sampled Values ID» лишаються з майстра. Потік без SV ID лишається таким, як у майстрі.
Результат:
    Omicron\\<IED>\\<IED>_Omicron_AI_Test_<n>.ohc   — один файл на кожен аркуш «Omicron AI …» з даними
    Omicron\\<IED>\\<IED>_Omicron.ohc               — якщо даних SV немає (лише GOOSE)
    Omicron\\_omicron_report.txt                   — звіт
Пристрій без даних Omicron пропускається (рядок у звіті). Бінарні входи/виходи однакові в усіх тестах пристрою.
Звіт попереджає про: SV ID, якого немає на аркуші SV General_PB; дві різні SV ID в одній групі V(n) / I(n); назви з недозволеними символами (дозволено A-Z a-z 0-9 _, перший символ — літера); канали з аркуша, яких немає в майстрі.

## 5. Create GOOSE_configurations (файли *_GOOSE.xml)
Що робить: для кожного вибраного пристрою створює конфігурацію модуля GOOSE Test Universe:
    • Підписки (CMC ПРИЙМАЄ GOOSE пристрою) — аркуш «Omicron Binary Inputs»: кожен GCB пристрою з цього аркуша;
    • Симуляції (CMC ПУБЛІКУЄ GOOSE замість інших пристроїв) — аркуш «Omicron Binary Outputs»: кожен GCB видавця (стовпці E — IED Name, F — GCB Name).
Порядок дій:
    1. Покладіть у папку CID проєкту CID пристрою і всіх видавців, на яких він підписаний.
    2. Натисніть кнопку, відмітьте пристрої, натисніть Create.
Звідки беруться дані:
    • з CID: MAC, APPID, GoID, посилання на GCB і набір даних, VLAN ID і пріоритет, ConfRev, MinTime/MaxTime (повтори Initial/Final) і ПОВНИЙ склад набору даних (stVal і якість .q);
    • з файлу пристрою: які сигнали на які входи/виходи CMC — стовпець B (ELEMENT) — сигнал, стовпець A — канал:
        Bin. in n / Bin. out n              → вхід/вихід CMC n
        vBin. in g-k / vBin. out g-k        → 24 + (g−1)·24 + k    (1-3 → 27, 7-24 → 192)
GCB шукається в CID за MAC-адресою (з аркуша або з General_PB, аркуш GEN), тому ім'я IED у CID може відрізнятися від назви в проєкті.
Перевірка CID: якщо для будь-якого потрібного GCB немає CID, програма НІЧОГО НЕ ЗАПИСУЄ і показує список відсутніх (англійською), наприклад:
    CID files are missing in …\\CID
      W1G_D60: MU02_QC1G / GCB VT (MAC 01-0C-CD-01-00-64) - no CID file with this GCB
Додайте файли й запустіть знову.
Результат: Omicron\\<IED>\\<IED>_GOOSE.xml і Omicron\\_goose_report.txt. Звіт також показує розбіжності General_PB ↔ CID (GoID, APPID, VLAN) — у файл іде значення з CID, тобто реальне значення пристрою; і сигнали з аркуша, яких немає в наборі даних CID.
Обмеження версії: підтримуються сигнали Boolean і якість (.q). GCB з іншими типами (Dbpos, INT, FLOAT, Timestamp) пропускається з приміткою у звіті.

## 6. Update GOOSE_configurations (оновлення за новими CID)
Навіщо: пристрій переконфігурували — у GCB додали, видалили або переставили сигнали, змінили MAC, APPID, VLAN чи ConfRev. Кнопка оновлює вже готовий *_GOOSE.xml і зберігає вашу ручну роботу в Test Universe, тож не доводиться перебивати конфігурацію вручну.
Порядок дій:
    1. Натисніть кнопку — відкриється вікно «Відкрити файли»; виберіть один або кілька *_GOOSE.xml (Ctrl / Shift).
    2. Виберіть папку з НОВИМИ CID (можна покласти лише ті, що змінились).
    Обидві папки програма запам'ятовує.
Як оновлюється кожен блок GCB: GCB шукається в нових CID за посиланням (IED + LD + назва GCB), а якщо не знайдено — за MAC.
    Знайдено:    з CID — MAC, APPID, GoID, посилання, VLAN, ConfRev, Initial/Final, склад і порядок сигналів;
                 з XML — Enable, SimulateFlag, Multiplier і всі прив'язки до входів/виходів CMC (номер, Bit, Value, Inverted)
       сигнал залишився        → прив'язка зберігається
       сигнал додано           → додається без прив'язки
       сигнал видалено         → зникає, у звіті «mapping … lost»
       змінився тип сигналу    → прив'язка знімається (у звіті)
       сигнали переставлено    → порядок як у CID, прив'язки переходять разом із сигналами
    CID пристрою є, а GCB у ньому немає   → блок видаляється («GCB removed»)
    CID пристрою немає в папці            → блок не змінюється («no CID for this device, left unchanged»)
Перейменований сигнал (Ind053 → Ind056) виглядає як «видалено + додано» — його прив'язку треба зробити заново. Нові сигнали автоматично на вільні входи НЕ ставляться.
Результат: поруч з оригіналом — <IED>_GOOSE_updated_<дата>_<час>.xml (оригінал не змінюється). Файл записується, лише якщо щось змінилось; інакше у звіті «nothing to update». Звіт _goose_update_report_<дата>_<час>.txt — у тій самій папці та в журналі вікна.
Excel і General_PB не змінюються. Якщо склад набору даних змінився, оновіть його і в General_PB.

## 7. Open Omicron folder
Відкриває папку <Проєкт>\\Omicron у Провіднику. Якщо її ще немає, програма про це повідомить.

## 8. Типові повідомлення і що робити
    Project folder not found / General_PB.xlsm not found  → вкажіть папку, де лежать General_PB.xlsm і DOC
    No master .ohc file / Master file not found           → покладіть default_01.ohc або default_goose.xml у Library
    Folder with the CID files not found                   → створіть папку CID у проєкті
    CID files are missing …                               → додайте CID названих пристроїв (розділ 5)
    SV ID … not on General_PB SV                          → SV ID з аркуша тесту немає в General_PB; потік лишається як у майстрі
    data type(s) … not supported yet                      → у наборі даних тип, який ще не підтримується (розділ 5)
    not read - …                                          → файл пошкоджений або нечитабельний
Відкритий в Excel файл програма читає зі своєї тимчасової копії — закривати Excel не треба.

[[Євген Герман|%(author)s]]
""" % dict(src=URL_SOURCE, gpb=URL_GENERAL_PB, author=URL_AUTHOR)

# =====================================================================================
MANUAL["en"] = """# IED → OMICRON Auto Config  {version} — user manual

Source code of the program: [[OMICRON_CONFIGURATION_FOR_IED|%(src)s]]

## 1. About
This program was created as an attempt to automate part of the work on preparing and directly carrying out tests at a substation.
It is essentially an MVP (minimum viable product) for a future project. It is not a finished product, so there may be inaccuracies. Always check the created files in OMICRON Test Universe before using them on site.
The program takes the data of a General_PB (Excel) project and the CID files and creates files for OMICRON Test Universe. What General_PB is: [[IEC61850_text_configurator|%(gpb)s]]
    • hardware configuration (*.ohc) — names of binary inputs/outputs, analog channels and the Sampled Values (SV) streams;
    • GOOSE configuration (*.xml) — GOOSE subscriptions and simulations mapped to the CMC inputs/outputs;
    • update of existing GOOSE configurations from new CID files.
The program only reads the Excel and CID files and never changes them.
Independent tool, not affiliated with or endorsed by OMICRON electronics. OMICRON and Test Universe are trademarks of OMICRON electronics GmbH.

## 2. Requirements
Start: the exe file IED2OMICRON (nothing to install) or from PyCharm / Python: ied2omicron.py. Next to it: omicron_configuration.py, omicron_goose_config.py and ied_doc_manual.py; the package openpyxl is needed.
Project folder:
    <Project>\\
      General_PB.xlsm          main workbook of the project
      DOC\\                     device files <IED>.xlsm
      Library\\                 master files:
         default_01.ohc        master of the hardware configuration
         default_goose.xml     master of the GOOSE configuration
      CID\\                     CID files (*.cid, *.icd, *.scd, *.iid) of all devices taking part in GOOSE
      Omicron\\                 CREATED BY THE PROGRAM — results
         <IED>\\                one folder per device
The master files are empty configurations saved in Test Universe. The program does not build them from scratch, it fills them:
    • default_01.ohc — with the GOOSE virtual inputs/outputs enabled (7×24) and three SV streams (6 groups of 4 analog channels);
    • default_goose.xml — an empty configuration of the GOOSE module.
General settings (CMC, PTP, IRIG-B, network port etc.) are set once in the master and go into every created file.

## 3. Main window
    Title             name and version of the program
    Button  i         this manual (choose the language at the top of the panel)
    Project folder    path of the project folder
    Browse...         choose the project folder; the last folder is remembered
    Create Omicron_configurations   — section 4
    Create GOOSE_configurations     — section 5
    Update GOOSE_configurations     — section 6
    Open Omicron folder             — section 7
    Log               progress and the final report of every action
    Status line       Ready / Working… / result / Stopped
While an action runs the buttons are disabled and the window stays open (the log can be scrolled). A message appears at the end.

## 4. Create Omicron_configurations (*.ohc files)
What it does: for every selected device it creates a Test Universe hardware configuration — names of binary inputs and outputs, names of the SV analog channels and the parameters of the SV streams.
Steps:
    1. Set the project folder.
    2. Press the button — the list of devices (files DOC\\*.xlsm) with check boxes appears; buttons Select all / None.
    3. If Library has several *.ohc files, a master choice appears above the list.
    4. Tick the devices and press Create (Cancel — abort).
Where the data come from (device file DOC\\<IED>.xlsm):
    • "Omicron Binary Inputs": A — channel (Bin. in 1…12, vBin. in 1-1 … 7-24), D — Omicron ID → name of the CMC binary input;
    • "Omicron Binary Outputs": A — channel (Bin. out 1…8, vBin. out 1-1 … 7-24), D — Omicron ID → name of the CMC binary output;
    • "Omicron AI <test>" (Test 1, Test 2, …): A — channel, D — Omicron ID → analog channel name, F — SV ID → stream of the group:
        V L1, V L2, V L3, V E      → SV Voltage 1      I L1 … I E      → SV Current 1
        V(2)-1 … V(2)-E           → SV Voltage 2      I(2)-1 … I(2)-E → SV Current 2
        V(3)-1 … V(3)-E           → SV Voltage 3      I(3)-1 … I(3)-E → SV Current 3
      (group n = SV stream n)
Empty Omicron ID → the channel keeps the name of the master ("vBin. in 1-2", "V(2)-1" …).
SV streams (from General_PB, sheet SV, by SV ID) — in the Test Universe dialog "Configure Sampled Values", tab Stream n:
    svID, Destination MAC, APPID, VLAN ID, VLAN priority     — from the sheet SV
    Dataset reference        = IEDName + Dataset   (MU01_QW1GMU01/LLN0$PhsMeas1)
    Control block reference  = IEDName + LD/LLN0$MS$MSVCB0<LD number>
    ConfRev = 1, port ETH1
The Dataset Mapping is filled too — which CMC channel goes into which attribute of the data set. Attribute captions (MU01.TCTR1.AmpSv…) and the line "Select IED / Sampled Values ID" stay from the master. A stream without SV ID stays as in the master.
Result:
    Omicron\\<IED>\\<IED>_Omicron_AI_Test_<n>.ohc   — one file per sheet "Omicron AI …" with data
    Omicron\\<IED>\\<IED>_Omicron.ohc               — when there is no SV data (GOOSE only)
    Omicron\\_omicron_report.txt                   — report
A device without Omicron data is skipped (a line in the report). The binary inputs/outputs are the same in all tests of a device.
The report warns about: SV ID not on the sheet SV of General_PB; two different SV IDs in one group V(n) / I(n); names with characters that are not allowed (A-Z a-z 0-9 _, first character a letter); channels of the sheet that the master does not have.

## 5. Create GOOSE_configurations (*_GOOSE.xml files)
What it does: for every selected device it creates a configuration of the Test Universe GOOSE module:
    • Subscriptions (the CMC RECEIVES the GOOSE of the device) — sheet "Omicron Binary Inputs": every GCB of the device on this sheet;
    • Simulations (the CMC PUBLISHES GOOSE instead of other devices) — sheet "Omicron Binary Outputs": every publisher GCB (columns E — IED Name, F — GCB Name).
Steps:
    1. Put the CID of the device and of all publishers it subscribes to into the CID folder of the project.
    2. Press the button, tick the devices, press Create.
Where the data come from:
    • from the CID: MAC, APPID, GoID, GCB and data set references, VLAN ID and priority, ConfRev, MinTime/MaxTime (repetition Initial/Final) and the WHOLE data set (stVal and quality .q);
    • from the device file: which signals go to which CMC inputs/outputs — column B (ELEMENT) is the signal, column A the channel:
        Bin. in n / Bin. out n              → CMC input/output n
        vBin. in g-k / vBin. out g-k        → 24 + (g−1)·24 + k    (1-3 → 27, 7-24 → 192)
A GCB is found in the CID files by its MAC address (from the sheet or from General_PB, sheet GEN), so the IED name inside the CID may differ from the project name.
CID check: if any needed GCB has no CID file, the program writes NOTHING and lists the missing ones, for example:
    CID files are missing in …\\CID
      W1G_D60: MU02_QC1G / GCB VT (MAC 01-0C-CD-01-00-64) - no CID file with this GCB
Add the files and run again.
Result: Omicron\\<IED>\\<IED>_GOOSE.xml and Omicron\\_goose_report.txt. The report also shows differences General_PB ↔ CID (GoID, APPID, VLAN) — the file gets the CID value, i.e. the real value of the device; and signals of the sheet that are not in the CID data set.
Limits of this version: Boolean signals and quality (.q) are supported. A GCB with other types (Dbpos, INT, FLOAT, Timestamp) is skipped with a note in the report.

## 6. Update GOOSE_configurations (from new CID files)
Why: a device was reconfigured — signals were added, removed or moved in a GCB, MAC, APPID, VLAN or ConfRev changed. The button updates an existing *_GOOSE.xml and keeps your manual work in Test Universe, so nothing has to be re-entered by hand.
Steps:
    1. Press the button — the "Open files" dialog appears; choose one or more *_GOOSE.xml (Ctrl / Shift).
    2. Choose the folder with the NEW CID files (only the changed ones are enough).
    Both folders are remembered.
How each GCB block is updated: the GCB is searched in the new CID files by its reference (IED + LD + GCB name), else by MAC.
    Found:       from the CID — MAC, APPID, GoID, references, VLAN, ConfRev, Initial/Final, signals and their order;
                 from the XML — Enable, SimulateFlag, Multiplier and every CMC mapping (number, Bit, Value, Inverted)
       signal still there      → mapping kept
       signal added            → added without a mapping
       signal removed          → removed, report "mapping … lost"
       signal type changed     → mapping removed (in the report)
       signals moved           → order as in the CID, the mappings move with their signals
    CID of the device present, GCB not   → block removed ("GCB removed")
    no CID of the device in the folder   → block unchanged ("no CID for this device, left unchanged")
A renamed signal (Ind053 → Ind056) looks like "removed + added" — map it again. New signals are NOT put on free inputs automatically.
Result: next to the original — <IED>_GOOSE_updated_<date>_<time>.xml (the original is not changed). The file is written only when something changed; otherwise the report says "nothing to update". Report _goose_update_report_<date>_<time>.txt — in the same folder and in the log.
Excel and General_PB are not changed. If a data set changed, update it in General_PB too.

## 7. Open Omicron folder
Opens the folder <Project>\\Omicron in Explorer. If it does not exist yet, the program says so.

## 8. Typical messages and what to do
    Project folder not found / General_PB.xlsm not found  → choose the folder with General_PB.xlsm and DOC
    No master .ohc file / Master file not found           → put default_01.ohc or default_goose.xml into Library
    Folder with the CID files not found                   → create the folder CID in the project
    CID files are missing …                               → add the CID files of the listed devices (section 5)
    SV ID … not on General_PB SV                          → the SV ID of the test sheet is not in General_PB; the stream stays as in the master
    data type(s) … not supported yet                      → the data set has a type that is not supported yet (section 5)
    not read - …                                          → the file is damaged or cannot be read
A file open in Excel is read from a temporary copy — Excel does not have to be closed.

[[Ievgen German|%(author)s]]
""" % dict(src=URL_SOURCE, gpb=URL_GENERAL_PB, author=URL_AUTHOR)

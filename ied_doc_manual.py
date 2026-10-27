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
    • оновлення вже готових конфігурацій GOOSE за новими CID;
    • об'єкти випробувань (*.xrio) — паспорт пристрою, місце встановлення і номінальні значення.
Файли Excel і CID програма тільки читає й ніколи не змінює.
Незалежний інструмент, не пов'язаний з OMICRON electronics і не схвалений нею. OMICRON і Test Universe — торговельні марки OMICRON electronics GmbH.

## 2. Що потрібно
Запуск: exe-файл IED2OMICRON (нічого встановлювати не треба) або з PyCharm / Python: ied2omicron.py. Поруч мають лежати omicron_configuration.py, omicron_goose_config.py, omicron_xrio.py та ied_doc_manual.py; потрібен пакет openpyxl.
Структура папки проєкту:
    <Проєкт>\\
      General_PB.xlsm          головна книга проєкту
      DOC\\                     файли пристроїв <IED>.xlsm
      Library\\                 майстер-файли:
         default_01.ohc        майстер конфігурації обладнання
         default_goose.xml     майстер конфігурації GOOSE
         test_object_default.xrio   майстер об'єкта випробувань
      CID\\                     CID (*.cid, *.icd, *.scd, *.iid) усіх пристроїв, що беруть участь у GOOSE
      Omicron\\                 СТВОРЮЄТЬСЯ ПРОГРАМОЮ — результати
         <IED>\\                окрема папка для кожного пристрою
Майстер-файли — це порожні конфігурації, збережені в Test Universe. Програма не створює їх з нуля, а заповнює:
    • default_01.ohc — з увімкненими віртуальними входами/виходами GOOSE (7×24) і трьома потоками SV (6 груп аналогових каналів по 4);
    • default_goose.xml — порожня конфігурація модуля GOOSE;
    • test_object_default.xrio — об'єкт випробувань, збережений у Test Universe (Device Settings).
Загальні налаштування (CMC, PTP, IRIG-B, мережевий порт тощо) задаються один раз у майстрі й переходять у всі створені файли.

## 3. Головне вікно
    Заголовок         назва та версія програми
    Кнопка  i         цей посібник (мову можна обрати вгорі панелі)
    Project folder    шлях до папки проєкту
    Browse...         вибір папки проєкту; програма запам'ятовує останню папку
    Create Omicron_configurations   — розділ 4
    Create GOOSE_configurations     — розділ 5
    Update GOOSE_configurations     — розділ 6
    Create XRIO_test_objects        — розділ 7
    Open Omicron folder             — розділ 8
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
    ConfRev = 1
Також заповнюється Dataset Mapping — який канал CMC іде в який атрибут набору даних.
Як і при ручному налаштуванні, з майстра лишаються: Control block reference, Dataset reference, рядок «Select IED / Sampled Values ID», підписи атрибутів (MU01.TCTR1.AmpSv…) і службові таблиці звіту. Потік без SV ID лишається таким, як у майстрі.
Кожен перейменований канал (BI, BO, аналоговий) отримує позначку «назву змінив користувач» — як це робить Test Universe.
Джерело часу (Time Source, PTP) береться з майстра: налаштуйте його в default_01.ohc і збережіть.
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
Кожен IED береться ЛИШЕ З ЙОГО ВЛАСНОГО ФАЙЛУ: CID / ICD / IID дає тільки свій (перший) IED. Інші IED усередині CID — це копії видавців, які пристрій зберігає для своїх підписок (без VLAN, з часами повторів за замовчуванням), вони не використовуються. SCD дає всі свої IED.
Перевірка CID: пристрій, для якого бракує власного CID будь-якого потрібного IED (або MAC у GEN), НЕ СТВОРЮЄТЬСЯ — програма показує повідомлення з причиною (англійською), решта пристроїв створюються, наприклад:
    W1G_D60: NOT CREATED - W1G_F60 / GCB BAY (MAC 01-0C-CD-01-00-55) - no own CID file of W1G_F60 in the CID folder; it is only a copy inside D60.cid
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

## 7. Create XRIO_test_objects (файли *.xrio)
Що робить: для кожного вибраного пристрою створює об'єкти випробувань Test Universe — по одному на кожен аркуш «Omicron AI <тест>», у якому є підписки на потоки SV (стовпець F — SV ID).
Порядок дій: як у розділі 4 — кнопка, список пристроїв з галочками, Create.
Що записується (вікно Test Universe «Device Settings»); усе інше в майстрі лишається без змін:
    Name/description       ← IEDName
    Manufacturer           ← Manufacturer типу пристрою (General_PB, Device_Type); не задано — порожньо
    Device type            ← Order code пристрою (General_PB, GEN), інакше order_code типу; не задано — порожньо
    Device address         ← IP Station Bus (GEN, IP SB), якщо немає — IP Process Bus (IP PB)
    Substation: Name       ← Project name   (General_PB Editor → Project)
    Substation: Address    ← Project address (General_PB Editor → Project)
    Bay: Name              ← BAY пристрою (Settings_IED)
Номінальні значення беруться з аркуша SV General_PB (стовпці N…W) для потоків, на які підписаний тест:
    f nom                  ← F, Hz
    V nom (secondary)      ← Vn secondary f-f, V
    V primary              ← Vn primary f-f, kV × 1000
    I nom (secondary)      ← In secondary, A
    I primary              ← In primary, A
    Residual voltage/current factors:
       VLN / VN            = (Vneutral primary × 1000 / Vneutral secondary) / (V primary / V nom)
       IN / I nom          = (I primary / I nom) / (Ineutral primary / Ineutral secondary)
       VN sec, IN sec, instrument transformer ← Vneutral secondary, Ineutral secondary, Instrumental transformer
У General_PB ці значення задаються в типі пристрою (Device_Type, блок SVCB<n>; за замовчуванням — значення test_object_default.xrio) і переносяться на аркуш SV кнопками Create / Update Gen/SV. Update зберігає значення, змінені вручну на аркуші SV.
Перевірка: усі потоки SV одного тесту мають однакові номінальні значення. Якщо значення різні або порожні, файл для цього тесту НЕ створюється, програма показує повідомлення, а причина записується у звіт, наприклад:
    W1G_D60 / Test 2: SV streams have different nominal values - test object cannot be created for this configuration: In secondary, A MU02_QW1G_01 1 <> MU01_QC1G_02 5
Результат:
    Omicron\\<IED>\\test_object_<IED>_AI_Test_<n>.xrio
    Omicron\\_xrio_report.txt                     — звіт
Пристрій без жодної підписки на SV (аркушів «Omicron AI» немає або в них немає SV ID): створюється ОДИН файл Omicron\\<IED>\\test_object_<IED>.xrio — заповнюється лише перша частина (паспорт і місце встановлення), номінальні значення лишаються з майстра. Якщо в пристрою є тести з підписками, аркуш без SV ID просто пропускається (примітка у звіті).
Якщо такий файл уже є (можливо, ви доповнили його в Test Universe уставками), він НЕ перезаписується: поруч створюється файл з датою і часом у назві (…_<дата>_<час>.xrio).

## 8. Open Omicron folder
Відкриває папку <Проєкт>\\Omicron у Провіднику. Якщо її ще немає, програма про це повідомить.

## 9. Типові повідомлення і що робити
    Project folder not found / General_PB.xlsm not found  → вкажіть папку, де лежать General_PB.xlsm і DOC
    No master .ohc file / Master file not found           → покладіть default_01.ohc або default_goose.xml у Library
    Master test object not found                          → покладіть test_object_default.xrio у Library
    SV streams have different nominal values / nominal values empty → виправте номінали на аркуші SV General_PB (розділ 7)
    Folder with the CID files not found                   → створіть папку CID у проєкті
    NOT CREATED - no own CID file of …                    → додайте власний CID названого IED (розділ 5)
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
    • update of existing GOOSE configurations from new CID files;
    • test objects (*.xrio) — nameplate, location and nominal values of the device.
The program only reads the Excel and CID files and never changes them.
Independent tool, not affiliated with or endorsed by OMICRON electronics. OMICRON and Test Universe are trademarks of OMICRON electronics GmbH.

## 2. Requirements
Start: the exe file IED2OMICRON (nothing to install) or from PyCharm / Python: ied2omicron.py. Next to it: omicron_configuration.py, omicron_goose_config.py, omicron_xrio.py and ied_doc_manual.py; the package openpyxl is needed.
Project folder:
    <Project>\\
      General_PB.xlsm          main workbook of the project
      DOC\\                     device files <IED>.xlsm
      Library\\                 master files:
         default_01.ohc        master of the hardware configuration
         default_goose.xml     master of the GOOSE configuration
         test_object_default.xrio   master of the test object
      CID\\                     CID files (*.cid, *.icd, *.scd, *.iid) of all devices taking part in GOOSE
      Omicron\\                 CREATED BY THE PROGRAM — results
         <IED>\\                one folder per device
The master files are empty configurations saved in Test Universe. The program does not build them from scratch, it fills them:
    • default_01.ohc — with the GOOSE virtual inputs/outputs enabled (7×24) and three SV streams (6 groups of 4 analog channels);
    • default_goose.xml — an empty configuration of the GOOSE module;
    • test_object_default.xrio — a test object saved in Test Universe (Device Settings).
General settings (CMC, PTP, IRIG-B, network port etc.) are set once in the master and go into every created file.

## 3. Main window
    Title             name and version of the program
    Button  i         this manual (choose the language at the top of the panel)
    Project folder    path of the project folder
    Browse...         choose the project folder; the last folder is remembered
    Create Omicron_configurations   — section 4
    Create GOOSE_configurations     — section 5
    Update GOOSE_configurations     — section 6
    Create XRIO_test_objects        — section 7
    Open Omicron folder             — section 8
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
    ConfRev = 1
The Dataset Mapping is filled too — which CMC channel goes into which attribute of the data set.
As when it is done by hand, these stay from the master: Control block reference, Dataset reference, the line "Select IED / Sampled Values ID", the attribute captions (MU01.TCTR1.AmpSv…) and the report tables. A stream without SV ID stays as in the master.
Every renamed channel (BI, BO, analog) gets the flag "name changed by the user" — as Test Universe does.
The time source (Time Source, PTP) comes from the master: set it in default_01.ohc and save it.
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
Every IED is taken ONLY FROM ITS OWN FILE: a CID / ICD / IID gives only its own (first) IED. The other IEDs inside a CID are copies of publishers kept by the device for its subscriptions (no VLAN, default repetition times); they are never used. An SCD gives all its IEDs.
CID check: a device that needs an IED without its own CID file (or without a MAC in GEN) is NOT CREATED - the program shows a message with the reason; the other devices are made, for example:
    W1G_D60: NOT CREATED - W1G_F60 / GCB BAY (MAC 01-0C-CD-01-00-55) - no own CID file of W1G_F60 in the CID folder; it is only a copy inside D60.cid
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

## 7. Create XRIO_test_objects (*.xrio files)
What it does: for every selected device it creates Test Universe test objects — one per sheet "Omicron AI <test>" that subscribes to SV streams (column F — SV ID).
Steps: as in section 4 — the button, the list of devices with check boxes, Create.
What is written (Test Universe dialog "Device Settings"); everything else of the master stays as it is:
    Name/description       ← IEDName
    Manufacturer           ← Manufacturer of the device type (General_PB, Device_Type); not set — empty
    Device type            ← Order code of the device (General_PB, GEN), else order_code of the type; not set — empty
    Device address         ← Station Bus IP (GEN, IP SB), if none — Process Bus IP (IP PB)
    Substation: Name       ← Project name   (General_PB Editor → Project)
    Substation: Address    ← Project address (General_PB Editor → Project)
    Bay: Name              ← BAY of the device (Settings_IED)
The nominal values come from the sheet SV of General_PB (columns N…W) of the streams the test subscribes to:
    f nom                  ← F, Hz
    V nom (secondary)      ← Vn secondary f-f, V
    V primary              ← Vn primary f-f, kV × 1000
    I nom (secondary)      ← In secondary, A
    I primary              ← In primary, A
    Residual voltage/current factors:
       VLN / VN            = (Vneutral primary × 1000 / Vneutral secondary) / (V primary / V nom)
       IN / I nom          = (I primary / I nom) / (Ineutral primary / Ineutral secondary)
       VN sec, IN sec, instrument transformer ← Vneutral secondary, Ineutral secondary, Instrumental transformer
In General_PB these values are set in the device type (Device_Type, block SVCB<n>; default — the values of test_object_default.xrio) and go to the sheet SV with Create / Update Gen/SV. Update keeps values changed by hand on the sheet SV.
Check: all SV streams of one test must have the same nominal values. If they differ or are empty, NO file is created for this test, the program shows a message and the reason goes into the report, for example:
    W1G_D60 / Test 2: SV streams have different nominal values - test object cannot be created for this configuration: In secondary, A MU02_QW1G_01 1 <> MU01_QC1G_02 5
Result:
    Omicron\\<IED>\\test_object_<IED>_AI_Test_<n>.xrio
    Omicron\\_xrio_report.txt                     — report
A device without any SV subscription (no "Omicron AI" sheets or no SV ID on them): ONE file Omicron\\<IED>\\test_object_<IED>.xrio — only the first part is filled (nameplate and location), the nominal values stay as in the master. If the device has tests with subscriptions, a sheet without SV ID is just skipped (a note in the report).
If the file already exists (you may have added relay settings to it in Test Universe), it is NOT overwritten: a file with the date and time in its name (…_<date>_<time>.xrio) is written next to it.

## 8. Open Omicron folder
Opens the folder <Project>\\Omicron in Explorer. If it does not exist yet, the program says so.

## 9. Typical messages and what to do
    Project folder not found / General_PB.xlsm not found  → choose the folder with General_PB.xlsm and DOC
    No master .ohc file / Master file not found           → put default_01.ohc or default_goose.xml into Library
    Master test object not found                          → put test_object_default.xrio into Library
    SV streams have different nominal values / nominal values empty → correct the nominal values on the sheet SV of General_PB (section 7)
    Folder with the CID files not found                   → create the folder CID in the project
    NOT CREATED - no own CID file of …                    → add the own CID file of the named IED (section 5)
    SV ID … not on General_PB SV                          → the SV ID of the test sheet is not in General_PB; the stream stays as in the master
    data type(s) … not supported yet                      → the data set has a type that is not supported yet (section 5)
    not read - …                                          → the file is damaged or cannot be read
A file open in Excel is read from a temporary copy — Excel does not have to be closed.

[[Ievgen German|%(author)s]]
""" % dict(src=URL_SOURCE, gpb=URL_GENERAL_PB, author=URL_AUTHOR)

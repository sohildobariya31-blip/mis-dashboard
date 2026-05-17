# 📊 Purchase MIS Dashboard — વાપરવાની રીત

## શરૂઆત (1 મિનિટ માં)

1. **`Purchase-MIS-Dashboard.html`** ફાઇલ કોઇ પણ બ્રાઉઝર (Chrome / Edge / Firefox) માં ખોલો.
2. **📂 Choose Excel file** બટન દબાવો, અથવા Excel ફાઇલ સીધી પેજ પર drag કરો.
3. જો ફાઇલ માં ઘણા sheet હોય તો **BOM** sheet પસંદ કરો. → બધું આપોઆપ ભરાઇ જશે.

ફાઇલ next time ખોલશો ત્યારે **છેલ્લે upload કરેલી ફાઇલ** આપોઆપ યાદ રહે છે. નવી Excel આવે ત્યારે ફરી upload કરી લો.

---

## મુખ્ય Features

| જે જોઇએ તે | કેવી રીતે |
|---|---|
| **Filter** (Vendor / Item / Category / Status…) | નીચે ની chip પર click — ફરી click કરો તો unselect |
| **Date range** | Today / 7d / 30d / 90d / YTD / All ની chip, અથવા From-To પસંદ કરો |
| **Chart customize** | દરેક chart ઉપર Group / Metric / Top N / Type (Donut / Pie / Bar / Polar) ની dropdown |
| **Labels (chart પર value)** | Auto / All / Top 10 / Off — congested લાગે તો Top 10 કે Off રાખો |
| **Theme બદલવો** | Header માં 🎨 Theme — 7 themes (Midnight default શ્રેષ્ઠ) |
| **Table column choose** | ⚙ Columns બટન — checkbox + drag to reorder |
| **Search row** | Table ઉપર search box, અથવા દરેક column ની filter row |
| **Export** | 📤 Export → Excel / CSV / PDF / PowerPoint |

---

## 💱 Currency (INR સિવાય ની currency હોય તો)

- ઉપર પીળી/લાલ flowing **caution banner** આવશે.
- **Set rates** click → દરેક currency × month માટે `1 USD = ? INR` rate type કરો.
- **Convert non-INR to INR** toggle ON કરો → બધા charts/KPI માં INR માં convert થઇ ને દેખાશે.
- Rate ન જોઇતું હોય તો **Show INR only** click → non-INR rows hide થઇ જશે.
- Rate yaad rakhe છે — ફરીથી type કરવાની જરૂર નહીં.

---

## Reset / થી શરૂ કરવું

- Header માં **Reset** બટન — બધા filter, date, columns reset થાય.
- ↻ બટન (Upload Excel ની બાજુમાં) — છેલ્લી ફાઇલ ફરી load.

---

## ⚠️ ધ્યાન માં રાખવાનું

- Excel ની column headers same હોવી જોઇએ (Category, PO no, PO Date, Supplier, Item Code, Qty, UOM, Status, Payable, Paid, Outstanding…).
- File browser માં જ open થાય છે — **server પર upload નથી થતી**, data safe છે.
- પહેલી વાર open કરો ત્યારે internet જોઇએ (charts/export libraries માટે). પછી offline ચાલે.

---

## 🆘 Problem આવે તો

- Charts blank દેખાય → date range "All" પર સેટ કરો, અથવા Reset દબાવો.
- Currency banner ન જાય → 💱 Currency માં rates ભરો અથવા "Show INR only" click કરો.
- Excel load ન થાય → BOM sheet ની column names ચેક કરો.

કોઇ વધારાની મદદ માટે **MIS team** ને contact કરો.

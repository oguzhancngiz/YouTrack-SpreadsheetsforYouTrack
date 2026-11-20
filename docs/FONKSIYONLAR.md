# Spreadsheets for YouTrack – fonksiyon listesi

Toplam **314** fonksiyon adı (292 ayrı fonksiyon + 22 eski/alternatif ad, ör. STDEV = STDEV.S). Koddan otomatik çıkarıldı.

## Matematik (78)

ABS, ACOS, ACOSH, AGGREGATE, ARABIC, ASIN, ASINH, ATAN, ATAN2, ATANH, BASE, BIN2DEC, CEILING, CEILING.MATH, CEILING.PRECISE, COMBIN, COMBINA, COS, COSH, COT, CSC, DEC2BIN, DEC2HEX, DEC2OCT, DECIMAL, DEGREES, EVEN, EXP, FACT, FACTDOUBLE, FLOOR, FLOOR.MATH, FLOOR.PRECISE, GCD, HEX2DEC, INT, LCM, LN, LOG, LOG10, MDETERM, MINVERSE, MMULT, MOD, MROUND, MUNIT, OCT2DEC, ODD, PERMUT, PI, POWER, PRODUCT, QUOTIENT, RADIANS, RAND, RANDARRAY, RANDBETWEEN, ROMAN, ROUND, ROUNDDOWN, ROUNDUP, SEC, SEQUENCE, SIGN, SIN, SINH, SQRT, SQRTPI, SUBTOTAL, SUM, SUMPRODUCT, SUMSQ, SUMX2MY2, SUMX2PY2, SUMXMY2, TAN, TANH, TRUNC

## İstatistik (88)

AVEDEV, AVERAGE, AVERAGEA, AVERAGEIF, AVERAGEIFS, BINOM.DIST, BINOMDIST, CONFIDENCE, CONFIDENCE.NORM, CORREL, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, COVAR, COVARIANCE.P, COVARIANCE.S, DEVSQ, EXPON.DIST, EXPONDIST, FISHER, FISHERINV, FORECAST, FORECAST.LINEAR, FREQUENCY, GAMMA, GAMMALN, GAMMALN.PRECISE, GEOMEAN, HARMEAN, INTERCEPT, KURT, LARGE, MAX, MAXA, MAXIFS, MEDIAN, MIN, MINA, MINIFS, MODE, MODE.MULT, MODE.SNGL, NORM.DIST, NORM.INV, NORM.S.DIST, NORM.S.INV, NORMDIST, NORMINV, NORMSDIST, NORMSINV, PEARSON, PERCENTILE, PERCENTILE.EXC, PERCENTILE.INC, PERCENTRANK, PERCENTRANK.INC, POISSON, POISSON.DIST, PROB, QUARTILE, QUARTILE.EXC, QUARTILE.INC, RANK, RANK.AVG, RANK.EQ, RSQ, SKEW, SLOPE, SMALL, STANDARDIZE, STDEV, STDEV.P, STDEV.S, STDEVA, STDEVP, STDEVPA, STEYX, SUMIF, SUMIFS, TREND, VAR, VAR.P, VAR.S, VARA, VARP, VARPA

## Finans (22)

CUMIPMT, CUMPRINC, DB, DDB, EFFECT, FV, IPMT, IRR, MIRR, NOMINAL, NPER, NPV, PDURATION, PMT, PPMT, PV, RATE, RRI, SLN, SYD, XIRR, XNPV

## Mantıksal (13)

AND, CHOOSE, FALSE, IF, IFERROR, IFNA, IFS, LET, NOT, OR, SWITCH, TRUE, XOR

## Metin (37)

ARRAYTOTEXT, CHAR, CLEAN, CODE, CONCAT, CONCATENATE, DOLLAR, EXACT, FIND, FIXED, LEFT, LEN, LOWER, MID, N, NUMBERVALUE, PROPER, REGEXEXTRACT, REGEXREPLACE, REGEXTEST, REPLACE, REPT, RIGHT, SEARCH, SUBSTITUTE, T, TEXT, TEXTAFTER, TEXTBEFORE, TEXTJOIN, TEXTSPLIT, TRIM, UNICHAR, UNICODE, UPPER, VALUE, VALUETOTEXT

## Tarih ve saat (25)

DATE, DATEDIF, DATEVALUE, DAY, DAYS, DAYS360, EDATE, EOMONTH, HOUR, ISOWEEKNUM, MINUTE, MONTH, NETWORKDAYS, NETWORKDAYS.INTL, NOW, SECOND, TIME, TIMEVALUE, TODAY, WEEKDAY, WEEKNUM, WORKDAY, WORKDAY.INTL, YEAR, YEARFRAC

## Arama ve başvuru (22)

ADDRESS, AREAS, CELL, COLUMN, COLUMNS, FORMULATEXT, HLOOKUP, HYPERLINK, INDEX, INDIRECT, ISFORMULA, ISREF, LOOKUP, MATCH, OFFSET, ROW, ROWS, SHEET, SHEETS, VLOOKUP, XLOOKUP, XMATCH

## Dinamik diziler (taşan sonuçlar) (16)

CHOOSECOLS, CHOOSEROWS, DROP, EXPAND, FILTER, HSTACK, SORT, SORTBY, TAKE, TOCOL, TOROW, TRANSPOSE, UNIQUE, VSTACK, WRAPCOLS, WRAPROWS

## Bilgi (13)

ERROR.TYPE, ISBLANK, ISERR, ISERROR, ISEVEN, ISLOGICAL, ISNA, ISNONTEXT, ISNUMBER, ISODD, ISTEXT, NA, TYPE

## What-If analizi (Data menüsü)

Bunlar formül fonksiyonu değil, menüden açılan araçlardır.

**Solver** (Data > Solver…)

- **Hedef:** bir hücreyi en büyük (Max) veya en küçük (Min) yapmak, ya da belirli bir değere getirmek (Value Of). Hedef boş bırakılırsa yalnızca kısıtları sağlayan bir çözüm aranır.
- **Değişken hücreler:** virgülle ayrılmış birden çok aralık veya adlandırılmış aralık, en fazla 500 hücre. Hücre seçme düğmesiyle aralık sayfadan seçilebilir.
- **Kısıtlar:** `<=`, `=`, `>=`, `int` (tamsayı), `bin` (0/1) ve `dif` (AllDifferent: 1..N'nin bir permütasyonu). Sağ taraf sayı, hücre, aynı boyutta aralık veya formül olabilir. En fazla 2.000 kısıt hücresi. Add, Change, Delete, Reset All.
- **Make Unconstrained Variables Non-Negative:** alt sınırı olmayan değişkenler 0 veya üzeri olur.
- **Yöntemler:**
  - Simplex LP: iki fazlı simplex; tamsayılar için dal-sınır (branch and bound) kullanır.
  - GRG Nonlinear: artırılmış Lagrange ve BFGS; ileri veya merkezi türev; Multistart seçeneği.
  - Evolutionary: genetik algoritma ve yerel arama.
- **Options (3 sekme):**
  - All Methods: Constraint Precision, Use Automatic Scaling, Show Iteration Results, Ignore Integer Constraints, Integer Optimality (%), Max Time, Iterations, Max Subproblems, Max Feasible Solutions.
  - GRG Nonlinear: Convergence, Derivatives (Forward/Central), Use Multistart, Population Size, Random Seed, Require Bounds on Variables.
  - Evolutionary: Convergence, Mutation Rate, Population Size, Random Seed, Maximum Time without improvement, Require Bounds on Variables.
  - Excel'de olduğu gibi popülasyon, tohum, yakınsama ve sınır şartı iki sekmede ortaktır.
- **Çözüm sırasında:**
  - Stop düğmesi.
  - Show Iteration Results ile her iterasyonda duraklama (Continue, Stop, Save Scenario).
  - Süre veya iterasyon sınırında "devam edilsin mi?" sorusu.
- **Solver Results:** Keep Solver Solution veya Restore Original Values, Return to Solver Parameters Dialog, Save Scenario…
- **Raporlar** (yeni sayfa olarak eklenir):
  - Answer.
  - Sensitivity: Simplex LP'de gölge fiyat, indirgenmiş maliyet ve izin verilen artış/azalış; GRG'de Lagrange çarpanı ve indirgenmiş gradyan.
  - Limits.
  - Population (Evolutionary).
  - Feasibility ve Feasibility-Bounds: çözümsüz problemde birbiriyle çakışan en küçük kısıt kümesi.
  - Linearity: model Simplex LP için doğrusal değilse.
- **Load/Save:** model hücrelere kaydedilir ve hücrelerden yüklenir (Replace veya Merge).
- **Saklama:**
  - Model sayfayla birlikte saklanır ve issue'yu açan herkes görür.
  - Satır/sütun ekleme-silme ve sayfa adı değişikliklerinde referanslar güncellenir.
  - .xlsx dosyalarına Excel'in gizli `solver_*` adlarıyla yazılır, oradan okunur.

**Goal Seek** (Data > What-if analysis > Goal Seek…)

- Set cell, To value, By changing cell.
- Sonuç penceresinde Target value ve Current value görünür. OK değeri uygular, Cancel geri alır.

**Scenario Manager** (Data > What-if analysis > Scenario Manager…)

- Add, Edit, Delete, Show, Merge (başka sayfadan) ve Summary (sonuç hücreleriyle özet sayfası).
- Bir senaryoda en fazla 32 değişen hücre olur.
- Senaryolar .xlsx dosyasında saklanır.

## Olmayan fonksiyonlar

- LAMBDA ve yardımcıları: MAP, BYROW, BYCOL, REDUCE, SCAN, MAKEARRAY
- Veritabanı fonksiyonları: DSUM, DCOUNT, DAVERAGE, DGET vb.
- Küp (CUBE...) ve web fonksiyonları: WEBSERVICE, FILTERXML, IMAGE
- Mühendislik fonksiyonlarının çoğu (CONVERT, BESSEL, karmaşık sayılar IM...). Taban dönüşümleri var: BIN2DEC, DEC2BIN, DEC2HEX, DEC2OCT, HEX2DEC, OCT2DEC
- Bazı ileri istatistik testleri: T.TEST, CHISQ.TEST, F.TEST, T.DIST, CHISQ.DIST, LINEST, LOGEST, GROWTH
- Tarih/finans ekleri: PRICE, YIELD, ACCRINT, COUPDAYS gibi tahvil fonksiyonları
- GETPIVOTDATA (pivot tablo yok), INFO, ENCODEURL, STOCKHISTORY, TRANSLATE, DETECTLANGUAGE
- YouTrack issue alanlarını okuyan formüller (formüller yalnızca tablo hücreleriyle çalışır)
- What-If menüsünde olmayanlar: Data Table (veri tablosu, TABLE()), Forecast Sheet, senaryo PivotTable raporu ve Solver raporlarında "Outline Reports" (satır gruplama yok)

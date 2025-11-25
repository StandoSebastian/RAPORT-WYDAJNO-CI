# RAPORT-WYDAJNO-CI

## Raport Wydajności SAP

System raportowania wydajności produkcji dla SAP ERP. Program ABAP pozwala na:
- Import danych z plików TSV (Tab-Separated Values)
- Przetwarzanie i agregację danych produkcyjnych
- Wyświetlanie danych w formacie ALV Grid
- Generowanie podglądów HTML
- Wysyłanie raportów e-mailem (HTML w treści lub XLSX jako załącznik)

---

## Funkcjonalności

### Obsługiwane wydziały produkcyjne (WYDZIAL)
Program automatycznie rozpoznaje wydział na podstawie nazwy pliku TSV:
- **KOSTKA** - Wydział produkcji kostek
- **AEROZOLE** - Wydział aerozoli
- **KONFEKCJA** - Wydział konfekcji
- **WTRYSK** - Wtryskownia
- **BLACHARNIA** - Blacharnia

### Agregacja danych
Raport grupuje dane według:
- Data (`DATA_D`)
- Zmiana (`ZMIANA`)
- Linia produkcyjna (`LINIA_PLIK`)
- Numer zlecenia (`NR_ZLECENIA`)

### Obliczane wskaźniki
- **Wydajność ważona (%)** - procentowa wydajność produkcji
- **Ilość wyprodukowana** - suma sztuk wyprodukowanych
- **Ilość oczekiwana (Plan)** - planowana ilość na podstawie normy
- **Średnia osób rzeczywistych** - średnia ważona liczby osób na produkcji
- **Średnia osób normatywnych** - normatywna liczba osób

### Czasy rejestrowane
- Czas trwania zmiany (min)
- Czas przejścia (min)
- Czas przezbrojenia (min)
- Czas przerwy (min)
- Czas awarii (min)
- Czas organizacyjny (min)
- Czas efektywny pracy (min)

---

## Tabele SAP

### ZSTLINIENORMA
Tabela przechowująca normy produkcyjne dla linii:

| Pole | Typ | Opis |
|------|-----|------|
| CLIENT | CLNT(3) | Mandant SAP |
| WYDZIAL | CHAR(10) | Kod wydziału |
| LINIA | CHAR(30) | Nazwa linii produkcyjnej |
| NORMA | INT4 | Norma produkcyjna (szt/8h) |

### ZSTLINIESORT
Tabela kolejności wyświetlania linii w raporcie:

| Pole | Typ | Opis |
|------|-----|------|
| CLIENT | CLNT(3) | Mandant SAP |
| WYDZIAL | CHAR(10) | Kod wydziału |
| LINIA | CHAR(30) | Nazwa linii produkcyjnej |
| SORT | INT4 | Priorytet sortowania |
| SORT2 | INT4 | Dodatkowy priorytet sortowania |

---

## Parametry programu

### Plik wejściowy
- `P_FILE` - Ścieżka do pliku TSV z danymi

### Filtrowanie
- `P_ONLYOK` - Przetwarzaj tylko poprawne rekordy (domyślnie: TAK)
- `P_ERRLOG` - Pokaż log błędnych rekordów
- `P_FIXHH` - Napraw format godzin (HH:MM)

### Test mailowy
- `P_TMAIL` - Tryb testowy wysyłki maila
- `P_TSUBJ` - Temat testowego maila
- `P_TBODY` - Treść testowego maila

### Wysyłka HTML
- `P_HTML` - Wyślij raport jako HTML w treści maila
- `P_HSUB` - Temat maila HTML
- `P_HPRE` - Tekst przed tabelą HTML

### Wysyłka XLSX
- `P_XLS` - Wyślij raport jako załącznik XLSX
- `P_XSUB` - Temat maila XLSX
- `P_XMSG` - Treść maila XLSX
- `P_XFN` - Nazwa pliku załącznika

### Podgląd
- `P_PREV` - Szybki podgląd HTML (bez wysyłki maila)

### Adres e-mail
- `P_EMAIL` - Adres e-mail odbiorcy raportu

---

## Kolorowanie wydajności

Raport używa kolorów do szybkiej identyfikacji wydajności:
- 🔴 **Czerwony** - Wydajność < 85%
- 🟡 **Żółty** - Wydajność 85-95%
- 🟢 **Zielony** - Wydajność ≥ 95%

Kolorowanie czasu trwania:
- 🟢 **Zielony** - Pełna zmiana (480 min)
- 🟠 **Pomarańczowy** - Niepełna zmiana

---

## Struktura plików

```
RAPORT-WYDAJNO-CI/
├── README.md                    # Ta dokumentacja
├── REPORT zst_excel_agg..txt    # Kod źródłowy ABAP
├── ZSTLINIENORMA_*.txt          # Dokumentacja tabeli ZSTLINIENORMA
└── ZSTLINIESORT_*.txt           # Dokumentacja tabeli ZSTLINIESORT
```

---

## Autor
Raport Wydajności SAP dla produkcji Dramers

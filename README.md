# RAPORT-WYDAJNO-CI

## Raport Wydajności SAP

Program ABAP do generowania raportów wydajności dla wydziałów produkcyjnych w systemie SAP.

### Funkcjonalności

- **Załadowanie danych TSV** - Import danych z plików TSV (Tab-Separated Values)
- **Wyświetlanie ALV** - Przegląd danych w standardowym widoku ALV z pełną funkcjonalnością
- **Podgląd HTML** - Generowanie podglądu raportu w formacie HTML
- **Wysyłka email** - Wysyłanie raportu emailem z załącznikiem HTML lub XLSX

### Struktura projektu

```
├── src/
│   └── zrep_wydajnosc_sap.abap    # Główny program ABAP
├── sample_data/
│   └── wydajnosc_sample.tsv       # Przykładowe dane testowe
└── README.md                       # Dokumentacja
```

### Wymagany format pliku TSV

Plik TSV musi zawierać następujące kolumny (oddzielone tabulacją):

| Kolumna | Opis | Typ |
|---------|------|-----|
| Wydzial | Nazwa wydziału produkcyjnego | Tekst (20 znaków) |
| Pracownik | Imię i nazwisko pracownika | Tekst (50 znaków) |
| Nr_pracownika | Numer identyfikacyjny pracownika | Tekst (10 znaków) |
| Data | Data pracy (format: DD.MM.YYYY) | Data |
| Czas_pracy | Czas pracy w godzinach | Liczba dziesiętna |
| Ilosc_sztuk | Ilość wyprodukowanych sztuk | Liczba całkowita |
| Norma | Norma produkcyjna | Liczba całkowita |
| Uwagi | Uwagi dodatkowe (opcjonalne) | Tekst (100 znaków) |

### Obliczanie wydajności

Program automatycznie oblicza wydajność procentową według wzoru:

```
Wydajność % = (Ilość sztuk / Norma) × 100
```

**Statusy wydajności:**
- **POWYŻEJ** - wydajność >= 100%
- **OK** - wydajność >= 80% i < 100%
- **PONIŻEJ** - wydajność < 80%

### Instrukcja użytkowania

1. **Uruchom raport** w transakcji SE38 lub własnej transakcji
2. **Wybierz plik TSV** używając przycisku F4 przy polu "Plik"
3. **Podaj nazwę wydziału** (domyślnie: PRODUKCJA)
4. **Wybierz format załącznika** dla emaila (HTML lub XLSX)
5. **Wykonaj raport** (F8)
6. W widoku ALV użyj przycisków:
   - **Podgląd HTML** - otwiera raport w przeglądarce
   - **Wyślij email** - wysyła raport na podany adres

### Instalacja

1. Skopiuj zawartość pliku `src/zrep_wydajnosc_sap.abap` do nowego programu w SE38
2. Aktywuj program
3. Opcjonalnie: utwórz transakcję dla programu

### Wymagania

- SAP NetWeaver 7.40 lub nowszy
- Uprawnienia do wysyłania emaili (SCOT musi być skonfigurowany)
- Dostęp do GUI frontend dla operacji na plikach

### Autor

Raport wydajności SAP - wersja 1.0

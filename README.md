# RAPORT-WYDAJNO-CI

## Raport Wydajności SAP

Program ABAP do generowania raportów wydajności dla wydziałów produkcyjnych.

## Funkcjonalności

- **Wczytywanie danych z plików TSV** - Import danych produkcyjnych z plików TSV
- **Wyświetlanie w ALV** - Pełna funkcjonalność ALV Grid z sortowaniem, filtrowaniem i eksportem
- **Podgląd HTM** - Generowanie i wyświetlanie raportu w formacie HTML
- **Wysyłanie maili** - Możliwość wysłania raportu mailem w formacie HTM lub XLSX

## Struktura projektu

```
├── src/
│   └── zraport_wydajnosci.prog.abap   # Główny program ABAP
├── data/
│   └── sample_wydajnosc.tsv           # Przykładowe dane TSV
└── README.md
```

## Format pliku TSV

Plik TSV powinien zawierać następujące kolumny (oddzielone tabulatorem):

| Kolumna | Opis |
|---------|------|
| Wydzial | Nazwa wydziału produkcyjnego |
| Data | Data w formacie YYYYMMDD |
| Pracownik | Imię i nazwisko pracownika |
| Produkt | Nazwa produktu |
| Ilosc | Ilość wyprodukowana |
| Czas_pracy | Czas pracy w godzinach |
| Cel | Cel wydajności (szt/godz) |
| Uwagi | Dodatkowe uwagi |

## Użycie

1. Uruchom transakcję SE38 i wykonaj program `ZRAPORT_WYDAJNOSCI`
2. Wybierz plik TSV z danymi produkcyjnymi
3. Podaj nazwę wydziału do filtrowania
4. Wybierz format eksportu (HTM lub XLSX)
5. Kliknij "Wykonaj" aby wyświetlić raport w ALV

### Funkcje dostępne w ALV:

- **Podgląd HTM** - Generuje i otwiera raport w przeglądarce
- **Wyślij HTM** - Wysyła raport mailem w formacie HTML
- **Wyślij XLSX** - Wysyła raport mailem w formacie Excel

## Obliczane wartości

Program automatycznie oblicza:
- **Wydajność** = Ilość / Czas pracy (szt/godz)
- **% Celu** = (Wydajność / Cel) × 100

## Wymagania

- SAP NetWeaver 7.40 lub nowszy
- Uprawnienia do wysyłania maili (SAPconnect)
- Dostęp do GUI frontend services

## Autor

Raport wydajności SAP

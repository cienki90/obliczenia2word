# obliczenia2word

Program wstawia tabele obliczeń z pliku Excel do opisów w Wordzie. Tabela trafia do
rozdziału **„V. Zestawienie zbiorcze słupów”**, tak jak w `112.docx` i `115.docx`.

## Użycie: plik `Obliczenia2Word.exe`

Program nie wymaga instalacji ani Pythona. Działa na Windows 7, 8, 10 i 11, w wersji 32- i 64-bitowej.

1. Kliknij dwukrotnie `Obliczenia2Word.exe`.
2. **Krok 1:** wskaż plik Excel z obliczeniami.
3. **Krok 2:** wskaż folder z opisami (`.docx`).
4. **Krok 3:** wskaż folder, w którym zapisać gotowe opisy.
5. Na koniec program pokazuje raport, w którym:
   - **czerwone komórki** w tabelach z obliczeniami są wypisane z numerem arkusza, adresem komórki i treścią,
   - widać, które opisy wypełniono, a które pominięto (brak pasującego arkusza),
   - pojawi się ostrzeżenie, gdy Excel nie zapisał wyników formuł.

   Raport można zapisać do pliku `.txt`.

Jeśli Windows SmartScreen wyświetli ostrzeżenie, kliknij „Więcej informacji”, a potem „Uruchom mimo to”.
Program nie jest podpisany cyfrowo, dlatego pojawia się to ostrzeżenie.

## Co robi program

- Do opisu `181.docx` wstawia tabelę z arkusza `181`, do `191.docx` z arkusza `191` itd.
- Odwzorowuje formatowanie z Excela: scalenia, obramowania, czcionki, pogrubienia, kolory,
  wyrównanie, wysokości wierszy i format liczb.
- Jeśli opis ma już tabelę, wymienia ją na aktualną.
- **Pilnuje, żeby za stroną z obliczeniami nie powstała pusta strona.** Puste akapity i ręczny podział
  strony za tabelą zastępuje ustawieniem „podział strony przed” przy następnym rozdziale
  (OŚWIADCZENIE). Taki podział nie tworzy pustej strony, niezależnie od długości tabeli.
- Nagłówek obwodu nie zostaje sam na dole strony.
- Jeśli jako folder wynikowy wskażesz folder z opisami, program zapyta, czy nadpisać oryginały.

Przed uruchomieniem zapisz plik Excel, żeby zawierał aktualne wyniki formuł.

## Budowa exe

Najnowszy plik exe: https://github.com/cienki90/obliczenia2word/releases/tag/najnowsza-wersja

Plik exe buduje się automatycznie w GitHub Actions (`.github/workflows/build-exe.yml`)
na Windows. Workflow testuje go też na przykładowych plikach. Ręcznie (na Windows):

```
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name Obliczenia2Word --collect-data docx obliczenia2word.py
```

## Wiersz poleceń (opcjonalnie)

```
python obliczenia2word.py                                   # okienka
python obliczenia2word.py -x "Obliczenia.xlsx" -o wynik folder_z_opisami
python obliczenia2word.py --arkusz 181 projekt.docx         # gdy nazwa pliku ≠ nazwa arkusza
python obliczenia2word.py --nadpisz 181.docx                # zapis w oryginalnym pliku
python obliczenia2word.py ... --raport raport.txt           # zapis raportu do pliku
```

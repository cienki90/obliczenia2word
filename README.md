# obliczenia2word

Program automatycznie wstawia tabele obliczeń z pliku Excel do projektów Word,
w rozdziale **„V. Zestawienie zbiorcze słupów”** (tak jak w `112.docx` i `115.docx`).

- Plik `181.docx` dostaje tabelę z arkusza `181`, plik `191.docx` z arkusza `191` itd.
- Wstawiany jest zakres wydruku arkusza z formatowaniem z Excela: scalenia, obramowania,
  czcionki (Lato / Lato Light), pogrubienia, podkreślenia, kolory, wyrównanie,
  wysokości wierszy i format liczb (np. `130.62`).
- Tabela zajmuje całą szerokość strony. Kolumny są dopasowane do treści, a nagłówki
  obwodów nie zostają same na dole strony.
- Jeśli dokument ma już tabelę, program ją podmienia na aktualną.
- Oryginały zostają bez zmian, a wyniki trafiają do folderu `wynik`.

## Instalacja (jednorazowo)

1. Zainstaluj [Python 3](https://www.python.org/downloads/) i zaznacz „Add Python to PATH”.
2. W folderze programu uruchom: `pip install -r requirements.txt`

## Użycie

Wrzuć pliki `.docx` i plik `.xlsx` do jednego folderu razem z programem i kliknij dwukrotnie
`uruchom.bat`. Możesz też uruchomić go z wiersza poleceń:

```
python obliczenia2word.py                         # wszystkie .docx w folderze
python obliczenia2word.py 181.docx 191.docx       # wybrane pliki
python obliczenia2word.py -x "Obliczenia - zestawienie pw.xlsx" -o gotowe
python obliczenia2word.py --arkusz 181 projekt.docx   # gdy nazwa pliku ≠ nazwa arkusza
python obliczenia2word.py --nadpisz 181.docx          # zapis w oryginalnym pliku
```

Uwaga: program bierze wyniki formuł zapisane przez Excela. Przed uruchomieniem zapisz
plik Excel, żeby zawierał aktualne wartości.

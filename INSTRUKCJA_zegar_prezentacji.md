# Zegar czasu w PowerPoint (prawy górny róg) — instrukcja instalacji

## Co to robi
Podczas pokazu slajdów (F5) w prawym górnym rogu pojawia się licznik czasu:
- **zielony** — jesteś w harmonogramie
- **pomarańczowy** — zostało mniej niż 30 s do końca budżetu danego segmentu slajdów
- **czerwony z "+"** — przekroczyłeś budżet czasowy (pokazuje o ile)

Licznik **resetuje się automatycznie**, gdy wchodzisz w nowy zakres slajdów, i odlicza
czas przeznaczony wyłącznie na ten segment — dzięki temu prowadzący od razu widzi,
ile ma czasu na swoją część, niezależnie od tego, ile już minęło wcześniej. Np.:
- slajdy 1–5 → licznik startuje od **5:00** i odlicza w dół
- slajdy 6–10 → w momencie wejścia na slajd 6 licznik resetuje się i startuje od **15:00**

## Krok 1 — włącz kartę Deweloper
Plik → Opcje → Dostosowywanie Wstążki → zaznacz "Deweloper" → OK.

## Krok 2 — otwórz edytor VBA
Karta **Deweloper** → **Visual Basic** (albo Alt+F11).

## Krok 3 — zaimportuj moduł
W edytorze VBA: **Plik → Importuj plik…** → wybierz plik `PrezentacjaZegar.bas` (dołączony obok tej instrukcji).

Pojawi się moduł `ModZegar`. (Drugi plik, `ClsPresEvents.cls`, importujemy w Kroku 5).

## Krok 4 — dopasuj segmenty czasowe
W module `ModZegar`, w procedurze `GetSegmentInfo`, znajdziesz tablicę:

```vba
ranges = Array( _
    Array(1, 5, 5), _
    Array(6, 10, 15), _
    Array(11, 999, 30) _
)
```

Każdy wiersz to: **(slajd od, slajd do, czas trwania TEGO segmentu w minutach)**.
Licznik resetuje się do tej wartości za każdym razem, gdy wchodzisz w nowy zakres.
Dopisz/zmień wiersze pod swoją agendę, np.:

```vba
ranges = Array( _
    Array(1, 3, 3), _
    Array(4, 8, 12), _
    Array(9, 15, 25), _
    Array(16, 999, 30) _
)
```

## Krok 5 — zaimportuj moduł klasy zdarzeń
**Uwaga:** PowerPoint (w przeciwieństwie do Excela/Worda) nie ma wbudowanego `ThisPresentation`
— to normalne, że go nie widzisz. Zdarzenia pokazu slajdów przechwytujemy przez moduł klasy
podpięty pod obiekt `Application`.

W edytorze VBA: **Plik → Importuj plik…** → wybierz `ClsPresEvents.cls` (dołączony obok tej
instrukcji). Pojawi się jako klasa `KlasaZdarzen` w sekcji "Class Modules".

## Krok 6 — podepnij zdarzenia (RĘCZNIE, raz na sesję)
Ważne: `Auto_Open` **nie działa w zwykłych plikach .pptm** (działa tylko w dodatkach .ppam),
więc podpięcie trzeba odpalić ręcznie po każdym otwarciu pliku — ale tylko raz, przed
rozpoczęciem pokazu:

1. W PowerPoincie wciśnij **Alt+F8** (Windows) lub użyj **Narzędzia → Makra** (Mac).
2. Wybierz makro **`InitEvents`** → **Uruchom**.
3. Pojawi się okienko potwierdzające "Zegar podpięty" — od teraz możesz normalnie
   uruchomić pokaz slajdów (**F5**), a zegar wystartuje i będzie się aktualizował
   automatycznie przy każdej zmianie slajdu, bez dodatkowych kliknięć.

Musisz powtórzyć ten krok (Alt+F8 → InitEvents → Uruchom) **za każdym razem, gdy zamkniesz
i ponownie otworzysz plik** — to jednorazowa czynność na sesję pracy, nie na każdy pokaz.

## Krok 7 — zapisz jako .pptm
Plik → Zapisz jako → wybierz typ **Prezentacja programu PowerPoint z obsługą makr (*.pptm)**.
(Zwykły .pptx nie przechowa makra).

## Krok 8 — pierwsze uruchomienie
Przy pierwszym otwarciu pliku Windows/Office może zablokować makra — kliknij
**"Włącz zawartość"** na żółtym pasku ostrzeżeń. Potem wykonaj Krok 6 (Alt+F8 →
InitEvents → Uruchom) i dopiero uruchom pokaz slajdów (F5) — licznik pojawi się
automatycznie w prawym górnym rogu i będzie się aktualizował przy każdej zmianie slajdu.

## Uwagi praktyczne
- Zegar tworzy niewidoczne dla widowni pole tekstowe "ClockBox" tylko na tym slajdzie, który akurat jest wyświetlany podczas pokazu — nie zaśmieca pozostałych slajdów w widoku edycji.
- Jeśli chcesz mieć zegar widoczny **tylko na Twoim ekranie prezentera**, a nie dla widowni, to VBA tego nie rozdzieli (rysuje na tym samym slajdzie, który widzi każdy). W takim wypadku lepszym rozwiązaniem byłby osobny "Widok prezentera" PowerPoint + osobne narzędzie w przeglądarce na Twoim monitorze — daj znać, jeśli wolisz taką wersję zamiast makra.
- Zegar resetuje się tylko przy wejściu w **inny** segment (inny zakres slajdów) niż ten, w którym aktualnie odliczał czas. Poruszanie się między slajdami w obrębie tego samego segmentu (np. z 2 na 4 w zakresie 1–5) nie resetuje licznika — dalej odlicza czas dla tego segmentu.
- Jeśli cofniesz się do poprzedniego segmentu, licznik też się zresetuje i zacznie odliczać od nowa pełny czas tego segmentu (nie pamięta, ile już z niego zużyłeś).

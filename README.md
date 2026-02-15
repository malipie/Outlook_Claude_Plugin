# AI Mail Assistant — Outlook Add-in

Add-in do Outlooka, który generuje odpowiedzi na maile przy użyciu AI
(Claude, GPT, Gemini) w Twoim osobistym stylu pisania.

## Wymagania
- Node.js 18+
- Outlook na Macu lub Outlook Web
- Klucz API od co najmniej jednego providera:
  - [Anthropic Claude](https://console.anthropic.com/settings/keys)
  - [OpenAI GPT](https://platform.openai.com/api-keys)
  - [Google Gemini](https://aistudio.google.com/apikey)
- Konto GitHub (do hostingu produkcyjnego)

## Szybki start (dev, localhost)

1. Sklonuj repo i zainstaluj zależności:
   ```
   git clone https://github.com/TWOJ_USERNAME/ai-mail-assistant.git
   cd ai-mail-assistant
   npm install
   ```

2. Uruchom dev server:
   ```
   npm run dev
   ```

3. Zainstaluj add-in w Outlooku:
   - Outlook Mac: Narzędzia → Pobierz dodatki → Moje dodatki →
     Dodaj niestandardowy → Dodaj z pliku → wskaż manifest.xml
   - Outlook Web: Ustawienia → Zarządzaj dodatkami →
     Niestandardowe dodatki → Dodaj z pliku → wskaż manifest.xml

4. Otwórz dowolny mail → kliknij "AI Mail Assistant" na wstążce
5. Wybierz providera i wklej klucz API w ustawieniach
6. Gotowe!

## Produkcja (GitHub Pages — zero utrzymania)

Gdy add-in działa poprawnie na localhost, wrzuć go na GitHub Pages:

### Jednorazowa konfiguracja:

1. Stwórz repo na GitHubie: `ai-mail-assistant`

2. Włącz GitHub Pages:
   Repo → Settings → Pages → Source: "GitHub Actions"

3. W pliku `manifest.production.xml` podmień `GITHUB_USERNAME`
   na swój username GitHuba

4. W `package.json` podmień `GITHUB_USERNAME` w polu `homepage`

5. Push na main:
   ```
   git add .
   git commit -m "initial commit"
   git remote add origin https://github.com/TWOJ_USERNAME/ai-mail-assistant.git
   git push -u origin main
   ```

6. GitHub Actions automatycznie zbuduje i wdroży na:
   `https://TWOJ_USERNAME.github.io/ai-mail-assistant/`

7. W Outlooku usuń stary add-in (localhost) i zainstaluj nowy
   z pliku manifest.production.xml

### Aktualizacja:
Każdy push na branch `main` automatycznie buduje i wdraża nową wersję.
Żadnych serwerów, żadnych procesów na kompie.

## Jak to działa (architektura)

Outlook ładuje panel boczny (taskpane) z GitHub Pages.
Panel to statyczna strona HTML+JS — zero backendu.
Wywołania API lecą BEZPOŚREDNIO z przeglądarki Outlooka do wybranego API.
Klucz API trzymany w localStorage przeglądarki (wpisujesz raz per provider).

```
Outlook → [panel HTML z GitHub Pages] → [fetch do wybranego API] → odpowiedź

Obsługiwane API:
- api.anthropic.com (Claude)
- api.openai.com (GPT)
- generativelanguage.googleapis.com (Gemini)
```

Żaden pośredni serwer nie jest potrzebny.

## Bezpieczeństwo
- Klucze API trzymane w localStorage — akceptowalne dla narzędzia osobistego
- Repo może być PRYWATNE — GitHub Pages działa też z prywatnych repo
  (wymaga GitHub Pro lub GitHub Free z Pages włączonym dla prywatnych repo)
- Treść maili NIE jest nigdzie zapisywana — idzie tylko do API wybranego
  providera i wraca jako odpowiedź
- Nikt poza Tobą nie ma dostępu do Twoich kluczy API

## Dodawanie nowych providerów

Architektura Provider Adapter Pattern pozwala dodać nowego providera
(np. Mistral, Grok, Ollama) przez:
1. Dodanie obiektu adaptera w `src/taskpane/providers.js`
2. Dodanie wpisu w registry `PROVIDERS`

Żadnych zmian w UI ani w `taskpane.js`.

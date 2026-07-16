# Moje jídlo

Mobilní webová aplikace pro výběr aktuálních jídel kantýny ČS Antal. Nabídku čte z listu `Historie` v Google Sheetu a každý potvrzený výběr ukládá do listu `Moje jidlo` včetně data a času.

Produkční adresa frontendu: <https://blovak.github.io/jidlo/>

## Jak řešení funguje

- GitHub Pages hostuje statický frontend bez serveru.
- Google Apps Script funguje jako malé API nad tabulkou.
- `GET action=menu` vrací dnešní placené položky rozdělené podle sekcí.
- Chybějící energetické hodnoty odhadne dávkově OpenAI API a backend je po zbytek dne uchová v trvalé cache.
- `POST action=save` uloží každé vybrané jídlo jako samostatný řádek.
- `GET action=status` ověří, že byl zápis opravdu dokončen.
- ID výběru a zámek Apps Scriptu brání duplicitám při opakovaném požadavku.

## 1. Nasazení Google Apps Script API

1. Otevřete [zdrojový Google Sheet](https://docs.google.com/spreadsheets/d/1XE-ZsuxcRExU1Z9jymjbwt34abTSeQn2NvbKA74u9wg/edit).
2. Zvolte **Rozšíření → Apps Script**.
3. Přidejte nový soubor skriptu, například `WebApp.gs`.
4. Zkopírujte do něj celý obsah souboru [`apps-script/Code.gs`](apps-script/Code.gs). Kód může být ve stejném projektu jako existující automatický import menu; používá vlastní konstantu `WEBAPP_CONFIG`, aby se s importem nekřížil.
5. V editoru vyberte funkci `setup`, klikněte na **Spustit** a potvrďte oprávnění. Tím vznikne list `Moje jidlo` s hlavičkami.
6. V **Nastavení projektu → Vlastnosti skriptu** přidejte `OPENAI_API_KEY` s API klíčem. Klíč nikdy nevkládejte do `config.js` ani do repozitáře.
7. Volitelně přidejte `OPENAI_MODEL`; výchozí model je `gpt-5.6-luna`.
8. Klikněte na **Nasadit → Nové nasazení**.
9. Jako typ zvolte **Webová aplikace**.
10. Nastavte:
   - **Spouštět jako:** Já
   - **Kdo má přístup:** Kdokoli
11. Dokončete nasazení a zkopírujte adresu webové aplikace končící `/exec`.

> Při každé pozdější změně backendu vytvořte přes **Nasadit → Spravovat nasazení → Upravit** novou verzi. URL `/exec` zůstane stejná.

Energetická hodnota je orientační AI odhad typické porce podle názvu, sekce a alergenů, nikoli laboratorní nebo výrobcem garantovaný údaj. Když API klíč chybí nebo OpenAI dočasně selže, jídelníček se načte dál, pouze bez kcal. Odhady se ukládají do Script Properties podle data a ID jídla; struktura listu `Historie` se nemění.

## 2. Propojení frontendu

V souboru [`config.js`](config.js) nahraďte zástupnou hodnotu adresou z předchozího kroku:

```js
window.APP_CONFIG = Object.freeze({
  API_URL: 'https://script.google.com/macros/s/VAŠE_ID/exec'
});
```

Změnu commitněte a pushněte do větve `main`. GitHub Pages ji obvykle zveřejní během několika minut.

## 3. GitHub Pages

V repozitáři `Blovak/jidlo` otevřete **Settings → Pages** a nastavte:

- **Source:** Deploy from a branch
- **Branch:** `main`
- **Folder:** `/ (root)`

## Ukládaná data

List `Moje jidlo` obsahuje sloupce:

| Sloupec | Obsah |
| --- | --- |
| Datum a čas výběru | Okamžik potvrzení výběru v časovém pásmu Praha |
| Datum menu | Datum, ke kterému patří nabídka |
| Sekce | Polévka, hotová jídla, buffet apod. |
| Položka | Název jídla |
| Alergeny | Čísla alergenů převzatá z menu |
| Cena (Kč) | Cena položky |
| Klíč položky | Stabilní identifikátor řádku historie |
| ID výběru | Identifikátor jednoho potvrzení pro deduplikaci |

## Lokální spuštění

Projekt nemá build krok ani závislosti:

```bash
python3 -m http.server 8080
```

Potom otevřete <http://localhost:8080/jidlo/>.

## Poznámka k veřejnému zápisu

Webová aplikace Apps Script musí být dostupná anonymně, aby mohli návštěvníci GitHub Pages ukládat výběr bez přihlášení ke Googlu. To také znamená, že URL API je veřejná. Backend proto omezuje počet jídel v jednom výběru, ověřuje položky proti dnešní nabídce a deduplikuje požadavky; nejde však o náhradu uživatelské autentizace.

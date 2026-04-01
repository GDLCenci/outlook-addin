# Outlook Task Add-in

Add-in personalizzato per Outlook che permette di creare task in Notion + M365 To Do direttamente dalla lettura di un'email, con interpretazione del linguaggio naturale tramite Claude.

## Architettura

```
Outlook (task pane)                         Claude (sync agent)
┌──────────────────┐                       ┌──────────────────┐
│  Form + note     │                       │  Agenda Sync     │
│  in italiano     │── salva bozza ──>     │  (ogni ora)      │
│                  │   nella Drafts        │                  │
│  Office.js       │   folder via          │  1. Trova bozza  │
│  legge email     │   REST API            │     // TASK      │
│                  │                       │  2. Claude        │
│  Hosted su       │                       │     interpreta   │
│  GitHub Pages    │                       │  3. Crea Notion   │
└──────────────────┘                       │  4. Crea To Do   │
                                           │  5. Cancella      │
                                           │     bozza        │
                                           └──────────────────┘
```

**Zero backend. Zero database. Solo file statici + il sync agent già esistente.**

## Come funziona

1. Giuseppe apre un'email in Outlook
2. Il pannello add-in si apre sulla destra
3. Il form pre-compila il titolo dall'oggetto email
4. Giuseppe scrive note in linguaggio naturale + seleziona area/priorità/scadenza (opzionale)
5. Click "Crea Task"
6. L'add-in salva una bozza strutturata `// TASK` nella cartella Drafts
7. Il sync agent (ogni ora) trova la bozza, Claude la interpreta, crea task in Notion + To Do, cancella la bozza

## File

| File | Ruolo |
|------|-------|
| `manifest.xml` | Registra l'add-in in Outlook (sideload) |
| `src/taskpane.html` | UI del pannello laterale |
| `src/taskpane.css` | Stili (Fluent UI) |
| `src/taskpane.js` | Logica: Office.js + salvataggio bozza via REST API |
| `src/assets/` | Icone per il manifest |

## Setup

### 1. Hosting (GitHub Pages)
1. Creare un repo GitHub (es. `outlook-addin`)
2. Pushare il contenuto di questo progetto
3. Abilitare GitHub Pages (Settings → Pages → Source: main, root)
4. L'URL sarà: `https://{username}.github.io/outlook-addin/`

### 2. Aggiornare manifest.xml
Sostituire tutti i riferimenti a `giuseppedilollo.github.io` con il tuo URL reale.

### 3. Sideload in Outlook
**Desktop (Windows):**
1. Creare una cartella condivisa (es. `C:\ManifestShare\`)
2. Copiare `manifest.xml` nella cartella
3. Outlook → File → Options → Trust Center → Trusted Add-in Catalogs
4. Aggiungere `\\localhost\ManifestShare\` (o il path della cartella)
5. Riavviare Outlook → Get Add-ins → My Add-ins

**Outlook Web:**
1. Outlook Web → Settings → View all settings → Add-ins → Custom add-ins
2. "Add from file" → caricare `manifest.xml`

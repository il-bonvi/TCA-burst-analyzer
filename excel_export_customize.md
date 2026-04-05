# Come personalizzare colori e formato dell'export Excel

File da editare: `backend/excel_export.py`

---

## 1. Cambiare i colori delle metriche

All'inizio del file c'è un blocco con tutte le costanti colore:

```python
# METRIC ACCENT COLORS
C_WATT  = "FFC084FC"   # viola  — watt / potenza
C_HR    = "FFEF4444"   # rosso  — frequenza cardiaca
C_CAD   = "FF4DA6FF"   # blu    — cadenza
C_TIME  = "FF3CCF7E"   # verde  — tempo / durata
C_DELTA = "FF3CCF7E"   # verde  — delta sopra soglia
C_MAX   = "FF9555CC"   # viola scuro — max power
C_MIN   = "FFD9A8FF"   # lavanda — min power
```

Il formato è **`FFRRGGBB`** dove:
- `FF` = opacità piena (lascialo sempre `FF`)
- `RRGGBB` = colore in esadecimale standard

**Esempio** — vuoi l'HR arancione invece di rosso:
```python
C_HR = "FFFF8C00"   # arancione
```

Per trovare il codice HEX di un colore qualsiasi usa:
- [coolors.co](https://coolors.co) → copia il valore HEX e aggiungi `FF` davanti
- Qualsiasi color picker online

---

## 2. Cambiare i colori dello sfondo (chrome)

```python
BG_DARK  = "FF0D0F14"   # barre titolo molto scure
BG_MID   = "FF1A1D24"   # titolo foglio "All Bursts"
BG_HEAD  = "FF252932"   # righe intestazione colonne
```

---

## 3. Cambiare font, dimensione, grassetto

La funzione `_font()` è usata ovunque. Puoi cambiarla globalmente:

```python
def _font(color=T_DARK, bold=False, size=10, name="Arial") -> Font:
    return Font(bold=bold, size=size, color=color, name=name)
```

- Cambia `"Arial"` con `"Calibri"`, `"Helvetica"`, `"Courier New"`, ecc.
- Cambia `size=10` con `size=11` per testo più grande di default

---

## 4. Cambiare larghezza colonne

Ogni foglio ha una lista di tuple `(nome_colonna, larghezza)`.

**Esempio nel foglio per soglia:**
```python
GRID_COLS = [
    ("Duration (s)",  13),   # ← cambia il numero per allargare/restringere
    ("Count",         10),
    ("Avg Power (W)", 13),
    ...
]
```

---

## 5. Cambiare l'altezza delle righe

Cerca le righe tipo:
```python
ws.row_dimensions[row_num].height = 18
```

Aumenta il numero per righe più alte (es. `22`).

---

## 6. Aggiungere una colonna nuova

1. Aggiungi la tupla alla lista `*_COLS` del foglio:
   ```python
   ("Fatigue Idx", 13),
   ```
2. Aggiungi il valore corrispondente nella lista `row_vals`:
   ```python
   row_vals = [..., round(b.get("fatigue_idx", 1), 3)]
   ```
3. Aggiungi il colore in `tc_map` (usa l'indice colonna, 1-based):
   ```python
   tc_map[9] = C_WATT   # colore per la nuova colonna
   ```

---

## 7. Togliere un intero foglio

Commenta o rimuovi il blocco corrispondente in `build_excel()`:

```python
# ── LAST SHEET: ALL BURSTS DETAIL ─────────────────────
# ws_all = wb.create_sheet("All Bursts")
# ...
```

---

## Riferimento rapido colori utili

| Colore       | ARGB       |
|-------------|------------|
| Giallo oro  | `FFF0C040` |
| Viola       | `FFC084FC` |
| Rosso       | `FFEF4444` |
| Blu chiaro  | `FF4DA6FF` |
| Verde       | `FF3CCF7E` |
| Arancione   | `FFFB923C` |
| Bianco      | `FFFFFFFF` |
| Nero        | `FF000000` |
| Grigio      | `FF6B7280` |

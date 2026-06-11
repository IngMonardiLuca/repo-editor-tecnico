"""
Test headless dei pulsanti della toolbar del blocco testo dopo l'aggiunta del
pulsante "lettere greche". Verifica che:
  - il nuovo pulsante greco abbia lo stile FISSO identico agli altri pulsanti
    fissi (apice/pedice) e che NON cambi dopo update_format_buttons();
  - i pulsanti di stato B / I / U continuino a togglare sfondo
    (#fafafa inattivo / #e8e8e8 attivo) esattamente come prima;
  - gli altri pulsanti fissi non vengano alterati da update_format_buttons();
  - l'inserimento di una lettera greca finisca davvero nell'editor.

Esecuzione (dalla cartella PROGETTO EDITOR):
    QT_QPA_PLATFORM=offscreen ./.venv/Scripts/python.exe test_greek_buttons.py
"""

import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

import main


FIXED_STYLE = (
    "QPushButton { border: 1px solid #d6d6d6; border-radius: 4px; background: #fafafa; }"
    "QPushButton:hover { background: #f0f0f0; }"
    "QToolTip { background-color: #f2f2f2; color: black; border: 1px solid #c8c8c8; padding: 4px; }"
)

BOLD_INACTIVE = (
    "QPushButton { font-weight: bold; border: 1px solid #d6d6d6; background: #fafafa; }"
    "QToolTip { background-color: #f2f2f2; color: black; border: 1px solid #c8c8c8; padding: 4px; }"
)

BOLD_ACTIVE = (
    "QPushButton { font-weight: bold; border: 1px solid #999; background: #e8e8e8; }"
    "QToolTip { background-color: #f2f2f2; color: black; border: 1px solid #c8c8c8; padding: 4px; }"
)


class StubMainWindow:
    """Minimo indispensabile usato da TextBlockWidget in costruzione."""
    def ensure_block_defaults(self, block):
        block = dict(block or {})
        block.setdefault("id", "testid")
        block.setdefault("meta", {})
        block.setdefault("collapsed", False)
        return block

    def mark_dirty(self):
        pass


class StubEditor:
    def __init__(self, main_window):
        self.main_window = main_window


def main_test():
    app = QApplication.instance() or QApplication([])

    w = main.TextBlockWidget(StubEditor(StubMainWindow()))

    failures = []

    def check(cond, msg):
        if cond:
            print(f"  OK   {msg}")
        else:
            print(f"  FAIL {msg}")
            failures.append(msg)

    print("== Stato iniziale ==")
    # I pulsanti fissi (apice/pedice/greco) hanno lo stesso stile fisso atteso.
    check(w.btn_greek.styleSheet() == FIXED_STYLE,
          "il pulsante greco ha lo stile fisso atteso")
    check(w.btn_subscript.styleSheet() == FIXED_STYLE,
          "il pedice ha lo stile fisso atteso")
    check(w.btn_superscript.styleSheet() == FIXED_STYLE,
          "l'apice ha lo stile fisso atteso")
    check(w.btn_greek.styleSheet() == w.btn_subscript.styleSheet(),
          "greco e pedice hanno stile identico (niente ereditarieta')")

    # Snapshot prima di update_format_buttons()
    snap = {
        "greek": w.btn_greek.styleSheet(),
        "superscript": w.btn_superscript.styleSheet(),
        "subscript": w.btn_subscript.styleSheet(),
        "bullets": w.btn_bullets.styleSheet(),
        "numbers": w.btn_numbers.styleSheet(),
        "justify": w.btn_justify.styleSheet(),
        "center": w.btn_center.styleSheet(),
    }

    print("== Dopo update_format_buttons() su editor vuoto ==")
    w.update_format_buttons()

    for name, before in snap.items():
        btn = getattr(w, f"btn_{name}")
        check(btn.styleSheet() == before,
              f"il pulsante '{name}' NON e' cambiato dopo update_format_buttons()")

    check(w.btn_bold.styleSheet() == BOLD_INACTIVE,
          "B (grassetto) e' nello stato INATTIVO su editor vuoto")

    print("== Toggle grassetto: B deve diventare ATTIVO, il greco invariato ==")
    greek_before_toggle = w.btn_greek.styleSheet()
    w.toggle_bold()
    w.update_format_buttons()
    check(w.btn_bold.styleSheet() == BOLD_ACTIVE,
          "B (grassetto) passa allo stato ATTIVO (sfondo #e8e8e8)")
    check(w.btn_greek.styleSheet() == greek_before_toggle,
          "il pulsante greco resta invariato durante il toggle di B")

    print("== Inserimento lettera greca ==")
    w.insert_greek_letter("α")
    check("α" in w.editor.toPlainText(),
          "la lettera greca alfa risulta inserita nell'editor")

    print()
    if failures:
        print(f"RISULTATO: {len(failures)} test FALLITI")
        return 1
    print("RISULTATO: tutti i test PASSATI")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main_test())

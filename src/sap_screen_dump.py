from __future__ import annotations
from pathlib import Path
from datetime import datetime

IMPORTANT_TYPES = {
    "GuiTextField",
    "GuiCTextField",
    "GuiComboBox",
    "GuiCheckBox",
    "GuiRadioButton",
    "GuiLabel",
    "GuiButton",
}

def dump_screen(session, out_dir: str = "data/dumps") -> str:
    """
    Gera dump estruturado da tela SAP.
    Mantém dump completo + seções inteligentes para IA e auditoria.
    """
    Path(out_dir).mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    info = session.Info
    prog = getattr(info, "Program", "UNK")
    scr = getattr(info, "ScreenNumber", "UNK")
    tcode = getattr(info, "Transaction", "UNK")
    file_path = Path(out_dir) / f"DUMP_{tcode}_{prog}_{scr}_{ts}.txt"
    lines: list[str] = []

    # HEADER
    lines.append(f"TCODE={tcode} PROGRAM={prog} SCREEN={scr}")
    lines.append(f"TIMESTAMP={ts}")
    lines.append("-" * 80)

    def safe_get(obj, attr):
        try:
            return getattr(obj, attr)
        except Exception:
            return ""

    # STATUS BAR
    try:
        sbar = session.findById("wnd[0]/sbar")
        lines.append("[STATUS_BAR]")
        lines.append(
            f"TYPE={safe_get(sbar, 'MessageType')} | TEXT={safe_get(sbar, 'Text')}"
        )
    except Exception:
        pass

    lines.append("-" * 80)

    # POPUP (se existir)
    try:
        wnd1 = session.findById("wnd[1]")
        lines.append("[POPUP]")
        lines.append(f"TITLE={safe_get(wnd1, 'Text')}")
    except Exception:
        pass

    lines.append("-" * 80)

    usr = session.findById("wnd[0]/usr")

    all_objects: list[str] = []
    relevant_objects: list[str] = []

    def walk(obj, depth=0):
        try:
            count = obj.Children.Count
        except Exception:
            count = 0

        for i in range(count):
            try:
                child = obj.Children(i)
            except Exception:
                continue

            obj_id = safe_get(child, "Id")
            obj_type = safe_get(child, "Type")
            name = safe_get(child, "Name")
            text = safe_get(child, "Text")
            tip = safe_get(child, "Tooltip")
            indent = "  " * depth
            line = (
                f"{indent}ID={obj_id} | TYPE={obj_type} | "
                f"NAME={name} | TEXT={text} | TIP={tip}"
            )

            all_objects.append(line)

            if obj_type in IMPORTANT_TYPES and (text or name):
                relevant_objects.append(line)

            walk(child, depth + 1)

    walk(usr)

    # RELEVANT OBJECTS
    lines.append("[RELEVANT_OBJECTS]")
    lines.extend(relevant_objects or ["<NONE>"])
    lines.append("-" * 80)

    # FULL TREE (DEBUG)
    lines.append("[FULL_TREE]")
    lines.extend(all_objects)

    file_path.write_text("\n".join(lines), encoding="utf-8", errors="ignore")
    return str(file_path)
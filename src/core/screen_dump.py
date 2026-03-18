import os
from pathlib import Path
from datetime import datetime

class ScreenDumper:
    IMPORTANT_TYPES = {
        "GuiTextField", "GuiCTextField", "GuiComboBox", 
        "GuiCheckBox", "GuiRadioButton", "GuiLabel", "GuiButton"
    }

    @staticmethod
    def dump(session, out_dir: str = "data/dumps") -> str:
        Path(out_dir).mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        info = session.Info
        file_path = Path(out_dir) / f"DUMP_{info.Transaction}_{ts}.txt"
        
        lines = [f"TCODE={info.Transaction} PROG={info.Program} SCR={info.ScreenNumber}", "-"*50]
        
        try:
            sb = session.findById("wnd[0]/sbar")
            lines.append(f"[STATUSBAR] {sb.MessageType}: {sb.Text}")
        except: pass

        relevant = []
        def walk(obj):
            try:
                for i in range(obj.Children.Count):
                    child = obj.Children(i)
                    t = child.Type
                    if t in ScreenDumper.IMPORTANT_TYPES:
                        relevant.append(f"ID={child.Id} | TYPE={t} | TEXT={getattr(child, 'Text', '')}")
                    walk(child)
            except: pass
            
        walk(session.findById("wnd[0]/usr"))
        lines.append("[OBJECTS]")
        lines.extend(relevant)
        
        file_path.write_text("\n".join(lines), encoding="utf-8")
        return str(file_path)
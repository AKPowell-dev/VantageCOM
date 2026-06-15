"""One-off generator for the econs UserForms (.frm + .frx).

Builds the dialogs' controls and their cosmetic properties (fonts, colors,
header band, sunken fields) via Excel automation, then exports them to the repo
Forms folder. The control positions/sizes set here are just initial values: all
three dialogs reflow at run time from their own Relayout method (so they can be
resized). Code-behind is maintained as text inside each .frm and is NOT written
by this script, so re-running it overwrites the .frm/.frx and the code-behind
must be re-applied afterwards. Re-run only when the control set or styling needs
to change.
"""
import sys
from pathlib import Path

import pythoncom  # type: ignore
from win32com.client import DispatchEx  # type: ignore

FORMS_DIR = Path(__file__).resolve().parent.parent / "Forms"
MAX_ROWS = 5
VBEXT_CT_MSFORM = 3
FONT_NAME = "Garamond"


def bgr(r, g, b):
    """MSForms colors are stored as BGR longs."""
    return r | (g << 8) | (b << 16)


# Palette (deep blue accent: #1F4E79)
ACCENT = bgr(31, 78, 121)
ACCENT_TXT = bgr(255, 255, 255)
FORM_BG = bgr(246, 246, 249)
FIELD_BG = bgr(255, 255, 255)
BODY_FG = bgr(64, 64, 64)
SECTION_FG = bgr(31, 78, 121)
HILITE = bgr(214, 228, 245)   # light blue band behind the replace checkbox


def add(des, progid, name, **props):
    ctl = des.Controls.Add(progid, name, True)
    for key, value in props.items():
        setattr(ctl, key, value)
    return ctl


def font(ctl, size=11, bold=False):
    f = ctl.Font
    f.Name = FONT_NAME
    f.Size = size
    f.Bold = bold


def header(des, caption, width):
    """A blue band plus a transparent, vertically centred title label."""
    band = add(des, "Forms.Label.1", "lblHeaderBand", Left=0, Top=0,
               Width=width, Height=24, BackColor=ACCENT)
    band.BackStyle = 1  # opaque
    txt = add(des, "Forms.Label.1", "lblHeaderText", Caption=caption, Left=12,
              Top=3, Width=width - 24, Height=18, ForeColor=ACCENT_TXT,
              TextAlign=1)
    txt.BackStyle = 0   # transparent (so the band shows through)
    font(txt, size=14, bold=True)


def case_listbox(des, name, top, width):
    lst = add(des, "Forms.ListBox.1", name, Left=12, Top=top, Width=width,
              Height=80, BackColor=FIELD_BG)
    lst.MultiSelect = 1   # fmMultiSelectMulti
    lst.ListStyle = 1     # fmListStyleOption (checkboxes)
    font(lst)
    return lst


def build_main(des):
    header(des, "  Econs  -  PowerPoint Output", 444)

    reset = add(des, "Forms.CommandButton.1", "btnResetAll", Caption="Reset All",
                Left=444 - 12 - 70, Top=3, Width=70, Height=18)
    font(reset, size=10)

    y = 42
    lbl = add(des, "Forms.Label.1", "lblNum", Caption="Number of outputs:",
              Left=12, Top=y + 2, Width=120, Height=16, ForeColor=BODY_FG)
    font(lbl)
    cbo = add(des, "Forms.ComboBox.1", "cboNum", Left=140, Top=y, Width=56,
              Height=18, BackColor=FIELD_BG)
    cbo.Style = 2
    font(cbo)
    y += 28

    chk = add(des, "Forms.CheckBox.1", "chkReplace",
              Caption="REPLACE existing outputs on matching slides (otherwise append new slides)",
              Left=12, Top=y, Width=430, Height=22, ForeColor=ACCENT,
              BackColor=HILITE)
    chk.BackStyle = 1  # opaque highlight band so it stands out
    font(chk, size=12, bold=True)
    y += 28

    chk2 = add(des, "Forms.CheckBox.1", "chkCaseFirst",
               Caption="Slide title order: Case | Output  (unchecked = Output | Case)",
               Left=12, Top=y, Width=430, Height=18, ForeColor=BODY_FG,
               BackColor=FORM_BG)
    font(chk2)
    y += 24

    sec = add(des, "Forms.Label.1", "lblSecOutputs", Caption="OUTPUTS", Left=12,
              Top=y, Width=200, Height=16, ForeColor=SECTION_FG)
    font(sec, size=11, bold=True)
    y += 18

    lbl = add(des, "Forms.Label.1", "lblHdrName",
              Caption="Output name (shown in PowerPoint)", Left=34, Top=y,
              Width=190, Height=16, ForeColor=BODY_FG)
    font(lbl, size=11)
    lbl = add(des, "Forms.Label.1", "lblHdrRange", Caption="Named range",
              Left=230, Top=y, Width=120, Height=16, ForeColor=BODY_FG)
    font(lbl, size=11)
    y += 18

    for i in range(1, MAX_ROWS + 1):
        lbl = add(des, "Forms.Label.1", "lblRow%d" % i, Caption="%d." % i,
                  Left=12, Top=y + 2, Width=20, Height=16, ForeColor=BODY_FG)
        font(lbl)
        t = add(des, "Forms.TextBox.1", "txtName%d" % i, Left=34, Top=y,
                Width=190, Height=18, BackColor=FIELD_BG)
        font(t)
        t = add(des, "Forms.TextBox.1", "txtRange%d" % i, Left=230, Top=y,
                Width=120, Height=18, BackColor=FIELD_BG)
        font(t)
        y += 24

    y += 6
    sec = add(des, "Forms.Label.1", "lblSecCases", Caption="CASES", Left=12,
              Top=y, Width=200, Height=16, ForeColor=SECTION_FG)
    font(sec, size=11, bold=True)
    y += 18

    lbl = add(des, "Forms.Label.1", "lblCaseRange",
              Caption="Case input cell (named range):", Left=12, Top=y + 2,
              Width=168, Height=16, ForeColor=BODY_FG)
    font(lbl)
    t = add(des, "Forms.TextBox.1", "txtCaseRange", Left=186, Top=y, Width=180,
            Height=18, BackColor=FIELD_BG)
    font(t)
    btn = add(des, "Forms.CommandButton.1", "btnRefreshCases",
              Caption="Clear Cases", Left=444 - 12 - 80, Top=y - 1, Width=80,
              Height=20)
    font(btn, size=11)
    y += 24

    lbl = add(des, "Forms.Label.1", "lblCases",
              Caption="Cases to print (check to include):", Left=12, Top=y,
              Width=260, Height=16, ForeColor=BODY_FG)
    font(lbl)
    y += 18

    case_listbox(des, "lstCases", y, 424)

    adv = add(des, "Forms.CommandButton.1", "btnAdvanced", Caption="Advanced...",
              Left=12, Top=452, Width=96, Height=26)
    font(adv)
    ok = add(des, "Forms.CommandButton.1", "btnOK", Caption="RUN", Left=270,
             Top=452, Width=84, Height=26, BackColor=ACCENT, ForeColor=ACCENT_TXT)
    ok.Default = True
    font(ok, bold=True)
    cancel = add(des, "Forms.CommandButton.1", "btnCancel", Caption="Cancel",
                 Left=360, Top=452, Width=84, Height=26)
    cancel.Cancel = True
    font(cancel)

    hist = add(des, "Forms.CommandButton.1", "btnHistory", Caption="History",
               Left=114, Top=452, Width=84, Height=26)
    font(hist)


def build_advanced(des):
    width = 462
    header(des, "  Advanced  -  Per-Case Extra Outputs", width)

    y = 38
    for idx in (1, 2):
        sec = add(des, "Forms.Label.1", "lblSec%d" % idx,
                  Caption="EXTRA OUTPUT %d" % (idx + 2), Left=12, Top=y,
                  Width=300, Height=16, ForeColor=SECTION_FG)
        font(sec, size=11, bold=True)
        y += 20

        lbl = add(des, "Forms.Label.1", "lblAdvName%d" % idx, Caption="Name:",
                  Left=12, Top=y + 2, Width=40, Height=16, ForeColor=BODY_FG)
        font(lbl)
        t = add(des, "Forms.TextBox.1", "txtAdvName%d" % idx, Left=54, Top=y,
                Width=160, Height=18, BackColor=FIELD_BG)
        font(t)
        lbl = add(des, "Forms.Label.1", "lblAdvRange%d" % idx, Caption="Range:",
                  Left=224, Top=y + 2, Width=46, Height=16, ForeColor=BODY_FG)
        font(lbl)
        t = add(des, "Forms.TextBox.1", "txtAdvRange%d" % idx, Left=272, Top=y,
                Width=178, Height=18, BackColor=FIELD_BG)
        font(t)
        y += 26

        lbl = add(des, "Forms.Label.1", "lblApply%d" % idx,
                  Caption="Apply to cases:", Left=12, Top=y, Width=200,
                  Height=16, ForeColor=BODY_FG)
        font(lbl)
        clr = add(des, "Forms.CommandButton.1", "btnAdvClear%d" % idx,
                  Caption="Clear Cases", Left=width - 12 - 80, Top=y - 1,
                  Width=80, Height=20)
        font(clr, size=11)
        y += 18

        case_listbox(des, "lstAdvCases%d" % idx, y, 438)
        y += 88

    y += 8
    ok = add(des, "Forms.CommandButton.1", "btnAdvOK", Caption="OK", Left=278,
             Top=y, Width=84, Height=26, BackColor=ACCENT, ForeColor=ACCENT_TXT)
    ok.Default = True
    font(ok, bold=True)
    cancel = add(des, "Forms.CommandButton.1", "btnAdvCancel", Caption="Cancel",
                 Left=368, Top=y, Width=84, Height=26)
    cancel.Cancel = True
    font(cancel)


def build_history(des):
    width = 444
    header(des, "  Econs  -  History", width)

    lbl = add(des, "Forms.Label.1", "lblHistHint",
              Caption="Select a previous run to load. Cases this model no longer "
                      "offers will be skipped.",
              Left=12, Top=30, Width=width - 24, Height=16, ForeColor=BODY_FG)
    font(lbl, size=11)

    lst = add(des, "Forms.ListBox.1", "lstHistory", Left=12, Top=50,
              Width=width - 24, Height=230, BackColor=FIELD_BG)
    lst.ColumnCount = 2
    font(lst)

    det = add(des, "Forms.TextBox.1", "txtDetails", Left=12, Top=286,
              Width=width - 24, Height=64, BackColor=FORM_BG)
    det.MultiLine = True
    det.WordWrap = True
    det.Locked = True
    det.ScrollBars = 2  # fmScrollBarsVertical
    font(det)

    load = add(des, "Forms.CommandButton.1", "btnHistLoad", Caption="Load",
               Left=width - 180, Top=360, Width=84, Height=26, BackColor=ACCENT,
               ForeColor=ACCENT_TXT)
    load.Default = True
    font(load, bold=True)
    cancel = add(des, "Forms.CommandButton.1", "btnHistCancel", Caption="Cancel",
                 Left=width - 90, Top=360, Width=84, Height=26)
    cancel.Cancel = True
    font(cancel)


def build_form(vbproj, name, caption, width, height, builder):
    comp = vbproj.VBComponents.Add(VBEXT_CT_MSFORM)
    comp.Properties("Width").Value = width
    comp.Properties("Height").Value = height
    try:
        comp.Name = name
    except Exception:
        pass
    comp.Properties("Caption").Value = caption
    comp.Properties("BackColor").Value = FORM_BG
    builder(comp.Designer)
    out_frm = FORMS_DIR / (name + ".frm")
    comp.Export(str(out_frm))
    print("Exported:", out_frm)


def main() -> int:
    pythoncom.CoInitialize()
    excel = DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Add()
        vbproj = wb.VBProject
        FORMS_DIR.mkdir(parents=True, exist_ok=True)
        build_form(vbproj, "UF_EconsConfig", "Econs - PowerPoint Output",
                   468, 512, build_main)
        build_form(vbproj, "UF_EconsAdvanced", "Econs - Advanced",
                   480, 446, build_advanced)
        build_form(vbproj, "UF_EconsHistory", "Econs - History",
                   468, 420, build_history)
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""Per-device cell maps for the Excel Data Aggregator.

The KEY of each entry (e.g. "AGH") must match the code that
``extract_device_code()`` pulls out of the source filename.

Almost every inspection form shares one layout, so the cell maps are built by
``form()`` rather than written out by hand.  A device that needs unusual
coordinates passes ``extra={...}`` (or gets a literal dict), which makes the
genuinely odd forms easy to spot in an otherwise uniform list.

TODO (carried over from the previous version of this file) — verify the cell
coordinates for: Phototherapy, Sealing Machine serial, OR Table, EA / C-Arm,
C-pap, FJ, Refrigerator serial.
"""

CellMap = dict[str, str]


def form(
    row: int,
    status: str,
    *,
    col: str = "E",
    val: str = "K",
    date_gap: int = 2,
    extra: CellMap | None = None,
) -> CellMap:
    """Build a cell map for the standard inspection-form layout.

    ``row`` is the row holding the Model.  The rest of the fields sit at fixed
    offsets from it::

        Date          {col}{row - date_gap}     e.g. form(18, ...) -> E16
        Model         {col}{row}                                   -> E18
        Manufacturer  {col}{row + 2}                               -> E20
        S.N           {val}{row}                                   -> K18
        Location      {val}{row + 2}                               -> K20

    ``col`` is the left-hand (label) column, ``val`` the right-hand one.
    ``date_gap`` is 2 on most forms and 4 where an extra line precedes the Date.
    ``extra`` adds or overrides individual cells — use it for a second serial
    (``S.N2``), a second status (``Status2``), or a one-off Location.
    """
    cells: CellMap = {
        "Manufacturer": f"{col}{row + 2}",
        "Model": f"{col}{row}",
        "S.N": f"{val}{row}",
        "Location": f"{val}{row + 2}",
        "Date": f"{col}{row - date_gap}",
        "Status": status,
    }
    if extra:
        cells.update(extra)
    return cells


# ──────────────────────────────────────────────────────────────────────────────
# Device table
#
#   device_name  — written to the Device column of the output
#   cells        — where to read each field from in the source form
#   second_row   — optional: auto-generate a row for a sub-module of the same
#                  physical unit.  Its Status comes from the parent's "Status2"
#                  cell, and code_replace rewrites the device code.
#   alt_cells    — optional list of older layouts for the same device, used when
#                  a form was re-laid-out between rounds.  Purely additive:
#                  `cells` is always tried first and only a file it cannot read
#                  sensibly falls through to an alternate, so adding one can
#                  rescue old forms but can never change a file that already
#                  reads correctly.  Read by firebase_export.py; calist.py
#                  ignores it, so the register output is unaffected.
# ──────────────────────────────────────────────────────────────────────────────

DEVICE_CONFIGS: dict[str, dict] = {
    # ── Two-row devices (main unit + sub-module) ──────────────────────────────
    "AGH": {
        "device_name": "Patient Monitor",
        "cells": form(18, "D39", extra={"Status2": "J39"}),
        "second_row": {"device_name": "NIBP", "code_replace": ("AGH", "AGCB")},
    },
    "VAH": {
        "device_name": "Vital Sign (SPO2 Module)",
        "cells": form(18, "G38", extra={"Status2": "J38"}),
        "second_row": {
            "device_name": "Vital Sign (NIBP Module)",
            "code_replace": ("VAH", "VCB"),
        },
    },

    # ── Imaging: a second serial (tube / probe) shifts Location down ──────────
    "BB": {"device_name": "Ultrasound", "cells": form(17, "H30", col="F", val="L", extra={"S.N2": "L21"})},
    "BF": {"device_name": "X-ray",      "cells": form(18, "J27", extra={"S.N2": "K20", "Location": "K22"})},
    # NOTE: identical to BF. The trailing "()" looks like an unfinished edit —
    # confirm what this code is meant to be called before the next run.
    "CA": {"device_name": "X-ray ()",   "cells": form(18, "J27", extra={"S.N2": "K20", "Location": "K22"})},
    "EA": {"device_name": "C-Arm",      "cells": form(18, "J27", extra={"S.N2": "K20", "Location": "K22"})},

    # ── Forms using the D/J columns ───────────────────────────────────────────
    "AF": {"device_name": "ECG",          "cells": form(32, "F41", col="D", val="J")},
    "AL": {"device_name": "Phototherapy", "cells": form(28, "F37", col="D", val="J")},

    # ── Standard E/K forms ────────────────────────────────────────────────────
    "AC": {"device_name": "Defibrillator",          "cells": form(15, "G24")},
    "AA": {"device_name": "Anesthesia",             "cells": form(17, "G33")},
    "BP": {"device_name": "Balance",                "cells": form(18, "G30")},
    "EO": {"device_name": "Pipet",                  "cells": form(18, "H29")},
    "AH": {"device_name": "SPO2",                   "cells": form(14, "G27")},
    "EE": {"device_name": "Flowmeter",              "cells": form(17, "G31")},
    "GP": {"device_name": "Holter machines",        "cells": form(18, "G26")},
    "DV": {"device_name": "OR light",               "cells": form(15, "G24")},
    "AS": {"device_name": "Centrifuge",             "cells": form(18, "K25")},
    "AJ": {"device_name": "Suction",                "cells": form(23, "G32")},
    "AM": {"device_name": "Ventilator",             "cells": form(17, "G33")},
    "FG": {"device_name": "ACT",                    "cells": form(18, "H32")},
    "DG": {"device_name": "CBC Analyzer",           "cells": form(18, "H32")},
    "AU": {"device_name": "Chemistry analyzer",     "cells": form(18, "K22")},
    "AX": {"device_name": "Lab Incubator",          "cells": form(18, "H32")},
    "EY": {"device_name": "Freezer",                "cells": form(18, "H32")},
    "EP": {"device_name": "Refrigerator",           "cells": form(18, "H32")},
    "DL": {"device_name": "Sealing Machine",        "cells": form(18, "K22")},
    "BV": {"device_name": "Blood gas analyzer",     "cells": form(18, "K22")},
    "FQ": {"device_name": "C-pap",                  "cells": form(17, "G34")},
    # Two layouts in the wild. The original map is unchanged and still tried
    # first; the alternate sits two rows higher and rescues the forms where the
    # original reads the *location* into the serial ("ICU", "NICU"), the
    # manufacturer into the model, and nothing into the manufacturer.
    #
    # Measured over 120 random Infusion forms: 84 read plausibly either way,
    # 30 only with the alternate, and none only with the original — so adding
    # it can rescue files but can never cost one.
    "AI": {
        "device_name": "Infusion",
        "cells": form(27, "G36"),
        "alt_cells": [form(25, "G34")],
    },
    "DO": {"device_name": "O2 conc",                "cells": form(17, "G31")},
    "FJ": {"device_name": "OR Table",               "cells": form(18, "K22")},
    "AB": {"device_name": "Vaporizer",              "cells": form(18, "K22")},
    "AD": {"device_name": "Pacemaker",              "cells": form(22, "G31")},
    "AV": {"device_name": "Elisa reader",           "cells": form(18, "K22")},
    "FE": {"device_name": "Nebulizer",              "cells": form(17, "K31")},
    "AO": {"device_name": "Infrared",               "cells": form(18, "H32")},
    "GC": {"device_name": "Portable Data Logger",   "cells": form(20, "G33")},
    "DA": {"device_name": "Shaker",                 "cells": form(18, "H29")},
    "GI": {"device_name": "Bacteria Analyzer",      "cells": form(18, "K22")},
    "ED": {"device_name": "Heart lung Machine",     "cells": form(15, "G43")},
    # The forms in the archive all use the ordinary E/K layout — labels checked
    # in column D: "Model:" r18, "Manufacturer:" r20, "Serial No.:" I18,
    # "Location" I20. The original map is kept and still tried first.
    "FI": {
        "device_name": "Hormone Analyzer",
        "cells": form(15, "G43"),
        "alt_cells": [form(18, "K22")],
    },
    "GK": {"device_name": "Tornique",               "cells": form(23, "G32")},
    "AQ": {"device_name": "Water Bath",             "cells": form(18, "H32")},
    "EV": {"device_name": "Blood Mixer",            "cells": form(18, "K22")},
    "GD": {"device_name": "Protien Analyzer",       "cells": form(18, "K22")},
    "AR": {"device_name": "Electrolyte Analyzer",   "cells": form(18, "K22")},

    # ── Added from the archive ────────────────────────────────────────────────
    # Codes that appear in four years of forms but were never in this table, so
    # every one of their files was reported as an unrecognised device. Each map
    # below was derived from the form's own printed labels ("Model:", "Serial
    # No.:", "Location") across four files spread over different sites and
    # years — never copied from a similar device.
    "BM": {"device_name": "Hemodialysis Machine",   "cells": form(18, "K22")},
    "BN": {"device_name": "Therapeutic Ultrasound", "cells": form(18, "K22")},
    "CN": {"device_name": "Microwave",              "cells": form(18, "K22")},
    "GE": {"device_name": "Temperature Calibration Tester", "cells": form(18, "K22")},
    # Manufacturer sits two rows below Model rather than the usual one gap.
    "FZ": {
        "device_name": "Endoscope",
        "cells": {
            "Manufacturer": "E22",
            "Model": "E18",
            "S.N": "K18",
            "Location": "K22",
            "Date": "E14",
            "Status": "K26",
        },
    },
    # Location two rows lower than standard, and the status is over in J.
    "BE": {
        "device_name": "X-ray (Mobile)",
        "cells": {
            "Manufacturer": "E20",
            "Model": "E18",
            "S.N": "K18",
            "Location": "K22",
            "Date": "E14",
            "Status": "J27",
        },
    },

    # ── Forms with an extra line above the Date (date_gap=4) ──────────────────
    "BZ": {"device_name": "Syringe",                "cells": form(26, "G35", date_gap=4)},
    "CE": {"device_name": "Sphygmomanometer",       "cells": form(47, "H59", date_gap=4)},
    "CB": {"device_name": "Digital blood pressure", "cells": form(18, "G26", date_gap=4)},
    "AE": {"device_name": "ESU",                    "cells": form(15, "G24", date_gap=4)},
    "BL": {"device_name": "Autoclave",              "cells": form(18, "K22", date_gap=4)},
    "AN": {"device_name": "Thermo",                 "cells": form(18, "H32", date_gap=4)},
    "EC": {"device_name": "Laminar flow",           "cells": form(18, "F31", date_gap=4)},
    # NOTE: same device_name as "AO" above but a different Status cell.
    "CK": {"device_name": "Infrared",               "cells": form(18, "F31", date_gap=4)},

    # ── Odd one out ───────────────────────────────────────────────────────────
    # NOTE: Location is K19. Every other standard form puts it 2 rows below the
    # S.N (which is K18 here), i.e. K20 — this looks like a typo worth checking
    # against the actual Lab Oven form.
    "EU": {"device_name": "Lab Oven", "cells": form(18, "H32", extra={"Location": "K19"})},

    # ── Genuinely different forms — written out in full ───────────────────────
    # The incubator form was re-laid-out between the 2025 and 2026 rounds. The
    # map below is the *current* one and is unchanged; the alternate is the
    # older sheet, which every 2023 and 2024 form uses and most of 2025.
    #
    # Sampled 25 forms per year: 2023 and 2024 read only with the alternate,
    # 2026 only with the current map, and 2025 is the changeover — 21 old, 4
    # new. Field positions in the older sheet were confirmed against the
    # printed labels in column A ("Manufacturer:" r71, "Model:" r73,
    # "Location:" r75, "S.N:" r69), not inferred from the values.
    "AK": {
        "device_name": "Baby Incubator",
        "cells": {
            "Manufacturer": "D74",
            "Model": "D72",
            "S.N": "L72",
            "Location": "D78",   # in the label column, 6 rows down
            "Date": "D70",
            "Status": "F56",
        },
        "alt_cells": [
            {
                "Manufacturer": "D71",
                "Model": "D73",   # Model sits *below* Manufacturer on this one
                "S.N": "L69",
                "Location": "D75",
                "Date": "D67",
                "Status": "F56",  # unchanged across both layouts
            },
        ],
    },
    # NOTE: Model (D68) sits *below* Manufacturer (D66) here — inverted compared
    # with every other form — and S.N (L60) is 8 rows above the Model. Worth
    # re-checking against the paper form; one of these is likely transposed.
    "CF": {
        "device_name": "Baby Warmer",
        "cells": {
            "Manufacturer": "D66",
            "Model": "D68",
            "S.N": "L60",
            "Location": "D70",
            "Date": "D62",
            "Status": "F49",
        },
    },

    # "FV": {"device_name": "Endoscopy", "cells": form(18, "K22")},  # unverified
}

# "AG" is "AGH" with the H dropped — the same Patient Monitor form, typed short.
# 1,259 files in the archive use it, and every one was reported as an
# unrecognised device until now.
#
# The cell map is shared with AGH by reference so the two can never drift apart.
# The second row is *not*: code_replace rewrites the device code by literal
# substitution, so AGH's ("AGH" -> "AGCB") would find nothing in "AG001" and the
# NIBP row would silently keep the parent's code.
DEVICE_CONFIGS["AG"] = {
    "device_name": DEVICE_CONFIGS["AGH"]["device_name"],
    "cells": DEVICE_CONFIGS["AGH"]["cells"],
    "second_row": {"device_name": "NIBP", "code_replace": ("AG", "AGCB")},
}

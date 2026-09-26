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
#                  reads correctly.  Used by calist.read_best, so it does reach
#                  the register — a form that used to import blank now imports
#                  its readings.
#
# ── device_name vs. the master code list ──────────────────────────────────────
#
# `device_name` here and DEVICE_NAMES in device_names.py disagree for 24 codes,
# and that is expected rather than a fault to fix.  This file holds the wording
# engineers read in the register; the master document holds the wording the
# site's code list uses, and it carries typos of its own ("Utrasound,Abdomen",
# "Infrared lamb", "Bactrial identification device").
#
# Confirmed synonyms, not conflicts — do not "correct" these:
#
#   AH   SPO2 here, "Pulse Oximeter" in the master list.  Same device.
#   AE   ESU here, "Diathermy,Cryo" there.
#   CE   Sphygmomanometer here, "Mercury Meter" there.
#   BB   Ultrasound here, "Utrasound,Abdomen" there.
#
# Changing any device_name rewrites that text in every register rebuilt
# afterwards, so treat it as a data decision rather than a typo fix.
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
    "BB": {"device_name": "Ultrasound",
           "cells": form(17, "H30", col="F", val="L", extra={"S.N2": "L21"}),
           "alt_cells": [form(18, "H31", col="F", val="L", extra={"S.N2": "L22"}),
                         form(16, "H29", col="F", val="L", extra={"S.N2": "L20"})],
           # Some 2024 forms print "Date of receipt" one row lower (F16), with
           # the issue date at F14 and F15 empty. Only read when F15 holds no
           # date — see _pick_field_cells.
           "field_alternates": {"Date": ["F16", "F14"]}},
    "BF": {"device_name": "X-ray",      "cells": form(18, "J27", extra={"S.N2": "K20", "Location": "K22"})},
    # Was "X-ray ()" — an unfinished edit this file carried a TODO about. The
    # site's master code list settles it: CA is the dental x-ray. It shares BF's
    # layout, which is why it was copied from it in the first place.
    "CA": {"device_name": "Dental X-Ray", "cells": form(18, "J27", extra={"S.N2": "K20", "Location": "K22"})},
    "EA": {"device_name": "C-Arm",      "cells": form(18, "J27", extra={"S.N2": "K20", "Location": "K22"})},

    # ── Forms using the D/J columns ───────────────────────────────────────────
    "AF": {"device_name": "ECG",          "cells": form(32, "F41", col="D", val="J")},
    # This file's own TODO listed Phototherapy as unverified, and the import
    # bore that out: ~670 files no layout could read. Two are in circulation —
    # the same rows with an extra line above the date, and one shifted seven
    # rows down. Both read off the printed labels.
    "AL": {
        "device_name": "Phototherapy",
        "cells": form(28, "F37", col="D", val="J"),
        "alt_cells": [
            # There used to be a row-28 entry ahead of this one, differing from
            # the primary only in date_gap and Status. Its Model and S.N were
            # the primary's, so it could never be reached — see
            # test_an_alternate_that_repeats_the_primary_identity_is_never_added.
            #
            # Status was J43, which on these forms is the "Safety:" caption —
            # written into the register as the device's verdict on 183 of 300
            # sampled files. The answer box is F44/G44; C44/D44 beside it is
            # the printed legend, not an answer.
            form(35, "F44", col="D", val="J", date_gap=4),
            # Two more, from the certificates' formulas: 58 forms eight rows
            # down, 4 forms thirty-one rows down.
            form(36, "F45", col="D", val="J"),
            form(59, "F68", col="D", val="J"),
        ],
    },

    # ── Standard E/K forms ────────────────────────────────────────────────────
    "AC": {"device_name": "Defibrillator",          "cells": form(15, "G24")},
    "AA": {"device_name": "Anesthesia",             "cells": form(17, "G33")},
    "BP": {"device_name": "Balance",                "cells": form(18, "G30")},
    "EO": {"device_name": "Pipet",                  "cells": form(18, "H29")},
    # SPO2 and "Pulse Oximeter" (the master list's wording) are the same device.
    "AH": {"device_name": "SPO2",                   "cells": form(14, "G27")},
    "EE": {"device_name": "Flowmeter",              "cells": form(17, "G31")},
    "GP": {"device_name": "Holter machines",        "cells": form(18, "G26")},
    # The master code list calls it a light source; the engineers call it the
    # OR light, and that is the wording the register carries.
    # ~400 of the 1,005 files use the ordinary E/K block instead of row 15.
    # The alternate's K22 Status is correct — checked, 119/119 real statuses —
    # unlike the column changes in AL, BZ, GC, FW and DU, which land on captions.
    "DV": {
        "device_name": "OR Light",
        "cells": form(15, "G24"),
        "alt_cells": [form(18, "K22")],
    },
    # On the .xls variant K25 is the "Tested by" box — its "JTE-40" was being
    # written as the verdict. The form captions K22 as Status there.
    "AS": {"device_name": "Centrifuge",             "cells": form(18, "K25"),
           "field_alternates": {"Status": ["K22"]}},
    "AJ": {"device_name": "Suction",                "cells": form(23, "G32")},
    # Status was G33, which is right for the older template only: of 250 forms
    # sampled, 148 put the box at G31 and 69 at G33. The identity block does
    # NOT move with it, so read_best never falls through — the majority simply
    # read blank until the box was found by its own label.
    "AM": {"device_name": "Ventilator",             "cells": form(17, "G31")},
    "FG": {"device_name": "ACT",                    "cells": form(18, "H32"),
           "alt_cells": [form(17, "H31"), form(19, "H33")]},
    # Status was H32 — a row this form does not even reach — and read blank on
    # every one of 300 files sampled. The box is K22, labelled "Status:" at I22;
    # it now reads on 771 of 937. The alternate's H29 was wrong the same way:
    # that layout labels F27/G27 and answers at F29/G29.
    "DG": {"device_name": "CBC Analyzer",           "cells": form(18, "K22"),
           "alt_cells": [form(15, "G29")]},
    "AU": {"device_name": "Chemistry analyzer",     "cells": form(18, "K22")},
    "AX": {"device_name": "Lab Incubator",          "cells": form(18, "H32")},
    "EY": {"device_name": "Freezer",                "cells": form(18, "H32")},
    "EP": {"device_name": "Refrigerator",           "cells": form(18, "H32"),
           "alt_cells": [form(19, "H33")]},
    "DL": {"device_name": "Sealing Machine",        "cells": form(18, "K22")},
    "BV": {"device_name": "Blood gas analyzer",     "cells": form(18, "K22"),
           "alt_cells": [form(15, "K19")]},
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
        "alt_cells": [form(28, "G37"), form(26, "G35"), form(25, "G34")],
    },
    "DO": {"device_name": "O2 conc",                "cells": form(17, "G31")},
    "FJ": {"device_name": "OR Table",               "cells": form(18, "K22")},
    "AB": {"device_name": "Vaporizer",              "cells": form(18, "K22")},
    "AD": {"device_name": "Pacemaker",              "cells": form(22, "G31")},
    "AV": {"device_name": "Elisa reader",           "cells": form(18, "K22")},
    "FE": {"device_name": "Nebulizer",              "cells": form(17, "K31"),
           "alt_cells": [form(18, "K32")]},
    # Was "Infrared", which is what CK is. The master code list has AO as the
    # patient thermometer, and two codes sharing one name was the clue that one
    # of them had been copied from the other.
    "AO": {"device_name": "Thermometer, patient",   "cells": form(18, "H32")},
    # A second layout sits one row up with an extra line above the date. Its
    # status cell was guessed at K23, following the +4 offset forms of this
    # shape usually use. The guess was wrong: K23 read blank on all 132 sampled
    # files that use this layout, and the box is at F32/G32.
    "GC": {
        "device_name": "Portable Data Logger",
        "cells": form(20, "G33"),
        "alt_cells": [form(19, "G32", date_gap=4)],
    },
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
    "AQ": {"device_name": "Water Bath",             "cells": form(18, "H32"),
           "alt_cells": [form(17, "H31")]},
    "EV": {"device_name": "Blood Mixer",            "cells": form(18, "K22"),
           "alt_cells": [form(15, "K19")]},
    "GD": {"device_name": "Protien Analyzer",       "cells": form(18, "K22")},
    "AR": {"device_name": "Electrolyte Analyzer",   "cells": form(18, "K22"),
           "alt_cells": [form(15, "K19")]},

    # ── Added from the archive ────────────────────────────────────────────────
    # Codes that appear in four years of forms but were never in this table, so
    # every one of their files was reported as an unrecognised device. Each map
    # below was derived from the form's own printed labels ("Model:", "Serial
    # No.:", "Location") across four files spread over different sites and
    # years — never copied from a similar device.
    # "Final"-type workbooks state the verdict as a Conclusion sentence, whose
    # row moves (A119-A132). Identity: this map, then the cover page, then the
    # Word certificate beside the workbook — see calist.read_best.
    "BM": {"device_name": "Hemodialysis Machine",   "cells": form(18, "K22"),
           "conclusion_tab": r"^\s*final\s*$"},
    # The device block sits low on the "data entry" tab, in D/J, and its height
    # varies — so its identity is found by the form's printed labels (the map
    # below fits almost none of them, deliberately left as it was). Mapping
    # the block's rows directly was tried and read test readings ("4.1",
    # "6.9") on 105 fields. What the labels cannot place is the Date, which
    # sits two rows above the Model wherever the block is — hence
    # label_offsets (see calist._best_layout).
    "BN": {"device_name": "Therapeutic Ultrasound", "cells": form(18, "K22"),
           "label_offsets": {"Date": ("Model", -2)}},
    "CN": {"device_name": "Microwave",              "cells": form(18, "K22")},
    "GE": {"device_name": "Temperature Calibration Tester", "cells": form(18, "K22"),
           "alt_cells": [form(17, "K21")]},
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
    # The syringe-pump form drifted by a row in each direction over the years.
    # The original map is row 26 and stays primary; a four-year import found
    # ~820 files it could not read, and probing those showed rows 25 and 27 in
    # circulation. Confirmed against the printed labels, not inferred.
    # The status box is in column G on every one of these layouts. The alternates
    # named K, which is where "Safety:" and "Syringe brand:" are printed — so a
    # form that fell through to an alternate had a caption recorded as its
    # verdict. ~7.6% of the 18,219 syringe forms use alt 0, and its K35 read
    # 'Syringe brand:' on 50 of 61 and 'Safety:' on 7; the real box is G34.
    #
    # A fourth entry used to sit here with the primary's own identity cells and
    # only a different Status. read_best tries an alternate only when the whole
    # record is implausible, so an entry whose identity matches the primary can
    # never be reached — it was dead. A Status box that moves on its own is now
    # found by the form's printed label instead.
    "BZ": {
        "device_name": "Syringe",
        # Date is E24, the calibration date: 90 forms read the wrong day from
        # E22. Four forms print a date only at E22, so it stays a candidate.
        "cells": form(26, "G35"),
        "field_alternates": {"Date": ["E22"]},
        "alt_cells": [
            form(27, "G34", date_gap=4),
            form(25, "G35", date_gap=4),
        ],
    },
    "CE": {"device_name": "Sphygmomanometer",       "cells": form(47, "H59", date_gap=4)},
    "CB": {"device_name": "Digital blood pressure", "cells": form(18, "G26", date_gap=4)},
    # Date is E13, the calibration date; E11 matched it only by coincidence
    # (886 forms) and was wrong on 5. Proven by value on every form: 0 broken.
    "AE": {"device_name": "ESU",                    "cells": form(15, "G24"),
           "field_alternates": {"Date": ["E11"]}},
    "BL": {"device_name": "Autoclave",              "cells": form(18, "K22", date_gap=4)},
    "AN": {"device_name": "Thermo",                 "cells": form(18, "H32", date_gap=4),
           "alt_cells": [form(19, "H33", date_gap=4), form(20, "H34", date_gap=4)]},
    "EC": {"device_name": "Laminar flow",           "cells": form(18, "F31", date_gap=4)},
    "CK": {"device_name": "Infrared lamp",          "cells": form(18, "F31", date_gap=4)},

    # ── Odd one out ───────────────────────────────────────────────────────────
    # NOTE: Location is K19. Every other standard form puts it 2 rows below the
    # S.N (which is K18 here), i.e. K20 — this looks like a typo worth checking
    # against the actual Lab Oven form.
    # Location is K20 like every other standard form (68 forms fixed); two B14
    # forms print it at K19, so that stays a candidate the caption can choose.
    "EU": {"device_name": "Lab Oven", "cells": form(18, "H32"),
           "alt_cells": [form(17, "H31"), form(19, "H33")],
           "field_alternates": {"Location": ["K19"]}},

    # ── Devices the master code list names but nobody had mapped ──────────────
    #
    # 29 codes, 190 files across the 2025 and 2026 rounds, every one of which
    # imported blank. Each map below was read off the forms' own printed labels
    # with calist.locate_by_labels — never copied from a device that looked
    # similar, which is what produced the AI map that reported "ICU" as a serial
    # for years. The count after each name is how many archive files agreed.
    #
    # These forms put the Date four rows above the Model rather than two, hence
    # date_gap=4 throughout.

    # The lab bench: one shared layout, Status bottom-right at K22.
    "FA": {"device_name": "Elisa Washer",           "cells": form(18, "K22", date_gap=4)},
    "EQ": {"device_name": "Urine Analyzer",         "cells": form(18, "K22", date_gap=4)},
    "FV": {"device_name": "Cardiac Enzyme Analyzer", "cells": form(18, "K22", date_gap=4)},
    "FM": {"device_name": "PCR Rotor",              "cells": form(18, "K22", date_gap=4)},
    "FR": {"device_name": "Sodium & Potassium Analyzer", "cells": form(18, "K22", date_gap=4)},
    "FF": {"device_name": "EEG",                    "cells": form(18, "K22", date_gap=4)},
    "DE": {"device_name": "Colony Counter",         "cells": form(18, "K22", date_gap=4)},
    "FD": {"device_name": "Drugs Analyzer",         "cells": form(18, "K22", date_gap=4)},
    "DB": {"device_name": "Hot Plate",              "cells": form(18, "K22", date_gap=4)},
    "CZ": {"device_name": "Mixture Device",         "cells": form(18, "K22", date_gap=4)},
    "GM": {"device_name": "Corona Virus Analyzer",  "cells": form(18, "K22", date_gap=4)},
    "BQ": {"device_name": "Flatbed Platelet Agitator", "cells": form(18, "K22", date_gap=4)},
    "EZ": {"device_name": "Non-invasive Hemodynamic Monitor",
           "cells": form(18, "K22", date_gap=4)},

    # Physiotherapy, same sheet as the lab bench.
    "FU": {"device_name": "Joint Mobiliser",        "cells": form(18, "K22", date_gap=4)},
    "FT": {"device_name": "Biofeedback",            "cells": form(18, "K22", date_gap=4)},
    "CP": {"device_name": "Vertebral Column Stretcher", "cells": form(18, "K22", date_gap=4)},

    # Two forms in circulation for these; the older one sits three rows higher.
    # On that older layout J27 is the "Contact Person Name:" caption, not a
    # status — checked on every file of each that uses it. The box is at F29/G29.
    "FW": {"device_name": "Blood Culture System",   "cells": form(18, "K22", date_gap=4),
           "alt_cells": [form(15, "G29", date_gap=4)]},
    "DU": {"device_name": "Immunoassay Analyzer",   "cells": form(18, "K22", date_gap=4),
           "alt_cells": [form(15, "G29", date_gap=4)]},
    # AY is the same pair of layouts the other way round, and its K22 alternate
    # is correct — 40 of 48 real statuses. Left alone deliberately.
    "AY": {"device_name": "Virus & PCR Analyzer",   "cells": form(15, "J27", date_gap=4),
           "alt_cells": [form(18, "K22", date_gap=4)]},

    # Imaging. These sheets carry no Status field at all — they end at
    # "Tested By" and a comment box — so Status is mapped to nothing rather
    # than to whatever happens to sit at a borrowed coordinate.
    "BW": {"device_name": "CT",                     "cells": form(17, "", date_gap=4)},
    "BX": {"device_name": "MRI",                    "cells": form(17, "", date_gap=4)},
    "FP": {"device_name": "Dexa Scan",              "cells": form(17, "", date_gap=4)},
    "DW": {"device_name": "Heater Air Mattress",    "cells": form(17, "", date_gap=4)},

    # Respiratory: Status is mid-sheet at G34.
    "DS": {"device_name": "Spirometer",             "cells": form(17, "G34", date_gap=4)},
    "GH": {"device_name": "Bipap",                  "cells": form(17, "G34", date_gap=4)},

    # The D/J column pair, like the ECG and Phototherapy forms.
    "DX": {"device_name": "Fetal Doppler",
           "cells": form(27, "F36", col="D", val="J", date_gap=4)},
    "BC": {"device_name": "Ultrasound (Eye)",
           "cells": form(28, "F37", col="D", val="J", date_gap=4)},

    # Manufacturer and Location sit four rows below the Model here, not two.
    "CD": {"device_name": "Endoscopic Set",
           "cells": form(18, "K26",
                         extra={"Manufacturer": "E22", "Location": "K22"})},
    # Location two rows lower again, as on the other imaging-suite forms.
    "EN": {"device_name": "Catheter Lab",
           "cells": form(18, "J27", date_gap=4, extra={"Location": "K22"})},

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
            # Model sits *below* Manufacturer on this form, as it does on the
            # older one below — the standard E/K orientation is inverted here.
            # This read D72, which is the "Next calibration:" date, so the
            # register showed a date in the Model column. Confirmed against the
            # printed labels in column B: "Manufacturer:" r74, "Model:" r76.
            #
            # The alternate could not rescue it: D72 and L72 both hold values,
            # so the record looked plausible and read_best stopped at the primary.
            "Model": "D76",
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
        # The 2025 round prints the same block two rows higher.
        "alt_cells": [{
            "Manufacturer": "D64",
            "Model": "D66",
            "S.N": "L58",
            "Location": "D68",
            "Date": "D60",
            "Status": "F47",
        }],
    },

    # "FV": {"device_name": "Endoscopy", "cells": form(18, "K22")},  # unverified

    # ── Mapped in the 2026 archive audit ──────────────────────────────────────
    # High Flow Nasal Cannula. Three forms use the standard block; the 2023
    # batch keeps its only device block on a tab called "cert", reached by the
    # label search once "Equipment Data" counts as the device's heading.
    "FC": {"device_name": "High Flow Nasal Cannula", "cells": form(18, "K22")},
    # ICU electrical bed: the standard Data-entry block on every form.
    "GO": {"device_name": "ICU Bed",                 "cells": form(18, "K22")},
    # Mammography: the workbook is the vendor's QC template, whose header is
    # boilerplate on every file ("GE / Alpha st / Gona Hospital", 2012) and
    # whose "Model" caption belongs to the X-ray TUBE. The device is only in
    # the same-named Word certificate beside it, which carries no verdict.
    "BD": {"device_name": "Mammography", "source": "word",
           "cells": {"Manufacturer": "", "Model": "", "S.N": "", "Location": "",
                     "Date": "", "Status": ""}},
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

# "VAGH" is "AGH" with a stray V typed in front — 12 files at one site, confirmed
# by the owner. The second row gets its own code_replace because AGH's
# ("AGH" -> "AGCB") would rewrite "VAGH001" to "VAGCB001".
#
# The cell map used to be shared with AGH by reference, to stop the two drifting.
# It cannot be: these forms are filled in on the *vital signs* template, not the
# patient-monitor one. All 12 carry their verdicts at G38/J38 — VAH's cells —
# and AGH's D39/J39 are empty on every one, so both statuses read blank.
# The identity block is identical on the two templates (E18/K18/E20/K20), which
# is why nothing ever flagged it: the record looked perfectly plausible.
#
# The alternate that used to sit here named the same Model and S.N as the
# primary and differed only in Date and Status. read_best tries an alternate
# only when the whole record is implausible, so it could never be reached.
#
# Note for the owner: the model on these files is a Contec CMS 51000, a vital
# signs monitor. The device_name and the NIBP sub-row are left as they are,
# because changing them rewrites the text in every register built afterwards.
DEVICE_CONFIGS["VAGH"] = {
    "device_name": DEVICE_CONFIGS["AGH"]["device_name"],
    "cells": form(18, "G38", extra={"Status2": "J38"}),
    "second_row": {"device_name": "NIBP", "code_replace": ("VAGH", "AGCB")},
}

# "GJAF" is the aortic balloon pump — a GJ, confirmed by the owner.
#
# Its form prints "ECG Performance Test" near the top, which is not the device:
# an intra-aortic balloon pump inflates in time with the heart and triggers off
# the ECG signal, so testing that trigger is part of testing the pump. Reading
# that heading as the identity gets you "ECG" (code AF), which is wrong, and the
# cell layout is not AF's either — AF puts the Date at D30 and the Status at F41.
#
# The two layouts below were read off the files themselves: one anchored at row
# 32 and a later re-lay-out at row 60.
DEVICE_CONFIGS["GJAF"] = {
    "device_name": "Aortic balloon",
    "cells": form(32, "J70", col="D", val="J", date_gap=4),
    "alt_cells": [form(60, "J70", col="D", val="J", date_gap=4)],
}

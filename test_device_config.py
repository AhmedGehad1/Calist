"""The cell layout Calist reads is the layout the app writes.

The Calystra app files certificates for the balance and the patient monitor,
and has templates bundled for three more types. It builds their identity cells
with ``IdentityCells.standard`` in ``lib/models/device_spec.dart``, a Dart copy
of ``form()``. Nothing ties the two together at build time, so the output is
pinned here — and the app's ``test/device_spec_test.dart`` pins the same five
entries from its side, the way ``filed_blob_name`` is pinned against the app's
``filedWorkbookPath``.
"""

import device_config

# Exactly what the app's IdentityCells.standard produces for each code.
APP_LAYOUT = {
    "BP": {
        "Manufacturer": "E20",
        "Model": "E18",
        "S.N": "K18",
        "Location": "K20",
        "Date": "E16",
        "Status": "G30",
    },
    "AGH": {
        "Manufacturer": "E20",
        "Model": "E18",
        "S.N": "K18",
        "Location": "K20",
        "Date": "E16",
        "Status": "D39",
    },
    "CB": {
        "Manufacturer": "E20",
        "Model": "E18",
        "S.N": "K18",
        "Location": "K20",
        "Date": "E14",
        "Status": "G26",
    },
    "CE": {
        "Manufacturer": "E49",
        "Model": "E47",
        "S.N": "K47",
        "Location": "K49",
        "Date": "E43",
        "Status": "H59",
    },
    "AH": {
        "Manufacturer": "E16",
        "Model": "E14",
        "S.N": "K14",
        "Location": "K16",
        "Date": "E12",
        "Status": "G27",
    },
}


def test_form_matches_the_app():
    for code, expected in APP_LAYOUT.items():
        cells = device_config.DEVICE_CONFIGS[code]["cells"]
        # Extras such as the monitor's Status2 are Calist's alone; the app
        # pins the six cells form() itself lays out.
        laid_out = {key: cells[key] for key in expected}
        assert laid_out == expected, code

import os
import sys
import unittest
from src.config.config_provider import ConfigProvider
from src.word.table_handler import (
    set_paragraph_spacing,
    set_table_column_widths,
    set_tables_autofit_to_window,
    set_landscape_for_all_sections,
    copy_table_rows_excluding_header_into_table_with_id,
    copy_excel_rows_excluding_header_into_table_with_id,
)
from src.word.placeholder_replacer import replace_placeholders_using_config
from src.config.constants import (
    DOCX_EXTENSION,
    XLSX_EXTENSION,
    APP_DATA_FOLDER_NAME,
    ConfigKeys,
    WordTableDefaults,
    CONFIG_FILE_NAME,
)
from src.validation.docx_verifier import verify_normalized_protocol, _table_matches_headers


class _FakeCell:
    def __init__(self, text: str):
        self.text = text


class _FakeRow:
    def __init__(self, cells):
        self.cells = [_FakeCell(c) for c in cells]


class _FakeTable:
    def __init__(self, rows):
        self.rows = [_FakeRow(r) for r in rows]


class TestProtocolNormalization(unittest.TestCase):
    """Unit test for normalizing Word protocols and replacing placeholders."""

    def setUp(self):
        """Load configuration and set file paths for the test."""
        self.config = ConfigProvider.load_config_json()
        self.exported_word = self.config[ConfigKeys.EXPORTED_STD]
        self.template_ready_word = self.config[ConfigKeys.TEMPLATE_PROTOCOL]
        self.output_word = self.config[ConfigKeys.NORMALIZED_PROTOCOL]

        # Ensure the output file has a .docx extension
        if not self.output_word.endswith(DOCX_EXTENSION):
            self.output_word += DOCX_EXTENSION

    def test_document_normalization(self):
        """
        Validate that Word tables are normalized and placeholders are replaced.

        Steps:
        1. Set all sections to landscape orientation.
        2. Autofit all tables to window width.
        3. Copy rows from exported Word into the template (excluding headers).
        4. Set column widths for all tables.
        5. Adjust paragraph spacing.
        6. Replace placeholders using configuration.
        7. Verify complete document integrity.
        """
        required_inputs = [self.exported_word, self.template_ready_word]
        missing_files = [p for p in required_inputs if not os.path.exists(p)]
        if missing_files:
            self.skipTest(
                "Missing required input files from config.json: "
                + ", ".join(missing_files)
            )
        # Set landscape layout for all sections
        source_doc_for_layout = (
            self.template_ready_word
            if self.exported_word.lower().endswith(XLSX_EXTENSION)
            else self.exported_word
        )
        set_landscape_for_all_sections(source_doc_for_layout, self.output_word)

        # Make tables autofit to window
        set_tables_autofit_to_window(source_doc_for_layout, self.output_word)

        # Copy rows (excluding header) into template table
        if self.exported_word.lower().endswith(XLSX_EXTENSION):
            copy_excel_rows_excluding_header_into_table_with_id(
                self.exported_word,
                self.template_ready_word,
                self.output_word,
            )
        else:
            copy_table_rows_excluding_header_into_table_with_id(
                self.exported_word,
                self.template_ready_word,
                self.output_word,
            )

        # Adjust table column widths
        set_table_column_widths(
            self.output_word,
            self.output_word,
            widths_cm=WordTableDefaults.DEFAULT_COLUMN_WIDTHS_CM
        )

        # Adjust paragraph spacing in the document
        set_paragraph_spacing(self.output_word, self.output_word)

        # Replace placeholders with values from config
        replace_placeholders_using_config(self.output_word, self.output_word)

        # CI-grade verification of the final normalized protocol.
        # Any deviation in content, structure, formatting, placeholders,
        # or template preservation will cause this test to fail with a
        # precise diagnostic message.
        verify_normalized_protocol(
            exported_std_path=self.exported_word,
            template_protocol_path=self.template_ready_word,
            normalized_protocol_path=self.output_word,
        )

    def test_target_header_matching_allows_prefix_for_wider_tables(self):
        table = _FakeTable(rows=[["ID", "Step", "Expected result"]])
        self.assertTrue(_table_matches_headers(table, ["ID"]))

    def test_target_header_matching_allows_expected_text_contained_in_cell(self):
        table = _FakeTable(rows=[["ID #", "Name"]])
        self.assertTrue(_table_matches_headers(table, ["ID"]))

    def test_target_header_matching_requires_position_match(self):
        table = _FakeTable(rows=[["Name", "ID"]])
        self.assertFalse(_table_matches_headers(table, ["ID"]))


if __name__ == "__main__":
    from src.config.logging_config import setup_logging, get_logger
    setup_logging()
    log = get_logger(__name__)

    # Define appdata path for storing config
    appdata_folder = os.path.join(
        os.environ.get('APPDATA', os.path.expanduser('~\\AppData\\Roaming')),
        APP_DATA_FOLDER_NAME
    )
    config_path = os.path.join(appdata_folder, CONFIG_FILE_NAME)

    log.info("Started document normalization. Config path: %s", config_path)

    # Load config
    config = ConfigProvider.load_config_json(config_path)
    if config:
        exported = config.get(ConfigKeys.EXPORTED_STD, "")
        output = config.get(ConfigKeys.NORMALIZED_PROTOCOL, "")
        log.info("Inputs: Exported_STD=%s, Normalized_Protocol=%s", exported, output)

    # Create test suite and run
    suite = unittest.TestSuite()
    suite.addTest(TestProtocolNormalization('test_document_normalization'))
    result = unittest.TextTestRunner(verbosity=2).run(suite)

    if result.wasSuccessful():
        log.info("Output: Document normalization completed successfully.")
        sys.exit(0)
    else:
        failures = "; ".join(str(f[1]) for f in result.failures) if result.failures else "unknown"
        log.info("Output: Document normalization failed. %s", failures)
        sys.exit(1)

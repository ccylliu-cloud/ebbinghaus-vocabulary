"""Optional native print regression: VOCAB_SOFFICE=/path/to/soffice python this.py.

Uses synthetic long definitions, all four schemas, both group sizes, and every
available exporter. Requires LibreOffice with a CJK font configured. No files
are published; temporary workbooks and native print PDFs are removed on exit.
"""
import os
from pathlib import Path
import re
import subprocess
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pypdf import PdfReader
from unittest.mock import patch
import vocab as v


@unittest.skipUnless(os.environ.get('VOCAB_SOFFICE'), 'set VOCAB_SOFFICE for native print verification')
class OfficeLayoutTests(unittest.TestCase):
    def test_native_print_page_count_and_complete_text(self):
        for daily in (5, 10):
            data = []
            for i in range(100):
                word = {'english': f'sample{i:03d}', 'chinese': '示例词'}
                if i % daily == 0:
                    word['chinese'] = '（用于书信开头，向收信人表达礼貌及亲近关系）亲爱的朋友'
                schema = (i // daily) % 4
                if schema in (1, 3): word['ipa'] = '/ˈsɑːmpl/'
                if schema in (2, 3): word['pos'] = 'n.'
                data.append(word)
            layout = v.master_layout(v.schedule(data, daily)['groups'])
            self.assertGreater(len(layout), 1)
            canonical_columns = None
            for engine in ('xlsxwriter', 'artifact-tool'):
                if engine == 'artifact-tool' and not ((v.RUNTIME / 'node/node_modules/@oai/artifact-tool').exists() or os.environ.get('VOCAB_ARTIFACT_MODULE')):
                    continue
                with self.subTest(daily=daily, engine=engine), tempfile.TemporaryDirectory() as tmp:
                    root = Path(tmp)
                    xlsx = root / 'master.xlsx'
                    with patch.dict(os.environ, {'VOCAB_XLSX_ENGINE': engine}):
                        v.render_xlsx(layout, xlsx, root)
                    ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    with zipfile.ZipFile(xlsx) as z:
                        sheet = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
                    columns = [c.attrib for c in sheet.find('m:cols', ns)]
                    if canonical_columns is None: canonical_columns = columns
                    self.assertEqual(columns, canonical_columns)
                    result = subprocess.run([os.environ['VOCAB_SOFFICE'],
                        '-env:UserInstallation=' + (root / 'office-profile').as_uri(),
                        '--headless', '--convert-to', 'pdf', '--outdir', str(root), str(xlsx)],
                        capture_output=True, text=True, timeout=60)
                    self.assertEqual(result.returncode, 0, result.stderr)
                    pdf = PdfReader(root / 'master.pdf')
                    self.assertEqual(len(pdf.pages), len(layout))
                    for actual, expected in zip(pdf.pages, layout):
                        text = re.sub(r'\s+', '', actual.extract_text())
                        for block in expected:
                            self.assertIn(f"第{block['group']}组", text)
                            first = (block['group'] - 1) * daily
                            for word in data[first:first + daily]:
                                self.assertIn(word['english'], text)
                            self.assertIn(data[first]['chinese'], text)
                        self.assertAlmostEqual(float(actual.mediabox.width), 595.28, delta=1)
                        self.assertAlmostEqual(float(actual.mediabox.height), 841.89, delta=1)


if __name__ == '__main__':
    unittest.main(verbosity=2)

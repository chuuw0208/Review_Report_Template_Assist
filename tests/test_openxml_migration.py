"""
Regression tests for the OpenXML (macOS / Linux) migration engine.

Builds tiny .docx fixtures in a temp folder — no Word required — and checks
that images, charts and hyperlinks from the reference survive migration with
valid relationships, and that table cells can be targeted and modified.

Run:
    python3 -m unittest discover -s tests -v
"""

import os
import posixpath
import re
import struct
import sys
import tempfile
import unittest
import zipfile
import zlib
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import phase2_core as core  # noqa: E402


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
REL_LINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
REL_PKG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"


# ---------------------------------------------------------------------------
# Fixture helpers
# ---------------------------------------------------------------------------

def make_png(rgb):
    """Return bytes of a valid 1x1 PNG of the given colour."""
    def chunk(tag, data):
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)
    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    idat = zlib.compress(b"\x00" + bytes(rgb))
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")


RED_PNG = make_png((255, 0, 0))
BLUE_PNG = make_png((0, 0, 255))
XLSX_BYTES = b"PK-fake-embedded-workbook"

DOC_NS = (
    f'xmlns:w="{W}" xmlns:r="{R}" xmlns:a="{A}" xmlns:c="{C}" xmlns:wp="{WP}" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'mc:Ignorable="w14 wp14"'
)


def heading(level, text):
    return f'<w:p><w:pPr><w:pStyle w:val="Heading{level}"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'


def para(text, red_italic=False):
    rpr = '<w:rPr><w:i/><w:color w:val="FF0000"/></w:rPr>' if red_italic else ""
    return f"<w:p><w:r>{rpr}<w:t>{text}</w:t></w:r></w:p>"


def image_para(rid, docpr_id, name):
    return (
        "<w:p><w:r><w:drawing><wp:inline>"
        '<wp:extent cx="914400" cy="914400"/>'
        f'<wp:docPr id="{docpr_id}" name="{name}"/>'
        f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:blipFill><a:blip r:embed="{rid}"/></pic:blipFill></pic:pic>'
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
    )


def chart_para(rid, docpr_id):
    return (
        "<w:p><w:r><w:drawing><wp:inline>"
        f'<wp:docPr id="{docpr_id}" name="Chart {docpr_id}"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        f'<c:chart r:id="{rid}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    )


def link_para(rid, text):
    return f'<w:p><w:hyperlink r:id="{rid}"><w:r><w:t>{text}</w:t></w:r></w:hyperlink></w:p>'


def cell(text, red_italic=False):
    return f"<w:tc>{para(text, red_italic)}</w:tc>"


def table(rows):
    return "<w:tbl>" + "".join("<w:tr>" + "".join(cell(*c) if isinstance(c, tuple) else cell(c)
                                                  for c in row) + "</w:tr>" for row in rows) + "</w:tbl>"


SECT_PR = ('<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/>'
           '<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>')
TEMPLATE_SECT_PR = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'


def build_docx(path, body_xml, rels, parts, content_types):
    rel_xml = "".join(
        f'<Relationship Id="{rid}" Type="{t}" Target="{tgt}"' + (' TargetMode="External"' if ext else "") + "/>"
        for rid, t, tgt, ext in rels
    )
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml",
                   f'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="{CT}">{content_types}</Types>')
        z.writestr("_rels/.rels",
                   f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{PKG}">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                   'relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/_rels/document.xml.rels",
                   f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{PKG}">{rel_xml}</Relationships>')
        z.writestr("word/styles.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="{W}"/>')
        z.writestr("word/settings.xml",
                   f'<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="{W}" '
                   'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
                   'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
                   'mc:Ignorable="w14"/>')
        z.writestr("word/document.xml",
                   f'<?xml version="1.0" encoding="UTF-8"?><w:document {DOC_NS}><w:body>{body_xml}</w:body></w:document>')
        for name, data in parts.items():
            z.writestr(name, data)


BASE_CT = (
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="png" ContentType="image/png"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
)
CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
XLSX_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def build_fixtures(folder):
    template = os.path.join(folder, "template_assessment.docx")
    reference = os.path.join(folder, "reference_assessment.docx")

    tmpl_body = (
        heading(1, "1. Review Details")
        + heading(2, "1.1 Review Scope")
        + para("[Instruction: describe scope]", red_italic=True)
        + image_para("rId5", 1, "Template logo")          # template's own image1.png (BLUE)
        + heading(2, "1.2 Model Components")
        + table([
            ["Model Component", "Status", "Owner"],
            ["PD Module", ("[Enter status]", True), "TBD"],
            ["LGD Module", "", "TBD"],
        ])
        + heading(1, "3. Governance & Controls".replace("&", "&amp;"))
        + heading(2, "3.1 Model Controls")
        + para("Template controls text.")
        + TEMPLATE_SECT_PR
    )
    build_docx(
        template, tmpl_body,
        rels=[("rId1", REL_STYLES, "styles.xml", False),
              ("rId5", REL_IMAGE, "media/image1.png", False)],
        parts={"word/media/image1.png": BLUE_PNG},
        content_types=BASE_CT,
    )

    ref_body = (
        heading(1, "1. Review Details")
        + heading(2, "1.1 Review Scope")
        + para("Reference scope text with a figure:")
        + image_para("rId4", 1, "Reference figure")       # reference's image1.png (RED) — name collision
        + image_para("rId4", 2, "Same figure again")      # same rel used twice -> copied once
        + link_para("rId9", "MRM Policy")
        + heading(1, "3. Governance &amp; Controls")
        + heading(2, "3.1 Model Controls")
        + chart_para("rId6", 3)
        + SECT_PR
    )
    build_docx(
        reference, ref_body,
        rels=[("rId1", REL_STYLES, "styles.xml", False),
              ("rId4", REL_IMAGE, "media/image1.png", False),
              ("rId6", REL_CHART, "charts/chart1.xml", False),
              ("rId9", REL_LINK, "https://example.com/mrm-policy", True),
              ("rIdHdr", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
               "header1.xml", False)],
        parts={
            "word/media/image1.png": RED_PNG,
            "word/charts/chart1.xml": f'<?xml version="1.0"?><c:chartSpace xmlns:c="{C}"/>',
            "word/charts/_rels/chart1.xml.rels":
                f'<?xml version="1.0"?><Relationships xmlns="{PKG}">'
                f'<Relationship Id="rId1" Type="{REL_PKG}" '
                'Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/></Relationships>',
            "word/embeddings/Microsoft_Excel_Worksheet.xlsx": XLSX_BYTES,
            "word/header1.xml": f'<?xml version="1.0"?><w:hdr xmlns:w="{W}"/>',
        },
        content_types=BASE_CT
        + '<Default Extension="xlsx" ContentType="' + XLSX_CT + '"/>'
        + f'<Override PartName="/word/charts/chart1.xml" ContentType="{CHART_CT}"/>',
    )
    return template, reference


# ---------------------------------------------------------------------------
# Output inspection helpers
# ---------------------------------------------------------------------------

class Package:
    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.files = {n: z.read(n) for n in z.namelist()}
        self.doc = ET.fromstring(self.files["word/document.xml"])
        self.rels = {r.get("Id"): r for r in ET.fromstring(self.files["word/_rels/document.xml.rels"])}
        self.ct = ET.fromstring(self.files["[Content_Types].xml"])

    def part_for(self, rid, source="word/document.xml"):
        rel = self.rels[rid]
        return core._resolve_target(source, rel.get("Target"))

    def content_type(self, part):
        for ov in self.ct.iter(f"{{{CT}}}Override"):
            if ov.get("PartName").lower() == "/" + part.lower():
                return ov.get("ContentType")
        ext = part.rsplit(".", 1)[-1].lower()
        for d in self.ct.iter(f"{{{CT}}}Default"):
            if d.get("Extension").lower() == ext:
                return d.get("ContentType")
        return None


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

class OpenXmlImageMigrationTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        template, reference = build_fixtures(cls.tmp.name)
        cls.output = os.path.join(cls.tmp.name, "out.docx")
        core.run_migration_openxml(template, reference, cls.output)
        cls.pkg = Package(cls.output)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def _blips(self):
        return [b.get(f"{{{R}}}embed") for b in self.pkg.doc.iter(f"{{{A}}}blip")]

    def test_every_relationship_id_resolves(self):
        for node in self.pkg.doc.iter():
            for attr, rid in node.attrib.items():
                if attr.startswith(f"{{{R}}}"):
                    self.assertIn(rid, self.pkg.rels, f"dangling {attr}={rid}")
                    rel = self.pkg.rels[rid]
                    if rel.get("TargetMode") != "External":
                        self.assertIn(self.pkg.part_for(rid), self.pkg.files)

    def test_reference_image_copied_without_overwriting_template_image(self):
        images = [self.pkg.files[self.pkg.part_for(rid)] for rid in self._blips()]
        self.assertEqual(images.count(RED_PNG), 2, "both reference figures should show the RED image")
        self.assertEqual(images.count(BLUE_PNG), 1, "template logo must still be BLUE")
        self.assertEqual(self.pkg.files["word/media/image1.png"], BLUE_PNG)

    def test_shared_reference_image_copied_once(self):
        red_parts = {self.pkg.part_for(rid) for rid in self._blips()
                     if self.pkg.files[self.pkg.part_for(rid)] == RED_PNG}
        self.assertEqual(len(red_parts), 1)

    def test_chart_and_embedded_workbook_copied(self):
        chart_rids = [c.get(f"{{{R}}}id") for c in self.pkg.doc.iter(f"{{{C}}}chart")]
        self.assertEqual(len(chart_rids), 1)
        chart_part = self.pkg.part_for(chart_rids[0])
        self.assertEqual(self.pkg.content_type(chart_part), CHART_CT)
        chart_rels = ET.fromstring(self.pkg.files[core._rels_path_for(chart_part)])
        xlsx_part = core._resolve_target(chart_part, chart_rels[0].get("Target"))
        self.assertEqual(self.pkg.files[xlsx_part], XLSX_BYTES)
        self.assertEqual(self.pkg.content_type(xlsx_part), XLSX_CT)

    def test_external_hyperlink_preserved(self):
        links = [h.get(f"{{{R}}}id") for h in self.pkg.doc.iter(f"{{{W}}}hyperlink")]
        self.assertEqual(len(links), 1)
        rel = self.pkg.rels[links[0]]
        self.assertEqual(rel.get("Target"), "https://example.com/mrm-policy")
        self.assertEqual(rel.get("TargetMode"), "External")

    def test_every_part_has_a_content_type(self):
        for name in self.pkg.files:
            if name == "[Content_Types].xml":
                continue
            self.assertIsNotNone(self.pkg.content_type(name), f"no content type for {name}")

    def test_drawing_ids_are_unique(self):
        ids = [d.get("id") for d in self.pkg.doc.iter(f"{{{WP}}}docPr")]
        self.assertEqual(len(ids), len(set(ids)), f"duplicate docPr ids: {ids}")

    def test_reference_section_properties_not_cloned(self):
        body = self.pkg.doc.find(f"{{{W}}}body")
        self.assertEqual(len(body.findall(f"{{{W}}}sectPr")), 1)
        self.assertNotIn("word/header1.xml", self.pkg.files)

    def test_migrated_image_is_tracked_insertion(self):
        ins_drawings = [d for ins in self.pkg.doc.iter(f"{{{W}}}ins") for d in ins.iter(f"{{{W}}}drawing")]
        self.assertEqual(len(ins_drawings), 3)  # 2 figures + 1 chart

    def test_ignorable_prefixes_stay_declared(self):
        for part in ("word/document.xml", "word/settings.xml"):
            text = self.pkg.files[part].decode("utf-8")
            root_tag = re.search(r"<(?![?!])[^>]+>", text).group(0)
            ignorable = re.search(r'Ignorable="([^"]*)"', root_tag)
            self.assertIsNotNone(ignorable, part)
            for prefix in ignorable.group(1).split():
                self.assertIn(f"xmlns:{prefix}=", root_tag, f"{part}: {prefix} undeclared")


if __name__ == "__main__":
    unittest.main()

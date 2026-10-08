"""
Sheet copier for PO Translation System.

Builds the translated workbook by copying sheets inside the .xlsx package
itself. The sheet XML and everything attached to it (pictures, shapes, charts,
comments, print settings) are carried over unchanged, only the texts of the
translated cells are replaced. openpyxl cannot be used for this step because
it drops shapes and drawings when it saves a workbook.
"""

import posixpath
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional
from xml.etree import ElementTree
from xml.sax.saxutils import escape, quoteattr, unescape

RELS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_PART = "[Content_Types].xml"
XML_ENTITIES = {"&quot;": '"', "&apos;": "'"}

# Parts that several sheets can point to. Everything else is copied per sheet.
SHARED_RELATIONSHIP_TYPES = ("/image", "/pivotCacheDefinition")

CELL_PATTERN = re.compile(r'<c\s([^>]*?)(?:/>|>(.*?)</c>)', re.DOTALL)
FORMULA_PATTERN = re.compile(r'(<f\b[^>]*>)(.*?)(</f>)', re.DOTALL)
SHEET_PATTERN = re.compile(r'<sheet\s[^>]*?/>')
DEFINED_NAME_PATTERN = re.compile(r'<definedName\s([^>]*)>(.*?)</definedName>', re.DOTALL)


@dataclass
class SheetCopy:
    """One sheet of the output workbook."""

    source_title: str            # sheet of the source workbook to copy
    title: str                   # title of the copy
    cell_values: Dict[str, str]  # new texts by cell reference, e.g. {"C9": "..."}
    title_map: Dict[str, str]    # source title -> title of the copy in the same language


def retarget_references(text: str, title_map: Dict[str, str]) -> str:
    """Point sheet references in a formula to the translated copies of the sheets."""
    if "!" not in text:
        return text
    for old_title, new_title in title_map.items():
        new_reference = "'" + new_title.replace("'", "''") + "'!"
        quoted_old = "'" + old_title.replace("'", "''") + "'!"
        text = text.replace(quoted_old, new_reference)
        # Sheet names without spaces may be written without quotes
        if re.fullmatch(r'[A-Za-z_][A-Za-z0-9_.]*', old_title):
            pattern = r"(?<![A-Za-z0-9_.'])" + re.escape(old_title) + "!"
            text = re.sub(pattern, lambda _: new_reference, text)
    return text


def write_sheet_copies(source_path: Path, output_path: Path, copies: List[SheetCopy]) -> None:
    """Write a workbook that contains only the given sheet copies, in the given order."""
    _Package(source_path).write(output_path, copies)


class _Package:
    """The parts of an .xlsx file, held in memory by their path inside the zip."""

    def __init__(self, source_path: Path) -> None:
        with zipfile.ZipFile(source_path) as archive:
            self.parts: Dict[str, bytes] = {
                name: archive.read(name) for name in archive.namelist() if not name.endswith("/")
            }
        self.workbook_part = self._find_related("", "/officeDocument")
        self.shared_strings_part = self._find_related(self.workbook_part, "/sharedStrings")
        self.new_strings: Dict[str, int] = {}
        self.content_type_overrides: List[str] = []

    def write(self, output_path: Path, copies: List[SheetCopy]) -> None:
        workbook_xml = self._text(self.workbook_part)
        sheets = [self._attributes(tag) for tag in SHEET_PATTERN.findall(workbook_xml)]
        id_attribute = next(name for name in sheets[0] if name.endswith("id") and name != "sheetId")
        relationships = self._read_rels(self.workbook_part)
        sheet_parts = {
            rel["Id"]: self._resolve(self.workbook_part, rel["Target"]) for rel in relationships
        }
        relationship_types = {rel["Id"]: rel["Type"] for rel in relationships}
        source_index = {sheet["name"]: index for index, sheet in enumerate(sheets)}

        # The source sheets and the calculation chain (it lists cells by sheet) are left out
        dropped_ids = {sheet[id_attribute] for sheet in sheets}
        relationships = [
            rel for rel in relationships
            if rel["Id"] not in dropped_ids and not rel["Type"].endswith("/calcChain")
        ]

        sheet_tags = []
        for index, sheet_copy in enumerate(copies):
            source = sheets[source_index[sheet_copy.source_title]]
            new_part = self._duplicate(sheet_parts[source[id_attribute]])
            self.parts[new_part] = self._edit_sheet(self._text(new_part), sheet_copy).encode("utf-8")

            rel_id = f"rIdCopy{index + 1}"
            relationships.append({
                "Id": rel_id,
                "Type": relationship_types[source[id_attribute]],
                "Target": self._relative(self.workbook_part, new_part),
            })
            state = source.get("state", "visible")
            # Some writers declare the relationship namespace on every sheet tag
            namespaces = "".join(
                f" {name}={quoteattr(value)}" for name, value in source.items() if name.startswith("xmlns")
            )
            sheet_tags.append({"name": sheet_copy.title, "state": state, "id": rel_id,
                               "namespaces": namespaces})

        # Excel refuses to open a workbook without a visible sheet
        if all(tag["state"] != "visible" for tag in sheet_tags):
            sheet_tags[0]["state"] = "visible"
        active = next(i for i, tag in enumerate(sheet_tags) if tag["state"] == "visible")

        sheets_xml = "".join(
            f'<sheet{tag["namespaces"]} name={quoteattr(tag["name"])} sheetId="{index + 1}"'
            + (f' state="{tag["state"]}"' if tag["state"] != "visible" else "")
            + f' {id_attribute}="{tag["id"]}"/>'
            for index, tag in enumerate(sheet_tags)
        )
        workbook_xml = re.sub(r'<sheets>.*?</sheets>', lambda _: f"<sheets>{sheets_xml}</sheets>",
                              workbook_xml, flags=re.DOTALL)
        workbook_xml = re.sub(r'\sactiveTab="\d+"', f' activeTab="{active}"', workbook_xml)
        workbook_xml = re.sub(r'\sfirstSheet="\d+"', "", workbook_xml)
        workbook_xml = self._copy_defined_names(workbook_xml, copies, source_index)

        self.parts[self.workbook_part] = workbook_xml.encode("utf-8")
        self._write_rels(self.workbook_part, relationships)
        self._save_shared_strings()
        self._remove_unused_parts()
        self._save_content_types()

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr(CONTENT_TYPES_PART, self.parts[CONTENT_TYPES_PART])
            for name, data in self.parts.items():
                if name != CONTENT_TYPES_PART:
                    archive.writestr(name, data)

    # Sheet content

    def _edit_sheet(self, sheet_xml: str, sheet_copy: SheetCopy) -> str:
        """Write the translated texts into a copied sheet and fix its references."""

        def replace_cell(match: "re.Match[str]") -> str:
            attributes = match.group(1)
            reference = re.search(r'\br="([A-Z]+\d+)"', attributes)
            if not reference or reference.group(1) not in sheet_copy.cell_values:
                return match.group(0)
            # The style attribute carries the font, border, fill and alignment of the cell
            style = re.search(r'\bs="\d+"', attributes)
            start = f'<c r="{reference.group(1)}"' + (f" {style.group(0)}" if style else "")
            text = sheet_copy.cell_values[reference.group(1)]
            if self.shared_strings_part:
                return f'{start} t="s"><v>{self._shared_string_index(text)}</v></c>'
            return f'{start} t="inlineStr"><is><t xml:space="preserve">{escape(text)}</t></is></c>'

        def replace_formula(match: "re.Match[str]") -> str:
            formula = retarget_references(unescape(match.group(2), XML_ENTITIES), sheet_copy.title_map)
            return match.group(1) + escape(formula) + match.group(3)

        sheet_xml = CELL_PATTERN.sub(replace_cell, sheet_xml)
        sheet_xml = FORMULA_PATTERN.sub(replace_formula, sheet_xml)
        # Several selected tabs would open the workbook in group editing mode
        return re.sub(r'\stabSelected="(1|true)"', "", sheet_xml)

    def _shared_string_index(self, text: str) -> int:
        if text not in self.new_strings:
            existing = len(re.findall(r'<si\b', self._text(self.shared_strings_part)))
            self.new_strings[text] = existing + len(self.new_strings)
        return self.new_strings[text]

    def _save_shared_strings(self) -> None:
        if not self.new_strings:
            return
        xml = self._text(self.shared_strings_part)
        items = "".join(
            f'<si><t xml:space="preserve">{escape(text)}</t></si>' for text in self.new_strings
        )
        # The counts are optional, so they are removed instead of recalculated
        xml = re.sub(r'\s(count|uniqueCount)="\d+"', "", xml, count=2)
        if "</sst>" in xml:
            xml = xml.replace("</sst>", items + "</sst>")
        else:
            xml = re.sub(r'<sst\b([^>]*?)/>', lambda m: f"<sst{m.group(1)}>{items}</sst>", xml)
        self.parts[self.shared_strings_part] = xml.encode("utf-8")

    def _copy_defined_names(self, workbook_xml: str, copies: List[SheetCopy],
                            source_index: Dict[str, int]) -> str:
        """Give every copy the print area and print titles of its source sheet."""
        if "<definedNames>" not in workbook_xml:
            return workbook_xml

        names = []
        for attributes, value in DEFINED_NAME_PATTERN.findall(workbook_xml):
            value = unescape(value, XML_ENTITIES)
            local = re.search(r'\slocalSheetId="(\d+)"', " " + attributes)
            if local is None:
                # A workbook level name follows the first language
                names.append((attributes, retarget_references(value, copies[0].title_map)))
                continue
            for index, sheet_copy in enumerate(copies):
                if source_index[sheet_copy.source_title] == int(local.group(1)):
                    new_attributes = re.sub(r'localSheetId="\d+"', f'localSheetId="{index}"', attributes)
                    names.append((new_attributes, retarget_references(value, sheet_copy.title_map)))

        # A name that still points to a removed sheet would be reported as an error
        removed = [
            title for title in source_index if title not in {c.title for c in copies}
        ]
        names = [
            (attributes, value) for attributes, value in names
            if not any(f"'{title}'!" in value or re.search(r"(?<![\w.'])" + re.escape(title) + "!", value)
                       for title in removed)
        ]
        names_xml = "".join(
            f"<definedName {attributes}>{escape(value)}</definedName>" for attributes, value in names
        )
        replacement = f"<definedNames>{names_xml}</definedNames>" if names else ""
        return re.sub(r'<definedNames>.*?</definedNames>', lambda _: replacement,
                      workbook_xml, flags=re.DOTALL)

    # Package parts

    def _duplicate(self, part: str) -> str:
        """Copy a part together with the parts that belong to it. Returns the new part path."""
        new_part = self._new_part_name(part)
        self.parts[new_part] = self.parts[part]
        content_type = self._content_type_override(part)
        if content_type:
            self.content_type_overrides.append(
                f'<Override PartName="/{new_part}" ContentType="{content_type}"/>'
            )

        relationships = self._read_rels(part)
        for rel in relationships:
            if rel.get("TargetMode") == "External" or rel["Type"].endswith(SHARED_RELATIONSHIP_TYPES):
                target = self._resolve(part, rel["Target"])
            else:
                target = self._duplicate(self._resolve(part, rel["Target"]))
                if rel["Type"].endswith("/table"):
                    self._make_table_unique(target)
            if rel.get("TargetMode") != "External":
                rel["Target"] = self._relative(new_part, target)
        if relationships:
            self._write_rels(new_part, relationships)
        return new_part

    def _new_part_name(self, part: str) -> str:
        stem, extension = posixpath.splitext(part)
        stem = stem.rstrip("0123456789")
        number = 1
        while f"{stem}{number}{extension}" in self.parts:
            number += 1
        return f"{stem}{number}{extension}"

    def _make_table_unique(self, part: str) -> None:
        """Table ids and names must be unique inside a workbook."""
        number = int(re.search(r'(\d+)\.\w+$', part).group(1))

        def rename(match: "re.Match[str]") -> str:
            tag = re.sub(r'\sid="\d+"', f' id="{number}"', match.group(0), count=1)
            return re.sub(r'\s(name|displayName)="([^"]*)"',
                          lambda m: f' {m.group(1)}="{m.group(2)}_{number}"', tag)

        xml = re.sub(r'<table\s[^>]*>', rename, self._text(part), count=1)
        self.parts[part] = xml.encode("utf-8")

    def _remove_unused_parts(self) -> None:
        """Drop the parts that no relationship leads to any more (the source sheets)."""
        used = {CONTENT_TYPES_PART}
        pending = [""]
        while pending:
            part = pending.pop()
            rels_part = self._rels_part(part)
            if rels_part not in self.parts:
                continue
            used.add(rels_part)
            for rel in self._read_rels(part):
                if rel.get("TargetMode") == "External":
                    continue
                target = self._resolve(part, rel["Target"])
                if target in self.parts and target not in used:
                    used.add(target)
                    pending.append(target)
        self.parts = {name: data for name, data in self.parts.items() if name in used}

    def _save_content_types(self) -> None:
        xml = self._text(CONTENT_TYPES_PART)
        xml = xml.replace("</Types>", "".join(self.content_type_overrides) + "</Types>")
        xml = re.sub(
            r'<Override\s[^>]*?PartName="/([^"]*)"[^>]*?/>',
            lambda m: m.group(0) if unescape(m.group(1), XML_ENTITIES) in self.parts else "",
            xml,
        )
        self.parts[CONTENT_TYPES_PART] = xml.encode("utf-8")

    def _content_type_override(self, part: str) -> Optional[str]:
        match = re.search(
            r'<Override\s[^>]*?PartName="/' + re.escape(escape(part)) + r'"[^>]*?/>',
            self._text(CONTENT_TYPES_PART),
        )
        return self._attributes(match.group(0), unescape_values=False)["ContentType"] if match else None

    # Relationships

    def _find_related(self, part: str, type_suffix: str) -> Optional[str]:
        for rel in self._read_rels(part):
            if rel["Type"].endswith(type_suffix):
                return self._resolve(part, rel["Target"])
        return None

    def _read_rels(self, part: str) -> List[Dict[str, str]]:
        rels_part = self._rels_part(part)
        if rels_part not in self.parts:
            return []
        root = ElementTree.fromstring(self.parts[rels_part])
        return [dict(element.attrib) for element in root.iter(f"{{{RELS_NAMESPACE}}}Relationship")]

    def _write_rels(self, part: str, relationships: List[Dict[str, str]]) -> None:
        items = "".join(
            "<Relationship " + " ".join(f"{name}={quoteattr(value)}" for name, value in rel.items()) + "/>"
            for rel in relationships
        )
        self.parts[self._rels_part(part)] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<Relationships xmlns="{RELS_NAMESPACE}">{items}</Relationships>'
        ).encode("utf-8")

    @staticmethod
    def _rels_part(part: str) -> str:
        directory, name = posixpath.split(part)
        return posixpath.join(directory, "_rels", f"{name}.rels")

    @staticmethod
    def _resolve(part: str, target: str) -> str:
        """Path of a relationship target inside the zip."""
        if target.startswith("/"):
            return target.lstrip("/")
        return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))

    @staticmethod
    def _relative(part: str, target: str) -> str:
        return posixpath.relpath(target, posixpath.dirname(part) or ".")

    # XML helpers

    def _text(self, part: str) -> str:
        return self.parts[part].decode("utf-8")

    @staticmethod
    def _attributes(tag: str, unescape_values: bool = True) -> Dict[str, str]:
        return {
            name: unescape(value, XML_ENTITIES) if unescape_values else value
            for name, value in re.findall(r'([\w:]+)="([^"]*)"', tag)
        }

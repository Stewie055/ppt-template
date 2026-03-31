from __future__ import annotations

import uuid
from typing import Optional, Union

from lxml import etree

from .adapter.pptx_adapter import PptxAdapter
from .exceptions import OperationError, ShapeOperationError


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
DEFAULT_SECTION_NAME = "Section 1"


def _tag(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


class PptOperations:
    def __init__(self, presentation, adapter: Optional[PptxAdapter] = None):
        self.presentation = presentation
        self.adapter = adapter or PptxAdapter()

    @classmethod
    def load(cls, template_path: Optional[str] = None, template_bytes: Optional[bytes] = None):
        adapter = PptxAdapter()
        return cls(adapter.load(template_path=template_path, template_bytes=template_bytes), adapter=adapter)

    def save_to_bytes(self) -> bytes:
        return self.adapter.save_to_bytes(self.presentation)

    def save_to_path(self, output_path: str) -> None:
        self.adapter.save_to_path(self.presentation, output_path)

    def delete_slide(self, slide_index: int) -> int:
        slide_id_el = self._get_slide_id_element(slide_index)
        slide_id = int(slide_id_el.get("id"))
        rel_id = slide_id_el.get(_tag(R_NS, "id"))
        self.presentation.part.drop_rel(rel_id)
        self.presentation.slides._sldIdLst.remove(slide_id_el)
        if self._has_sections():
            groups = self._read_sections()
            for group in groups:
                group["slide_ids"] = [item for item in group["slide_ids"] if item != slide_id]
            self._write_sections([group for group in groups if group["slide_ids"]])
        return slide_id

    def insert_slide(self, target_index: int, layout_index: int):
        if target_index < 0 or target_index > len(self.presentation.slides):
            raise OperationError(f"target_index {target_index} out of range")
        if layout_index < 0 or layout_index >= len(self.presentation.slide_layouts):
            raise OperationError(f"layout_index {layout_index} out of range")

        previous_slide_ids = self._slide_ids_in_order()
        groups = self._read_sections()
        slide = self.presentation.slides.add_slide(self.presentation.slide_layouts[layout_index])
        slide_id = slide.slide_id
        slide_id_el = self.presentation.slides._sldIdLst[-1]
        self.presentation.slides._sldIdLst.remove(slide_id_el)
        self.presentation.slides._sldIdLst.insert(target_index, slide_id_el)

        if groups:
            if target_index >= len(previous_slide_ids):
                target_group = len(groups) - 1
                groups[target_group]["slide_ids"].append(slide_id)
            else:
                anchor_slide_id = previous_slide_ids[target_index]
                target_group = self._group_index_for_slide(groups, anchor_slide_id)
                anchor_pos = groups[target_group]["slide_ids"].index(anchor_slide_id)
                groups[target_group]["slide_ids"].insert(anchor_pos, slide_id)
            self._write_sections(groups)
        return slide

    def add_section(self, name: str, start_slide_index: int) -> None:
        slide_ids = self._slide_ids_in_order()
        if not slide_ids:
            raise OperationError("cannot add a section to an empty presentation")
        if start_slide_index < 0 or start_slide_index >= len(slide_ids):
            raise OperationError(f"start_slide_index {start_slide_index} out of range")

        start_slide_id = slide_ids[start_slide_index]
        groups = self._read_sections()
        if not groups:
            before = slide_ids[:start_slide_index]
            after = slide_ids[start_slide_index:]
            groups = []
            if before:
                groups.append(self._make_section(DEFAULT_SECTION_NAME, before))
            groups.append(self._make_section(name, after))
            self._write_sections(groups)
            return

        group_index = self._group_index_for_slide(groups, start_slide_id)
        group = groups[group_index]
        first_slide_id = group["slide_ids"][0]
        if first_slide_id == start_slide_id:
            group["name"] = name
            self._write_sections(groups)
            return

        split_at = group["slide_ids"].index(start_slide_id)
        before = group["slide_ids"][:split_at]
        after = group["slide_ids"][split_at:]
        groups[group_index] = self._make_section(group["name"], before, guid=group["id"])
        groups.insert(group_index + 1, self._make_section(name, after))
        self._write_sections(groups)

    def delete_section(self, section_index: int) -> None:
        groups = self._read_sections()
        if not groups:
            raise OperationError("presentation does not contain sections")
        if section_index < 0 or section_index >= len(groups):
            raise OperationError(f"section_index {section_index} out of range")
        if len(groups) == 1:
            self._write_sections([])
            return

        removed = groups.pop(section_index)
        if section_index == 0:
            groups[0]["slide_ids"] = removed["slide_ids"] + groups[0]["slide_ids"]
        else:
            groups[section_index - 1]["slide_ids"].extend(removed["slide_ids"])
        self._write_sections(groups)

    def delete_table_row(self, slide_index: int, shape_locator: Union[int, str], row_index: int) -> None:
        table = self._resolve_table(slide_index, shape_locator)
        self._ensure_unmerged_table(table)
        if row_index < 0 or row_index >= len(table.rows):
            raise OperationError(f"row_index {row_index} out of range")
        table._tbl.remove(table._tbl.tr_lst[row_index])

    def delete_table_column(self, slide_index: int, shape_locator: Union[int, str], column_index: int) -> None:
        table = self._resolve_table(slide_index, shape_locator)
        self._ensure_unmerged_table(table)
        if column_index < 0 or column_index >= len(table.columns):
            raise OperationError(f"column_index {column_index} out of range")
        table._tbl.tblGrid.remove(table._tbl.tblGrid.gridCol_lst[column_index])
        for tr in table._tbl.tr_lst:
            tr.remove(tr.tc_lst[column_index])

    def merge_table_cells(
        self,
        slide_index: int,
        shape_locator: Union[int, str],
        first_row: int,
        first_col: int,
        last_row: int,
        last_col: int,
    ) -> None:
        table = self._resolve_table(slide_index, shape_locator)
        self._validate_merge_bounds(table, first_row, first_col, last_row, last_col)
        table.cell(first_row, first_col).merge(table.cell(last_row, last_col))

    def _resolve_table(self, slide_index: int, shape_locator: Union[int, str]):
        slide = self.adapter.get_slide(self.presentation, slide_index)
        shape = self.adapter.find_shape(slide, shape_locator)
        if not getattr(shape, "has_table", False):
            raise ShapeOperationError("target shape is not a table")
        return shape.table

    @staticmethod
    def _validate_merge_bounds(table, first_row: int, first_col: int, last_row: int, last_col: int) -> None:
        if first_row > last_row or first_col > last_col:
            raise OperationError("merge bounds must define a top-left to bottom-right rectangle")
        if first_row < 0 or first_col < 0 or last_row >= len(table.rows) or last_col >= len(table.columns):
            raise OperationError("merge bounds out of range")

    @staticmethod
    def _ensure_unmerged_table(table) -> None:
        for tr in table._tbl.tr_lst:
            for tc in tr.tc_lst:
                if tc.get("rowSpan") or tc.get("gridSpan") or tc.get("hMerge") or tc.get("vMerge"):
                    raise OperationError("row/column deletion is not supported on merged tables")

    def _get_slide_id_element(self, slide_index: int):
        slide_ids = list(self.presentation.slides._sldIdLst)
        if slide_index < 0 or slide_index >= len(slide_ids):
            raise OperationError(f"slide_index {slide_index} out of range")
        return slide_ids[slide_index]

    def _slide_ids_in_order(self) -> list[int]:
        return [slide.slide_id for slide in self.presentation.slides]

    def _group_index_for_slide(self, groups: list[dict], slide_id: int) -> int:
        for index, group in enumerate(groups):
            if slide_id in group["slide_ids"]:
                return index
        raise OperationError(f"slide id {slide_id} is not assigned to any section")

    @staticmethod
    def _make_section(name: str, slide_ids: list[int], guid: Optional[str] = None) -> dict:
        return {"name": name, "id": guid or f"{{{str(uuid.uuid4()).upper()}}}", "slide_ids": slide_ids}

    def _has_sections(self) -> bool:
        return self._find_section_ext() is not None

    def _read_sections(self) -> list[dict]:
        section_ext = self._find_section_ext()
        if section_ext is None:
            return []

        section_lst = section_ext.find(_tag(P14_NS, "sectionLst"))
        if section_lst is None:
            return []

        groups = []
        for section in section_lst.findall(_tag(P14_NS, "section")):
            slide_ids = [int(sld_id.get("id")) for sld_id in section.find(_tag(P14_NS, "sldIdLst")).findall(_tag(P14_NS, "sldId"))]
            groups.append(
                {
                    "name": section.get("name") or DEFAULT_SECTION_NAME,
                    "id": section.get("id") or f"{{{str(uuid.uuid4()).upper()}}}",
                    "slide_ids": slide_ids,
                }
            )
        return groups

    def _write_sections(self, groups: list[dict]) -> None:
        presentation_el = self.presentation.part._element
        ext_lst = presentation_el.find(_tag(P_NS, "extLst"))
        section_ext = self._find_section_ext()
        if not groups:
            if section_ext is not None:
                ext_lst.remove(section_ext)
            if ext_lst is not None and len(ext_lst) == 0:
                presentation_el.remove(ext_lst)
            return

        ordered_ids = self._slide_ids_in_order()
        order_map = {slide_id: index for index, slide_id in enumerate(ordered_ids)}
        normalized_groups = []
        for group in groups:
            slide_ids = sorted(
                [slide_id for slide_id in group["slide_ids"] if slide_id in order_map],
                key=order_map.__getitem__,
            )
            if slide_ids:
                normalized_groups.append({**group, "slide_ids": slide_ids})

        if ext_lst is None:
            ext_lst = etree.SubElement(presentation_el, _tag(P_NS, "extLst"))
        if section_ext is None:
            section_ext = etree.SubElement(ext_lst, _tag(P_NS, "ext"), uri=SECTION_EXT_URI)
        else:
            for child in list(section_ext):
                section_ext.remove(child)

        section_lst = etree.SubElement(section_ext, _tag(P14_NS, "sectionLst"), nsmap={"p14": P14_NS})
        for group in normalized_groups:
            section_el = etree.SubElement(
                section_lst,
                _tag(P14_NS, "section"),
                name=group["name"],
                id=group["id"],
            )
            slide_id_lst = etree.SubElement(section_el, _tag(P14_NS, "sldIdLst"))
            for slide_id in group["slide_ids"]:
                etree.SubElement(slide_id_lst, _tag(P14_NS, "sldId"), id=str(slide_id))

    def _find_section_ext(self):
        presentation_el = self.presentation.part._element
        ext_lst = presentation_el.find(_tag(P_NS, "extLst"))
        if ext_lst is None:
            return None
        for ext in ext_lst.findall(_tag(P_NS, "ext")):
            if ext.get("uri") == SECTION_EXT_URI:
                return ext
        return None

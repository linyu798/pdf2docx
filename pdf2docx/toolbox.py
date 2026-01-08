"""PDF toolbox utilities."""

from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass
from typing import Iterable, List, Optional, Sequence, Tuple

import fitz
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

from .converter import Converter


RangeList = List[Tuple[int, int]]


@dataclass(frozen=True)
class PageSelection:
    ranges: RangeList
    indices: List[int]


def _parse_page_ranges(page_ranges: str, total_pages: int, zero_based: bool = False) -> PageSelection:
    if not page_ranges:
        raise ValueError('Page ranges are required.')

    ranges: RangeList = []
    for token in re.split(r"\s*,\s*", page_ranges.strip()):
        if not token:
            continue
        if '-' in token:
            start_text, end_text = token.split('-', 1)
        else:
            start_text, end_text = token, token

        if not start_text.strip().isdigit() or not end_text.strip().isdigit():
            raise ValueError(f'Invalid page range token: {token!r}')

        start = int(start_text)
        end = int(end_text)
        if not zero_based:
            start -= 1
            end -= 1

        if start < 0 or end < 0 or start > end or end >= total_pages:
            raise ValueError(f'Invalid page range: {token!r}')

        ranges.append((start, end))

    if not ranges:
        raise ValueError('No valid page ranges found.')

    indices: List[int] = []
    for start, end in ranges:
        indices.extend(range(start, end + 1))

    return PageSelection(ranges=ranges, indices=indices)


def _ensure_parent_dir(path: str) -> None:
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)


class PDFToolbox:
    """A collection of PDF processing utilities."""

    @staticmethod
    def merge(
        pdf_files: Sequence[str],
        output_file: str,
        password: Optional[str] = None,
    ) -> str:
        if not pdf_files:
            raise ValueError('No PDF files provided for merging.')
        _ensure_parent_dir(output_file)

        writer = PdfWriter()
        for pdf_file in pdf_files:
            logging.info('Merging file: %s', pdf_file)
            reader = PdfReader(pdf_file)
            if password:
                reader.decrypt(password)
            for page in reader.pages:
                writer.add_page(page)

        with open(output_file, 'wb') as fp:
            writer.write(fp)
        return output_file

    @staticmethod
    def extract_pages(
        pdf_file: str,
        page_ranges: str,
        output_file: str,
        zero_based: bool = False,
        password: Optional[str] = None,
    ) -> str:
        reader = PdfReader(pdf_file)
        if password:
            reader.decrypt(password)
        selection = _parse_page_ranges(page_ranges, len(reader.pages), zero_based)
        _ensure_parent_dir(output_file)

        writer = PdfWriter()
        for index in selection.indices:
            writer.add_page(reader.pages[index])

        with open(output_file, 'wb') as fp:
            writer.write(fp)
        return output_file

    @staticmethod
    def delete_pages(
        pdf_file: str,
        page_ranges: str,
        output_file: str,
        zero_based: bool = False,
        password: Optional[str] = None,
    ) -> str:
        reader = PdfReader(pdf_file)
        if password:
            reader.decrypt(password)
        selection = _parse_page_ranges(page_ranges, len(reader.pages), zero_based)
        _ensure_parent_dir(output_file)

        remove_indices = set(selection.indices)
        writer = PdfWriter()
        for index, page in enumerate(reader.pages):
            if index not in remove_indices:
                writer.add_page(page)

        with open(output_file, 'wb') as fp:
            writer.write(fp)
        return output_file

    @staticmethod
    def split_pages(
        pdf_file: str,
        page_ranges: str,
        output_dir: str,
        zero_based: bool = False,
        password: Optional[str] = None,
    ) -> List[str]:
        reader = PdfReader(pdf_file)
        if password:
            reader.decrypt(password)
        selection = _parse_page_ranges(page_ranges, len(reader.pages), zero_based)
        os.makedirs(output_dir, exist_ok=True)

        output_files: List[str] = []
        base_name = os.path.splitext(os.path.basename(pdf_file))[0]
        for start, end in selection.ranges:
            writer = PdfWriter()
            for index in range(start, end + 1):
                writer.add_page(reader.pages[index])
            output_file = os.path.join(output_dir, f'{base_name}_{start + 1}-{end + 1}.pdf')
            with open(output_file, 'wb') as fp:
                writer.write(fp)
            output_files.append(output_file)

        return output_files

    @staticmethod
    def to_word(pdf_file: str, docx_file: str, password: Optional[str] = None, **kwargs) -> str:
        _ensure_parent_dir(docx_file)
        cv = Converter(pdf_file, password=password)
        try:
            cv.convert(docx_file, **kwargs)
        finally:
            cv.close()
        return docx_file

    @staticmethod
    def to_images(
        pdf_file: str,
        output_dir: str,
        image_format: str = 'png',
        dpi: int = 200,
    ) -> List[str]:
        image_format = image_format.lower().lstrip('.')
        if image_format not in {'png', 'jpeg', 'jpg'}:
            raise ValueError('Image format must be png or jpg/jpeg.')

        os.makedirs(output_dir, exist_ok=True)
        doc = fitz.open(pdf_file)
        output_files: List[str] = []
        for index in range(doc.page_count):
            page = doc.load_page(index)
            pix = page.get_pixmap(dpi=dpi)
            ext = 'jpg' if image_format in {'jpg', 'jpeg'} else 'png'
            output_file = os.path.join(output_dir, f'page_{index + 1}.{ext}')
            pix.save(output_file)
            output_files.append(output_file)
        doc.close()
        return output_files

    @staticmethod
    def to_excel(
        pdf_file: str,
        output_file: str,
        password: Optional[str] = None,
        start: int = 0,
        end: Optional[int] = None,
        pages: Optional[Iterable[int]] = None,
        **kwargs,
    ) -> str:
        _ensure_parent_dir(output_file)
        cv = Converter(pdf_file, password=password)
        try:
            tables = cv.extract_tables(start=start, end=end, pages=pages, **kwargs)
        finally:
            cv.close()

        if not tables:
            raise ValueError('No tables detected in the selected pages.')

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for index, table in enumerate(tables, start=1):
                df = pd.DataFrame(table)
                sheet_name = f'Table_{index}'
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

        return output_file

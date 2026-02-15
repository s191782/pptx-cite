#!/usr/bin/env python3
"""
PPTX Citation Manager - Replace LaTeX-style citations in PowerPoint files
"""

import argparse
import sys
import re
import yaml
from copy import deepcopy
from pathlib import Path
from typing import Dict, Optional, List, Set
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt, Inches

# ANSI color codes for terminal output
RESET = '\x1b[0m'
BOLD = '\x1b[1m'
RED = '\x1b[91m'
GREEN = '\x1b[92m'
YELLOW = '\x1b[93m'

class BibTeXParser:
    """Parse BibTeX files and extract citation information."""
    
    def __init__(self, bib_file: Path):
        self.bib_file = bib_file
        self.entries: Dict[str, Dict[str, str]] = {}
        self.parse()
    
    def parse(self):
        """Parse the BibTeX file and extract entries."""
        with open(self.bib_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Match BibTeX entries: @type{key, ...} with proper brace matching
        # Find all entries by looking for @type{
        entry_pattern = r'@(\w+)\s*\{\s*([^,\s]+)'
        
        for entry_match in re.finditer(entry_pattern, content):
            entry_type = entry_match.group(1).lower()
            cite_key = entry_match.group(2).strip()
            
            # Find the matching closing brace for this entry
            start_pos = entry_match.end()
            brace_count = 1
            pos = start_pos
            
            while pos < len(content) and brace_count > 0:
                if content[pos] == '{':
                    brace_count += 1
                elif content[pos] == '}':
                    brace_count -= 1
                pos += 1
            
            # Extract the fields section
            fields_str = content[start_pos:pos-1]
            
            # Parse fields with proper brace matching
            # Store the BibTeX entry type under a dedicated key to avoid
            # collisions with a user-defined `type = {...}` field.
            fields = {'entry_type': entry_type}
            
            # Find field assignments
            field_start_pattern = r'(\w+)\s*=\s*'
            i = 0
            while i < len(fields_str):
                match = re.match(field_start_pattern, fields_str[i:])
                if match:
                    field_name = match.group(1).lower()
                    i += match.end()
                    
                    # Skip whitespace
                    while i < len(fields_str) and fields_str[i].isspace():
                        i += 1
                    
                    if i >= len(fields_str):
                        break
                    
                    # Determine delimiter: { or "
                    if fields_str[i] == '{':
                        # Match braces
                        brace_count = 1
                        start = i + 1
                        i += 1
                        while i < len(fields_str) and brace_count > 0:
                            if fields_str[i] == '{':
                                brace_count += 1
                            elif fields_str[i] == '}':
                                brace_count -= 1
                            i += 1
                        field_value = fields_str[start:i-1].strip()
                        fields[field_name] = field_value
                    elif fields_str[i] == '"':
                        # Match quotes
                        start = i + 1
                        i += 1
                        while i < len(fields_str) and fields_str[i] != '"':
                            if fields_str[i] == '\\':
                                i += 2  # Skip escaped character
                            else:
                                i += 1
                        field_value = fields_str[start:i].strip()
                        fields[field_name] = field_value
                        i += 1  # Skip closing quote
                    else:
                        # Bare value (number or string) - read until comma or end
                        start = i
                        while i < len(fields_str) and fields_str[i] not in ',}\n':
                            i += 1
                        field_value = fields_str[start:i].strip()
                        fields[field_name] = field_value
                else:
                    i += 1
            
            # Clean curly brackets from all fields except URL
            for field_name, field_value in fields.items():
                if field_name not in ['url'] and field_value:
                    fields[field_name] = self._clean_field_value(field_value)
            
            self.entries[cite_key] = fields
    
    def _clean_field_value(self, value: str) -> str:
        """Remove curly brackets from field values (used for LaTeX protection)."""
        # Remove curly brackets often used in BibTeX for protecting capitalization
        value = value.replace('{', '').replace('}', '')
        return value.strip()
    
    def get_entry(self, key: str) -> Optional[Dict[str, str]]:
        """Get a bibliography entry by its cite key."""
        return self.entries.get(key)
    
    def _get_year(self, entry: Dict[str, str]) -> Optional[str]:
        """Extract year from entry, handling both 'year' and 'date' fields."""
        if 'year' in entry:
            return entry['year']
        elif 'date' in entry:
            # Extract year from date field (format: YYYY-MM-DD or YYYY)
            date_str = entry['date']
            # Take first 4 digits
            year_match = re.match(r'(\d{4})', date_str)
            if year_match:
                return year_match.group(1)
        return None
    
    def _clean_pages(self, pages: str) -> str:
        """Clean page numbers by converting double dashes to single dashes."""
        # Convert LaTeX-style double dashes (--) to single dash
        return pages.replace('--', '-')
    
    def _clean_author_name(self, name: str) -> str:
        """Clean author name by removing curly brackets and extra spaces."""
        # Remove curly brackets often used in BibTeX for protection
        name = name.replace('{', '').replace('}', '')
        # Clean up extra whitespace
        name = ' '.join(name.split())
        return name.strip()
    
    def _parse_authors(self, author_string: str) -> List[str]:
        """Parse author string into list of author names."""
        if not author_string:
            return []
        
        # Split by ' and ' (BibTeX standard separator)
        authors = [self._clean_author_name(a.strip()) for a in author_string.split(' and ')]
        return [a for a in authors if a]  # Remove empty strings
    
    def _format_author_list(self, author_string: str, max_authors: int = 3, use_et_al: bool = True) -> str:
        """Format author list with et al. for many authors."""
        authors = self._parse_authors(author_string)
        
        if not authors:
            return "Unknown"
        
        if len(authors) <= max_authors or not use_et_al:
            return ', '.join(authors)
        else:
            # More than max_authors: use et al.
            return f"{authors[0]} et al."
    
    def _get_last_name(self, author_name: str) -> str:
        """Extract last name from author name."""
        author_name = self._clean_author_name(author_name)
        
        # Handle "Last, First" format
        if ',' in author_name:
            return author_name.split(',')[0].strip()
        
        # Handle "First Last" format - take last word
        parts = author_name.split()
        if parts:
            return parts[-1]
        
        return author_name
    
    def format_citation(self, key: str, style: str = 'numeric') -> str:
        """Format a single citation."""
        entry = self.get_entry(key)
        if not entry:
            return f"[{key}?]"
        
        if style == 'numeric':
            # Will be replaced with actual number later
            return f"[{key}]"
        elif style == 'author-year':
            author = entry.get('author', 'Unknown')
            year = self._get_year(entry) or 'n.d.'
            
            # Parse authors and get first author's last name
            authors = self._parse_authors(author)
            if authors:
                first_last = self._get_last_name(authors[0])
                if len(authors) > 1:
                    first_last += " et al."
            else:
                first_last = "Unknown"
            
            return f"({first_last}, {year})"
        elif style == 'author':
            author = entry.get('author', 'Unknown')
            
            # Parse authors and get first author's last name
            authors = self._parse_authors(author)
            if authors:
                first_last = self._get_last_name(authors[0])
                if len(authors) > 1:
                    first_last += " et al."
            else:
                first_last = "Unknown"
            
            return f"({first_last})"
        else:
            return f"[{key}]"
    
    def format_bibliography_entry(self, key: str, number: Optional[int] = None, style: str = 'numeric') -> str:
        """Format a full bibliography entry."""
        entry = self.get_entry(key)
        if not entry:
            if style == 'numeric':
                return f"[{number if number is not None else '?'}] {key}: Entry not found"
            elif style == 'numeric-parens':
                return f"({number if number is not None else '?'}) {key}: Entry not found"
            else:
                return f"{key}: Entry not found"
        
        # Format based on entry type
        entry_type = entry.get('entry_type', 'misc')
        result = []
        
        # Number (only for numeric styles)
        if number is not None and style in ['numeric', 'numeric-parens']:
            if style == 'numeric':
                result.append(f"[{number}]")
            else:  # numeric-parens
                result.append(f"({number})")
        
        # Author(s) or Editor(s)
        if 'author' in entry:
            authors_formatted = self._format_author_list(entry['author'], max_authors=3, use_et_al=True)
            result.append(f"{authors_formatted}.")
        elif 'editor' in entry:
            authors = self._parse_authors(entry['editor'])
            editors_formatted = self._format_author_list(entry['editor'], max_authors=3, use_et_al=True)
            suffix = " (Ed.)" if len(authors) == 1 else " (Eds.)"
            result.append(f"{editors_formatted}{suffix}.")
        
        # Title
        if 'title' in entry:
            result.append(f'"{entry["title"]}".')
        
        # Publication details based on type
        if entry_type == 'article':
            # Journal article
            pub_parts = []
            if 'journal' in entry:
                pub_parts.append(entry['journal'])
            if 'volume' in entry:
                vol_str = f"vol. {entry['volume']}"
                if 'number' in entry:
                    vol_str += f", no. {entry['number']}"
                pub_parts.append(vol_str)
            if 'pages' in entry:
                pub_parts.append(f"pp. {self._clean_pages(entry['pages'])}")
            year = self._get_year(entry)
            if year:
                pub_parts.append(year)
            if pub_parts:
                result.append(', '.join(pub_parts) + '.')
        
        elif entry_type == 'book':
            # Book
            pub_parts = []
            if 'edition' in entry:
                pub_parts.append(f"{entry['edition']} ed.")
            if 'volume' in entry:
                pub_parts.append(f"vol. {entry['volume']}")
            if 'publisher' in entry:
                pub_parts.append(entry['publisher'])
            if 'address' in entry:
                pub_parts.append(entry['address'])
            if 'year' in entry:
                pub_parts.append(entry['year'])
            if pub_parts:
                result.append(', '.join(pub_parts) + '.')
        
        elif entry_type in ['inproceedings', 'conference']:
            # Conference paper
            pub_parts = []
            if 'booktitle' in entry:
                pub_parts.append(f"in {entry['booktitle']}")
            if 'editor' in entry and 'author' in entry:
                editors = entry['editor'].replace(' and ', ', ')
                pub_parts.append(f"Eds. {editors}")
            if 'pages' in entry:
                pub_parts.append(f"pp. {self._clean_pages(entry['pages'])}")
            if 'organization' in entry:
                pub_parts.append(entry['organization'])
            if 'publisher' in entry:
                pub_parts.append(entry['publisher'])
            if 'address' in entry:
                pub_parts.append(entry['address'])
            year = self._get_year(entry)
            if year:
                pub_parts.append(year)
            if pub_parts:
                result.append(', '.join(pub_parts) + '.')
        
        elif entry_type in ['incollection', 'inbook']:
            # Chapter or part of a book
            pub_parts = []
            if 'booktitle' in entry:
                pub_parts.append(f"in {entry['booktitle']}")
            if 'editor' in entry and 'author' in entry:
                editors = entry['editor'].replace(' and ', ', ')
                pub_parts.append(f"Eds. {editors}")
            if 'chapter' in entry:
                pub_parts.append(f"ch. {entry['chapter']}")
            if 'pages' in entry:
                pub_parts.append(f"pp. {self._clean_pages(entry['pages'])}")
            if 'edition' in entry:
                pub_parts.append(f"{entry['edition']} ed.")
            if 'publisher' in entry:
                pub_parts.append(entry['publisher'])
            if 'address' in entry:
                pub_parts.append(entry['address'])
            year = self._get_year(entry)
            if year:
                pub_parts.append(year)
            if pub_parts:
                result.append(', '.join(pub_parts) + '.')
        
        elif entry_type in ['mastersthesis', 'phdthesis', 'thesis']:
            # Thesis
            if entry_type == 'phdthesis':
                thesis_type = 'Ph.D. dissertation'
            elif entry_type == 'mastersthesis':
                thesis_type = "Master's thesis"
            elif entry_type == 'thesis':
                # Use type field if available, otherwise default
                thesis_type = entry.get('type', 'Thesis')
            else:
                thesis_type = 'Thesis'
            
            pub_parts = [thesis_type]
            if 'school' in entry:
                pub_parts.append(entry['school'])
            elif 'institution' in entry:
                pub_parts.append(entry['institution'])
            if 'address' in entry:
                pub_parts.append(entry['address'])
            year = self._get_year(entry)
            if year:
                pub_parts.append(year)
            if pub_parts:
                result.append(', '.join(pub_parts) + '.')
        
        elif entry_type in ['techreport', 'report']:
            # Technical report
            pub_parts = []
            if 'number' in entry:
                pub_parts.append(f"Tech. Rep. {entry['number']}")
            if 'institution' in entry:
                pub_parts.append(entry['institution'])
            if 'address' in entry:
                pub_parts.append(entry['address'])
            year = self._get_year(entry)
            if year:
                pub_parts.append(year)
            if pub_parts:
                result.append(', '.join(pub_parts) + '.')
        
        elif entry_type == 'manual':
            # Technical manual
            pub_parts = []
            if 'organization' in entry:
                pub_parts.append(entry['organization'])
            if 'address' in entry:
                pub_parts.append(entry['address'])
            if 'edition' in entry:
                pub_parts.append(f"{entry['edition']} ed.")
            year = self._get_year(entry)
            if year:
                pub_parts.append(year)
            if pub_parts:
                result.append(', '.join(pub_parts) + '.')
        
        elif entry_type == 'proceedings':
            # Conference proceedings
            pub_parts = []
            if 'volume' in entry:
                pub_parts.append(f"vol. {entry['volume']}")
            if 'organization' in entry:
                pub_parts.append(entry['organization'])
            if 'publisher' in entry:
                pub_parts.append(entry['publisher'])
            if 'address' in entry:
                pub_parts.append(entry['address'])
            year = self._get_year(entry)
            if year:
                pub_parts.append(year)
            if pub_parts:
                result.append(', '.join(pub_parts) + '.')
        
        elif entry_type == 'unpublished':
            # Unpublished work
            pub_parts = []
            if 'note' in entry:
                pub_parts.append(entry['note'])
            year = self._get_year(entry)
            if year:
                pub_parts.append(year)
            if pub_parts:
                result.append(', '.join(pub_parts) + '.')
        
        elif entry_type in ['online', 'electronic', 'www']:
            # Online resource
            pub_parts = []
            if 'organization' in entry:
                pub_parts.append(entry['organization'])
            year = self._get_year(entry)
            if year:
                pub_parts.append(year)
            if pub_parts:
                result.append(', '.join(pub_parts) + '.')
            
            # URL and access date
            url_part = []
            if 'url' in entry:
                url_part.append(entry['url'])
            if 'urldate' in entry:
                # Format: [Accessed: date]
                url_part.append(f"[Accessed: {entry['urldate']}]")
            
            if url_part:
                result.append(' '.join(url_part) + '.')
        
        elif entry_type == 'patent':
            # Patent
            pub_parts = []
            if 'number' in entry:
                pub_parts.append(f"Patent {entry['number']}")
            if 'nationality' in entry:
                pub_parts.append(entry['nationality'])
            year = self._get_year(entry)
            if year:
                pub_parts.append(year)
            if pub_parts:
                result.append(', '.join(pub_parts) + '.')
        
        elif entry_type == 'booklet':
            # Booklet
            pub_parts = []
            if 'howpublished' in entry:
                pub_parts.append(entry['howpublished'])
            if 'address' in entry:
                pub_parts.append(entry['address'])
            year = self._get_year(entry)
            if year:
                pub_parts.append(year)
            if pub_parts:
                result.append(', '.join(pub_parts) + '.')
        
        elif entry_type in ['collection', 'reference', 'mvbook']:
            # Collection or multi-volume work
            pub_parts = []
            if 'volumes' in entry:
                pub_parts.append(f"{entry['volumes']} vols.")
            elif 'volume' in entry:
                pub_parts.append(f"vol. {entry['volume']}")
            if 'publisher' in entry:
                pub_parts.append(entry['publisher'])
            if 'address' in entry:
                pub_parts.append(entry['address'])
            year = self._get_year(entry)
            if year:
                pub_parts.append(year)
            if pub_parts:
                result.append(', '.join(pub_parts) + '.')
        
        else:
            # Generic fallback for misc and other types
            pub_parts = []
            if 'howpublished' in entry:
                pub_parts.append(entry['howpublished'])
            if 'journal' in entry:
                pub_parts.append(entry['journal'])
            elif 'booktitle' in entry:
                pub_parts.append(entry['booktitle'])
            elif 'publisher' in entry:
                pub_parts.append(entry['publisher'])
            elif 'institution' in entry:
                pub_parts.append(entry['institution'])
            elif 'organization' in entry:
                pub_parts.append(entry['organization'])
            year = self._get_year(entry)
            if year:
                pub_parts.append(year)
            if pub_parts:
                result.append(', '.join(pub_parts) + '.')
            if 'note' in entry:
                result.append(f"Note: {entry['note']}.")
        
        # Additional fields (DOI, URL) - only if not already added
        if 'doi' in entry:
            result.append(f"DOI: {entry['doi']}.")
        if 'url' in entry and 'doi' not in entry and entry_type not in ['online', 'electronic', 'www']:
            url_part = [entry['url']]
            if 'urldate' in entry:
                url_part.append(f"[Accessed: {entry['urldate']}]")
            result.append(' '.join(url_part) + '.')
        
        return ' '.join(result)  

class CitationProcessor:
    """Process citations in PowerPoint presentations."""
    
    def __init__(self, pptx_file: Path, bib_parser: BibTeXParser, config: Dict):
        self.pptx_file = pptx_file
        self.bib_parser = bib_parser
        self.config = config
        self.prs = Presentation(str(pptx_file))
        self.cited_keys: List[str] = []  # Track order of citations
        self.cited_set: Set[str] = set()  # Track unique citations
        self.missing_citations: Dict[str, List[int]] = {}  # Track missing citations with slide numbers
        self._citation_index_map: Dict[str, int] = {}  # Precomputed index map for numeric styles
    
    def process(self) -> None:
        """Process all slides for citations and bibliography commands."""
        bib_command = self.config.get('bibliography_command', r'\bibliography')
        
        # Check if bibliography command exists in presentation
        has_bibliography = False
        for slide in self.prs.slides:
            if self._slide_contains_bibliography_command(slide, bib_command):
                has_bibliography = True
                break
        
        # First pass: collect citations from all slides
        slides_to_process = list(enumerate(self.prs.slides))
        for slide_idx, slide in slides_to_process:
            # Process citations on all slides, but exclude shapes containing bibliography command
            self._process_citations_in_slide(slide, slide_idx, remove_only=not has_bibliography, skip_bib_shapes=True)
        
        # Print warnings for missing citations
        if self.missing_citations:
            for key, slide_numbers in sorted(self.missing_citations.items()):
                slides_str = ', '.join(map(str, slide_numbers))
                print(f"  {YELLOW}Warning:{RESET} Citation '{key}' not found in bibliography file (slides: {slides_str})")
        
        # Second pass: create bibliography (potentially adding slides) - only if bibliography exists
        if has_bibliography:
            for slide_idx, slide in slides_to_process:
                if self._slide_contains_bibliography_command(slide, bib_command):
                    additional_slides = self._create_bibliography_slide(slide, slide_idx)
                    if additional_slides > 0:
                        print(f"  Created {additional_slides} continuation slide(s) for bibliography")
                    break  # Only process first bibliography command
        else:
            print(f"  {YELLOW}Warning:{RESET} No bibliography command found - citations removed")
    
    def _slide_contains_bibliography_command(self, slide, command: str) -> bool:
        """Check if a slide contains the bibliography command."""
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                if command in shape.text:
                    return True
        return False
    
    def _process_citations_in_slide(self, slide, slide_idx: int, remove_only: bool = False, skip_bib_shapes: bool = False) -> None:
        """Process citations in a single slide.
        
        Args:
            slide: The slide to process
            slide_idx: The index of the slide
            remove_only: If True, remove citations without replacing them
            skip_bib_shapes: If True, skip shapes that contain the bibliography command
        """
        cite_command = self.config.get('cite_command', r'\cite')
        cite_style = self.config.get('cite_style', 'numeric')
        bib_command = self.config.get('bibliography_command', r'\bibliography')
        
        for shape in slide.shapes:
            # Skip shapes containing bibliography command if requested
            if skip_bib_shapes and hasattr(shape, "text") and bib_command in shape.text:
                continue
            
            # Process regular text frames
            if hasattr(shape, "text_frame"):
                self._process_text_frame(shape.text_frame, cite_command, cite_style, slide_idx, remove_only)
            
            # Process group shapes
            if hasattr(shape, 'shape_type') and hasattr(shape, 'shapes'):
                try:
                    # Check if it's a GROUP (type 6)
                    if shape.shape_type == MSO_SHAPE_TYPE.GROUP or shape.shape_type == 6:
                        self._process_group_shape(shape, cite_command, cite_style, slide_idx, remove_only)
                except (AttributeError, KeyError, TypeError):
                    # Some shapes may have unreliable shape_type access, try processing as group anyway
                    # This handles edge cases where shape_type property raises an exception
                    try:
                        self._process_group_shape(shape, cite_command, cite_style, slide_idx, remove_only)
                    except (AttributeError, TypeError, KeyError):
                        # Expected exceptions when shape doesn't support group operations
                        # Skip this shape if it can't be processed as a group
                        pass
            
            # Process tables within GraphicFrame
            if hasattr(shape, 'has_table') and shape.has_table:
                self._process_table(shape.table, cite_command, cite_style, slide_idx, remove_only)
    
    def _process_table(self, table, cite_command: str, cite_style: str, slide_idx: int, remove_only: bool = False) -> None:
        """Process citations in table cells."""
        try:
            for row in table.rows:
                for cell in row.cells:
                    if hasattr(cell, 'text_frame'):
                        self._process_text_frame(cell.text_frame, cite_command, cite_style, slide_idx, remove_only)
        except (AttributeError, TypeError):
            pass
    
    def _process_group_shape(self, group_shape, cite_command: str, cite_style: str, slide_idx: int, remove_only: bool = False) -> None:
        """Process citations in grouped shapes."""
        try:
            # Try to access shapes in the group
            if hasattr(group_shape, 'shapes'):
                for shape in group_shape.shapes:
                    if hasattr(shape, "text_frame"):
                        self._process_text_frame(shape.text_frame, cite_command, cite_style, slide_idx, remove_only)
                    # Recursively process nested groups
                    if hasattr(shape, 'shapes'):
                        self._process_group_shape(shape, cite_command, cite_style, slide_idx, remove_only)
        except (AttributeError, TypeError):
            # If we can't access the group shapes, skip silently
            pass
    
    def _process_text_frame(self, text_frame, cite_command: str, cite_style: str, slide_idx: int, remove_only: bool = False) -> None:
        """Process citations in a text frame."""
        for paragraph in text_frame.paragraphs:
            # Get full paragraph text
            full_text = paragraph.text
            
            # Find all citations in the paragraph
            # Pattern: \cite{key1,key2} or \cite{key1}
            escaped_command = re.escape(cite_command)
            pattern = rf'{escaped_command}\{{([^}}]+)\}}'
            
            matches = list(re.finditer(pattern, full_text))
            if not matches:
                continue
            
            # First pass: track citations in forward order (order of appearance) - only if not removing
            if not remove_only:
                for match in matches:
                    keys_str = match.group(1)
                    keys = [k.strip() for k in keys_str.split(',')]
                    
                    # Track citations and check if they exist
                    for key in keys:
                        if key not in self.cited_set:
                            self.cited_keys.append(key)
                            self.cited_set.add(key)
                            # Build index map incrementally for new keys to avoid O(n) lookups
                            self._citation_index_map[key] = len(self._citation_index_map)
                        # Check if citation exists in bibliography and track missing ones
                        if not self.bib_parser.get_entry(key):
                            if key not in self.missing_citations:
                                self.missing_citations[key] = []
                            # Add slide number (1-indexed for user display)
                            if slide_idx + 1 not in self.missing_citations[key]:
                                self.missing_citations[key].append(slide_idx + 1)
            
            # Second pass: replace citations in reverse order to maintain text indices
            for match in reversed(matches):
                if remove_only:
                    # Remove citations entirely
                    replacement = ""
                else:
                    keys_str = match.group(1)
                    keys = [k.strip() for k in keys_str.split(',')]
                    
                    if cite_style in ('numeric', 'numeric-parens'):
                        # Numeric styles: use precomputed index mapping to avoid O(n) lookups
                        citation_start = self.config.get('citation_start_number', 1)
                        numbers = [str(self._citation_index_map[key] + citation_start) for key in keys]
                        if cite_style == 'numeric':
                            # Replace with numeric references in square brackets
                            replacement = f"[{', '.join(numbers)}]"
                        else:
                            # numeric-parens: replace with numeric references in parentheses
                            replacement = f"({', '.join(numbers)})"
                    else:
                        # Non-numeric styles: format full citations
                        citations = [self.bib_parser.format_citation(key, cite_style) for key in keys]
                        replacement = ', '.join(citations)
                
                # Replace in the paragraph
                self._replace_text_in_paragraph(paragraph, match.start(), match.end(), replacement)
    
    def _replace_text_in_paragraph(self, paragraph, start: int, end: int, replacement: str) -> None:
        """Replace text in a paragraph while preserving formatting."""
        # Find which run(s) contain the text to replace
        runs = paragraph.runs
        if not runs:
            # No runs, create one with the replacement
            paragraph.add_run().text = replacement
            return
        
        # Calculate position of each run in the paragraph
        run_positions = []
        pos = 0
        for run in runs:
            run_start = pos
            run_end = pos + len(run.text)
            run_positions.append((run, run_start, run_end))
            pos = run_end
        
        # Find runs that overlap with the replacement range [start, end)
        affected_runs = []
        for run, run_start, run_end in run_positions:
            if run_end > start and run_start < end:
                affected_runs.append((run, run_start, run_end))
        
        if not affected_runs:
            # No runs affected (shouldn't happen), fallback
            return
        
        # Get the first affected run to use as template for formatting
        first_run, first_start, first_end = affected_runs[0]
        
        # Calculate text for the first affected run
        # It should contain: [text before match] + [replacement] + [text after match if it's the only run]
        before_text = first_run.text[:max(0, start - first_start)]
        
        if len(affected_runs) == 1:
            # Single run contains the entire match
            after_text = first_run.text[end - first_start:]
            first_run.text = before_text + replacement + after_text
        else:
            # Multiple runs - first run gets: before + replacement
            first_run.text = before_text + replacement
            
            # Get text after the match from the last affected run
            last_run, last_start, last_end = affected_runs[-1]
            after_text = last_run.text[end - last_start:]
            last_run.text = after_text
            
            # Clear middle runs if any
            for run, run_start, run_end in affected_runs[1:-1]:
                run.text = ""
    
    def _create_bibliography_slide(self, slide, slide_index: int) -> int:
        """Create or update a bibliography slide.
        
        Returns the number of additional slides created for overflow.
        """
        bib_command = self.config.get('bibliography_command', r'\bibliography')
        
        # Find the shape and paragraph containing the bibliography command
        for shape in slide.shapes:
            if hasattr(shape, "text_frame") and bib_command in shape.text:
                text_frame = shape.text_frame
                
                # Get text frame height for accurate capacity calculation
                # Use text_frame margins if available, otherwise estimate from shape
                if hasattr(text_frame, 'margin_top') and hasattr(text_frame, 'margin_bottom'):
                    available_height = shape.height - text_frame.margin_top - text_frame.margin_bottom
                else:
                    # Fallback: estimate margins (typically ~0.1 inches top/bottom)
                    available_height = shape.height - Inches(0.2)
                
                # Find the paragraph with the command to extract all formatting
                template_para = None
                template_run = None
                
                for i, para in enumerate(text_frame.paragraphs):
                    if bib_command in para.text:
                        template_para = para
                        # Get formatting from first run that has text
                        for run in para.runs:
                            if run.text.strip():
                                template_run = run
                                break
                        break
                
                if template_para is None:
                    return 0
                
                # Calculate line height from font size
                line_height = Pt(12)  # Default
                if template_run and template_run.font.size:
                    line_height = template_run.font.size * 1.2  # Add 20% for line spacing
                
                # Estimate average characters per line based on font size and text box width
                # Average character width is approximately 0.5 * font_size for most fonts
                font_size = template_run.font.size if template_run and template_run.font.size else Pt(12)
                avg_char_width = font_size * 0.5
                effective_width = shape.width
                if hasattr(text_frame, 'margin_left') and hasattr(text_frame, 'margin_right'):
                    effective_width = shape.width - text_frame.margin_left - text_frame.margin_right
                # Account for hanging indent (0.5 inches) - subsequent lines are indented
                effective_width_wrapped = effective_width - Inches(0.5)
                chars_per_line_first = max(50, int(effective_width / avg_char_width))
                chars_per_line_wrapped = max(40, int(effective_width_wrapped / avg_char_width))
                
                # Calculate how many entries can fit, considering line wrapping
                # Minimal safety margin (1%) to maximize space usage
                total_lines_available = int(available_height / line_height * 0.99)
                
                # Estimate lines per entry by checking actual entry lengths
                entries_to_add = list(self.cited_keys)
                entries_on_this_slide = []
                overflow_entries = []
                lines_used = 0
                
                citation_start = self.config.get('citation_start_number', 1)
                for i, key in enumerate(entries_to_add):
                    entry_text = self.bib_parser.format_bibliography_entry(key, i + citation_start, self.config.get('cite_style', 'numeric'))
                    # Estimate how many lines this entry will take
                    entry_chars = len(entry_text)
                    # First line uses full width, subsequent lines are indented
                    if entry_chars <= chars_per_line_first:
                        lines_for_entry = 1
                    else:
                        # First line + wrapped lines
                        remaining_chars = entry_chars - chars_per_line_first
                        lines_for_entry = 1 + max(1, (remaining_chars + chars_per_line_wrapped - 1) // chars_per_line_wrapped)
                    
                    # Check if we have room for this entry
                    if lines_used + lines_for_entry <= total_lines_available:
                        entries_on_this_slide.append(key)
                        lines_used += lines_for_entry
                    else:
                        overflow_entries.append(key)
                
                # Ensure at least one entry on first slide
                if not entries_on_this_slide and entries_to_add:
                    entries_on_this_slide = [entries_to_add[0]]
                    overflow_entries = entries_to_add[1:]
                
                # First, replace the command paragraph with first entry
                if entries_on_this_slide:
                    # Clear existing runs and set new text
                    for run in template_para.runs:
                        run.text = ""
                    
                    citation_start = self.config.get('citation_start_number', 1)
                    entry_text = self.bib_parser.format_bibliography_entry(entries_on_this_slide[0], citation_start, self.config.get('cite_style', 'numeric'))
                    run = template_para.add_run()
                    run.text = entry_text
                    
                    # Set hanging indent for wrapped lines using XML
                    # Get or create paragraph properties
                    pPr = template_para._element.get_or_add_pPr()
                    
                    # Set indentation (914400 EMU = 1 inch, so 457200 EMU = 0.5 inch)
                    indent_emu = 457200  # 0.5 inches in EMUs
                    pPr.set('marL', str(indent_emu))  # Left margin
                    pPr.set('indent', str(-indent_emu))  # First line indent (negative for hanging)
                    
                    # Copy font formatting
                    if template_run:
                        if template_run.font.size:
                            run.font.size = template_run.font.size
                        if template_run.font.name:
                            run.font.name = template_run.font.name
                        if template_run.font.bold is not None:
                            run.font.bold = template_run.font.bold
                        if template_run.font.italic is not None:
                            run.font.italic = template_run.font.italic
                        if template_run.font.color and template_run.font.color.type is not None:
                            run.font.color.rgb = template_run.font.color.rgb
                
                # Add remaining entries for this slide
                citation_start = self.config.get('citation_start_number', 1)
                for idx, key in enumerate(entries_on_this_slide[1:], citation_start + 1):
                    entry_text = self.bib_parser.format_bibliography_entry(key, idx, self.config.get('cite_style', 'numeric'))
                    p = text_frame.add_paragraph()
                    run = p.add_run()
                    run.text = entry_text
                    
                    # Set hanging indent for wrapped lines using XML
                    pPr = p._element.get_or_add_pPr()
                    indent_emu = 457200  # 0.5 inches
                    pPr.set('marL', str(indent_emu))
                    pPr.set('indent', str(-indent_emu))
                    
                    # Copy all formatting from template
                    if template_run:
                        if template_run.font.size:
                            run.font.size = template_run.font.size
                        if template_run.font.name:
                            run.font.name = template_run.font.name
                        if template_run.font.bold is not None:
                            run.font.bold = template_run.font.bold
                        if template_run.font.italic is not None:
                            run.font.italic = template_run.font.italic
                        if template_run.font.color and template_run.font.color.type is not None:
                            run.font.color.rgb = template_run.font.color.rgb
                    
                    # Copy complete paragraph properties at XML level
                    if template_para:
                        template_pPr = template_para._element.pPr
                        
                        if template_pPr is not None:
                            # Remove the default pPr if it exists
                            new_pPr = p._element.pPr
                            if new_pPr is not None:
                                p._element.remove(new_pPr)
                            
                            # Copy the entire pPr element from template
                            p._element.insert(0, deepcopy(template_pPr))
                
                # Handle overflow - create continuation slide(s) immediately after this slide
                if overflow_entries:
                    print(f"  {GREEN}Info:{RESET} Bibliography has {len(overflow_entries)} overflow entries.")
                    print("  Creating continuation slide(s)...")
                    
                    # Get the current slide index
                    current_slide_index = slide_index + 1  # +1 because slide_index is 0-based, insertion is 1-based
                    
                    # Track overflow
                    slides_created = 0
                    remaining = overflow_entries
                    citation_start = self.config.get('citation_start_number', 1)
                    start_idx = len(entries_on_this_slide) + citation_start
                    
                    while remaining:
                        # Create a new slide with the same layout as the original bibliography slide.
                        # The layout carries background and placeholder formatting, which is safely managed
                        # by python-pptx without needing to clone raw XML.
                        new_slide = self.prs.slides.add_slide(slide.slide_layout)
                        
                        # Optionally copy the title text using the high-level API (if both slides have a title).
                        try:
                            orig_title = slide.shapes.title
                        except Exception:
                            orig_title = None
                        try:
                            new_title = new_slide.shapes.title
                        except Exception:
                            new_title = None
                        if orig_title is not None and new_title is not None:
                            new_title.text = orig_title.text
                        
                        # Add a text box for the bibliography content at the same position as the original
                        # Use the high-level API to safely create the shape
                        new_text_box = new_slide.shapes.add_textbox(
                            shape.left, shape.top, shape.width, shape.height
                        )
                        new_text_frame = new_text_box.text_frame
                        
                        # Then move it to the correct position (right after current slide + already created overflow slides)
                        target_position = current_slide_index + slides_created
                        
                        # Get the slide parts
                        slide_id_list = self.prs.slides._sldIdLst
                        slide_ids = list(slide_id_list)
                        
                        # Move the last slide (newly added) to the target position
                        if len(slide_ids) > target_position:
                            last_slide = slide_ids[-1]
                            slide_id_list.remove(last_slide)
                            slide_id_list.insert(target_position, last_slide)
                        
                        slides_created += 1
                        
                        # Clear the bibliography text frame completely (no blank lines)
                        new_text_frame.clear()
                        
                        # Calculate how many entries fit on this overflow slide
                        lines_used_overflow = 0
                        entries_for_this_slide = []
                        
                        for key in remaining:
                            entry_text = self.bib_parser.format_bibliography_entry(key, start_idx + len(entries_for_this_slide), self.config.get('cite_style', 'numeric'))
                            entry_chars = len(entry_text)
                            # First line uses full width, subsequent lines are indented
                            if entry_chars <= chars_per_line_first:
                                lines_for_entry = 1
                            else:
                                remaining_chars = entry_chars - chars_per_line_first
                                lines_for_entry = 1 + max(1, (remaining_chars + chars_per_line_wrapped - 1) // chars_per_line_wrapped)
                            
                            if lines_used_overflow + lines_for_entry <= total_lines_available:
                                entries_for_this_slide.append(key)
                                lines_used_overflow += lines_for_entry
                            else:
                                break
                        
                        # Update remaining entries
                        remaining = remaining[len(entries_for_this_slide):]
                        
                        for i, key in enumerate(entries_for_this_slide):
                            entry_idx = start_idx + i
                            entry_text = self.bib_parser.format_bibliography_entry(key, entry_idx, self.config.get('cite_style', 'numeric'))
                            
                            # Use the first existing paragraph for the first entry, add new ones for the rest
                            if i == 0:
                                # Use the existing paragraph from clear() for first entry
                                p = new_text_frame.paragraphs[0]
                            else:
                                # Add new paragraph for subsequent entries
                                p = new_text_frame.add_paragraph()
                            
                            run = p.add_run()
                            run.text = entry_text
                            
                            # Copy formatting from template first
                            if template_run:
                                if template_run.font.size:
                                    run.font.size = template_run.font.size
                                if template_run.font.name:
                                    run.font.name = template_run.font.name
                                if template_run.font.bold is not None:
                                    run.font.bold = template_run.font.bold
                                if template_run.font.italic is not None:
                                    run.font.italic = template_run.font.italic
                                if template_run.font.color and template_run.font.color.type is not None:
                                    run.font.color.rgb = template_run.font.color.rgb
                            
                            # Copy paragraph properties from template
                            if template_para:
                                template_pPr = template_para._element.pPr
                                if template_pPr is not None:
                                    new_pPr = p._element.pPr
                                    if new_pPr is not None:
                                        p._element.remove(new_pPr)
                                    p._element.insert(0, deepcopy(template_pPr))
                            
                            # Now add hanging indent for wrapped lines
                            pPr = p._element.get_or_add_pPr()
                            indent_emu = 457200  # 0.5 inches
                            pPr.set('marL', str(indent_emu))
                            pPr.set('indent', str(-indent_emu))
                        
                        start_idx += len(entries_for_this_slide)
                    
                    return slides_created
                
                return 0
        
        # If no existing text box found, create a new one
        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(9)
        height = Inches(5)
        
        text_box = slide.shapes.add_textbox(left, top, width, height)
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        
        # Add title
        p = text_frame.paragraphs[0]
        run = p.add_run()
        run.text = "References"
        run.font.bold = True
        run.font.size = Pt(18)
        
        # Add bibliography entries
        citation_start = self.config.get('citation_start_number', 1)
        for idx, key in enumerate(self.cited_keys, citation_start):
            entry_text = self.bib_parser.format_bibliography_entry(key, idx, self.config.get('cite_style', 'numeric'))
            p = text_frame.add_paragraph()
            run = p.add_run()
            run.text = entry_text
            run.font.size = Pt(12)
            p.level = 0
        
        return 0
    
    def save(self, output_file: Optional[Path] = None) -> None:
        """Save the modified presentation."""
        if output_file is None:
            output_file = self.pptx_file.parent / f"{self.pptx_file.stem}_cited{self.pptx_file.suffix}"
        
        self.prs.save(str(output_file))
        print(f"Saved to: {output_file}")

def load_config(config_file: Optional[Path]) -> Dict:
    """Load configuration from YAML file."""
    default_config = {
        'cite_command': r'\cite',
        'bibliography_command': r'\bibliography',
        'cite_style': 'numeric',  # Options: numeric, numeric-parens, author-year, author
    }
    
    if config_file and config_file.exists():
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                user_config = yaml.safe_load(f)
                # Treat an empty configuration file as an empty dict; reject non-dict types
                if user_config is None:
                    user_config = {}
                elif not isinstance(user_config, dict):
                    print(f"{RED}Error: Configuration file '{config_file}' must be a valid YAML dictionary.{RESET}")
                    sys.exit(1)
                default_config.update(user_config)
        except yaml.YAMLError as e:
            print(f"{RED}Error: Configuration file '{config_file}' contains invalid YAML syntax:{RESET}")
            print(f"{RED}{e}{RESET}")
            sys.exit(1)
        except (IOError, OSError, PermissionError) as e:
            print(f"{RED}Error: Failed to read configuration file '{config_file}': {e}{RESET}")
            sys.exit(1)
    
    return default_config

def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description='Replace LaTeX-style citations in PowerPoint files using BibTeX references'
    )
    parser.add_argument('pptx_file', type=Path, help='Path to the PowerPoint file (.pptx)')
    parser.add_argument('bib_file', type=Path, nargs='?', help='Path to the BibTeX file (.bib). Defaults to the pptx filename with .bib extension')
    parser.add_argument('-c', '--config', type=Path, help='Path to configuration file (optional)')
    parser.add_argument('-o', '--output', type=Path, help='Output file path (optional)')
    parser.add_argument('-s', '--start-number', type=int, default=1, 
                        help='Starting number for numeric citation styles (default: 1)')
    
    args = parser.parse_args()
    
    # Validate PowerPoint file
    if not args.pptx_file.exists():
        print(f"Error: PowerPoint file not found: {args.pptx_file}", file=sys.stderr)
        sys.exit(1)
    
    # Default bib_file to pptx_file with .bib extension if not provided
    if args.bib_file is None:
        args.bib_file = args.pptx_file.with_suffix('.bib')
    
    # Validate BibTeX file
    if not args.bib_file.exists():
        print(f"Error: BibTeX file not found: {args.bib_file}", file=sys.stderr)
        sys.exit(1)
    
    # Load configuration
    config = load_config(args.config)
    
    # Override citation_start_number from CLI if provided
    config['citation_start_number'] = args.start_number
    
    print(f"Processing: {args.pptx_file}")
    print(f"Using bibliography: {args.bib_file}")
    
    # Parse BibTeX file
    bib_parser = BibTeXParser(args.bib_file)
    print(f"Loaded {len(bib_parser.entries)} bibliography entries")
    
    # Process PowerPoint
    processor = CitationProcessor(args.pptx_file, bib_parser, config)
    processor.process()
    processor.save(args.output)
    
    # Print citation count with bold if all entries are cited
    num_citations = len(processor.cited_keys)
    num_entries = len(bib_parser.entries)
    if num_citations == num_entries:
        print(f"Found {BOLD}{num_citations}{RESET} citations")
    else:
        print(f"Found {num_citations} citations")

if __name__ == '__main__':
    main()

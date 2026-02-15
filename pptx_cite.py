#!/usr/bin/env python3
"""
PPTX Citation Manager - Replace LaTeX-style citations in PowerPoint files
"""

import re
from pathlib import Path
from typing import Dict, Optional, List

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
            fields = {'type': entry_type}
            
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
                return f"[{number if number else '?'}] {key}: Entry not found"
            elif style == 'numeric-parens':
                return f"({number if number else '?'}) {key}: Entry not found"
            else:
                return f"{key}: Entry not found"
        
        # Format based on entry type
        entry_type = entry.get('type', 'misc')
        result = []
        
        # Number (only for numeric styles)
        if number and style in ['numeric', 'numeric-parens']:
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

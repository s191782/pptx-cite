#!/usr/bin/env python3
"""
Tests for the BibTeXParser class
"""

import unittest
import tempfile
from pathlib import Path
import sys

# Add parent directory to path to import pptx_cite
sys.path.insert(0, str(Path(__file__).parent.parent))
from pptx_cite import BibTeXParser


class TestBibTeXParser(unittest.TestCase):
    """Test cases for BibTeXParser class"""
    
    def setUp(self):
        """Set up test fixtures"""
        # Create a temporary BibTeX file for testing
        self.temp_dir = tempfile.mkdtemp()
        self.bib_file = Path(self.temp_dir) / "test.bib"
        
        # Sample BibTeX content with various entry types
        self.bib_content = """
@article{smith2024,
  author = {John Smith and Jane Doe},
  title = {Test Article},
  journal = {Test Journal},
  year = {2024},
  volume = {15},
  number = {3},
  pages = {123--145}
}

@book{johnson2023,
  author = {Alice Johnson},
  title = {Test Book},
  publisher = {Academic Press},
  year = {2023},
  address = {New York},
  edition = {2nd}
}

@inproceedings{brown2022,
  author = {Bob Brown and Charlie Davis and Eve Wilson},
  title = {Conference Paper},
  booktitle = {Proceedings of the Test Conference},
  year = {2022},
  pages = {45--52}
}

@article{multiple2024,
  author = {Author One and Author Two and Author Three and Author Four},
  title = {Paper with Many Authors},
  journal = {Multi-Author Journal},
  year = {2024},
  volume = {1},
  pages = {1--10}
}

@article{dateformat2024,
  author = {Test Author},
  title = {Article with Date Field},
  journal = {Date Test Journal},
  date = {2024-05-15},
  volume = {10},
  pages = {100--110}
}

@book{curlybraces2023,
  author = {Jane {Doe} Smith},
  title = {{Test} {Book} with {Curly} {Braces}},
  publisher = {Test Publisher},
  year = {2023}
}

@article{lastname2024,
  author = {von der Berg, Hans and Smith, John and Mary Jane Watson},
  title = {Testing Last Names},
  journal = {Name Test Journal},
  year = {2024},
  pages = {1--5}
}

@phdthesis{thesis2023,
  author = {PhD Student},
  title = {Test Thesis},
  school = {Test University},
  year = {2023}
}

@incollection{chapter2022,
  author = {Chapter Author},
  title = {Test Chapter},
  booktitle = {Edited Book Title},
  editor = {Book Editor},
  publisher = {Chapter Publisher},
  year = {2022},
  pages = {10--25},
  chapter = {3}
}
"""
        
        with open(self.bib_file, 'w', encoding='utf-8') as f:
            f.write(self.bib_content)
    
    def tearDown(self):
        """Clean up test fixtures"""
        import shutil
        if Path(self.temp_dir).exists():
            shutil.rmtree(self.temp_dir)
    
    def test_parser_initialization(self):
        """Test that parser initializes correctly"""
        parser = BibTeXParser(self.bib_file)
        self.assertIsNotNone(parser.entries)
        self.assertIsInstance(parser.entries, dict)
        self.assertGreater(len(parser.entries), 0)
    
    def test_parse_basic_entry(self):
        """Test basic entry parsing"""
        parser = BibTeXParser(self.bib_file)
        entry = parser.get_entry('smith2024')
        
        self.assertIsNotNone(entry)
        self.assertEqual(entry['entry_type'], 'article')
        self.assertEqual(entry['author'], 'John Smith and Jane Doe')
        self.assertEqual(entry['title'], 'Test Article')
        self.assertEqual(entry['journal'], 'Test Journal')
        self.assertEqual(entry['year'], '2024')
        self.assertEqual(entry['volume'], '15')
        self.assertEqual(entry['number'], '3')
        self.assertEqual(entry['pages'], '123--145')
    
    def test_parse_multiple_entries(self):
        """Test that multiple entries are parsed"""
        parser = BibTeXParser(self.bib_file)
        
        self.assertIn('smith2024', parser.entries)
        self.assertIn('johnson2023', parser.entries)
        self.assertIn('brown2022', parser.entries)
        self.assertIn('multiple2024', parser.entries)
    
    def test_get_entry_nonexistent(self):
        """Test getting a nonexistent entry returns None"""
        parser = BibTeXParser(self.bib_file)
        entry = parser.get_entry('nonexistent')
        self.assertIsNone(entry)
    
    def test_get_year_from_year_field(self):
        """Test extracting year from year field"""
        parser = BibTeXParser(self.bib_file)
        entry = parser.get_entry('smith2024')
        year = parser._get_year(entry)
        self.assertEqual(year, '2024')
    
    def test_get_year_from_date_field(self):
        """Test extracting year from date field"""
        parser = BibTeXParser(self.bib_file)
        entry = parser.get_entry('dateformat2024')
        year = parser._get_year(entry)
        self.assertEqual(year, '2024')
    
    def test_clean_pages(self):
        """Test cleaning page numbers (double dash to single)"""
        parser = BibTeXParser(self.bib_file)
        cleaned = parser._clean_pages('123--145')
        self.assertEqual(cleaned, '123-145')
    
    def test_clean_field_value(self):
        """Test removing curly brackets from field values"""
        parser = BibTeXParser(self.bib_file)
        entry = parser.get_entry('curlybraces2023')
        
        # Check that curly braces have been removed
        self.assertNotIn('{', entry['author'])
        self.assertNotIn('}', entry['author'])
        self.assertNotIn('{', entry['title'])
        self.assertNotIn('}', entry['title'])
    
    def test_parse_authors(self):
        """Test parsing author string into list"""
        parser = BibTeXParser(self.bib_file)
        
        # Test simple case
        authors = parser._parse_authors('John Smith and Jane Doe')
        self.assertEqual(len(authors), 2)
        self.assertEqual(authors[0], 'John Smith')
        self.assertEqual(authors[1], 'Jane Doe')
        
        # Test three authors
        authors = parser._parse_authors('Bob Brown and Charlie Davis and Eve Wilson')
        self.assertEqual(len(authors), 3)
    
    def test_get_last_name_comma_format(self):
        """Test extracting last name from 'Last, First' format"""
        parser = BibTeXParser(self.bib_file)
        last_name = parser._get_last_name('von der Berg, Hans')
        self.assertEqual(last_name, 'von der Berg')
    
    def test_get_last_name_first_last_format(self):
        """Test extracting last name from 'First Last' format"""
        parser = BibTeXParser(self.bib_file)
        last_name = parser._get_last_name('John Smith')
        self.assertEqual(last_name, 'Smith')
        
        last_name = parser._get_last_name('Mary Jane Watson')
        self.assertEqual(last_name, 'Watson')
    
    def test_format_author_list_few_authors(self):
        """Test formatting author list with few authors"""
        parser = BibTeXParser(self.bib_file)
        
        # Two authors
        formatted = parser._format_author_list('John Smith and Jane Doe')
        self.assertEqual(formatted, 'John Smith, Jane Doe')
    
    def test_format_author_list_many_authors(self):
        """Test formatting author list with many authors uses et al."""
        parser = BibTeXParser(self.bib_file)
        
        # Four authors with max_authors=3
        formatted = parser._format_author_list(
            'Author One and Author Two and Author Three and Author Four',
            max_authors=3
        )
        self.assertEqual(formatted, 'Author One et al.')
    
    def test_format_citation_numeric(self):
        """Test formatting citation in numeric style"""
        parser = BibTeXParser(self.bib_file)
        citation = parser.format_citation('smith2024', style='numeric')
        self.assertEqual(citation, '[smith2024]')
    
    def test_format_citation_author_year(self):
        """Test formatting citation in author-year style"""
        parser = BibTeXParser(self.bib_file)
        citation = parser.format_citation('smith2024', style='author-year')
        self.assertIn('Smith', citation)
        self.assertIn('2024', citation)
        self.assertIn('et al.', citation)  # Two authors
    
    def test_format_citation_author(self):
        """Test formatting citation in author style"""
        parser = BibTeXParser(self.bib_file)
        citation = parser.format_citation('smith2024', style='author')
        self.assertIn('Smith', citation)
        self.assertIn('et al.', citation)
        self.assertNotIn('2024', citation)
    
    def test_format_citation_nonexistent(self):
        """Test formatting citation for nonexistent entry"""
        parser = BibTeXParser(self.bib_file)
        citation = parser.format_citation('nonexistent', style='numeric')
        self.assertEqual(citation, '[nonexistent?]')
    
    def test_format_bibliography_entry_article(self):
        """Test formatting bibliography entry for article"""
        parser = BibTeXParser(self.bib_file)
        entry = parser.format_bibliography_entry('smith2024', number=1, style='numeric')
        
        self.assertIn('[1]', entry)
        self.assertIn('John Smith', entry)
        self.assertIn('Test Article', entry)
        self.assertIn('Test Journal', entry)
        self.assertIn('2024', entry)
        self.assertIn('vol. 15', entry)
        self.assertIn('no. 3', entry)
        self.assertIn('pp. 123-145', entry)  # Should have single dash
    
    def test_format_bibliography_entry_book(self):
        """Test formatting bibliography entry for book"""
        parser = BibTeXParser(self.bib_file)
        entry = parser.format_bibliography_entry('johnson2023', number=2, style='numeric')
        
        self.assertIn('[2]', entry)
        self.assertIn('Alice Johnson', entry)
        self.assertIn('Test Book', entry)
        self.assertIn('Academic Press', entry)
        self.assertIn('2023', entry)
        self.assertIn('2nd ed.', entry)
        self.assertIn('New York', entry)
    
    def test_format_bibliography_entry_inproceedings(self):
        """Test formatting bibliography entry for conference paper"""
        parser = BibTeXParser(self.bib_file)
        entry = parser.format_bibliography_entry('brown2022', number=3, style='numeric')
        
        self.assertIn('[3]', entry)
        self.assertIn('Bob Brown', entry)
        self.assertIn('Conference Paper', entry)
        self.assertIn('Proceedings of the Test Conference', entry)
        self.assertIn('2022', entry)
    
    def test_format_bibliography_entry_thesis(self):
        """Test formatting bibliography entry for thesis"""
        parser = BibTeXParser(self.bib_file)
        entry = parser.format_bibliography_entry('thesis2023', number=4, style='numeric')
        
        self.assertIn('[4]', entry)
        self.assertIn('PhD Student', entry)
        self.assertIn('Test Thesis', entry)
        self.assertIn('Ph.D. dissertation', entry)
        self.assertIn('Test University', entry)
    
    def test_format_bibliography_entry_incollection(self):
        """Test formatting bibliography entry for book chapter"""
        parser = BibTeXParser(self.bib_file)
        entry = parser.format_bibliography_entry('chapter2022', number=5, style='numeric')
        
        self.assertIn('[5]', entry)
        self.assertIn('Chapter Author', entry)
        self.assertIn('Test Chapter', entry)
        self.assertIn('Edited Book Title', entry)
        self.assertIn('ch. 3', entry)
        self.assertIn('pp. 10-25', entry)
    
    def test_format_bibliography_entry_numeric_parens(self):
        """Test formatting bibliography entry with parentheses style"""
        parser = BibTeXParser(self.bib_file)
        entry = parser.format_bibliography_entry('smith2024', number=1, style='numeric-parens')
        
        self.assertIn('(1)', entry)
        self.assertNotIn('[1]', entry)
    
    def test_format_bibliography_entry_nonexistent(self):
        """Test formatting bibliography entry for nonexistent entry"""
        parser = BibTeXParser(self.bib_file)
        entry = parser.format_bibliography_entry('nonexistent', number=99, style='numeric')
        
        self.assertIn('[99]', entry)
        self.assertIn('nonexistent', entry)
        self.assertIn('Entry not found', entry)
    
    def test_empty_author_field(self):
        """Test handling of empty/missing author field"""
        # Create a temporary bib file with no author
        temp_bib = Path(self.temp_dir) / "no_author.bib"
        with open(temp_bib, 'w') as f:
            f.write('@article{noauthor, title={No Author}, journal={Test}, year={2024}}')
        
        parser = BibTeXParser(temp_bib)
        authors = parser._parse_authors('')
        self.assertEqual(authors, [])
        
        # Test citation formatting with no author
        citation = parser.format_citation('noauthor', style='author-year')
        self.assertIn('Unknown', citation)
    
    def test_format_author_list_no_et_al(self):
        """Test formatting author list without et al."""
        parser = BibTeXParser(self.bib_file)
        
        # Four authors without et al.
        formatted = parser._format_author_list(
            'Author One and Author Two and Author Three and Author Four',
            max_authors=3,
            use_et_al=False
        )
        self.assertIn('Author One', formatted)
        self.assertIn('Author Four', formatted)
        self.assertNotIn('et al.', formatted)


class TestBibTeXParserEdgeCases(unittest.TestCase):
    """Test edge cases and error handling"""
    
    def test_parse_nested_braces(self):
        """Test parsing entries with nested braces"""
        with tempfile.TemporaryDirectory() as temp_dir:
            bib_file = Path(temp_dir) / "nested.bib"
            content = """
@article{nested2024,
  author = {Test Author},
  title = {Title with {Nested {Braces} Inside}},
  journal = {Test Journal},
  year = {2024}
}
"""
            with open(bib_file, 'w') as f:
                f.write(content)
            
            parser = BibTeXParser(bib_file)
            entry = parser.get_entry('nested2024')
            self.assertIsNotNone(entry)
            # Curly braces should be removed
            self.assertNotIn('{', entry['title'])
            self.assertNotIn('}', entry['title'])
    
    def test_parse_quoted_fields(self):
        """Test parsing fields with quoted values"""
        with tempfile.TemporaryDirectory() as temp_dir:
            bib_file = Path(temp_dir) / "quoted.bib"
            content = """
@article{quoted2024,
  author = "Test Author",
  title = "Quoted Title",
  journal = "Quoted Journal",
  year = "2024"
}
"""
            with open(bib_file, 'w') as f:
                f.write(content)
            
            parser = BibTeXParser(bib_file)
            entry = parser.get_entry('quoted2024')
            self.assertIsNotNone(entry)
            self.assertEqual(entry['author'], 'Test Author')
            self.assertEqual(entry['title'], 'Quoted Title')
    
    def test_parse_bare_numeric_values(self):
        """Test parsing bare numeric values"""
        with tempfile.TemporaryDirectory() as temp_dir:
            bib_file = Path(temp_dir) / "numeric.bib"
            content = """
@article{numeric2024,
  author = {Test Author},
  title = {Test Title},
  year = 2024,
  volume = 15
}
"""
            with open(bib_file, 'w') as f:
                f.write(content)
            
            parser = BibTeXParser(bib_file)
            entry = parser.get_entry('numeric2024')
            self.assertIsNotNone(entry)
            self.assertEqual(entry['year'], '2024')
            self.assertEqual(entry['volume'], '15')
    
    def test_clean_author_name_with_extra_spaces(self):
        """Test cleaning author names with extra whitespace"""
        with tempfile.TemporaryDirectory() as temp_dir:
            bib_file = Path(temp_dir) / "spaces.bib"
            content = """
@article{spaces2024,
  author = {  John   Smith  },
  title = {Test},
  year = {2024}
}
"""
            with open(bib_file, 'w') as f:
                f.write(content)
            
            parser = BibTeXParser(bib_file)
            entry = parser.get_entry('spaces2024')
            # Extra spaces should be cleaned
            self.assertEqual(entry['author'].strip(), 'John   Smith')
            
            cleaned = parser._clean_author_name('  John   Smith  ')
            self.assertEqual(cleaned, 'John Smith')


if __name__ == '__main__':
    unittest.main()

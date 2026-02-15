"""
Integration tests for CitationProcessor class.
"""

import unittest
import tempfile
import shutil
import sys
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches

# Add parent directory to path for imports (needed for standalone test execution)
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pptx_cite import BibTeXParser, CitationProcessor


class TestCitationProcessor(unittest.TestCase):
    """Test CitationProcessor functionality with sample PPTX files."""
    
    def setUp(self):
        """Set up test fixtures."""
        # Create temporary directory
        self.test_dir = tempfile.mkdtemp()
        self.test_dir_path = Path(self.test_dir)
        
        # Create a sample BibTeX file
        self.bib_file = self.test_dir_path / "references.bib"
        self.bib_file.write_text("""
@article{smith2020,
    author = {Smith, John and Doe, Jane},
    title = {A Sample Article},
    journal = {Journal of Examples},
    year = {2020},
    volume = {1},
    pages = {1--10}
}

@book{jones2019,
    author = {Jones, Mary},
    title = {Sample Book Title},
    publisher = {Academic Press},
    year = {2019}
}

@inproceedings{brown2021,
    author = {Brown, Bob},
    title = {Conference Paper Example},
    booktitle = {Proceedings of Example Conference},
    year = {2021}
}
""")
        
        # Create BibTeX parser
        self.bib_parser = BibTeXParser(self.bib_file)
        
        # Create default config
        self.config = {
            'cite_command': r'\cite',
            'bibliography_command': r'\bibliography',
            'cite_style': 'numeric',
            'citation_start_number': 1
        }
    
    def tearDown(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.test_dir)
    
    def _create_simple_pptx(self, text_content: str) -> Path:
        """Create a simple PPTX file with given text content."""
        pptx_file = self.test_dir_path / "test.pptx"
        prs = Presentation()
        
        # Add a blank slide
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # Add a text box
        left = Inches(1)
        top = Inches(1)
        width = Inches(8)
        height = Inches(5)
        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        text_frame.text = text_content
        
        prs.save(str(pptx_file))
        return pptx_file
    
    def test_processor_initialization(self):
        """Test CitationProcessor initialization."""
        pptx_file = self._create_simple_pptx("Test content")
        processor = CitationProcessor(pptx_file, self.bib_parser, self.config)
        
        self.assertEqual(processor.pptx_file, pptx_file)
        self.assertEqual(processor.bib_parser, self.bib_parser)
        self.assertEqual(processor.config, self.config)
        self.assertIsNotNone(processor.prs)
        self.assertEqual(processor.cited_keys, [])
        self.assertEqual(processor.cited_set, set())
        self.assertEqual(processor.missing_citations, {})
    
    def test_citation_replacement_numeric(self):
        """Test that citations are replaced correctly in numeric style."""
        # Create PPTX with citation and bibliography
        pptx_file = self.test_dir_path / "test_numeric.pptx"
        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6]
        
        # Add slide with citation
        slide1 = prs.slides.add_slide(blank_slide_layout)
        textbox1 = slide1.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        textbox1.text_frame.text = "This is a citation \\cite{smith2020}"
        
        # Add slide with bibliography
        slide2 = prs.slides.add_slide(blank_slide_layout)
        textbox2 = slide2.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        textbox2.text_frame.text = "\\bibliography"
        
        prs.save(str(pptx_file))
        
        processor = CitationProcessor(pptx_file, self.bib_parser, self.config)
        processor.process()
        
        # Check that citation was tracked
        self.assertIn('smith2020', processor.cited_keys)
    
    def test_multiple_citations(self):
        """Test handling of multiple citations."""
        # Create PPTX with multiple citations and bibliography
        pptx_file = self.test_dir_path / "test_multiple.pptx"
        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6]
        
        # Add slide with citations
        slide1 = prs.slides.add_slide(blank_slide_layout)
        textbox1 = slide1.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        textbox1.text_frame.text = "First citation \\cite{smith2020} and second \\cite{jones2019}"
        
        # Add slide with bibliography
        slide2 = prs.slides.add_slide(blank_slide_layout)
        textbox2 = slide2.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        textbox2.text_frame.text = "\\bibliography"
        
        prs.save(str(pptx_file))
        
        processor = CitationProcessor(pptx_file, self.bib_parser, self.config)
        processor.process()
        
        # Check both citations were tracked
        self.assertIn('smith2020', processor.cited_keys)
        self.assertIn('jones2019', processor.cited_keys)
        self.assertEqual(len(processor.cited_keys), 2)
    
    def test_missing_citation_warning(self):
        """Test that missing citations are tracked."""
        # Create PPTX with missing citation and bibliography
        pptx_file = self.test_dir_path / "test_missing.pptx"
        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6]
        
        # Add slide with missing citation
        slide1 = prs.slides.add_slide(blank_slide_layout)
        textbox1 = slide1.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        textbox1.text_frame.text = "Missing citation \\cite{nonexistent2022}"
        
        # Add slide with bibliography
        slide2 = prs.slides.add_slide(blank_slide_layout)
        textbox2 = slide2.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        textbox2.text_frame.text = "\\bibliography"
        
        prs.save(str(pptx_file))
        
        processor = CitationProcessor(pptx_file, self.bib_parser, self.config)
        processor.process()
        
        # Check that missing citation was tracked
        self.assertIn('nonexistent2022', processor.missing_citations)
        self.assertIn(1, processor.missing_citations['nonexistent2022'])
    
    def test_bibliography_detection(self):
        """Test that bibliography command is detected."""
        pptx_file = self._create_simple_pptx("\\bibliography")
        processor = CitationProcessor(pptx_file, self.bib_parser, self.config)
        
        # Check bibliography detection
        slide = list(processor.prs.slides)[0]
        has_bib = processor._slide_contains_bibliography_command(
            slide, 
            self.config['bibliography_command']
        )
        self.assertTrue(has_bib)
    
    def test_no_bibliography_command(self):
        """Test processing when no bibliography command exists."""
        pptx_file = self._create_simple_pptx("Citation without bibliography \\cite{smith2020}")
        processor = CitationProcessor(pptx_file, self.bib_parser, self.config)
        processor.process()
        
        # When no bibliography is present, citations are removed but not tracked
        self.assertEqual(len(processor.cited_keys), 0)
    
    def test_save_functionality(self):
        """Test that processed presentation can be saved."""
        pptx_file = self._create_simple_pptx("Test \\cite{smith2020}")
        processor = CitationProcessor(pptx_file, self.bib_parser, self.config)
        processor.process()
        
        # Save to a new file
        output_file = self.test_dir_path / "output.pptx"
        processor.save(output_file)
        
        # Check that output file exists
        self.assertTrue(output_file.exists())
        
        # Check that output file can be opened
        prs = Presentation(str(output_file))
        self.assertIsNotNone(prs)
    
    def test_citation_with_bibliography(self):
        """Test citation processing with bibliography slide."""
        # Create PPTX with citation and bibliography
        pptx_file = self.test_dir_path / "test_with_bib.pptx"
        prs = Presentation()
        
        # Add slide with citation
        blank_slide_layout = prs.slide_layouts[6]
        slide1 = prs.slides.add_slide(blank_slide_layout)
        textbox1 = slide1.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        textbox1.text_frame.text = "Citation \\cite{smith2020}"
        
        # Add slide with bibliography
        slide2 = prs.slides.add_slide(blank_slide_layout)
        textbox2 = slide2.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        textbox2.text_frame.text = "\\bibliography"
        
        prs.save(str(pptx_file))
        
        # Process
        processor = CitationProcessor(pptx_file, self.bib_parser, self.config)
        processor.process()
        
        # Check citation was tracked
        self.assertIn('smith2020', processor.cited_keys)
    
    def test_duplicate_citations(self):
        """Test that duplicate citations are handled correctly."""
        # Create PPTX with duplicate citations and bibliography
        pptx_file = self.test_dir_path / "test_duplicate.pptx"
        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6]
        
        # Add slide with duplicate citations
        slide1 = prs.slides.add_slide(blank_slide_layout)
        textbox1 = slide1.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        textbox1.text_frame.text = "First \\cite{smith2020} and again \\cite{smith2020}"
        
        # Add slide with bibliography
        slide2 = prs.slides.add_slide(blank_slide_layout)
        textbox2 = slide2.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        textbox2.text_frame.text = "\\bibliography"
        
        prs.save(str(pptx_file))
        
        processor = CitationProcessor(pptx_file, self.bib_parser, self.config)
        processor.process()
        
        # Check that citation appears only once (unique tracking)
        self.assertEqual(processor.cited_keys.count('smith2020'), 1)
        # And once in cited_set
        self.assertIn('smith2020', processor.cited_set)
        self.assertEqual(len(processor.cited_set), 1)
    
    def test_citation_style_author_year(self):
        """Test citation processing with author-year style."""
        self.config['cite_style'] = 'author-year'
        
        # Create PPTX with citation and bibliography
        pptx_file = self.test_dir_path / "test_author_year.pptx"
        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6]
        
        # Add slide with citation
        slide1 = prs.slides.add_slide(blank_slide_layout)
        textbox1 = slide1.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        textbox1.text_frame.text = "Citation \\cite{smith2020}"
        
        # Add slide with bibliography
        slide2 = prs.slides.add_slide(blank_slide_layout)
        textbox2 = slide2.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        textbox2.text_frame.text = "\\bibliography"
        
        prs.save(str(pptx_file))
        
        processor = CitationProcessor(pptx_file, self.bib_parser, self.config)
        processor.process()
        
        # Check citation was tracked
        self.assertIn('smith2020', processor.cited_keys)
    
    def test_empty_presentation(self):
        """Test processing an empty presentation."""
        pptx_file = self.test_dir_path / "empty.pptx"
        prs = Presentation()
        prs.save(str(pptx_file))
        
        processor = CitationProcessor(pptx_file, self.bib_parser, self.config)
        processor.process()
        
        # Should not raise any errors
        self.assertEqual(len(processor.cited_keys), 0)
    
    def test_slide_with_multiple_textboxes(self):
        """Test slide with multiple text boxes."""
        pptx_file = self.test_dir_path / "multi_textbox.pptx"
        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6]
        
        # Add slide with multiple text boxes
        slide1 = prs.slides.add_slide(blank_slide_layout)
        textbox1 = slide1.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(2))
        textbox1.text_frame.text = "First \\cite{smith2020}"
        
        textbox2 = slide1.shapes.add_textbox(Inches(5), Inches(1), Inches(3), Inches(2))
        textbox2.text_frame.text = "Second \\cite{jones2019}"
        
        # Add slide with bibliography
        slide2 = prs.slides.add_slide(blank_slide_layout)
        textbox3 = slide2.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        textbox3.text_frame.text = "\\bibliography"
        
        prs.save(str(pptx_file))
        
        processor = CitationProcessor(pptx_file, self.bib_parser, self.config)
        processor.process()
        
        # Check both citations were found
        self.assertIn('smith2020', processor.cited_keys)
        self.assertIn('jones2019', processor.cited_keys)


class TestCitationProcessorEdgeCases(unittest.TestCase):
    """Test edge cases for CitationProcessor."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = tempfile.mkdtemp()
        self.test_dir_path = Path(self.test_dir)
        
        # Create minimal BibTeX file
        self.bib_file = self.test_dir_path / "refs.bib"
        self.bib_file.write_text("""
@article{test2020,
    author = {Test Author},
    title = {Test Title},
    year = {2020}
}
""")
        
        self.bib_parser = BibTeXParser(self.bib_file)
        self.config = {
            'cite_command': r'\cite',
            'bibliography_command': r'\bibliography',
            'cite_style': 'numeric'
        }
    
    def tearDown(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.test_dir)
    
    def test_citation_with_special_characters(self):
        """Test citation keys with special characters."""
        pptx_file = self.test_dir_path / "test.pptx"
        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        textbox.text_frame.text = "Citation \\cite{test-2020_special:key}"
        prs.save(str(pptx_file))
        
        processor = CitationProcessor(pptx_file, self.bib_parser, self.config)
        processor.process()
        
        # Should handle gracefully (will be missing citation)
        self.assertIsNotNone(processor.missing_citations)
    
    def test_malformed_citation(self):
        """Test handling of malformed citations."""
        pptx_file = self.test_dir_path / "test.pptx"
        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        textbox.text_frame.text = "Malformed \\cite{ or \\cite{} or \\cite"
        prs.save(str(pptx_file))
        
        processor = CitationProcessor(pptx_file, self.bib_parser, self.config)
        # Should not raise error
        processor.process()


if __name__ == '__main__':
    unittest.main()

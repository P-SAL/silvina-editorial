"""
word_counter.py
Accurate Word character/word/paragraph counting using COM automation.
Part of Silvina Editorial Assistant v0.8
"""

import os
from typing import Dict, Optional

# Try to import win32com (Windows only)
try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False


class WordCounter:
    """Get accurate Word statistics using COM automation."""
    
    def __init__(self):
        """Initialize Word counter."""
        self.word_app = None
        self.doc = None
    
    
    
    def get_accurate_counts(self, docx_path: str) -> Optional[Dict[str, int]]:
        """
        Get accurate character/word/paragraph counts from Word.
        
        Args:
            docx_path: Path to .docx file
            
        Returns:
            Dict with 'char_count', 'word_count', 'paragraph_count'
            Returns None if win32com not available or error occurs
        """
        if not WIN32COM_AVAILABLE:
            return None
        
        if not os.path.exists(docx_path):
            return None
        
        # Try up to 2 times (in case Word is busy)
        for attempt in range(2):
            try:
                # Kill any existing Word instances that might interfere
                if attempt == 1:
                    try:
                        import pythoncom
                        pythoncom.CoUninitialize()
                        pythoncom.CoInitialize()
                    except:
                        pass
                
                # Launch Word with DispatchEx for isolated instance
                self.word_app = win32com.client.DispatchEx("Word.Application")
                self.word_app.DisplayAlerts = 0  # Suppress all alerts
                
                try:
                    self.word_app.Visible = False
                except:
                    pass  # Some Windows versions don't allow setting Visible
                            
                
                # Open document (read-only, no updates)
                self.doc = self.word_app.Documents.Open(
                    os.path.abspath(docx_path),
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    ConfirmConversions=False
                )
                
                # Get counts
                char_count = self._get_character_count()
                word_count = self._get_word_count()
                paragraph_count = self._get_paragraph_count()
                
                # Close document
                self.doc.Close(False)
                self.word_app.Quit()
                
                # Clean up COM
                self.word_app = None
                self.doc = None
                
                return {
                    'char_count': char_count,
                    'word_count': word_count,
                    'paragraph_count': paragraph_count
                }
                
            except Exception as e:
                # Clean up on error
                try:
                    if self.doc:
                        self.doc.Close(False)
                except:
                    pass
                
                try:
                    if self.word_app:
                        self.word_app.Quit()
                except:
                    pass
                
                # Reset
                self.word_app = None
                self.doc = None
                
                # If this was the last attempt, report error
                if attempt == 1:
                    print(f"   ⚠ No se pudo obtener conteos precisos de Word")
                    print(f"   ℹ️  Usando conteo aproximado desde python-docx")
                    return None
                
                # Otherwise retry
                import time
                time.sleep(0.5)
        
        return None
    
    def _get_character_count(self) -> int:
        """Get accurate Word character count (including footnotes/endnotes)."""
        if not self.doc:
            return 0
        
        try:
            # Base character count
            total = self.doc.Characters.Count
            
            # Add footnotes
            for fn in self.doc.Footnotes:
                total += len(fn.Range.Text)
            
            # Add endnotes
            for en in self.doc.Endnotes:
                total += len(en.Range.Text)
            
            return total
        except:
            return 0
    
    def _get_word_count(self) -> int:
        """Get accurate Word word count (including footnotes/endnotes)."""
        if not self.doc:
            return 0
        
        try:
            # Start with main document word count
            total = self.doc.ComputeStatistics(0)  # wdStatisticWords
            
            # ADD footnote words (might be excluded by ComputeStatistics)
            try:
                for fn in self.doc.Footnotes:
                    total += fn.Range.ComputeStatistics(0)
            except:
                pass
            
            # ADD endnote words
            try:
                for en in self.doc.Endnotes:
                    total += en.Range.ComputeStatistics(0)
            except:
                pass
            
            return total
        except:
            return 0
  
    def _get_paragraph_count(self) -> int:
        """Get Word paragraph count INCLUDING footnotes, endnotes, text boxes."""
        if not self.doc:
            return 0

        try:
            # 4 = wdStatisticParagraphs
            return self.doc.ComputeStatistics(4)
        except Exception as e:
            print(f"   ⚠ Error contando párrafos Word: {e}")
            return 0

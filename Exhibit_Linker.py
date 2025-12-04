import os
import re
import win32com.client
import tkinter as tk
from tkinter import filedialog, messagebox, ttk as tk_ttk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import difflib
import pythoncom
import time
import tempfile
import shutil
from pypdf import PdfReader, PdfWriter
from pypdf.generic import DictionaryObject, ArrayObject, TextStringObject

class WordAutoLinkerCOM:
    def __init__(self):
        self.word_app = None
        self.doc = None
        self.original_doc = None
        self.target_folder = None
        self.doc_folder = None
        self.original_doc_path = None
        self.use_black_hyperlinks = False
        self.original_word_settings = {}
        
        # Bates mode settings
        self.bates_mode = False
        self.bates_prefix = ""

        # Page number automation settings
        self.page_automation_enabled = False
        self.exemplary_citation = ""
        self.exemplary_page_number = None
        self.page_pattern_regex = None
        self.exhibit_group_index = None
        self.page_group_index = None
        
        
        # Original exhibit patterns
        self.exhibit_patterns = [
            r'Ex\.\s*(\d+[A-Z]?)',        # Ex. 1, Ex. 2, Ex. 1A, Ex. 2B
            r'Ex\.\s*([A-Z])',            # Ex. A, Ex. B
            r'Exhibit\s*(\d+[A-Z]?)',     # Exhibit 1, Exhibit 2, Exhibit 1A, Exhibit 2B
            r'Exhibit\s*([A-Z])',         # Exhibit A, Exhibit B
            
            r'Ex\.(\d+[A-Z]?)',           # Ex.1, Ex.2A (no space)
            r'Ex\.([A-Z])',               # Ex.A, Ex.B (no space)
            r'Ex\s+(\d+[A-Z]?)',          # Ex 1, Ex 2A (space instead of period)
            r'Ex\s+([A-Z])',              # Ex A, Ex B (space instead of period)
            r'Ex_(\d+[A-Z]?)',            # Ex_1, Ex_2A (underscore)
            r'Ex_([A-Z])',                # Ex_A, Ex_B (underscore)
        ]
        
        # Track hyperlinks we create for PDF processing
        self.created_hyperlinks = []
        
        # Bates PDF mapping (filename -> starting page number)
        self.bates_pdf_map = {}

    def set_black_hyperlinks(self, use_black):
        """Set whether to use black hyperlinks"""
        self.use_black_hyperlinks = use_black
        print(f"Black hyperlinks mode: {'enabled' if use_black else 'disabled'}")

    def set_bates_mode(self, enabled, prefix=""):
        """Set Bates mode on/off with prefix"""
        self.bates_mode = enabled
        self.bates_prefix = prefix.strip()
        if self.bates_mode:
            print(f"Bates mode enabled with prefix: '{self.bates_prefix}'")
            # Build the PDF mapping when Bates mode is enabled
            self.build_bates_pdf_map()
        else:
            print("Bates mode disabled - using exhibit mode")

    def set_page_automation(self, enabled, exemplary_citation="", exemplary_page_number=None):
        """Configure page number automation"""
        self.page_automation_enabled = enabled
        self.exemplary_citation = exemplary_citation.strip()
        self.exemplary_page_number = exemplary_page_number
        
        if enabled and self.exemplary_citation and self.exemplary_page_number:
            self.build_page_pattern()
            print(f"Page automation enabled:")
            print(f"  Exemplary citation: '{self.exemplary_citation}'")
            print(f"  Exemplary page: {self.exemplary_page_number}")
            print(f"  Generated pattern: {self.page_pattern_regex}")
        else:
            self.page_pattern_regex = None
            self.exhibit_group_index = None
            self.page_group_index = None
            print("Page automation disabled")

    def build_page_pattern(self):
        """Build flexible regex pattern for legal citation formats"""
        try:
            citation = self.exemplary_citation
            page_num = str(self.exemplary_page_number)
            
            print(f"Building flexible pattern from: '{citation}' with page {page_num}")
            
            # Step 1: Find exhibit identifier using existing patterns
            exhibit_id = None
            for pattern in self.exhibit_patterns:
                match = re.search(pattern, citation, re.IGNORECASE)
                if match:
                    exhibit_id = match.group(1)
                    print(f"  Found exhibit ID: '{exhibit_id}'")
                    break
            
            if not exhibit_id:
                print("  ERROR: Could not find exhibit identifier in exemplary citation")
                return False
            
            # Step 2: Find page number with common legal citation patterns
            page_patterns = [
                rf'\bat\s+p\.?\s*{re.escape(page_num)}\b',      # "at p. 25", "at p 25"
                rf'\bat\s+pp\.?\s*{re.escape(page_num)}\b',     # "at pp. 25", "at pp 25" 
                rf'\bat\s+{re.escape(page_num)}\b',             # "at 25"
                rf'\bp\.?\s*{re.escape(page_num)}\b',           # "p. 25", "p 25"
                rf'\bpp\.?\s*{re.escape(page_num)}\b',          # "pp. 25", "pp 25"
                rf'\bpage\s+{re.escape(page_num)}\b',           # "page 25"
                rf'\bpages?\s+{re.escape(page_num)}\b',         # "pages 25"
                rf'\b{re.escape(page_num)}\b'                   # just "25" (fallback)
            ]
            
            page_match_info = None
            for i, pattern in enumerate(page_patterns):
                matches = list(re.finditer(pattern, citation, re.IGNORECASE))
                if matches:
                    # Use the last match (most likely the page reference)
                    page_match_info = {
                        'match': matches[-1],
                        'pattern_type': i,
                        'full_match': matches[-1].group(0)
                    }
                    print(f"  Found page reference: '{page_match_info['full_match']}' using pattern type {i}")
                    break
            
            if not page_match_info:
                print(f"  ERROR: Could not find page number '{page_num}' in citation")
                return False
            
            # Step 3: Build flexible pattern that matches legal citation formats
            # Pattern structure: Ex. [ID] [misc text] [page reference with NUMBER]
            
            # Create exhibit capture group (flexible for numbers, letters, combinations)
            exhibit_capture = r'(\d+[A-Za-z]*|[A-Za-z]+\d*|[A-Za-z]+)'
            
            # Create page number capture group based on the pattern type found
            pattern_type = page_match_info['pattern_type']
            if pattern_type == 0:  # "at p. NUMBER" - cover both p. and pp.
                page_capture = r'at\s+pp?\.?\s*(\d+)'
            elif pattern_type == 1:  # "at pp. NUMBER" - cover both p. and pp.
                page_capture = r'at\s+pp?\.?\s*(\d+)'
            elif pattern_type == 2:  # "at NUMBER"
                page_capture = r'at\s+(\d+)'
            elif pattern_type == 3:  # "p. NUMBER"
                page_capture = r'p\.?\s*(\d+)'
            elif pattern_type == 4:  # "pp. NUMBER"
                page_capture = r'pp\.?\s*(\d+)'
            elif pattern_type == 5:  # "page NUMBER"
                page_capture = r'page\s+(\d+)'
            elif pattern_type == 6:  # "pages NUMBER"
                page_capture = r'pages?\s+(\d+)'
            else:  # just NUMBER (fallback)
                page_capture = r'(\d+)'
            

            # Use negative lookahead to stop at next exhibit boundary, allow more flexible middle content
            full_pattern = rf'(?:Ex\.|Exhibit)\s*{exhibit_capture}(?:(?!(?:Ex\.|Exhibit)\s*\w)[^.]*?){page_capture}'
            loose_pattern = rf'{page_capture}'
            
            print(f"  Generated strict pattern: {full_pattern}")
            print(f"  Generated loose pattern: {loose_pattern}")
            
            print(f"  Generated flexible pattern: {full_pattern}")
            
            #Step 4: Test the pattern
            test_match = re.search(full_pattern, citation, re.IGNORECASE)
            if not test_match:
                # Try with DOTALL flag for multiline text
                print("  Trying with DOTALL flag for multiline matching...")
                test_match = re.search(full_pattern, citation, re.IGNORECASE | re.DOTALL)
            if test_match and len(test_match.groups()) >= 2:
                found_exhibit = test_match.group(1)
                found_page = test_match.group(2)
                
                print(f"  Pattern test successful:")
                print(f"    Group 1 (exhibit): '{found_exhibit}'")
                print(f"    Group 2 (page): '{found_page}'")
                
                # Verify the groups match our expected values
                if (found_exhibit.upper() == exhibit_id.upper() and 
                    found_page == page_num):
                    self.page_pattern_regex = full_pattern
                    self.exhibit_group_index = 1
                    self.page_group_index = 2
                    print("  ✓ Flexible pattern validation successful!")
                    return True
                else:
                    print("  ✗ Pattern validation failed - groups don't match expected values")
                    print(f"    Expected: exhibit='{exhibit_id}', page='{page_num}'")
                    print(f"    Found: exhibit='{found_exhibit}', page='{found_page}'")
                    return False
            else:
                print("  ✗ Pattern test failed - doesn't match original citation or insufficient groups")
                return False
                
        except Exception as e:
            print(f"Error building flexible pattern: {e}")
            import traceback
            traceback.print_exc()
            return False

    def find_page_number_in_text(self, text):
        """Extract exhibit ID and page number using the pattern - ENHANCED VERSION"""
        if not self.page_automation_enabled or not self.page_pattern_regex:
            print(f"    Page automation disabled or no pattern (enabled={self.page_automation_enabled}, pattern={self.page_pattern_regex is not None})")
            return None, None
        
        try:
            print(f"    Testing pattern '{self.page_pattern_regex}' against text: '{text}'")
            
            # Find ALL pattern matches, not just the first one
            matches = list(re.finditer(self.page_pattern_regex, text, re.IGNORECASE))
            
            if matches:
                print(f"    Found {len(matches)} pattern matches:")
                for i, match in enumerate(matches):
                    exhibit_id = match.group(self.exhibit_group_index)
                    page_number = int(match.group(self.page_group_index))
                    match_text = match.group(0)
                    print(f"      Match {i+1}: '{match_text}' -> Exhibit '{exhibit_id}' Page {page_number}")
                
                # Return all matches for the caller to decide which one to use
                results = []
                for match in matches:
                    exhibit_id = match.group(self.exhibit_group_index)
                    page_number = int(match.group(self.page_group_index))
                    results.append((exhibit_id, page_number, match))
                
                return results
            else:
                print(f"    ✗ Pattern did not match text")
                return []
                
        except Exception as e:
            print(f"    ✗ Error in page number extraction: {e}")
            import traceback
            traceback.print_exc()
        
        return []

    def build_bates_pdf_map(self):
        """Build mapping of Bates PDFs to their starting page numbers"""
        self.bates_pdf_map = {}
        
        if not self.target_folder or not self.bates_prefix:
            return
        
        try:
            files_in_folder = os.listdir(self.target_folder)
            bates_files = []
            
            # Find all PDF files matching the Bates prefix pattern
            bates_pattern = rf'^{re.escape(self.bates_prefix)}(\d+)\.pdf$'
            
            for filename in files_in_folder:
                match = re.match(bates_pattern, filename, re.IGNORECASE)
                if match:
                    bates_number = int(match.group(1))
                    full_path = os.path.join(self.target_folder, filename)
                    bates_files.append((bates_number, filename, full_path))
            
            # Sort by Bates number
            bates_files.sort(key=lambda x: x[0])
            
            # Build the mapping
            for i, (bates_number, filename, full_path) in enumerate(bates_files):
                self.bates_pdf_map[bates_number] = {
                    'filename': filename,
                    'path': full_path,
                    'start_page': bates_number
                }
            
            print(f"Built Bates PDF map for {len(bates_files)} files:")
            for bates_num, info in self.bates_pdf_map.items():
                print(f"  {info['filename']} starts at page {bates_num}")
                
        except Exception as e:
            print(f"Error building Bates PDF map: {e}")
    
    def find_bates_pdf_and_page(self, bates_number):
        """Find which PDF contains the given Bates number and calculate the page"""
        if not self.bates_pdf_map:
            return None, None
        
        # Sort PDF starting numbers in descending order
        sorted_starts = sorted(self.bates_pdf_map.keys(), reverse=True)
        
        # Find the PDF that contains this Bates number
        for start_page in sorted_starts:
            if bates_number >= start_page:
                pdf_info = self.bates_pdf_map[start_page]
                # Calculate the page within this PDF (1-based)
                page_in_pdf = bates_number - start_page + 1
                
                print(f"Bates {bates_number} -> {pdf_info['filename']} page {page_in_pdf}")
                return pdf_info['path'], page_in_pdf
        
        print(f"No PDF found for Bates number {bates_number}")
        return None, None
    
    def get_bates_patterns(self):
        """Get regex patterns for Bates numbering"""
        if not self.bates_prefix:
            return []
        
        escaped_prefix = re.escape(self.bates_prefix)
        return [
            rf'{escaped_prefix}(\d+)',  # SMITH_0001, SMITH_123, etc.
        ]

    def initialize_word(self):
        """Initialize Word COM application with enhanced error handling and corporate performance optimizations"""
        if self.word_app is not None:
            return True
            
        try:
            print("Initializing Word COM application...")
            
            # Clear COM cache first to fix CLSIDtoPackageMap errors
            try:
                gen_py_path = os.path.join(tempfile.gettempdir(), "gen_py")
                if os.path.exists(gen_py_path):
                    shutil.rmtree(gen_py_path)
                    print("COM cache cleared")
            except Exception as cache_error:
                print(f"Could not clear COM cache: {cache_error}")
            
            # Initialize COM
            pythoncom.CoInitialize()
            
            # FIXED: Always create a NEW Word instance to ensure clean settings
            # This prevents connecting to an instance that already has disabled spell check
            try:
                print("Creating new Word instance (to ensure clean settings)...")
                self.word_app = win32com.client.Dispatch("Word.Application")
                print("Created new Word instance with late binding")
            except Exception as dispatch_error:
                # Try dynamic dispatch as fallback
                print("Trying dynamic dispatch as fallback...")
                self.word_app = win32com.client.dynamic.Dispatch("Word.Application")
                print("Created new Word instance with dynamic dispatch")
            
            # Keep Word hidden but responsive
            self.word_app.Visible = False
            self.word_app.DisplayAlerts = False
            
            # CRITICAL: SAVE ORIGINAL SETTINGS IMMEDIATELY AFTER CONNECTING
            print("Saving original Word settings...")
            try:
                # Initialize the dictionary first
                self.original_word_settings = {}
                
                # Save each setting individually with error handling
                try:
                    self.original_word_settings['CheckGrammarAsYouType'] = self.word_app.Options.CheckGrammarAsYouType
                    print(f"  Saved CheckGrammarAsYouType = {self.original_word_settings['CheckGrammarAsYouType']}")
                except Exception as e:
                    print(f"  Could not save CheckGrammarAsYouType: {e}")
                
                try:
                    self.original_word_settings['CheckSpellingAsYouType'] = self.word_app.Options.CheckSpellingAsYouType
                    print(f"  Saved CheckSpellingAsYouType = {self.original_word_settings['CheckSpellingAsYouType']}")
                except Exception as e:
                    print(f"  Could not save CheckSpellingAsYouType: {e}")
                
                try:
                    self.original_word_settings['SuggestSpellingCorrections'] = self.word_app.Options.SuggestSpellingCorrections
                    print(f"  Saved SuggestSpellingCorrections = {self.original_word_settings['SuggestSpellingCorrections']}")
                except Exception as e:
                    print(f"  Could not save SuggestSpellingCorrections: {e}")
                
                try:
                    self.original_word_settings['AutoFormatAsYouTypeApplyHeadings'] = self.word_app.Options.AutoFormatAsYouTypeApplyHeadings
                    print(f"  Saved AutoFormatAsYouTypeApplyHeadings = {self.original_word_settings['AutoFormatAsYouTypeApplyHeadings']}")
                except Exception as e:
                    print(f"  Could not save AutoFormatAsYouTypeApplyHeadings: {e}")
                
                try:
                    self.original_word_settings['AutoFormatAsYouTypeApplyBulletedLists'] = self.word_app.Options.AutoFormatAsYouTypeApplyBulletedLists
                    print(f"  Saved AutoFormatAsYouTypeApplyBulletedLists = {self.original_word_settings['AutoFormatAsYouTypeApplyBulletedLists']}")
                except Exception as e:
                    print(f"  Could not save AutoFormatAsYouTypeApplyBulletedLists: {e}")
                
                try:
                    self.original_word_settings['AutoFormatAsYouTypeApplyNumberedLists'] = self.word_app.Options.AutoFormatAsYouTypeApplyNumberedLists
                    print(f"  Saved AutoFormatAsYouTypeApplyNumberedLists = {self.original_word_settings['AutoFormatAsYouTypeApplyNumberedLists']}")
                except Exception as e:
                    print(f"  Could not save AutoFormatAsYouTypeApplyNumberedLists: {e}")
                
                try:
                    self.original_word_settings['ScreenUpdating'] = self.word_app.ScreenUpdating
                    print(f"  Saved ScreenUpdating = {self.original_word_settings['ScreenUpdating']}")
                except Exception as e:
                    print(f"  Could not save ScreenUpdating: {e}")
                
                try:
                    self.original_word_settings['EnableEvents'] = self.word_app.EnableEvents
                    print(f"  Saved EnableEvents = {self.original_word_settings['EnableEvents']}")
                except Exception as e:
                    print(f"  Could not save EnableEvents: {e}")
                
                print(f"✓ Successfully saved {len(self.original_word_settings)} Word settings")
                
            except Exception as e:
                print(f"Warning: Could not save original settings: {e}")
                # Initialize empty dict so restore doesn't crash
                self.original_word_settings = {}
            
            # PERFORMANCE SETTINGS FOR PROCESSING (TEMPORARY)
            print("Applying temporary performance optimizations...")
            try:
                self.word_app.ScreenUpdating = False  # Stop screen refreshing
                self.word_app.EnableEvents = False    # Disable Word events
                print("✓ Screen updating and events disabled")
            except Exception as e:
                print(f"Could not disable screen updating/events: {e}")
            
            try:
                # Disable spell/grammar checking that slows down corporate environments
                self.word_app.Options.CheckGrammarAsYouType = False
                self.word_app.Options.CheckSpellingAsYouType = False
                self.word_app.Options.SuggestSpellingCorrections = False
                print("✓ Spell/grammar checking temporarily disabled")
            except Exception as e:
                print(f"Could not disable spell/grammar checking: {e}")
            
            try:
                # Additional corporate environment optimizations
                self.word_app.Options.AutoFormatAsYouTypeApplyHeadings = False
                self.word_app.Options.AutoFormatAsYouTypeApplyBulletedLists = False
                self.word_app.Options.AutoFormatAsYouTypeApplyNumberedLists = False
                print("✓ Auto-formatting temporarily disabled")
            except Exception as e:
                print(f"Could not disable auto-formatting: {e}")
            
            # Test that Word is working
            doc_count = self.word_app.Documents.Count
            print(f"Word initialized successfully with performance optimizations. Current documents: {doc_count}")
            
            return True
            
        except Exception as e:
            print(f"Error initializing Word: {e}")
            self.word_app = None
            raise Exception(f"Could not initialize Microsoft Word: {str(e)}\n\nPlease ensure:\n1. Word is installed\n2. Word is not currently busy\n3. You have proper permissions\n4. Try closing all Word instances and restarting")

    def restore_word_settings(self):
        """Restore original Word settings"""
        if not self.word_app:
            print("No Word application to restore settings to")
            return
        
        if not hasattr(self, 'original_word_settings') or not self.original_word_settings:
            print("No Word settings were saved to restore")
            return
        
        try:
            print("Restoring original Word settings...")
            
            # Restore all saved settings
            for setting_name, original_value in self.original_word_settings.items():
                try:
                    if setting_name in ['ScreenUpdating', 'EnableEvents']:
                        # These are direct properties of the application
                        setattr(self.word_app, setting_name, original_value)
                        print(f"  ✓ Restored {setting_name} = {original_value}")
                    else:
                        # These are Options properties
                        setattr(self.word_app.Options, setting_name, original_value)
                        print(f"  ✓ Restored {setting_name} = {original_value}")
                except Exception as e:
                    print(f"  ✗ Could not restore {setting_name}: {e}")
            
            print("✓ Word settings restoration completed")
            
            # Force Word to save these settings by creating and closing a temporary document
            try:
                print("Forcing Word to save restored settings...")
                temp_doc = self.word_app.Documents.Add()
                temp_doc.Close(SaveChanges=False)
                print("✓ Settings save forced")
            except Exception as e:
                print(f"Could not force settings save: {e}")
            
        except Exception as e:
            print(f"Error restoring Word settings: {e}")

    def select_word_document(self):
        """Select the Word document to process"""
        # Initialize Word if not already done
        if not self.initialize_word():
            return None
            
        file_path = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word documents", "*.docx *.doc"), ("All files", "*.*")]
        )
        
        if not file_path:
            return None
            
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"File does not exist: {file_path}")
            return None
            
        try:
            print(f"Opening original document: {file_path}")
            
            # Convert to absolute path and normalize
            abs_path = os.path.abspath(file_path)
            print(f"Absolute path: {abs_path}")
            
            # Close any existing documents we have open
            if self.doc:
                try:
                    self.doc.Close(SaveChanges=False)
                except:
                    pass
            if self.original_doc:
                try:
                    self.original_doc.Close(SaveChanges=False)
                except:
                    pass
            
            # Open the original document as READ-ONLY for reference
            print("Opening original document as read-only reference...")
            self.original_doc = self.word_app.Documents.Open(abs_path, ReadOnly=True)
            
            # Create a working copy to preserve the original
            print("Creating working copy...")
            
            # Generate working copy name using same convention as PDF output
            original_dir = os.path.dirname(abs_path)
            original_name = os.path.basename(abs_path)
            name_without_ext = os.path.splitext(original_name)[0]
            ext = os.path.splitext(original_name)[1]
            
            # Create working copy name (same pattern as PDF default)
            mode_suffix = "_with_bates_links" if self.bates_mode else "_with_links"
            working_copy_name = f"{name_without_ext}{mode_suffix}{ext}"
            working_copy_path = os.path.join(original_dir, working_copy_name)
            
            # Handle existing files by adding counter
            counter = 1
            while os.path.exists(working_copy_path):
                working_copy_name = f"{name_without_ext}{mode_suffix}_{counter}{ext}"
                working_copy_path = os.path.join(original_dir, working_copy_name)
                counter += 1
            
            print(f"Working copy path: {working_copy_path}")
            
            # Save original as the working copy
            self.original_doc.SaveAs2(working_copy_path)
            
            # Close the original (we don't need it open anymore)
            self.original_doc.Close(SaveChanges=False)
            
            # Now open the working copy as our working document
            print("Opening working copy...")
            self.doc = self.word_app.Documents.Open(working_copy_path, ReadOnly=False)
            
            # Reopen original as read-only reference (if needed for comparison)
            self.original_doc = self.word_app.Documents.Open(abs_path, ReadOnly=True)
            
            # Store paths - CRITICAL FIX: Always point to original document's location
            self.doc_folder = original_dir  # Use original document's directory, not working copy
            self.original_doc_path = abs_path
            self.working_copy_path = working_copy_path  # Store working copy path
            self.target_folder = original_dir  # Default to original document's folder, not working copy
            
            print(f"Working copy created successfully. Paragraphs: {self.doc.Paragraphs.Count}")
            print(f"Original document remains untouched at: {abs_path}")
            print(f"Working copy saved at: {working_copy_path}")
            print(f"Document folder set to: {self.doc_folder}")
            print(f"Target folder set to: {self.target_folder}")
            
            # Build Bates PDF map if in Bates mode
            if self.bates_mode:
                self.build_bates_pdf_map()
            
            return file_path
            
        except Exception as e:
            print(f"Error opening document: {e}")
            messagebox.showerror("Error", f"Could not open document: {str(e)}")
            return None    

    def select_exhibit_folder(self):
        """Select the folder containing exhibit files"""
        folder_path = filedialog.askdirectory(
            title="Select Exhibit Folder" if not self.bates_mode else "Select Bates PDF Folder",
            initialdir=self.doc_folder if self.doc_folder else "."
        )
        if folder_path:
            self.target_folder = folder_path
            # Rebuild Bates map if in Bates mode
            if self.bates_mode:
                self.build_bates_pdf_map()
            return folder_path
        return None

    def find_matching_files(self, reference_text):
        """Find files that match the reference (Exhibit or Bates mode)"""
        if not self.target_folder:
            return []
        
        if self.bates_mode:
            return self.find_matching_bates_files(reference_text)
        else:
            return self.find_matching_exhibit_files(reference_text)

    def find_matching_exhibit_files(self, reference_text):
        """Find files in the target folder that match the exhibit reference - ENHANCED VERSION"""
        matching_files = []
        try:
            files_in_folder = os.listdir(self.target_folder)
        except Exception as e:
            print(f"Error reading folder {self.target_folder}: {e}")
            return []
        
        for pattern in self.exhibit_patterns:
            match = re.search(pattern, reference_text, re.IGNORECASE)
            if match:
                identifier = match.group(1)
                
                print(f"REFERENCE: '{reference_text}' -> EXTRACTED: '{identifier}'")
                
                # ENHANCED: Try multiple filename patterns for each identifier
                possible_prefixes = [
                    f"Ex. {identifier}",     # Ex. 1, Ex. A
                    f"Ex.{identifier}",      # Ex.1, Ex.A
                    f"Ex {identifier}",      # Ex 1, Ex A
                    f"Ex_{identifier}",      # Ex_1, Ex_A
                    f"Exhibit {identifier}", # Exhibit 1, Exhibit A
                    f"Exhibit_{identifier}", # Exhibit_1, Exhibit_A
                ]
                
                for target_prefix in possible_prefixes:
                    print(f"  Trying prefix: '{target_prefix}'")
                    
                    for filename in files_in_folder:
                        # Check if filename starts with our target prefix
                        if filename.startswith(target_prefix):
                            prefix_len = len(target_prefix)
                            
                            if prefix_len >= len(filename):
                                # Exact match - filename is exactly our target
                                full_path = os.path.join(self.target_folder, filename)
                                matching_files.append(full_path)
                                print(f"    ✓ EXACT MATCH: '{reference_text}' -> '{filename}'")
                            else:
                                # Check what comes after our prefix
                                next_char = filename[prefix_len]
                                # Allow common separators and extensions
                                if next_char in ['_', '-', '.', ' ']:
                                    full_path = os.path.join(self.target_folder, filename)
                                    matching_files.append(full_path)
                                    print(f"    ✓ PARTIAL MATCH: '{reference_text}' -> '{filename}'")
                    
                    # If we found matches with this prefix pattern, stop trying other patterns
                    if matching_files:
                        print(f"  Found {len(matching_files)} matches with prefix '{target_prefix}'")
                        break
                
                # If we found matches with this regex pattern, stop trying other patterns
                if matching_files:
                    break
        
        if not matching_files:
            print(f"✗ NO MATCH: '{reference_text}'")
        
        return matching_files    

    def find_matching_bates_files(self, reference_text):
        """Find Bates PDF and page number for the reference"""
        matching_files = []
        
        bates_patterns = self.get_bates_patterns()
        for pattern in bates_patterns:
            match = re.search(pattern, reference_text, re.IGNORECASE)
            if match:
                bates_number = int(match.group(1))
                print(f"BATES REFERENCE: '{reference_text}' -> EXTRACTED: {bates_number}")
                
                pdf_path, page_number = self.find_bates_pdf_and_page(bates_number)
                if pdf_path and page_number:
                    # Create a special entry that includes page information
                    matching_files.append({
                        'type': 'bates',
                        'path': pdf_path,
                        'page': page_number,
                        'bates_number': bates_number
                    })
                    print(f"✓ BATES MATCHED: '{reference_text}' -> {os.path.basename(pdf_path)} page {page_number}")
                else:
                    print(f"✗ NO BATES MATCH: '{reference_text}' -> Bates {bates_number}")
                
                break  # Stop after first match
        
        return matching_files

    def set_word_hyperlink_base_for_relative_links(self):
        """Set Word document properties to force relative hyperlinks"""
        try:
            print("Setting Word document to use relative hyperlinks...")
            
            # Set Hyperlink Base to the original document's directory
            if hasattr(self, 'original_doc_path') and self.original_doc_path:
                base_path = os.path.dirname(self.original_doc_path)
            else:
                base_path = self.doc_folder
            
            print(f"Setting Hyperlink Base to: {base_path}")
            
            # Access built-in document properties and set Hyperlink Base
            builtin_props = self.doc.BuiltInDocumentProperties
            hyperlink_base_prop = builtin_props("Hyperlink base")
            hyperlink_base_prop.Value = base_path
            
            print("✓ Hyperlink Base set successfully")
            return True
            
        except Exception as e:
            print(f"Error setting Hyperlink Base: {e}")
            return False

    def get_relative_path_from_original_doc(self, file_path):
        """Calculate relative path from ORIGINAL document location for consistent linking"""
        # Use the original document's directory as the reference point
        if hasattr(self, 'original_doc_path') and self.original_doc_path:
            reference_dir = os.path.dirname(self.original_doc_path)
            print(f"Using original doc directory as reference: {reference_dir}")
        else:
            reference_dir = self.doc_folder
            print(f"Using doc_folder as reference: {reference_dir}")
        
        if not reference_dir:
            return os.path.basename(file_path)  # Just filename as fallback
        
        try:
            relative_path = os.path.relpath(file_path, reference_dir)
            normalized_path = relative_path.replace('\\', '/')
            print(f"Calculated relative path: {file_path} -> {normalized_path}")
            return normalized_path
        except ValueError:
            # Different drives - use just filename
            print(f"Different drives detected, using filename: {os.path.basename(file_path)}")
            return os.path.basename(file_path)

    def get_relative_path(self, file_path):
        """Convert absolute path to relative path from document location"""
        if not self.doc_folder:
            return file_path
        
        try:
            relative_path = os.path.relpath(file_path, self.doc_folder)
            normalized_path = relative_path.replace('\\', '/')
            return normalized_path
        except ValueError:
            return file_path.replace('\\', '/')
    
    def safe_range_operation(self, operation, *args, **kwargs):
        """Safely perform a range operation with error handling"""
        max_retries = 3
        for attempt in range(max_retries):
            try:
                return operation(*args, **kwargs)
            except Exception as e:
                print(f"Range operation failed (attempt {attempt + 1}): {e}")
                if attempt == max_retries - 1:
                    raise
                time.sleep(0.1)  # Brief pause before retry

    def process_range_for_hyperlinks(self, range_obj, range_name=""):
        """Process a Word range (paragraph, footnote, etc.) for hyperlinks - FIXED VERSION"""
        if not range_obj:
            return 0
        
        try:
            range_text = range_obj.Text
        except Exception as e:
            print(f"Error reading range text for {range_name}: {e}")
            return 0
        
        # Check for relevant patterns based on mode
        if self.bates_mode:
            if not self.bates_prefix or self.bates_prefix not in range_text:
                return 0
            patterns = self.get_bates_patterns()
        else:
            if not range_text or not ('Ex.' in range_text or 'Exhibit' in range_text):
                return 0
            patterns = self.exhibit_patterns
        
        print(f"\nProcessing {range_name}: '{range_text[:100]}...'")
        
        # Find all references in this range
        references = []
        matched_positions = set()
        
        for pattern in patterns:
            for match in re.finditer(pattern, range_text, re.IGNORECASE):
                start_pos = match.start()
                end_pos = match.end()
                reference = match.group(0)
                
                # Skip if this position overlaps with a previously matched reference
                if any(start <= start_pos < end for start, end in matched_positions):
                    continue
                
                matching_files = self.find_matching_files(reference)
                if matching_files:
                    references.append({
                        'reference': reference,
                        'start_pos': start_pos,
                        'end_pos': end_pos,
                        'file_info': matching_files[0]
                    })
                    matched_positions.add((start_pos, end_pos))
                    print(f"  Found reference: '{reference}' at positions {start_pos}-{end_pos}")
        
        if not references:
            return 0
        
        # Sort by position (reverse order to avoid position shifts)
        references.sort(key=lambda x: x['start_pos'], reverse=True)
        
        links_added = 0
        
        # CRITICAL: Process each reference with improved range handling
        for ref in references:
            try:
                # Re-read the range text to account for any changes from previous hyperlinks
                current_range_text = range_obj.Text
                
                # Verify the text still matches at the expected position
                expected_text = ref['reference']
                actual_text_at_pos = current_range_text[ref['start_pos']:ref['end_pos']]
                
                print(f"  Expected: '{expected_text}' vs Actual: '{actual_text_at_pos}'")
                
                # If the text doesn't match exactly, try to find it nearby
                if actual_text_at_pos != expected_text:
                    print(f"  Position mismatch detected, searching for correct position...")
                    
                    # Search for the exact text in a small window around the expected position
                    search_window_start = max(0, ref['start_pos'] - 5)
                    search_window_end = min(len(current_range_text), ref['end_pos'] + 5)
                    search_window = current_range_text[search_window_start:search_window_end]
                    
                    # Try to find the exact match within the window
                    local_match = re.search(re.escape(expected_text), search_window, re.IGNORECASE)
                    if local_match:
                        # Adjust positions based on the local match
                        corrected_start = search_window_start + local_match.start()
                        corrected_end = search_window_start + local_match.end()
                        print(f"  Corrected position: {corrected_start}-{corrected_end}")
                        ref['start_pos'] = corrected_start
                        ref['end_pos'] = corrected_end
                    else:
                        print(f"  Could not find exact match, skipping this reference")
                        continue
                
                try:
                    print("  Trying precise offset method...")
                    ref_range = range_obj.Duplicate
                    ref_range.Start = range_obj.Start + ref['start_pos']
                    ref_range.End = range_obj.Start + ref['end_pos']
                    
                    # A quick verification to ensure we got the right text. If this fails,
                    # it will trigger the fallback method.
                    if ref_range.Text.strip().lower() != expected_text.strip().lower():
                        raise ValueError("Text mismatch, likely due to a complex object.")
                    
                    print("  ✓ Precise offset method succeeded.")

                # If the precise method fails (usually due to an image/chart messing up
                # character counts), fall back to the slower but more robust Find.Execute method.
                except Exception as e:
                    print(f"  ⚠️ Offset method failed: {e}. Falling back to robust Find.Execute method...")
                    ref_range = range_obj.Duplicate
                    
                    # Use Word's built-in Find to locate the text, which is better at
                    # handling complex layouts with images and other objects.
                    find_success = ref_range.Find.Execute(
                        FindText=expected_text,
                        MatchCase=False,
                        MatchWholeWord=False,
                        MatchWildcards=False,
                        Forward=True,
                        Wrap=0  # wdFindStop
                    )

                    if not find_success:
                        print(f"  ✗ Fallback Find.Execute also failed for '{expected_text}'. Skipping.")
                        continue # Move to the next reference
                # --- HYBRID METHOD END ---
                
                # Double-check the range text before creating hyperlink
                final_range_text = ref_range.Text
                print(f"  Final range text: '{final_range_text}' (expected: '{expected_text}')")
                
                # Only proceed if we have the right text
                if final_range_text.strip().lower() != expected_text.strip().lower():
                    print(f"  Final text verification failed, skipping hyperlink creation")
                    continue
                
                # Handle different file info types
                file_info = ref['file_info']
                if isinstance(file_info, dict) and file_info.get('type') == 'bates':
                    # Bates mode - link to specific page
                    file_path = file_info['path']
                    page_number = file_info['page']
                    relative_path = self.get_relative_path_from_original_doc(file_path)
                    link_target = f"{relative_path}#page={page_number}"
                    screen_tip = f"Link to {os.path.basename(file_path)} page {page_number} (Bates {file_info['bates_number']})"

                else:
                    # Regular exhibit mode
                    exhibit_id = None
                    page_number = None
                    
                    # Only check for page automation if it's enabled
                    if self.page_automation_enabled:
                        context_start = ref['start_pos']  # Start exactly at the exhibit reference
                        context_end = min(len(range_text), ref['end_pos'] + 150)  # Look 150 chars after
                        context_text = range_text[context_start:context_end]

                        print(f"  Checking page automation for exhibit: '{expected_text}'")
                        print(f"  Using context: '{context_text}'")
                        
                        # Get the current exhibit ID for comparison
                        current_exhibit_id = None
                        for pattern in self.exhibit_patterns:
                            match = re.search(pattern, expected_text, re.IGNORECASE)
                            if match:
                                current_exhibit_id = match.group(1)
                                break
                        
                        # Get all possible matches
                        all_matches = self.find_page_number_in_text(context_text)
                        
                        # Find the match that corresponds to our current exhibit
                        if all_matches and current_exhibit_id:
                            for match_exhibit_id, match_page_number, match_obj in all_matches:
                                if match_exhibit_id.upper() == current_exhibit_id.upper():
                                    exhibit_id, page_number = match_exhibit_id, match_page_number
                                    print(f"  ✓ Found matching exhibit ID: '{exhibit_id}' -> Page {page_number}")
                                    break
                            
                            if not exhibit_id:
                                print(f"  ✗ No matching exhibit ID found for '{current_exhibit_id}' in context")
                    
                    print(f"  Page automation result: exhibit_id='{exhibit_id}', page_number={page_number}")
                    
                    # Set up paths and targets
                    file_path = file_info  # Set this first, used in all cases
                    relative_path = self.get_relative_path_from_original_doc(file_path)
                    
                    if exhibit_id and page_number:
                        # Page automation mode - link to specific page
                        link_target = f"{relative_path}#page={page_number}"
                        screen_tip = f"Link to {os.path.basename(file_path)} page {page_number}"
                        print(f"  Using page automation: '{expected_text}' -> {link_target}")
                    else:
                        # Regular exhibit mode - link to file
                        link_target = relative_path
                        screen_tip = f"Link to {os.path.basename(file_path)}"


                print(f"  Creating hyperlink: '{link_target}' for text '{final_range_text}'")

                try:
                    # ENHANCED SOLUTION: Better Word hyperlink handling for page fragments
                    if '#page=' in link_target:
                        # Split the URL into address and fragment
                        parts = link_target.split('#page=')
                        base_address = parts[0]  # File path without fragment
                        page_num = parts[1]      # Page number
                        
                        print(f"  Splitting URL: Address='{base_address}', Page='{page_num}'")
                        
                        try:
                            # Method 1: Use SubAddress parameter (Word's preferred method)
                            hyperlink = range_obj.Hyperlinks.Add(
                                Anchor=ref_range,
                                Address=base_address,
                                SubAddress=f"page={page_num}",
                                TextToDisplay=expected_text,
                                ScreenTip=screen_tip
                            )
                            print(f"  ✓ Created hyperlink with SubAddress: page={page_num}")
                            
                        except Exception as subaddress_error:
                            print(f"  SubAddress failed: {subaddress_error}, trying alternative...")
                            
                            try:
                                # Method 2: Create with original target and fix encoding immediately
                                hyperlink = range_obj.Hyperlinks.Add(
                                    Anchor=ref_range,
                                    Address=link_target,
                                    TextToDisplay=expected_text,
                                    ScreenTip=screen_tip
                                )
                                
                                # Check if Word encoded it and fix immediately without changing state
                                current_address = hyperlink.Address
                                if '%23page=' in current_address:
                                    # Delete and recreate to avoid "visited" state
                                    hyperlink.Delete()
                                    
                                    # Try creating with the file path only
                                    hyperlink = range_obj.Hyperlinks.Add(
                                        Anchor=ref_range,
                                        Address=base_address,
                                        TextToDisplay=expected_text,
                                        ScreenTip=f"Link to {os.path.basename(base_address)} page {page_num}"
                                    )
                                    print(f"  ✓ Created hyperlink without page fragment to preserve color")
                                else:
                                    print(f"  ✓ Created hyperlink with correct fragment")
                                
                            except Exception as alternative_error:
                                print(f"  Alternative method failed: {alternative_error}")
                                # Fallback - create without fragment
                                hyperlink = range_obj.Hyperlinks.Add(
                                    Anchor=ref_range,
                                    Address=base_address,
                                    TextToDisplay=expected_text,
                                    ScreenTip=screen_tip
                                )
                                print(f"  ✓ Created hyperlink without page fragment (fallback)")
                    
                    else:
                        # No page fragment - simple hyperlink
                        hyperlink = range_obj.Hyperlinks.Add(
                            Anchor=ref_range,
                            Address=link_target,
                            TextToDisplay=expected_text,
                            ScreenTip=screen_tip
                        )
                        print(f"  ✓ Created simple hyperlink (no page fragment)")

                    # Ensure hyperlink has proper unvisited appearance
                    try:
                        # Force hyperlink to appear as unvisited (bright blue)
                        hyperlink_range = hyperlink.Range
                        hyperlink_range.Font.Color = 16711680  # Bright blue (BGR format)
                        hyperlink_range.Font.Underline = True
                        print(f"    ✓ Applied fresh hyperlink formatting")
                    except Exception as format_error:
                        print(f"    Could not apply fresh formatting: {format_error}")

                except Exception as e:
                    print(f"  Error creating hyperlink: {e}")
                    continue
                
                print(f"  ✓ Added hyperlink for '{expected_text}'")
                links_added += 1
                
            except Exception as e:
                print(f"  ✗ Error adding hyperlink for '{ref['reference']}': {e}")
        
        return links_added

    def process_document(self, progress_callback=None):
        """Process the document for exhibit hyperlinks using COM with progress updates"""
        if not self.doc or not self.target_folder:
            return 0
        
        self.set_word_hyperlink_base_for_relative_links()

        mode_text = "BATES" if self.bates_mode else "EXHIBIT"
        print(f"=== PROCESSING DOCUMENT IN {mode_text} MODE ===")
        print(f"Target folder: {self.target_folder}")
        
        if self.bates_mode:
            print(f"Bates prefix: '{self.bates_prefix}'")
            print(f"Bates PDF map: {len(self.bates_pdf_map)} PDFs")
        
        try:
            para_count = self.doc.Paragraphs.Count
            print(f"Document has {para_count} paragraphs")
        except Exception as e:
            print(f"Error accessing document: {e}")
            return 0
        
        # List available files
        try:
            files_in_folder = os.listdir(self.target_folder)
            if self.bates_mode:
                relevant_files = [f for f in files_in_folder if f.startswith(self.bates_prefix) and f.endswith('.pdf')]
                print(f"Available Bates PDF files: {relevant_files}")
            else:
                relevant_files = [f for f in files_in_folder if f.startswith('Ex.')]
                print(f"Available exhibit files: {relevant_files}")
        except Exception as e:
            print(f"Error reading folder: {e}")
            return 0
        
        total_links_added = 0
        
        # Calculate total work for progress tracking
        try:
            footnote_count = self.doc.Footnotes.Count
            endnote_count = self.doc.Endnotes.Count
        except:
            footnote_count = 0
            endnote_count = 0
        
        total_items = para_count + footnote_count + endnote_count
        processed_items = 0
        
        # Helper function to update progress
        def update_progress(current, total, status_text="Processing..."):
            if progress_callback:
                progress_percent = int((current / total) * 100) if total > 0 else 0
                progress_callback(progress_percent, status_text)
        
        # Process main document paragraphs with progress updates
        print(f"\n--- Processing main document paragraphs in {mode_text} mode ---")
        update_progress(0, total_items, f"Processing {para_count} paragraphs...")
        
        for i in range(1, para_count + 1):
            try:
                paragraph = self.doc.Paragraphs(i)
                paragraph_range = paragraph.Range
                
                links_in_para = self.process_range_for_hyperlinks(
                    paragraph_range, 
                    f"Paragraph {i}"
                )
                total_links_added += links_in_para
                processed_items += 1
                
                # Update progress every 10 paragraphs or at milestones
                if i % 10 == 0 or i == para_count:
                    update_progress(
                        processed_items, 
                        total_items, 
                        f"Processed {i}/{para_count} paragraphs, {total_links_added} links found"
                    )
                
            except Exception as e:
                print(f"Error processing paragraph {i}: {e}")
                processed_items += 1
        
        # Process footnotes with progress updates
        print("\n--- Processing footnotes ---")
        update_progress(processed_items, total_items, f"Processing {footnote_count} footnotes...")
        
        try:
            footnotes = self.doc.Footnotes
            footnote_count = footnotes.Count
            print(f"Found {footnote_count} footnotes")
            
            for i in range(1, footnote_count + 1):
                try:
                    footnote = footnotes(i)
                    footnote_range = footnote.Range
                    
                    links_in_footnote = self.process_range_for_hyperlinks(
                        footnote_range,
                        f"Footnote {i}"
                    )
                    total_links_added += links_in_footnote
                    processed_items += 1
                    
                    # Update progress for footnotes
                    update_progress(
                        processed_items, 
                        total_items, 
                        f"Processed footnote {i}/{footnote_count}, {total_links_added} links total"
                    )
                    
                except Exception as e:
                    print(f"Error processing footnote {i}: {e}")
                    processed_items += 1
                    
        except Exception as e:
            print(f"Error accessing footnotes: {e}")
        
        # Process endnotes with progress updates
        print("\n--- Processing endnotes ---")
        update_progress(processed_items, total_items, f"Processing {endnote_count} endnotes...")
        
        try:
            endnotes = self.doc.Endnotes
            endnote_count = endnotes.Count
            print(f"Found {endnote_count} endnotes")
            
            for i in range(1, endnote_count + 1):
                try:
                    endnote = endnotes(i)
                    endnote_range = endnote.Range
                    
                    links_in_endnote = self.process_range_for_hyperlinks(
                        endnote_range,
                        f"Endnote {i}"
                    )
                    total_links_added += links_in_endnote
                    processed_items += 1
                    
                    # Update progress for endnotes
                    update_progress(
                        processed_items, 
                        total_items, 
                        f"Processed endnote {i}/{endnote_count}, {total_links_added} links total"
                    )
                    
                except Exception as e:
                    print(f"Error processing endnote {i}: {e}")
                    processed_items += 1
                    
        except Exception as e:
            print(f"Error accessing endnotes: {e}")
        
        # Final progress update
        update_progress(100, 100, f"Processing complete! {total_links_added} links created")
        
        print(f"\n=== PROCESSING COMPLETE ===")
        print(f"Total links added: {total_links_added}")
        
        # Save the working copy with hyperlinks
        if total_links_added > 0:
            try:
                update_progress(100, 100, "Saving document with hyperlinks...")
                print("Saving working copy with hyperlinks...")
                self.doc.Save()
                print("Working copy saved successfully with hyperlinks")
            except Exception as e:
                print(f"Could not save working copy: {e}")
        
        return total_links_added


    def export_to_pdf_with_relative_links(self, word_pdf_path):
        """Export Word document to PDF with OneDrive compatibility fixes - FIXED for spaces and parameters"""
        if not self.doc:
            return False
        
        try:
            print("\n=== WORD PDF EXPORT + MANUAL FIX ===")
            print(f"Exporting to: {word_pdf_path}")
            
            # Normalize the path and ensure directory exists
            normalized_path = os.path.normpath(word_pdf_path)
            target_dir = os.path.dirname(normalized_path)
            
            print(f"Normalized path: {normalized_path}")
            print(f"Target directory: {target_dir}")
            
            # Ensure target directory exists
            if not os.path.exists(target_dir):
                print(f"Creating directory: {target_dir}")
                os.makedirs(target_dir, exist_ok=True)
            
            # Method 1: Try direct export with minimal parameters (most compatible)
            try:
                print("Attempting direct PDF export with minimal parameters...")
                
                # Use only the most basic parameters that all Word versions support
                self.doc.ExportAsFixedFormat(
                    OutputFileName=normalized_path,
                    ExportFormat=17  # wdExportFormatPDF - just the basics
                )
                
                print("✅ Direct PDF export succeeded")
                actual_pdf_path = normalized_path
                
            except Exception as direct_error:
                print(f"Direct export failed: {direct_error}")
                print("Trying temporary directory method...")
                
                # Method 2: Export to temp directory first, then copy
                try:
                    import tempfile
                    import shutil
                    
                    # Create temp file in system temp directory
                    temp_dir = tempfile.gettempdir()
                    temp_filename = f"word_export_{int(time.time())}.pdf"
                    temp_path = os.path.join(temp_dir, temp_filename)
                    
                    print(f"Temporary export path: {temp_path}")
                    
                    # Export to temp location with minimal parameters
                    self.doc.ExportAsFixedFormat(
                        OutputFileName=temp_path,
                        ExportFormat=17  # wdExportFormatPDF
                    )
                    
                    print("✅ Temporary PDF export succeeded")
                    
                    # Verify temp file exists
                    if not os.path.exists(temp_path):
                        raise Exception("Temporary PDF file was not created")
                    
                    # Copy to final location
                    print(f"Copying from temp to final location...")
                    shutil.copy2(temp_path, normalized_path)
                    
                    # Verify final file exists
                    if not os.path.exists(normalized_path):
                        raise Exception("Final PDF file was not created")
                    
                    print("✅ PDF copied to final location")
                    actual_pdf_path = normalized_path
                    
                    # Clean up temp file
                    try:
                        os.remove(temp_path)
                        print("✅ Temporary file cleaned up")
                    except:
                        print("Warning: Could not clean up temporary file")
                    
                except Exception as temp_error:
                    print(f"Temporary directory method failed: {temp_error}")
                    
                    # Method 3: Try Print to PDF as last resort
                    try:
                        print("Trying Print to PDF method...")
                        
                        # Save current printer
                        original_printer = self.word_app.ActivePrinter
                        
                        # Set to Microsoft Print to PDF
                        self.word_app.ActivePrinter = "Microsoft Print to PDF"
                        
                        # Print to file
                        self.doc.PrintOut(
                            OutputFileName=normalized_path,
                            PrintToFile=True
                        )
                        
                        # Restore original printer
                        self.word_app.ActivePrinter = original_printer
                        
                        print("✅ Print to PDF method succeeded")
                        actual_pdf_path = normalized_path
                        
                    except Exception as print_error:
                        print(f"Print to PDF failed: {print_error}")
                        print("❌ All PDF export methods failed")
                        return False
            
            # CRITICAL FIX: Check for space-encoded filename and rename if needed
            try:
                # Check if Word created a file with URL-encoded spaces
                expected_filename = os.path.basename(normalized_path)
                directory = os.path.dirname(normalized_path)
                
                # Look for files in the directory that might be our PDF with encoded spaces
                if not os.path.exists(normalized_path):
                    print("Expected PDF not found, checking for URL-encoded version...")
                    
                    import urllib.parse
                    encoded_filename = urllib.parse.quote(expected_filename, safe='.')
                    encoded_path = os.path.join(directory, encoded_filename)
                    
                    print(f"Checking for encoded version: {encoded_path}")
                    
                    if os.path.exists(encoded_path):
                        print(f"Found URL-encoded PDF, renaming to correct filename...")
                        shutil.move(encoded_path, normalized_path)
                        print(f"✅ Renamed {encoded_filename} to {expected_filename}")
                        actual_pdf_path = normalized_path
                    else:
                        # Look for any PDF files created around this time
                        print("Searching for any recently created PDF files...")
                        import glob
                        pdf_pattern = os.path.join(directory, "*.pdf")
                        pdf_files = glob.glob(pdf_pattern)
                        
                        # Find the most recently created PDF
                        if pdf_files:
                            newest_pdf = max(pdf_files, key=os.path.getctime)
                            creation_time = os.path.getctime(newest_pdf)
                            current_time = time.time()
                            
                            # If it was created within the last 30 seconds, it's probably ours
                            if current_time - creation_time < 30:
                                print(f"Found recently created PDF: {newest_pdf}")
                                if newest_pdf != normalized_path:
                                    print(f"Renaming to correct filename...")
                                    shutil.move(newest_pdf, normalized_path)
                                    print(f"✅ Renamed to {expected_filename}")
                                actual_pdf_path = normalized_path
                            else:
                                print("No recent PDF files found")
                                return False
                        else:
                            print("No PDF files found in directory")
                            return False
                else:
                    print("✅ PDF found at expected location")
                    actual_pdf_path = normalized_path
            
            except Exception as rename_error:
                print(f"Error checking/renaming PDF: {rename_error}")
                # Continue anyway, maybe the file is there
                actual_pdf_path = normalized_path
            
            # Now try to fix the hyperlink encoding if pypdf is available
            try:
                print("\n=== ANALYZING WORD'S OUTPUT ===")
                
                from pypdf import PdfReader
                
                if not os.path.exists(actual_pdf_path):
                    print(f"PDF file not found at: {actual_pdf_path}")
                    return False
                
                reader = PdfReader(actual_pdf_path)
                print(f"PDF has {len(reader.pages)} pages")
                
                total_links = 0
                needs_fix = False
                
                for page_num, page in enumerate(reader.pages):
                    if "/Annots" in page:
                        annots = page["/Annots"]
                        
                        for annot in annots:
                            if "/A" in annot and "/URI" in annot["/A"]:
                                uri = str(annot["/A"]["/URI"])
                                print(f"  📎 Page {page_num + 1} link: {uri}")
                                total_links += 1
                                
                                if "%23page=" in uri:
                                    print("    ⚠️  Contains %23page= (needs fix)")
                                    needs_fix = True
                                elif "#page=" in uri:
                                    print("    ✅ Contains #page= (already good)")
                
                print(f"\n📊 Found {total_links} links")
                
                if needs_fix:
                    print("\n🔧 APPLYING MANUAL FIX...")
                    fix_success = self.fix_word_pdf_encoding(actual_pdf_path)
                    
                    if fix_success:
                        print("✅ Manual fix applied successfully!")
                        print("🎉 Links should now work in both Chrome and Adobe!")
                    else:
                        print("❌ Manual fix failed")
                else:
                    print("✅ No fix needed - links already in correct format")
                
            except ImportError:
                print("pypdf not available - cannot analyze or fix links")
                print("PDF created successfully but links may need manual testing")
            except Exception as e:
                print(f"Link analysis error: {e}")
                print("PDF created successfully but could not analyze links")
            
            return True
            
        except Exception as e:
            print(f"PDF export completely failed: {e}")
            print("This might be due to OneDrive sync issues or permissions")
            return False

    def fix_word_pdf_encoding(self, pdf_path):
        """Fix %23page= encoding AND convert absolute paths back to relative"""
        try:
            print(f"Fixing encoding and converting to relative paths in: {pdf_path}")
            
            # Read PDF as binary
            with open(pdf_path, 'rb') as f:
                pdf_bytes = f.read()
            
            # Convert to string for replacement
            pdf_text = pdf_bytes.decode('latin-1', errors='replace')
            
            # Step 1: Fix %23page= encoding
            before_encoding_count = pdf_text.count('%23page=')
            print(f"Found {before_encoding_count} instances of '%23page=' to fix")
            
            fixed_text = pdf_text.replace('%23page=', '#page=')
            
            # Step 2: Convert absolute file:// paths back to relative paths
            print("Converting absolute paths to relative paths...")
            
            # Get the directory where the PDF is located
            pdf_dir = os.path.dirname(os.path.abspath(pdf_path))
            print(f"PDF directory: {pdf_dir}")
            
            # Pattern to match file:// URLs
            import re
            
            # Find all file:// URLs in the PDF
            # Look for file:/// followed by path, optionally followed by #page=number
            file_url_pattern = r'file:///([^\s\)>#]+)(#page=\d+)?'
            
            def convert_to_relative(match):
                full_path = match.group(1)  # The path part after file:///
                page_fragment = match.group(2) if match.group(2) else ""  # The #page=X part
                
                print(f"  Raw path captured: '{full_path}'")
                print(f"  Page fragment: '{page_fragment}'")
                
                if not full_path:
                    print(f"  ERROR: Empty path captured")
                    return match.group(0)  # Return original if empty
                
                # Convert back to Windows path format
                windows_path = full_path.replace('/', '\\')
                print(f"  Windows path: '{windows_path}'")
                
                try:
                    # Calculate relative path from PDF location
                    relative_path = os.path.relpath(windows_path, pdf_dir)
                    # Convert back to forward slashes for consistency
                    relative_path = relative_path.replace('\\', '/')
                    
                    print(f"  Converting: file:///{full_path}{page_fragment}")
                    print(f"         To: {relative_path}{page_fragment}")
                    
                    return relative_path + page_fragment
                    
                except Exception as e:
                    print(f"  Could not convert {full_path}: {e}")
                    # Return original if conversion fails
                    return f"file:///{full_path}{page_fragment}"
            
            # Apply the conversion
            fixed_text = re.sub(file_url_pattern, convert_to_relative, fixed_text)
            
            # Count changes made
            after_encoding_count = fixed_text.count('%23page=')
            encoding_fixes = before_encoding_count - after_encoding_count
            
            # Check for remaining file:// URLs
            remaining_file_urls = len(re.findall(r'file:///', fixed_text))
            
            print(f"Encoding fixes made: {encoding_fixes}")
            print(f"Remaining absolute file:// URLs: {remaining_file_urls}")
            
            if encoding_fixes > 0 or remaining_file_urls == 0:
                # Create temporary backup
                backup_path = pdf_path + '.backup'
                backup_created = False
                try:
                    with open(backup_path, 'wb') as f:
                        f.write(pdf_bytes)
                    backup_created = True
                    print(f"Temporary backup created")
                except:
                    print("Could not create backup (continuing anyway)")
                
                # Write fixed version
                fixed_bytes = fixed_text.encode('latin-1', errors='replace')
                with open(pdf_path, 'wb') as f:
                    f.write(fixed_bytes)
                
                # Verify fix worked
                with open(pdf_path, 'rb') as f:
                    verify_bytes = f.read()
                verify_text = verify_bytes.decode('latin-1', errors='replace')
                
                final_encoding_count = verify_text.count('%23page=')
                final_file_urls = len(re.findall(r'file:///', verify_text))
                
                # Clean up backup file
                if backup_created:
                    try:
                        os.remove(backup_path)
                        print("✅ Temporary backup cleaned up")
                    except:
                        print("Could not remove backup file")
                
                print(f"\n✅ FINAL RESULTS:")
                print(f"  %23page= instances: {final_encoding_count} (should be 0)")
                print(f"  Absolute file:// URLs: {final_file_urls} (should be 0)")
                
                if final_encoding_count == 0 and final_file_urls == 0:
                    print("🎉 Perfect! All links are now relative with correct encoding!")
                    return True
                elif final_encoding_count == 0:
                    print("✅ Encoding fixed, but some absolute URLs remain")
                    return True
                else:
                    print("⚠️  Some issues remain")
                    return False
            else:
                print("No changes needed")
                return True
                
        except Exception as e:
            print(f"Fix failed: {e}")
            import traceback
            traceback.print_exc()
            return False

    # Alternative approach: Try to prevent Word from creating absolute paths in the first place
    def create_relative_hyperlinks_in_word(self, range_obj, file_info, ref_text):
        """Create hyperlinks in Word using more relative-friendly format"""
        
        if isinstance(file_info, dict) and file_info.get('type') == 'bates':
            # Bates mode
            target_file = file_info['path']
            page_number = file_info['page_number']
            
            # Try using just the filename + page, not full path
            filename = os.path.basename(target_file)
            relative_path = f"{filename}#page={page_number}"
            
        else:
            # Exhibit mode
            target_file = file_info
            filename = os.path.basename(target_file)
            relative_path = filename
        
        print(f"Creating Word hyperlink with relative path: {relative_path}")
        
        try:
            hyperlink = range_obj.Hyperlinks.Add(
                Anchor=range_obj,
                Address=relative_path,  # Use just filename, not full path
                TextToDisplay=ref_text,
                ScreenTip=f"Link to {filename}"
            )
            return True
        except Exception as e:
            print(f"Failed to create relative hyperlink: {e}")
            # Fallback to original method
            return False
            
    def simple_pdf_export(self, word_pdf_path):
        """Simple Word export - works in Chrome, may not work in Adobe"""
        try:
            self.doc.ExportAsFixedFormat(word_pdf_path, 17)
            print("✅ Simple Word export completed")
            print("ℹ️  Links work in Chrome, may not work in Adobe (due to %23 encoding)")
            return True
        except Exception as e:
            print(f"Export failed: {e}")
            return False

    def save_document(self, output_path=None, progress_dialog=None):
        """Enhanced save with progress dialog management"""
        if not self.doc:
            return False
        
        if not output_path:
            # Generate default names but let user choose
            if self.original_doc_path:
                original_dir = os.path.dirname(self.original_doc_path)
                original_name = os.path.basename(self.original_doc_path)
                name_without_ext = os.path.splitext(original_name)[0]
                mode_suffix = "_with_bates_links" if self.bates_mode else "_with_links"
                default_word_name = f"{name_without_ext}{mode_suffix}.docx"
                default_pdf_name = f"{name_without_ext}{mode_suffix}.pdf"
                print(f"Save dialog: Using original document directory: {original_dir}")
                print(f"Save dialog: Default Word filename: {default_word_name}")
            else:
                print("WARNING: original_doc_path is not available - using current directory")
                original_dir = os.getcwd()
                default_word_name = "processed_document.docx"
                default_pdf_name = "processed_document.pdf"
            
            # HIDE PROGRESS DIALOG BEFORE SHOWING SAVE DIALOGS
            if progress_dialog:
                progress_dialog.hide_temporarily()
            
            try:
                # Ask user where to save Word document
                from tkinter import filedialog
                word_output = filedialog.asksaveasfilename(
                    title="Save Word Document with Links",
                    defaultextension=".docx",
                    filetypes=[("Word documents", "*.docx *.doc"), ("All files", "*.*")],
                    initialdir=original_dir,
                    initialfile=default_word_name
                )
                
                if not word_output:
                    print("User cancelled Word save")
                    # RESTORE PROGRESS DIALOG BEFORE RETURNING
                    if progress_dialog:
                        progress_dialog.show_again()
                    return False
                
                # Ask user where to save PDF
                word_dir = os.path.dirname(word_output)
                pdf_output = filedialog.asksaveasfilename(
                    title="Save PDF Export",
                    defaultextension=".pdf",
                    filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
                    initialdir=word_dir,
                    initialfile=default_pdf_name
                )
                
                if not pdf_output:
                    print("User cancelled PDF save")
                    # RESTORE PROGRESS DIALOG BEFORE RETURNING
                    if progress_dialog:
                        progress_dialog.show_again()
                    return False
                
                print(f"User selected Word path: {word_output}")
                print(f"User selected PDF path: {pdf_output}")
                
            finally:
                # ALWAYS RESTORE PROGRESS DIALOG AFTER SAVE DIALOGS
                if progress_dialog:
                    progress_dialog.show_again()
                    progress_dialog.update_progress(100, "Saving files...")
            
        else:
            word_output = output_path
            pdf_output = output_path.replace('.docx', '.pdf').replace('.doc', '.pdf')
        
        try:
            # Update progress if we have it
            if progress_dialog:
                progress_dialog.update_progress(100, "Saving Word document...")
            
            # Save Word document
            self.doc.SaveAs2(word_output)
            print("Word document saved successfully")
            
            # CRITICAL FIX: Clear hyperlink base AFTER saving to ensure relative links
            try:
                print("Clearing hyperlink base to ensure relative links...")
                builtin_props = self.doc.BuiltInDocumentProperties
                hyperlink_base_prop = builtin_props("Hyperlink base")
                hyperlink_base_prop.Value = ""  # Clear the hyperlink base
                
                # Save again to persist the cleared hyperlink base
                self.doc.Save()
                print("✓ Hyperlink base cleared and document re-saved")
            except Exception as e:
                print(f"Warning: Could not clear hyperlink base: {e}")
            
            word_saved = True
            
            # Update progress for PDF export
            if progress_dialog:
                progress_dialog.update_progress(100, "Exporting to PDF...")
            
            # Export to PDF with enhanced error handling
            print(f"Exporting to PDF: {pdf_output}")
            
            pdf_saved = self.export_to_pdf_with_relative_links(pdf_output)
            
            if word_saved and pdf_saved:
                print("Both Word and PDF files saved successfully!")
                if progress_dialog:
                    progress_dialog.update_progress(100, "Save complete!")
                return True
            elif word_saved:
                print("Word saved successfully, PDF export failed")
                if progress_dialog:
                    progress_dialog.update_progress(100, "Word saved, PDF export failed")
                
                from tkinter import messagebox
                messagebox.showwarning("Partial Success", 
                    f"Word document saved successfully!\n"
                    f"PDF export failed (possibly due to OneDrive sync issues).\n"
                    f"The Word file has working hyperlinks.\n\n"
                    f"Word file: {word_output}\n\n"
                    f"To create PDF manually:\n"
                    f"1. Open the Word file\n"
                    f"2. Go to File > Export > Create PDF/XPS\n"
                    f"3. Save as PDF")
                return True
            else:
                print("Both saves failed")
                if progress_dialog:
                    progress_dialog.update_progress(100, "Save failed")
                return False
                
        except Exception as e:
            print(f"Error saving documents: {e}")
            if progress_dialog:
                progress_dialog.update_progress(100, f"Error: {str(e)}")
            
            from tkinter import messagebox
            messagebox.showerror("Error", f"Could not save documents: {str(e)}")
            return False


    def cleanup(self):
        """Clean up COM objects and ensure all documents are properly closed"""
        try:
            print("Starting cleanup...")

            self.restore_word_settings()
            
            # Store working copy path before closing (CRITICAL FIX)
            working_copy_to_delete = None
            if hasattr(self, 'working_copy_path') and self.working_copy_path:
                working_copy_to_delete = self.working_copy_path
                print(f"Will delete working copy: {working_copy_to_delete}")
            
            # Close our specific documents first
            if self.doc:
                try:
                    print(f"Closing working document: {self.doc.Name}")
                    self.doc.Close(SaveChanges=False)
                    print("Working document closed successfully")
                except Exception as e:
                    print(f"Error closing working document: {e}")
                finally:
                    self.doc = None
                    
            if self.original_doc:
                try:
                    print(f"Closing original document: {self.original_doc.Name}")
                    self.original_doc.Close(SaveChanges=False)
                    print("Original document closed successfully")
                except Exception as e:
                    print(f"Error closing original document: {e}")
                finally:
                    self.original_doc = None
            
            # Force close any remaining documents that might be hanging around
            if self.word_app:
                try:
                    # Get count of open documents
                    doc_count = self.word_app.Documents.Count
                    print(f"Word has {doc_count} documents still open")
                    
                    # Close all documents (be more aggressive)
                    while self.word_app.Documents.Count > 0:
                        try:
                            doc = self.word_app.Documents(1)  # Get first document
                            doc_name = doc.Name
                            print(f"Force closing document: {doc_name}")
                            doc.Close(SaveChanges=False)
                        except Exception as e:
                            print(f"Error force closing document: {e}")
                            break  # Avoid infinite loop
                    
                    # Now quit Word application
                    print("Quitting Word application...")
                    self.word_app.Quit(SaveChanges=False)
                    print("Word application quit successfully")
                    
                except Exception as e:
                    print(f"Error during Word cleanup: {e}")
                finally:
                    self.word_app = None
            
            # Force COM cleanup
            import gc
            gc.collect()
            
            print("Cleanup completed")
            
            # Always try to uninitialize COM
            try:
                pythoncom.CoUninitialize()
                print("COM uninitialized")
            except Exception as e:
                print(f"Error uninitializing COM: {e}")
            
            # CRITICAL FIX: Delete the working copy file after Word is closed (like Excel does)
            if working_copy_to_delete and os.path.exists(working_copy_to_delete):
                try:
                    print(f"Deleting working copy file: {working_copy_to_delete}")
                    
                    # Wait a moment for Word to fully release the file
                    import time
                    time.sleep(1)
                    
                    # Try to delete the file
                    os.remove(working_copy_to_delete)
                    print("✓ Working copy file deleted successfully")
                    
                except Exception as e:
                    print(f"✗ Could not delete working copy file: {e}")
                    print("You may need to delete it manually")
            
        except Exception as e:
            print(f"Error during cleanup: {e}")
            
        # Note: Original document is preserved, only working copy is cleaned up
        if hasattr(self, 'original_doc_path') and self.original_doc_path:
            print(f"Original document preserved at: {self.original_doc_path}")

class FileRenamer:
    """Utility class to rename files for better Chrome PDF compatibility"""
    
    @staticmethod
    def normalize_filename(filename):
        """
        Convert filenames to Chrome-friendly format:
        - Ex. A Letter.pdf -> Ex._A_Letter.pdf
        - Ex. 55 Email.docx -> Ex._55_Email.docx
        - Exhibit 12 Memo.pdf -> Exhibit_12_Memo.pdf
        """
        # Split filename and extension
        name, ext = os.path.splitext(filename)
        
        # Skip files that don't look like exhibits
        if not (name.lower().startswith(('ex.', 'ex ', 'exhibit')) or 
                re.match(r'^ex[._\s]', name.lower())):
            return filename
        
        print(f"Processing: '{filename}'")
        
        # Step 1: Handle common exhibit patterns
        # Ex. A Letter -> Ex_A_Letter
        # Ex.106 -> Ex_106
        # Ex. 55 Email -> Ex_55_Email  
        # Exhibit 12 Memo -> Exhibit_12_Memo
        
        normalized = name
        
        # Replace "Ex." followed by optional spaces with "Ex_" (period + any spaces = one underscore)
        normalized = re.sub(r'^(Ex)\.(\s*)', r'\1_', normalized, flags=re.IGNORECASE)
        
        # Replace "Ex " (space without period) with "Ex_"
        normalized = re.sub(r'^(Ex)\s+', r'\1_', normalized, flags=re.IGNORECASE)
        
        # Replace "Exhibit " with "Exhibit_"
        normalized = re.sub(r'^(Exhibit)\s+', r'\1_', normalized, flags=re.IGNORECASE)
        
        # Step 2: Replace remaining spaces with underscores
        # But be smart about it - don't create double underscores
        normalized = re.sub(r'\s+', '_', normalized)
        
        # Step 3: Clean up any double underscores
        normalized = re.sub(r'_{2,}', '_', normalized)
        
        # Step 4: Remove trailing underscores
        normalized = normalized.rstrip('_')
        
        new_filename = normalized + ext
        
        if new_filename != filename:
            print(f"  Will rename: '{filename}' -> '{new_filename}'")
        else:
            print(f"  No change needed: '{filename}'")
        
        return new_filename
    
    @staticmethod
    def rename_files_in_folder(folder_path, dry_run=True):
        """
        Rename files in folder for Chrome compatibility
        
        Args:
            folder_path: Path to folder containing files
            dry_run: If True, only show what would be renamed without actually renaming
            
        Returns:
            tuple: (successful_renames, failed_renames, unchanged_files)
        """
        if not os.path.exists(folder_path):
            raise Exception(f"Folder does not exist: {folder_path}")
        
        try:
            files = os.listdir(folder_path)
        except Exception as e:
            raise Exception(f"Cannot read folder: {e}")
        
        successful_renames = []
        failed_renames = []
        unchanged_files = []
        
        print(f"\n{'DRY RUN - ' if dry_run else ''}Processing files in: {folder_path}")
        print(f"Found {len(files)} files")
        
        for filename in files:
            # Skip directories
            full_path = os.path.join(folder_path, filename)
            if os.path.isdir(full_path):
                continue
            
            new_filename = FileRenamer.normalize_filename(filename)
            
            if new_filename == filename:
                unchanged_files.append(filename)
                continue
            
            new_full_path = os.path.join(folder_path, new_filename)
            
            # Check if target filename already exists
            if os.path.exists(new_full_path):
                error_msg = f"Target file already exists: {new_filename}"
                failed_renames.append((filename, new_filename, error_msg))
                print(f"  ✗ CONFLICT: {error_msg}")
                continue
            
            if not dry_run:
                try:
                    os.rename(full_path, new_full_path)
                    successful_renames.append((filename, new_filename))
                    print(f"  ✓ RENAMED: '{filename}' -> '{new_filename}'")
                except Exception as e:
                    failed_renames.append((filename, new_filename, str(e)))
                    print(f"  ✗ FAILED: '{filename}' -> '{new_filename}' ({e})")
            else:
                successful_renames.append((filename, new_filename))
                print(f"  ✓ WOULD RENAME: '{filename}' -> '{new_filename}'")
        
        print(f"\nSummary:")
        print(f"  Files that would be renamed: {len(successful_renames)}")
        print(f"  Files that would fail: {len(failed_renames)}")
        print(f"  Files unchanged: {len(unchanged_files)}")
        
        return successful_renames, failed_renames, unchanged_files


class ExcelAutoLinker:
    def __init__(self):
        self.excel_app = None
        self.workbook = None
        self.worksheet = None
        self.target_folder = None
        self.selected_column_index = None
        self.selected_column_letter = None
        self.excel_file_path = None
        self.mode = "exhibit"  # "exhibit" or "bates"
        self.bates_prefix = ""
        self.bates_pdf_map = {}
        self.use_black_hyperlinks = False
        self.page_automation_enabled = False
        self.exemplary_citation = ""
        self.exemplary_page_number = None
        self.page_pattern_regex = None
        self.exhibit_group_index = None
        self.page_group_index = None

        
        # Exhibit patterns (reuse from Word class)
        self.exhibit_patterns = [
            # More flexible patterns with word boundaries
            r'\bEx\.\s*(\d+[A-Z]?)\b',        # Ex. 1, Ex. 2, Ex. 1A, Ex. 2B (word boundaries)
            r'\bEx\.\s*([A-Z])\b',            # Ex. A, Ex. B (single letters only)
            r'\bExhibit\s*(\d+[A-Z]?)\b',     # Exhibit 1, Exhibit 2, Exhibit 1A, Exhibit 2B
            r'\bExhibit\s*([A-Z])\b',         # Exhibit A, Exhibit B (single letters only)
            
            # NEW: Additional flexible patterns for Excel
            r'\bEx\.(\d+[A-Z]?)\b',           # Ex.1, Ex.2A (no space)
            r'\bEx\.([A-Z])\b',               # Ex.A, Ex.B (no space)
            r'\bEx\s+(\d+[A-Z]?)\b',          # Ex 1, Ex 2A (space instead of period)
            r'\bEx\s+([A-Z])\b',              # Ex A, Ex B (space instead of period)
            r'\bEx_(\d+[A-Z]?)\b',            # Ex_1, Ex_2A (underscore)
            r'\bEx_([A-Z])\b',                # Ex_A, Ex_B (underscore)
        ]

        
        # Track created hyperlinks
        self.created_hyperlinks = []

    def set_black_hyperlinks(self, use_black):
        """Set whether to use black hyperlinks"""
        self.use_black_hyperlinks = use_black
        print(f"Black hyperlinks mode: {'enabled' if use_black else 'disabled'}")


    def initialize_excel(self):
        """Initialize Excel COM application - FIXED to stay hidden"""
        try:
            print("Initializing Excel COM application...")
            pythoncom.CoInitialize()
            
            try:
                self.excel_app = win32com.client.GetActiveObject("Excel.Application")
                print("Connected to existing Excel instance")
            except:
                self.excel_app = win32com.client.Dispatch("Excel.Application")
                print("Created new Excel instance")
            
            # CRITICAL FIX: Keep Excel completely hidden
            self.excel_app.Visible = False
            self.excel_app.DisplayAlerts = False
            self.excel_app.ScreenUpdating = False  # Disable screen updates for performance
            self.excel_app.EnableEvents = False    # Disable events for performance
            
            # Additional settings to ensure Excel stays hidden
            try:
                self.excel_app.WindowState = -4140  # xlMinimized
            except:
                pass  # Some versions might not support this
            
            workbook_count = self.excel_app.Workbooks.Count
            print(f"Excel initialized successfully (hidden). Current workbooks: {workbook_count}")
            
            return True
            
        except Exception as e:
            print(f"Error initializing Excel: {e}")
            raise Exception(f"Could not initialize Microsoft Excel: {str(e)}")

    def select_excel_file(self):
        """Select Excel file to process - FIXED to create working copy"""
        if not self.initialize_excel():
            return None
            
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")]
        )
        
        if not file_path:
            return None
            
        try:
            print(f"Opening Excel file: {file_path}")
            
            # Convert to absolute path and normalize
            abs_path = os.path.abspath(file_path)
            self.original_excel_path = abs_path  # Store original path
            print(f"Absolute path: {abs_path}")
            
            # Close existing workbooks
            if self.workbook:
                try:
                    self.workbook.Close(SaveChanges=False)
                except:
                    pass
            
            # CRITICAL FIX: Create working copy like Word version does
            print("Creating working copy of Excel file...")
            
            # Generate working copy name
            original_dir = os.path.dirname(abs_path)
            original_name = os.path.basename(abs_path)
            name_without_ext = os.path.splitext(original_name)[0]
            ext = os.path.splitext(original_name)[1]
            
            # Create working copy name
            mode_suffix = "_with_bates_links" if self.mode == "bates" else "_with_exhibit_links"
            working_copy_name = f"{name_without_ext}{mode_suffix}{ext}"
            working_copy_path = os.path.join(original_dir, working_copy_name)
            
            # Handle existing files by adding counter
            counter = 1
            while os.path.exists(working_copy_path):
                working_copy_name = f"{name_without_ext}{mode_suffix}_{counter}{ext}"
                working_copy_path = os.path.join(original_dir, working_copy_name)
                counter += 1
            
            print(f"Working copy path: {working_copy_path}")
            
            # Create the working copy using file system copy
            import shutil
            shutil.copy2(abs_path, working_copy_path)
            print("Working copy created successfully")
            
            # Now open the WORKING COPY for editing (not the original)
            print("Opening working copy for editing...")
            self.workbook = self.excel_app.Workbooks.Open(working_copy_path)
            self.worksheet = self.workbook.ActiveSheet
            
            # Store paths - IMPORTANT: Keep track of both original and working copy
            self.excel_file_path = working_copy_path  # Point to working copy for processing
            self.working_copy_path = working_copy_path
            
            print(f"Excel working copy opened successfully")
            print(f"Active sheet: {self.worksheet.Name}")
            print(f"Original file remains untouched at: {abs_path}")
            print(f"Working on copy at: {working_copy_path}")
            
            return file_path  # Return original path for display purposes
            
        except Exception as e:
            print(f"Error opening Excel file: {e}")
            raise Exception(f"Could not open Excel file: {str(e)}")

    def get_column_letter(self, col_index):
        """Convert column index to letter (1=A, 2=B, etc.)"""
        result = ""
        while col_index > 0:
            col_index -= 1
            result = chr(col_index % 26 + ord('A')) + result
            col_index //= 26
        return result

    def get_available_columns(self):
        """Get list of available columns with their headers"""
        if not self.worksheet:
            return []
        
        try:
            used_range = self.worksheet.UsedRange
            if used_range.Rows.Count < 1:
                return []
            
            columns = []
            first_row = used_range.Rows(1)
            
            for i in range(1, first_row.Columns.Count + 1):
                try:
                    cell_value = first_row.Cells(1, i).Value
                    if cell_value is None:
                        cell_value = f"(Empty)"
                    
                    column_letter = self.get_column_letter(i)
                    columns.append({
                        'index': i,
                        'letter': column_letter,
                        'header': str(cell_value),
                        'display': f"Column {column_letter}: {cell_value}"
                    })
                except Exception as e:
                    print(f"Error reading column {i}: {e}")
                    continue
            
            return columns
            
        except Exception as e:
            print(f"Error getting columns: {e}")
            return []

    def set_mode(self, mode, bates_prefix=""):
        """Set processing mode"""
        self.mode = mode
        self.bates_prefix = bates_prefix.strip()
        if mode == "bates" and self.target_folder:
            self.build_bates_pdf_map()

    def build_bates_pdf_map(self):
        """Build mapping of Bates PDFs - reuse logic from Word class"""
        self.bates_pdf_map = {}
        
        if not self.target_folder or not self.bates_prefix:
            return
        
        try:
            files_in_folder = os.listdir(self.target_folder)
            bates_files = []
            
            bates_pattern = rf'^{re.escape(self.bates_prefix)}(\d+)\.pdf$'
            
            for filename in files_in_folder:
                match = re.match(bates_pattern, filename, re.IGNORECASE)
                if match:
                    bates_number = int(match.group(1))
                    full_path = os.path.join(self.target_folder, filename)
                    bates_files.append((bates_number, filename, full_path))
            
            bates_files.sort(key=lambda x: x[0])
            
            for i, (bates_number, filename, full_path) in enumerate(bates_files):
                self.bates_pdf_map[bates_number] = {
                    'filename': filename,
                    'path': full_path,
                    'start_page': bates_number
                }
            
            print(f"Built Bates PDF map for {len(bates_files)} files")
                
        except Exception as e:
            print(f"Error building Bates PDF map: {e}")

    def set_page_automation(self, enabled, exemplary_citation="", exemplary_page_number=None):
        """Configure page number automation for Excel"""
        self.page_automation_enabled = enabled
        self.exemplary_citation = exemplary_citation.strip()
        self.exemplary_page_number = exemplary_page_number
        
        if enabled and self.exemplary_citation and self.exemplary_page_number:
            self.build_page_pattern()
        else:
            self.page_pattern_regex = None
            self.exhibit_group_index = None
            self.page_group_index = None

    def build_page_pattern(self):
        """Build flexible regex pattern for legal citation formats"""
        try:
            citation = self.exemplary_citation
            page_num = str(self.exemplary_page_number)
            
            print(f"Building flexible pattern from: '{citation}' with page {page_num}")
            
            # Step 1: Find exhibit identifier using existing patterns
            exhibit_id = None
            for pattern in self.exhibit_patterns:
                match = re.search(pattern, citation, re.IGNORECASE)
                if match:
                    exhibit_id = match.group(1)
                    print(f"  Found exhibit ID: '{exhibit_id}'")
                    break
            
            if not exhibit_id:
                print("  ERROR: Could not find exhibit identifier in exemplary citation")
                return False
            
            # Step 2: Find page number with common legal citation patterns
            page_patterns = [
                rf'\bat\s+p\.?\s*{re.escape(page_num)}\b',      # "at p. 25", "at p 25"
                rf'\bat\s+pp\.?\s*{re.escape(page_num)}\b',     # "at pp. 25", "at pp 25" 
                rf'\bat\s+{re.escape(page_num)}\b',             # "at 25"
                rf'\bp\.?\s*{re.escape(page_num)}\b',           # "p. 25", "p 25"
                rf'\bpp\.?\s*{re.escape(page_num)}\b',          # "pp. 25", "pp 25"
                rf'\bpage\s+{re.escape(page_num)}\b',           # "page 25"
                rf'\bpages?\s+{re.escape(page_num)}\b',         # "pages 25"
                rf'\b{re.escape(page_num)}\b'                   # just "25" (fallback)
            ]
            
            page_match_info = None
            for i, pattern in enumerate(page_patterns):
                matches = list(re.finditer(pattern, citation, re.IGNORECASE))
                if matches:
                    # Use the last match (most likely the page reference)
                    page_match_info = {
                        'match': matches[-1],
                        'pattern_type': i,
                        'full_match': matches[-1].group(0)
                    }
                    print(f"  Found page reference: '{page_match_info['full_match']}' using pattern type {i}")
                    break
            
            if not page_match_info:
                print(f"  ERROR: Could not find page number '{page_num}' in citation")
                return False
            
            # Step 3: Build flexible pattern that matches legal citation formats
            # Pattern structure: Ex. [ID] [misc text] [page reference with NUMBER]
            
            # Create exhibit capture group (flexible for numbers, letters, combinations)
            exhibit_capture = r'(\d+[A-Za-z]*|[A-Za-z]+\d*|[A-Za-z]+)'
            
            # Create page number capture group based on the pattern type found
            pattern_type = page_match_info['pattern_type']
            if pattern_type == 0:  # "at p. NUMBER" - cover both p. and pp.
                page_capture = r'at\s+pp?\.?\s*(\d+)'
            elif pattern_type == 1:  # "at pp. NUMBER" - cover both p. and pp.
                page_capture = r'at\s+pp?\.?\s*(\d+)'
            elif pattern_type == 2:  # "at NUMBER"
                page_capture = r'at\s+(\d+)'
            elif pattern_type == 3:  # "p. NUMBER"
                page_capture = r'p\.?\s*(\d+)'
            elif pattern_type == 4:  # "pp. NUMBER"
                page_capture = r'pp\.?\s*(\d+)'
            elif pattern_type == 5:  # "page NUMBER"
                page_capture = r'page\s+(\d+)'
            elif pattern_type == 6:  # "pages NUMBER"
                page_capture = r'pages?\s+(\d+)'
            else:  # just NUMBER (fallback)
                page_capture = r'(\d+)'
            

            # Use negative lookahead to stop at next exhibit boundary, allow more flexible middle content
            full_pattern = rf'(?:Ex\.|Exhibit)\s*{exhibit_capture}(?:(?!(?:Ex\.|Exhibit)\s*\w)[^.]*?){page_capture}'
            loose_pattern = rf'{page_capture}'
            
            print(f"  Generated strict pattern: {full_pattern}")
            print(f"  Generated loose pattern: {loose_pattern}")
            
            print(f"  Generated flexible pattern: {full_pattern}")
            
            # Step 4: Test the pattern
            test_match = re.search(full_pattern, citation, re.IGNORECASE)
            if not test_match:
                # Try with DOTALL flag for multiline text
                print("  Trying with DOTALL flag for multiline matching...")
                test_match = re.search(full_pattern, citation, re.IGNORECASE | re.DOTALL)
            if test_match and len(test_match.groups()) >= 2:
                found_exhibit = test_match.group(1)
                found_page = test_match.group(2)
                
                print(f"  Pattern test successful:")
                print(f"    Group 1 (exhibit): '{found_exhibit}'")
                print(f"    Group 2 (page): '{found_page}'")
                
                # Verify the groups match our expected values
                if (found_exhibit.upper() == exhibit_id.upper() and 
                    found_page == page_num):
                    self.page_pattern_regex = full_pattern
                    self.exhibit_group_index = 1
                    self.page_group_index = 2
                    print("  ✓ Flexible pattern validation successful!")
                    return True
                else:
                    print("  ✗ Pattern validation failed - groups don't match expected values")
                    print(f"    Expected: exhibit='{exhibit_id}', page='{page_num}'")
                    print(f"    Found: exhibit='{found_exhibit}', page='{found_page}'")
                    return False
            else:
                print("  ✗ Pattern test failed - doesn't match original citation or insufficient groups")
                return False
                
        except Exception as e:
            print(f"Error building flexible pattern: {e}")
            import traceback
            traceback.print_exc()
            return False

    def find_page_number_in_text(self, text):
            """Extract exhibit ID and page number using the pattern - ENHANCED VERSION"""
            if not self.page_automation_enabled or not self.page_pattern_regex:
                print(f"    Page automation disabled or no pattern (enabled={self.page_automation_enabled}, pattern={self.page_pattern_regex is not None})")
                return None, None
            
            try:
                print(f"    Testing pattern '{self.page_pattern_regex}' against text: '{text}'")
                
                # Find ALL pattern matches, not just the first one
                matches = list(re.finditer(self.page_pattern_regex, text, re.IGNORECASE))
                
                if matches:
                    print(f"    Found {len(matches)} pattern matches:")
                    for i, match in enumerate(matches):
                        exhibit_id = match.group(self.exhibit_group_index)
                        page_number = int(match.group(self.page_group_index))
                        match_text = match.group(0)
                        print(f"      Match {i+1}: '{match_text}' -> Exhibit '{exhibit_id}' Page {page_number}")
                    
                    # Return all matches for the caller to decide which one to use
                    results = []
                    for match in matches:
                        exhibit_id = match.group(self.exhibit_group_index)
                        page_number = int(match.group(self.page_group_index))
                        results.append((exhibit_id, page_number, match))
                    
                    return results
                else:
                    print(f"    ✗ Pattern did not match text")
                    return []
                    
            except Exception as e:
                print(f"    ✗ Error in page number extraction: {e}")
                import traceback
                traceback.print_exc()
            
            return []


    def find_bates_pdf_and_page(self, bates_number):
        """Find PDF and page for Bates number - reuse from Word class"""
        if not self.bates_pdf_map:
            return None, None
        
        sorted_starts = sorted(self.bates_pdf_map.keys(), reverse=True)
        
        for start_page in sorted_starts:
            if bates_number >= start_page:
                pdf_info = self.bates_pdf_map[start_page]
                page_in_pdf = bates_number - start_page + 1
                return pdf_info['path'], page_in_pdf
        
        return None, None

    def get_relative_path(self, file_path):
        """Convert to file URL for Excel hyperlinks - FIXED for local files"""
        if not self.excel_file_path:
            return file_path
        
        try:
            print(f"\n=== PATH CONVERSION DEBUG ===")
            print(f"Input file_path: {file_path}")
            
            # Check if it's already a web URL
            if file_path.startswith(('http://', 'https://')):
                print("Web URL detected - returning as-is")
                return file_path
            
            # For local files, ALWAYS use file:// protocol for Excel compatibility
            # Convert to absolute path first to ensure it works
            abs_file_path = os.path.abspath(file_path)
            print(f"Absolute file path: {abs_file_path}")
            
            # Create proper file:// URL - FIXED VERSION
            # Replace backslashes with forward slashes but DON'T encode colons for file://
            file_url = f"file:///{abs_file_path.replace('\\', '/')}"
            print(f"Created file URL: {file_url}")
            
            return file_url
            
        except Exception as e:
            print(f"Error creating file URL: {e}")
            # Fallback - still try to create a file URL without encoding
            try:
                fallback_url = f"file:///{file_path.replace('\\', '/')}"
                print(f"Using fallback file URL: {fallback_url}")
                return fallback_url
            except:
                return file_path

        """Convert to file URL for Excel hyperlinks - FIXED for local files"""
        if not self.excel_file_path:
            return file_path
        
        try:
            print(f"\n=== PATH CONVERSION DEBUG ===")
            print(f"Input file_path: {file_path}")
            
            # Check if it's already a web URL
            if file_path.startswith(('http://', 'https://')):
                print("Web URL detected - returning as-is")
                return file_path
            
            # For local files, ALWAYS use file:// protocol for Excel compatibility
            # Convert to absolute path first to ensure it works
            abs_file_path = os.path.abspath(file_path)
            print(f"Absolute file path: {abs_file_path}")
            
            # Create proper file:// URL
            # Replace backslashes with forward slashes and encode colons
            file_url = f"file:///{abs_file_path.replace('\\', '/').replace(':', '%3A')}"
            print(f"Created file URL: {file_url}")
            
            return file_url
            
        except Exception as e:
            print(f"Error creating file URL: {e}")
            # Fallback - still try to create a file URL
            try:
                fallback_url = f"file:///{file_path.replace('\\', '/').replace(':', '%3A')}"
                print(f"Using fallback file URL: {fallback_url}")
                return fallback_url
            except:
                return file_path

    def find_matching_files(self, reference_text):
        """Find matching files based on mode"""
        if not self.target_folder:
            return []
        
        if self.mode == "bates":
            return self.find_matching_bates_files(reference_text)
        else:
            return self.find_matching_exhibit_files(reference_text)

    def find_matching_exhibit_files(self, reference_text):
        """Find exhibit files - ENHANCED VERSION with flexible naming patterns"""
        matching_files = []
        try:
            files_in_folder = os.listdir(self.target_folder)
            print(f"DEBUG: Files in folder: {files_in_folder}")
        except Exception as e:
            print(f"Error reading folder: {e}")
            return []
        
        print(f"EXCEL PROCESSING: '{reference_text}' (type: {type(reference_text)})")
        
        # Clean up the reference text and handle Excel number conversion
        cleaned_ref = str(reference_text).strip()
        
        # Handle Excel's float conversion (10.0 -> 10, 155.0 -> 155)
        if cleaned_ref.endswith('.0'):
            potential_number = cleaned_ref.replace('.0', '')
            if potential_number.replace('-', '').isdigit():
                cleaned_ref = potential_number
                print(f"FLOAT CONVERSION: '{reference_text}' -> '{cleaned_ref}'")
        
        # Handle cases where Excel gives us a pure number
        try:
            if isinstance(reference_text, (int, float)):
                cleaned_ref = str(int(reference_text))
                print(f"DIRECT NUMBER CONVERSION: {reference_text} -> '{cleaned_ref}'")
            elif cleaned_ref.replace('-', '').replace('.', '').isdigit():
                num_val = float(cleaned_ref)
                if num_val == int(num_val):
                    cleaned_ref = str(int(num_val))
                    print(f"STRING NUMBER CONVERSION: '{reference_text}' -> '{cleaned_ref}'")
        except (ValueError, OverflowError):
            pass
        
        print(f"CLEANED: '{cleaned_ref}'")
        
        # Skip processing if this looks like a header or non-exhibit text
        skip_words = ['exhibit', 'exhibits', 'ex', 'number', 'ref', 'reference', 'document', 'file']
        if cleaned_ref.lower() in skip_words:
            print(f"SKIPPING HEADER/NON-EXHIBIT: '{cleaned_ref}'")
            return []
        
        # Also skip if it's too long to be a reasonable exhibit reference
        if len(cleaned_ref) > 125:
            print(f"SKIPPING TOO LONG: '{cleaned_ref}'")
            return []
        
        # First, try the ENHANCED patterns with word boundaries
        for pattern in self.exhibit_patterns:
            # Use the full original text for pattern matching to get proper context
            match = re.search(pattern, str(reference_text), re.IGNORECASE)
            if match:
                identifier = match.group(1)
                print(f"PATTERN MATCHED: '{reference_text}' -> identifier: '{identifier}'")
                
                # ENHANCED: Try multiple filename patterns
                possible_prefixes = [
                    f"Ex. {identifier}",     # Ex. 1, Ex. A
                    f"Ex.{identifier}",      # Ex.1, Ex.A
                    f"Ex {identifier}",      # Ex 1, Ex A
                    f"Ex_{identifier}",      # Ex_1, Ex_A
                    f"Exhibit {identifier}", # Exhibit 1, Exhibit A
                    f"Exhibit_{identifier}", # Exhibit_1, Exhibit_A
                ]
                
                for target_prefix in possible_prefixes:
                    print(f"  Trying prefix: '{target_prefix}'")
                    
                    for filename in files_in_folder:
                        if filename.startswith(target_prefix):
                            prefix_len = len(target_prefix)
                            
                            if prefix_len >= len(filename):
                                # Exact match
                                full_path = os.path.join(self.target_folder, filename)
                                matching_files.append(full_path)
                                print(f"    ✓ EXACT MATCH: '{reference_text}' -> '{filename}'")
                            else:
                                next_char = filename[prefix_len]
                                # Allow common separators and extensions
                                if next_char in ['_', '-', '.', ' ']:
                                    full_path = os.path.join(self.target_folder, filename)
                                    matching_files.append(full_path)
                                    print(f"    ✓ PARTIAL MATCH: '{reference_text}' -> '{filename}'")
                    
                    # Stop if we found matches with this prefix
                    if matching_files:
                        break
                
                # Stop if we found matches with this pattern
                if matching_files:
                    break
        
        # If no matches found with standard patterns, try bare number/letter matching
        if not matching_files:
            print(f"No standard pattern match, trying bare reference...")
            
            if cleaned_ref:
                identifier = None
                
                # Handle pure numbers (155 -> Ex. 155)
                if cleaned_ref.isdigit():
                    identifier = cleaned_ref
                # Handle pure letters (A, B, C) - but only single letters
                elif cleaned_ref.isalpha() and len(cleaned_ref) == 1:
                    identifier = cleaned_ref.upper()
                # Handle alphanumeric combinations (1A, 2B) - reasonable length limit
                elif re.match(r'^[A-Za-z0-9]+$', cleaned_ref) and 1 <= len(cleaned_ref) <= 5:
                    identifier = cleaned_ref.upper()
                
                if identifier:
                    print(f"BARE REFERENCE DETECTED: '{cleaned_ref}' -> identifier: '{identifier}'")
                    
                    # Try the same multiple filename patterns
                    possible_prefixes = [
                        f"Ex. {identifier}",
                        f"Ex.{identifier}",
                        f"Ex {identifier}",
                        f"Ex_{identifier}",
                        f"Exhibit {identifier}",
                        f"Exhibit_{identifier}",
                    ]
                    
                    for target_prefix in possible_prefixes:
                        print(f"  Trying bare prefix: '{target_prefix}'")
                        
                        for filename in files_in_folder:
                            if filename.startswith(target_prefix):
                                prefix_len = len(target_prefix)
                                
                                if prefix_len >= len(filename):
                                    full_path = os.path.join(self.target_folder, filename)
                                    matching_files.append(full_path)
                                    print(f"    ✓ BARE EXACT MATCH: '{cleaned_ref}' -> '{filename}'")
                                else:
                                    next_char = filename[prefix_len]
                                    if next_char in ['_', '-', '.', ' ']:
                                        full_path = os.path.join(self.target_folder, filename)
                                        matching_files.append(full_path)
                                        print(f"    ✓ BARE PARTIAL MATCH: '{cleaned_ref}' -> '{filename}'")
                        
                        # Stop if we found matches
                        if matching_files:
                            break
                else:
                    print(f"BARE REFERENCE REJECTED: '{cleaned_ref}' doesn't match simple patterns")
        
        if not matching_files:
            print(f"✗ NO MATCH FOUND for: '{cleaned_ref}'")
        else:
            print(f"✓ FINAL RESULT: Found {len(matching_files)} matches for '{cleaned_ref}'")
            for match in matching_files:
                print(f"  Matched file: {match}")
        
        return matching_files

    def find_matching_bates_files(self, reference_text):
        """Find Bates files - reuse logic from Word class"""
        matching_files = []
        
        if not self.bates_prefix:
            return []
        
        escaped_prefix = re.escape(self.bates_prefix)
        bates_pattern = rf'{escaped_prefix}(\d+)'
        
        match = re.search(bates_pattern, reference_text, re.IGNORECASE)
        if match:
            bates_number = int(match.group(1))
            pdf_path, page_number = self.find_bates_pdf_and_page(bates_number)
            if pdf_path and page_number:
                matching_files.append({
                    'type': 'bates',
                    'path': pdf_path,
                    'page': page_number,
                    'bates_number': bates_number
                })
        
        return matching_files

    def get_relative_path_for_excel(self, file_path):
        """Convert to relative path for Excel hyperlinks - FIXED VERSION"""
        if not self.excel_file_path:
            return file_path
        
        try:
            print(f"\n=== EXCEL HYPERLINK PATH DEBUG ===")
            print(f"Target file: {file_path}")
            print(f"Excel working copy: {self.excel_file_path}")
            print(f"Original Excel file: {getattr(self, 'original_excel_path', 'Not set')}")
            
            # CRITICAL: Use the original Excel file location for path calculation
            # because that's where the user will likely keep the final files
            if hasattr(self, 'original_excel_path') and self.original_excel_path:
                excel_reference_path = self.original_excel_path
                print(f"Using original file location as reference: {excel_reference_path}")
            else:
                excel_reference_path = self.excel_file_path
                print(f"Using working copy location as reference: {excel_reference_path}")
            
            # Get the directory containing the Excel file
            excel_dir = os.path.dirname(os.path.abspath(excel_reference_path))
            target_dir = os.path.dirname(os.path.abspath(file_path))
            
            print(f"Excel directory: {excel_dir}")
            print(f"Target directory: {target_dir}")
            
            # Check if files are in the same directory
            if os.path.normpath(excel_dir) == os.path.normpath(target_dir):
                # Same directory - just use filename
                relative_path = os.path.basename(file_path)
                print(f"Same directory - using filename: {relative_path}")
                return relative_path
            
            # Calculate relative path from Excel file to target file
            try:
                relative_path = os.path.relpath(file_path, excel_dir)
                print(f"Calculated relative path: {relative_path}")
                
                # Convert to forward slashes for Excel - CRITICAL FIX: Don't URL encode!
                excel_relative_path = relative_path.replace('\\', '/')
                print(f"Excel-formatted path: {excel_relative_path}")
                
                # Verify the path exists
                test_absolute = os.path.abspath(os.path.join(excel_dir, relative_path))
                print(f"Verification - reconstructed absolute path: {test_absolute}")
                print(f"Original file exists: {os.path.exists(file_path)}")
                print(f"Reconstructed path exists: {os.path.exists(test_absolute)}")
                
                return excel_relative_path
                
            except ValueError as e:
                print(f"Relative path calculation failed: {e}")
                # Files are on different drives - use absolute path as file:// URL
                abs_path = os.path.abspath(file_path)
                file_url = f"file:///{abs_path.replace('\\', '/')}"
                print(f"Using absolute file:// URL: {file_url}")
                return file_url
            
        except Exception as e:
            print(f"Error in path calculation: {e}")
            import traceback
            traceback.print_exc()
            # Ultimate fallback - just the filename
            return os.path.basename(file_path)

    def process_excel_column(self):
        """Process selected column for hyperlinks - COMPLETE FIXED VERSION"""
        if not self.worksheet or self.selected_column_index is None:
            return 0
        
        try:
            used_range = self.worksheet.UsedRange
            total_rows = used_range.Rows.Count
            
            print(f"\n=== EXCEL PROCESSING DEBUG ===")
            print(f"Processing column {self.selected_column_letter} in {self.mode} mode")
            print(f"Excel file: {self.excel_file_path}")
            print(f"Target folder: {self.target_folder}")
            print(f"Excel UsedRange reports {total_rows} total rows")
            
            # Check beyond UsedRange to catch data Excel might miss
            extended_check_rows = max(total_rows + 10, 50)
            actual_last_row = total_rows
            
            print(f"Checking extended range up to row {extended_check_rows} to find actual data...")
            
            # Find the real last row with data in our column
            for check_row in range(1, extended_check_rows + 1):
                try:
                    cell = self.worksheet.Cells(check_row, self.selected_column_index)
                    cell_value = cell.Value
                    
                    if cell_value is not None:
                        cell_text = str(cell_value).strip()
                        if cell_text and cell_text.lower() not in ['', 'none', 'null', '#n/a', '#value!', '#ref!']:
                            actual_last_row = max(actual_last_row, check_row)
                            if check_row > total_rows:
                                print(f"  Found data in row {check_row}: '{cell_text}' (beyond Excel's UsedRange!)")
                                
                except Exception as e:
                    break
            
            print(f"Actual last row with data: {actual_last_row}")
            print(f"Will process rows 2 to {actual_last_row} (skipping header row 1)")
            
            if actual_last_row < 2:
                print("No data rows found to process")
                return 0
            
            links_added = 0
            successful_links = []
            failed_links = []
            
            # Process each row
            for row in range(2, actual_last_row + 1):
                try:
                    cell = self.worksheet.Cells(row, self.selected_column_index)
                    cell_value = cell.Value
                    
                    print(f"\n=== ROW {row} ===")
                    print(f"Raw cell_value: {repr(cell_value)} (type: {type(cell_value)})")
                    
                    # Check for various "empty" conditions
                    if cell_value is None:
                        print(f"Row {row}: SKIPPED - cell_value is None")
                        continue
                    
                    # Convert to string and strip whitespace
                    cell_text_raw = str(cell_value).strip()
                    
                    if not cell_text_raw or cell_text_raw.lower() in ['', 'none', 'null', '#n/a', '#value!', '#ref!']:
                        print(f"Row {row}: SKIPPED - empty or error value: '{cell_text_raw}'")
                        continue
                    
                    # Store original value for display
                    original_value = cell_text_raw
                    
                    # Handle Excel's float conversion (10.0 -> 10) for matching AND display
                    cell_text = original_value
                    display_text = original_value  # This will be what shows in the cell
                    
                    if cell_text.endswith('.0') and cell_text.replace('.0', '').replace('-', '').isdigit():
                        cell_text = cell_text.replace('.0', '')
                        display_text = cell_text  # Use the clean version (10) instead of (10.0)
                        print(f"Row {row}: Excel float conversion '{original_value}' -> '{cell_text}' (display: '{display_text}')")
                    
                    # Also handle the case where Excel gives us a float object directly
                    if isinstance(cell_value, float) and cell_value == int(cell_value):
                        display_text = str(int(cell_value))  # Convert 10.0 -> "10"
                        print(f"Row {row}: Direct float conversion {cell_value} -> display: '{display_text}'")
                    
                    print(f"Row {row}: Processing '{cell_text}'")

                    # Find matching files using the converted cell_text
                    matching_files = self.find_matching_files(cell_text)
                    
                    if matching_files:
                        file_info = matching_files[0]
                        print(f"Row {row}: Found matching file: {file_info}")
                        
                        # Create hyperlink based on mode - FIXED FOR BATES PAGE LINKS
                        if isinstance(file_info, dict) and file_info.get('type') == 'bates':
                            # Bates mode - link to specific page
                            target_file = file_info['path']
                            page_number = file_info['page']
                            relative_path = self.get_relative_path_for_excel(target_file)
                            link_target = f"{relative_path}#page={page_number}"
                            screen_tip = f"Bates {file_info['bates_number']} - Page {page_number} of {os.path.basename(target_file)}"
                            print(f"  Bates link target: {link_target}")

                        else:

                            # Check for page number automation first - but only if enabled
                            exhibit_id = None
                            page_number = None

                            if self.page_automation_enabled:
                                page_matches = self.find_page_number_in_text(cell_text)
                                
                                # Get the current exhibit ID for comparison
                                current_exhibit_id = None
                                for pattern in self.exhibit_patterns:
                                    match = re.search(pattern, cell_text, re.IGNORECASE)
                                    if match:
                                        current_exhibit_id = match.group(1)
                                        break
                                
                                # Find the match that corresponds to our current exhibit
                                if page_matches and current_exhibit_id:
                                    for match_exhibit_id, match_page_number, match_obj in page_matches:
                                        if match_exhibit_id.upper() == current_exhibit_id.upper():
                                            exhibit_id, page_number = match_exhibit_id, match_page_number
                                            print(f"  ✓ Found matching exhibit ID: '{exhibit_id}' -> Page {page_number}")
                                            break
                                    
                                    if not exhibit_id:
                                        print(f"  ✗ No matching exhibit ID found for '{current_exhibit_id}' in page automation")

                            
                            if exhibit_id and page_number:
                                # Page automation mode - link to specific page
                                target_file = file_info
                                relative_path = self.get_relative_path_for_excel(target_file)
                                link_target = f"{relative_path}#page={page_number}"
                                screen_tip = f"Link to {os.path.basename(target_file)} page {page_number}"
                                print(f"  Page automation link target: {link_target}")

                            else:
                                # Regular exhibit mode
                                exhibit_id = None
                                page_number = None

                                # Only check for page automation if it's enabled
                                if self.page_automation_enabled:
                                    page_matches = self.find_page_number_in_text(cell_text)
                                    
                                    # Get the current exhibit ID for comparison
                                    current_exhibit_id = None
                                    for pattern in self.exhibit_patterns:
                                        match = re.search(pattern, cell_text, re.IGNORECASE)
                                        if match:
                                            current_exhibit_id = match.group(1)
                                            break
                                    
                                    # Find the match that corresponds to our current exhibit
                                    if page_matches and current_exhibit_id:
                                        for match_exhibit_id, match_page_number, match_obj in page_matches:
                                            if match_exhibit_id.upper() == current_exhibit_id.upper():
                                                exhibit_id, page_number = match_exhibit_id, match_page_number
                                                print(f"  ✓ Found matching exhibit ID: '{exhibit_id}' -> Page {page_number}")
                                                break
                                        
                                        if not exhibit_id:
                                            print(f"  ✗ No matching exhibit ID found for '{current_exhibit_id}' in page automation")
                                
                                print(f"  Page automation result: exhibit_id='{exhibit_id}', page_number={page_number}")
                                
                                # Set up paths and targets
                                target_file = file_info
                                relative_path = self.get_relative_path_for_excel(target_file)
                                
                                if exhibit_id and page_number:
                                    # Page automation mode - link to specific page
                                    link_target = f"{relative_path}#page={page_number}"
                                    screen_tip = f"Link to {os.path.basename(target_file)} page {page_number}"
                                    print(f"  Page automation link target: {link_target}")
                                else:
                                    # Regular exhibit mode - link to file
                                    link_target = relative_path
                                    screen_tip = f"Link to {os.path.basename(file_info)}"
                                    print(f"  Regular exhibit link target: {link_target}")
                        
                        # ENHANCED: Create Excel hyperlink with better debugging
                        try:
                            print(f"  Attempting to create hyperlink:")
                            print(f"    Cell: {cell.Address}")
                            print(f"    Target file: {target_file}")
                            print(f"    Link target: {link_target}")
                            print(f"    Display text: {display_text}")
                            print(f"    Screen tip: {screen_tip}")
                            
                            # Remove any existing hyperlinks first
                            if cell.Hyperlinks.Count > 0:
                                print(f"    Removing {cell.Hyperlinks.Count} existing hyperlinks")
                                cell.Hyperlinks.Delete()
                            
                            # Try the most reliable method for Excel hyperlinks
                            try:
                                # Method 1: Use HYPERLINK formula (most reliable)
                                print(f"    Trying HYPERLINK formula method...")
                                
                                # Escape quotes and special characters
                                safe_display = str(display_text).replace('"', '""')
                                safe_target = str(link_target).replace('"', '""')
                                
                                # Create HYPERLINK formula
                                hyperlink_formula = f'=HYPERLINK("{safe_target}","{safe_display}")'
                                print(f"    Formula: {hyperlink_formula}")

                                # Set the formula
                                cell.Formula = hyperlink_formula

                                # Apply formatting AFTER setting the formula
                                if self.use_black_hyperlinks:
                                    # Force black color and remove underline after setting formula
                                    cell.Font.Color = 0  # Black
                                    cell.Font.Underline = False  # No underline for black mode
                                    print(f"    Applied black formatting")
                                else:
                                    print(f"    Using default hyperlink formatting")
                                
                                print(f"    ✓ HYPERLINK formula method succeeded")
                                links_added += 1
                                successful_links.append({
                                    'row': row,
                                    'cell': cell.Address,
                                    'text': display_text,
                                    'target': target_file,
                                    'relative_path': link_target,
                                    'method': 'HYPERLINK formula'
                                })
                                
                            except Exception as formula_error:
                                print(f"    HYPERLINK formula failed: {formula_error}")
                                
                                try:
                                    # Method 2: Traditional Hyperlinks.Add
                                    print(f"    Trying Hyperlinks.Add method...")
                                    
                                    hyperlink = self.worksheet.Hyperlinks.Add(
                                        Anchor=cell,
                                        Address=link_target,
                                        TextToDisplay=display_text,
                                        ScreenTip=screen_tip
                                    )

                                    # Apply black formatting if needed
                                    if self.use_black_hyperlinks:
                                        cell.Font.Color = 0  # Black
                                        cell.Font.Underline = False  # No underline
                                        print(f"    Applied black formatting to Hyperlinks.Add method")

                                    print(f"    ✓ Hyperlinks.Add method succeeded")
                                    links_added += 1
                                    successful_links.append({
                                        'row': row,
                                        'cell': cell.Address,
                                        'text': display_text,
                                        'target': target_file,
                                        'relative_path': link_target,
                                        'method': 'Hyperlinks.Add'
                                    })
                                    
                                except Exception as add_error:
                                    print(f"    Hyperlinks.Add failed: {add_error}")
                                    
                                    # Method 3: Set value only and log for manual linking
                                    try:
                                        print(f"    Setting cell value without hyperlink...")
                                        cell.Value = display_text
                                        failed_links.append({
                                            'row': row,
                                            'cell': cell.Address,
                                            'text': display_text,
                                            'target': target_file,
                                            'relative_path': link_target,
                                            'error': str(add_error)
                                        })
                                        print(f"    Cell value set (no hyperlink created)")
                                        
                                    except Exception as value_error:
                                        print(f"    Even setting cell value failed: {value_error}")
                                        failed_links.append({
                                            'row': row,
                                            'cell': cell.Address,
                                            'text': display_text,
                                            'target': target_file if 'target_file' in locals() else 'unknown',
                                            'relative_path': link_target if 'link_target' in locals() else 'unknown',
                                            'error': f"All methods failed: {value_error}"
                                        })
                            
                        except Exception as e:
                            print(f"  ✗ Error creating hyperlink for '{cell_text}': {e}")
                            failed_links.append({
                                'row': row,
                                'cell': cell.Address,
                                'text': display_text,
                                'target': target_file if 'target_file' in locals() else 'unknown',
                                'error': str(e)
                            })
                    else:
                        print(f"  ✗ No match found for '{cell_text}'")
                    
                except Exception as e:
                    print(f"Error processing row {row}: {e}")
                    continue
            
            # Summary report
            print(f"\n=== PROCESSING SUMMARY ===")
            print(f"Total hyperlinks created: {links_added}")
            print(f"Successful links: {len(successful_links)}")
            print(f"Failed links: {len(failed_links)}")
            
            if successful_links:
                print(f"\nSuccessful hyperlinks:")
                for link in successful_links:
                    print(f"  Row {link['row']}: '{link['text']}' -> {link['relative_path']} ({link['method']})")
            
            if failed_links:
                print(f"\nFailed hyperlinks:")
                for link in failed_links:
                    print(f"  Row {link['row']}: '{link['text']}' -> {link.get('relative_path', 'unknown')} (Error: {link['error']})")
            
            return links_added
            
        except Exception as e:
            print(f"Error in process_excel_column: {e}")
            import traceback
            traceback.print_exc()
            return 0

    def save_excel_with_links(self, output_path=None):
        """Save Excel file with hyperlinks and export to PDF - ENHANCED CLEANUP VERSION"""
        if not self.workbook or not self.excel_file_path:
            return False, False
        
        try:
            if not output_path:
                # Generate default names
                original_dir = os.path.dirname(self.original_excel_path) if hasattr(self, 'original_excel_path') else os.path.dirname(self.excel_file_path)
                original_name = os.path.basename(self.original_excel_path) if hasattr(self, 'original_excel_path') else os.path.basename(self.excel_file_path)
                name_without_ext = os.path.splitext(original_name)[0]
                ext = os.path.splitext(original_name)[1]
                
                mode_suffix = "_with_bates_links" if self.mode == "bates" else "_with_exhibit_links"
                default_excel_name = f"{name_without_ext}{mode_suffix}{ext}"
                default_pdf_name = f"{name_without_ext}{mode_suffix}.pdf"
                
                print(f"Default save location: {original_dir}")
                print(f"Default Excel name: {default_excel_name}")
                
                # Ask user where to save Excel file
                from tkinter import filedialog
                excel_output = filedialog.asksaveasfilename(
                    title="Save Excel File with Links",
                    defaultextension=ext,
                    filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")],
                    initialdir=original_dir,
                    initialfile=default_excel_name
                )
                
                if not excel_output:
                    print("User cancelled Excel save")
                    return False, False
                
                # Ask user where to save PDF
                pdf_dir = os.path.dirname(excel_output)
                pdf_output = filedialog.asksaveasfilename(
                    title="Save PDF Export",
                    defaultextension=".pdf",
                    filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
                    initialdir=pdf_dir,
                    initialfile=default_pdf_name
                )
                
                if not pdf_output:
                    print("User cancelled PDF save")
                    return False, False
                        
            else:
                excel_output = output_path
                pdf_output = output_path.replace('.xlsx', '.pdf').replace('.xls', '.pdf')
            
            print(f"Attempting to save Excel to: {excel_output}")
            print(f"Attempting to save PDF to: {pdf_output}")
            
            # Save Excel with links
            excel_saved = False
            
            try:
                print("Attempting to save Excel file using temp method...")
                
                # Create temp file
                temp_dir = tempfile.gettempdir()
                temp_filename = f"excel_temp_{int(time.time())}.xlsx"
                temp_path = os.path.join(temp_dir, temp_filename)
                
                print(f"Saving to temp file: {temp_path}")
                
                # Save to temp location
                self.workbook.SaveAs(temp_path, FileFormat=51)  # 51 = xlOpenXMLWorkbook
                print("Temp file saved successfully")
                
                # Verify temp file exists
                if not os.path.exists(temp_path):
                    raise Exception("Temp file was not created")
                
                # Copy from temp to final location
                print(f"Copying from temp to final location: {excel_output}")
                
                # Make sure target directory exists
                os.makedirs(os.path.dirname(excel_output), exist_ok=True)
                
                # Copy the file
                shutil.copy2(temp_path, excel_output)
                
                # Verify final file exists
                if not os.path.exists(excel_output):
                    raise Exception("Final file was not created")
                
                print("Excel file saved successfully")
                
                # Clean up temp file
                try:
                    os.remove(temp_path)
                    print("Temp file cleaned up")
                except:
                    print("Could not clean up temp file (not critical)")
                
                excel_saved = True
                
            except Exception as e:
                print(f"Excel save failed: {e}")
                excel_saved = False

            # Export to PDF
            print(f"Attempting to export PDF: {pdf_output}")
            pdf_saved = False
            
            try:
                # Make sure target directory exists
                os.makedirs(os.path.dirname(pdf_output), exist_ok=True)
                
                # Remove existing PDF if it exists
                if os.path.exists(pdf_output):
                    os.remove(pdf_output)
                
                # Export to PDF using temp method
                temp_pdf_dir = tempfile.gettempdir()
                temp_pdf_name = f"excel_pdf_{int(time.time())}.pdf"
                temp_pdf_path = os.path.join(temp_pdf_dir, temp_pdf_name)
                
                print(f"Exporting to temp PDF: {temp_pdf_path}")
                
                # LANDSCAPE: Set page setup for landscape and fit-to-page before exporting
                print("Configuring page setup for landscape and fit-to-page...")
                try:
                    # Configure the active worksheet's page setup
                    page_setup = self.worksheet.PageSetup
                    
                    # Set to landscape orientation
                    page_setup.Orientation = 2  # xlLandscape (1 = xlPortrait, 2 = xlLandscape)
                    
                    # Fit all columns on one page
                    page_setup.FitToPagesWide = 1  # Fit to 1 page wide
                    page_setup.FitToPagesTall = False  # Allow multiple pages tall if needed
                    
                    # Optional: Set to fit all content on one page (both width and height)
                    # Uncomment the next line if you want everything on exactly one page
                    # page_setup.FitToPagesTall = 1
                    
                    # Ensure we're not using scaling (use fit-to-page instead)
                    page_setup.Zoom = False  # Disable zoom to enable fit-to-page
                    
                    # Set reasonable margins for more content space
                    page_setup.LeftMargin = 36   # 0.5 inch in points (72 points per inch)
                    page_setup.RightMargin = 36  # 0.5 inch
                    page_setup.TopMargin = 54    # 0.75 inch
                    page_setup.BottomMargin = 54 # 0.75 inch
                    
                    print("✓ Page setup configured: Landscape, fit all columns to width")
                    
                except Exception as setup_error:
                    print(f"Warning: Could not configure page setup: {setup_error}")
                    print("PDF will use default settings")
                
                # Use ExportAsFixedFormat with enhanced settings
                self.workbook.ExportAsFixedFormat(
                    Type=0,  # xlTypePDF
                    Filename=temp_pdf_path,
                    Quality=0,  # xlQualityStandard
                    IncludeDocProperties=True,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False
                )
                
                print("Temp PDF created successfully")
                
                # Verify temp PDF exists
                if not os.path.exists(temp_pdf_path):
                    raise Exception("Temp PDF was not created")
                
                # Copy to final location
                print(f"Copying PDF from temp to final location: {pdf_output}")
                shutil.copy2(temp_pdf_path, pdf_output)
                
                # Verify final PDF exists
                if not os.path.exists(pdf_output):
                    raise Exception("Final PDF was not created")
                
                print("PDF export completed successfully")
                
                # Clean up temp PDF
                try:
                    os.remove(temp_pdf_path)
                    print("Temp PDF cleaned up")
                except:
                    print("Could not clean up temp PDF (not critical)")
                
                pdf_saved = True
                
            except Exception as e:
                print(f"PDF export failed: {e}")
                pdf_saved = False
                
            return excel_saved, pdf_saved
            
        except Exception as e:
            print(f"Error in save_excel_with_links: {e}")
            return False, False

    def cleanup(self):
        """Clean up Excel COM objects and remove working copy file"""
        try:
            print("Starting Excel cleanup...")
            
            # Store working copy path before closing
            working_copy_to_delete = None
            if hasattr(self, 'working_copy_path') and self.working_copy_path:
                working_copy_to_delete = self.working_copy_path
                print(f"Will delete working copy: {working_copy_to_delete}")
            
            # Close workbook first
            if self.workbook:
                try:
                    workbook_name = self.workbook.Name
                    print(f"Closing workbook: {workbook_name}")
                    self.workbook.Close(SaveChanges=False)
                    print("Workbook closed successfully")
                except Exception as e:
                    print(f"Error closing workbook: {e}")
                finally:
                    self.workbook = None
            
            # Quit Excel application
            if self.excel_app:
                try:
                    # Re-enable settings before quitting
                    self.excel_app.ScreenUpdating = True
                    self.excel_app.EnableEvents = True
                    
                    # Close any remaining workbooks
                    while self.excel_app.Workbooks.Count > 0:
                        try:
                            wb = self.excel_app.Workbooks(1)
                            wb_name = wb.Name
                            print(f"Force closing: {wb_name}")
                            wb.Close(SaveChanges=False)
                        except Exception as e:
                            print(f"Error force closing workbook: {e}")
                            break
                    
                    print("Quitting Excel application...")
                    self.excel_app.Quit()
                    print("Excel quit successfully")
                    
                except Exception as e:
                    print(f"Error quitting Excel: {e}")
                finally:
                    self.excel_app = None
            
            # Force COM cleanup
            import gc
            gc.collect()
            
            try:
                pythoncom.CoUninitialize()
                print("COM uninitialized")
            except Exception as e:
                print(f"Error uninitializing COM: {e}")
            
            # CRITICAL FIX: Delete the working copy file after Excel is closed
            if working_copy_to_delete and os.path.exists(working_copy_to_delete):
                try:
                    print(f"Deleting working copy file: {working_copy_to_delete}")
                    
                    # Wait a moment for Excel to fully release the file
                    import time
                    time.sleep(1)
                    
                    # Try to delete the file
                    os.remove(working_copy_to_delete)
                    print("✓ Working copy file deleted successfully")
                    
                except Exception as e:
                    print(f"✗ Could not delete working copy file: {e}")
                    print("You may need to delete it manually")
            
            print("Excel cleanup completed")
            
        except Exception as e:
            print(f"Error during cleanup: {e}")
            import traceback
            traceback.print_exc()

class ExhibitAnchorApp:
    def __init__(self):
            self.root = ttk.Window(themename="cosmo")
            self.root.title("Exhibit Linker")
            
            # Set custom icon for main window if available
            try:
                icon_path = os.path.join(os.path.dirname(__file__), "images", "icon.ico")
                if os.path.exists(icon_path):
                    self.root.iconbitmap(icon_path)
                    self.icon_path = icon_path  # Store for use in other windows
                else:
                    self.icon_path = None
            except Exception as e:
                print(f"Could not set main window icon: {e}")
                self.icon_path = None
            
            # DYNAMIC SIZING: Detect screen resolution and set appropriate window size
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            # Calculate window size based on screen resolution
            if screen_width <= 1366:  # Small/laptop screens
                window_width = min(1050, int(screen_width * 0.9))  # 750 * 1.4 = 40% bigger
                window_height = min(800, int(screen_height * 0.85))
                self.is_small_screen = True
            elif screen_width <= 1920:  # Standard screens
                window_width = 1120  # 800 * 1.4 = 40% bigger
                window_height = 980
                self.is_small_screen = False
            else:  # Large/high-res screens
                window_width = 1190  # 850 * 1.4 = 40% bigger
                window_height = 1000
                self.is_small_screen = False
            
            print(f"Detected screen: {screen_width}x{screen_height}")
            print(f"Setting window size: {window_width}x{window_height}")
            
            self.root.geometry(f"{window_width}x{window_height}")
            self.root.resizable(True, True)
            
            # Set minimum window size to prevent UI breaking
            self.root.minsize(600, 500)
            
            # Store window dimensions for responsive elements
            self.window_width = window_width
            self.window_height = window_height
            
            self.use_black_hyperlinks = tk.BooleanVar(value=False)
            
            # Center the window on screen
            self.center_window()

            # Maximize the window on startup
            self.root.state('zoomed')  # For Windows
            
            # Initialize processors lazily
            self.word_linker = None
            self.excel_linker = None
            
            # UI variables
            self.processing_mode = tk.StringVar(value="word")
            self.word_submode_var = tk.StringVar(value="exhibit")  
            self.doc_path = tk.StringVar(value="No document selected")
            self.folder_path = tk.StringVar(value="No folder selected")
            self.status_text = tk.StringVar(value="Ready to process documents")
            
            # Mode-specific variables
            self.bates_prefix_var = tk.StringVar()
            self.selected_column_var = tk.StringVar(value="No column selected")
            self.excel_submode_var = tk.StringVar(value="exhibit")

            self.page_automation_var = tk.BooleanVar(value=False)
            self.exemplary_citation_var = tk.StringVar()
            self.exemplary_page_var = tk.StringVar()

            self.info_text_var = tk.StringVar()
            
            # Dynamic UI elements (will be created as needed)
            self.bates_prefix_frame = None
            self.excel_controls_frame = None
            self.column_selection_frame = None
            
            self.create_widgets()
            
            # Cleanup on close
            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
  
    def center_window(self):
        """Center the window on the screen with dynamic sizing"""
        self.root.update_idletasks()
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Get actual window size (might be different from requested if screen is small)
        window_width = self.root.winfo_width()
        window_height = self.root.winfo_height()
        
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        # Ensure window stays on screen
        x = max(0, min(x, screen_width - window_width))
        y = max(0, min(y, screen_height - window_height))
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def get_responsive_font_size(self, base_size):
        """Get font size based on screen size - LESS AGGRESSIVE"""
        if self.is_small_screen:
            return max(8, base_size - 1)  # Only reduce by 1 for small screens
        else:
            return base_size  # Keep original size for normal/large screens

    def get_responsive_padding(self, base_padding):
        """Get padding based on screen size - LESS AGGRESSIVE"""
        if self.is_small_screen:
            return max(8, base_padding - 3)  # Reduce less padding
        else:
            return base_padding  # Keep original padding for normal/large screens

    def get_responsive_wraplength(self, base_length):
        """Get text wrap length based on window width - IMPROVED"""
        if self.is_small_screen:
            return min(500, self.window_width - 80)  # Less aggressive reduction
        else:
            return min(base_length, self.window_width - 50)  # Scale with window width

    def on_mode_changed(self):
        """Handle mode selection changes - RESPONSIVE VERSION"""
        mode = self.processing_mode.get()
        
        # Hide all dynamic controls first
        if hasattr(self, 'word_controls_frame') and self.word_controls_frame:
            self.word_controls_frame.pack_forget()
        if hasattr(self, 'excel_controls_frame') and self.excel_controls_frame:
            self.excel_controls_frame.pack_forget()
        
        # Update UI based on mode
        if mode == "word":
            self.step1_frame.config(text="Step 1: Select Word Document")
            self.doc_label_text.config(text="Selected Document:")
            self.browse_doc_button.config(text="Browse Word Docs")
            self.word_controls_frame.pack(fill=X, pady=(5, 0))
            
            # Hide Excel column selection
            if hasattr(self, 'excel_column_frame'):
                if self.is_small_screen:
                    self.excel_column_frame.pack_forget()
                else:
                    if hasattr(self, 'excel_separator_frame'):
                        self.excel_separator_frame.grid_forget()
                    self.excel_column_frame.grid_forget()
            
            self.on_word_submode_changed()
            self.status_text.set("Ready to process Word document")
            
        elif mode == "excel":
            self.step1_frame.config(text="Step 1: Select Excel File & Column")
            self.doc_label_text.config(text="Selected Excel File:")
            self.browse_doc_button.config(text="Browse Excel File")
            self.excel_controls_frame.pack(fill=X, pady=(5, 0))
            
            # Show column selection
            if hasattr(self, 'excel_column_frame'):
                if self.is_small_screen:
                    self.excel_column_frame.pack(fill=X, pady=(10, 0))
                else:
                    if hasattr(self, 'excel_separator_frame'):
                        self.excel_separator_frame.grid(row=0, column=1, sticky="ns", padx=(10, 10))
                    self.excel_column_frame.grid(row=0, column=2, sticky="nw", padx=(10, 0))
            
            self.on_excel_submode_changed()
            self.status_text.set("Ready to process Excel file")
        
        # Update info text and reset file selection
        self.update_info_text()
        self.doc_path.set("No document selected")

    def on_word_submode_changed(self):
        """Handle Word sub-mode changes"""
        if self.processing_mode.get() != "word":
            return
            
        submode = self.word_submode_var.get()
        
        if submode == "bates":
            if hasattr(self, 'word_bates_frame'):
                self.word_bates_frame.pack(fill=X, pady=(5, 0))
            # Hide page automation in Bates mode
            if hasattr(self, 'page_auto_check'):
                self.page_auto_check.pack_forget()
            if hasattr(self, 'page_automation_frame'):
                self.page_automation_frame.pack_forget()
        else:
            if hasattr(self, 'word_bates_frame'):
                self.word_bates_frame.pack_forget()
            # Show page automation in Exhibit mode
            if hasattr(self, 'page_auto_check'):
                self.page_auto_check.pack(anchor='w', pady=(2, 0))
            if hasattr(self, 'page_automation_frame'):
                self.page_automation_frame.pack(fill=X, pady=(5, 0))
        
        self.update_info_text()        

    def on_excel_submode_changed(self):
        """Handle Excel sub-mode changes"""
        if self.processing_mode.get() != "excel":
            return
            
        submode = self.excel_submode_var.get()
        
        if submode == "bates":
            if hasattr(self, 'excel_bates_frame'):
                self.excel_bates_frame.pack(fill=X, pady=(5, 0))
            # Hide page automation in Bates mode
            if hasattr(self, 'page_auto_check'):
                self.page_auto_check.pack_forget()
            if hasattr(self, 'page_automation_frame'):
                self.page_automation_frame.pack_forget()
        else:
            if hasattr(self, 'excel_bates_frame'):
                self.excel_bates_frame.pack_forget()
            # Show page automation in Exhibit mode
            if hasattr(self, 'page_auto_check'):
                self.page_auto_check.pack(anchor='w', pady=(2, 0))
            if hasattr(self, 'page_automation_frame'):
                self.page_automation_frame.pack(fill=X, pady=(5, 0))
        
        self.update_info_text()

    def on_bates_prefix_changed(self, *args):
        """Handle Bates prefix changes"""
        prefix = self.bates_prefix_var.get().strip()
        if prefix:
            if self.processing_mode.get() == "word_bates":
                self.status_text.set(f"Word/Bates mode with prefix: '{prefix}'")
            elif self.processing_mode.get() == "excel" and self.excel_submode_var.get() == "bates":
                self.status_text.set(f"Excel/Bates mode with prefix: '{prefix}'")

    def on_page_automation_toggled(self):
        """Handle page automation toggle - controls enablement, not visibility"""
        if self.page_automation_var.get():
            # Enable the controls
            if hasattr(self, 'citation_entry'):
                self.citation_entry.config(state='normal')
            if hasattr(self, 'page_entry'):
                self.page_entry.config(state='normal')
            self.status_text.set("Page automation enabled - enter exemplary citation and page number")
        else:
            # Disable the controls but keep them visible
            if hasattr(self, 'citation_entry'):
                self.citation_entry.config(state='disabled')
            if hasattr(self, 'page_entry'):
                self.page_entry.config(state='disabled')
            self.status_text.set("Page automation disabled")

    def show_file_renamer_dialog(self):
        """Show file renaming dialog for Chrome PDF compatibility"""
        # Check if we have a target folder
        folder_path = None
        
        # Try to get folder from current processor
        mode = self.processing_mode.get()
        if mode == "word":
            linker = self.get_word_linker()
            if linker and linker.target_folder:
                folder_path = linker.target_folder
        elif mode == "excel":
            linker = self.get_excel_linker()
            if linker and linker.target_folder:
                folder_path = linker.target_folder
        
        # If no folder selected, let user choose
        if not folder_path:
            folder_path = filedialog.askdirectory(
                title="Select Folder to Rename Files",
                initialdir="."
            )
            if not folder_path:
                return
        
        try:
            # First, do a dry run to show what would happen
            successful, failed, unchanged = FileRenamer.rename_files_in_folder(folder_path, dry_run=True)
            
            if not successful and not failed:
                messagebox.showinfo("No Changes Needed", 
                    "No files in this folder need renaming for Chrome compatibility.")
                return
            
            # Create preview dialog
            preview_dialog = tk.Toplevel(self.root)
            preview_dialog.title("File Renaming Preview - Chrome PDF Compatibility")
            preview_dialog.geometry("800x675")

            # Set custom icon if available
            if hasattr(self, 'icon_path') and self.icon_path:
                try:
                    preview_dialog.iconbitmap(self.icon_path)
                except Exception as e:
                    print(f"Could not set preview dialog icon: {e}")

            preview_dialog.transient(self.root)
            preview_dialog.grab_set()
            preview_dialog.resizable(True, True)
            
            # Center dialog
            preview_dialog.update_idletasks()
            x = (preview_dialog.winfo_screenwidth() - 800) // 2
            y = (preview_dialog.winfo_screenheight() - 600) // 2
            preview_dialog.geometry(f"800x675+{x}+{y}")
            
            # Main frame
            main_frame = ttk.Frame(preview_dialog, padding=20)
            main_frame.pack(fill=BOTH, expand=True)
            
            # Title and explanation
            title_label = ttk.Label(
                main_frame, 
                text="File Renaming Preview - Chrome PDF Compatibility", 
                font=("Helvetica", 14, "bold")
            )
            title_label.pack(pady=(0, 10))
            
            explanation = ttk.Label(
                main_frame,
                text="This tool standardizes filenames to improve Chrome PDF link compatibility.\n" +
                    "Chrome sometimes has issues with periods and spaces in filenames when following hyperlinks.\n" +
                    "Examples: 'Ex. A Letter.pdf' → 'Ex._A_Letter.pdf', 'Ex. 55 Email.docx' → 'Ex._55_Email.docx'",
                font=("Helvetica", 10),
                justify=CENTER,
                wraplength=750
            )
            explanation.pack(pady=(0, 15))
            
            # Folder info
            folder_label = ttk.Label(
                main_frame,
                text=f"Folder: {folder_path}",
                font=("Helvetica", 9),
                bootstyle="secondary"
            )
            folder_label.pack(pady=(0, 15))
            
            # Create notebook for different categories
            notebook = ttk.Notebook(main_frame)
            notebook.pack(fill=BOTH, expand=True, pady=(0, 15))
            
            # Tab 1: Files to be renamed
            if successful:
                rename_frame = ttk.Frame(notebook)
                notebook.add(rename_frame, text=f"Files to Rename ({len(successful)})")
                
                # Scrollable list
                rename_list_frame = ttk.Frame(rename_frame)
                rename_list_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
                
                rename_text = tk.Text(
                    rename_list_frame,
                    wrap=tk.NONE,
                    font=("Consolas", 9),
                    bg="#f8f9fa"
                )
                
                rename_scrollbar_y = ttk.Scrollbar(rename_list_frame, orient=tk.VERTICAL, command=rename_text.yview)
                rename_scrollbar_x = ttk.Scrollbar(rename_list_frame, orient=tk.HORIZONTAL, command=rename_text.xview)
                rename_text.config(yscrollcommand=rename_scrollbar_y.set, xscrollcommand=rename_scrollbar_x.set)
                
                # Add content
                for old_name, new_name in successful:
                    rename_text.insert(tk.END, f"'{old_name}'\n  → '{new_name}'\n\n")
                
                rename_text.config(state=tk.DISABLED)
                
                # Pack scrollbars and text
                rename_text.pack(side=tk.LEFT, fill=BOTH, expand=True)
                rename_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
                rename_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
            
            # Tab 2: Conflicts/Failures
            if failed:
                failed_frame = ttk.Frame(notebook)
                notebook.add(failed_frame, text=f"Conflicts ({len(failed)})")
                
                failed_list_frame = ttk.Frame(failed_frame)
                failed_list_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
                
                failed_text = tk.Text(
                    failed_list_frame,
                    wrap=tk.WORD,
                    font=("Consolas", 9),
                    bg="#fff5f5"
                )
                
                failed_scrollbar = ttk.Scrollbar(failed_list_frame, orient=tk.VERTICAL, command=failed_text.yview)
                failed_text.config(yscrollcommand=failed_scrollbar.set)
                
                for old_name, new_name, error in failed:
                    failed_text.insert(tk.END, f"'{old_name}' → '{new_name}'\nError: {error}\n\n")
                
                failed_text.config(state=tk.DISABLED)
                
                failed_text.pack(side=tk.LEFT, fill=BOTH, expand=True)
                failed_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Tab 3: Unchanged files
            if unchanged:
                unchanged_frame = ttk.Frame(notebook)
                notebook.add(unchanged_frame, text=f"No Changes Needed ({len(unchanged)})")
                
                unchanged_list_frame = ttk.Frame(unchanged_frame)
                unchanged_list_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
                
                unchanged_text = tk.Text(
                    unchanged_list_frame,
                    wrap=tk.WORD,
                    font=("Consolas", 9),
                    bg="#f0fff0"
                )
                
                unchanged_scrollbar = ttk.Scrollbar(unchanged_list_frame, orient=tk.VERTICAL, command=unchanged_text.yview)
                unchanged_text.config(yscrollcommand=unchanged_scrollbar.set)
                
                for filename in unchanged:
                    unchanged_text.insert(tk.END, f"'{filename}'\n")
                
                unchanged_text.config(state=tk.DISABLED)
                
                unchanged_text.pack(side=tk.LEFT, fill=BOTH, expand=True)
                unchanged_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Buttons frame
            buttons_frame = ttk.Frame(main_frame)
            buttons_frame.pack(pady=(10, 0))
            
            # Result storage
            result = [False]  # Use list to modify from inner functions
            
            def proceed_with_rename():
                try:
                    # Perform actual rename
                    actual_successful, actual_failed, _ = FileRenamer.rename_files_in_folder(folder_path, dry_run=False)
                    
                    if actual_failed:
                        error_summary = "\n".join([f"'{old}' → '{new}': {error}" for old, new, error in actual_failed])
                        messagebox.showerror("Some Renames Failed", 
                            f"Successfully renamed {len(actual_successful)} files.\n\n" +
                            f"Failed to rename {len(actual_failed)} files:\n{error_summary}")
                    else:
                        messagebox.showinfo("Rename Complete", 
                            f"Successfully renamed {len(actual_successful)} files for Chrome PDF compatibility!")
                    
                    result[0] = True
                    preview_dialog.destroy()
                    
                except Exception as e:
                    messagebox.showerror("Rename Failed", f"Error during renaming: {str(e)}")
            
            def cancel_rename():
                result[0] = False
                preview_dialog.destroy()
            
            # Buttons
            if successful:
                ttk.Button(
                    buttons_frame,
                    text=f"Rename {len(successful)} Files",
                    command=proceed_with_rename,
                    bootstyle="warning",
                    width=20
                ).pack(side=tk.LEFT, padx=(0, 10))
            
            ttk.Button(
                buttons_frame,
                text="Cancel",
                command=cancel_rename,
                bootstyle="secondary",
                width=15
            ).pack(side=tk.LEFT)
            
            # Warning if there are conflicts
            if failed:
                warning_frame = ttk.Frame(main_frame)
                warning_frame.pack(pady=(10, 0))
                
                warning_label = ttk.Label(
                    warning_frame,
                    text=f"⚠️ {len(failed)} files have conflicts and will be skipped",
                    font=("Helvetica", 10, "bold"),
                    bootstyle="warning"
                )
                warning_label.pack()
            
            # Wait for dialog to close
            self.root.wait_window(preview_dialog)
            
            # Update status and refresh folder info if renames were made
            if result[0]:
                self.status_text.set(f"File renaming completed - {len(successful)} files renamed for Chrome compatibility")
                # Refresh folder status
                if mode == "word" and linker and linker.target_folder:
                    self.update_folder_status(linker.target_folder, mode)
                elif mode == "excel" and linker and linker.target_folder:
                    self.update_folder_status(linker.target_folder, mode)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error analyzing files: {str(e)}")

    def show_help_popup(self):
        """Show help information popup with comparison table"""
        # Create help dialog
        help_dialog = tk.Toplevel(self.root)
        help_dialog.title("Help - Export Information")
        
        # Set custom icon if available
        if hasattr(self, 'icon_path') and self.icon_path:
            try:
                help_dialog.iconbitmap(self.icon_path)
            except Exception as e:
                print(f"Could not set help dialog icon: {e}")
        
        help_dialog.transient(self.root)
        help_dialog.grab_set()
        help_dialog.resizable(True, True)
        
        # Set dialog dimensions
        dialog_width = 790
        dialog_height = 650  # Increased height from 500 to 650
        
        # Center dialog on screen
        help_dialog.update_idletasks()
        screen_width = help_dialog.winfo_screenwidth()
        screen_height = help_dialog.winfo_screenheight()
        x = (screen_width - dialog_width) // 2
        y = (screen_height - dialog_height) // 2
        
        # Make sure dialog stays on screen
        x = max(0, min(x, screen_width - dialog_width))
        y = max(0, min(y, screen_height - dialog_height))
        
        help_dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        
        # Main frame with padding
        main_frame = ttk.Frame(help_dialog, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="Export Information & Hyperlink Types", 
            font=("Helvetica", 14, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # Create comparison table using Frame with borders instead of custom styles
        table_frame = ttk.Frame(main_frame, relief="solid", borderwidth=1)
        table_frame.pack(fill=X, pady=(0, 20))
        
        # Helper function to create table cells
        def create_table_cell(parent, text, row, col, is_header=False, wraplength=None):
            cell_frame = ttk.Frame(parent, relief="solid", borderwidth=1)
            cell_frame.grid(row=row, column=col, sticky="nsew", padx=0, pady=0)
            
            if is_header:
                cell_frame.configure(style="info.TFrame")  # Use ttkbootstrap style
                label = ttk.Label(cell_frame, text=text, font=("Helvetica", 10, "bold"), 
                                anchor="center", background="#e8f4fd")
            else:
                label = ttk.Label(cell_frame, text=text, font=("Helvetica", 9), 
                                anchor="center", wraplength=wraplength if wraplength else 200)
            
            label.pack(fill=BOTH, expand=True, padx=8, pady=6)
            return cell_frame
        
        # Create table grid
        # Headers
        create_table_cell(table_frame, "Input", 0, 0, is_header=True)
        create_table_cell(table_frame, "Word", 0, 1, is_header=True)
        create_table_cell(table_frame, "Excel", 0, 2, is_header=True)
        create_table_cell(table_frame, "PDF", 0, 3, is_header=True)
        
        # Row 1: Word
        create_table_cell(table_frame, "Word", 1, 0)
        create_table_cell(table_frame, "Relative Links", 1, 1)
        create_table_cell(table_frame, "N/A", 1, 2)
        create_table_cell(table_frame, "Relative Links", 
                        1, 3, wraplength=300)
        
        # Row 2: Excel
        create_table_cell(table_frame, "Excel", 2, 0)
        create_table_cell(table_frame, "N/A", 2, 1)
        create_table_cell(table_frame, "Relative Links", 2, 2)
        create_table_cell(table_frame, "Relative Links", 2, 3, wraplength=300)
        
        # Configure grid weights for proper resizing
        for i in range(4):
            table_frame.grid_columnconfigure(i, weight=1)
        for i in range(3):
            table_frame.grid_rowconfigure(i, weight=1)
        
        # Separator
        ttk.Separator(main_frame, orient='horizontal').pack(fill=X, pady=(10, 20))
        
        # Create scrollable text area for explanations
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=BOTH, expand=True)
        
        # Text widget
        text_widget = tk.Text(
            text_frame,
            wrap=tk.WORD,
            font=("Helvetica", 10),
            bg="#f8f9fa",
            fg="#2c3e50",
            relief=tk.FLAT,
            borderwidth=0,
            padx=15,
            pady=15,
            height=8
        )
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.config(yscrollcommand=scrollbar.set)
        
        # Pack text and scrollbar
        text_widget.pack(side=tk.LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Help content - focused on link types
        help_content = """Relative Hyperlinks: These hyperlinks will work if the file is brought to another location/PC so long as the PDF and linked files are in the same relative orientation. E.g., if your PDF with hyperlinks is in a folder called "Memo" and the exhibits are in the "Exhibits" subfolder thereof, so long as that basic orientation is retained, the linking should remain. As such, these hyperlinks are ideal if you are sending your files along to another individual.

Page Citations: This script can create citations to page numbers (be it via the Page Citation or Bates stamp functions) in output PDF and Excel files.  The links will bring you to the correct page if opened via Chrome or most other browsers.  Acrobat, however, is incompatible with page linking so a specific page hyperlink opened via Acrobat will simply bring you to page 1 of the operative exhibit.  Note also that page citation mode only works with convention page numbers (e.g., page 1 or p. 55) and not alphaneumeric descriptors like JOHN_005.

Filenames:  If you are linking to files in the same  folder as the "parent" document, you should process the filenames which will replace periods/spaces with underscores.  This ensures compatability with browser PDF viewers which can misinterpret hyperlinks otherwise.  If the exhibits are in a subfolder, however, there should be no need to process the filenames.

Word:  The PDF output by this script will work out of the box.  However, if you add links to your Word document and wish to make further modifications and then save your Word document as a PDF, you should use Word's Save As feature and save as PDF.  Do not use Save as Adobe PDF or Print to PDF as they can strip away hyperlink functionality. 

File Linking:  Linking is done by searching for specific terms in the document. For example, if you have a Word document with references like 'Ex. 1', 'Exhibit A', or Bates numbers like 'SMITH_011', the script will automatically convert these into clickable hyperlinks that point to the corresponding files in the same folder.  For exhibits, this can include longer file names, so, e.g., a cite to Ex. 55 would link to both Ex. 55.pdf and Ex. 55 Letter to the Court.pdf"""


        # Insert content
        text_widget.insert(tk.END, help_content)
        
        # Apply formatting to specific terms
        def format_bold_term(term):
            start = "1.0"
            while True:
                pos = text_widget.search(term, start, tk.END)
                if not pos:
                    break
                end_pos = f"{pos}+{len(term)}c"
                text_widget.tag_add("bold_header", pos, end_pos)
                start = end_pos
        
        # Apply bold formatting to headers
        format_bold_term("Relative Hyperlinks:")
        format_bold_term("Page Citations:")
        format_bold_term("File Linking:")
        format_bold_term("Filenames:")
        format_bold_term("Word:")
        
        # Configure the bold style
        text_widget.tag_config("bold_header", font=("Helvetica", 10, "bold"))
        
        # Make text widget read-only
        text_widget.config(state=tk.DISABLED)
        
        # Close button
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(15, 0))
        
        ttk.Button(
            button_frame,
            text="Close",
            command=help_dialog.destroy,
            bootstyle="primary",
            width=15
        ).pack()


    def update_info_text(self):
        """Update information panel based on current mode - UPDATED"""
        mode = self.processing_mode.get()
        
        if mode == "word":
            submode = self.word_submode_var.get()
            if submode == "exhibit":
                info_text = """Word/Exhibit Mode: Converts exhibit references to clickable hyperlinks in Word documents.

    Features:
    • Finds references like 'Ex. 1', 'Exhibit A' in your Word document
    • Links to files named 'Ex. 1.pdf', 'Exhibit A.docx', etc.
    • Exports Word doc + PDF with relative links that work on any computer
    • Original Word document remains unchanged

    Requirements: Save output files in same folder as original Word document."""
            else:  # bates
                info_text = """Word/Bates Mode: Links Bates numbers to specific pages in PDF files.

    Features:
    • Enter Bates prefixes (muse use underscores like 'SMITH_') to match cites
    • Documents must be consecutively paginated across range
    • Exports to Word and PDF with page-specific links (Chrome opens to exact page, Acrobat to first page of the relevant PDF)
    • If SMITH_011 is 5 pages, the script knows that a cite to SMITH_012 is to the 2nd page of SMITH_011.pdf"""
                
        else:  # excel mode
            submode = self.excel_submode_var.get()
            if submode == "exhibit":
                info_text = """Excel/Exhibit Mode: Adds hyperlinks to exhibit references in Excel columns.

    Features:
    • Select column containing exhibit references
    • Converts numbers and letters (whether preceded by Ex./Exhibit or not) to links to 'Ex. A.pdf', 'Ex. 1.docx'
    • Citations to Bates numbers bring you to operative PDF (even if Bates number is mid-PDF)
    • Exports Excel and PDF files with working links 
    • Links in PDF work when files are moved together """
            else:
                info_text = """Excel/Bates Mode: Links Bates numbers in Excel to specific documents.

    Features:
    • Enter Bates prefix to match your PDFs (must use underscores like 'SMITH_')
    • Select column with Bates numbers
    • Exports Excel + PDF with working links
    • Exports to PDF with page-specific links (Chrome opens to exact page, Acrobat to first page of the relevant PDF)

    Requirement: Bates PDFs must be numbered sequentially"""
        
        info_text += "\n\nCopyright © Alex Owens, 2025. All rights reserved."
        self.info_text_var.set(info_text)

    def get_word_linker(self):
        """Get or create Word linker"""
        if self.word_linker is None:
            try:
                self.word_linker = WordAutoLinkerCOM()
            except Exception as e:
                messagebox.showerror("Error", str(e))
                return None
        return self.word_linker

    def get_excel_linker(self):
        """Get or create Excel linker"""
        if self.excel_linker is None:
            try:
                self.excel_linker = ExcelAutoLinker()
            except Exception as e:
                messagebox.showerror("Error", str(e))
                return None
        return self.excel_linker

    def browse_document(self):
        """Handle document/file selection based on mode - SIMPLIFIED"""
        mode = self.processing_mode.get()
        
        if mode == "word":
            self.browse_word_document()
        elif mode == "excel":
            self.browse_excel_file()

    def browse_word_document(self):
        """Browse for Word document - UPDATED to check sub-mode"""
        linker = self.get_word_linker()
        if not linker:
            return
            
        try:
            # Set mode in linker based on sub-mode
            is_bates = self.word_submode_var.get() == "bates"  # Changed from processing_mode
            prefix = self.bates_prefix_var.get() if is_bates else ""
            linker.set_bates_mode(is_bates, prefix)
            
            self.status_text.set("Creating working copy of document...")
            self.root.update()
            
            file_path = linker.select_word_document()
            if file_path:
                original_name = os.path.basename(file_path)
                name_without_ext = os.path.splitext(original_name)[0]
                mode_suffix = "_with_bates_links" if is_bates else "_with_links"
                working_copy_display = f"{name_without_ext}{mode_suffix} (working copy)"
                self.doc_path.set(working_copy_display)
                self.folder_path.set(os.path.dirname(file_path))
                
                folder = os.path.dirname(file_path)
                file_count = len([f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))])
                
                mode_text = "Bates mode" if is_bates else "Exhibit mode"
                self.status_text.set(f"Working copy created in {mode_text} - {file_count} files in folder")
            else:
                self.status_text.set("No document selected")
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting document: {str(e)}")
            self.status_text.set("Error selecting document")

    def browse_excel_file(self):
        """Browse for Excel file"""
        linker = self.get_excel_linker()
        if not linker:
            return
            
        try:
            self.status_text.set("Opening Excel file...")
            self.root.update()
            
            file_path = linker.select_excel_file()
            if file_path:
                original_name = os.path.basename(file_path)
                self.doc_path.set(original_name)
                self.folder_path.set(os.path.dirname(file_path))
                
                # Enable column selection
                self.select_column_button.config(state='normal')
                
                folder = os.path.dirname(file_path)
                file_count = len([f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))])
                
                self.status_text.set(f"Excel file opened - {file_count} files in folder - select column to process")
            else:
                self.status_text.set("No Excel file selected")
        except Exception as e:
            messagebox.showerror("Error", f"Error opening Excel file: {str(e)}")
            self.status_text.set("Error opening Excel file")

    def select_excel_column(self):
        """Show dialog to select Excel column"""
        linker = self.get_excel_linker()
        if not linker or not linker.worksheet:
            messagebox.showwarning("Warning", "Please select an Excel file first")
            return
        
        try:
            columns = linker.get_available_columns()
            if not columns:
                messagebox.showerror("Error", "No columns found in Excel file")
                return
            
            # Create column selection dialog
            dialog = tk.Toplevel(self.root)
            dialog.title("Select Column")
            dialog.geometry("400x300")

            # Set custom icon if available
            if hasattr(self, 'icon_path') and self.icon_path:
                try:
                    dialog.iconbitmap(self.icon_path)
                except Exception as e:
                    print(f"Could not set column dialog icon: {e}")

            dialog.transient(self.root)
            dialog.grab_set()
            
            # Center dialog
            dialog.update_idletasks()
            x = (dialog.winfo_screenwidth() - 400) // 2
            y = (dialog.winfo_screenheight() - 300) // 2
            dialog.geometry(f"400x300+{x}+{y}")
            
            ttk.Label(dialog, text="Select column to process:", font=("Helvetica", 12, "bold")).pack(pady=10)
            
            # Listbox for columns
            listbox_frame = ttk.Frame(dialog)
            listbox_frame.pack(fill=BOTH, expand=True, padx=20, pady=10)
            
            listbox = tk.Listbox(listbox_frame, font=("Helvetica", 10))
            scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=listbox.yview)
            listbox.config(yscrollcommand=scrollbar.set)
            
            listbox.pack(side=tk.LEFT, fill=BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Populate listbox
            for col in columns:
                listbox.insert(tk.END, col['display'])
            
            # Select first item by default
            if columns:
                listbox.selection_set(0)
            
            # Buttons
            button_frame = ttk.Frame(dialog)
            button_frame.pack(pady=10)
            
            selected_column = [None]  # Use list to modify from inner function
            
            def on_select():
                selection = listbox.curselection()
                if selection:
                    selected_column[0] = columns[selection[0]]
                    dialog.destroy()
            
            def on_cancel():
                dialog.destroy()
            
            ttk.Button(button_frame, text="Select", command=on_select, bootstyle="primary").pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="Cancel", command=on_cancel, bootstyle="secondary").pack(side=tk.LEFT, padx=5)
            
            # Wait for dialog to close
            self.root.wait_window(dialog)
            
            # Process selection
            if selected_column[0]:
                col_info = selected_column[0]
                linker.selected_column_index = col_info['index']
                linker.selected_column_letter = col_info['letter']
                
                self.selected_column_var.set(f"Column {col_info['letter']}: {col_info['header']}")
                self.status_text.set(f"Column {col_info['letter']} selected - ready to process")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting column: {str(e)}")

    def browse_folder(self):
        """Handle file folder selection based on mode - UPDATED"""
        mode = self.processing_mode.get()
        
        if mode == "word":
            linker = self.get_word_linker()
            if not linker or not linker.doc:
                messagebox.showwarning("Warning", "Please select a Word document first")
                return
        elif mode == "excel":
            linker = self.get_excel_linker()
            if not linker or not linker.excel_file_path:
                messagebox.showwarning("Warning", "Please select an Excel file first")
                return
        else:
            return
        
        # Update mode settings in linker
        if mode == "word" and self.word_submode_var.get() == "bates":
            linker.set_bates_mode(True, self.bates_prefix_var.get())
        elif mode == "excel" and self.excel_submode_var.get() == "bates":
            linker.set_mode("bates", self.bates_prefix_var.get())
        elif mode == "excel":
            linker.set_mode("exhibit", "")
        else:
            linker.set_bates_mode(False, "")
        
        folder_path = None
        if hasattr(linker, 'select_exhibit_folder'):
            folder_path = linker.select_exhibit_folder()
        else:
            # For Excel, use file dialog
            initial_dir = os.path.dirname(linker.excel_file_path) if linker.excel_file_path else "."
            folder_path = filedialog.askdirectory(
                title="Select Files Folder",
                initialdir=initial_dir
            )
            if folder_path:
                linker.target_folder = folder_path
        
        if folder_path:
            self.folder_path.set(folder_path)
            self.update_folder_status(folder_path, mode)

    def use_document_folder(self):
        """Use document's folder for files - UPDATED"""
        mode = self.processing_mode.get()
        
        if mode == "word":
            linker = self.get_word_linker()
            if not linker or not linker.doc_folder:
                messagebox.showwarning("Warning", "Please select a Word document first")
                return
            folder_path = linker.doc_folder
            linker.target_folder = folder_path
        elif mode == "excel":
            linker = self.get_excel_linker()
            if not linker or not linker.excel_file_path:
                messagebox.showwarning("Warning", "Please select an Excel file first")
                return
            folder_path = os.path.dirname(linker.excel_file_path)
            linker.target_folder = folder_path
        else:
            return
        
        # Update mode settings
        if mode == "word" and self.word_submode_var.get() == "bates":
            linker.set_bates_mode(True, self.bates_prefix_var.get())
        elif mode == "excel" and self.excel_submode_var.get() == "bates":
            linker.set_mode("bates", self.bates_prefix_var.get())
        elif mode == "excel":
            linker.set_mode("exhibit", "")
        else:
            linker.set_bates_mode(False, "")
        
        self.folder_path.set(folder_path)
        self.update_folder_status(folder_path, mode)


    def update_folder_status(self, folder_path, mode):
        """Update status based on folder selection and mode - UPDATED"""
        try:
            file_count = len([f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))])
            
            # Check for Bates mode in either Word or Excel
            is_bates_mode = False
            if mode == "word" and self.word_submode_var.get() == "bates":
                is_bates_mode = True
            elif mode == "excel" and self.excel_submode_var.get() == "bates":
                is_bates_mode = True
            
            if is_bates_mode:
                prefix = self.bates_prefix_var.get().strip()
                if prefix:
                    bates_count = len([f for f in os.listdir(folder_path) 
                                    if f.startswith(prefix) and f.endswith('.pdf')])
                    self.status_text.set(f"Folder selected - {bates_count} Bates PDFs found with prefix '{prefix}' ({file_count} total files)")
                else:
                    self.status_text.set(f"Folder selected - enter Bates prefix ({file_count} total files)")
            else:
                exhibit_count = len([f for f in os.listdir(folder_path) if f.startswith('Ex.')])
                self.status_text.set(f"Folder selected - {exhibit_count} exhibit files found ({file_count} total files)")
        except Exception as e:
            self.status_text.set(f"Error reading folder: {e}")

    def create_progress_dialog(self, title):
        """Create a modern progress dialog with enhanced positioning"""
        class ProgressDialog:
            def __init__(self, parent, title):
                self.parent = parent
                self.dialog = tk.Toplevel(parent)
                self.dialog.title(title)
                
                # Set custom icon - try to find it directly
                try:
                    import os
                    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "images", "icon.ico")
                    if os.path.exists(icon_path):
                        self.dialog.iconbitmap(icon_path)
                        print(f"Progress dialog icon set from: {icon_path}")
                    else:
                        print(f"Icon not found at: {icon_path}")
                except Exception as e:
                    print(f"Could not set progress dialog icon: {e}")
                
                self.dialog.transient(parent)
                self.dialog.grab_set()
                self.dialog.resizable(False, False)
                
                # Store original parent position
                try:
                    self.original_parent_geometry = parent.geometry()
                except:
                    self.original_parent_geometry = None
                
                # Set initial size
                dialog_width = 500
                dialog_height = 150
                
                # ENHANCED CENTERING - Multiple methods for reliability
                self.dialog.withdraw()  # Hide while positioning
                
                # Method 1: Center relative to parent window
                try:
                    parent.update_idletasks()
                    parent_x = parent.winfo_rootx()
                    parent_y = parent.winfo_rooty()
                    parent_width = parent.winfo_width()
                    parent_height = parent.winfo_height()
                    
                    # Calculate center position relative to parent
                    x = parent_x + (parent_width - dialog_width) // 2
                    y = parent_y + (parent_height - dialog_height) // 2
                    
                    # Make sure it stays on screen
                    screen_width = self.dialog.winfo_screenwidth()
                    screen_height = self.dialog.winfo_screenheight()
                    
                    x = max(0, min(x, screen_width - dialog_width))
                    y = max(0, min(y, screen_height - dialog_height))
                    
                    print(f"Centering progress dialog at: {x}, {y} (relative to parent)")
                    
                except Exception as e:
                    print(f"Parent-relative centering failed: {e}, using screen center")
                    
                    # Method 2: Fallback to screen center
                    self.dialog.update_idletasks()
                    screen_width = self.dialog.winfo_screenwidth()
                    screen_height = self.dialog.winfo_screenheight()
                    
                    x = (screen_width - dialog_width) // 2
                    y = (screen_height - dialog_height) // 2
                
                # Set geometry and show
                self.dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
                self.dialog.deiconify()  # Show the dialog
                
                # Force to front and center again (belt and suspenders)
                self.dialog.lift()
                self.dialog.focus_force()
                self.dialog.update()
                
                # Main frame
                main_frame = ttk.Frame(self.dialog, padding=20)
                main_frame.pack(fill=BOTH, expand=True)
                
                # Title
                self.title_label = ttk.Label(
                    main_frame,
                    text=title,
                    font=("Helvetica", 12, "bold")
                )
                self.title_label.pack(pady=(0, 10))
                
                # Status text
                self.status_label = ttk.Label(
                    main_frame,
                    text="Initializing...",
                    font=("Helvetica", 10)
                )
                self.status_label.pack(pady=(0, 10))
                
                # Progress bar
                self.progress_bar = ttk.Progressbar(
                    main_frame,
                    mode='determinate',
                    length=400,
                    bootstyle="success-striped"
                )
                self.progress_bar.pack(pady=(0, 10))
                
                # Percentage label
                self.percent_label = ttk.Label(
                    main_frame,
                    text="0%",
                    font=("Helvetica", 9),
                    bootstyle="secondary"
                )
                self.percent_label.pack()
                
                # Make dialog stay on top and prevent closing
                self.dialog.protocol("WM_DELETE_WINDOW", lambda: None)
                self.dialog.attributes('-topmost', True)
            
            def update_progress(self, percent, status_text):
                """Update progress bar and status"""
                try:
                    self.progress_bar['value'] = percent
                    self.status_label.config(text=status_text)
                    self.percent_label.config(text=f"{percent}%")
                    self.dialog.update()
                except:
                    pass  # Dialog might be destroyed
            
            def hide_temporarily(self):
                """Hide the progress dialog temporarily"""
                try:
                    self.dialog.withdraw()
                    print("Progress dialog hidden for file save dialogs")
                except:
                    pass
            
            def show_again(self):
                """Show the progress dialog again"""
                try:
                    self.dialog.deiconify()
                    self.dialog.lift()
                    self.dialog.focus_force()
                    print("Progress dialog restored after file save")
                except:
                    pass
            
            def close(self):
                """Close the progress dialog"""
                try:
                    self.dialog.destroy()
                except:
                    pass
        
        return ProgressDialog(self.root, title)

    def process_document(self):
        """Handle processing based on mode - SIMPLIFIED"""
        mode = self.processing_mode.get()
        
        if mode == "word":
            self.process_word_document()
        elif mode == "excel":
            self.process_excel_document()

    def process_word_document(self):
        """Process Word document with enhanced progress tracking"""
        linker = self.get_word_linker()
        if not linker or not linker.doc or not linker.target_folder:
            messagebox.showerror("Error", "Please select a Word document and files folder first")
            return
        
        # Validate Bates mode requirements based on sub-mode
        if self.word_submode_var.get() == "bates":
            prefix = self.bates_prefix_var.get().strip()
            if not prefix:
                messagebox.showerror("Error", "Please enter a Bates prefix for Bates mode")
                return
            linker.set_bates_mode(True, prefix)
        else:
            linker.set_bates_mode(False, "")
        
        # DON'T CREATE PROGRESS DIALOG HERE ANYMORE!
        
        # Set black hyperlinks option
        linker.set_black_hyperlinks(self.use_black_hyperlinks.get())
        
        # Configure page automation if enabled
        if not self.word_submode_var.get() == "bates" and self.page_automation_var.get():
            citation = self.exemplary_citation_var.get().strip()
            page_str = self.exemplary_page_var.get().strip()
            
            if citation and page_str:
                try:
                    page_num = int(page_str)
                    linker.set_page_automation(True, citation, page_num)
                except ValueError:
                    messagebox.showerror("Error", "Page number must be a valid integer")
                    return
            else:
                messagebox.showerror("Error", "Please enter both exemplary citation and page number for page automation")
                return
        else:
            linker.set_page_automation(False)
        
        # NOW CREATE THE PROGRESS DIALOG HERE - AFTER ALL CHECKS!
        progress_dialog = self.create_progress_dialog("Processing Word Document")
        
        mode_text = "Bates mode" if self.word_submode_var.get() == "bates" else "Exhibit mode"
        
        def progress_callback(percent, status_text):
            """Update progress dialog"""
            try:
                progress_dialog.update_progress(percent, status_text)
                self.root.update()  # Keep GUI responsive
            except:
                pass  # Dialog might be closed
        
        try:
            # Process with progress updates
            links_added = linker.process_document(progress_callback)
            
            # Update progress for saving
            progress_callback(100, "Processing complete! Now saving files...")
            
            if links_added is not None and links_added >= 0:
                # PASS PROGRESS DIALOG TO SAVE METHOD
                if linker.save_document(progress_dialog=progress_dialog):
                    # Close progress dialog after everything is done
                    progress_dialog.close()
                    
                    link_type = "Bates links" if self.word_submode_var.get() == "bates" else "exhibit links"
                    self.status_text.set(f"Success! {links_added} {link_type} added. Files saved.")
                    
                    success_message = f"Word document processed successfully!\n\n"
                    success_message += f"• {links_added} relative hyperlinks added\n"
                    success_message += f"• Mode: {mode_text}\n"
                    success_message += f"• PDF and Word files saved with links\n"
                    success_message += f"• Original document unchanged"
                    
                    messagebox.showinfo("Processing Complete", success_message)

                    # Job complete ASCII art in console
                    print("\n")
                    print("     ██╗ ██████╗ ██████╗      ██████╗ ██████╗ ███╗   ███╗██████╗ ██╗     ███████╗████████╗███████╗██╗")
                    print("     ██║██╔═══██╗██╔══██╗    ██╔════╝██╔═══██╗████╗ ████║██╔══██╗██║     ██╔════╝╚══██╔══╝██╔════╝██║")
                    print("     ██║██║   ██║██████╔╝    ██║     ██║   ██║██╔████╔██║██████╔╝██║     █████╗     ██║   █████╗  ██║")
                    print("██   ██║██║   ██║██╔══██╗    ██║     ██║   ██║██║╚██╔╝██║██╔═══╝ ██║     ██╔══╝     ██║   ██╔══╝  ╚═╝")
                    print("╚█████╔╝╚██████╔╝██████╔╝    ╚██████╗╚██████╔╝██║ ╚═╝ ██║██║     ███████╗███████╗   ██║   ███████╗██╗")
                    print(" ╚════╝  ╚═════╝ ╚═════╝      ╚═════╝ ╚═════╝ ╚═╝     ╚═╝╚═╝     ╚══════╝╚══════╝   ╚═╝   ╚══════╝╚═╝")
                    print(f"Word processing complete. {links_added} links created.\n")
                else:
                    progress_dialog.close()
                    self.status_text.set("Document processed but not saved")
                    messagebox.showwarning("Warning", "Document processed but not saved.")
            else:
                progress_dialog.close()
                self.status_text.set("Processing completed with errors")
                messagebox.showwarning("Warning", "Processing completed but may have encountered errors.")
                
        except Exception as e:
            progress_dialog.close()
            self.status_text.set(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Processing failed: {str(e)}")

    def process_excel_document(self):
        """Process Excel document"""
        linker = self.get_excel_linker()
        if not linker or not linker.excel_file_path or not linker.target_folder:
            messagebox.showerror("Error", "Please select an Excel file and files folder first")
            return
        
        if linker.selected_column_index is None:
            messagebox.showerror("Error", "Please select a column to process")
            return
        
        # Set mode and validate Bates requirements
        submode = self.excel_submode_var.get()
        if submode == "bates":
            prefix = self.bates_prefix_var.get().strip()
            if not prefix:
                messagebox.showerror("Error", "Please enter a Bates prefix for Bates mode")
                return
            linker.set_mode("bates", prefix)
        else:
            linker.set_mode("exhibit", "")
        
        # Set black hyperlinks option
        linker.set_black_hyperlinks(self.use_black_hyperlinks.get())
        
        # Configure page automation if enabled - CHECK FIRST!
        if submode != "bates" and self.page_automation_var.get():
            citation = self.exemplary_citation_var.get().strip()
            page_str = self.exemplary_page_var.get().strip()
            
            if citation and page_str:
                try:
                    page_num = int(page_str)
                    linker.set_page_automation(True, citation, page_num)
                except ValueError:
                    messagebox.showerror("Error", "Page number must be a valid integer")
                    return
            else:
                messagebox.showerror("Error", "Please enter both exemplary citation and page number for page automation")
                return
        else:
            linker.set_page_automation(False)
        
        # NOW show progress AFTER all validation!
        self.progress.pack(pady=10)
        self.progress.start()
        
        mode_text = f"Excel {submode.title()} mode"
        self.status_text.set(f"Processing Excel file in {mode_text}...")
        self.root.update()
        
        try:
            links_added = linker.process_excel_column()
            
            self.progress.stop()
            self.progress.pack_forget()
            
            if links_added >= 0:
                excel_saved, pdf_saved = linker.save_excel_with_links()
                
                if excel_saved:
                    link_type = "Bates links" if submode == "bates" else "exhibit links"
                    self.status_text.set(f"Success! {links_added} {link_type} added to Excel file.")
                    
                    success_message = f"Excel file processed successfully!\n\n"
                    success_message += f"• {links_added} relative hyperlinks added\n"
                    success_message += f"• Mode: {mode_text}\n"
                    success_message += f"• Column: {linker.selected_column_letter}\n"
                    success_message += f"• Excel file saved with working links\n"
                    if pdf_saved:
                        success_message += f"• PDF export completed\n"
                    else:
                        success_message += f"• PDF export failed (Excel file still has links)\n"
                    success_message += f"• Links work when files are moved together"
                    
                    messagebox.showinfo("Processing Complete", success_message)

                    # Job complete ASCII art in console
                    print("\n")
                    print("     ██╗ ██████╗ ██████╗      ██████╗ ██████╗ ███╗   ███╗██████╗ ██╗     ███████╗████████╗███████╗██╗")
                    print("     ██║██╔═══██╗██╔══██╗    ██╔════╝██╔═══██╗████╗ ████║██╔══██╗██║     ██╔════╝╚══██╔══╝██╔════╝██║")
                    print("     ██║██║   ██║██████╔╝    ██║     ██║   ██║██╔████╔██║██████╔╝██║     █████╗     ██║   █████╗  ██║")
                    print("██   ██║██║   ██║██╔══██╗    ██║     ██║   ██║██║╚██╔╝██║██╔═══╝ ██║     ██╔══╝     ██║   ██╔══╝  ╚═╝")
                    print("╚█████╔╝╚██████╔╝██████╔╝    ╚██████╗╚██████╔╝██║ ╚═╝ ██║██║     ███████╗███████╗   ██║   ███████╗██╗")
                    print(" ╚════╝  ╚═════╝ ╚═════╝      ╚═════╝ ╚═════╝ ╚═╝     ╚═╝╚═╝     ╚══════╝╚══════╝   ╚═╝   ╚══════╝╚═╝")
                    print(f"Word processing complete! {links_added} links created.\n")

                    # Force cleanup to close Excel and remove working copy
                    try:
                        linker.cleanup()
                    except Exception as e:
                        print(f"Error during final cleanup: {e}")

                    
                else:
                    self.status_text.set("Excel processing failed")
                    messagebox.showerror("Error", "Failed to save Excel file")
            else:
                self.status_text.set("Excel processing completed with errors")
                messagebox.showwarning("Warning", "Processing completed but may have encountered errors.")
                
        except Exception as e:
            self.progress.stop()
            self.progress.pack_forget()
            self.status_text.set(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Excel processing failed: {str(e)}")

    def on_closing(self):
        """Handle application closing"""
        try:
            if self.word_linker:
                self.word_linker.cleanup()
            if self.excel_linker:
                self.excel_linker.cleanup()
        except:
            pass
        self.root.destroy()

    def create_word_controls(self):
        """Create Word-specific controls - NEW METHOD"""
        self.word_controls_frame = ttk.Frame(self.dynamic_controls_frame)
        
        # Word sub-mode selection
        submode_frame = ttk.Frame(self.word_controls_frame)
        submode_frame.pack(fill=X, pady=(0, 10))
        
        ttk.Label(submode_frame, text="Word Mode:", font=("Helvetica", 10, "bold")).pack(side=LEFT, padx=(0, 15))
        
        ttk.Radiobutton(
            submode_frame,
            text="Exhibit Links",
            variable=self.word_submode_var,
            value="exhibit",
            command=self.on_word_submode_changed,
            bootstyle="info"
        ).pack(side=LEFT, padx=(0, 20))
        
        ttk.Radiobutton(
            submode_frame,
            text="Bates Links",
            variable=self.word_submode_var,
            value="bates",
            command=self.on_word_submode_changed,
            bootstyle="info"
        ).pack(side=LEFT)
            
        # Bates prefix for Word (initially hidden)
        self.word_bates_frame = ttk.Frame(self.word_controls_frame)
        
        ttk.Label(self.word_bates_frame, text="Bates Prefix:", font=("Helvetica", 10, "bold")).pack(side=LEFT, padx=(0, 10))
        
        self.word_bates_entry = ttk.Entry(
            self.word_bates_frame,
            textvariable=self.bates_prefix_var,
            width=20
        )
        self.word_bates_entry.pack(side=LEFT, padx=(0, 10))
        
        ttk.Label(
            self.word_bates_frame,
            text="(e.g., SMITH_, DOC_) *CASE SENSITIVE*",
            font=("Helvetica", 9),
            bootstyle="secondary"
        ).pack(side=LEFT)

    def create_widgets(self):
        # Main container with responsive padding
        padding = self.get_responsive_padding(20)
        main_frame = ttk.Frame(self.root, padding=padding)
        main_frame.pack(fill=BOTH, expand=True)
        
        # Enhanced Header - RESPONSIVE VERSION
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=X, pady=(0, 8))
        
        # Configure custom styles
        style = ttk.Style()
        
        # Header styling - responsive
        outer_header = ttk.Frame(header_frame, style="HeaderOuter.TFrame")
        outer_header.pack(fill=X, expand=True, pady=1)
        
        style.configure("HeaderOuter.TFrame", background="#0099FF", relief="flat", borderwidth=1)
        
        middle_header = ttk.Frame(outer_header, style="HeaderMiddle.TFrame")
        header_padding = self.get_responsive_padding(2)
        middle_header.pack(fill=X, expand=True, padx=header_padding, pady=header_padding)
        
        style.configure("HeaderMiddle.TFrame", background="#0099FF", relief="flat")
        
        inner_padding = self.get_responsive_padding(15)
        inner_header = ttk.Frame(middle_header, style="HeaderInner.TFrame")
        inner_header.pack(fill=X, expand=True, padx=inner_padding, pady=inner_padding)
        
        style.configure("HeaderInner.TFrame", background="#0099FF")
        
        # Icon and title container - responsive layout
        title_container = ttk.Frame(inner_header, style="HeaderInner.TFrame")
        title_container.pack(fill=X, expand=True)
        
        # Left side: Icon and title
        left_content = ttk.Frame(title_container, style="HeaderInner.TFrame")
        left_content.pack(side=LEFT)
        
        # Responsive icon size - LESS AGGRESSIVE
        icon_size = 22 if self.is_small_screen else 24  # Minimal difference
        icon_label = ttk.Label(
            left_content,
            text="🔗",
            font=("Segoe UI Emoji", icon_size, "normal"),
            foreground="#FFFFFF",
            background="#0099FF"
        )
        icon_padding = 8 if self.is_small_screen else 10  # Minimal difference
        icon_label.pack(side=LEFT, padx=(0, icon_padding))
        
        # Title text container
        text_container = ttk.Frame(left_content, style="HeaderInner.TFrame")
        text_container.pack(side=LEFT, fill=Y)
        
        # Responsive title font
        title_font_size = self.get_responsive_font_size(20)
        title_label = ttk.Label(
            text_container,
            text="Exhibit Linker ",
            font=("Segoe UI", title_font_size, "bold"),
            foreground="#FFFFFF",
            background="#0099FF"
        )
        title_label.pack(anchor=W)
        
        # Responsive subtitle
        subtitle_font_size = self.get_responsive_font_size(9)
        subtitle_label = ttk.Label(
            text_container,
            text="Word + Excel Hyperlink Automation  ",
            font=("Segoe UI", subtitle_font_size, "normal"),
            foreground="#FFFFFF",
            background="#0099FF"
        )
        subtitle_label.pack(anchor=W, pady=(1, 0))
        
        # Bottom accent bars
        accent_frame = ttk.Frame(main_frame, height=3, style="AccentTop.TFrame")
        accent_frame.pack(fill=X)
        style.configure("AccentTop.TFrame", background="#0099FF")
        
        accent_frame2 = ttk.Frame(main_frame, height=2, style="AccentBottom.TFrame")
        accent_frame2.pack(fill=X, pady=(0, 8))
        style.configure("AccentBottom.TFrame", background="#0099FF")
        
        # MODE SELECTION SECTION - RESPONSIVE
        section_padding = self.get_responsive_padding(15)
        mode_frame = ttk.LabelFrame(main_frame, text="Processing Mode (Exit Word/Excel First)", padding=section_padding)
        mode_frame.pack(fill=X, pady=(0, 15))
        
        # Radio buttons - responsive spacing
        mode_container = ttk.Frame(mode_frame)
        mode_container.pack(fill=X, pady=(0, 10))
        
        radio_spacing = 30 if self.is_small_screen else 50
        
        ttk.Radiobutton(
            mode_container,
            text="Word Document",
            variable=self.processing_mode,
            value="word",
            command=self.on_mode_changed,
            bootstyle="primary"
        ).pack(side=LEFT, padx=(0, radio_spacing))
        
        ttk.Radiobutton(
            mode_container,
            text="Excel File",
            variable=self.processing_mode,
            value="excel",
            command=self.on_mode_changed,
            bootstyle="primary"
        ).pack(side=LEFT)
        
        # Dynamic controls container
        self.dynamic_controls_frame = ttk.Frame(mode_frame)
        self.dynamic_controls_frame.pack(fill=X)
        
        # Create mode-specific UI elements
        self.create_word_controls()
        self.create_excel_controls()
        
        # RESPONSIVE LAYOUT - Steps 1 and 2 side by side
        steps_container = ttk.Frame(main_frame)
        steps_container.pack(fill=X, pady=(0, 15))
        
        # Configure grid weights
        steps_container.grid_columnconfigure(0, weight=1)
        steps_container.grid_columnconfigure(1, weight=2)  # Step 2 gets more space
        
        # Step 1: Document Selection (left side)
        self.step1_frame = ttk.LabelFrame(steps_container, text="Step 1: Select Document", padding=section_padding)
        self.step1_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        doc_main_frame = ttk.Frame(self.step1_frame)
        doc_main_frame.pack(fill=X, pady=(0, 10))

        if self.is_small_screen:
            # SMALL SCREEN: Stack vertically instead of side-by-side
            print("Using small-screen vertical layout")
            
            # Document selection (full width)
            doc_left_frame = ttk.Frame(doc_main_frame)
            doc_left_frame.pack(fill=X, pady=(0, 10))
            
            self.doc_label_text = ttk.Label(doc_left_frame, text="Selected Document:", 
                                        font=("Helvetica", self.get_responsive_font_size(10), "bold"))
            self.doc_label_text.pack(anchor=W)

            doc_label = ttk.Label(
                doc_left_frame, 
                textvariable=self.doc_path, 
                font=("Helvetica", self.get_responsive_font_size(9)),
                bootstyle="info",
                wraplength=self.get_responsive_wraplength(400)
            )
            doc_label.pack(anchor=W, pady=(2, 8))

            self.browse_doc_button = ttk.Button(
                doc_left_frame,
                text="Browse Document",
                command=self.browse_document,
                bootstyle="primary-outline",
                width=20
            )
            self.browse_doc_button.pack(anchor=W)
            
            # Excel column selection (full width, separate section)
            self.excel_column_frame = ttk.Frame(doc_main_frame)
            # Will be shown/hidden by on_mode_changed - no grid needed for small screen
            
            ttk.Label(self.excel_column_frame, text="Selected Column:", 
                    font=("Helvetica", self.get_responsive_font_size(10), "bold")).pack(anchor=W)

            column_info_label = ttk.Label(
                self.excel_column_frame,
                textvariable=self.selected_column_var,
                font=("Helvetica", self.get_responsive_font_size(9)),
                bootstyle="secondary",
                wraplength=self.get_responsive_wraplength(250)
            )
            column_info_label.pack(anchor=W, pady=(2, 8))

            self.select_column_button = ttk.Button(
                self.excel_column_frame,
                text="Select Column",
                command=self.select_excel_column,
                bootstyle="info-outline",
                width=15,
                state='disabled'
            )
            self.select_column_button.pack(anchor=W)
            
            # No separator needed for vertical layout
            self.excel_separator_frame = None
            
        else:
            # LARGE SCREEN: Use original grid layout
            print("Using large-screen grid layout")
            
            # Configure grid columns
            doc_main_frame.grid_columnconfigure(0, weight=1)
            doc_main_frame.grid_columnconfigure(1, weight=0)
            doc_main_frame.grid_columnconfigure(2, weight=1)

            # Left side: Document info and browse button
            doc_left_frame = ttk.Frame(doc_main_frame)
            doc_left_frame.grid(row=0, column=0, sticky="nw", padx=(0, 10))

            self.doc_label_text = ttk.Label(doc_left_frame, text="Selected Document:", 
                                        font=("Helvetica", 10, "bold"))
            self.doc_label_text.pack(anchor=W)

            doc_label = ttk.Label(
                doc_left_frame, 
                textvariable=self.doc_path, 
                font=("Helvetica", 9),
                bootstyle="info",
                wraplength=400
            )
            doc_label.pack(anchor=W, pady=(2, 8))

            self.browse_doc_button = ttk.Button(
                doc_left_frame,
                text="Browse Document",
                command=self.browse_document,
                bootstyle="primary-outline",
                width=20
            )
            self.browse_doc_button.pack(anchor=W)

            # Separator
            self.excel_separator_frame = ttk.Frame(doc_main_frame)
            separator = ttk.Separator(self.excel_separator_frame, orient='vertical')
            separator.pack(fill=Y, expand=True)

            # Right side: Excel column selection
            self.excel_column_frame = ttk.Frame(doc_main_frame)

            ttk.Label(self.excel_column_frame, text="Selected Column:", font=("Helvetica", 10, "bold")).pack(anchor=W)

            column_info_label = ttk.Label(
                self.excel_column_frame,
                textvariable=self.selected_column_var,
                font=("Helvetica", 9),
                bootstyle="secondary",
                wraplength=250
            )
            column_info_label.pack(anchor=W, pady=(2, 8))

            self.select_column_button = ttk.Button(
                self.excel_column_frame,
                text="Select Column",
                command=self.select_excel_column,
                bootstyle="info-outline",
                width=15,
                state='disabled'
            )
            self.select_column_button.pack(anchor=W)

        # Update UI based on initial mode
        self.on_mode_changed()

        # Step 2: Folder Selection - RESPONSIVE
        step2_frame = ttk.LabelFrame(steps_container, text="Step 2: Select Linked Files Folder", padding=section_padding)
        step2_frame.grid(row=0, column=1, sticky="nsew")
        
        folder_info_frame = ttk.Frame(step2_frame)
        folder_info_frame.pack(fill=X, pady=(0, 10))
        
        ttk.Label(folder_info_frame, text="Linked Files Folder:", 
                font=("Helvetica", self.get_responsive_font_size(10), "bold")).pack(anchor=W)
        
        folder_label = ttk.Label(
            folder_info_frame, 
            textvariable=self.folder_path, 
            font=("Helvetica", self.get_responsive_font_size(9)),
            bootstyle="secondary",
            wraplength=self.get_responsive_wraplength(700)
        )
        folder_label.pack(anchor=W, pady=(2, 0))
        
        folder_button_frame = ttk.Frame(step2_frame)
        folder_button_frame.pack(fill=X)
        
        # Responsive button sizing and spacing - LESS AGGRESSIVE
        button_width = 19 if self.is_small_screen else 20  # Minimal difference
        button_spacing = 8 if self.is_small_screen else 10  # Minimal difference
        
        ttk.Button(
            folder_button_frame,
            text="Browse Files Folder",
            command=self.browse_folder,
            bootstyle="secondary-outline",
            width=button_width
        ).pack(side=LEFT, padx=(0, button_spacing))
        
        ttk.Button(
            folder_button_frame,
            text="Use Step 1 Folder",
            command=self.use_document_folder,
            bootstyle="info-outline",
            width=button_width
        ).pack(side=LEFT, padx=(0, button_spacing))
        
        ttk.Button(
            folder_button_frame,
            text="Process Filenames",
            command=self.show_file_renamer_dialog,
            bootstyle="warning-outline",
            width=button_width + 2
        ).pack(side=LEFT)
        
        # Step 3: Process - RESPONSIVE
        step3_frame = ttk.LabelFrame(main_frame, text="Step 3: Process Document", padding=section_padding)
        step3_frame.pack(fill=X, pady=(0, 15))
        
        status_process_frame = ttk.Frame(step3_frame)
        status_process_frame.pack(fill=X)
        
        if self.is_small_screen:
            # SMALL SCREEN: Stack controls vertically
            left_controls_frame = ttk.Frame(status_process_frame)
            left_controls_frame.pack(fill=X, pady=(0, 10))

            process_btn = ttk.Button(
                left_controls_frame,
                text="Process Document", 
                command=self.process_document,
                bootstyle="success",
                width=20
            )
            process_btn.pack(pady=(0, 5))

            black_links_check = ttk.Checkbutton(
                left_controls_frame,
                text="Hidden Hyperlinks (Black/No Underline)",
                variable=self.use_black_hyperlinks,
                bootstyle="info-round-toggle"
            )
            black_links_check.pack(anchor='w', pady=(2, 0))
            
            # Page automation toggle
            self.page_auto_check = ttk.Checkbutton(
                left_controls_frame,
                text="Automate Page Number Links",
                variable=self.page_automation_var,
                command=self.on_page_automation_toggled,
                bootstyle="warning-round-toggle"
            )
            self.page_auto_check.pack(anchor='w', pady=(2, 0))
            
            # Page automation controls (always visible but conditionally enabled)
            self.page_automation_frame = ttk.Frame(left_controls_frame)
            self.page_automation_frame.pack(fill=X, pady=(5, 0))  # Always pack it
            
            # Exemplary citation
            citation_frame = ttk.Frame(self.page_automation_frame)
            citation_frame.pack(fill=X, pady=(5, 2))
            
            ttk.Label(citation_frame, text="Example Citation:", font=("Helvetica", 9, "bold")).pack(anchor=W)
            self.citation_entry = ttk.Entry(
                citation_frame,
                textvariable=self.exemplary_citation_var,
                width=30 if not self.is_small_screen else 25,
                font=("Helvetica", 8)
            )
            self.citation_entry.pack(fill=X, pady=(2, 0))
            
            # Exemplary page number
            page_frame = ttk.Frame(self.page_automation_frame)
            page_frame.pack(fill=X, pady=(2, 0))
            
            ttk.Label(page_frame, text="Page Number:", font=("Helvetica", 9, "bold")).pack(side=LEFT)
            self.page_entry = ttk.Entry(
                page_frame,
                textvariable=self.exemplary_page_var,
                width=8,
                font=("Helvetica", 8)
            )
            self.page_entry.pack(side=LEFT, padx=(5, 0))
            
            ttk.Label(
                page_frame,
                text="(e.g., 'Ex. 5, Memo, at p. 25' and '25')",
                font=("Helvetica", 8),
                bootstyle="secondary"
            ).pack(side=LEFT, padx=(10, 0))
            
            # Status below
            status_right_frame = ttk.Frame(status_process_frame)
            status_right_frame.pack(fill=X)
            
        else:
            # LARGE SCREEN: Horizontal layout
            left_controls_frame = ttk.Frame(status_process_frame)
            left_controls_frame.pack(side=LEFT, padx=(0, 20))

            process_btn = ttk.Button(
                left_controls_frame,
                text="Process Document", 
                command=self.process_document,
                bootstyle="success",
                width=25
            )
            process_btn.pack(anchor=W, pady=(0, 5))

            black_links_check = ttk.Checkbutton(
                left_controls_frame,
                text="Hidden Hyperlinks (Black/No Underline)",
                variable=self.use_black_hyperlinks,
                bootstyle="info-round-toggle"
            )
            black_links_check.pack(anchor='w', pady=(2, 0))
            
            # Page automation toggle
            self.page_auto_check = ttk.Checkbutton(
                left_controls_frame,
                text="Automate Page Number Links",
                variable=self.page_automation_var,
                command=self.on_page_automation_toggled,
                bootstyle="warning-round-toggle"
            )
            self.page_auto_check.pack(anchor='w', pady=(2, 0))
            
            # Page automation controls (always visible but conditionally enabled)
            self.page_automation_frame = ttk.Frame(left_controls_frame)
            self.page_automation_frame.pack(fill=X, pady=(5, 0))  # Always pack it
            
            # Exemplary citation
            citation_frame = ttk.Frame(self.page_automation_frame)
            citation_frame.pack(fill=X, pady=(5, 2))
            
            ttk.Label(citation_frame, text="Example Citation:", font=("Helvetica", 9, "bold")).pack(anchor=W)
            self.citation_entry = ttk.Entry(
                citation_frame,
                textvariable=self.exemplary_citation_var,
                width=30,
                font=("Helvetica", 8)
            )
            self.citation_entry.pack(fill=X, pady=(2, 0))
            
            # Exemplary page number 
            page_frame = ttk.Frame(self.page_automation_frame)
            page_frame.pack(fill=X, pady=(2, 0))
            
            ttk.Label(page_frame, text="Page Number:", font=("Helvetica", 9, "bold")).pack(side=LEFT)
            self.page_entry = ttk.Entry(
                page_frame,
                textvariable=self.exemplary_page_var,
                width=8,
                font=("Helvetica", 8)
            )
            self.page_entry.pack(side=LEFT, padx=(5, 0))
            
            ttk.Label(
                page_frame,
                text="(e.g., 'Ex. 5, Memo, at p. 25' and '25')",
                font=("Helvetica", 8),
                bootstyle="secondary"
            ).pack(side=LEFT, padx=(10, 0))
            
            # Status on right
            status_right_frame = ttk.Frame(status_process_frame)
            status_right_frame.pack(side=LEFT, fill=X, expand=True)
        
        ttk.Label(status_right_frame, text="Status:", 
                font=("Helvetica", self.get_responsive_font_size(10), "bold")).pack(anchor=W)
        status_label = ttk.Label(
            status_right_frame, 
            textvariable=self.status_text, 
            font=("Helvetica", self.get_responsive_font_size(9)),
            bootstyle="secondary"
        )
        status_label.pack(anchor=W, pady=(2, 0))
        
        # Progress bar
        self.progress = ttk.Progressbar(
            step3_frame,
            mode='indeterminate',
            bootstyle="success-striped"
        )
        
        # Information Panel - RESPONSIVE
        info_padding = self.get_responsive_padding(12)
        info_frame = ttk.LabelFrame(main_frame, text="Information", padding=info_padding)
        info_frame.pack(fill=BOTH, expand=True, pady=(0, 12))
        
        self.info_text_var = tk.StringVar()
        self.update_info_text()
        
        info_font_size = self.get_responsive_font_size(9)
        self.info_label = ttk.Label(
                    info_frame, 
                    textvariable=self.info_text_var,
                    justify=LEFT, 
                    anchor='w',  # Anchor to west (left) side
                    wraplength=self.get_responsive_wraplength(700),
                    font=("Helvetica", info_font_size)
                )
        info_label_padding = self.get_responsive_padding(8)
        self.info_label.pack(anchor='w', fill='x', pady=info_label_padding, padx=info_label_padding)
        
        # Help button
        help_button = ttk.Button(
            main_frame,
            text="?",
            command=self.show_help_popup,
            bootstyle="info",
            width=3
        )
        help_button.pack(side=RIGHT, anchor=SE, padx=(0, 5), pady=(0, 5))


    def create_excel_controls(self):
        """Create Excel-specific controls"""
        self.excel_controls_frame = ttk.Frame(self.dynamic_controls_frame)
        
        # Excel sub-mode selection
        submode_frame = ttk.Frame(self.excel_controls_frame)
        submode_frame.pack(fill=X, pady=(0, 10))
        
        ttk.Label(submode_frame, text="Excel Mode:", font=("Helvetica", 10, "bold")).pack(side=LEFT, padx=(0, 15))
        
        ttk.Radiobutton(
            submode_frame,
            text="Exhibit Links",
            variable=self.excel_submode_var,
            value="exhibit",
            command=self.on_excel_submode_changed,
            bootstyle="info"
        ).pack(side=LEFT, padx=(0, 20))
        
        ttk.Radiobutton(
            submode_frame,
            text="Bates Links",
            variable=self.excel_submode_var,
            value="bates",
            command=self.on_excel_submode_changed,
            bootstyle="info"
        ).pack(side=LEFT)
               
        # Bates prefix for Excel (initially hidden)
        self.excel_bates_frame = ttk.Frame(self.excel_controls_frame)
        
        ttk.Label(self.excel_bates_frame, text="Bates Prefix:", font=("Helvetica", 10, "bold")).pack(side=LEFT, padx=(0, 10))
        
        self.excel_bates_entry = ttk.Entry(
            self.excel_bates_frame,
            textvariable=self.bates_prefix_var,
            width=20
        )
        self.excel_bates_entry.pack(side=LEFT, padx=(0, 10))
        
        ttk.Label(
            self.excel_bates_frame,
            text="(e.g., SMITH_, DOC_) *CASE SENSITIVE*",
            font=("Helvetica", 9),
            bootstyle="secondary"
        ).pack(side=LEFT)

def main():
    """Main function"""
    try:
        # Terminal welcome message with ASCII art
        print("\nWelcome to")
        print("""███████╗██╗  ██╗██╗  ██╗██╗██████╗ ██╗████████╗    ██╗     ██╗███╗   ██╗██╗  ██╗███████╗██████╗ 
██╔════╝╚██╗██╔╝██║  ██║██║██╔══██╗██║╚══██╔══╝    ██║     ██║████╗  ██║██║ ██╔╝██╔════╝██╔══██╗
█████╗   ╚███╔╝ ███████║██║██████╔╝██║   ██║       ██║     ██║██╔██╗ ██║█████╔╝ █████╗  ██████╔╝
██╔══╝   ██╔██╗ ██╔══██║██║██╔══██╗██║   ██║       ██║     ██║██║╚██╗██║██╔═██╗ ██╔══╝  ██╔══██╗
███████╗██╔╝ ██╗██║  ██║██║██████╔╝██║   ██║       ███████╗██║██║ ╚████║██║  ██╗███████╗██║  ██║
╚══════╝╚═╝  ╚═╝╚═╝  ╚═╝╚═╝╚═════╝ ╚═╝   ╚═╝       ╚══════╝╚═╝╚═╝  ╚═══╝╚═╝  ╚═╝╚══════╝╚═╝  ╚═╝""")
        print("Word + Excel Hyperlink Automation")
        print("Copyright © Alexander Owens, 2025\n")
        
        app = ExhibitAnchorApp()
        app.root.mainloop()
    except Exception as e:
        messagebox.showerror("Startup Error", f"Could not start application: {str(e)}")
        
if __name__ == "__main__":
    main()

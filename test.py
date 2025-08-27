import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pygetwindow as gw
import cv2
import numpy as np
import pytesseract
from PIL import Image, ImageTk
import win32gui
import win32ui
import win32con
import win32api
import threading
import re

class WindowTextFinder:
    def __init__(self, root):
        self.root = root
        self.root.title("Window Text Finder")
        self.root.geometry("800x600")
        
        # Variables
        self.selected_window = None
        self.windows_list = []
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Windows list section
        ttk.Label(main_frame, text="Available Windows:").grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        # Listbox for windows
        self.windows_listbox = tk.Listbox(main_frame, height=8)
        self.windows_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        self.windows_listbox.bind('<<ListboxSelect>>', self.on_window_select)
        
        # Refresh button
        ttk.Button(main_frame, text="Refresh Windows", command=self.refresh_windows).grid(row=2, column=0, sticky=tk.W, pady=(0, 10))
        
        # Action buttons frame
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.experience_btn = ttk.Button(action_frame, text="Experience", command=self.experience_mode, state='disabled')
        self.experience_btn.grid(row=0, column=0, padx=(0, 10))
        
        self.search_destroy_btn = ttk.Button(action_frame, text="Search and Destroy", command=self.search_destroy_mode, state='disabled')
        self.search_destroy_btn.grid(row=0, column=1)
        
        # Search frame (initially hidden)
        self.search_frame = ttk.LabelFrame(main_frame, text="Search and Destroy", padding="10")
        
        ttk.Label(self.search_frame, text="Enter text to search for:").grid(row=0, column=0, sticky=tk.W)
        
        self.search_entry = ttk.Entry(self.search_frame, width=50)
        self.search_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 10))
        
        ttk.Button(self.search_frame, text="Find Text", command=self.find_text).grid(row=2, column=0, sticky=tk.W)
        
        # Results area
        ttk.Label(self.search_frame, text="Results:").grid(row=3, column=0, sticky=tk.W, pady=(10, 5))
        
        self.results_text = scrolledtext.ScrolledText(self.search_frame, height=10, width=70)
        self.results_text.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Configure search frame grid
        self.search_frame.columnconfigure(0, weight=1)
        self.search_frame.rowconfigure(4, weight=1)
        
        # Progress bar
        self.progress = ttk.Progressbar(self.search_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Initialize windows list
        self.refresh_windows()
    
    def refresh_windows(self):
        """Refresh the list of available windows"""
        self.windows_listbox.delete(0, tk.END)
        self.windows_list = []
        
        try:
            # Get all windows
            windows = gw.getAllWindows()
            
            for window in windows:
                # Filter out empty titles and system windows
                if window.title and window.title.strip() and window.visible:
                    # Avoid including the current application
                    if "Window Text Finder" not in window.title:
                        self.windows_list.append(window)
                        self.windows_listbox.insert(tk.END, f"{window.title}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to get windows list: {str(e)}")
    
    def on_window_select(self, event):
        """Handle window selection"""
        selection = event.widget.curselection()
        if selection:
            index = selection[0]
            self.selected_window = self.windows_list[index]
            self.experience_btn.config(state='normal')
            self.search_destroy_btn.config(state='normal')
    
    def experience_mode(self):
        """Experience mode - currently blank as requested"""
        if not self.selected_window:
            messagebox.showwarning("Warning", "Please select a window first")
            return
        
        # Hide search frame if visible
        self.search_frame.grid_remove()
        
        messagebox.showinfo("Experience Mode", f"Experience mode for window: {self.selected_window.title}\n\nThis feature is not yet implemented.")
    
    def search_destroy_mode(self):
        """Search and destroy mode - show search interface"""
        if not self.selected_window:
            messagebox.showwarning("Warning", "Please select a window first")
            return
        
        # Show search frame
        self.search_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # Configure main frame to accommodate search frame
        self.root.geometry("800x800")
    
    def capture_window(self, window):
        """Capture screenshot of the selected window"""
        try:
            # Get window handle
            hwnd = window._hWnd
            
            # Get window dimensions
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top
            
            # Create device context
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()
            
            # Create bitmap
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)
            
            # Copy window content
            result = saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)
            
            # Convert to numpy array
            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)
            
            img = np.frombuffer(bmpstr, dtype='uint8')
            img.shape = (height, width, 4)
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2RGB)
            
            # Cleanup
            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwndDC)
            
            return img
            
        except Exception as e:
            raise Exception(f"Failed to capture window: {str(e)}")
    
    def find_text_in_image(self, image, search_text):
        """Find text in image using OCR and return bounding boxes"""
        try:
            # Convert to PIL Image for tesseract
            pil_image = Image.fromarray(image)
            
            # Use pytesseract to get detailed data
            data = pytesseract.image_to_data(pil_image, output_type=pytesseract.Output.DICT)
            
            # Find matching text
            results = []
            search_text_lower = search_text.lower().strip()
            
            # Combine words to check for multi-word phrases
            words = []
            for i in range(len(data['text'])):
                if int(data['conf'][i]) > 30:  # Filter low confidence detections
                    words.append({
                        'text': data['text'][i].strip(),
                        'left': data['left'][i],
                        'top': data['top'][i],
                        'width': data['width'][i],
                        'height': data['height'][i],
                        'conf': data['conf'][i]
                    })
            
            # Search for exact matches and partial matches
            for i, word in enumerate(words):
                if not word['text']:
                    continue
                
                # Check single word match
                if search_text_lower in word['text'].lower():
                    x1, y1 = word['left'], word['top']
                    x2, y2 = word['left'] + word['width'], word['top'] + word['height']
                    results.append({
                        'text': word['text'],
                        'coordinates': [x1, y1, x2, y2],
                        'confidence': word['conf']
                    })
                
                # Check multi-word phrases
                if ' ' in search_text_lower:
                    phrase_words = search_text_lower.split()
                    if len(phrase_words) > 1:
                        # Look for consecutive words that match the phrase
                        for j in range(len(phrase_words)):
                            if i + j < len(words):
                                combined_text = ' '.join([w['text'].lower() for w in words[i:i+len(phrase_words)]])
                                if search_text_lower in combined_text:
                                    # Calculate bounding box for the entire phrase
                                    first_word = words[i]
                                    last_word = words[i + len(phrase_words) - 1]
                                    x1 = first_word['left']
                                    y1 = min([w['top'] for w in words[i:i+len(phrase_words)]])
                                    x2 = last_word['left'] + last_word['width']
                                    y2 = max([w['top'] + w['height'] for w in words[i:i+len(phrase_words)]])
                                    
                                    results.append({
                                        'text': combined_text,
                                        'coordinates': [x1, y1, x2, y2],
                                        'confidence': min([w['conf'] for w in words[i:i+len(phrase_words)]])
                                    })
                                    break
            
            return results
            
        except Exception as e:
            raise Exception(f"OCR processing failed: {str(e)}")
    
    def find_text_thread(self, search_text):
        """Thread function for finding text"""
        try:
            self.progress.start()
            
            # Capture window
            self.results_text.insert(tk.END, f"Capturing window: {self.selected_window.title}...\n")
            self.results_text.see(tk.END)
            self.root.update()
            
            image = self.capture_window(self.selected_window)
            
            self.results_text.insert(tk.END, f"Searching for text: '{search_text}'...\n")
            self.results_text.see(tk.END)
            self.root.update()
            
            # Find text in image
            results = self.find_text_in_image(image, search_text)
            
            # Display results
            self.results_text.insert(tk.END, f"\n--- SEARCH RESULTS ---\n")
            if results:
                self.results_text.insert(tk.END, f"Found {len(results)} match(es):\n\n")
                for i, result in enumerate(results, 1):
                    x1, y1, x2, y2 = result['coordinates']
                    self.results_text.insert(tk.END, f"Match {i}:\n")
                    self.results_text.insert(tk.END, f"  Text: '{result['text']}'\n")
                    self.results_text.insert(tk.END, f"  Coordinates: [{x1}, {y1}, {x2}, {y2}]\n")
                    self.results_text.insert(tk.END, f"  Confidence: {result['confidence']:.1f}%\n\n")
            else:
                self.results_text.insert(tk.END, "No matches found.\n")
            
            self.results_text.insert(tk.END, "--- SEARCH COMPLETE ---\n\n")
            self.results_text.see(tk.END)
            
        except Exception as e:
            self.results_text.insert(tk.END, f"Error: {str(e)}\n")
        finally:
            self.progress.stop()
    
    def find_text(self):
        """Find text in the selected window"""
        if not self.selected_window:
            messagebox.showwarning("Warning", "Please select a window first")
            return
        
        search_text = self.search_entry.get().strip()
        if not search_text:
            messagebox.showwarning("Warning", "Please enter text to search for")
            return
        
        # Clear previous results
        self.results_text.delete(1.0, tk.END)
        
        # Run search in a separate thread to prevent UI freezing
        thread = threading.Thread(target=self.find_text_thread, args=(search_text,))
        thread.daemon = True
        thread.start()

def main():
    # Check if required dependencies are available
    try:
        import pytesseract
        import cv2
        import pygetwindow
        import win32gui
    except ImportError as e:
        print(f"Missing required dependency: {e}")
        print("\nPlease install the required packages:")
        print("pip install opencv-python pytesseract pygetwindow pywin32 pillow")
        print("\nAlso make sure Tesseract is installed and in your PATH:")
        print("https://github.com/UB-Mannheim/tesseract/wiki")
        return
    
    root = tk.Tk()
    app = WindowTextFinder(root)
    root.mainloop()

if __name__ == "__main__":
    main()
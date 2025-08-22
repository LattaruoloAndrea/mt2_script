import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import keyboard
from typing import Dict, List
import json
import os

class AutoKeyPresser:
    def __init__(self, root):
        self.root = root
        self.root.title("Auto Key Presser")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # Application state
        self.is_running = False
        self.is_starting = False
        self.countdown_active = False
        self.spacebar_active = False
        self.threads: Dict[str, threading.Thread] = {}
        self.stop_events: Dict[str, threading.Event] = {}
        self.key_sequences: List[Dict[str, str]] = []
        self.spacebar_thread = None
        
        self.setup_ui()
        self.load_config()
        self.setup_global_hotkey()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Auto Key Presser", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Control buttons frame
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=1, column=0, columnspan=3, pady=(0, 10), sticky=(tk.W, tk.E))
        
        # Start/Stop button
        self.start_stop_btn = ttk.Button(control_frame, text="Start", command=self.handle_button_click)
        self.start_stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Spacebar hold switch button
        # self.spacebar_btn = ttk.Button(control_frame, text="Spacebar: OFF", command=self.toggle_spacebar_button)
        # self.spacebar_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(control_frame, text="Status: Stopped", foreground="red")
        self.status_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # Key sequences frame
        sequences_frame = ttk.LabelFrame(main_frame, text="Key Sequences", padding="10")
        sequences_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        sequences_frame.columnconfigure(0, weight=1)
        sequences_frame.rowconfigure(0, weight=1)
        
        # Treeview for displaying key sequences
        columns = ("Key", "Interval (seconds)")
        self.tree = ttk.Treeview(sequences_frame, columns=columns, show="headings", height=10)
        
        # Define column headings and widths
        self.tree.heading("Key", text="Key")
        self.tree.heading("Interval (seconds)", text="Interval (seconds)")
        self.tree.column("Key", width=200)
        self.tree.column("Interval (seconds)", width=150)
        
        # Scrollbar for treeview
        scrollbar = ttk.Scrollbar(sequences_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Add/Remove buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=3, column=0, columnspan=3, pady=(0, 10))
        
        # ttk.Button(buttons_frame, text="Add Key Sequence", command=self.add_key_sequence).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(buttons_frame, text="Remove Selected", command=self.remove_key_sequence).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(buttons_frame, text="Edit Selected", command=self.edit_key_sequence).pack(side=tk.LEFT)
        
        # Input frame for adding sequences
        input_frame = ttk.LabelFrame(main_frame, text="Add New Key Sequence", padding="10")
        input_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Key input
        ttk.Label(input_frame, text="Key:").grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        self.key_var = tk.StringVar()
        key_entry = ttk.Entry(input_frame, textvariable=self.key_var, width=20)
        key_entry.grid(row=0, column=1, padx=(0, 20), sticky=(tk.W, tk.E))
        
        # Interval input
        ttk.Label(input_frame, text="Interval (seconds):").grid(row=0, column=2, padx=(0, 10), sticky=tk.W)
        self.interval_var = tk.StringVar()
        interval_entry = ttk.Entry(input_frame, textvariable=self.interval_var, width=10)
        interval_entry.grid(row=0, column=3, padx=(0, 20), sticky=(tk.W, tk.E))
        
        # Add button
        ttk.Button(input_frame, text="Add", command=self.add_sequence_from_input).grid(row=0, column=4)
        
        # Configure input frame grid weights
        input_frame.columnconfigure(1, weight=1)
        input_frame.columnconfigure(3, weight=1)
        
    def add_key_sequence(self):
        """Open dialog to add new key sequence"""
        self.show_key_sequence_dialog()
        
    def show_key_sequence_dialog(self, edit_item=None):
        """Show dialog for adding or editing key sequence"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Key Sequence" if not edit_item else "Edit Key Sequence")
        dialog.geometry("300x150")
        dialog.resizable(False, False)
        dialog.grab_set()  # Make dialog modal
        
        # Center dialog on parent window
        dialog.transient(self.root)
        
        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Key input
        ttk.Label(frame, text="Key:").grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        key_var = tk.StringVar()
        if edit_item:
            key_var.set(self.tree.item(edit_item, 'values')[0])
        key_entry = ttk.Entry(frame, textvariable=key_var, width=25)
        key_entry.grid(row=0, column=1, pady=(0, 10), sticky=(tk.W, tk.E))
        key_entry.focus()
        
        # Interval input
        ttk.Label(frame, text="Interval (seconds):").grid(row=1, column=0, sticky=tk.W, pady=(0, 10))
        interval_var = tk.StringVar()
        if edit_item:
            interval_var.set(self.tree.item(edit_item, 'values')[1])
        interval_entry = ttk.Entry(frame, textvariable=interval_var, width=25)
        interval_entry.grid(row=1, column=1, pady=(0, 10), sticky=(tk.W, tk.E))
        
        # Buttons frame
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=(10, 0))
        
        def save_sequence():
            key = key_var.get().strip()
            interval_str = interval_var.get().strip()
            
            if not key or not interval_str:
                messagebox.showerror("Error", "Please fill in both fields")
                return
                
            try:
                interval = float(interval_str)
                if interval <= 0:
                    raise ValueError("Interval must be positive")
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid positive number for interval")
                return
            
            if edit_item:
                # Update existing item
                self.tree.item(edit_item, values=(key, interval_str))
                # Update the key_sequences list
                for i, seq in enumerate(self.key_sequences):
                    if seq['id'] == self.tree.item(edit_item, 'tags')[0]:
                        self.key_sequences[i] = {'id': seq['id'], 'key': key, 'interval': interval}
                        break
            else:
                # Add new item
                item_id = str(len(self.key_sequences))
                item = self.tree.insert("", tk.END, values=(key, interval_str), tags=(item_id,))
                self.key_sequences.append({'id': item_id, 'key': key, 'interval': interval})
            
            self.save_config()
            dialog.destroy()
        
        def cancel():
            dialog.destroy()
        
        ttk.Button(btn_frame, text="Save", command=save_sequence).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="Cancel", command=cancel).pack(side=tk.LEFT)
        
        frame.columnconfigure(1, weight=1)
        
        # Bind Enter key to save
        dialog.bind('<Return>', lambda e: save_sequence())
        dialog.bind('<Escape>', lambda e: cancel())
    
    def handle_button_click(self):
        """Handle start/stop/cancel button clicks"""
        if self.is_starting and self.countdown_active:
            # Cancel the countdown
            self.cancel_start()
        elif not self.is_running and not self.is_starting:
            # Start with delay
            self.start_application_with_delay()
        elif self.is_running:
            # Stop the application
            self.stop_application()
        
    def add_sequence_from_input(self):
        """Add sequence from the input fields at the bottom"""
        key = self.key_var.get().strip()
        interval_str = self.interval_var.get().strip()
        
        if not key or not interval_str:
            messagebox.showerror("Error", "Please fill in both fields")
            return
            
        try:
            interval = float(interval_str)
            if interval <= 0:
                raise ValueError("Interval must be positive")
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid positive number for interval")
            return
        
        # Add to treeview and list
        item_id = str(len(self.key_sequences))
        self.tree.insert("", tk.END, values=(key, interval_str), tags=(item_id,))
        self.key_sequences.append({'id': item_id, 'key': key, 'interval': interval})
        
        # Clear input fields
        self.key_var.set("")
        self.interval_var.set("")
        
        self.save_config()
        
    def remove_key_sequence(self):
        """Remove selected key sequence"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select a sequence to remove")
            return
        
        if messagebox.askyesno("Confirm", "Are you sure you want to remove the selected sequence(s)?"):
            for item in selected_items:
                # Remove from key_sequences list
                item_id = self.tree.item(item, 'tags')[0]
                self.key_sequences = [seq for seq in self.key_sequences if seq['id'] != item_id]
                # Remove from treeview
                self.tree.delete(item)
            
            self.save_config()
            
    def edit_key_sequence(self):
        """Edit selected key sequence"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select a sequence to edit")
            return
        
        if len(selected_items) > 1:
            messagebox.showwarning("Warning", "Please select only one sequence to edit")
            return
            
        self.show_key_sequence_dialog(selected_items[0])
        
    def toggle_application(self):
        """Start or stop the key pressing application"""
        if not self.is_running and not self.is_starting:
            self.start_application_with_delay()
        elif self.is_running:
            self.stop_application()
    
    def start_application_with_delay(self):
        """Start the application with a 5-second countdown"""
        if not self.key_sequences:
            messagebox.showwarning("Warning", "Please add at least one key sequence before starting")
            return
        
        self.is_starting = True
        self.countdown_active = True
        self.start_stop_btn.config(text="Cancel", state="normal")
        
        # Start countdown in a separate thread
        countdown_thread = threading.Thread(target=self.countdown_worker, daemon=True)
        countdown_thread.start()
    
    def countdown_worker(self):
        """Worker function for the 5-second countdown"""
        for i in range(5, 0, -1):
            if not self.countdown_active:  # If cancelled
                return
            
            # Update UI on main thread
            self.root.after(0, lambda count=i: self.update_countdown_ui(count))
            time.sleep(1)
        
        # If countdown completed without cancellation
        if self.countdown_active:
            self.root.after(0, self.start_application_after_countdown)
    
    def update_countdown_ui(self, count):
        """Update UI during countdown"""
        self.status_label.config(text=f"Starting in {count} seconds...", foreground="orange")
    
    def start_application_after_countdown(self):
        """Start the actual application after countdown completes"""
        if not self.countdown_active:  # If cancelled during countdown
            return
        
        self.countdown_active = False
        self.is_starting = False
        self.start_application()
    
    def cancel_start(self):
        """Cancel the startup countdown"""
        self.countdown_active = False
        self.is_starting = False
        self.start_stop_btn.config(text="Start")
        self.status_label.config(text="Status: Stopped", foreground="red")
    
    def toggle_spacebar_button(self):
        """Toggle spacebar holding on/off with button"""
        print("Toggling spacebar hold", self.spacebar_active)
        self.spacebar_active = not self.spacebar_active
        if self.spacebar_active:
            self.spacebar_btn.config(text="Spacebar: ON")
        else:
            self.spacebar_btn.config(text="Spacebar: OFF")
    
    def start_spacebar_hold(self):
        """Start holding the spacebar"""
        if not self.spacebar_active:
            self.spacebar_active = True
            self.spacebar_btn.config(text="Spacebar: ON")
            
            # Start spacebar holding thread
            self.spacebar_thread = threading.Thread(target=self.spacebar_hold_worker, daemon=True)
            self.spacebar_thread.start()
    
    def stop_spacebar_hold(self):
        """Stop holding the spacebar"""
        if self.spacebar_active:
            self.spacebar_active = False
            self.spacebar_btn.config(text="Spacebar: OFF")
            
            # Release the spacebar key
            try:
                keyboard.release('space')
            except Exception as e:
                print(f"Error releasing spacebar: {e}")
    
    def update_spacebar_checkbox(self):
        """Update button text to match actual spacebar state"""
        if self.spacebar_active:
            self.spacebar_btn.config(text="Spacebar: ON")
        else:
            self.spacebar_btn.config(text="Spacebar: OFF")
    
    def spacebar_hold_worker(self):
        """Worker function that keeps spacebar pressed"""
        try:
            # Press and keep spacebar held down
            print("Holding spacebar down")
            keyboard.press('space')
            
            # Keep checking if we should stop
            while self.spacebar_active and self.spacebar_btn['text'] == "Spacebar: ON":
                time.sleep(0.1)  # Small delay to prevent excessive CPU usage
                
        except Exception as e:
            print(f"Error in spacebar hold worker: {e}")
        finally:
            # Always release spacebar when stopping
            try:
                keyboard.release('space')
            except:
                pass
    
    def setup_global_hotkey(self):
        """Setup global hotkey for closing the application"""
        try:
            # Register Alt+Q hotkey to close the application
            keyboard.add_hotkey('alt+q', self.force_close_app, suppress=True)
        except Exception as e:
            print(f"Error setting up global hotkey: {e}")
    
    def force_close_app(self):
        """Force close the application when Alt+Q is pressed"""
        # Schedule the close operation on the main thread
        self.root.after(0, self.emergency_close)
            
    def start_application(self):
        """Start all key pressing threads (called after countdown)"""
        self.is_running = True
        self.start_stop_btn.config(text="Stop")
        self.status_label.config(text="Status: Running", foreground="green")
        space_stop_event = threading.Event()
        # Start a thread for each key sequence
        for seq in self.key_sequences:
            stop_event = threading.Event()
            self.stop_events[seq['id']] = stop_event
            
            thread = threading.Thread(
                target=self.key_press_worker,
                args=(seq['key'], seq['interval'], stop_event),
                daemon=True
            )
            self.threads[seq['id']] = thread
            thread.start()
            
    def stop_application(self):
        """Stop all key pressing threads"""
        self.is_running = False
        self.start_stop_btn.config(text="Start")
        self.status_label.config(text="Status: Stopped", foreground="red")
        
        # Signal all threads to stop
        for stop_event in self.stop_events.values():
            stop_event.set()
            
        # Wait for all threads to finish
        for thread in self.threads.values():
            thread.join(timeout=1.0)
            
        # Clear thread dictionaries
        self.threads.clear()
        self.stop_events.clear()
        
    def key_press_worker(self, key, interval, stop_event):
        """Worker function that presses a key at specified intervals"""
        while not stop_event.is_set():
            try:
                keyboard.press_and_release(key)
            except Exception as e:
                print(f"Error pressing key '{key}': {e}")
            
            # Use stop_event.wait() instead of time.sleep() for responsive stopping
            if stop_event.wait(interval):
                break
                
    def save_config(self):
        """Save configuration to file"""
        try:
            config = {
                'key_sequences': self.key_sequences
            }
            with open('auto_key_presser_config.json', 'w') as f:
                json.dump(config, f, indent=2)
        except Exception as e:
            print(f"Error saving config: {e}")
            
    def load_config(self):
        """Load configuration from file"""
        try:
            if os.path.exists('auto_key_presser_config.json'):
                with open('auto_key_presser_config.json', 'r') as f:
                    config = json.load(f)
                    
                self.key_sequences = config.get('key_sequences', [])
                
                # Populate treeview with loaded data
                for seq in self.key_sequences:
                    self.tree.insert("", tk.END, values=(seq['key'], seq['interval']), tags=(seq['id'],))
                    
        except Exception as e:
            print(f"Error loading config: {e}")
            self.key_sequences = []
            
    def on_closing(self):
        """Handle application closing"""
        if self.is_starting and self.countdown_active:
            self.cancel_start()
        if self.is_running:
            self.stop_application()
        if self.spacebar_active:
            self.stop_spacebar_hold()
        self.cleanup_hotkeys()
        self.save_config()
        self.root.destroy()
    
    def emergency_close(self):
        """Emergency close function for Alt+Q hotkey"""
        try:
            # Stop all operations immediately
            self.countdown_active = False
            self.is_starting = False
            
            if self.is_running:
                self.stop_application()
            
            if self.spacebar_active:
                self.spacebar_active = False  # Set directly to avoid checkbox issues
                try:
                    keyboard.release('space')
                except:
                    pass
            
            self.cleanup_hotkeys()
            self.save_config()
            self.root.quit()  # Force quit
            self.root.destroy()
        except Exception as e:
            print(f"Error during emergency close: {e}")
            # Force exit if all else fails
            import sys
            sys.exit(0)
    
    def cleanup_hotkeys(self):
        """Clean up global hotkeys"""
        try:
            keyboard.unhook_all_hotkeys()
        except Exception as e:
            print(f"Error cleaning up hotkeys: {e}")

def main():
    root = tk.Tk()
    app = AutoKeyPresser(root)
    
    # Handle window closing
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    root.mainloop()

if __name__ == "__main__":
    main()
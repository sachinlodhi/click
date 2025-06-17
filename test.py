import pyautogui
import time
import os
import json
import sys
from datetime import datetime
import pandas as pd
import pyperclip

class DataExtractor:
    def __init__(self):
        self.textbox_position = None
        self.cell_positions = []  # List to store 10 cell positions
        self.extracted_data = []  # Store all test results
        self.positions_file = "saved_positions.json"  # File to save positions
        
        # Create output directory
        self.output_dir = f"extracted_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        os.makedirs(self.output_dir, exist_ok=True)
        
        # Disable pyautogui failsafe (optional)
        pyautogui.FAILSAFE = True
        
    def save_positions(self):
        """Save textbox and cell positions to file"""
        try:
            positions_data = {
                'textbox_position': self.textbox_position,
                'cell_positions': self.cell_positions,
                'saved_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
            with open(self.positions_file, 'w') as f:
                json.dump(positions_data, f, indent=2)
            
            print(f"✓ Positions saved to {self.positions_file}")
            return True
            
        except Exception as e:
            print(f"Error saving positions: {e}")
            return False
    
    def load_positions(self):
        """Load textbox and cell positions from file"""
        try:
            if not os.path.exists(self.positions_file):
                return False
                
            with open(self.positions_file, 'r') as f:
                positions_data = json.load(f)
            
            self.textbox_position = tuple(positions_data['textbox_position']) if positions_data['textbox_position'] else None
            self.cell_positions = [tuple(pos) for pos in positions_data['cell_positions']]
            
            print(f"✓ Positions loaded from {self.positions_file}")
            print(f"  Saved on: {positions_data.get('saved_date', 'Unknown')}")
            print(f"  Textbox: {self.textbox_position}")
            print(f"  Cell positions: {len(self.cell_positions)} cells loaded")
            
            return True
            
        except Exception as e:
            print(f"Error loading positions: {e}")
            return False
    
    def ask_load_or_set_positions(self):
        """Ask user if they want to load saved positions or set new ones"""
        if os.path.exists(self.positions_file):
            print(f"\n📁 Found saved positions file: {self.positions_file}")
            
            # Try to load and show preview
            try:
                with open(self.positions_file, 'r') as f:
                    data = json.load(f)
                print(f"  Saved on: {data.get('saved_date', 'Unknown')}")
                print(f"  Textbox position: {data.get('textbox_position')}")
                print(f"  Cell positions: {len(data.get('cell_positions', []))} cells")
            except:
                print("  (File exists but couldn't preview)")
            
            choice = input("\nDo you want to load saved positions? (y/n): ").lower().strip()
            
            if choice in ['y', 'yes']:
                if self.load_positions():
                    # Verify positions are valid
                    if self.textbox_position and len(self.cell_positions) == 10:
                        print("✓ All positions loaded successfully!")
                        return True
                    else:
                        print("❌ Loaded positions seem incomplete, will set manually")
                        return False
                else:
                    print("❌ Failed to load positions, will set manually")
                    return False
            else:
                print("📝 Will set positions manually")
                return False
        else:
            print(f"\n📝 No saved positions found, will set manually")
            return False
        
    def set_textbox_position(self):
        """Let user set the textbox position"""
        print("Step 1: Set the TEXTBOX position (where you'll type values)")
        print("Move mouse to the textbox and press ENTER...")
        
        try:
            input("Press ENTER when mouse is over the textbox: ")
            self.textbox_position = pyautogui.position()
            print(f"Textbox position set: {self.textbox_position}")
            return True
            
        except Exception as e:
            print(f"Error setting textbox position: {e}")
            return False
    
    def set_cell_positions(self):
        """Let user set positions for table cells (5 rows × 2 columns = 10 cells)"""
        print(f"\nStep 2: Set positions for table cells (5 rows × 2 columns)")
        print("You'll position mouse on each cell one by one...")
        
        cell_labels = [
            "Row 1, Col 1", "Row 1, Col 2",
            "Row 2, Col 1", "Row 2, Col 2", 
            "Row 3, Col 1", "Row 3, Col 2",
            "Row 4, Col 1", "Row 4, Col 2",
            "Row 5, Col 1", "Row 5, Col 2"
        ]
        
        self.cell_positions = []
        
        for i, label in enumerate(cell_labels):
            print(f"\nSetting position {i+1}/10: {label}")
            print("Move mouse to this cell and press ENTER...")
            
            try:
                input(f"Press ENTER when mouse is over {label}: ")
                pos = pyautogui.position()
                self.cell_positions.append(pos)
                print(f"Position {i+1} set: {pos}")
                
            except Exception as e:
                print(f"Error setting position {i+1}: {e}")
                return False
        
        print(f"\nAll 10 cell positions set successfully!")
        return True
    
    def extract_cell_value(self, cell_position, cell_name):
        """Extract text from a cell using double-click and copy"""
        try:
            # Move to cell and double-click to select text
            pyautogui.click(cell_position)
            #time.sleep(0.02)
            pyautogui.doubleClick(cell_position)
            #time.sleep(0.02)
            
            # Copy selected text
            pyautogui.hotkey('ctrl', 'c')
            #time.sleep(0.05)  # Give time for copy to complete
            
            # Get copied text from clipboard
            copied_text = pyperclip.paste()
            
            print(f"  {cell_name}: '{copied_text}'")
            return copied_text.strip()
            
        except Exception as e:
            print(f"  Error extracting {cell_name}: {e}")
            return ""
    
    def extract_all_cells(self):
        """Extract values from all 10 table cells"""
        print("  Extracting data from all cells...")
        
        cell_values = []
        cell_names = [f"Cell_{i+1}" for i in range(10)]
        
        for i, (position, name) in enumerate(zip(self.cell_positions, cell_names)):
            value = self.extract_cell_value(position, name)
            cell_values.append(value)
        
        return cell_values
    
    def auto_test_parameters(self, start_val, end_val, step_val, delay=0.01):
        """Main parameter testing loop with data extraction"""
        if self.textbox_position is None or len(self.cell_positions) != 10:
            print("Textbox position or cell positions not set!")
            return False
            
        # Generate test values normally
        if step_val == int(step_val):  # Integer step
            test_values = list(range(int(start_val), int(end_val) + 1, int(step_val)))
        else:  # Float step
            test_values = []
            val = start_val
            while val <= end_val:
                test_values.append(round(val, 3))
                val += step_val
        
        print(f"\nStarting parameter testing...")
        test_values.insert(0, 5) # manually inserting
        print(f"Test values: {test_values}")
        print(f"Will test {len(test_values)} different parameters")
        print("Note: First value will be tested twice (saving only second result)")
        print("Press Ctrl+C to stop early\n")
        
        try:
            first_time = False
            for i, test_value in enumerate(test_values):
                print(f"\nTest {i+1}/{len(test_values)}: Testing value {test_value}")
                
                # 1. Double-click on textbox to select existing text
                pyautogui.doubleClick(self.textbox_position)
                #time.sleep(0.01)
                
                # 2. Press backspace multiple times to ensure clearing
                pyautogui.press('backspace', presses=10)  # Clear up to 10 characters
                #time.sleep(0.01)
                
                # 3. Type new value
                pyautogui.typewrite(str(test_value))
                
                # 4. Press Enter to execute
                pyautogui.press('enter')
                
                # 5. Wait for processing
                #time.sleep(delay)
                if first_time == False:
                  print("first time detected, skipping saving")
                  first_time = True
                  continue
                
                # 5. Extract data from all cells
                cell_values = self.extract_all_cells()
                
                # 6. Store results - convert 10 cell values to 5x2 table format
                # Each test parameter generates 5 rows of data (row1: col1,col2; row2: col1,col2; etc.)
                for row_idx in range(5):
                    col1_value = cell_values[row_idx * 2]      # Even indices: 0,2,4,6,8
                    col2_value = cell_values[row_idx * 2 + 1]  # Odd indices: 1,3,5,7,9
                    
                    result_row = {
                        #'test_parameter': test_value,
                        #'row_number': row_idx + 1,
                        'I': col1_value,
                        'V': col2_value
                    }
                    
                    self.extracted_data.append(result_row)
                
                print(f"✓ Completed test for value {test_value}")
                
            print(f"\nParameter testing complete! {len(test_values)} parameters tested")
            print(f"Generated {len(self.extracted_data)} total rows (5 rows per parameter)")
            
            # Save extracted data to CSV
            self.save_results_to_csv()
            
            return True
            
        except KeyboardInterrupt:
            completed_tests = len(set(row['test_parameter'] for row in self.extracted_data))
            print(f"\nParameter testing stopped by user.")
            print(f"Completed {completed_tests} parameters, collected {len(self.extracted_data)} rows")
            self.save_results_to_csv()
            return False
        except Exception as e:
            print(f"Error during parameter testing: {e}")
            return False
    
    def save_results_to_csv(self):
        """Save extracted data to CSV file"""
        if not self.extracted_data:
            print("No data to save!")
            return
            
        try:
            df = pd.DataFrame(self.extracted_data)
            csv_filename = os.path.join(self.output_dir, 'extracted_data.csv')
            df.to_csv(csv_filename, index=False)
            
            print(f"\n✓ Data saved to: {csv_filename}")
            print(f"✓ Total rows in CSV: {len(self.extracted_data)}")
            print(f"✓ Parameters tested: {len(set(row['test_parameter'] for row in self.extracted_data))}")
            print("\nData preview:")
            print(df.head(10).to_string(index=False))  # Show first 10 rows
            
        except Exception as e:
            print(f"Error saving CSV: {e}")
    
    def run(self):
        """Main program workflow"""
        print("=== Simple Parameter Testing Data Extractor ===")
        print("This program will:")
        print("1. Set textbox position (where to type test values)")
        print("2. Set 10 cell positions (5 rows × 2 columns table)")
        print("3. Prime textbox with first value (don't save)")
        print("4. Test all values including first one again (save all)")
        print("5. Each test parameter creates 5 rows in CSV (one per table row)\n")
        print("CSV structure: test_parameter, row_number, column_1, column_2")
        
        # Try to load saved positions first
        if self.ask_load_or_set_positions():
            print("✓ Using loaded positions")
        else:
            # Step 1: Set textbox position
            if not self.set_textbox_position():
                print("Failed to set textbox position. Exiting.")
                return
                
            # Step 2: Set cell positions
            if not self.set_cell_positions():
                print("Failed to set cell positions. Exiting.")
                return
            
            # Ask if user wants to save positions
            save_choice = input("\n💾 Save these positions for future runs? (y/n): ").lower().strip()
            if save_choice in ['y', 'yes']:
                self.save_positions()
        
        # Step 3: Configure parameter range
        try:
            print(f"\nParameter Range Configuration:")
            start_val = float(input(f"Start value (e.g., 1): ") or "1")
            end_val = float(input(f"End value (e.g., 10): ") or "10")
            step_val = float(input(f"Step size (e.g., 1 or 0.5): ") or "1")
            delay = float(input(f"Delay after pressing Enter (seconds, default 0.2): ") or "0.2")
            
            print(f"\nTest range: {start_val} to {end_val}, step {step_val}")
            print("Method: Prime with first value → Test all values (including first again)")
            
        except ValueError:
            print("Invalid input, using defaults (1 to 10, step 1, delay 0.2)")
            start_val, end_val, step_val, delay = 1, 10, 1, 0.01
            
        # Step 4: Countdown
        print(f"\nStarting parameter testing in 2 seconds...")
        for i in range(2, 0, -1):
            print(f"{i}...")
            time.sleep(1)
        print("Starting!")
        
        # Step 5: Start automated parameter testing
        self.auto_test_parameters(start_val, end_val, step_val, delay)

if __name__ == "__main__":
    extractor = DataExtractor()
    
    # Check for command line arguments
    if len(sys.argv) > 1:
        if sys.argv[1] == "--clear-positions":
            if os.path.exists(extractor.positions_file):
                os.remove(extractor.positions_file)
                print(f"✓ Cleared saved positions ({extractor.positions_file})")
            else:
                print("No saved positions to clear")
            sys.exit(0)
        elif sys.argv[1] == "--show-positions":
            if os.path.exists(extractor.positions_file):
                with open(extractor.positions_file, 'r') as f:
                    data = json.load(f)
                print("Saved positions:")
                print(json.dumps(data, indent=2))
            else:
                print("No saved positions found")
            sys.exit(0)
    
    extractor.run()
    print("position")

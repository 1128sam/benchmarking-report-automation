import os
import time
import shutil
import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

WATCH_FOLDER = r"C:/automation testing"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = r"C:/automation testing"

def get_unique_path(path):
    """Return a path that doesn't exist yet, adding (1), (2), ... before the extension."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    counter = 1
    while os.path.exists(f"{base} ({counter}){ext}"):
        counter += 1
    return f"{base} ({counter}){ext}"

class PDFDropHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        
        # Only trigger for PDF files
        if not event.src_path.lower().endswith('.pdf'):
            return

        print(f"\nNew PDF Detected: {os.path.basename(event.src_path)}")
        
        # Debounce: Wait to ensure the file is completely dropped and written to disk
        stable_count = 0
        last_size = -1
        while stable_count < 3:
            time.sleep(1)  # Check every second
            current_size = os.path.getsize(event.src_path)
            if current_size == last_size:
                stable_count += 1
            else:
                stable_count = 0
            last_size = current_size
        print("File download complete. Starting pipeline...")

        try:
            for item in os.listdir(WATCH_FOLDER):
                item_path = os.path.join(WATCH_FOLDER, item)
                if os.path.isdir(item_path) and item.endswith("_output_files_1"):
                    try:
                        shutil.rmtree(item_path)
                        print(f"Cleared previous run folder: {item}")
                    except Exception as e:
                        print(f"Could not remove {item}: {e}")
            # 2. Run PDF extraction code
            print(f"Step 1: Extracting text & generating CSVs from "f"{event.src_path}...")
            subprocess.run(["python", "app_test.py", event.src_path], check=True)

            # 3. Run chart creation code
            print("Step 2: Creating native charts and compiling presentation (make_chart.py)...")
            subprocess.run(["python", "make_chart.py"], check=True)

            # 4. Move final PPT to the user's desktop
            last_output_file = os.path.join(OUTPUT_DIR, "last_output.txt")
            with open(last_output_file, "r") as f:
                generated_ppt = f.read().strip()

            desktop_path = os.path.expanduser("~/Desktop")
            final_destination = get_unique_path(
                os.path.join(desktop_path, os.path.basename(generated_ppt))
            )

            shutil.copy(generated_ppt, final_destination)
            print(f"Success! PowerPoint saved to your Desktop: {os.path.basename(final_destination)}\n")
            
        except subprocess.CalledProcessError as e:
            print(f"Automation failed during script execution: {e}")
        except Exception as e:
            print(f"Unexpected Error: {e}")

if __name__ == "__main__":
    if not os.path.exists(WATCH_FOLDER):
        os.makedirs(WATCH_FOLDER)
        print(f"Created watcher folder at: {WATCH_FOLDER}")

    event_handler = PDFDropHandler()
    observer = Observer()
    observer.schedule(event_handler, path=WATCH_FOLDER, recursive=False)
    
    print(f"Service Active. Drag & drop your benchmarking PDFs into '{WATCH_FOLDER}'...")
    observer.start()
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
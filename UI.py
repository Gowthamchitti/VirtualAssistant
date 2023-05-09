import tkinter as tk
from PIL import Image, ImageTk
import subprocess
import threading
import io
import sys
class App:
    def __init__(self, master):
        self.master = master
        master.title("Virtual Assistant")
        self.start_button = tk.Button(master, text="Start", command=self.start_thread, fg="white", bg="green")
        self.start_button.pack(side="left", padx=100, pady=100)
        self.stop_button = tk.Button(master, text="Stop", command=self.stop_process, fg="white", bg="red")
        self.stop_button.pack(side="right", padx=100, pady=100)
        self.terminate_button = tk.Button(master, text="Exit", command= master.destroy,bg="black",fg="white")
        self.terminate_button.pack(side="bottom", padx=10, pady=100)
        self.process = None
        self.running = False
        self.thread = None
        self.label = None
        self.text = tk.Text(self.master)
        self.text.pack(expand=True, fill="both")

    def start_thread(self):
        if self.thread is None or not self.thread.is_alive():
            self.running = True
            self.text.delete("1.0", "end") # clear text widget
            self.thread = threading.Thread(target=self.run_path_code)
            self.thread.start()
            self.show_loading_animation()

    def run_path_code(self):
        process = subprocess.Popen(['python', 'C:/Users/gouth/PycharmProjects/VirtualAssistant/Tempimplemnt.py'], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, universal_newlines=True)
        for line in iter(process.stdout.readline, b''):
            if not self.running:
                break
            if line.startswith("DISPLAY_IMAGE!"):
                image_path = line.split("!")[1].strip()
                print(image_path)
                self.display_image(image_path,width=200,height=250)
            else:
                self.text.insert("end", line)
                self.text.see("end")
        process.kill()

    def stop_process(self):
        self.running = False
        if self.process is not None:
            self.process.kill()
        if self.label is not None:
            self.label.destroy()

    def show_loading_animation(self):
        self.label = tk.Label(self.master,text="Running....")
        self.label.pack()

    def animate_loading(self, frame):
        self.label.configure(text="Running....")
        self.master.after(100, self.animate_loading, frame)

    def display_image(self, image_path, width=None, height=None):
        # Open image using Pillow
        image = Image.open(image_path)
        # Resize image to fit label
        if width and height:
            image = image.resize((width, height))
        # Create PhotoImage object from PIL image
        photo = ImageTk.PhotoImage(image)
        # Create label with the PhotoImage object
        self.image_label=tk.Label(self.text,image=photo)
        self.image_label.photo=photo
        self.text.window_create("end",window=self.image_label)
    def remove_image(self):
        self.image_label.destroy()

root = tk.Tk()
app = App(root)
root.mainloop()

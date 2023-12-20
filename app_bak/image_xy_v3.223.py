import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, IntVar, Radiobutton
from PIL import Image, ImageTk
import openpyxl
from datetime import datetime
import os

class PixelExtractor:
    def __init__(self, root):
        self.root = root
        self.application_running = True
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.title("Pixel Extractor")
        self.root.geometry("1850x950")

        self.image_folder_path = ""
        self.image_list = []
        self.current_image_index = 0
        self.click_count = 0

        self.alpha, self.beta = 0, 0

        self.window_frame = tk.Frame(root, width=1850, height=950)
        self.window_frame.pack(fill="both", expand=True)

        self.image_frame = tk.Frame(self.window_frame, width=1600, height=900)
        self.image_frame.pack(side="left", fill="both", expand=True)

        self.info_frame = tk.Frame(self.window_frame, width=250, height=900, bg="lightgray")
        self.info_frame.pack(side="right", fill="both", expand=False)

        self.canvas = tk.Canvas(self.image_frame, width=1600, height=900)
        self.canvas.pack(side="top", fill="both", expand=True)
        self.scrollbar_y = tk.Scrollbar(self.image_frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar_y.pack(side="right", fill="y")
        self.scrollbar_x = tk.Scrollbar(self.image_frame, orient="horizontal", command=self.canvas.xview)
        self.scrollbar_x.pack(side="bottom", fill="x")
        self.canvas.configure(xscrollcommand=self.scrollbar_x.set, yscrollcommand=self.scrollbar_y.set)
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.canvas.bind("<Button-1>", self.on_click)

        self.info_label = tk.Label(self.info_frame, text="", justify="left", bg="lightgray", padx=5, pady=5)
        self.info_label.pack(side="top", fill="both", expand=False)

        self.image_name_label = tk.Label(self.info_frame, text="Image Name:", bg="lightgray", padx=5, pady=5)
        self.image_name_label.pack(side="top", fill="both", expand=False)

        self.tk_image = None
        self.click_coords = []

        self.scale_var = IntVar()
        self.scale_var.set(-1)

        self.dogak_list_file_path = ""

        scale_frame = tk.Frame(self.info_frame, bg="lightgray")
        scale_frame.pack(side="top", fill="both", expand=False)

        scale_label = tk.Label(scale_frame, text="Scale:", bg="lightgray", padx=5, pady=5)
        scale_label.grid(row=0, column=0)

        scales = [-1, 500, 600, 1000, 1200, 6000]
        for i, scale in enumerate(scales):
            rb = Radiobutton(scale_frame, text=str(scale), variable=self.scale_var, value=scale, bg="lightgray")
            rb.grid(row=i + 1, column=0, sticky="w")

        self.load_images()

    def on_close(self):
        self.application_running = False
        self.root.destroy()

    def get_alpha_beta(self, selected_scale):
        if selected_scale == 500:
            self.alpha, self.beta = 200, 150
        elif selected_scale == 600:
            self.alpha, self.beta = 250, 200
        elif selected_scale == 1000:
            self.alpha, self.beta = 400, 300
        elif selected_scale == 1200:
            self.alpha, self.beta = 500, 400
        elif selected_scale == 6000:
            self.alpha, self.beta = 2500, 2000
        return self.alpha, self.beta

    def load_images(self):
        # Get the image files
        image_files = filedialog.askopenfilenames(
            title="Select Image Files",
            filetypes=[("JPEG files", "*.jpg")]
        )

        self.image_list = sorted(list(image_files))

        if not self.image_list:
            messagebox.showinfo("Error", "No jpg files selected.")
            self.root.destroy()
            return

        # Get the single text file path from the user
        self.dogak_list_file_path = filedialog.askopenfilename(
            title="Select a text file",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )

        if not self.dogak_list_file_path:
            messagebox.showinfo("Error", "No text file selected.")
            self.root.destroy()
            return

        # Load the first image
        self.load_image()
        self.load_target_dogak_coordinates(
            os.path.basename(self.image_list[self.current_image_index]), self.dogak_list_file_path
        )

    def load_image(self):
        if self.current_image_index < len(self.image_list):
            image_path = self.image_list[self.current_image_index]
            self.original_image = Image.open(image_path).copy()
            self.tk_image = ImageTk.PhotoImage(self.original_image)
            self.image_item = self.canvas.create_image(0, 0, anchor=tk.NW, image=self.tk_image)
            image_filename = os.path.splitext(os.path.basename(image_path))[0]
            self.image_name_label.config(text=f"Image Name: {image_filename}")

            if os.path.exists("image.xlsx"):
                self.workbook = openpyxl.load_workbook("image.xlsx")
            else:
                self.workbook = openpyxl.Workbook()
                self.workbook.save("image.xlsx")

            self.sheet = self.workbook.active

            self.show_position(f"Click on image {self.current_image_index + 1}, point {self.click_count + 1}")

    def save_to_excel(self, image_filename, positions):
        clicked_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for i, pos in enumerate(positions):
            user_x, user_y, pixel_x, pixel_y = pos
            self.sheet.append([image_filename, f"Point {i + 1}", f"{user_x}, {user_y}", f"{pixel_x}, {pixel_y}", clicked_time])

        self.workbook.save("image.xlsx")

    def save_coordinates_to_file(self, alpha, beta):
        if self.click_coords:
            current_image_index = self.current_image_index
            image_path = self.image_list[current_image_index]
            image_filename = os.path.splitext(os.path.basename(image_path))[0]
            points_filename = image_filename + ".points"
            image_directory = os.path.dirname(image_path)
            filepath = os.path.join(image_directory, points_filename)

            try:
                with open(filepath, "w") as file:
                    header_line = "mapX,mapY,pixelX,pixelY,enable,dX,dY,residual\n"
                    file.write(header_line)

                    mapx1, mapy1, pixelx1, pixely1 = self.click_coords[0]
                    coords_line = f"{mapx1},{mapy1},{pixelx1},{-pixely1},1,0,0,0\n"
                    file.write(coords_line)

                    _, _, pixelx2, pixely2 = self.click_coords[1]
                    mapx2 = mapx1 + alpha
                    mapy2 = mapy1
                    coords_line = f"{mapx2},{mapy2},{pixelx2},{-pixely2},1,0,0,0\n"  # 수정된 부분
                    file.write(coords_line)

                    _, _, pixelx3, pixely3 = self.click_coords[2]
                    mapx3 = mapx1 + alpha
                    mapy3 = mapy1 + beta
                    coords_line = f"{mapx3},{mapy3},{pixelx3},{-pixely3},1,0,0,0\n"  # 수정된 부분
                    file.write(coords_line)

            except Exception as e:
                messagebox.showinfo("Error", f"Error saving coordinates to file: {e}")
            finally:
                if 'file' in locals():
                    if not file.closed:
                        file.close()

    def load_target_dogak_coordinates(self, target_filename, dogak_list_file_path):
        user_x, user_y = 0, 0  # Default values

        try:
            print(f"Looking for {target_filename} in {dogak_list_file_path}")
            with open(dogak_list_file_path, "r") as file:
                lines = file.readlines()

                # Iterate through lines to find the image name
                for line in lines:
                    fields = line.strip().split(',')
                    if len(fields) >= 3:
                        image_filename, x, y = fields[0], float(fields[1]), float(fields[2])

                        # Print each filename in the text file for debugging
                        print(f"Checking filename in the text file: {image_filename}")

                        # Check if the image filename matches the target filename
                        if os.path.splitext(image_filename)[:15] == os.path.splitext(target_filename)[:15]:
                            user_x, user_y = x, y
                            print(f"Found coordinates for {target_filename}: user_x={user_x}, user_y={user_y}")
                            break  # Stop searching once a match is found

        except Exception as e:
            messagebox.showinfo("Error", f"Error loading coordinates from the text file: {e}")

        print(f"Final user_x: {user_x}, user_y: {user_y}")
        return user_x, user_y


    def on_click(self, event):
        if not self.application_running:
            return

        selected_scale = self.scale_var.get()
        alpha, beta = self.get_alpha_beta(selected_scale)

        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        user_x, user_y = 0, 0

        # Move to the bottom of the image on click_count == 0
        self.canvas.xview_moveto(0.0)
        self.canvas.yview_moveto(1.1)

        if self.click_count == 0:
            user_x, user_y = self.load_target_dogak_coordinates(os.path.basename(self.image_list[self.current_image_index]), self.dogak_list_file_path)
            pixelx1, pixely1 = x, y
            self.click_coords.append((user_x, user_y, pixelx1, pixely1))

            # Move to the top of the image on click_count == 1
            self.canvas.xview_moveto(1.1)

        elif self.click_count == 1:
            x1, y1, pixelx1, pixely1 = self.click_coords[0]
            user_x = x1 + alpha
            user_y = y1
            pixelx2, pixely2 = x, y
            self.click_coords.append((user_x, user_y, pixelx2, pixely2))

            # Move to the left of the image on click_count == 2
            self.canvas.xview_moveto(1.1)
            self.canvas.yview_moveto(0.0)

        elif self.click_count == 2:
            _, _, pixelx2, pixely2 = self.click_coords[1]
            pixelx3, pixely3 = x, y
            user_x = pixelx2 + alpha
            user_y = pixely2 + beta
            self.click_coords.append((user_x, user_y, pixelx3, pixely3))

        self.click_count += 1

        if self.click_count == 3:
            image_filename = os.path.splitext(os.path.basename(self.image_list[self.current_image_index]))[0]
            self.save_to_excel(image_filename, self.click_coords)
            self.save_coordinates_to_file(alpha, beta)

            self.click_count = 0
            self.current_image_index += 1
            self.click_coords = []
            self.canvas.delete("all")

            if self.current_image_index < len(self.image_list):
                self.load_image()
            else:
                self.workbook.save("image.xlsx")
                self.application_running = False
                self.root.destroy()

        if self.application_running and self.root and self.root.winfo_exists():
            info_text = f"Image: {self.current_image_index + 1}\nPoint: {self.click_count + 1}\nCoordinates: ({x}, {y})"
            self.info_label.config(text=info_text)

            if self.scrollbar_y.get()[1] < 1:
                self.scrollbar_y.set(self.scrollbar_y.get()[0], self.scrollbar_y.get()[1] + 0.01)


    def on_canvas_configure(self, event):
        if self.tk_image:
            self.canvas.config(scrollregion=(0, 0, self.original_image.width, self.original_image.height))
            self.scrollbar_y.place(x=self.image_frame.winfo_width() - 20, y=0, relheight=1)

    def show_position(self, text):
        if self.application_running and hasattr(self, 'info_label') and self.info_label.winfo_exists() and self.root and self.root.winfo_exists():
            self.info_label.config(text=text)

if __name__ == "__main__":
    root = tk.Tk()
    app = PixelExtractor(root)
    root.mainloop()

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

        # 축척 라디오 버튼 정의
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
            self.alpha, self.beta = 200, 150
        elif selected_scale == 1200:
            self.alpha, self.beta = 500, 400
        elif selected_scale == 6000:
            self.alpha, self.beta = 2500, 2000
        return self.alpha, self.beta

    def load_images(self):
        # 이미지 파일을 선택하는 대화상자 호출
        image_files = filedialog.askopenfilenames(title="Select Image Files", filetypes=[("JPEG files", "*.jpg")])
        self.image_list = sorted(list(image_files))

        if not self.image_list:
            messagebox.showinfo("Error", "No jpg files selected.")
            self.root.destroy()
            return

        self.load_image()

    def load_image(self):
        # 이미지를 불러오고 화면에 표시
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

            # 사용자가 첫 번째 지점을 클릭하도록 안내
            self.show_position(f"Click on image {self.current_image_index + 1}, point {self.click_count + 1}")

    def save_to_excel(self, image_filename, positions):
        clicked_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for i, pos in enumerate(positions, start=1):
            user_x1, user_y1, x1, y1 = pos
            self.sheet.append([image_filename, f"Point {i}", f"{user_x1}, {user_y1}", f"{x1}, {y1}", clicked_time])

        self.workbook.save("image.xlsx")

    def on_click(self, event):
        if not self.application_running:
            return

        selected_scale = self.scale_var.get()
        alpha, beta = self.get_alpha_beta(selected_scale)

        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)

        # Initialize user_x and user_y outside the if statement
        user_x, user_y = 0, 0

        if self.click_count < 3:
            if self.click_count == 0:
                user_x = simpledialog.askfloat(f"X Coordinate for point {self.click_count + 1}",
                                               f"Enter X coordinate for point {self.click_count + 1}",
                                               parent=self.image_frame)
                user_y = simpledialog.askfloat(f"Y Coordinate for point {self.click_count + 1}",
                                               f"Enter Y coordinate for point {self.click_count + 1}",
                                               parent=self.image_frame)
                self.click_coords.append((user_x, user_y, x, y))
            elif self.click_count == 1:
                x1, y1, _, _ = self.click_coords[0]
                user_x = x1 + alpha
                user_y = y1 + beta
                self.click_coords.append((user_x, user_y, x, y))
            elif self.click_count == 2:
                _, _, x2, y2 = self.click_coords[1]
                user_x = x2 + alpha
                user_y = y2
                self.click_coords.append((user_x, user_y, x, y))

            self.click_count += 1
        else:
            image_filename = os.path.splitext(os.path.basename(self.image_list[self.current_image_index]))[0]
            self.save_to_excel(image_filename, self.click_coords)
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

        self.save_coordinates_to_file()

        if self.application_running and self.root and self.root.winfo_exists():
            info_text = f"Image: {self.current_image_index + 1}\nPoint: {self.click_count + 1}\nCoordinates: ({x}, {y})"
            self.info_label.config(text=info_text)

            if self.scrollbar_y.get()[1] < 1:
                self.scrollbar_y.set(self.scrollbar_y.get()[0], self.scrollbar_y.get()[1] + 0.01)

    def save_to_excel(self, image_filename, positions):
        clicked_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for i, pos in enumerate(positions, start=1):
            user_x1, user_y1, x1, y1 = pos
            self.sheet.append([image_filename, f"Point {i}", f"{user_x1}, {user_y1}", f"{x1}, {y1}", clicked_time])

        self.workbook.save("image.xlsx")

        # 파일 입력 형태에 따라서 파일에 저장
        with open(f"{image_filename}.points", "w") as file:
            file.write(f"{user_x1},{user_y1},{x1},{y1},1,0,0,0\n")
            file.write(f"{user_x1 + alpha},{user_y1 + beta},{x1},{y1},1,0,0,0\n")
            file.write(f"{user_x1 + alpha},{user_y1},{x1},{y1},1,0,0,0\n")

    def save_coordinates_to_file(self):
        if self.click_coords:
            current_image_index = self.current_image_index
            image_path = self.image_list[current_image_index]
            image_filename = os.path.splitext(os.path.basename(image_path))[0]
            points_filename = image_filename + ".points"
            image_directory = os.path.dirname(image_path)
            filepath = os.path.join(image_directory, points_filename)

            try:
                with open(filepath, "w") as file:
                    if len(self.click_coords) > 0:
                        additional_data_line = "mapX,mapY,pixelX,pixelY,enable,dX,dY,residual\n"
                        file.write(additional_data_line)

                    for coords_set in self.click_coords:
                        user_x1, user_y1, x1, y1 = coords_set
                        coords_line = f"{user_x1},{user_y1},{x1},{y1},1,0,0,0\n"
                        file.write(coords_line)

            except Exception as e:
                messagebox.showinfo("Error", f"Error saving coordinates to file: {e}")
            finally:
                if 'file' in locals():
                    if not file.closed:
                        file.close()


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

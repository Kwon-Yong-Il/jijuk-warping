import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from PIL import Image, ImageTk
import openpyxl
from datetime import datetime
import os

class PixelExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("Pixel Extractor")
        self.root.geometry("1850x950")  # Change the size to 1850x950

        self.image_folder_path = ""
        self.image_list = []
        self.current_image_index = 0
        self.click_count = 0  # To track the number of clicks on the current image

        # Frame for the entire window
        self.window_frame = tk.Frame(root, width=1850, height=950)
        self.window_frame.pack(fill="both", expand=True)

        # Frame for image display
        self.image_frame = tk.Frame(self.window_frame, width=1600, height=900)
        self.image_frame.pack(side="left", fill="both", expand=True)

        # Canvas for image
        self.canvas = tk.Canvas(self.image_frame, width=1600, height=900)
        self.canvas.pack(side="top", fill="both", expand=True)

        # Scrollbars for the canvas
        self.scrollbar_y = tk.Scrollbar(self.image_frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar_y.pack(side="right", fill="y")
        self.scrollbar_x = tk.Scrollbar(self.image_frame, orient="horizontal", command=self.canvas.xview)
        self.scrollbar_x.pack(side="bottom", fill="x")
        self.canvas.configure(xscrollcommand=self.scrollbar_x.set, yscrollcommand=self.scrollbar_y.set)
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.canvas.bind("<Button-1>", self.on_click)

        # Frame for information display (on the right)
        self.info_frame = tk.Frame(self.window_frame, width=250, height=900, bg="lightgray")
        self.info_frame.pack(side="right", fill="both", expand=False)

        # Label for displaying information
        self.info_label = tk.Label(self.info_frame, text="", justify="left", bg="lightgray", padx=5, pady=5)
        self.info_label.pack(side="top", fill="both", expand=False)

        # Label for displaying image name
        self.image_name_label = tk.Label(self.info_frame, text="Image Name:", bg="lightgray", padx=5, pady=5)
        self.image_name_label.pack(side="top", fill="both", expand=False)

        self.tk_image = None  # Added tk_image as an instance variable
        self.click_coords = []

        self.load_images()

    def load_images(self):
        # Ask the user to select multiple .jpg files
        image_files = filedialog.askopenfilenames(title="Select Image Files", filetypes=[("JPEG files", "*.jpg")])

        # Convert the result to a list for consistency with the rest of the code
        self.image_list = list(image_files)

        # Display the image name in the label without the path
        image_filename = os.path.basename(self.image_list[self.current_image_index])
        self.info_label.config(text=f"Image Name: {image_filename}")

        if not self.image_list:
            messagebox.showinfo("Error", "No jpg files selected.")
            self.root.destroy()
            return

        # Load the first image
        self.load_image()



    def load_image(self):
        if self.current_image_index < len(self.image_list):
            image_path = os.path.join(self.image_folder_path, self.image_list[self.current_image_index])
            self.original_image = Image.open(image_path)

            

            # Display the image in its original size
            self.tk_image = ImageTk.PhotoImage(self.original_image)
            self.image_item = self.canvas.create_image(0, 0, anchor=tk.NW, image=self.tk_image)

            # Set the canvas scrolling region to the canvas size
            self.canvas.config(scrollregion=(0, 0, self.original_image.width, self.original_image.height))

            # Update the label with the current image name
            self.image_name_label.config(text=f"{self.image_list[self.current_image_index]}")

            # Prepare Excel file
            if os.path.exists("image.xlsx"):
                self.workbook = openpyxl.load_workbook("image.xlsx")
            else:
                self.workbook = openpyxl.Workbook()
                self.workbook.save("image.xlsx")

            self.sheet = self.workbook.active

            self.show_position(f"Click on image {self.current_image_index + 1}, point {self.click_count + 1}")

    def save_to_excel(self, image_filename, position):
        clicked_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.sheet.append([image_filename, *position, clicked_time])
        self.workbook.save("image.xlsx")  # Save after each click

    def on_click(self, event):
        if self.current_image_index >= len(self.image_list):
            # All images processed, close the application
            self.workbook.save("image.xlsx")
            self.save_coordinates_to_file()  # Save coordinates for the last image
            self.root.destroy()
            return

        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)

        if self.click_count == 0:
            user_x = simpledialog.askfloat("X Coordinate", f"Enter X coordinate for point {self.click_count + 1}",
                                           parent=self.image_frame)
            user_y = simpledialog.askfloat("Y Coordinate", f"Enter Y coordinate for point {self.click_count + 1}",
                                           parent=self.image_frame)
            image_filename = self.image_list[self.current_image_index]
            self.save_to_excel(image_filename, (user_x, user_y, x, y))
            self.click_coords.append((user_x, user_y, x, y))
            self.click_count += 1
        elif self.click_count == 1:
            user_x = simpledialog.askfloat("X Coordinate", f"Enter X coordinate for point {self.click_count + 1}",
                                           parent=self.image_frame)
            user_y = simpledialog.askfloat("Y Coordinate", f"Enter Y coordinate for point {self.click_count + 1}",
                                           parent=self.image_frame)
            if user_x is not None and user_y is not None:
                image_filename = self.image_list[self.current_image_index]
                self.click_coords.append((user_x, user_y, x, y))

                # Save the data to Excel for the first point
                self.save_to_excel(image_filename, (user_x, user_y, x, y))

                # Save the additional data to Excel
                self.save_to_excel(image_filename, (self.click_coords[1][0], self.click_coords[0][1],
                                                    self.click_coords[1][2], self.click_coords[0][3]))

                # Save the coordinates to the points file for both points
                self.save_coordinates_to_file()

                self.click_count = 0
                self.current_image_index += 1
                self.click_coords = []  # Clear click coordinates
                self.canvas.delete("all")  # Clear canvas

                if self.current_image_index < len(self.image_list):
                    self.load_image()
                else:
                    # All images processed, close the application
                    self.workbook.save("image.xlsx")
                    self.root.destroy()
            else:
                self.click_count = 0
                self.click_coords = []

        # Display information in the right frame
        info_text = f"Image: {self.current_image_index + 1}\nPoint: {self.click_count + 1}\nCoordinates: ({x}, {y})"
        
        # Check if the label widget exists before configuring it
        if self.info_label.winfo_exists():
            self.info_label.config(text=info_text)

        # Move the vertical scrollbar to the right
        if self.scrollbar_y.get()[1] < 1:
            self.scrollbar_y.set(self.scrollbar_y.get()[0], self.scrollbar_y.get()[1] + 0.01)

        # If two points are clicked, move to the next image
        if self.click_count == 2:
            self.click_count = 0
            self.current_image_index += 1
            self.click_coords = []  # Clear click coordinates
            self.canvas.delete("all")  # Clear canvas

            if self.current_image_index < len(self.image_list):
                self.load_image()
            else:
                # All images processed, close the application
                self.workbook.save("image.xlsx")
                self.root.destroy()
        else:
            # Display instructions for the next point
            self.show_position(f"Click on image {self.current_image_index + 1}, point {self.click_count + 1}")

    def save_coordinates_to_file(self):
        if self.click_coords:  # Check if coordinates are available
            current_image_index = self.current_image_index
            filename = self.image_list[current_image_index].replace(".jpg", ".points")
            filepath = os.path.join(self.image_folder_path, filename)

            with open(filepath, "w") as file:
                # Write the field name to the points file
                if len(self.click_coords) == 2:
                    additional_data_line = f"mapX,mapY,pixelX,pixelY,enable,dX,dY,residual\n"
                    file.write(additional_data_line)
                    
                # Write the coordinates for both points
                for coords_set in self.click_coords:
                    # Write the coordinates including user-entered and image location values
                    x1, y1, x2, y2 = coords_set
                    coords_line = f"{x1},{y1},{x2},{y2 if y2 < 0 else '-' + str(y2)}\n"
                    file.write(coords_line)

                # Append the additional data to the points file
                if len(self.click_coords) == 2:
                    x2, y1, x1, y2 = self.click_coords[1][0], self.click_coords[0][1], self.click_coords[1][2], self.click_coords[0][3]
                    additional_data_line = f"{x2},{y1},{x1},{y2 if y2 < 0 else '-' + str(y2)}\n"
                    file.write(additional_data_line)

    def on_canvas_configure(self, event):
        # Adjust the canvas scrolling region when the window is resized
        if self.tk_image:
            self.canvas.config(scrollregion=(0, 0, self.original_image.width, self.original_image.height))
            self.scrollbar_y.place(x=self.image_frame.winfo_width() - 20, y=0, relheight=1)  # Adjusted position

    def show_position(self, text):
        self.info_label.config(text=text)

        

if __name__ == "__main__":
    root = tk.Tk()
    app = PixelExtractor(root)
    root.mainloop()

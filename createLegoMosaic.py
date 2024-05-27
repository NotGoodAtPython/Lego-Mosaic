from PIL import Image, ImageOps, ImageTk
import xlsxwriter, os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from xlsxwriter.utility import xl_rowcol_to_cell
import numpy as np

def open_image():
    filepath = filedialog.askopenfilename(title = 'Select your image:')
    if not filepath:
        return None

    image = Image.open(filepath)
    
    # Prompt user for new size
    new_size = get_new_size()
    if new_size is None:
        return None

    width, height = new_size
    image = resize_image(image, new_size=(width, height))
    image = recolor_image(image, colors=color_palette)
    return display_image(image, filepath)

def get_new_size():
    size_dialog = tk.Toplevel(root)
    size_dialog.title("Enter Image Size")
    
    tk.Label(size_dialog, text="Width:").pack(padx=10, pady=5)
    width_entry = tk.Entry(size_dialog)
    width_entry.pack(padx=10, pady=5)
    
    tk.Label(size_dialog, text="Height:").pack(padx=10, pady=5)
    height_entry = tk.Entry(size_dialog)
    height_entry.pack(padx=10, pady=5)

    result = []
    
    def on_ok():
        try:
            width = int(width_entry.get())
            height = int(height_entry.get())
            if width > 0 and height > 0:
                result.append((width, height))
                size_dialog.destroy()
            else:
                messagebox.showerror("Invalid Size", "Width and height must be positive integers.")
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid integers for width and height.")

    def on_close():
        result.append(None)
        size_dialog.destroy()
        
    def on_cancel():
        result.append(None)
        size_dialog.destroy()    

    btn_frame = tk.Frame(size_dialog)
    btn_frame.pack(fill=tk.X, pady=10)
    
    btn_ok = tk.Button(btn_frame, text="Okay", command=on_ok)
    btn_ok.pack(side=tk.LEFT, padx=10)
    
    btn_cancel = tk.Button(btn_frame, text="Cancel", command=on_cancel)
    btn_cancel.pack(side=tk.RIGHT, padx=10)

    size_dialog.protocol("WM_DELETE_WINDOW", on_close)
    size_dialog.wait_window(size_dialog)
    return result[0] if result else None

def resize_image(image, new_size):
    return image.resize(new_size, Image.LANCZOS)

def recolor_image(image, colors):
    image = image.convert('RGB')
    data = np.array(image)
    
    def find_nearest_color(color):
        color_diffs = np.linalg.norm(colors - color, axis=1)
        return colors[np.argmin(color_diffs)]
    
    new_data = np.apply_along_axis(find_nearest_color, 2, data)
    return Image.fromarray(new_data.astype('uint8'), 'RGB')

def display_image(image, original_filepath, zoom_factor=11):
    result = []
    window = tk.Toplevel()
    window.title("Image Preview")
    
    zoomed_size = (image.width * zoom_factor, image.height * zoom_factor)
    zoomed_image = image.resize(zoomed_size, Image.NEAREST)

    canvas = tk.Canvas(window, width=zoomed_image.width, height=zoomed_image.height)
    canvas.pack()
    
    tk_image = ImageTk.PhotoImage(zoomed_image)
    canvas.create_image(0, 0, anchor=tk.NW, image=tk_image)
    
    def on_ok():
        new_filepath = save_image(image, original_filepath)
        result.append(new_filepath)
        window.destroy()
    
    def on_cancel():
        result.append(None)
        window.destroy()

    def on_close():
        result.append(None)
        window.destroy()

    btn_frame = tk.Frame(window)
    btn_frame.pack(fill=tk.X, pady=5)
    
    btn_ok = tk.Button(btn_frame, text="Okay", command=on_ok)
    btn_ok.pack(side=tk.LEFT, padx=5)
    
    btn_cancel = tk.Button(btn_frame, text="Cancel", command=on_cancel)
    btn_cancel.pack(side=tk.RIGHT, padx=5)
    
    window.protocol("WM_DELETE_WINDOW", on_cancel)
    window.wait_window(window)

    return result[0] if result else None

def save_image(image, original_filepath):
    directory, filename = original_filepath.rsplit('/', 1)
    new_filename = f"{filename.rsplit('.', 1)[0]}_legomosaic.png"
    new_filepath = f"{directory}/{new_filename}"
    image.save(new_filepath)
    return new_filepath

def rgb_to_hex(rgb):
    return '#{:02x}{:02x}{:02x}'.format(*rgb)

def create_excel_from_image(image_path, excel_path):
    img = Image.open(image_path)
    img = img.convert('RGB')
    pixels = img.load()

    workbook = xlsxwriter.Workbook(excel_path)
    worksheet = workbook.add_worksheet('Raw')
    worksheet1 = workbook.add_worksheet('Conversion')
    worksheet2 = workbook.add_worksheet('Count')

    worksheet2.write(0, 0, 'Code Number')
    worksheet2.write(0, 1, 'Lego.com Index')
    worksheet2.write(0, 2, 'HTML Color Notation')
    worksheet2.write(0, 3, 'Count')
    worksheet2.write(1, 1, "307001")
    worksheet2.write(1, 2, "d9dadc")
    worksheet2.write(2, 1, "307021")
    worksheet2.write(2, 2, "be0606")
    worksheet2.write(3, 1, "307024")
    worksheet2.write(3, 2, "f6c500")
    worksheet2.write(4, 1, "307026")
    worksheet2.write(4, 2, "303030")
    worksheet2.write(5, 1, "4125253")
    worksheet2.write(5, 2, "b6a36f")
    worksheet2.write(6, 1, "4206330")
    worksheet2.write(6, 2, "2562b3")
    worksheet2.write(7, 1, "4210848")
    worksheet2.write(7, 2, "666666")
    worksheet2.write(8, 1, "4211288")
    worksheet2.write(8, 2, "582b17")
    worksheet2.write(9, 1, "4211415")
    worksheet2.write(9, 2, "939393")
    worksheet2.write(10, 1, "4527526")
    worksheet2.write(10, 2, "769ace")
    worksheet2.write(11, 1, "4537251")
    worksheet2.write(11, 2, "98bf43")
    worksheet2.write(12, 1, "4550169")
    worksheet2.write(12, 2, "6f2732")
    worksheet2.write(13, 1, "4558593")
    worksheet2.write(13, 2, "30a251")
    worksheet2.write(14, 1, "4558595")
    worksheet2.write(14, 2, "dd7e2a")
    worksheet2.write(15, 1, "4631385")
    worksheet2.write(15, 2, "253e67")
    worksheet2.write(16, 1, "4655243")
    worksheet2.write(16, 2, "61a3ba")
    worksheet2.write(17, 1, "6055171")
    worksheet2.write(17, 2, "35533b")
    worksheet2.write(18, 1, "6055172")
    worksheet2.write(18, 2, "887964")
    worksheet2.write(19, 1, "6065504")
    worksheet2.write(19, 2, "f39d30")
    worksheet2.write(20, 1, "6097301")
    worksheet2.write(20, 2, "8d62a3")
    worksheet2.write(21, 1, "6099364")
    worksheet2.write(21, 2, "8e3c7c")
    worksheet2.write(22, 1, "6133726")
    worksheet2.write(22, 2, "d837a1")
    worksheet2.write(23, 1, "6138232")
    worksheet2.write(23, 2, "7b858e")
    worksheet2.write(24, 1, "6143431")
    worksheet2.write(24, 2, "8a4e31")
    worksheet2.write(25, 1, "6151658")
    worksheet2.write(25, 2, "3685af")
    worksheet2.write(26, 1, "6167457")
    worksheet2.write(26, 2, "5c4394")
    worksheet2.write(27, 1, "6172375")
    worksheet2.write(27, 2, "59af44")
    worksheet2.write(28, 1, "6177146")
    worksheet2.write(28, 2, "856043")
    worksheet2.write(29, 1, "6211403")
    worksheet2.write(29, 2, "b48fc2")
    worksheet2.write(30, 1, "6213782")
    worksheet2.write(30, 2, "1e8f8b")
    worksheet2.write(31, 1, "6223913")
    worksheet2.write(31, 2, "769e84")
    worksheet2.write(32, 1, "6251846")
    worksheet2.write(32, 2, "bad3cd")
    worksheet2.write(33, 1, "6251940")
    worksheet2.write(33, 2, "ef97c7")
    worksheet2.write(34, 1, "6275876")
    worksheet2.write(34, 2, "fe6488")
    worksheet2.write(35, 1, "6275877")
    worksheet2.write(35, 2, "fde26f")
    worksheet2.write(36, 1, "6304896")
    worksheet2.write(36, 2, "bbc882")
    worksheet2.write(37, 1, "6316569")
    worksheet2.write(37, 2, "9fb9dd")
    worksheet2.write(38, 1, "6376232")
    worksheet2.write(38, 2, "dfdf05")
    worksheet2.write(39, 1, "6419170")
    worksheet2.write(39, 2, "c0835c")
    worksheet2.write(40, 1, "6475042")
    worksheet2.write(40, 2, "b75a17")

    for y in range(40):
        worksheet2.write(y+1, 0, str(y+1))
        worksheet2.write_formula(y+1, 3, '=COUNTIF(Raw!$A$1:' + str(xl_rowcol_to_cell(img.height,img.width)) + ',"#" & Count!' + str(xl_rowcol_to_cell(y+1,2)) + ')')

    for y in range(img.height):
        for x in range(img.width):
            rgb = pixels[x, y]  # Get RGB (ignore alpha if present)
            hex_color = rgb_to_hex(rgb)
            worksheet.write(y, x, hex_color)
            worksheet1.write_formula(y, x, '=XLOOKUP(SUBSTITUTE(Raw!' + str(xl_rowcol_to_cell(y,x)) + ',"#",""),Count!$C$2:Count!$C$41,Count!$A$2:$A$41)')
    workbook.close()

def show_completion_message(directory):
    msg_dialog = tk.Toplevel(root)
    msg_dialog.title("Process Complete")
    
    msg = f"Your image and Excel file have been created and are stored in the directory:\n{directory}"
    tk.Label(msg_dialog, text=msg, padx=20, pady=20).pack()

    def close_app():
        msg_dialog.destroy()
        root.quit()
        root.destroy()
    
    btn_ok = tk.Button(msg_dialog, text="Okay", command=close_app)
    btn_ok.pack(pady=10)
    
    msg_dialog.protocol("WM_DELETE_WINDOW", close_app)
    msg_dialog.wait_window(msg_dialog)

def show_error_message(error_message):
    error_dialog = tk.Toplevel(root)
    error_dialog.title("Error")
    
    tk.Label(error_dialog, text=error_message, padx=20, pady=20).pack()

    def close_app():
        error_dialog.destroy()
        root.quit()
        root.destroy()
    
    btn_ok = tk.Button(error_dialog, text="Okay", command=close_app)
    btn_ok.pack(pady=10)
    
    error_dialog.protocol("WM_DELETE_WINDOW", close_app)
    error_dialog.wait_window(error_dialog)

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    
    color_palette = np.array([
        [217, 218, 220], [190, 6, 6], [246, 197, 0], [48, 48, 48], [182, 163, 111],
        [37, 98, 179], [102, 102, 102], [88, 43, 23], [147, 147, 147], [118, 154, 206],
        [152, 191, 67], [111, 39, 50], [48, 162, 81], [221, 126, 42], [37, 62, 103],
        [97, 163, 186], [53, 83, 59], [136, 121, 100], [243, 157, 48], [141, 98, 163],
        [142, 60, 124], [216, 55, 161], [123, 133, 142], [138, 78, 49], [54, 133, 175],
        [92, 67, 148], [89, 175, 68], [133, 96, 67], [180, 143, 194], [30, 143, 139],
        [118, 158, 132], [186, 211, 205], [239, 151, 199], [254, 100, 136], [253, 226, 111],
        [187, 200, 130], [159, 185, 221], [223, 223, 5], [192, 131, 92], [183, 90, 23]
    ])
    
    new_image_path = open_image()
    if new_image_path:
        excel_path = filedialog.asksaveasfilename(defaultextension = ".xlsx", filetypes = (("Excel Workbook", "*.xlsx"),("All Files", "*.*")))
        if excel_path:
            create_excel_from_image(new_image_path, excel_path)
            directory = os.path.dirname(new_image_path)
            show_completion_message(directory)
        else:
           show_error_message("Error: No excel file was created.")
    else:
        show_error_message("Error: No image was created.")


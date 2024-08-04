import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import cv2
import os
from PIL import Image, ImageTk
from io import BytesIO
import win32clipboard
import tempfile

class VideoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("動画から画像抽出")
        self.root.geometry("960x640")
        self.root.resizable(False, False)  # 画面サイズを固定
        # フォント設定
        self.font = ("Yu Gothic", 12)

        self.video_path = ""
        self.capture = None
        self.frame_rate = 0
        self.total_frames = 0
        self.duration = 0
        self.current_frame = 0
        self.video_width = 0
        self.video_height = 0

        # 小数点以下の値を保持する変数
        self.slider_value = tk.DoubleVar()

        # UI構築
        self.create_ui()

    def create_ui(self):
        frame_top = ctk.CTkFrame(self.root)
        frame_top.pack(pady=10)

        self.open_button = ctk.CTkButton(frame_top, text="動画を開く", command=self.load_video, font=self.font)
        self.open_button.grid(row=0, column=0, padx=5)

        self.video_label = ctk.CTkLabel(frame_top, text="", font=self.font)
        self.video_label.grid(row=0, column=1, padx=5)

        frame_middle = ctk.CTkFrame(self.root)
        frame_middle.pack(pady=10)

        self.scale = ctk.CTkSlider(frame_middle, from_=0, to=100, variable=self.slider_value, number_of_steps=100000, command=self.update_preview)
        self.scale.grid(row=1, column=0, padx=5)

        self.time_entry = ctk.CTkEntry(frame_middle, font=self.font)
        self.time_entry.grid(row=1, column=1, padx=5)

        self.preview_label = ctk.CTkLabel(frame_middle, text="プレビュー", font=self.font)
        self.preview_label.grid(row=2, column=0, columnspan=2, pady=10)

        self.preview_canvas = tk.Canvas(frame_middle, width=640, height=360)
        self.preview_canvas.grid(row=3, column=0, columnspan=2, pady=10)

        self.preview_canvas.bind("<Button-3>", self.copy_to_clipboard)

        frame_bottom = ctk.CTkFrame(self.root)
        frame_bottom.pack(pady=10)

        self.save_button = ctk.CTkButton(frame_bottom, text="画像を保存", command=self.save_image, font=self.font)
        self.save_button.pack(pady=10)

    def load_video(self):
        self.video_path = filedialog.askopenfilename(filetypes=[("Video files", "*.mp4 *.avi *.mov")])
        if not self.video_path:
            return

        self.capture = cv2.VideoCapture(self.video_path)
        self.frame_rate = self.capture.get(cv2.CAP_PROP_FPS)
        self.total_frames = int(self.capture.get(cv2.CAP_PROP_FRAME_COUNT))
        self.duration = self.total_frames / self.frame_rate
        self.video_width = int(self.capture.get(cv2.CAP_PROP_FRAME_WIDTH))
        self.video_height = int(self.capture.get(cv2.CAP_PROP_FRAME_HEIGHT))

        self.video_label.configure(text=f"動画: {os.path.basename(self.video_path)} | 長さ: {self.duration:.3f}秒 | 解像度: {self.video_width}x{self.video_height}")
        self.scale.configure(to=self.duration)
        self.update_preview(0)

    def update_preview(self, val):
        if self.capture:
            self.current_frame = int(float(val) * self.frame_rate)
            self.capture.set(cv2.CAP_PROP_POS_FRAMES, self.current_frame)
            ret, frame = self.capture.read()
            if ret:
                self.show_frame(frame)
                self.time_entry.delete(0, tk.END)
                self.time_entry.insert(0, f"{float(val):.3f}秒")

    def show_frame(self, frame):
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(frame)
        img = img.resize((640, 360), Image.Resampling.LANCZOS)
        imgtk = ImageTk.PhotoImage(image=img)
        self.preview_canvas.create_image(0, 0, anchor=tk.NW, image=imgtk)
        self.preview_canvas.image = imgtk

    def save_image(self):
        if not self.video_path:
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png"), ("JPEG files", "*.jpg")])
        if not save_path:
            return

        time_val = self.time_entry.get()
        try:
            seconds = float(time_val.replace("秒", ""))
            self.current_frame = int(seconds * self.frame_rate)
        except ValueError:
            messagebox.showerror("エラー", "有効な秒数を入力してください")
            return

        self.capture.set(cv2.CAP_PROP_POS_FRAMES, self.current_frame)
        ret, frame = self.capture.read()
        if ret:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_file:
                tmp_path = tmp_file.name
                cv2.imwrite(tmp_path, frame)

            # Use os.replace to ensure atomic operation
            try:
                os.replace(tmp_path, save_path)
                messagebox.showinfo("完了", "画像が保存されました")
            except OSError as e:
                messagebox.showerror("エラー", f"画像の保存に失敗しました: {e}")
                os.remove(tmp_path)
        else:
            messagebox.showerror("エラー", "画像の保存に失敗しました")

    def copy_to_clipboard(self, event=None):
        if not self.capture:
            return

        time_val = self.time_entry.get()
        try:
            seconds = float(time_val.replace("秒", ""))
            self.current_frame = int(seconds * self.frame_rate)
        except ValueError:
            messagebox.showerror("エラー", "有効な秒数を入力してください")
            return

        self.capture.set(cv2.CAP_PROP_POS_FRAMES, self.current_frame)
        ret, frame = self.capture.read()
        if ret:
            # Create a temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix=".bmp") as tmp_file:
                tmp_path = tmp_file.name
                cv2.imwrite(tmp_path, frame)

            image = Image.open(tmp_path)
            output = BytesIO()
            image.convert("RGB").save(output, "BMP")
            data = output.getvalue()
            output.close()

            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data[14:])
            win32clipboard.CloseClipboard()

            os.remove(tmp_path)
            messagebox.showinfo("完了", "画像がクリップボードにコピーされました")

if __name__ == "__main__":
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    app = VideoApp(root)
    root.mainloop()

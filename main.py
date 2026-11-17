import customtkinter as ctk
from ui.main_window import MainWindow


def main() -> None:
    # 設定暗黑質感外觀模式與深藍色調主題
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()

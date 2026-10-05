import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

class ExcelViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("채용 및 현장 현황 대시보드")
        self.root.geometry("1000x650")

        # 1. 고정 3개 시트 이름 설정
        self.target_sheets = ["26년 채용비율", "반도체현장", "예전채용비율"]
        self.excel_path = None

        self._create_widgets()

        # 같은 폴더에 '채용비율.xlsx' 파일이 있으면 자동 로드
        default_file = "채용비율.xlsx"
        if os.path.exists(default_file):
            self.load_excel(default_file)

    def _create_widgets(self):
        # 상단 컨트롤 프레임 (파일 선택 버튼 및 상태)
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(fill=tk.X)

        btn_load = ttk.Button(top_frame, text="📁 엑셀 파일 불러오기", command=self.browse_file)
        btn_load.pack(side=tk.LEFT, padx=5)

        self.lbl_status = ttk.Label(top_frame, text="파일을 선택해주세요.", font=("맑은 고딕", 10))
        self.lbl_status.pack(side=tk.LEFT, padx=10)

        # 2. 시트 탭 (Notebook) 생성
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_frames = {}
        self.tree_views = {}

        for sheet_name in self.target_sheets:
            frame = ttk.Frame(self.notebook, padding=5)
            self.notebook.add(frame, text=f"  {sheet_name}  ")
            self.tab_frames[sheet_name] = frame

            # 각 탭에 테이블(Treeview) 및 스크롤바 배치
            tree_frame = ttk.Frame(frame)
            tree_frame.pack(fill=tk.BOTH, expand=True)

            vsb = ttk.Scrollbar(tree_frame, orient="vertical")
            hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

            tree = ttk.Treeview(
                tree_frame,
                yscrollcommand=vsb.set,
                xscrollcommand=hsb.set,
                selectmode="extended"
            )

            vsb.config(command=tree.yview)
            hsb.config(command=tree.xview)

            vsb.pack(side=tk.RIGHT, fill=tk.Y)
            hsb.pack(side=tk.BOTTOM, fill=tk.X)
            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            self.tree_views[sheet_name] = tree

    def browse_file(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if filepath:
            self.load_excel(filepath)

    def load_excel(self, filepath):
        self.excel_path = filepath
        self.lbl_status.config(text=f"현재 로드된 파일: {os.path.basename(filepath)}")

        try:
            excel_file = pd.ExcelFile(filepath)
            available_sheets = excel_file.sheet_names

            # 각 지정 탭에 데이터 매핑
            for sheet_name in self.target_sheets:
                tree = self.tree_views[sheet_name]
                tree.delete(*tree.get_children())
                tree["columns"] = []

                if sheet_name in available_sheets:
                    df = pd.read_excel(filepath, sheet_name=sheet_name)
                    df = df.fillna("")  # 빈값 처리

                    cols = list(df.columns)
                    tree["columns"] = cols
                    tree["show"] = "headings"

                    for col in cols:
                        tree.heading(col, text=str(col))
                        tree.column(col, width=120, anchor="center")

                    for _, row in df.iterrows():
                        tree.insert("", tk.END, values=list(row))
                else:
                    # 파일 내에 해당 이름의 시트가 없을 때 예외 안내
                    tree["columns"] = ["안내"]
                    tree["show"] = "headings"
                    tree.heading("안내", text="시트 정보")
                    tree.column("안내", width=300, anchor="center")
                    tree.insert("", tk.END, values=[f"'{sheet_name}' 시트가 파일에 존재하지 않습니다."])

        except Exception as e:
            messagebox.showerror("오류", f"엑셀 파일을 읽는 중 오류가 발생했습니다:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewerApp(root)
    root.mainloop()

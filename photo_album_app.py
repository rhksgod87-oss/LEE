import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from PIL import Image, ImageTk, ImageOps

import album_core as core
import auto_classify

THUMB_SIZE = (150, 112)


class ThumbButton(tk.Frame):
    def __init__(self, master, path, on_toggle):
        super().__init__(master, bd=3, relief="flat", bg="white")
        self.path = path
        self.selected = False
        self.on_toggle = on_toggle

        img = ImageOps.exif_transpose(Image.open(path))
        img.thumbnail(THUMB_SIZE)
        self.photo = ImageTk.PhotoImage(img)

        self.img_label = tk.Label(self, image=self.photo, bd=0, bg="white")
        self.img_label.pack()
        self.name_label = tk.Label(
            self, text=os.path.basename(path), font=("Malgun Gothic", 8),
            wraplength=150, bg="white"
        )
        self.name_label.pack(fill="x")

        for w in (self, self.img_label, self.name_label):
            w.bind("<Button-1>", self.toggle)

    def toggle(self, event=None):
        self.set_selected(not self.selected)
        self.on_toggle(self)

    def set_selected(self, val):
        self.selected = val
        color = "#2d7dd2" if val else "white"
        fg = "white" if val else "black"
        self.config(bg=color)
        self.img_label.config(bg=color)
        self.name_label.config(bg=color, fg=fg)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("사진대지 자동 생성기")
        self.geometry("1150x760")

        self.template_path = tk.StringVar()
        self.photos_root = tk.StringVar()
        self.output_path = tk.StringVar()
        self.project_name = tk.StringVar()
        self.gongjong = tk.StringVar(value="포장공")
        self.owner = tk.StringVar()

        self.stage_gongjong_var = tk.StringVar()
        self.stage_count_var = tk.IntVar(value=2)

        self.output_path_is_auto = True  # 저장 위치를 직접 지정하기 전까지는 공사명을 따라감
        self.project_name.trace_add("write", self._on_project_name_change)

        self.folders = []
        self.folder_order = []          # [base_name, ...] 사용자가 지정한 처리 순서
        self.sub_stages = {}            # base_name -> [세부단계 이름, ...] (순서 있음)
        # 아래 3개 dict는 "단계 키"로 관리: 하위단계 없는 폴더는 base_name 그대로,
        # 하위단계는 "base_name::세부단계이름" 형태의 키를 씀.
        self.folder_manual_picks = {}   # 단계키 -> [선택된 사진 절대경로, ...] (최대 stage_count개)
        self.folder_gongjong = {}       # 단계키 -> 공종 문자열 (자동추정 또는 직접수정)
        self.folder_photo_count = {}    # 단계키 -> 2 또는 4 (대표사진 장수)
        self.thumb_widgets = {}         # path -> ThumbButton (현재 보이는 단계 것만)
        self.current_stage_key = None
        self.current_source_folder = None

        self._build_ui()

    # ---------------------------------------------------------- UI 구성
    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        def path_row(label, var, cmd, r):
            ttk.Label(top, text=label, width=10).grid(row=r, column=0, sticky="w", pady=3)
            ttk.Entry(top, textvariable=var, width=75).grid(row=r, column=1, padx=5)
            ttk.Button(top, text="찾아보기", command=cmd).grid(row=r, column=2)

        path_row("템플릿 파일", self.template_path, self.browse_template, 0)
        path_row("사진 폴더", self.photos_root, self.browse_photos_root, 1)
        path_row("저장 위치", self.output_path, self.browse_output, 2)

        info = ttk.Frame(self, padding=(10, 0, 10, 5))
        info.pack(fill="x")
        ttk.Label(info, text="공사명").grid(row=0, column=0, sticky="w")
        ttk.Entry(info, textvariable=self.project_name, width=38).grid(row=0, column=1, padx=5)
        ttk.Label(info, text="기본 공종").grid(row=0, column=2, sticky="w")
        ttk.Entry(info, textvariable=self.gongjong, width=12).grid(row=0, column=3, padx=5)
        ttk.Label(info, text="발주처").grid(row=0, column=4, sticky="w")
        ttk.Entry(info, textvariable=self.owner, width=14).grid(row=0, column=5, padx=5)
        ttk.Label(
            info, text="(단계별 공종은 오른쪽에서 폴더 선택 시 자동 추정되며, 추정이 안 될 때 이 값이 쓰입니다)",
            foreground="#666666"
        ).grid(row=1, column=0, columnspan=6, sticky="w", pady=(2, 0))

        ttk.Separator(self).pack(fill="x", padx=10)

        main = ttk.Frame(self)
        main.pack(fill="both", expand=True, padx=10, pady=8)

        left = ttk.Frame(main, width=260)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        ttk.Label(left, text="공정 단계").pack(anchor="w")
        ttk.Label(
            left,
            text="(폴더명에 번호 없이 이름만 써도 됩니다. 폴더를 선택하고\n"
                 "'+하위단계'로 폴더 안 사진을 세부 단계로 나눌 수 있습니다.\n"
                 "하위단계마다 기준이 될 사진을 먼저 클릭해두면(4장 추천),\n"
                 "'자동 분류'로 나머지 사진을 생김새 기준으로 나눠줍니다.\n"
                 "완벽하지 않으니 결과는 꼭 확인하세요.)",
            foreground="#666666", font=("Malgun Gothic", 8), justify="left"
        ).pack(anchor="w")

        list_row = ttk.Frame(left)
        list_row.pack(fill="both", expand=True, pady=4)
        self.folder_tree = ttk.Treeview(list_row, show="tree", selectmode="browse")
        self.folder_tree.pack(side="left", fill="both", expand=True)
        self.folder_tree.bind("<<TreeviewSelect>>", self.on_stage_select)

        order_btns = ttk.Frame(list_row)
        order_btns.pack(side="left", fill="y", padx=(4, 0))
        ttk.Button(order_btns, text="▲ 위로", command=self.move_stage_up, width=10).pack(pady=2)
        ttk.Button(order_btns, text="▼ 아래로", command=self.move_stage_down, width=10).pack(pady=2)
        ttk.Separator(order_btns, orient="horizontal").pack(fill="x", pady=6)
        ttk.Button(order_btns, text="+ 하위단계", command=self.add_sub_stage, width=10).pack(pady=2)
        ttk.Button(order_btns, text="- 하위단계 삭제", command=self.remove_sub_stage, width=10).pack(pady=2)
        ttk.Separator(order_btns, orient="horizontal").pack(fill="x", pady=6)
        self.auto_classify_btn = ttk.Button(
            order_btns, text="자동 분류", command=self.start_auto_classify, width=10
        )
        self.auto_classify_btn.pack(pady=2)

        right = ttk.Frame(main)
        right.pack(side="left", fill="both", expand=True, padx=(10, 0))

        gj_row = ttk.Frame(right)
        gj_row.pack(anchor="w", fill="x", pady=(0, 6))
        ttk.Label(gj_row, text="이 단계 공종 (자동 추정됨 - 틀리면 직접 수정하세요):").pack(side="left")
        ttk.Entry(gj_row, textvariable=self.stage_gongjong_var, width=20).pack(side="left", padx=5)

        ttk.Label(gj_row, text="   대표사진 장수:").pack(side="left", padx=(15, 0))
        ttk.Radiobutton(
            gj_row, text="2장", variable=self.stage_count_var, value=2,
            command=self.on_stage_count_change
        ).pack(side="left")
        ttk.Radiobutton(
            gj_row, text="4장", variable=self.stage_count_var, value=4,
            command=self.on_stage_count_change
        ).pack(side="left")

        self.thumb_hint = ttk.Label(
            right,
            text="대표사진으로 쓸 사진을 클릭해서 선택하세요 (최대 2장, 선택 안 하면 자동으로 선택됩니다)"
        )
        self.thumb_hint.pack(anchor="w")

        canvas_frame = ttk.Frame(right)
        canvas_frame.pack(fill="both", expand=True, pady=4)
        self.canvas = tk.Canvas(canvas_frame, bg="#f0f0f0", highlightthickness=0)
        vsb = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.thumb_frame = tk.Frame(self.canvas, bg="#f0f0f0")
        self.canvas_window = self.canvas.create_window((0, 0), window=self.thumb_frame, anchor="nw")
        self.thumb_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)

        bottom = ttk.Frame(self, padding=10)
        bottom.pack(fill="x")
        self.generate_btn = ttk.Button(bottom, text="사진대지 생성하기", command=self.on_generate)
        self.generate_btn.pack(side="left")
        ttk.Button(bottom, text="결과 파일 열기", command=self.open_output).pack(side="left", padx=5)

        self.log_text = tk.Text(self, height=8)
        self.log_text.pack(fill="x", padx=10, pady=(0, 10))

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-event.delta / 120), "units")

    # ---------------------------------------------------------- 경로 선택
    def browse_template(self):
        p = filedialog.askopenfilename(title="템플릿 엑셀 파일 선택", filetypes=[("Excel 파일", "*.xlsx")])
        if p:
            self.template_path.set(p)

    def browse_photos_root(self):
        p = filedialog.askdirectory(title="단계별 사진 폴더들이 들어있는 폴더 선택")
        if p:
            self.photos_root.set(p)
            self.folder_manual_picks = {}
            self.folder_gongjong = {}
            self.folder_photo_count = {}
            self.folder_order = []
            self.sub_stages = {}
            self.current_stage_key = None
            self.current_source_folder = None
            self.load_folders()
            if self.output_path_is_auto:
                self._update_auto_output_path()

    def browse_output(self):
        p = filedialog.asksaveasfilename(
            title="저장할 파일 위치", defaultextension=".xlsx", filetypes=[("Excel 파일", "*.xlsx")]
        )
        if p:
            self.output_path.set(p)
            self.output_path_is_auto = False  # 직접 지정했으니 이후 공사명이 바뀌어도 안 따라감

    def _on_project_name_change(self, *args):
        if self.output_path_is_auto:
            self._update_auto_output_path()

    def _update_auto_output_path(self):
        """저장 위치를 직접 지정하지 않았다면, 파일명을 공사명으로 자동으로 맞춰줌."""
        if not self.photos_root.get():
            return
        name = core.sanitize_filename(self.project_name.get(), fallback="output")
        self.output_path.set(os.path.join(self.photos_root.get(), f"{name}.xlsx"))

    # ---------------------------------------------------------- 폴더/단계 트리
    def load_folders(self):
        self.folders = core.list_stage_folders(self.photos_root.get(), order=self.folder_order or None)
        self.folder_order = [os.path.basename(f) for f in self.folders]
        self._refresh_tree()
        children = self.folder_tree.get_children("")
        if children:
            self.folder_tree.selection_set(children[0])
            self.folder_tree.focus(children[0])

    def _refresh_tree(self, select_iid=None):
        self.folder_tree.delete(*self.folder_tree.get_children())
        for f in self.folders:
            base = os.path.basename(f)
            n = len(core.list_photos(f))
            subs = self.sub_stages.get(base, [])
            label = f"{base} ({n}장)" if not subs else f"{base} ({n}장, 하위단계 {len(subs)}개)"
            self.folder_tree.insert("", "end", iid=base, text=label, open=True)
            for name in subs:
                self.folder_tree.insert(base, "end", iid=f"{base}::{name}", text=f"    ↳ {name}")
        if select_iid and self.folder_tree.exists(select_iid):
            self.folder_tree.selection_set(select_iid)
            self.folder_tree.focus(select_iid)

    def _stage_info(self, iid):
        """트리 항목 id로부터 (단계키, 사진이 들어있는 실제 폴더, 기본 세부설명)을 계산."""
        if "::" in iid:
            base, name = iid.split("::", 1)
            folder = next((f for f in self.folders if os.path.basename(f) == base), None)
            return iid, folder, " ".join(list(name))
        folder = next((f for f in self.folders if os.path.basename(f) == iid), None)
        return iid, folder, (core.folder_caption(folder) if folder else iid)

    def move_stage_up(self):
        self._move_stage(-1)

    def move_stage_down(self):
        self._move_stage(1)

    def _move_stage(self, direction):
        sel = self.folder_tree.selection()
        if not sel:
            return
        iid = sel[0]
        if "::" in iid:
            base, name = iid.split("::", 1)
            subs = self.sub_stages.get(base, [])
            if name not in subs:
                return
            i = subs.index(name)
            j = i + direction
            if 0 <= j < len(subs):
                subs[i], subs[j] = subs[j], subs[i]
                self._refresh_tree(select_iid=iid)
        else:
            i = next((idx for idx, f in enumerate(self.folders) if os.path.basename(f) == iid), None)
            if i is None:
                return
            j = i + direction
            if 0 <= j < len(self.folders):
                self.folders[i], self.folders[j] = self.folders[j], self.folders[i]
                self.folder_order = [os.path.basename(f) for f in self.folders]
                self._refresh_tree(select_iid=iid)

    def add_sub_stage(self):
        sel = self.folder_tree.selection()
        if not sel:
            messagebox.showinfo("안내", "하위 단계를 추가할 폴더를 먼저 선택하세요.")
            return
        base = sel[0].split("::")[0]  # 하위단계가 선택돼 있으면 그 부모 폴더 기준
        name = simpledialog.askstring(
            "하위 단계 추가",
            f"'{base}' 폴더 안 사진을 나눌 세부 단계 이름을 입력하세요\n(예: 노면절삭, 택코팅, 아스콘포설및다짐)",
            parent=self,
        )
        if not name:
            return
        name = name.strip()
        if not name:
            return
        subs = self.sub_stages.setdefault(base, [])
        if name in subs:
            messagebox.showwarning("안내", "이미 있는 이름입니다.")
            return
        subs.append(name)
        self._refresh_tree(select_iid=f"{base}::{name}")
        self.on_stage_select(None)

    def remove_sub_stage(self):
        sel = self.folder_tree.selection()
        if not sel or "::" not in sel[0]:
            messagebox.showinfo("안내", "삭제할 하위 단계를 목록에서 선택하세요.")
            return
        iid = sel[0]
        base, name = iid.split("::", 1)
        if not messagebox.askyesno("삭제 확인", f"'{name}' 하위 단계를 삭제할까요?\n(사진 파일 자체는 지워지지 않습니다)"):
            return
        subs = self.sub_stages.get(base, [])
        if name in subs:
            subs.remove(name)
        self.folder_manual_picks.pop(iid, None)
        self.folder_gongjong.pop(iid, None)
        self.folder_photo_count.pop(iid, None)
        self._refresh_tree(select_iid=base)
        self.on_stage_select(None)

    # ---------------------------------------------------------- 자동 분류 (오프라인)
    def start_auto_classify(self):
        sel = self.folder_tree.selection()
        if not sel:
            messagebox.showinfo("안내", "하위 단계가 있는 폴더나 하위 단계를 선택하세요.")
            return
        base = sel[0].split("::")[0]
        subs = self.sub_stages.get(base, [])
        if not subs:
            messagebox.showinfo("안내", f"'{base}' 폴더에 하위 단계를 먼저 만드세요 ('+ 하위단계').")
            return

        self._save_current_stage_settings()

        reference_photos = {}
        missing = []
        for name in subs:
            key = f"{base}::{name}"
            picks = self.folder_manual_picks.get(key)
            if picks:
                reference_photos[name] = list(picks)
            else:
                missing.append(name)
        if missing:
            messagebox.showinfo(
                "안내",
                "먼저 아래 하위 단계마다 대표사진을 클릭해서 선택해주세요"
                "(그 사진들을 기준으로 나머지를 분류합니다. 정확도를 높이려면"
                " '대표사진 장수'를 4장으로 하고 4장 다 클릭해두는 걸 추천합니다):"
                f"\n\n{', '.join(missing)}",
            )
            return

        folder = next((f for f in self.folders if os.path.basename(f) == base), None)
        if folder is None:
            return
        photos = core.list_photos(folder)
        if not photos:
            messagebox.showinfo("안내", f"'{base}' 폴더에 사진이 없습니다.")
            return

        if not messagebox.askyesno(
            "자동 분류",
            f"'{base}' 폴더의 사진 {len(photos)}장을 하위 단계 {subs} 로 자동 분류합니다.\n"
            f"인터넷/API 없이 이 PC에서만 처리됩니다(사전학습 이미지 인식 모델 사용).\n"
            f"완벽하지 않을 수 있으니 결과를 꼭 확인해주세요. 계속할까요?",
        ):
            return

        self.auto_classify_btn.config(state="disabled")
        self.log_text.delete("1.0", "end")
        self.log(f"[자동 분류] '{base}' 사진 {len(photos)}장 -> {subs}")
        threading.Thread(
            target=self._run_auto_classify,
            args=(photos, reference_photos, subs, base),
            daemon=True,
        ).start()

    def _run_auto_classify(self, photos, reference_photos, subs, base):
        def thread_log(msg):
            self.after(0, self.log, msg)

        try:
            assignments = auto_classify.classify_photos(photos, reference_photos, log=thread_log)
        except Exception as e:
            self.after(0, self._auto_classify_failed, str(e))
            return
        self.after(0, self._auto_classify_done, assignments, subs, base)

    def _auto_classify_failed(self, message):
        self.auto_classify_btn.config(state="normal")
        self.log(f"\n자동 분류 실패: {message}")
        messagebox.showerror("오류", f"자동 분류 중 문제가 발생했습니다:\n{message}")

    def _auto_classify_done(self, assignments, subs, base):
        groups = {name: [] for name in subs}
        unknown = []
        for path, name in assignments.items():
            if name in groups:
                groups[name].append(path)
            else:
                unknown.append(path)

        for name, matched_photos in groups.items():
            key = f"{base}::{name}"
            count = self.folder_photo_count.get(key, 2)
            self.folder_manual_picks[key] = core.pick_representatives(matched_photos, n=count)

        self._refresh_tree()
        sel = self.folder_tree.selection()
        if sel:
            self.on_stage_select(None)

        self.auto_classify_btn.config(state="normal")
        summary = ", ".join(f"{name} {len(p)}장" for name, p in groups.items())
        self.log(f"\n완료: {summary}" + (f" | 미상 {len(unknown)}장" if unknown else ""))
        messagebox.showinfo(
            "자동 분류 완료",
            "자동 분류가 완료되어 각 하위단계에 대표사진이 자동 지정됐습니다.\n"
            "왼쪽에서 하위단계를 눌러 확인하고, 틀린 사진은 직접 클릭해서 바꿔주세요."
            + (f"\n\n분류가 안 된 사진 {len(unknown)}장은 어느 단계에도 배정되지 않았습니다." if unknown else ""),
        )

    def on_stage_select(self, event):
        self._save_current_stage_settings()

        sel = self.folder_tree.selection()
        if not sel:
            return
        key, folder, caption_default = self._stage_info(sel[0])
        if folder is None:
            return
        self.current_stage_key = key
        self.current_source_folder = folder

        if key in self.folder_gongjong:
            self.stage_gongjong_var.set(self.folder_gongjong[key])
        else:
            guess = core.guess_gongjong(caption_default, self.gongjong.get())
            self.stage_gongjong_var.set(guess)

        self.stage_count_var.set(self.folder_photo_count.get(key, 2))
        self._update_thumb_hint()

        self.show_thumbnails(folder, key)

    def _save_current_stage_settings(self):
        """단계를 바꾸거나 생성하기 전에, 현재 화면에 보이는 공종/장수 값을 저장해둠."""
        if self.current_stage_key is not None:
            self.folder_gongjong[self.current_stage_key] = self.stage_gongjong_var.get()
            self.folder_photo_count[self.current_stage_key] = self.stage_count_var.get()

    def on_stage_count_change(self):
        """2장/4장 라디오버튼을 바꿨을 때: 장수를 줄이면 이미 선택된 사진이
        새 장수보다 많을 경우 뒤쪽(나중에 선택한) 것부터 선택 해제함."""
        if self.current_stage_key is None:
            return
        count = self.stage_count_var.get()
        self.folder_photo_count[self.current_stage_key] = count

        picks = self.folder_manual_picks.get(self.current_stage_key, [])
        while len(picks) > count:
            removed = picks.pop()
            btn = self.thumb_widgets.get(removed)
            if btn:
                btn.set_selected(False)

        self._update_thumb_hint()

    def _update_thumb_hint(self):
        n = self.stage_count_var.get()
        self.thumb_hint.config(
            text=f"대표사진으로 쓸 사진을 클릭해서 선택하세요 (최대 {n}장, 선택 안 하면 자동으로 선택됩니다)"
        )

    def show_thumbnails(self, folder, key):
        for w in self.thumb_frame.winfo_children():
            w.destroy()
        self.thumb_widgets = {}

        photos = core.list_photos(folder)
        selected_paths = set(self.folder_manual_picks.get(key, []))

        cols = 6
        for idx, p in enumerate(photos):
            try:
                btn = ThumbButton(self.thumb_frame, p, self.on_thumb_toggle)
            except Exception:
                continue
            btn.grid(row=idx // cols, column=idx % cols, padx=4, pady=4)
            if p in selected_paths:
                btn.set_selected(True)
            self.thumb_widgets[p] = btn

    def on_thumb_toggle(self, btn):
        key = self.current_stage_key
        picks = self.folder_manual_picks.setdefault(key, [])
        max_count = self.stage_count_var.get()
        if btn.selected:
            if len(picks) >= max_count:
                oldest = picks.pop(0)
                old_btn = self.thumb_widgets.get(oldest)
                if old_btn:
                    old_btn.set_selected(False)
            picks.append(btn.path)
        else:
            if btn.path in picks:
                picks.remove(btn.path)

    # ---------------------------------------------------------- 로그
    def log(self, msg):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.update_idletasks()

    # ---------------------------------------------------------- 단계 목록 구성
    def _build_stage_list(self):
        """트리에 표시된 대로(폴더 순서 + 하위단계 유무)를 반영해 build_album에
        넘길 stage_list를 만듦. 하위단계가 있는 폴더는 그 폴더 자체 대신
        하위단계들이 그 자리에 펼쳐져 들어감."""
        stage_list = []
        for f in self.folders:
            base = os.path.basename(f)
            subs = self.sub_stages.get(base, [])
            if subs:
                for name in subs:
                    stage_list.append({
                        "key": f"{base}::{name}",
                        "folder": f,
                        "caption": " ".join(list(name)),
                    })
            else:
                stage_list.append({
                    "key": base,
                    "folder": f,
                    "caption": core.folder_caption(f),
                })
        return stage_list

    # ---------------------------------------------------------- 생성
    def on_generate(self):
        self.log_text.delete("1.0", "end")

        if not self.template_path.get():
            messagebox.showerror("오류", "템플릿 파일을 선택하세요.")
            return
        if not self.photos_root.get():
            messagebox.showerror("오류", "사진 폴더를 선택하세요.")
            return

        self._save_current_stage_settings()

        if not self.output_path.get():
            self._update_auto_output_path()
        output_path = self.output_path.get() or os.path.join(self.photos_root.get(), "output.xlsx")

        stage_list = self._build_stage_list()
        if not stage_list:
            messagebox.showerror("오류", "공정 단계 폴더가 없습니다.")
            return

        manual = {k: v for k, v in self.folder_manual_picks.items() if v}

        # 아직 한 번도 클릭 안 해본 단계도 자동추정/기본값을 채워넣는다.
        gongjong_by_folder = {}
        photo_count_by_folder = {}
        for stage in stage_list:
            key = stage["key"]
            val = self.folder_gongjong.get(key)
            if not val:
                val = core.guess_gongjong(stage["caption"], self.gongjong.get())
            gongjong_by_folder[key] = val
            photo_count_by_folder[key] = self.folder_photo_count.get(key, 2)

        self.generate_btn.config(state="disabled")
        try:
            items = core.build_album(
                template_path=self.template_path.get(),
                output_path=output_path,
                project_name=self.project_name.get(),
                gongjong=self.gongjong.get(),
                owner=self.owner.get(),
                photos_root=self.photos_root.get(),
                manual_picks=manual,
                gongjong_by_folder=gongjong_by_folder,
                photo_count_by_folder=photo_count_by_folder,
                stage_list=stage_list,
                log=self.log,
            )
            self.output_path.set(output_path)
            self.log(f"\n완료! 총 {len(items)}장 -> {output_path}")
            messagebox.showinfo("완료", f"사진대지 생성이 완료됐습니다.\n\n{output_path}")
        except Exception as e:
            self.log(f"\n오류: {e}")
            self.log(traceback.format_exc())
            messagebox.showerror("오류", str(e))
        finally:
            self.generate_btn.config(state="normal")

    def open_output(self):
        p = self.output_path.get()
        if p and os.path.isfile(p):
            os.startfile(p)
        else:
            messagebox.showwarning("알림", "아직 생성된 파일이 없습니다. 먼저 생성해주세요.")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

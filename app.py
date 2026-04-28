import html
import io
import re
import time
import urllib.parse
import urllib.request
import zipfile
from datetime import datetime
from pathlib import Path
from tempfile import mkdtemp

import streamlit as st
from PIL import Image
from pptx import Presentation

PIXEL_TO_EMU = 9525
USER_AGENT = "Mozilla/5.0"


def sanitize_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*]', "_", name).strip()
    return name or "converted_presentation"


def extract_file_id(url: str) -> str:
    m = re.search(r"/file/d/([a-zA-Z0-9_-]+)", url)
    if not m:
        raise ValueError("無法從連結中解析檔案 ID，請確認是 Google Drive 的 /file/d/... 連結。")
    return m.group(1)


def fetch_text(url: str, timeout: int = 45) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read().decode("utf-8", "ignore")


def parse_drive_view_metadata(drive_url: str) -> tuple[str, int, str, str]:
    file_id = extract_file_id(drive_url)
    view_url = f"https://drive.google.com/file/d/{file_id}/view"

    raw_html = fetch_text(view_url)

    title_match = re.search(r"<title>(.*?)</title>", raw_html, flags=re.IGNORECASE | re.DOTALL)
    title = html.unescape(title_match.group(1)).replace(" - Google 雲端硬碟", "").strip() if title_match else file_id

    page_matches = re.findall(r"共\s*(\d+)\s*頁", raw_html)
    if not page_matches:
        raise RuntimeError("無法解析總頁數，請確認檔案可公開檢視。")
    total_pages = max(int(x) for x in page_matches)

    upload_match = re.search(r"viewer/upload\?([^\"']+)", raw_html)
    if not upload_match:
        raise RuntimeError("無法從頁面取得 viewer/upload 參數。")

    upload_path = upload_match.group(0)
    upload_path = upload_path.encode("utf-8").decode("unicode_escape")
    upload_path = upload_path.replace("\\u003d", "=").replace("\\u0026", "&")
    upload_url = urllib.parse.urljoin("https://drive.google.com/", upload_path)

    upload_query = urllib.parse.parse_qs(urllib.parse.urlparse(upload_url).query)
    dsmi = upload_query.get("dsmi", ["texmex"])[0]

    upload_resp = fetch_text(upload_url)
    token_match = re.search(r"ACFr[0-9A-Za-z\-_=%%]+", upload_resp)
    if not token_match:
        raise RuntimeError("無法從 viewer/upload 回應中取得圖像 token。")
    token = token_match.group(0)

    return title, total_pages, token, dsmi


def download_slide_images(
    token: str,
    dsmi: str,
    total_pages: int,
    image_width: int,
    retries: int,
    output_dir: Path,
    log,
    update_progress,
) -> tuple[list[Path], tuple[int, int]]:
    output_dir.mkdir(parents=True, exist_ok=True)
    image_paths: list[Path] = []
    size_set: set[tuple[int, int]] = set()

    base = "https://drive.google.com/viewer/img"

    for idx in range(total_pages):
        params = {
            "id": token,
            "dsmi": dsmi,
            "auditContext": "forDisplay",
            "page": str(idx),
            "skiphighlight": "true",
            "w": str(image_width),
            "webp": "true",
        }
        url = base + "?" + urllib.parse.urlencode(params)

        data = None
        last_err = ""
        for attempt in range(1, retries + 1):
            try:
                req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
                with urllib.request.urlopen(req, timeout=45) as resp:
                    data = resp.read()
                break
            except Exception as err:  # noqa: BLE001
                last_err = str(err)
                time.sleep(0.6 * attempt)

        if data is None:
            raise RuntimeError(f"第 {idx + 1} 頁下載失敗: {last_err}")

        image = Image.open(io.BytesIO(data)).convert("RGB")
        size_set.add(image.size)
        out_path = output_dir / f"slide_{idx + 1:03d}.png"
        image.save(out_path, "PNG", optimize=True)
        image_paths.append(out_path)

        update_progress(idx + 1, total_pages)
        if (idx + 1) % 5 == 0 or (idx + 1) == total_pages:
            log(f"已下載 {idx + 1}/{total_pages} 頁")

    if len(size_set) != 1:
        raise RuntimeError(f"下載頁面尺寸不一致: {sorted(size_set)}")

    return image_paths, next(iter(size_set))


def build_pptx_from_images(
    image_paths: list[Path],
    size: tuple[int, int],
    output_path: Path,
    log,
    update_progress,
) -> Path:
    width, height = size
    prs = Presentation()
    prs.slide_width = width * PIXEL_TO_EMU
    prs.slide_height = height * PIXEL_TO_EMU
    blank = prs.slide_layouts[6]

    total = len(image_paths)
    for idx, image_path in enumerate(image_paths, 1):
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(
            str(image_path),
            left=0,
            top=0,
            width=prs.slide_width,
            height=prs.slide_height,
        )

        update_progress(idx, total)
        if idx % 5 == 0 or idx == total:
            log(f"已建立投影片 {idx}/{total}")

    try:
        prs.save(str(output_path))
        return output_path
    except PermissionError:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback = output_path.with_stem(output_path.stem + f"_fixed_{ts}")
        prs.save(str(fallback))
        log(f"原檔被占用，已改存為: {fallback.name}")
        return fallback


def build_output_name(naming_mode: str, title: str, custom_name: str) -> str:
    if naming_mode == "沿用原始檔名":
        return sanitize_filename(title) + ".pptx"
    if naming_mode == "原始檔名加日期":
        ts = datetime.now().strftime("%Y%m%d")
        return sanitize_filename(f"{title}_{ts}") + ".pptx"
    if not custom_name.strip():
        raise ValueError("請輸入自訂檔名。")
    return sanitize_filename(custom_name.strip()) + ".pptx"


def history_to_csv_bytes(records: list[dict]) -> bytes:
    headers = [
        "時間",
        "狀態",
        "來源連結",
        "標題",
        "輸出檔名",
        "頁數",
        "尺寸",
        "檔案大小MB",
        "訊息",
    ]
    lines = [",".join(headers)]
    for r in records:
        row = [
            str(r.get("時間", "")),
            str(r.get("狀態", "")),
            str(r.get("來源連結", "")),
            str(r.get("標題", "")),
            str(r.get("輸出檔名", "")),
            str(r.get("頁數", "")),
            str(r.get("尺寸", "")),
            str(r.get("檔案大小MB", "")),
            str(r.get("訊息", "")),
        ]
        escaped = []
        for cell in row:
            if any(ch in cell for ch in [",", '"', "\n"]):
                cell = '"' + cell.replace('"', '""') + '"'
            escaped.append(cell)
        lines.append(",".join(escaped))
    return ("\n".join(lines)).encode("utf-8-sig")


def bundle_zip_bytes(items: list[tuple[str, bytes]]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in items:
            zf.writestr(name, data)
    return buf.getvalue()


st.set_page_config(page_title="Drive 簡報轉換器", page_icon="📎", layout="wide")

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap');
html, body, [class*="css"] { font-family: 'Noto Sans TC', sans-serif; }
.gov-card {
  border: 1px solid #D7DEE3;
  border-radius: 10px;
  padding: 14px;
  background: #FFFFFF;
}
.gov-kpi {
  border-left: 4px solid #0F4C5C;
  padding: 8px 12px;
  background: #F7FAFC;
  border-radius: 6px;
}
.small-note { color: #5F6B73; font-size: 0.9rem; }
</style>
""",
    unsafe_allow_html=True,
)

st.title("Google Drive 簡報轉換平台")
st.caption("專業公務風 | 公開可檢視連結轉換為 PPTX")

with st.sidebar:
    st.subheader("任務設定")
    mode = st.radio("作業模式", ["單一連結", "批次連結"], index=0)
    drive_url = ""
    batch_links_text = ""
    if mode == "單一連結":
        drive_url = st.text_input("Google Drive 連結", placeholder="https://drive.google.com/file/d/.../view")
    else:
        batch_links_text = st.text_area(
            "批次連結（每行一個）",
            placeholder="https://drive.google.com/file/d/.../view\nhttps://drive.google.com/file/d/.../view",
            height=150,
        )
    naming_mode = st.radio("輸出檔名規則", ["沿用原始檔名", "原始檔名加日期", "自訂名稱"], index=0)
    custom_name = st.text_input("自訂檔名", value="") if naming_mode == "自訂名稱" else ""
    image_width = st.selectbox("圖片寬度", [1600, 1920, 1280], index=0)
    retries = st.slider("重試次數", 1, 5, 3)
    save_local = st.checkbox("另存至伺服器工作目錄", value=True)
    start = st.button("開始轉換", type="primary", use_container_width=True)
    retry_failed = st.button("重試失敗項目", use_container_width=True)
    clear_history = st.button("清除任務歷程", use_container_width=True)

step_col1, step_col2, step_col3, step_col4 = st.columns(4)
step_col1.markdown("<div class='gov-kpi'>1. 解析連結</div>", unsafe_allow_html=True)
step_col2.markdown("<div class='gov-kpi'>2. 下載頁面</div>", unsafe_allow_html=True)
step_col3.markdown("<div class='gov-kpi'>3. 重建檔案</div>", unsafe_allow_html=True)
step_col4.markdown("<div class='gov-kpi'>4. 完成交付</div>", unsafe_allow_html=True)

status_box = st.empty()
download_progress = st.progress(0, text="下載進度: 0%")
build_progress = st.progress(0, text="建檔進度: 0%")
log_box = st.empty()

if "logs" not in st.session_state:
    st.session_state.logs = []
if "history" not in st.session_state:
    st.session_state.history = []

if clear_history:
    st.session_state.history = []


def add_log(message: str) -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    st.session_state.logs.append(f"[{ts}] {message}")
    log_box.code("\n".join(st.session_state.logs[-200:]))


run_mode = None
if start:
    run_mode = "start"
elif retry_failed:
    run_mode = "retry"

if run_mode:
    st.session_state.logs = []
    log_box.code("")
    download_progress.progress(0, text="下載進度: 0%")
    build_progress.progress(0, text="建檔進度: 0%")

    try:
        if run_mode == "start":
            if mode == "單一連結":
                links = [drive_url.strip()] if drive_url.strip() else []
            else:
                links = [x.strip() for x in batch_links_text.splitlines() if x.strip()]

            if not links:
                raise ValueError("請至少輸入一個 Google Drive 連結。")

            operation_name = "批次作業"
        else:
            links = []
            seen = set()
            for rec in st.session_state.history:
                if rec.get("狀態") == "成功":
                    continue
                link = str(rec.get("來源連結", "")).strip()
                if link and link not in seen:
                    seen.add(link)
                    links.append(link)

            if not links:
                raise ValueError("沒有可重試的失敗項目。")

            add_log(f"啟動失敗項目重試，共 {len(links)} 筆")
            operation_name = "重試作業"

        artifacts: list[tuple[str, bytes]] = []
        success_count = 0

        for idx, link in enumerate(links, 1):
            row = {
                "時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "狀態": "失敗",
                "來源連結": link,
                "標題": "",
                "輸出檔名": "",
                "頁數": "",
                "尺寸": "",
                "檔案大小MB": "",
                "訊息": "",
            }
            try:
                status_box.info(f"({idx}/{len(links)}) 系統正在解析連結資訊。")
                add_log(f"[{idx}/{len(links)}] 開始解析 Drive 連結")
                title, total_pages, token, dsmi = parse_drive_view_metadata(link)
                add_log(f"[{idx}/{len(links)}] 解析完成：標題={title}，頁數={total_pages}")

                output_name = build_output_name(naming_mode, title, custom_name)
                temp_root = Path(mkdtemp(prefix="drive_ppt_"))
                image_dir = temp_root / "slides"
                output_file = temp_root / output_name

                status_box.info(f"({idx}/{len(links)}) 系統正在下載投影片頁面。")
                image_paths, size = download_slide_images(
                    token=token,
                    dsmi=dsmi,
                    total_pages=total_pages,
                    image_width=image_width,
                    retries=retries,
                    output_dir=image_dir,
                    log=add_log,
                    update_progress=lambda done, total: download_progress.progress(
                        int(done * 100 / total), text=f"下載進度: {done}/{total}"
                    ),
                )

                status_box.info(f"({idx}/{len(links)}) 系統正在重建 PPTX。")
                final_file = build_pptx_from_images(
                    image_paths=image_paths,
                    size=size,
                    output_path=output_file,
                    log=add_log,
                    update_progress=lambda done, total: build_progress.progress(
                        int(done * 100 / total), text=f"建檔進度: {done}/{total}"
                    ),
                )

                pptx_bytes = final_file.read_bytes()
                artifacts.append((final_file.name, pptx_bytes))
                success_count += 1

                row.update(
                    {
                        "狀態": "成功",
                        "標題": title,
                        "輸出檔名": final_file.name,
                        "頁數": str(total_pages),
                        "尺寸": f"{size[0]} x {size[1]}",
                        "檔案大小MB": f"{len(pptx_bytes) / 1024 / 1024:.2f}",
                        "訊息": "完成",
                    }
                )

                if save_local:
                    local_path = Path.cwd() / final_file.name
                    try:
                        local_path.write_bytes(pptx_bytes)
                        add_log(f"已另存至：{local_path}")
                    except PermissionError:
                        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                        fallback = Path.cwd() / f"{Path(final_file.name).stem}_fixed_{ts}.pptx"
                        fallback.write_bytes(pptx_bytes)
                        add_log(f"原檔被占用，已另存為：{fallback}")

            except Exception as err:  # noqa: BLE001
                row["訊息"] = str(err)
                add_log(f"[{idx}/{len(links)}] 錯誤：{err}")

            st.session_state.history.append(row)

        if success_count == len(links):
            status_box.success(f"{operation_name}完成：{success_count}/{len(links)} 成功。")
        elif success_count > 0:
            status_box.warning(f"{operation_name}完成：{success_count}/{len(links)} 成功，請檢查失敗項目。")
        else:
            status_box.error("作業未完成：全部連結皆失敗。")

        if len(artifacts) == 1:
            fname, data = artifacts[0]
            st.download_button(
                label="下載轉換後 PPTX",
                data=data,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )
        elif len(artifacts) > 1:
            zip_bytes = bundle_zip_bytes(artifacts)
            zip_name = f"drive_ppt_batch_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
            st.download_button(
                label="下載批次成果 ZIP",
                data=zip_bytes,
                file_name=zip_name,
                mime="application/zip",
                use_container_width=True,
            )

    except Exception as err:  # noqa: BLE001
        status_box.error(f"作業未完成：{err}")
        add_log(f"錯誤：{err}")

if st.session_state.history:
    st.subheader("任務歷程")
    st.dataframe(st.session_state.history, use_container_width=True)
    csv_bytes = history_to_csv_bytes(st.session_state.history)
    st.download_button(
        label="下載任務歷程 CSV",
        data=csv_bytes,
        file_name=f"task_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv",
        use_container_width=True,
    )

with st.expander("作業說明", expanded=False):
    st.markdown(
        """
- 僅支援 Google Drive 可公開檢視連結。
- 若原始檔限制下載，本工具會透過檢視頁圖像重建為新 PPTX。
- 若同名檔案被占用，系統會自動附加 `_fixed_時間戳` 後存檔。
"""
    )

# docx → markdown (pandoc)

Chuyển tài liệu Word (`.docx`) sang markdown (GFM) kèm ảnh, không làm mất các hình vẽ bằng
Word Shapes (canvas, group, autoshape, SmartArt, chart, Visio/OLE, EMF/WMF). Nếu chỉ chạy
pandoc, các hình này bị bỏ qua mà không có cảnh báo nào.

Tài liệu này là **hướng dẫn sử dụng**. Chi tiết kỹ thuật của từng bước, các quy tắc làm sạch
và cách mở rộng nằm trong [scripts/README.md](scripts/README.md). Các vấn đề đã biết nằm trong
[scripts/BACKLOG.md](scripts/BACKLOG.md).

## Yêu cầu hệ thống

### Máy chạy pipeline (bắt buộc)

| Thành phần | Phiên bản | Ghi chú |
|---|---|---|
| **Windows** | 10 hoặc 11 | Bước convert điều khiển Word qua COM và dùng clipboard của Windows, nên không chạy được trên macOS/Linux |
| **Microsoft Word desktop** (Office) | 2013 trở lên, 32 hoặc 64 bit | Phải là bản cài trên máy (Microsoft 365 Apps, Office 2016/2019/2021/2024…) **đã kích hoạt bản quyền**. Word for the web và Word trên Mac không dùng được. Word chưa kích hoạt chỉ mở ở chế độ xem nên không lưu được file |
| **Windows PowerShell** | 5.1 | Có sẵn trên Windows 10/11. File `.bat` gọi `powershell.exe` (5.1), không phải PowerShell 7 (`pwsh`) |
| **pandoc** | ≥ 3.0, khuyến nghị ≥ 3.11 | `winget install --id JohnMacFarlane.Pandoc`. Với bản cũ hơn 3.11, tham chiếu chéo ra chữ thường thay vì link |
| **uv** | bản mới | `winget install --id astral-sh.uv -e`. **Không cần cài Python**: lần chạy đầu, uv tự tải Python 3.12 và thư viện đúng phiên bản (cần internet lần đầu; nếu có proxy thì đặt `HTTPS_PROXY`) |

Sau khi cài pandoc và uv, **mở lại cửa sổ terminal** để PATH được cập nhật. Kiểm tra:

```bat
pandoc --version
uv --version
```

Điều kiện khác trên máy chạy:

- PowerShell phải ở chế độ **FullLanguage**. Máy bị khóa bằng AppLocker/WDAC (Constrained
  Language Mode) không chạy được bước convert; script sẽ báo lỗi ngay từ đầu.
- Word không bị chặn bởi hộp thoại lần đầu (đăng nhập, chọn định dạng file, kích hoạt).
  Nên mở Word bằng tay một lần và đóng hết các hộp thoại trước khi chạy.
- **Không dùng clipboard** (copy/paste) trong lúc bước convert đang chạy, và tắt các ứng dụng
  quản lý clipboard. Script lấy hình của shape floating qua clipboard, nên nội dung clipboard
  hiện có sẽ bị xóa.

### Bước nào cần gì

| Bước | Windows + Word | pandoc | uv |
|---|:---:|:---:|:---:|
| [1] preflight, [8] gate (`Test-DocxDrawings.ps1`) | Windows (PowerShell), không cần Word | | |
| [2] cleanup, [4] prep (Python) | | | ✔ (chạy được cả trên macOS/Linux) |
| [3] convert (`Convert-ShapesToPictures.ps1`) | ✔ | | |
| [5] pandoc | | ✔ | |
| [6] media (`Convert-MediaToPng.ps1`) | Windows (GDI+), không cần Word | | |
| [7] publish (`Publish-Output.ps1`) | Windows (PowerShell) | | |

Trên macOS/Linux chỉ phát triển và chạy test được phần Python (`cd scripts/cleanup && uv run pytest`).

## Cách dùng

### 1. Chạy

Kéo thả file `.docx` vào một trong các file `.bat` trong `scripts\`, hoặc chạy từ `cmd`:

```bat
cd scripts

rem Tài liệu có khung đỏ của người review và bảng 1 ô bọc hình: profile review-markup
run-review-markup.bat D:\data\input.docx

rem Tài liệu thường: không làm sạch
run-all.bat D:\data\input.docx
```

Với một bộ dữ liệu mới, nên chạy thử ở chế độ **chỉ báo cáo** trước, kiểm tra
`input.cleanup-manifest.csv`, rồi mới chạy thật:

```bat
run-all.bat D:\data\input.docx -Profile review-markup -CleanupMode Report
```

Một tài liệu dài có thể mất vài phút ở bước convert. Cửa sổ Word chạy ẩn; không cần mở
hay thao tác gì với nó.

### 2. Lấy kết quả

Kết quả nằm cạnh file input:

```
D:\data\
  input.docx          file gốc, không bao giờ bị sửa
  input.out\          KẾT QUẢ BÀN GIAO (tạo lại sau mỗi lần chạy)
    input.md
    images\           chỉ các ảnh PNG mà input.md dùng
  input.work\         file trung gian, log, manifest để kiểm tra
```

Chỉ cần dùng thư mục `input.out\`. Đừng lưu file của mình vào đó: thư mục bị xóa và tạo lại
mỗi lần chạy.

### 3. Đọc kết quả chạy

Dòng cuối trên console cho biết kết quả:

| Mã thoát | Dòng cuối | Ý nghĩa |
|---|---|---|
| `0` | `[PASS] ...` | Mọi hình vẽ đã được chuyển thành ảnh |
| `3` | `[CHECK] Finished with issues ...` | Đã có kết quả, nhưng cần kiểm tra: có hình convert lỗi (`FAILED`), link ảnh hỏng, hoặc còn EMF/WMF |
| `1` | `[ERROR] ...` | Lỗi, không có kết quả mới. Đọc thông báo lỗi |

Khi gặp mã `3`, xem các file trong `input.work\`:

| File | Xem gì |
|---|---|
| `input.pipeline.log` | Toàn bộ log của lần chạy |
| `input.shapes\manifest.csv` | Từng hình vẽ: đã convert hay chưa, lỗi gì (cột `Action`, `Error`) |
| `input.cleanup-manifest.csv` | Các khung đỏ đã xóa, các bảng bọc đã gỡ, và các trường hợp chỉ báo cáo |

### Tham số thường dùng

Thêm sau tên file, ví dụ `run-all.bat input.docx -Dpi 300 -NoPageInfo`:

| Tham số | Tác dụng |
|---|---|
| `-Profile review-markup` | Bật làm sạch khung đỏ và bảng bọc (`run-review-markup.bat` đã có sẵn) |
| `-CleanupMode Report` | Làm sạch chỉ báo cáo, không sửa gì |
| `-Dpi 300` | Ảnh nét hơn (mặc định 200) |
| `-NoPageInfo` | Chạy nhanh hơn với tài liệu dài (bỏ số trang trong manifest) |
| `-IncludeTextBoxes` | Chuyển cả text box đứng riêng thành ảnh (mặc định giữ dạng chữ) |
| `-UnlinkShapeFields` | Dùng khi ảnh hiện `Error! Reference source not found.` |
| `-NoHeadingNumbers` | Không ghi số mục ("7.5") vào tiêu đề |
| `-OutputDir <thư mục>` | Đổi nơi ghi kết quả (mặc định `input.out` cạnh file input) |
| `-OutputFormat markdown` | Xuất Pandoc Markdown thay vì GFM |

Danh sách đầy đủ: chạy `run-all.bat` không có tham số, hoặc xem
[scripts/README.md](scripts/README.md).

Chạy không dừng (ví dụ trong script khác hoặc CI): `set NOPAUSE=1` trước khi gọi file `.bat`.

## Xử lý sự cố nhanh

| Triệu chứng | Cách xử lý |
|---|---|
| `pandoc not found` / `uv not found` | Cài theo bảng yêu cầu ở trên, rồi mở lại cửa sổ terminal |
| `running scripts is disabled` | Chạy qua file `.bat` (đã có `-ExecutionPolicy Bypass`) |
| `needs FullLanguage` | Máy bị khóa AppLocker/WDAC. Nhờ IT cho phép, hoặc chạy trên máy khác |
| `Document is protected` | Gỡ Restrict Editing trong Word rồi chạy lại |
| Treo, hoặc lỗi COM | Đóng hết Word (`taskkill /IM WINWORD.EXE /F`), chạy lại với `-Visible` để xem Word đang hiện hộp thoại gì |
| Nhiều dòng `FAILED` liên quan clipboard | Không copy/paste trong lúc chạy; tắt ứng dụng quản lý clipboard |
| File tải từ mạng không mở được | Chuột phải → Properties → Unblock, hoặc `Unblock-File .\input.docx` |

Các trường hợp khác: xem mục "Xử lý sự cố" trong [scripts/README.md](scripts/README.md).

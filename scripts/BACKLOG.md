# Backlog

Các vấn đề đã biết và điểm cải tiến **chưa xử lý**, sẽ thống nhất và làm sau.
Mức ưu tiên: **P1** là ảnh hưởng trực tiếp tới nội dung đầu ra; **P2** là chất lượng hoặc
khả năng sử dụng; **P3** là kỹ thuật hoặc vận hành.

## Vấn đề đã biết

| ID | Ưu tiên | Vấn đề | Dữ kiện đã có | Bước tiếp theo |
|---|---|---|---|---|
| – | – | (không còn vấn đề đã biết nào đang mở) | | |

## Quyết định còn mở

| ID | Ưu tiên | Câu hỏi | Lựa chọn |
|---|---|---|---|
| D-02 | P2 | **Chữ trong hình** cho các bước sau (LLM…) | Hiện tại nằm trong thuộc tính `alt`. Các lựa chọn khác: một khối text ngay dưới ảnh; dùng LLM đọc ảnh để sinh mô tả; chuyển sơ đồ trạng thái (như Fig 7‑30) sang Mermaid |

## Cải tiến theo bước

### Cleanup

| ID | Ưu tiên | Nội dung |
|---|---|---|
| C-01 | P1 | **Hiệu chỉnh ReviewerBoxes trên dữ liệu NS thật** (màu, độ trong suốt, `includeInGroups`) dựa trên các dòng `low`/`medium` trong `cleanup-manifest.csv` |
| C-02 | P2 | UnwrapLayoutTables: bảng bọc **nhiều cột** (hai hình đặt cạnh nhau) hiện không xử lý. Có thể tách thành các hình liên tiếp |
| C-03 | P3 | UnwrapLayoutTables: bảng 1 ô **chỉ chứa chữ** (khung Note) hiện chỉ báo cáo. Có thể chuyển thành blockquote |
| C-04 | P3 | ReviewerBoxes: chế độ `Mark`, tức đánh dấu phần nội dung được khoanh (`<mark>`) thay vì chỉ xóa khung. Chỉ làm nếu phần khoanh đỏ có ý nghĩa |
| C-05 | P3 | Khung đỏ đã được **vẽ sẵn vào ảnh chụp màn hình** (raster): ngoài phạm vi. Có thể chỉ đánh dấu các ảnh có nhiều pixel đỏ để kiểm tra thủ công |

### Convert (Word)

| ID | Ưu tiên | Nội dung |
|---|---|---|
| V-01 | P2 | Chưa xử lý shape trong header, footer, footnote, comment |
| V-02 | P2 | Ảnh EMF/WMF dạng **floating** chưa được chuyển sang PNG |
| V-03 | P2 | Gom shape rời thành một hình dựa vào caption: hình không có caption và neo ở nhiều paragraph có thể bị tách. Caption nằm trong text box floating có thể bị render vào ảnh |
| V-04 | P3 | Đang dùng lại dấu phân cách ` \| ` trong alt text (`Build-AltText`), và `figures.lua` dựa vào dấu này để chuyển thành `; `. Nếu chạy không có filter thì `\|` vẫn xuất hiện |
| V-05 | P3 | Hiệu năng: mở và đóng Word một lần cho cả lô tài liệu, thay vì mỗi file một lần |

### Pandoc / đầu ra

| ID | Ưu tiên | Nội dung |
|---|---|---|
| P-01 | P3 | Pandoc đổi tên ảnh thành `imageN.png`, nên chỉ còn thuộc tính `data-shape` nối ảnh với manifest. Có thể dùng filter để đặt tên file theo Id (`S001.png`) |
| P-02 | P2 | **Giảm token**: GFM giữ lại các thuộc tính chỉ để trình bày (`style="width:…"`, `<colgroup>`, `style` của bảng). Có thể dùng filter để bỏ bớt |
| P-03 | P2 | **Yêu cầu với RAG**: giữ thẻ HTML, không cắt chunk giữa `<table>` hoặc `<figure>`. Cần kiểm tra khi tích hợp |
| P-04 | P3 | Bảng không có hàng tiêu đề trong Word bị xuất thành bảng HTML, dù rất đơn giản. Có thể coi hàng đầu tiên là tiêu đề khi hàng đó in đậm hoặc có tô nền |
| P-06 | P3 | Với pandoc cũ hơn 3.11, tham chiếu `REF` ra chữ thường thay vì link. Nên thống nhất dùng pandoc ≥ 3.11 trên máy chạy |
| P-05 | P3 | GitHub thêm tiền tố `user-content-` vào `id`, nên link tham chiếu chéo (`#_Ref…`) có thể không nhảy đúng khi xem trên GitHub. VS Code hiển thị bình thường |

### Kiểm tra & vận hành

| ID | Ưu tiên | Nội dung |
|---|---|---|
| Q-01 | P2 | Phép đếm caption và ảnh ở bước gate chỉ là heuristic. Có thể mở rộng sang bảng và đếm trên AST của pandoc thay vì regex |
| Q-02 | P3 | Chuyển `Test-DocxDrawings.ps1` sang package Python để dùng chung bộ phân loại XML với bước cleanup |
| Q-03 | P3 | CI: `uv run pytest`, PSScriptAnalyzer cho các script `.ps1`, và một bài test tích hợp trên Windows với file mẫu đã ẩn thông tin |
| Q-04 | P3 | Chạy theo lô nhiều file docx, kèm báo cáo tổng hợp |

## Đã xử lý

| Vấn đề | Cách xử lý |
|---|---|
| Pandoc bỏ qua shape vẽ bằng Word mà không báo | `Convert-ShapesToPictures.ps1` |
| Khung đỏ của người review làm hình bị tách, sinh ảnh khung đỏ | Quy tắc cleanup `ReviewerBoxes` |
| Hình và bảng bị đóng trong bảng 1 ô (grid table) | Quy tắc cleanup `UnwrapLayoutTables` |
| Alt text và title làm vỡ cú pháp ảnh (`\|`, ngắt dòng, `shape2png`) | `pandoc\figures.lua` + `--wrap=none` |
| Ảnh render ra `Error! Reference source not found.` thay cho tham chiếu chéo trong hình (ví dụ "refer to 7.5"). Nguyên nhân: Word cập nhật field khi chuyển/render shape, lúc đó không thấy bookmark | Bước convert khóa field trong shape (`wdTextFrameStory`) trước khi chuyển đổi, và tắt `UpdateFieldsAtPrint`. Dự phòng: `-UnlinkShapeFields` |
| Mất số caption và nhãn "Table", mất tham chiếu chéo (trước là B-01). Nguyên nhân: 149 field `w:fldSimple` có sẵn trong tài liệu gốc (75 STYLEREF, 74 SEQ); pandoc bỏ chữ của `fldSimple` | Bước prep, quy tắc `ExpandSimpleFields` (`profiles\pandoc.json`): chuyển về dạng field đầy đủ |
| Caption bảng bị đưa xuống dưới bảng (bảng pipe của GFM không có caption) | `figures.lua`: caption thành đoạn văn phía trên bảng |
| Markdown đúng nhưng không hiển thị được (grid table, `{…}`, `: caption`). Trước là D-01 | Xuất `gfm`: bảng đơn giản thành bảng pipe, bảng phức tạp và hình thành HTML. Bookmark của caption bảng (trước là B-02) thành `id` HTML hợp lệ |
| Đường dẫn sai trong `.bat` sau lệnh `shift` | Lưu thư mục script trước khi `shift` |

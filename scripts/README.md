# Chuyển shape/drawing trong Word sang PNG trước khi chạy pandoc (Option A)

Pandoc bỏ qua các hình vẽ bằng Word Shapes mà không cảnh báo (canvas, group, autoshape,
SmartArt, chart). Script `Convert-ShapesToPictures.ps1` dùng chính Word (qua COM) để
render các hình đó ra PNG, rồi thay vào docx. Sau đó pandoc trích xuất ảnh bình thường.

Yêu cầu: Windows, Word desktop 2013 trở lên, Windows PowerShell 5.1 (có sẵn), pandoc.

## Cách nhanh: dùng file .bat

Double-click, hoặc kéo thả file docx vào file `.bat`:

```bat
rem Chạy đủ pipeline: preflight -> convert -> pandoc -> gate
run-all.bat input.docx
rem kết quả: input.shapes.docx, input.shapes\ (PNG + manifest.csv), input.md, images\media\

rem Chỉ convert (dùng được -DryRun và mọi tham số khác)
convert.bat input.docx -DryRun
convert.bat input.docx -Dpi 300 -IncludeTextBoxes
```

Mã thoát của `run-all.bat`: `0` là PASS, `3` là chạy xong nhưng có vấn đề (convert lỗi một
phần hoặc gate không đạt), `1` là lỗi (thiếu file, thiếu pandoc, convert hoặc pandoc thất
bại). Đặt `set NOPAUSE=1` để chạy không dừng, ví dụ khi dùng trong CI.
Logic nằm hoàn toàn trong các file `.ps1`; file `.bat` chỉ là lớp bao để gọi cho tiện.

## Các bước chạy thủ công

```powershell
cd <thư mục chứa script>

# 0. Preflight: đếm số đối tượng pandoc sẽ bỏ qua (không cần Word)
powershell -ExecutionPolicy Bypass -File .\Test-DocxDrawings.ps1 -Docx .\input.docx

# 1. Dry run: chỉ liệt kê, không thay đổi gì -> input.shapes\manifest.csv
powershell -ExecutionPolicy Bypass -File .\Convert-ShapesToPictures.ps1 -InputPath .\input.docx -DryRun

# 2. Convert -> input.shapes.docx + input.shapes\S001.png, S001.emf, ..., manifest.csv
powershell -ExecutionPolicy Bypass -File .\Convert-ShapesToPictures.ps1 -InputPath .\input.docx

# 3. Chạy pandoc trên file đã convert
pandoc -f docx -t markdown --extract-media=./images .\input.shapes.docx -o output.md

# 4. Gate: fail nếu còn đối tượng pandoc sẽ bỏ qua; cảnh báo (không fail) nếu số caption > số ảnh
powershell -ExecutionPolicy Bypass -File .\Test-DocxDrawings.ps1 -Docx .\input.shapes.docx -Markdown .\output.md
```

File gốc không bị sửa. Trong lúc script chạy, **không dùng clipboard**, vì script cần
copy/paste để lấy hình của các shape floating.

## Tham số

| Tham số | Mặc định | Ý nghĩa |
|---|---|---|
| `-OutputPath` | `<input>.shapes.docx` | File docx kết quả |
| `-ImageDir` | `<input>.shapes\` | Thư mục chứa PNG, EMF (để debug) và `manifest.csv` |
| `-Dpi` | 200 | Độ phân giải PNG |
| `-IncludeTextBoxes` | tắt | Chuyển cả text box đứng riêng. Mặc định giữ lại để không mất văn bản |
| `-KeepMetafiles` | tắt | Giữ nguyên OLE (Visio...) và ảnh EMF/WMF. Mặc định sẽ chuyển sang PNG vì trình xem markdown không hiển thị được EMF |
| `-NoCluster` | tắt | Không gộp các shape rời của cùng một hình |
| `-NoTrim` | tắt | Không cắt viền trắng |
| `-DryRun` | tắt | Chỉ liệt kê |
| `-Visible` | tắt | Hiện cửa sổ Word để debug |

Exit code: `0` là OK, `2` là có đối tượng convert thất bại (xem cột `Error` trong manifest).

## Cách hoạt động

1. **Shape floating** (`Document.Shapes`):
   - Các shape rời thuộc cùng một hình được **gộp thành một group** trước khi render. Hai
     shape được coi là cùng hình nếu cùng caption `Fig./Figure/Hình` và cùng trang, hoặc
     cùng paragraph anchor. Mục đích là không để một hình bị cắt thành nhiều ảnh nhỏ.
   - Cách lấy hình được thử theo thứ tự:
     1. `ConvertToInlineShape()`, chỉ áp dụng được cho picture/OLE.
     2. Copy rồi đọc EMF từ clipboard qua Win32.
     3. Paste Special dạng EMF vào một document tạm.
   - PNG được chèn thành một paragraph riêng (style Normal) ngay trước paragraph anchor,
     sau đó xóa shape gốc.
2. **Shape inline** (`Document.InlineShapes`): script phân loại theo XML
   (`Range.WordOpenXML`) chứ không theo kiểu COM, vì Word không phân loại đúng các
   shape/group/canvas DrawingML nằm inline. Hình được lấy qua `Range.EnhMetaFileBits`
   và thay tại chỗ.
3. EMF được render sang PNG bằng GDI+. PNG được gán DPI để Word giữ đúng kích thước in.
4. **Text trong hình** (các ô trạng thái, nhãn mũi tên...) được ghi vào alt text của ảnh.
   Pandoc xuất alt text thành `![Drawing converted to image. Text: Blank | Redisplaying ...](images/media/imageN.png)`,
   nên thông tin vẫn tìm kiếm được và dùng được cho RAG/LLM.

## Giới hạn đã biết

- Nếu caption nằm trong một text box floating cạnh hình, text box đó có thể bị gộp vào
  cluster và render thành ảnh. Caption khi đó chỉ còn trong alt text. Nên kiểm tra các
  cluster trong manifest.
- Macro trong file input không bao giờ chạy (`AutomationSecurity = ForceDisable`).
- Máy bị khóa bằng AppLocker/WDAC (PowerShell ở Constrained Language Mode) không chạy
  được `Add-Type`/COM. Script sẽ báo lỗi rõ ràng ngay từ đầu.
- Tài liệu bị Restrict Editing có mật khẩu: phải gỡ bảo vệ trong Word trước khi chạy.
- Chỉ xử lý nội dung chính của document. Shape nằm trong header, footer, footnote và
  comment chưa được xử lý.
- Việc gộp shape rời dựa vào caption. Nếu hình không có caption và các shape neo ở nhiều
  paragraph khác nhau, chúng có thể ra thành nhiều ảnh. Hãy xem manifest; nếu cần, gộp
  thủ công trong Word (Select, rồi Group).
- Word không cho group một drawing canvas với shape khác. Khi đó cluster được tách ra và
  convert từng shape riêng (manifest ghi `Action=split`).
- Ảnh EMF/WMF dạng **floating** chưa được chuyển (chỉ ảnh inline được chuyển).
- Font render bằng GDI+ có thể hơi khác so với Word. Nếu cần giống tuyệt đối, có thể xuất
  PDF từ Word rồi crop.

## Khi test xong, gửi lại

- `manifest.csv`
- Log trên console, nhất là các dòng `FAILED`
- Output của `Test-DocxDrawings.ps1` trước và sau khi convert
- 1–2 ảnh PNG, ví dụ ảnh của Fig 7‑30, để đánh giá chất lượng render

## Xử lý sự cố

| Triệu chứng | Cách xử lý |
|---|---|
| `running scripts is disabled` | Chạy qua `powershell -ExecutionPolicy Bypass -File ...` |
| File tải từ mạng không mở được | `Unblock-File .\input.docx` |
| Treo hoặc lỗi COM | Đóng hết Word (`Get-Process WINWORD \| Stop-Process`), rồi chạy lại với `-Visible` để xem Word đang báo gì |
| Nhiều dòng `FAILED` với `clipboard`/`Paste Special` | Không dùng clipboard trong lúc chạy; tắt các app quản lý clipboard |
| PNG bị cắt hoặc thừa viền | Chạy với `-NoTrim`, rồi so với file `.emf` tương ứng |

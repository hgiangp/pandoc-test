# Pipeline docx → markdown (pandoc) cho dữ liệu NS

Pandoc bỏ qua các hình vẽ bằng Word Shapes mà không cảnh báo (canvas, group, autoshape,
SmartArt, chart). Ngoài ra, dữ liệu của khách hàng NS có các khung đỏ do người review tự vẽ
để đánh dấu. Các khung này không mang thông tin, nhưng làm hình bị tách nhỏ hoặc sinh ra
ảnh khung đỏ thừa. Pipeline dưới đây xử lý cả hai vấn đề **trước** khi chạy pandoc.

```
input.docx
 ├─[1] preflight   Test-DocxDrawings.ps1        đếm các hình pandoc sẽ bỏ qua
 ├─[2] cleanup     cleanup\ (Python)            làm sạch dữ liệu theo profile, bật/tắt được → input.clean.docx
 ├─[3] convert     Convert-ShapesToPictures.ps1 chuyển hình vẽ sang PNG bằng Word  → input.shapes.docx
 ├─[4] pandoc                                                                    → input.md + images\
 └─[5] gate        Test-DocxDrawings.ps1        kiểm tra không còn hình nào bị mất
```

- Bước **cleanup** xử lý vấn đề riêng của **dữ liệu**. Bước **convert** vá giới hạn của
  **pandoc**. Hai bước được tách riêng: tắt cleanup thì pipeline chạy y như trước.
- Cleanup chạy trước convert, để khung đỏ không bị gom chung với hình thật.
- Mỗi bước ghi ra file docx mới. Mọi file trung gian đều được giữ lại để mở bằng Word kiểm
  tra. File gốc không bao giờ bị sửa.

## Cài đặt (một lần trên máy chạy)

| Thành phần                          | Dùng cho                                          | Cài đặt                                                                           |
| ------------------------------------- | -------------------------------------------------- | ------------------------------------------------------------------------------------ |
| Windows + Word desktop 2013 trở lên | convert                                            | –                                                                                   |
| Windows PowerShell 5.1                | toàn bộ                                          | Có sẵn trên Windows                                                               |
| pandoc                                | pandoc                                             | `winget install --id JohnMacFarlane.Pandoc`                                        |
| [uv](https://docs.astral.sh/uv/) | cleanup (chỉ khi dùng profile có bật quy tắc) | `winget install --id astral-sh.uv -e`. **Không cần cài Python riêng**: lần chạy đầu, uv tự tải Python 3.12 (ghim trong `cleanup\.python-version`) và các thư viện đúng phiên bản trong `cleanup\uv.lock` (cần internet lần đầu) |

## Cách chạy nhanh: dùng file .bat

Double-click, hoặc kéo thả file docx vào file `.bat`:

| File                                                    | Việc làm                                                                |
| ------------------------------------------------------- | ------------------------------------------------------------------------- |
| `run-ns.bat input.docx`                               | **Pipeline đầy đủ với profile NS** (có làm sạch khung đỏ) |
| `run-all.bat input.docx [tham số]`                   | Pipeline đầy đủ. Mặc định**không** làm sạch               |
| `cleanup.bat input.docx --profile ns [--mode Report]` | Chỉ chạy bước làm sạch (`uv run ... docx-cleanup`) |
| `convert.bat input.docx [-DryRun]`                    | Chỉ chạy bước convert                                                 |

```bat
rem Lần đầu với dữ liệu mới: chỉ báo cáo, chưa xóa gì, để kiểm tra kết quả nhận diện
run-all.bat input.docx -Profile ns -CleanupMode Report

rem Chạy thật với profile NS
run-ns.bat input.docx

rem So sánh với khi không làm sạch
run-all.bat input.docx -NoCleanup
```

Kết quả nằm cạnh `input.docx`:

| File                                     | Nội dung                                                  |
| ---------------------------------------- | ---------------------------------------------------------- |
| `input.clean.docx`                     | Sau bước làm sạch (chỉ có khi profile bật quy tắc) |
| `input.cleanup-manifest.csv`           | Các quy tắc làm sạch đã tìm thấy và xử lý gì   |
| `input.shapes.docx`, `input.shapes\` | Sau bước convert: PNG, EMF và`manifest.csv`           |
| `input.md`, `images\media\`          | Kết quả của pandoc                                      |
| `input.pipeline.log`                   | Log toàn bộ lần chạy                                   |

Mã thoát: `0` là PASS, `3` là chạy xong nhưng có vấn đề, `1` là lỗi. Đặt `set NOPAUSE=1`
để chạy không dừng, ví dụ khi dùng trong CI. Logic nằm trong `Invoke-Pipeline.ps1`;
file `.bat` chỉ là lớp bao để gọi cho tiện.

## Làm sạch dữ liệu (cleanup)

### Profile

Mỗi profile là một file JSON trong `profiles\`. Trong profile, mỗi quy tắc có một `mode`:

| mode       | Hành vi                                                 |
| ---------- | -------------------------------------------------------- |
| `Off`    | Không chạy                                             |
| `Report` | Chỉ nhận diện và ghi vào manifest, không sửa docx |
| `Apply`  | Nhận diện và xử lý (xóa)                           |

- `default.json`: mọi quy tắc `Off`. Bước cleanup được bỏ qua và không cần uv/Python.
- `ns.json`: `ReviewerBoxes` = `Apply`.
- Tham số `-CleanupMode Report` (trong `run-all.bat`) hoặc `--mode Report` (trong
  `cleanup.bat`) ép mọi quy tắc **đang bật** chạy ở chế độ Report. Tham số này không bao giờ
  bật một quy tắc đang `Off`.
- Nếu profile có tên tham số sai (gõ nhầm), chương trình dừng với lỗi và báo rõ, không âm
  thầm dùng giá trị mặc định.

### Quy tắc `ReviewerBoxes`: khung đỏ của người review

Một shape được coi là khung đánh dấu khi thỏa **tất cả** điều kiện sau:

| Điều kiện                                                                                                             | Tham số                                        | Mặc định                                  |
| ------------------------------------------------------------------------------------------------------------------------ | ----------------------------------------------- | -------------------------------------------- |
| Hình chữ nhật hoặc chữ nhật bo góc                                                                                | `geometries`                                  | `["rect", "roundRect"]`                    |
| Không tô nền, hoặc nền gần trong suốt                                                                             | `maxFillOpacity`                              | `0.1`                                      |
| Viền đỏ: gần một màu trong danh sách, hoặc có sắc đỏ rõ                                                     | `colors`, `colorTolerance`, `matchRedHue` | `FF0000, C00000, E60012`; `60`; `true` |
| Không chứa chữ                                                                                                        | –                                              | –                                           |
| Không có connector (mũi tên) nối vào. Ô trong sơ đồ thường có mũi tên nối, khung đánh dấu thì không | –                                              | –                                           |

| Mức tin cậy | Khi nào                                                              | Xử lý ở chế độ Apply             |
| ------------- | --------------------------------------------------------------------- | -------------------------------------- |
| `high`      | Thỏa hết điều kiện, là shape đứng riêng                      | Xóa                                   |
| `medium`    | Thỏa hết điều kiện, nhưng nằm trong group hoặc canvas         | Chỉ xóa khi`includeInGroups: true` |
| `low`       | Viền màu đỏ hoặc gần đỏ, và trượt đúng một điều kiện | Chỉ báo cáo, dùng để tinh chỉnh |

Màu được đọc cả khi đặt trực tiếp (`srgbClr`), khi lấy từ theme (`schemeClr`, có tính các
biến thể sáng/tối), khi lấy từ style của shape (`lnRef`) và khi là shape VML. Phần
`mc:Fallback` (bản sao VML của cùng một shape) không bị đếm lặp.

Khi xóa:

- **Shape đứng riêng**: xóa cả khối `mc:AlternateContent`, gồm cả phần Fallback. Chữ nằm
  bên dưới khung vẫn giữ nguyên.
- **Shape trong group/canvas**: chỉ xóa shape con. Phần Fallback của group bị bỏ, vì không
  còn khớp với nội dung mới; Word sẽ tạo lại phần này ở bước convert. Nếu group không còn
  shape nào thì xóa luôn group.

### Đọc `input.cleanup-manifest.csv`

| Cột                                      | Ý nghĩa                                                                                                                      |
| ----------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| `Action`                                | `removed`, `reported` hoặc `reported (set includeInGroups to remove)`                                                   |
| `Confidence` / `Reason`               | Mức tin cậy; với mức`low`, cột này ghi điều kiện bị trượt (`fails: fill`, `fails: text`, …)                 |
| `Part`, `Paragraph`                   | Phần tài liệu (document, header, …) và số thứ tự paragraph chứa shape                                                 |
| `Location`                              | Chữ của paragraph đó, hoặc của paragraph không rỗng kế tiếp (thường là caption), giúp tìm lại shape trong Word |
| `ShapeId`, `ShapeName`, `Container` | Id và tên trong Word;`top`, `group` hoặc `canvas`                                                                     |
| `Features`                              | Các đặc trưng đã đo (JSON): hình dạng, màu viền, độ dày viền, kiểu nét, độ trong suốt của nền, …        |

**Cách tinh chỉnh**:

- Nếu thấy dòng `low` với `fails: color` mà đúng là khung đánh dấu, thêm màu ở `Features.lineColor` vào `colors`.
- Nếu thấy dòng `medium` đúng là khung đánh dấu, đặt `includeInGroups: true`.
- Nếu shape bị xóa nhầm, chuyển quy tắc về `Report` và gửi dòng manifest tương ứng cho mình.

### Thêm quy tắc mới

1. Tạo `cleanup\docx_cleanup\rules\<ten_quy_tac>.py`, kế thừa `Rule` (`model.py`) và cài đặt
   `detect()` (không được sửa docx) và `apply()`.
2. Đăng ký class trong `registry.py`.
3. Thêm cấu hình vào profile (`profiles\ns.json`).
4. Viết test trong `cleanup\tests\`. `conftest.py` có sẵn các hàm tạo docx mẫu (shape,
   group, canvas, VML).

Các thao tác xóa an toàn (xử lý AlternateContent/Fallback, dọn run rỗng, kiểm tra file sau
khi ghi) nằm tập trung trong `DocxPackage` (`package.py`). Quy tắc mới nên gọi các hàm này,
không tự thao tác trực tiếp trên XML.

### Quản lý package bằng uv

`cleanup\` là một uv project:

| File | Vai trò |
|---|---|
| `pyproject.toml` | Khai báo dependency (`lxml`), nhóm dev (`pytest`) và lệnh `docx-cleanup` |
| `uv.lock` | Ghim chính xác phiên bản mọi thư viện, **phải commit** |
| `.python-version` | Phiên bản Python uv dùng (3.12). Code tương thích Python ≥ 3.9 |

Pipeline gọi `uv run --project cleanup --locked --no-dev docx-cleanup ...`:
- `--locked`: báo lỗi nếu `uv.lock` không khớp `pyproject.toml`, thay vì âm thầm cài phiên bản khác.
- `--no-dev`: không cài pytest trên máy xử lý.

```bash
cd scripts/cleanup
uv sync                                  # tạo .venv (có cả nhóm dev)
uv run pytest                            # chạy test, không cần Windows
uv run --isolated --python 3.9 pytest    # kiểm tra tương thích Python 3.9
uv add <package>                         # thêm dependency (tự cập nhật uv.lock)
uv lock --upgrade                        # nâng phiên bản các thư viện
```

## Các bước chạy thủ công (bước convert)

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

Trong lúc bước convert chạy, **không dùng clipboard**, vì script cần copy/paste để lấy hình
của các shape floating.

## Tham số của bước convert

| Tham số              | Mặc định             | Ý nghĩa                                                                                                                            |
| --------------------- | ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------ |
| `-OutputPath`       | `<input>.shapes.docx` | File docx kết quả                                                                                                                  |
| `-ImageDir`         | `<input>.shapes\`     | Thư mục chứa PNG, EMF (để debug) và`manifest.csv`                                                                            |
| `-Dpi`              | 200                     | Độ phân giải PNG                                                                                                                 |
| `-IncludeTextBoxes` | tắt                    | Chuyển cả text box đứng riêng. Mặc định giữ lại để không mất văn bản                                                 |
| `-KeepMetafiles`    | tắt                    | Giữ nguyên OLE (Visio...) và ảnh EMF/WMF. Mặc định sẽ chuyển sang PNG vì trình xem markdown không hiển thị được EMF |
| `-NoCluster`        | tắt                    | Không gộp các shape rời của cùng một hình                                                                                    |
| `-NoTrim`           | tắt                    | Không cắt viền trắng                                                                                                             |
| `-DryRun`           | tắt                    | Chỉ liệt kê                                                                                                                       |
| `-Visible`          | tắt                    | Hiện cửa sổ Word để debug                                                                                                       |

Exit code: `0` là OK, `2` là có đối tượng convert thất bại (xem cột `Error` trong manifest).

## Cách hoạt động của bước convert

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

- `input.pipeline.log`
- `input.cleanup-manifest.csv`, nhất là các dòng khung đỏ bị xóa nhầm hoặc bị bỏ sót
- `input.shapes\manifest.csv`
- Log trên console, nhất là các dòng `FAILED`
- Output của `Test-DocxDrawings.ps1` trước và sau khi convert
- 1–2 ảnh PNG, ví dụ ảnh của Fig 7‑30, để đánh giá chất lượng render

## Xử lý sự cố

| Triệu chứng                                               | Cách xử lý                                                                                                             |
| ----------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------- |
| `uv not found in PATH` | `winget install --id astral-sh.uv -e`, mở lại cửa sổ terminal; hoặc chạy với `-NoCleanup` |
| `The lockfile ... needs to be updated` | `uv.lock` không khớp `pyproject.toml`: chạy `uv lock` trong `cleanup\` rồi commit |
| Lần đầu chạy cleanup bị lỗi tải | uv cần internet để tải Python và thư viện lần đầu. Nếu có proxy, đặt `HTTPS_PROXY` |
| `running scripts is disabled`                             | Chạy qua`powershell -ExecutionPolicy Bypass -File ...`                                                                 |
| File tải từ mạng không mở được                      | `Unblock-File .\input.docx`                                                                                             |
| Treo hoặc lỗi COM                                         | Đóng hết Word (`Get-Process WINWORD \| Stop-Process`), rồi chạy lại với `-Visible` để xem Word đang báo gì |
| Nhiều dòng`FAILED` với `clipboard`/`Paste Special` | Không dùng clipboard trong lúc chạy; tắt các app quản lý clipboard                                                |
| PNG bị cắt hoặc thừa viền                              | Chạy với`-NoTrim`, rồi so với file `.emf` tương ứng                                                            |

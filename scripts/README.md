# Pipeline docx → markdown (pandoc) cho dữ liệu NS

Pandoc bỏ qua các hình vẽ bằng Word Shapes mà không cảnh báo (canvas, group, autoshape,
SmartArt, chart). Ngoài ra, dữ liệu của khách hàng NS có hai đặc trưng riêng:
- Các **khung đỏ** do người review tự vẽ để đánh dấu. Khung không mang thông tin, nhưng làm
  hình bị tách nhỏ hoặc sinh ra ảnh khung đỏ thừa.
- Hình và bảng được đặt cùng caption trong một **bảng 1 ô dùng để dàn trang**. Pandoc vì vậy
  xuất cả khối thành grid table, và ảnh cùng bảng hiển thị như bị đóng trong một ô.

Pipeline dưới đây xử lý các vấn đề này trước và trong bước pandoc.

```
input.docx
 ├─[1] preflight   Test-DocxDrawings.ps1        đếm các hình pandoc sẽ bỏ qua
 ├─[2] cleanup     cleanup\ (Python)            làm sạch dữ liệu theo profile, bật/tắt được → input.clean.docx
 │                   ReviewerBoxes → UnwrapLayoutTables   (profile ns)
 ├─[3] convert     Convert-ShapesToPictures.ps1 chuyển hình vẽ sang PNG bằng Word  → input.shapes.docx
 ├─[4] pandoc      -t gfm --wrap=none --lua-filter=pandoc\figures.lua             → input.md + images\
 └─[5] gate        Test-DocxDrawings.ps1        kiểm tra không còn hình nào bị mất
```

- Bước **cleanup** xử lý vấn đề riêng của **dữ liệu**. Bước **convert** vá giới hạn của
  **pandoc**. Hai bước được tách riêng: tắt cleanup thì pipeline chạy y như trước.
- Cleanup chạy trước convert, để khung đỏ không bị gom chung với hình thật, và để ảnh PNG
  được chèn ra ngoài bảng, ngay trước caption.
- Mỗi bước ghi ra file docx mới. Mọi file trung gian đều được giữ lại để mở bằng Word kiểm
  tra. File gốc không bao giờ bị sửa.

### Nguyên tắc: xử lý ở bước nào?

Khi gặp một trường hợp mới, chọn bước theo **bản chất** của vấn đề:

| Bản chất của vấn đề | Xử lý ở | Ví dụ |
|---|---|---|
| **Nhiễu cấu trúc của dữ liệu nguồn**: những thứ không phải nội dung, do cách soạn thảo hoặc review | **Cleanup** (quy tắc Python sửa XML của docx, bật/tắt theo profile) | Khung đỏ của người review, bảng 1 ô dùng để dàn trang |
| **Giới hạn của pandoc** với một loại đối tượng Word | **Convert** (dùng Word) | Shape, canvas, SmartArt, chart, OLE, EMF |
| **Cách trình bày đầu ra**, phụ thuộc vào ai đọc markdown | **Lua filter trong pandoc** (sửa trên AST) | Alt text, title và id của ảnh, định dạng bảng, caption của bảng |
| Chỉnh sửa thuần văn bản trên file `.md` | Bước hậu xử lý bằng chữ, **chỉ khi thật cần** | Chuẩn hóa khoảng trắng |

Lý do:
- **Sửa cấu trúc ở nguồn thì pandoc hiểu đúng cấu trúc.** Ví dụ: sau khi gỡ bảng bọc ở
  docx, pandoc tự ghép ảnh với caption thành figure, tự gắn caption cho bảng và giữ bookmark.
  Gỡ ở sau thì phải tự dựng lại những thứ đó.
- Ở mức docx còn đọc được thông tin Word (style Caption, field SEQ, màu và hình dạng shape);
  sang markdown thì các thông tin này đã mất.
- **Không sửa cấu trúc bằng regex trên file `.md`.** Grid table, ký tự escape và ngắt dòng làm
  cách này rất dễ vỡ. Việc cần làm sau pandoc thì làm trên AST bằng Lua filter.
- Logic riêng của một bộ dữ liệu phải nằm trong profile của cleanup, không trộn vào các bước
  chung (convert, pandoc).

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

rem So sánh với khi không làm sạch / không dùng filter của pandoc
run-all.bat input.docx -NoCleanup
run-all.bat input.docx -Profile ns -NoFigureFilter
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
- `ns.json`: `ReviewerBoxes` rồi `UnwrapLayoutTables`, cả hai `Apply`. **Các quy tắc chạy
  theo đúng thứ tự trong profile.** ReviewerBoxes phải chạy trước, để khung đỏ nằm trong bảng
  bọc bị xóa trước khi nội dung của bảng được đưa ra ngoài.
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

### Quy tắc `UnwrapLayoutTables`: bảng 1 ô bọc hình/bảng và caption

Một bảng được coi là **bảng bọc để dàn trang** khi thỏa tất cả điều kiện sau:

| Điều kiện | Tham số | Mặc định |
|---|---|---|
| Chỉ có 1 cột (mỗi hàng đúng 1 ô) | – | – |
| Không quá `maxRows` hàng, ví dụ hình ở hàng 1 và caption ở hàng 2 | `maxRows` | `4` |
| Chứa hình (drawing, picture, OLE) hoặc một bảng con | – | – |
| Có caption: style Caption (hoặc style kế thừa từ nó), hoặc đoạn văn chứa field `SEQ` | – | – |

Style Caption được tìm theo **tên** (`caption`), không theo id. Word bản địa hóa (tiếng Nhật,
tiếng Việt…) lưu id dạng `a3`, nhưng tên style dựng sẵn thì luôn là tiếng Anh.

| Mức tin cậy | Khi nào | Xử lý ở chế độ Apply |
|---|---|---|
| `high` | Thỏa hết điều kiện | Gỡ bảng bọc: nội dung trong ô (đoạn văn, ảnh, bảng con, bookmark) được đưa ra ngoài theo đúng thứ tự |
| `medium` | Có hình/bảng con nhưng không có caption | Chỉ gỡ khi `unwrapWithoutCaption: true` |
| `low` | Vượt `maxRows`, hoặc chỉ chứa chữ (ví dụ khung Note) | Chỉ báo cáo |

Bảng nhiều cột, ví dụ hai hình đặt cạnh nhau, **không bao giờ** bị gỡ.

Bookmark `_Ref…` (đích của các link "Fig. 7‑30", "Table 1‑2") đi theo caption, nên tham
chiếu chéo vẫn hoạt động sau khi gỡ.

### Đọc `input.cleanup-manifest.csv`

| Cột                                      | Ý nghĩa                                                                                                                      |
| ----------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| `Action` | `removed`, `unwrapped`, `reported`, hoặc `reported (set … to …)` khi quy tắc cần bật thêm tham số mới xử lý |
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

## Bước pandoc: định dạng đầu ra và `pandoc\figures.lua`

### Định dạng đầu ra: GFM

Pipeline xuất **GFM** (GitHub Flavored Markdown, `-t gfm`), không xuất Pandoc Markdown
(`-t markdown`). Lý do là cần cân bằng giữa người đọc và AI:

| Tiêu chí | Pandoc Markdown | **GFM (mặc định)** |
|---|---|---|
| Người đọc (GitHub, VS Code, GitLab) | ❌ Grid table `+---+`, `{…}` và `: caption` hiện thành chữ thô | ✅ Hiển thị đúng |
| AI, bảng đơn giản | Grid table: nhiều dấu cách căn cột, tốn token | Bảng pipe: gọn nhất |
| AI, bảng phức tạp (ô nhiều đoạn, có caption, không có hàng tiêu đề) | Grid table: nội dung một ô bị tách qua nhiều dòng | Bảng HTML `<table>` + `<caption>`: ranh giới ô rõ ràng |
| AI, hình | Ảnh + thuộc tính `{alt=…}` | `<figure>` gồm ảnh (có alt text) và `<figcaption>` |

Các giới hạn của bảng pipe trong GFM: mỗi ô chỉ một dòng, bắt buộc có hàng tiêu đề, không có
caption. Bảng nào vi phạm, pandoc sẽ xuất thành bảng HTML. Muốn bảng đơn giản ra bảng pipe
thì trong Word, hàng đầu tiên phải được đánh dấu "Repeat as header row".

**Yêu cầu với hệ thống phía sau (RAG/LLM):** phải giữ thẻ HTML trong markdown, và không được
cắt chunk giữa một `<table>` hay `<figure>`.

Muốn xuất Pandoc Markdown để so sánh: `run-all.bat input.docx -OutputFormat markdown`.

Gỡ bảng bọc (cleanup) vẫn **cần thiết** với GFM. Nếu không gỡ, hình và bảng dữ liệu trở
thành bảng HTML lồng trong một bảng HTML khác: người đọc vẫn thấy khung bao quanh, còn AI
không biết caption thuộc về hình hay bảng nào.

### `figures.lua`

Lua filter chạy trong pandoc, xử lý các ảnh do bước convert tạo ra:

| Việc | Trước | Sau |
|---|---|---|
| Title `shape2png:S001` (dùng để đối chiếu với manifest) chuyển thành thuộc tính | `"shape2png:S001"` | `data-shape="S001"` |
| Alt text chứa chữ trong hình: giữ lại cho LLM, bỏ dấu `\|` vốn gây xung đột với cú pháp bảng | `Drawing converted to image. Text: A \| B` | `Text in figure: A; B` |
| Bookmark trong caption của figure chuyển lên figure | `<span id="_Ref1" class="anchor">` nằm trong caption | `<figure id="_Ref1">` |

Khi ảnh đứng ngay trước một đoạn caption, pandoc tự ghép hai phần thành một figure.
Kết quả với GFM:

```html
<figure id="_Ref240186339">
<img src="images/media/image19.png" style="width:2.5in;height:2.38in" data-shape="S001" alt="Text in figure: Sound span; Not sound span; S1; S4; S3; S2" />
<figcaption><p>Fig. 7‑30 Transition of alarm sound status</p></figcaption>
</figure>
```

Tương tự, caption đứng ngay trước một bảng sẽ thành `<caption>` của bảng đó. Link tham chiếu
chéo (`[Fig. 7‑30](#_Ref240186339)`) trỏ tới `id` của figure hoặc của bảng.

`--wrap=none`: không ngắt dòng giữa cú pháp ảnh hay link.

Đã kiểm thử với pandoc 3.11, cả `gfm` và `markdown`. Filter cần pandoc ≥ 3.0.

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
pandoc -f docx -t gfm --wrap=none --extract-media=./images --lua-filter=.\pandoc\figures.lua .\input.shapes.docx -o output.md

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

Các điểm cải tiến và vấn đề đã biết nhưng chưa xử lý được theo dõi trong [BACKLOG.md](BACKLOG.md).

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

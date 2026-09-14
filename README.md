# Trợ lý Dữ liệu Quyết toán

Ứng dụng desktop Windows mở Trợ lý ảo bằng file BAT, tự nhận file JSON tải
xuống trong `Output`, kiểm tra dữ liệu và hỗ trợ xem/sửa/xác nhận kết quả.

Luồng hiện tại không còn mở URL, cấu hình trình duyệt, mở thư mục nhận file ở
Bước 1 hoặc chọn JSON thủ công.

## Chức năng chính

- Bước 1 chỉ có nút **Mở Trợ lý ảo**.
- File BAT do người dùng chọn trong Cài đặt và được chạy tách rời qua
  `cmd.exe`.
- Trước khi chạy BAT, ứng dụng cập nhật nguyên tử Chrome `Preferences` của
  profile trong bundle để tải thẳng về thư mục `Output`.
- Theo dõi `ket_qua_boc_tach*.json`, bỏ qua file tải tạm và chờ file ghi ổn
  định.
- Mỗi lượt tải thay thế file JSON trước đó và được đặt tên
  `ket_qua_boc_tach_YYYYMMDD_HHMMSS.json`.
- Kiểm tra schema và các quy tắc nghiệp vụ hiện có; hỗ trợ xem, thêm, sửa, xóa,
  xem JSON thô, lưu và xác nhận.
- Bước 2 hiển thị trạng thái file và thời điểm lưu gần nhất theo định dạng
  `HH:mm ngày dd/MM/yyyy`.
- Bước 3 cho phép chọn nhiều sheet tháng trong cùng một tác vụ khi đồng bộ
  Hàng ngày → BK, nhập khoản chi → BK và BK → Thanh toán. Một tháng lỗi sẽ
  chặn toàn bộ tác vụ; các workbook gốc chỉ được thay sau khi mọi tháng đã
  ghi và kiểm tra thành công.
- Khi nhập khoản chi, ứng dụng phân nhóm theo chứng từ nguồn/hóa đơn, đề xuất
  sheet từ ngày hóa đơn và để người dùng quyết định tháng làm mốc dò cuối cùng.
  Chỉ các sheet BK đã tồn tại được chọn cho luồng này.
- Với từng chứng từ/hóa đơn, ứng dụng tìm dòng trong sheet tháng được chọn và hai
  tháng liền trước, rồi ghi khoản chi ngay tại sheet chứa dòng tìm thấy. Ứng dụng
  không sao chép hoặc tạo dòng kế hoạch mới.
- Nguồn Bảng kê giữ cơ chế đối chiếu riêng trên toàn bộ các sheet tháng.
- Container xuất hiện nhiều lần trong phạm vi dò phải được chọn đúng sheet/dòng/SQT.
  Nhiều dòng JSON cùng nhắm một ô phí không được cộng; người dùng chọn đúng một
  dòng để ghi, còn ô đã có giá trị chỉ cho giữ nguyên, ghi đè hoặc bỏ qua.
- Bước 4 (RPA) vẫn chỉ cho chọn một sheet BK, tổng hợp các dòng theo SQT, chọn nhiều SQT và gọi
  BAT riêng để chạy PAD nhập khoản chi lên phần mềm quyết toán. Dòng `Đã nhập`
  vẫn được phép chạy lại.
- Mọi lần ghi BK đều dùng backup, working copy, kiểm tra lại và thay file
  nguyên tử; tác vụ chạy nền để không khóa giao diện.
- Dòng cước biển thiếu số cont có một hành động duy nhất **Đối soát số cont**.
  Người dùng chọn tháng/năm, ứng dụng tự chọn sheet `TMM YY` và tìm cont bằng
  **tên tàu + số chuyến**.
- Màn hình kiểm tra chỉ hiển thị một cột **Tàu/chuyến** ghép từ hai trường trên;
  khi sửa dòng, giá trị AI nguyên văn vẫn được giữ để truy vết.
- Nếu exact match không thấy, ứng dụng đối chiếu tên tàu và phần số/hậu tố chuyến
  trong đúng sheet BK. Chỉ một kết quả khác do thiếu tiền tố như `V.`, `MB`, `BS`
  hoặc `VPN` thì tự dùng giá trị canonical của BK; nhiều kết quả thì người dùng
  chọn một lần. Gợi ý gần đúng do sai tên tàu vẫn chỉ để tham khảo và sửa tay.
- Khi bắt đầu đối soát, người dùng có thể chọn nhiều hóa đơn cước biển cùng
  tàu/chuyến từ nhiều PDF để đóng góp vào một hồ sơ. Mỗi hóa đơn vẫn giữ tên
  chứng từ, số HĐ, ngày, số cont và số tiền riêng để truy vết.
- Cửa sổ đối soát hiển thị đồng thời danh sách HĐ, cont tìm được trong BK và
  bảng kết quả dự kiến. HĐ có thể thêm, sửa, xóa và lưu trực tiếp; hồ sơ chưa
  hoàn tất được giữ qua các lần khởi động mà không cần trung tâm chờ riêng.
- Tại màn hình kiểm tra file bóc tách, người dùng được xóa mọi loại cước dù batch
  đã có hồ sơ cước biển. Thao tác chỉ có hiệu lực khi bấm **Lưu**; HĐ cước biển
  bị xóa đồng thời được loại khỏi hồ sơ, hồ sơ rỗng được hủy nhưng vẫn giữ lịch
  sử, và HĐ vẫn được phép đối soát/nhập lại nếu xuất hiện ở lần sau.
- Hồ sơ trong cửa sổ đối soát được cố định theo đúng dòng đã chọn ở màn hình
  kiểm tra dữ liệu bóc tách. Muốn xem hoặc đối soát file khác, người dùng phải
  đóng hồ sơ hiện tại rồi bấm **Xem hồ sơ** hoặc **Đối soát** tại dòng tương ứng.
- Khi mở Trợ lý bóc tách trong hồ sơ, JSON mới tải về `Output` được tự kiểm tra.
  Chỉ HĐ cước biển cùng tàu/chuyến được thêm vào hồ sơ; HĐ trùng bị bỏ qua và
  các dòng khác vẫn nằm trong batch bóc tách mới.
- Mỗi lần mở Trợ lý tạo một phiên riêng và chỉ cho phép một phiên hoạt động. Khi
  JSON hợp lệ đã được tiếp nhận từ trang chủ, hoặc có ít nhất một HĐ phù hợp đã
  được thêm vào hồ sơ đối soát, phần mềm yêu cầu extension đóng đúng cửa sổ GPT
  của phiên đó. File lỗi hoặc file không có HĐ phù hợp không làm đóng cửa sổ.
- Chỉ khi tổng số cont trên HĐ bằng số cont hợp lệ duy nhất trong BK, phần mềm
  mới cho xác nhận. Tổng tiền của mọi HĐ được chia đều; phần lẻ dồn vào cont
  cuối. Mọi dòng kết quả ghi chung danh sách số HĐ, B/L và bên vận tải.
- Sau khi xác nhận đối soát, file bóc tách gốc vẫn là phiên làm việc hiện hành.
  Dòng cước biển gốc vẫn giữ nguyên một dòng và có nút **Xem hồ sơ**; JSON cont
  được lưu như batch con nội bộ trong `Output\_system\Ready`, không tự mở lên
  màn hình kiểm tra và không xuất hiện như một file bóc tách độc lập trong Lịch sử.
- Khi nhập khoản chi vào BK, ứng dụng tạo một gói dữ liệu duy nhất: giữ các khoản
  bình thường của file gốc, bỏ dòng cước biển nguyên bản đã được hồ sơ quản lý và
  ghép các dòng cont từ batch con hiện hành. Toàn bộ gói dùng chung một lần phân
  tích, một backup và một lần thay file BK; nguồn batch/dòng vẫn được lưu riêng để
  chống nhập trùng.
- Hồ sơ đã hoàn tất hoặc đã ghi BK được phép **Đối soát lại**. Mỗi lần tạo một
  phiên mới, giữ nguyên phiên và nhật ký cũ. Nếu BK đã có giá trị khác, mặc định
  giữ nguyên và chỉ ghi đè khi người dùng chọn rõ trong màn hình xử lý xung đột.
- HĐ cước biển xuất hiện lại trong một batch khác được đánh dấu cảnh báo màu vàng
  và liên kết về đúng hồ sơ tàu/chuyến hiện có. Liên kết này chỉ phục vụ truy vết,
  không cộng lại tiền hoặc số cont và không chặn người dùng chọn nhập lại vào BK.
- Khi một hồ sơ bị hủy, các HĐ được giải phóng để có thể đối soát lại và tạo
  kết quả mới. Hồ sơ, HĐ và nhật ký cũ vẫn được giữ trong SQLite để tra cứu;
  batch kết quả của hồ sơ đã hủy được lưu trữ và không còn được phép ghi BK.

## Yêu cầu

- Windows 10 hoặc Windows 11.
- Python 3.12 bản 64-bit.
- Google Chrome và bundle Trợ lý ảo đã giải nén đúng cấu trúc.
- Không cần quyền Administrator.

## Cài đặt và chạy

Nhấp đúp `run_app.bat`, hoặc chạy:

```bat
run_app.bat
```

Script kiểm tra Python 3.12 64-bit, tạo `.venv` nếu cần, cài dependency runtime
và mở ứng dụng. Lần đầu cần Internet để tải thư viện. Các lần sau không gọi
`pip` nếu `requirements.txt` không thay đổi và môi trường vẫn hợp lệ.

Để cài dependency phục vụ test:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
```

Mặc định toàn bộ dữ liệu nội bộ được đặt ngay trong thư mục chứa source hoặc
file `.exe`, nên có thể copy/đóng gói cả bundle mà không phụ thuộc AppData.
Có thể chủ động chọn vùng dữ liệu khác khi khởi động:

```powershell
.\run_app.bat --data-root "D:\DuLieu\TroLyQuyetToan"
$env:TRO_LY_DATA_ROOT = "D:\DuLieu\TroLyQuyetToan"
.\run_app.bat
```

Không đổi `data_root` khi ứng dụng đang chạy. Thư mục này dùng cho Config,
Database, Logs và cây dữ liệu nội bộ bên dưới `Output\_system`.

## Thiết lập lần đầu

1. Mở **Cài đặt**.
2. Chọn file BAT của bundle, ví dụ
   `RPA_ChatGPT_Launcher\Mo_Tro_Ly_RPA.bat`.
3. Giữ `Output` mặc định hoặc chọn thư mục khác.
4. Nhấn **Kiểm tra cấu hình**, sau đó **Lưu cấu hình**.
5. Về trang Quy trình và nhấn **Mở Trợ lý ảo**.
6. Làm việc trong cửa sổ Trợ lý ảo rồi nhấn tải file kết quả. Ứng dụng sẽ tự
   nhận file; không cần chọn file bằng tay.

Để dùng Bước 3, cấu hình thêm:

- **File Hàng ngày trên LAN** (`.xlsx` hoặc `.xlsm`).
- **File BK Tổng hợp cá nhân** (`.xlsx` hoặc `.xlsm`).

Để dùng Bước 4, chọn thêm **BAT RPA nhập quyết toán**. Phần mềm luôn ghi dữ liệu
PAD vào `Output\_system\RPA\rpa_input_selection.json`. Hợp đồng dữ liệu và lệnh
PAD gọi sau khi web lưu thành công được mô tả tại
[`docs/RPA_FLOW4_CONTRACT.md`](docs/RPA_FLOW4_CONTRACT.md).

Nút **Kiểm tra cấu hình** chỉ đọc workbook/bản sao tạm, không lưu thử vào file
gốc. Định dạng Excel cũ `.xls` không được hỗ trợ.

Không còn BAT hoặc Custom GPT riêng để bóc số container. Khi cần bổ sung hóa
đơn, nút trong hồ sơ **Đối soát số cont** mở lại **Trợ lý ảo chính**; JSON mới
tiếp tục qua watcher và chỉ các HĐ đúng tàu/chuyến được nhập vào hồ sơ.

Bundle BAT/PowerShell và extension phối hợp với phần mềm bằng một mã phiên ngẫu
nhiên qua kết nối nội bộ `127.0.0.1`. Phần mềm quyết định khi nào dữ liệu đã được
tiếp nhận; extension dùng mã phiên để đóng đúng cửa sổ GPT, không tắt các cửa sổ
Chrome khác. Chrome `Preferences` trong `RPA_ChatGPT_Profile\Default` vẫn được cập
nhật để đặt:

```json
{
  "download": {
    "default_directory": "<đường dẫn Output>",
    "prompt_for_download": false,
    "directory_upgrade": true
  }
}
```

Các khóa khác trong `Preferences` được giữ nguyên.

## Hợp đồng JSON

Schema ghi công khai hiện tại là v3:

```json
{
  "v": 3,
  "d": [
    {
      "source_document_id": "DOC_001",
      "source_document_name": "Hoa_don_A.pdf",
      "container": null,
      "bl": "OOLU1234567890",
      "vessel_voyage_raw": "PROSPER 2625S",
      "vessel_name": "PROSPER",
      "voyage_no": "2625S",
      "invoice_container_count": 4,
      "container_count_basis": "EXPLICIT",
      "fee": "CB",
      "rule": "HD",
      "invoice_no": "INV-001",
      "invoice_date": "2026-07-15",
      "carrier": "HÃNG TÀU",
      "amount": 32000000
    }
  ]
}
```

Root chỉ có `v` và `d`; mỗi object v3 phải có đủ đúng 15 khóa theo thứ tự trên.
Hai trường nguồn phải là chuỗi không rỗng; dữ liệu nghiệp vụ thiếu dùng `null`.
`voyage_no` luôn là chuỗi; `invoice_container_count` là số nguyên dương hoặc
`null`; căn cứ nhận `EXPLICIT`, `CALCULATED`, `UNKNOWN`.

Schema v1 mảng 7 vị trí và schema v2 object 13 khóa vẫn được đọc để tương thích.
Dữ liệu cũ được gắn nguồn legacy theo batch. Sau khi user lưu/chỉnh sửa, phần
mềm luôn serialize thành v3; các file JSON lịch sử không bị sửa hàng loạt.

Không dùng các root `metadata`, `du_lieu_boc_tach`, `canh_bao` hoặc
`raw_data`.

Khi đối soát, phần mềm tự đọc sheet BK theo tháng đã chọn, kiểm tra ISO 6346,
bỏ trùng container và lưu snapshot. Số cont sai/trùng được báo bằng tiếng Việt.
Trước khi xác nhận và trước khi ghi dữ liệu, BK được đọc lại; nếu workbook hoặc
danh sách container đổi, hồ sơ bị khóa và yêu cầu người dùng kiểm tra lại.

Hai bản instructions Custom GPT tương ứng nằm tại
`gpt_custom_instructions_chi_tiet.txt` và `gpt_custom_instructions_ngan.txt`;
bản ngắn được giữ dưới 8.000 ký tự. Chỉ cài instructions v3 sau khi ứng dụng đã
được nâng cấp hỗ trợ v3, hoặc phát hành đồng thời cả hai.

## Vòng đời file Output

Thư mục mặc định là:

```text
<thư mục chứa main.py hoặc .exe>\Output
```

Khi có file hoàn chỉnh mới:

1. Các file kết quả JSON cũ trong `Output` bị loại bỏ.
2. File mới được đặt tên theo timestamp nhận file.
3. Nội dung được kiểm tra và tự động mở trên màn hình review.

Quy tắc mới nhất luôn thắng. Nếu hai lượt tải chồng nhau, ứng dụng bỏ qua sự
kiện cũ. Nếu file mới sai schema, file cũ vẫn bị loại bỏ theo quy tắc đã chốt
và Bước 2 hiển thị nguyên nhân lỗi, không hiển thị lưu thành công.

Khi lưu từ màn hình review, ứng dụng ghi nguyên tử trực tiếp vào một file JSON
timestamp mới rồi xóa tên cũ. Không tạo archive, working copy, `.bak` hoặc
snapshot Ready. Nếu file mới được tải khi cửa sổ cũ còn thay đổi chưa lưu, dữ
liệu cũ bị bỏ và cửa sổ review chuyển ngay sang dữ liệu mới.

## Thư mục dữ liệu

Mặc định dữ liệu nội bộ nằm cùng thư mục ứng dụng:

```text
khoanchi_pm_project\                 (hoặc thư mục chứa file .exe)
├── Config\settings.json
├── Database\app_state.db
├── Logs\
└── Output\
    ├── ket_qua_boc_tach_YYYYMMDD_HHMMSS.json
    ├── <tên bất kỳ>.json
    └── _system\
        ├── Excel\
        │   ├── Temp\
        │   ├── Backup\
        │   └── Reports\
        └── RPA\
```

Khi cả bundle được chuyển vị trí, ứng dụng tự cập nhật `data_root`, `Output`
mặc định và các đường dẫn batch trong SQLite sang vị trí mới. Đường dẫn ngoài
bundle do người dùng tự chọn vẫn được giữ nguyên.

- `Output\ket_qua_boc_tach_YYYYMMDD_HHMMSS.json`: JSON hiện hành duy nhất.
- `Output\_system\Excel\Temp`: snapshot nguồn và working copy ngắn hạn.
- `Output\_system\Excel\Backup`: backup BK tạo ngay trước mỗi lần ghi.
- `Output\_system\Excel\Reports`: vùng dành cho báo cáo xử lý Excel.
- `Output\_system\RPA`: JSON đầu vào riêng cho từng lượt chạy PAD.
- `Output\<tên bất kỳ>.json`: có thể là kết quả container nếu nội dung đúng
  contract `{"containers": [...]}`; tên file không được dùng để phân loại.
- `Database`: metadata batch và trạng thái ứng dụng.

## Chạy kiểm thử

```powershell
.\.venv\Scripts\python.exe -m pytest -q
```

Test sử dụng thư mục tạm của pytest, không ghi vào Output hay database vận hành
của người dùng.

## Xử lý lỗi thường gặp

### Không mở được Trợ lý ảo

Kiểm tra đã chọn file `.bat`, file nằm trong đúng bundle và các thư mục launcher,
extension vẫn đầy đủ. Dùng **Kiểm tra cấu hình** để xem thông báo cụ thể.

### Không nhận file

Kiểm tra tên file khớp `ket_qua_boc_tach*.json`, file đã tải xong và Output
trong Cài đặt đúng với thư mục đang được theo dõi. Các hậu tố `.crdownload`,
`.part`, `.tmp` và `.download` được xem là file tạm.

### JSON sai schema

File phải có đúng root `{v, d}` và tuân theo schema v3 object 15 khóa; schema
v2 object 13 khóa và schema v1 mảng 7 vị trí chỉ được đọc theo chế độ tương
thích. Bước 2 hiển thị lỗi cấu trúc và khóa xem/sửa cho đến khi có file mới đọc
được.

### Không có quyền ghi

Chọn Output trong vùng người dùng có quyền ghi; tránh `Program Files`, thư mục
hệ thống và thư mục mạng chỉ đọc. Kiểm tra Controlled Folder Access nếu quyền
Windows có vẻ đúng nhưng ứng dụng vẫn báo lỗi.

## API cho mô-đun tích hợp sau

`app.services.reviewed_batch_provider.ReviewedBatchProvider` cung cấp:

```text
get_latest_ready_json_path() -> Path | None
get_ready_json_path(batch_id: int) -> Path | None
list_ready_batches() -> list[BatchMetadata]
```

Các tên API cũ được giữ để tương thích, nhưng path trả về trỏ tới file JSON
hiện hành đã xác nhận trong `Output`.

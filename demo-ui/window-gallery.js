"use strict";

const descriptions = {
  review:{title:"Kiểm tra dữ liệu",subtitle:"Mở từ nút “Xem file bóc tách” ở bước 2.",kept:"Giữ bảng khoản chi, lọc/tìm kiếm, thêm/sửa/xóa dòng, xem nguồn và nút lưu xác nhận.",improved:"Thông tin chứng từ và lỗi cần sửa nổi bật hơn; Batch ID, SHA-256 và JSON thô rời khỏi phần đầu cửa sổ."},
  edit:{title:"Sửa khoản chi",subtitle:"Mở khi chọn một dòng ở màn kiểm tra.",kept:"Giữ các trường hóa đơn, container, loại phí, số tiền và quy tắc kiểm tra dữ liệu.",improved:"Nhóm trường theo cách người dùng đọc chứng từ; ghi ngày theo ngày/tháng/năm; lỗi gắn ngay với trường cần sửa."},
  reconcile:{title:"Đối soát số container",subtitle:"Mở từ dòng cước biển cần đối soát.",kept:"Giữ hồ sơ tàu/chuyến, hóa đơn, nguồn BK, số container và kết quả phân bổ.",improved:"Đặt hai số lượng cạnh nhau, nói rõ thiếu ở phía nào; nút xác nhận chỉ rõ khi nào kết quả đã đủ điều kiện."},
  month:{title:"Chọn tháng Excel",subtitle:"Xuất hiện trước bước phân tích Excel như hiện tại.",kept:"Giữ khả năng chọn nhiều sheet tháng và ghi nhớ lựa chọn gần nhất.",improved:"Viết rõ tháng nguồn, tháng đích và việc sẽ xảy ra sau khi bấm Phân tích; giảm thuật ngữ “ánh xạ”."},
  conflict:{title:"Xử lý chênh lệch",subtitle:"Xuất hiện trong tác vụ “Khoản chi → BK” khi dữ liệu chưa khớp.",kept:"Giữ các lựa chọn giữ giá trị, ghi đè hoặc bỏ qua, cùng thông tin dòng/ô để truy vết.",improved:"Hiện tiền hiện có và tiền từ chứng từ cạnh nhau; đưa sheet, dòng, ô vào chi tiết thay vì dàn trên bảng 15 cột."},
  confirmation:{title:"Xác nhận cập nhật Excel",subtitle:"Bước xem trước trước khi ghi vào workbook.",kept:"Giữ thống kê sẽ ghi, ghi đè, giữ nguyên, bỏ qua và cảnh báo thay đổi.",improved:"Một câu tóm tắt nói rõ file nào sẽ thay đổi; cảnh báo ghi đè nổi bật; tên nút gọi đúng thao tác."},
  rpa:{title:"Chọn số quyết toán",subtitle:"Mở sau khi chọn sheet BK ở bước 4.",kept:"Giữ bảng số quyết toán, trạng thái đã nhập/chưa nhập và khả năng chọn chạy lại.",improved:"Tổng số mục và tổng tiền đã chọn luôn hiện trước nút chạy; phân biệt “đã khởi chạy” với “đã nhập thành công”."}
};
const demo = {view:"review",selectedRow:3,search:"",invoiceDate:"",invoiceAmount:"1200000",editDateDraft:"",editAmountDraft:"1200000",reconMoreMonth:false,confirmed:false,monthChoice:"T09 26",choices:{a:"",b:""},selectedJobs:new Set(),toastTimer:null};
const preview=document.getElementById("gallery-preview");
const esc=value=>String(value??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[c]));
const money=value=>new Intl.NumberFormat("vi-VN").format(value)+" đ";
const pill=(label,tone)=>`<span class="pill ${tone}">${label}</span>`;
const windowFrame=(title,subtitle,body,footer="",size="")=>`<section class="mock-window ${size}"><header class="mock-titlebar"><div><h2>${title}</h2><p>${subtitle}</p></div><span class="mock-close" aria-hidden="true">×</span></header><div class="mock-body">${body}</div>${footer?`<footer class="mock-footer">${footer}</footer>`:""}</section>`;
const footer=(note,buttons)=>`<p>${note}</p><div class="mock-footer-actions">${buttons}</div>`;

function reviewView(){
  const rows=[
    {id:1,document:"Hóa đơn CMA-CG-201",fee:"Cước biển",container:"CMAU7312094",amount:12750000,status:"Đã kiểm tra",tone:"success",issue:"Thông tin đã đầy đủ."},
    {id:2,document:"Hóa đơn CMA-CG-201",fee:"Phí vệ sinh",container:"CMAU7312094",amount:550000,status:"Cần xem lại",tone:"warning",issue:"Số tiền từ chứng từ khác giá trị đang có trong BK."},
    {id:3,document:"Hóa đơn ONE-0926",fee:"Phí nâng hạ",container:"OOLU9182736",amount:Number(demo.invoiceAmount.replace(/[^0-9]/g,""))||1200000,status:demo.invoiceDate?"Đã kiểm tra":"Cần sửa",tone:demo.invoiceDate?"success":"error",issue:demo.invoiceDate?"Đã bổ sung ngày hóa đơn. Hãy đối chiếu lại chứng từ gốc.":"Thiếu ngày hóa đơn. Hãy bổ sung trước khi xác nhận."},
    {id:4,document:"Hóa đơn ONE-0926",fee:"Cước biển",container:demo.reconMoreMonth?"4 container":"Chưa có",amount:2300000,status:demo.reconMoreMonth?"Đã đối soát":"Cần đối soát",tone:demo.reconMoreMonth?"success":"warning",issue:demo.reconMoreMonth?"Đã tìm đủ 4 container trong các tháng BK được chọn.":"Thiếu số container cho cước biển. Hãy đối soát số cont."}
  ];
  const shown=rows.filter(row=>`${row.document} ${row.fee} ${row.container}`.toLocaleLowerCase("vi").includes(demo.search.toLocaleLowerCase("vi")));
  const selected=rows.find(row=>row.id===demo.selectedRow)||rows[0];
  const body=`<div class="mock-banner"><div><strong>Dữ liệu chứng từ tháng 09/2026</strong><br><span>2 chứng từ · 4 khoản chi minh họa · Nhận lúc 09:30</span></div><div class="mock-summary">${pill(`${demo.invoiceDate?0:1} cần sửa`,demo.invoiceDate?"success":"error")}${pill(`${demo.reconMoreMonth?1:2} cần xem lại`,"warning")}</div></div>
  <div class="mock-toolbar"><div class="mock-search"><input id="gallery-search" type="search" placeholder="Tìm hóa đơn, container, loại phí…" value="${esc(demo.search)}"></div><div class="mock-flex"><button class="btn small" data-action="add-example">Thêm dòng</button><button class="btn small" data-view="edit">Sửa dòng</button><button class="btn small danger" data-action="delete-example">Xóa dòng</button></div></div>
  <div class="mock-review-grid"><div class="mock-table-wrap"><table class="mock-table"><thead><tr><th>Chứng từ</th><th>Khoản chi</th><th style="text-align:right">Số tiền</th><th>Trạng thái</th></tr></thead><tbody>${shown.map(row=>`<tr data-action="pick-row" data-id="${row.id}" class="${row.id===selected.id?"chosen":""}"><td><span class="main">${row.document}</span><span class="sub">Hóa đơn ${row.id<3?"000321":"000849"}</span></td><td>${row.fee}<span class="sub">${row.container}</span></td><td class="amount">${money(row.amount)}</td><td>${pill(row.status,row.tone)}</td></tr>`).join("")}</tbody></table></div><aside class="mock-detail"><h3>${selected.fee}</h3><div class="caption">${selected.document}</div><div class="detail-pair"><span>Số tiền</span><strong>${money(selected.amount)}</strong></div><div class="detail-pair"><span>Container</span><strong>${selected.container}</strong></div><div class="callout ${selected.tone}">${selected.issue}</div><button class="btn small primary" data-view="${selected.id===4?"reconcile":"edit"}">${selected.id===4?"Đối soát số cont":"Sửa thông tin"}</button></aside></div>`;
  return windowFrame("Kiểm tra dữ liệu bóc tách","Cửa sổ mở từ bước 2 · Giữ cách kiểm tra từng dòng",body,footer(demo.confirmed?"Dữ liệu đã được xác nhận trong bản demo.":demo.invoiceDate&&demo.reconMoreMonth?"Đã xử lý các mục chặn. Có thể lưu và xác nhận.":"Cần sửa hóa đơn và đối soát cước biển trước khi xác nhận.",'<button class="btn small" data-action="technical-info">Thông tin kỹ thuật</button><button class="btn small primary" data-action="review-confirm">Lưu và xác nhận</button>'));
}

function editView(){
  const body=`<div class="mock-banner"><div><strong>Hóa đơn ONE-0926 · Khoản phí nâng hạ</strong><br><span>Chỉnh sửa đúng dòng đang chọn ở màn kiểm tra.</span></div>${pill(demo.invoiceDate?"Đã kiểm tra":"Cần sửa",demo.invoiceDate?"success":"error")}</div>
  <h3 class="mock-section-title">Thông tin chứng từ</h3><div class="mock-form-grid"><div class="mock-field"><label for="edit-doc">Số hóa đơn</label><input id="edit-doc" value="000849"></div><div class="mock-field"><label for="edit-date">Ngày hóa đơn</label><input id="edit-date" type="date" value="${esc(demo.editDateDraft)}"><div class="mock-field-hint">Kiểm tra ngày trên hóa đơn gốc.</div></div><div class="mock-field"><label for="edit-container">Số container</label><input id="edit-container" value="OOLU9182736"></div><div class="mock-field"><label for="edit-bl">Vận đơn (B/L)</label><input id="edit-bl" value="OOLU9921845"></div></div>
  <hr class="mock-divider"><h3 class="mock-section-title">Khoản chi</h3><div class="mock-form-grid"><div class="mock-field"><label for="edit-fee">Loại phí</label><select id="edit-fee"><option>Phí nâng hạ</option><option>Cước biển</option><option>Phí lưu container</option></select></div><div class="mock-field"><label for="edit-money">Số tiền (VND)</label><input id="edit-money" inputmode="numeric" value="${esc(demo.editAmountDraft)}"><div class="mock-field-hint">Ví dụ: 1.200.000</div></div></div><div id="edit-feedback" class="callout ${demo.editDateDraft?"success":"error"}" style="margin-bottom:0">${demo.editDateDraft?"Thông tin bắt buộc đã đầy đủ. Hãy đối chiếu lại với hóa đơn trước khi lưu.":"Thiếu ngày hóa đơn. Hãy bổ sung để lưu khoản chi."}</div>`;
  return windowFrame("Sửa thông tin khoản chi","Hóa đơn 000849 · Phí nâng hạ",body,footer("Các quy tắc kiểm tra dữ liệu hiện có vẫn được áp dụng.",'<button class="btn" data-view="review">Hủy</button><button class="btn primary" data-action="save-edit">Lưu dòng</button>'),"compact");
}

function reconcileView(){
  const found=demo.reconMoreMonth?4:3;
  const body=`<div class="mock-banner"><div><strong>Hồ sơ tàu/chuyến: ONE HARMONY 078W</strong><br><span>Đối soát hóa đơn cước biển với dữ liệu container trong BK.</span></div>${pill(found===4?"Đã khớp số lượng":"Chưa khớp số lượng",found===4?"success":"warning")}</div>
  <div class="mock-metrics"><div class="mock-metric"><small>HÓA ĐƠN</small><strong>2</strong></div><div class="mock-metric"><small>SỐ CONT TRÊN HĐ</small><strong>4</strong></div><div class="mock-metric"><small>CONT TÌM ĐƯỢC TRONG BK</small><strong>${found}</strong></div><div class="mock-metric"><small>CHÊNH LỆCH</small><strong style="color:${found===4?"#087f78":"#b86e2c"}">${4-found}</strong></div></div>
  ${found===4?'<div class="callout success">Số container trên hóa đơn đã khớp với BK. Hãy kiểm tra danh sách bên dưới trước khi xác nhận.</div>':'<div class="callout warning">Hóa đơn ghi 4 container; BK đang tìm thấy 3 trong tháng đã chọn. Kiểm tra thêm tháng đối soát hoặc thông tin tàu/chuyến.</div>'}
  <div class="mock-two-col"><section class="mock-panel"><h3>Hóa đơn cước biển</h3><p>Giữ từng hóa đơn riêng để có thể truy vết.</p><div class="mock-mini-list"><div class="mock-mini-row"><div><strong>HĐ 000849</strong><small>2 cont · 12.000.000 đ</small></div>${pill("Đã nhận","neutral")}</div><div class="mock-mini-row"><div><strong>HĐ 000850</strong><small>2 cont · 10.000.000 đ</small></div>${pill("Đã nhận","neutral")}</div></div></section><section class="mock-panel"><h3>Nguồn bảng kê</h3><p>Chọn đúng tháng dữ liệu để dò số container.</p><div class="mock-mini-list"><div class="mock-mini-row"><div><strong>T09 26</strong><small>ONE HARMONY 078W · 3 cont</small></div>${pill("Đang dùng","success")}</div>${demo.reconMoreMonth?'<div class="mock-mini-row"><div><strong>T08 26</strong><small>ONE HARMONY 078W · 1 cont</small></div>'+pill("Đã thêm","success")+'</div>':''}</div><button class="btn small" style="margin-top:11px" data-action="add-month">${demo.reconMoreMonth?"Bỏ tháng bổ sung":"Thêm tháng đối soát"}</button></section></div>
  <div class="mock-panel" style="margin-top:13px"><h3>Kết quả dự kiến</h3><p>${found===4?"4 container sẽ được phân bổ từ 2 hóa đơn.":"Chưa tạo kết quả cuối cùng khi số lượng còn lệch."}</p><div class="mock-progress"><span style="width:${found/4*100}%"></span></div></div>`;
  return windowFrame("Đối soát số container","Cước biển · ONE HARMONY 078W",body,footer("Chỉ xác nhận khi hóa đơn và BK khớp số lượng.",`<button class="btn" data-action="refresh-bk">Đọc lại BK</button><button class="btn primary" data-action="finish-reconcile" ${found===4?"":"disabled"}>Xác nhận đối soát</button>`),"medium");
}

function monthView(){
  const body=`<div class="callout info" style="margin-top:0">Chọn sheet BK cho từng tháng nguồn. Các lựa chọn gần nhất có thể được điền sẵn để giảm thao tác lặp.</div><div class="mock-mini-list"><div class="mock-mini-row"><div><strong>Hàng ngày · Tháng 09/2026</strong><small>58 dòng dữ liệu vận chuyển</small></div><select aria-label="Tháng BK cho tháng 09"><option>T09 26 · Tháng 09/2026</option><option>T08 26 · Tháng 08/2026</option></select></div><div class="mock-mini-row"><div><strong>Hàng ngày · Tháng 08/2026</strong><small>46 dòng dữ liệu vận chuyển</small></div><select aria-label="Tháng BK cho tháng 08"><option>T08 26 · Tháng 08/2026</option><option>T09 26 · Tháng 09/2026</option></select></div></div><div class="mock-info" style="margin-top:14px">Sau khi bấm <strong>Phân tích</strong>, phần mềm kiểm tra hai tháng đã chọn và cho xem trước kết quả. File gốc chưa thay đổi ở bước này.</div>`;
  return windowFrame("Chọn tháng nhận dữ liệu","Tác vụ Hàng ngày → BK",body,footer("Đã điền sẵn lựa chọn từ lần xử lý gần nhất.",'<button class="btn" data-action="reset-months">Đặt lại</button><button class="btn primary" data-action="analyze-months">Phân tích</button>'),"compact");
}

function conflictView(){
  const entries=[{id:"a",title:"Phí vệ sinh · CMAU7312094",invoice:"Hóa đơn 000321",current:500000,newValue:550000},{id:"b",title:"Phí lưu container · OOLU9182736",invoice:"Hóa đơn 000849",current:2000000,newValue:2300000}];
  const remaining=entries.filter(item=>!demo.choices[item.id]).length;
  const body=`<div class="mock-banner"><div><strong>${remaining?`Còn ${remaining} khoản cần chọn cách xử lý`:"Đã chọn cách xử lý cho tất cả khoản"}</strong><br><span>So sánh tiền đang có trong BK với tiền từ chứng từ.</span></div>${pill(`${entries.length-remaining}/${entries.length} đã chọn`,remaining?"warning":"success")}</div>
  <div class="mock-two-col"><div>${entries.map(item=>`<div class="mock-choice"><h3>${item.title}</h3><p>${item.invoice} · BK tháng 09/2026</p><div class="mock-compare"><div><small>ĐANG CÓ TRONG BK</small><strong>${money(item.current)}</strong></div><div><small>TỪ CHỨNG TỪ</small><strong>${money(item.newValue)}</strong></div></div><select data-conflict="${item.id}" aria-label="Cách xử lý ${item.title}"><option value="" ${!demo.choices[item.id]?"selected":""}>Chọn cách xử lý…</option><option value="keep" ${demo.choices[item.id]==="keep"?"selected":""}>Giữ số tiền trong BK</option><option value="replace" ${demo.choices[item.id]==="replace"?"selected":""}>Dùng số tiền từ chứng từ</option><option value="skip" ${demo.choices[item.id]==="skip"?"selected":""}>Bỏ qua khoản này</option></select></div>`).join("")}</div><aside class="mock-panel"><h3>Chi tiết khi cần đối chiếu</h3><p>Thông tin Excel vẫn có để truy vết, nhưng không chen vào phần ra quyết định.</p><div class="detail-pair"><span>Sheet</span><strong>T09 26</strong></div><div class="detail-pair"><span>Dòng</span><strong>184</strong></div><div class="detail-pair"><span>Ô tiền</span><strong>N184</strong></div><div class="callout info">Các lựa chọn chỉ áp dụng khi bạn xác nhận cập nhật Excel ở bước sau.</div></aside></div>`;
  return windowFrame("Xử lý các khoản chênh lệch","Tác vụ Khoản chi → BK",body,footer("Cần chọn cách xử lý cho từng khoản trước khi tiếp tục.",`<button class="btn" data-view="month">Quay lại chọn tháng</button><button class="btn primary" data-view="confirmation" ${remaining?"disabled":""}>Xem tóm tắt cập nhật</button>`),"medium");
}

function confirmationView(){
  const body=`<div class="mock-banner"><div><strong>File BK Tổng hợp tháng 09/2026 sẽ được cập nhật</strong><br><span>Dữ liệu từ 2 chứng từ · 1 sheet tháng · Xem trước trước khi ghi.</span></div>${pill("Cần xác nhận","warning")}</div><div class="mock-metrics"><div class="mock-metric"><small>KHOẢN SẼ GHI</small><strong>4</strong></div><div class="mock-metric"><small>GHI MỚI</small><strong>2</strong></div><div class="mock-metric"><small>GHI ĐÈ</small><strong style="color:#af6c29">1</strong></div><div class="mock-metric"><small>GIỮ NGUYÊN</small><strong>1</strong></div></div><div class="callout warning">1 khoản sẽ thay thế số tiền đang có trong BK. Hãy kiểm tra lựa chọn ở màn xử lý chênh lệch.</div><div class="mock-panel"><h3>Kết quả theo tháng</h3><div class="mock-table-wrap"><table class="mock-table"><thead><tr><th>Sheet BK</th><th>Ghi mới</th><th>Ghi đè</th><th>Giữ nguyên</th><th>Bỏ qua</th></tr></thead><tbody><tr><td class="main">T09 26</td><td>2</td><td>1</td><td>1</td><td>0</td></tr></tbody></table></div></div>`;
  return windowFrame("Xác nhận nhập khoản chi vào BK","Bước cuối trước khi thay đổi file Excel",body,footer("Bản demo không ghi file BK.",'<button class="btn" data-view="conflict">Sửa lựa chọn</button><button class="btn primary" data-action="confirm-excel">Xác nhận cập nhật BK</button>'),"compact");
}

function rpaView(){
  const jobs=[{id:"2451",name:"CMA CGM",amount:13300000,status:"Chưa nhập",tone:"neutral"},{id:"2452",name:"Ocean Network Express",amount:3500000,status:"Chưa nhập",tone:"neutral"},{id:"2448",name:"DHL Global Forwarding",amount:8870000,status:"Đã nhập",tone:"success"}];
  const picked=jobs.filter(job=>demo.selectedJobs.has(job.id));
  const body=`<div class="mock-banner"><div><strong>BK tháng 09/2026 · Sheet T09 26</strong><br><span>Chọn một hoặc nhiều số quyết toán để nhập lên phần mềm.</span></div>${pill("1 số đã nhập","success")}</div><div class="mock-toolbar"><div class="mock-select-actions"><button class="btn small" data-action="select-pending">Chọn số chưa nhập</button><button class="btn small" data-action="clear-jobs">Bỏ chọn</button></div><span class="mock-info">Có thể chọn chạy lại số đã nhập.</span></div><div class="mock-table-wrap"><table class="mock-table"><thead><tr><th>Chọn</th><th>Số quyết toán</th><th>Bên vận tải</th><th style="text-align:right">Tổng khoản chi</th><th>Trạng thái</th></tr></thead><tbody>${jobs.map(job=>`<tr><td><input type="checkbox" data-job="${job.id}" ${demo.selectedJobs.has(job.id)?"checked":""} aria-label="Chọn số quyết toán ${job.id}"></td><td class="main">${job.id}</td><td>${job.name}</td><td class="amount">${money(job.amount)}</td><td>${pill(job.status,job.tone)}</td></tr>`).join("")}</tbody></table></div><div class="callout ${picked.some(job=>job.status==="Đã nhập")?"warning":"info"}" style="margin-bottom:0">${picked.length?`Đã chọn ${picked.length} số quyết toán · Tổng khoản chi ${money(picked.reduce((sum,job)=>sum+job.amount,0))}.${picked.some(job=>job.status==="Đã nhập")?" Có số đã nhập được chọn để chạy lại.":""}`:"Chọn số quyết toán để xem tổng tiền của lượt nhập."}</div>`;
  return windowFrame("Chọn số quyết toán cần nhập","Mở sau khi chọn sheet BK ở bước 4",body,footer("Chỉ báo “Đã nhập” sau khi thao tác trên phần mềm đích thành công.",`<button class="btn" data-action="clear-jobs">Hủy chọn</button><button class="btn primary" data-action="preview-rpa" ${picked.length?"":"disabled"}>Xem trước lượt nhập</button>`),"medium");
}

const renderers={review:reviewView,edit:editView,reconcile:reconcileView,month:monthView,conflict:conflictView,confirmation:confirmationView,rpa:rpaView};
function render(){const item=descriptions[demo.view];document.getElementById("gallery-title").textContent=item.title;document.getElementById("gallery-subtitle").textContent=item.subtitle;document.getElementById("gallery-kept").textContent=item.kept;document.getElementById("gallery-improved").textContent=item.improved;document.querySelectorAll(".gallery-nav-item").forEach(button=>button.classList.toggle("active",button.dataset.view===demo.view));preview.innerHTML=renderers[demo.view]();}
function show(view){if(!renderers[view])return;if(view==="edit"){demo.editDateDraft=demo.invoiceDate;demo.editAmountDraft=demo.invoiceAmount;}demo.view=view;render();document.querySelector(".gallery-main").scrollIntoView({block:"start",behavior:"smooth"});}
function notify(message){document.querySelector(".gallery-toast")?.remove();const toast=document.createElement("div");toast.className="gallery-toast";toast.textContent=message;document.body.append(toast);clearTimeout(demo.toastTimer);demo.toastTimer=setTimeout(()=>toast.remove(),3800);}

document.addEventListener("click",event=>{
  const viewButton=event.target.closest("[data-view]");if(viewButton){show(viewButton.dataset.view);return;}
  const target=event.target.closest("[data-action]");if(!target)return;
  const action=target.dataset.action;
  if(action==="pick-row"){demo.selectedRow=Number(target.dataset.id);render();return;}
  if(action==="save-edit"){
    const value=document.getElementById("edit-date")?.value||"";
    if(!value){notify("Hãy bổ sung ngày hóa đơn trước khi lưu.");return;}
    demo.invoiceDate=value;demo.invoiceAmount=document.getElementById("edit-money")?.value||"1200000";demo.confirmed=false;notify("Đã lưu thông tin minh họa. Dữ liệu thật chưa thay đổi.");show("review");return;
  }
  if(action==="add-month"){demo.reconMoreMonth=!demo.reconMoreMonth;demo.confirmed=false;render();return;}
  if(action==="finish-reconcile"){notify("Đã xem trước kết quả đối soát. Dữ liệu thật chưa thay đổi.");return;}
  if(action==="analyze-months"){show("conflict");return;}
  if(action==="select-pending"){demo.selectedJobs=new Set(["2451","2452"]);render();return;}
  if(action==="clear-jobs"){demo.selectedJobs.clear();render();return;}
  if(action==="preview-rpa"){
    const jobs={"2451":13300000,"2452":3500000,"2448":8870000};
    const total=[...demo.selectedJobs].reduce((sum,id)=>sum+jobs[id],0);
    notify(`Lượt nhập mẫu: ${demo.selectedJobs.size} số quyết toán · ${money(total)}. Không gửi dữ liệu.`);return;
  }
  if(action==="review-confirm"){if(!demo.invoiceDate||!demo.reconMoreMonth){notify("Còn khoản cần sửa hoặc cước biển cần đối soát trước khi xác nhận.");return;}demo.confirmed=true;render();notify("Đã xác nhận dữ liệu trong bản demo. File vận hành chưa thay đổi.");return;}
  if(action==="confirm-excel"){notify("Đây là bản xem trước. Không có file Excel nào được ghi.");return;}
  if(action==="reset-months"){notify("Đã đặt lại lựa chọn tháng trong bản demo.");return;}
  if(action==="technical-info"){notify("Mã batch, SHA-256 và JSON thô được đặt trong phần chi tiết kỹ thuật.");return;}
  if(action==="delete-example"||action==="add-example"||action==="refresh-bk"){notify("Vị trí thao tác được giữ như phần mềm hiện tại; bản demo không thay đổi dữ liệu.");}
});
document.addEventListener("change",event=>{
  const target=event.target;
  if(target.matches("select[data-conflict]")){demo.choices[target.dataset.conflict]=target.value;render();}
  if(target.matches("input[data-job]")){target.checked?demo.selectedJobs.add(target.dataset.job):demo.selectedJobs.delete(target.dataset.job);render();}
  if(target.id==="edit-date"){demo.editDateDraft=target.value;const feedback=document.getElementById("edit-feedback");if(feedback){feedback.className="callout "+(target.value?"success":"error");feedback.textContent=target.value?"Thông tin bắt buộc đã đầy đủ. Hãy đối chiếu lại với hóa đơn trước khi lưu.":"Thiếu ngày hóa đơn. Hãy bổ sung để lưu khoản chi.";}}
});
document.addEventListener("input",event=>{if(event.target.id==="gallery-search"){demo.search=event.target.value;const position=event.target.selectionStart;render();const input=document.getElementById("gallery-search");input.focus();input.setSelectionRange(position,position);}if(event.target.id==="edit-money")demo.editAmountDraft=event.target.value;});
if(renderers[location.hash.slice(1)])demo.view=location.hash.slice(1);
render();

var DocTienBangChu = function () {
    this.ChuSo = new Array(" không ", " một ", " hai ", " ba ", " bốn ", " năm ", " sáu ", " bảy ", " tám ", " chín ");
    this.Tien = new Array("", " nghìn", " triệu", " tỷ", " nghìn tỷ", " triệu tỷ");
};

DocTienBangChu.prototype.docSo3ChuSo = function (baso) {
    var tram;
    var chuc;
    var donvi;
    var KetQua = "";
    tram = parseInt(baso / 100);
    chuc = parseInt((baso % 100) / 10);
    donvi = baso % 10;
    if (tram == 0 && chuc == 0 && donvi == 0) return "";
    if (tram != 0) {
        KetQua += this.ChuSo[tram] + " trăm ";
        if ((chuc == 0) && (donvi != 0)) KetQua += " linh ";
    }
    if ((chuc != 0) && (chuc != 1)) {
        KetQua += this.ChuSo[chuc] + " mươi";
        if ((chuc == 0) && (donvi != 0)) KetQua = KetQua + " linh ";
    }
    if (chuc == 1) KetQua += " mười ";
    switch (donvi) {
        case 1:
            if ((chuc != 0) && (chuc != 1)) {
                KetQua += " mốt ";
            }
            else {
                KetQua += this.ChuSo[donvi];
            }
            break;
        case 5:
            if (chuc == 0) {
                KetQua += this.ChuSo[donvi];
            }
            else {
                KetQua += " lăm ";
            }
            break;
        default:
            if (donvi != 0) {
                KetQua += this.ChuSo[donvi];
            }
            break;
    }
    return KetQua;
}

DocTienBangChu.prototype.doc = function (SoTien) {
    var lan = 0;
    var i = 0;
    var so = 0;
    var KetQua = "";
    var tmp = "";
    var soAm = false;
    var ViTri = new Array();
    if (SoTien < 0) soAm = true;//return "Số tiền âm !";
    if (SoTien == 0) return "Không đồng";//"Không đồng !";
    if (SoTien > 0) {
        so = SoTien;
    }
    else {
        so = -SoTien;
    }
    if (SoTien > 8999999999999999) {
        //SoTien = 0;
        return "";//"Số quá lớn!";
    }
    ViTri[5] = Math.floor(so / 1000000000000000);
    if (isNaN(ViTri[5]))
        ViTri[5] = "0";
    so = so - parseFloat(ViTri[5].toString()) * 1000000000000000;
    ViTri[4] = Math.floor(so / 1000000000000);
    if (isNaN(ViTri[4]))
        ViTri[4] = "0";
    so = so - parseFloat(ViTri[4].toString()) * 1000000000000;
    ViTri[3] = Math.floor(so / 1000000000);
    if (isNaN(ViTri[3]))
        ViTri[3] = "0";
    so = so - parseFloat(ViTri[3].toString()) * 1000000000;
    ViTri[2] = parseInt(so / 1000000);
    if (isNaN(ViTri[2]))
        ViTri[2] = "0";
    ViTri[1] = parseInt((so % 1000000) / 1000);
    if (isNaN(ViTri[1]))
        ViTri[1] = "0";
    ViTri[0] = parseInt(so % 1000);
    if (isNaN(ViTri[0]))
        ViTri[0] = "0";
    if (ViTri[5] > 0) {
        lan = 5;
    }
    else if (ViTri[4] > 0) {
        lan = 4;
    }
    else if (ViTri[3] > 0) {
        lan = 3;
    }
    else if (ViTri[2] > 0) {
        lan = 2;
    }
    else if (ViTri[1] > 0) {
        lan = 1;
    }
    else {
        lan = 0;
    }
    for (i = lan; i >= 0; i--) {
        tmp = this.docSo3ChuSo(ViTri[i]);
        KetQua += tmp;
        if (ViTri[i] > 0) KetQua += this.Tien[i];
        if ((i > 0) && (tmp.length > 0)) KetQua += '';//',';//&& (!string.IsNullOrEmpty(tmp))
    }
    if (KetQua.substring(KetQua.length - 1) == ',') {
        KetQua = KetQua.substring(0, KetQua.length - 1);
    }
    KetQua = KetQua.substring(1, 2).toUpperCase() + KetQua.substring(2);
    if (soAm) {
        return "Âm " + KetQua + " đồng";//.substring(0, 1);//.toUpperCase();// + KetQua.substring(1);
    }
    else {
        return KetQua + " đồng";//.substring(0, 1);//.toUpperCase();// + KetQua.substring(1);
    }
}

function formatNumber(numberString) {
    if (!numberString) return '';
    // Loại bỏ tất cả dấu chấm
    const num = numberString.replace(/\./g, '');
    const formatted = parseFloat(num).toString();
    return formatted.replace('.', ',');
}

function formatWithCommas(numberString) {
    if (!numberString) return '';
    const num = numberString.replace(',', '.');
    return parseFloat(num).toLocaleString('it-IT');
}

const SPREADSHEET_ID = '1kb0cieDcUElLmMaEcsNPptIXJs0ZNAu_aaCcmU0aDAU';
const RANGE = 'xuat_chuyen_kho!A:L'; // Mở rộng phạm vi đến cột L
const RANGE_CHITIET = 'xuat_chuyen_kho_chi_tiet!C:Q'; // Dải dữ liệu từ sheet 'don_hang_chi_tiet'
const API_KEY = 'AIzaSyA9g2qFUolpsu3_HVHOebdZb0NXnQgXlFM';

// Lấy giá trị từ URI sau dấu "?" cho các tham số cụ thể
function getDataFromURI() {
    const url = window.location.href;

    // Sử dụng RegEx để trích xuất giá trị của ma_phieu_xuat, xuat_tai_kho, và nhap_tai_kho
    const maPhieuxuatURIMatch = url.match(/ma_phieu_xuat=([^?&]*)/);
    const xuatTaiKhoMatch = url.match(/xuat_tai_kho=([^?&]*)/);
    const nhapTaiKhoMatch = url.match(/nhap_tai_kho=([^?&]*)/);

    // Gán các giá trị vào các biến
    const maPhieuxuatURI = maPhieuxuatURIMatch ? decodeURIComponent(maPhieuxuatURIMatch[1]) : null;
    const xuatTaiKhoURI = xuatTaiKhoMatch ? decodeURIComponent(xuatTaiKhoMatch[1]) : null;
    const nhapTaiKhoURI = nhapTaiKhoMatch ? decodeURIComponent(nhapTaiKhoMatch[1]) : null;

    // Trả về một đối tượng chứa các giá trị
    return {
        maPhieuxuatURI,
        xuatTaiKhoURI,
        nhapTaiKhoURI
    };
}

function extractDay(dateString) {
    if (!dateString) return '';

    // Xử lý cho định dạng dd/mm/yyyy hoặc d/m/yyyy
    const parts = dateString.split('/');
    if (parts.length !== 3) {
        // Thử tách bằng dấu gạch ngang
        const dashParts = dateString.split('-');
        if (dashParts.length === 3) {
            // Nếu định dạng là yyyy-mm-dd
            if (dashParts[0].length === 4) {
                // Định dạng yyyy-mm-dd
                return dashParts[2].padStart(2, '0'); // Lấy ngày, thêm số 0 nếu cần
            } else if (dashParts[2].length === 4) {
                // Định dạng dd-mm-yyyy
                return dashParts[0].padStart(2, '0'); // Lấy ngày
            }
        }
        return '';
    }

    // Phân tích định dạng dd/mm/yyyy
    let day, month, year;

    // Xác định phần nào là ngày, tháng, năm
    if (parts[2].length === 4) {
        // Định dạng dd/mm/yyyy hoặc d/m/yyyy
        day = parts[0];
        month = parts[1];
        year = parts[2];
    } else if (parts[0].length === 4) {
        // Định dạng yyyy/mm/dd
        year = parts[0];
        month = parts[1];
        day = parts[2];
    } else {
        // Không xác định, mặc định là dd/mm/yyyy
        day = parts[0];
        month = parts[1];
        year = parts[2];
    }

    // Trả về ngày đã được định dạng (2 chữ số)
    return parseInt(day).toString().padStart(2, '0');
}

// Thêm hàm để lấy tháng từ ngày
function extractMonth(dateString) {
    if (!dateString) return '';

    const parts = dateString.split('/');
    if (parts.length !== 3) {
        const dashParts = dateString.split('-');
        if (dashParts.length === 3) {
            if (dashParts[0].length === 4) {
                return dashParts[1].padStart(2, '0');
            } else if (dashParts[2].length === 4) {
                return dashParts[1].padStart(2, '0');
            }
        }
        return '';
    }

    let month;
    if (parts[2].length === 4) {
        month = parts[1];
    } else if (parts[0].length === 4) {
        month = parts[1];
    } else {
        month = parts[1];
    }

    return parseInt(month).toString().padStart(2, '0');
}

// Thêm hàm để lấy năm từ ngày
function extractYear(dateString) {
    if (!dateString) return '';

    const parts = dateString.split('/');
    if (parts.length !== 3) {
        const dashParts = dateString.split('-');
        if (dashParts.length === 3) {
            if (dashParts[0].length === 4) {
                return dashParts[0];
            } else if (dashParts[2].length === 4) {
                return dashParts[2];
            }
        }
        return '';
    }

    let year;
    if (parts[2].length === 4) {
        year = parts[2];
    } else if (parts[0].length === 4) {
        year = parts[0];
    } else {
        year = parts[2];
    }

    return year;
}


// Hàm để tải Google API Client
function loadGapiAndInitialize() {
    const script = document.createElement('script');
    script.src = "https://apis.google.com/js/api.js"; // Đường dẫn đến Google API Client
    script.onload = initialize; // Gọi hàm `initialize` sau khi thư viện được tải xong
    script.onerror = () => console.error('Failed to load Google API Client.');
    document.body.appendChild(script); // Gắn thẻ script vào tài liệu
}

// Hàm khởi tạo sau khi Google API Client được tải
function initialize() {
    gapi.load('client', async () => {
        try {
            await gapi.client.init({
                apiKey: API_KEY,
                discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4']
            });

            const uriData = getDataFromURI();
            if (!uriData.maPhieuxuatURI) {
                updateContent('No valid data found in URI.');
                return;
            }

            findRowInSheet(uriData.maPhieuxuatURI);
            findDetailsInSheet(uriData.maPhieuxuatURI);

        } catch (error) {
            updateContent('Initialization error: ' + error.message);
            console.error('Initialization Error:', error);
        }
    });
}

// Gọi hàm tải Google API Client khi DOM đã sẵn sàng
document.addEventListener('DOMContentLoaded', () => {
    loadGapiAndInitialize();
});

function updateContent(message) {
    const contentElement = document.getElementById('content'); // Thay 'content' bằng ID của phần tử HTML cần hiển thị
    if (contentElement) {
        contentElement.textContent = message;
    } else {
        console.warn('Element with ID "content" not found.');
    }
}


// Tìm chỉ số dòng chứa dữ liệu khớp trong cột B và lấy các giá trị từ các cột khác
let orderDetails = null; // Thông tin đơn hàng chính
let orderItems = [];

async function findRowInSheet(maPhieuxuatURI) {
    const uriData = getDataFromURI();

    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: RANGE,
        });

        const rows = response.result.values;
        if (!rows || rows.length === 0) {
            updateContent('No data found.');
            return;
        }

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];

            const bColumnValue = row[0]; // Cột A
            if (bColumnValue === maPhieuxuatURI) {
                // Lấy ngày tháng năm từ cột G
                const ngayXuatFull = row[6] || '';

                // Lưu dữ liệu vào biến toàn cục
                orderDetails = {
                    maPhieuXuat: row[0] || '', // Cột A
                    xuongSanXuat: row[2] || '', // Cột C
                    ngayXuat: ngayXuatFull, // Toàn bộ ngày từ cột G
                    // Sử dụng hàm extract để lấy ngày, tháng, năm từ ngàyXuat
                    thangXuat: row[5] || '', // Cột F (có thể giữ nguyên hoặc dùng từ ngàyXuat)
                    namXuat: row[4] || '', // Cột E (có thể giữ nguyên hoặc dùng từ ngàyXuat)
                    xuatTaiKho: uriData.xuatTaiKhoURI || row[10] || '', // Cột K
                    nhapTaiKho: uriData.nhapTaiKhoURI || row[11] || '', // Cột L
                    ghiChu: row[12] || '', // Cột M
                };

                // Cập nhật nội dung HTML
                document.getElementById('maPhieuXuat').textContent = orderDetails.maPhieuXuat;
                document.getElementById('xuongSanXuat').textContent = orderDetails.xuongSanXuat;

                // Sử dụng hàm extract để lấy ngày, tháng, năm
                document.getElementById('ngayXuat').textContent = extractDay(orderDetails.ngayXuat);

                // Ưu tiên lấy tháng từ cột G, nếu không có thì dùng từ cột F
                const thangTuNgayXuat = extractMonth(orderDetails.ngayXuat);
                document.getElementById('thangXuat').textContent = thangTuNgayXuat || orderDetails.thangXuat;

                // Ưu tiên lấy năm từ cột G, nếu không có thì dùng từ cột E
                const namTuNgayXuat = extractYear(orderDetails.ngayXuat);
                document.getElementById('namXuat').textContent = namTuNgayXuat || orderDetails.namXuat;

                document.getElementById('xuatTaiKho').textContent = orderDetails.xuatTaiKho;
                document.getElementById('nhapTaiKho').textContent = orderDetails.nhapTaiKho;
                document.getElementById('ghiChu').textContent = orderDetails.ghiChu;

                return; // Dừng khi tìm thấy
            }
        }

        updateContent(`No matching data found for "${maPhieuxuatURI}".`);
    } catch (error) {
        updateContent('Error fetching data: ' + error.message);
        console.error('Fetch Error:', error);
    }
}

// Tìm chi tiết trong bảng tính
async function findDetailsInSheet(maPhieuxuatURI) {
    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: RANGE_CHITIET,
        });

        const rows = response.result.values;
        if (!rows || rows.length === 0) {
            updateContent('No detail data found.');
            return;
        }

        const filteredRows = rows.filter(row => row[0] === maPhieuxuatURI); // Lọc các dòng có giá trị cột F khớp với maPhieuxuatURI
        orderItems = filteredRows.map(extractDetailDataFromRow);
        if (filteredRows.length > 0) {
            displayDetailData(filteredRows);
        } else {
            updateContent('No matching data found.');
        }
    } catch (error) {
        console.error('Error fetching detail data:', error);
        updateContent('Error fetching detail data.');
    }
}

// Hàm chuyển đổi chuỗi số từ định dạng Google Sheets (dấu . phân cách nghìn, dấu , thập phân)
function parseNumberFromSheet(value) {
    if (!value) return '';

    // Nếu đã là số thì trả về luôn
    if (typeof value === 'number') return value;

    // Chuẩn hóa: loại bỏ tất cả dấu . (phân cách nghìn)
    let normalized = value.toString().replace(/\./g, '');

    // Thay dấu , thành . để JavaScript có thể parse thành số
    normalized = normalized.replace(',', '.');

    // Parse thành số
    const num = parseFloat(normalized);

    // Kiểm tra nếu là NaN thì trả về chuỗi rỗng
    return isNaN(num) ? '' : num;
}

// Hàm định dạng số hiển thị theo kiểu Việt Nam (không có dấu phân cách nghìn, dấu , thập phân)
function formatNumberForDisplay(number) {
    if (number === '' || number === null || number === undefined) return '';

    // Chuyển số thành chuỗi
    let str = number.toString();

    // Nếu có dấu . thập phân, thay bằng dấu ,
    if (str.includes('.')) {
        // Tách phần nguyên và phần thập phân
        const parts = str.split('.');

        // Giữ nguyên phần nguyên (không thêm dấu phân cách nghìn)
        let integerPart = parts[0];

        // Giữ nguyên phần thập phân
        let decimalPart = parts[1];

        // Ghép lại với dấu ,
        str = integerPart + ',' + decimalPart;
    }

    return str;
}

// Trích xuất dữ liệu từ hàng
function extractDetailDataFromRow(row) {
    // Parse số lượng từ Google Sheets
    const slXuatRaw = row[6];
    const slXuatParsed = parseNumberFromSheet(slXuatRaw);
    const slXuatFormatted = formatNumberForDisplay(slXuatParsed);

    const slXuatQuydoiRaw = row[8];
    const slXuatQuydoiParsed = parseNumberFromSheet(slXuatQuydoiRaw);
    const slXuatQuydoiFormatted = formatNumberForDisplay(slXuatQuydoiParsed);

    return {
        sttTrongdon: row[1],
        maVattu: row[2],
        tenVattu: row[3],
        dvt: row[7],
        slXuat: slXuatFormatted, // Đã định dạng đúng
        dvtQuydoi: row[9],
        slXuatQuydoi: slXuatQuydoiFormatted, // Đã định dạng đúng
        slXuatQuydoiParsed: slXuatQuydoiParsed, // Giữ lại số để tính tổng
        vitriKehang: row[10],
        ghiChuItem: row[12],
        huongdandGhinhan: row[13],
    };
}

function displayDetailData(filteredRows) {
    const tableBody = document.getElementById('itemTableBody');
    tableBody.innerHTML = ''; // Xóa dữ liệu cũ nếu có

    let totalSlXuatQuydoi = 0;

    filteredRows.forEach(row => {
        const item = extractDetailDataFromRow(row);

        // Cộng dồn số lượng (dùng giá trị đã parse)
        totalSlXuatQuydoi += parseFloat(item.slXuatQuydoiParsed) || 0;

        tableBody.innerHTML += `
        <tr class="bordered-table">
            <td class="borderedcol-1">${item.sttTrongdon || ''}</td>
            <td class="borderedcol-2">${item.maVattu || ''}</td>
            <td class="borderedcol-3">${item.tenVattu || ''}</td>
            <td class="borderedcol-4">${item.dvt || ''}</td>
            <td class="borderedcol-5">${item.slXuat || ''}</td>
            <td class="borderedcol-6">${item.dvtQuydoi || ''}</td>
            <td class="borderedcol-7">${item.slXuatQuydoi || ''}</td>
            <td class="borderedcol-8">${item.vitriKehang || ''}</td>
            <td class="borderedcol-9">${item.ghiChuItem || ''}</td>
            <td class="borderedcol-10">${item.huongdandGhinhan || ''}</td>
        </tr>
    `;
    });

    // Định dạng tổng cho hiển thị
    const totalFormatted = formatNumberForDisplay(totalSlXuatQuydoi);

    // Thêm dòng tổng vào cuối bảng
    tableBody.innerHTML += `
        <tr class="bordered-table font-bold bg-gray-200">
            <th class="borderedcol-1" colspan="4" style="text-align: right;">Tổng:</th>
            <th class="borderedcol-5"></th>
            <th class="borderedcol-6"></th>
            <th class="borderedcol-7">${totalFormatted}</th>
            <th class="borderedcol-8"></th>
            <th class="borderedcol-9"></th>
            <th class="borderedcol-10"></th>
        </tr>
        `;
}

// Hàm cập nhật nội dung DOM
function updateElement(elementId, value) {
    const element = document.getElementById(elementId);
    if (element) {
        element.innerText = value;
    }
}
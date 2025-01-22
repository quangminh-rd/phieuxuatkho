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

    // Chuẩn hóa định dạng ngày về "DD/MM/YYYY"
    const parts = dateString.split(/[-/]/); // Chấp nhận cả "-" và "/"
    if (parts.length === 3) {
        let day, month, year;

        if (parts[0].length === 4) {
            // Định dạng ban đầu là "YYYY/MM/DD" hoặc "YYYY-MM-DD"
            [year, month, day] = parts;
        } else if (parts[1].length === 4) {
            // Định dạng ban đầu là "DD/MM/YYYY" (đã đúng)
            [day, month, year] = parts;
        } else {
            // Giả định định dạng "MM/DD/YYYY"
            [month, day, year] = parts;
        }

        // Đảm bảo các phần đều đủ 2 chữ số (nếu cần)
        day = day.padStart(2, '0');
        month = month.padStart(2, '0');

        // Chuẩn hóa thành "DD/MM/YYYY"
        dateString = `${day}/${month}/${year}`;
    }

    // Trích xuất ngày từ định dạng "DD/MM/YYYY"
    return dateString.split('/')[0];
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
                // Lưu dữ liệu vào biến toàn cục
                orderDetails = {
                    maPhieuXuat: row[0] || '', // Cột A
                    xuongSanXuat: row[2] || '', // Cột C
                    ngayXuat: row[6] || '', // Cột G
                    thangXuat: row[5] || '', // Cột F
                    namXuat: row[4] || '', // Cột E
                    xuatTaiKho: uriData.xuatTaiKhoURI || row[10] || '', // Cột K
                    nhapTaiKho: uriData.nhapTaiKhoURI || row[11] || '', // Cột L
                    ghiChu: row[12] || '', // Cột M
                };

                // Cập nhật nội dung HTML
                document.getElementById('maPhieuXuat').textContent = orderDetails.maPhieuXuat;
                document.getElementById('xuongSanXuat').textContent = orderDetails.xuongSanXuat;
                document.getElementById('ngayXuat').textContent = extractDay(orderDetails.ngayXuat);
                document.getElementById('thangXuat').textContent = orderDetails.thangXuat;
                document.getElementById('namXuat').textContent = orderDetails.namXuat;
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

    function updateContent(message) {
        // Hàm để xử lý thông báo lỗi hoặc cập nhật chung
        alert(message);
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

function displayDetailData(filteredRows) {
    const tableBody = document.getElementById('itemTableBody');
    tableBody.innerHTML = ''; // Xóa dữ liệu cũ nếu có

    filteredRows.forEach(row => {
        const item = extractDetailDataFromRow(row);;

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
}


// Trích xuất dữ liệu từ hàng
function extractDetailDataFromRow(row) {
    return {
        sttTrongdon: row[1],
        maVattu: row[2],
        tenVattu: row[3],
        dvt: row[7],
        slXuat: row[6],
        dvtQuydoi: row[9],
        slXuatQuydoi: row[8],
        vitriKehang: row[10],
        ghiChuItem: row[12],
        huongdandGhinhan: row[13],
    };
}

// Hàm cập nhật nội dung DOM
function updateElement(elementId, value) {
    const element = document.getElementById(elementId);
    if (element) {
        element.innerText = value;
    }
}
/**
 * Client-side script for Cao Dai Funeral Calendar Calculator
 * Manages UI interactions, live updates, and DOCX export requests
 */

const BTRE_SAU_SAT_NHAP = [
  "Xã Phú Túc", "Xã Giao Long", "Xã Tiên Thủy", "Xã Tân Phú", "Xã Phú Phụng",
  "Xã Chợ Lách", "Xã Vĩnh Thành", "Xã Hưng Khánh Trung", "Xã Phước Mỹ Trung",
  "Xã Tân Thành Bình", "Xã Nhuận Phú Tân", "Xã Đồng Khởi", "Xã Mỏ Cày",
  "Xã Thành Thới", "Xã An Định", "Xã Hương Mỹ", "Xã Đại Điền", "Xã Quới Điền",
  "Xã Thạnh Phú", "Xã An Qui", "Xã Thạnh Hải", "Xã Thạnh Phong", "Xã Tân Thủy",
  "Xã Bảo Thạnh", "Xã Ba Tri", "Xã Tân Xuân", "Xã Mỹ Chánh Hòa", "Xã An Ngãi Trung",
  "Xã An Hiệp", "Xã Hưng Nhượng", "Xã Giồng Trôm", "Xã Tân Hào", "Xã Phước Long",
  "Xã Lương Phú", "Xã Châu Hoà", "Xã Lương Hoà", "Xã Thới Thuận", "Xã Thạnh Phước",
  "Xã Bình Đại", "Xã Thạnh Trị", "Xã Lộc Thuận", "Xã Châu Hưng", "Xã Phú Thuận",
  "Phường An Hội", "Phường Phú Khương", "Phường Bến Tre", "Phường Sơn Đông", "Phường Phú Tân"
];

const BTRE_DATA = {
  "Thành phố Bến Tre": [
    "Phường Phú Khương",
    "Phường Phú Tân",
    "Phường 8",
    "Phường 6",
    "Phường An Hội",
    "Phường 7",
    "Xã Sơn Đông",
    "Xã Phú Hưng",
    "Xã Bình Phú",
    "Xã Mỹ Thạnh An",
    "Xã Nhơn Thạnh",
    "Xã Phú Nhuận"
  ],
  "Huyện Châu Thành": [
    "Xã Tân Thạch",
    "Xã Qưới Sơn",
    "Thị trấn Châu Thành",
    "Xã Giao Long",
    "Xã Phú Túc",
    "Xã Phú Đức",
    "Xã An Phước",
    "Xã Tam Phước",
    "Xã Thành Triệu",
    "Xã Tân Phú",
    "Xã Quới Thành",
    "Xã Phước Thạnh",
    "Xã Tiên Long",
    "Xã Tường Đa",
    "Xã Hữu Định",
    "Thị trấn Tiên Thủy"
  ],
  "Huyện Chợ Lách": [
    "Thị trấn Chợ Lách",
    "Xã Phú Phụng",
    "Xã Sơn Định",
    "Xã Vĩnh Bình",
    "Xã Hòa Nghĩa",
    "Xã Long Thới",
    "Xã Phú Sơn",
    "Xã Tân Thiềng",
    "Xã Vĩnh Thành",
    "Xã Vĩnh Hòa",
    "Xã Hưng Khánh Trung B"
  ],
  "Huyện Mỏ Cày Nam": [
    "Thị trấn Mỏ Cày",
    "Xã Định Thủy",
    "Xã Đa Phước Hội",
    "Xã Tân Hội",
    "Xã Phước Hiệp",
    "Xã Bình Khánh",
    "Xã An Thạnh",
    "Xã An Định",
    "Xã Thành Thới B",
    "Xã Tân Trung",
    "Xã An Thới",
    "Xã Thành Thới A",
    "Xã Minh Đức",
    "Xã Ngãi Đăng",
    "Xã Cẩm Sơn",
    "Xã Hương Mỹ"
  ],
  "Huyện Giồng Trôm": [
    "Thị trấn Giồng Trôm",
    "Xã Phong Nẫm",
    "Xã Mỹ Thạnh",
    "Xã Châu Hòa",
    "Xã Lương Hòa",
    "Xã Lương Quới",
    "Xã Lương Phú",
    "Xã Châu Bình",
    "Xã Thuận Điền",
    "Xã Sơn Phú",
    "Xã Bình Hoà",
    "Xã Phước Long",
    "Xã Hưng Phong",
    "Xã Long Mỹ",
    "Xã Tân Hào",
    "Xã Bình Thành",
    "Xã Tân Thanh",
    "Xã Tân Lợi Thạnh",
    "Xã Thạnh Phú Đông",
    "Xã Hưng Nhượng",
    "Xã Hưng Lễ"
  ],
  "Huyện Bình Đại": [
    "Thị trấn Bình Đại",
    "Xã Tam Hiệp",
    "Xã Long Định",
    "Xã Long Hòa",
    "Xã Phú Thuận",
    "Xã Vang Quới Tây",
    "Xã Vang Quới Đông",
    "Xã Châu Hưng",
    "Xã Lộc Thuận",
    "Xã Định Trung",
    "Xã Thới Lai",
    "Xã Bình Thới",
    "Xã Phú Long",
    "Xã Bình Thắng",
    "Xã Thạnh Trị",
    "Xã Đại Hòa Lộc",
    "Xã Thừa Đức",
    "Xã Thạnh Phước",
    "Xã Thới Thuận"
  ],
  "Huyện Ba Tri": [
    "Thị trấn Ba Tri",
    "Xã Mỹ Hòa",
    "Xã Tân Xuân",
    "Xã Mỹ Chánh",
    "Xã Bảo Thạnh",
    "Xã An Phú Trung",
    "Xã Mỹ Thạnh",
    "Xã Mỹ Nhơn",
    "Xã Phước Ngãi",
    "Xã An Ngãi Trung",
    "Xã Phú Lễ",
    "Xã An Bình Tây",
    "Xã Bảo Thuận",
    "Xã Tân Hưng",
    "Xã An Ngãi Tây",
    "Xã An Hiệp",
    "Xã Vĩnh Hòa",
    "Xã Tân Thủy",
    "Xã Vĩnh An",
    "Xã An Đức",
    "Xã An Hòa Tây",
    "Thị trấn Tiệm Tôm"
  ],
  "Huyện Thạnh Phú": [
    "Thị trấn Thạnh Phú",
    "Xã Phú Khánh",
    "Xã Đại Điền",
    "Xã Quới Điền",
    "Xã Tân Phong",
    "Xã Mỹ Hưng",
    "Xã An Thạnh",
    "Xã Thới Thạnh",
    "Xã Hòa Lợi",
    "Xã An Điền",
    "Xã Bình Thạnh",
    "Xã An Thuận",
    "Xã An Quy",
    "Xã Thạnh Hải",
    "Xã An Nhơn",
    "Xã Giao Thạnh",
    "Xã Thạnh Phong",
    "Xã Mỹ An"
  ],
  "Huyện Mỏ Cày Bắc": [
    "Xã Phú Mỹ",
    "Xã Hưng Khánh Trung A",
    "Xã Thanh Tân",
    "Xã Thạnh Ngãi",
    "Xã Tân Phú Tây",
    "Thị trấn Phước Mỹ Trung",
    "Xã Tân Thành Bình",
    "Xã Thành An",
    "Xã Hòa Lộc",
    "Xã Tân Thanh Tây",
    "Xã Tân Bình",
    "Xã Nhuận Phú Tân",
    "Xã Khánh Thạnh Tân"
  ]
};

document.addEventListener("DOMContentLoaded", () => {
    // UI Elements
    const hoTenInput = document.getElementById("hoTen");
    const gioiTinhSelect = document.getElementById("gioiTinh");
    const phamDaoSelect = document.getElementById("phamDao");
    const gioMatSelect = document.getElementById("gioMat");
    const ngayAlSelect = document.getElementById("ngayMatAl");
    const thangAlSelect = document.getElementById("thangMatAl");
    const namAlSelect = document.getElementById("namMatAl");
    const nhuanCheckbox = document.getElementById("thangMatAlNhuan");
    const sinhNamSelect = document.getElementById("sinhNam");
    const tieuDeXungSelect = document.getElementById("tieuDeXung");
    const ageDisplayBox = document.getElementById("age-display");

    // Nơi sinh elements
    const noiSinhHuyenSelect = document.getElementById("noiSinhHuyen");
    const noiSinhXaSelect = document.getElementById("noiSinhXa");

    // Sớ Cầu Siêu elements
    const soTuanCuuSelect = document.getElementById("soTuanCuu");
    const noiCungSelect = document.getElementById("noiCung");
    const banCungSelect = document.getElementById("banCung");
    const soCungXaSelect = document.getElementById("soCungXa");
    const btnExportCauSieu = document.getElementById("btn-export-causieu");
    
    // Preview Labels
    const lblSolarEquivalent = document.getElementById("solar-equivalent-date");
    const lblPreviewBadge = document.getElementById("preview-title-badge");
    const lblNgayAl = document.getElementById("lblNgayAl");
    const lblThangAl = document.getElementById("lblThangAl");
    const lblNamAl = document.getElementById("lblNamAl");
    const lblCanChiAl = document.getElementById("lblCanChiAl");
    const lblGioMat = document.getElementById("lblGioMat");
    const lblNgayDl = document.getElementById("lblNgayDl");
    const lblThangDl = document.getElementById("lblThangDl");
    const lblNamDl = document.getElementById("lblNamDl");
    
    const previewTableBody = document.getElementById("preview-table-body");
    const btnExport = document.getElementById("btn-export");
    const loadingOverlay = document.getElementById("loading-overlay");

    // Global Sino-Vietnamese Age Translation
    function toSinoVietnamese(num) {
        const units = {
            1: "Nhứt", 2: "Nhị", 3: "Tam", 4: "Tứ", 5: "Ngũ",
            6: "Lục", 7: "Thất", 8: "Bát", 9: "Cửu"
        };
        if (num < 10) return units[num] || "";
        if (num === 10) return "Thập";
        if (num < 20) return "Thập " + units[num - 10];
        if (num < 100) {
            const ten = Math.floor(num / 10);
            const unit = num % 10;
            const tenStr = units[ten] + " Thập";
            return unit === 0 ? tenStr : tenStr + " " + units[unit];
        }
        if (num === 100) return "Bá";
        if (num < 110) return "Bá " + units[num % 10];
        return num.toString();
    }

    // Initialize Dropdowns
    // Days 1 to 30
    for (let d = 1; d <= 30; d++) {
        const opt = document.createElement("option");
        opt.value = d;
        opt.textContent = String(d).padStart(2, '0');
        if (d === 15) opt.selected = true; // Default 15th
        ngayAlSelect.appendChild(opt);
    }

    // Months 1 to 12
    for (let m = 1; m <= 12; m++) {
        const opt = document.createElement("option");
        opt.value = m;
        opt.textContent = String(m).padStart(2, '0');
        if (m === 7) opt.selected = true; // Default 7th month
        thangAlSelect.appendChild(opt);
    }

    // Death Years (from 1980 to 2080, default to 2026)
    const currentYear = 2026;
    for (let y = 1980; y <= 2080; y++) {
        const opt = document.createElement("option");
        opt.value = y;
        opt.textContent = `${y} (${CaoDaiCalendar.getCanChiYear(y)})`;
        if (y === currentYear) opt.selected = true;
        namAlSelect.appendChild(opt);
    }

    // Birth Years (from 1900 to 2080, default to 1945)
    for (let y = 1900; y <= 2080; y++) {
        const opt = document.createElement("option");
        opt.value = y;
        opt.textContent = `${y} (${CaoDaiCalendar.getCanChiYear(y)})`;
        if (y === 1945) opt.selected = true;
        sinhNamSelect.appendChild(opt);
    }

    // Populate Huyện nơi sinh
    for (const huyen in BTRE_DATA) {
        const opt = document.createElement("option");
        opt.value = huyen;
        opt.textContent = huyen;
        if (huyen === "Huyện Giồng Trôm") {
            opt.selected = true;
        }
        noiSinhHuyenSelect.appendChild(opt);
    }

    // Populate Xã nơi sinh based on selected Huyện
    function updateNoiSinhXa(defaultXa = null) {
        const huyen = noiSinhHuyenSelect.value;
        const xas = BTRE_DATA[huyen] || [];
        noiSinhXaSelect.innerHTML = "";
        xas.forEach(xa => {
            const opt = document.createElement("option");
            opt.value = xa;
            opt.textContent = xa;
            if (defaultXa && xa === defaultXa) {
                opt.selected = true;
            } else if (!defaultXa && xa === "Xã Mỹ Thạnh") {
                opt.selected = true;
            }
            noiSinhXaSelect.appendChild(opt);
        });
    }

    // Initial population for Xã nơi sinh
    updateNoiSinhXa("Xã Mỹ Thạnh");

    // Populate SoCungXa from BTRE_SAU_SAT_NHAP
    BTRE_SAU_SAT_NHAP.forEach(xa => {
        const opt = document.createElement("option");
        opt.value = xa;
        opt.textContent = xa;
        if (xa === "Xã Lương Phú") {
            opt.selected = true;
        }
        soCungXaSelect.appendChild(opt);
    });

    // Recalculate function
    let currentCalculation = null;
    let currentAge = 0;
    let currentBirthCanChi = "";
    let currentAgeTerm = "Hưởng thọ";
    let lastCalculatedAge = null;

    function runRecalculation() {
        const name = hoTenInput.value.trim() || "Chưa nhập họ tên";
        const title = phamDaoSelect.value;
        const gioMat = gioMatSelect.value;
        const lDay = parseInt(ngayAlSelect.value);
        const lMonth = parseInt(thangAlSelect.value);
        const lYear = parseInt(namAlSelect.value);
        const lLeap = nhuanCheckbox.checked ? 1 : 0;
        const birthYear = parseInt(sinhNamSelect.value);

        // Recalculate Calendar dates
        const results = CaoDaiCalendar.calculateCaoDaiDates(lDay, lMonth, lYear, lLeap);
        currentCalculation = results;

        // Calculate Age (Newborn is 1 year old)
        currentAge = lYear - birthYear + 1;
        currentBirthCanChi = CaoDaiCalendar.getCanChiYear(birthYear);

        // Auto-detect term based on age if the age changes
        if (currentAge !== lastCalculatedAge && currentAge >= 1) {
            let calculatedTerm = "Hưởng thọ";
            if (currentAge < 61) {
                calculatedTerm = currentAge < 18 ? "Hưởng linh" : "Hưởng dương";
            } else if (currentAge >= 61 && currentAge <= 70) {
                calculatedTerm = "Hưởng thọ";
            } else if (currentAge >= 71 && currentAge <= 80) {
                calculatedTerm = "Hưởng trung thọ";
            } else if (currentAge >= 81 && currentAge <= 90) {
                calculatedTerm = "Hưởng thượng thọ";
            } else if (currentAge >= 91) {
                calculatedTerm = "Hưởng mạo";
            }
            tieuDeXungSelect.value = calculatedTerm;
            lastCalculatedAge = currentAge;
        }

        // Read value from selector (allows manual override)
        currentAgeTerm = tieuDeXungSelect.value;

        // Update Age Display Box
        if (currentAge >= 1) {
            ageDisplayBox.textContent = `${currentBirthCanChi} — ${currentAge} tuổi`;
        } else {
            ageDisplayBox.textContent = `${currentBirthCanChi} — Năm sinh không hợp lệ`;
        }

        // Update Preview Labels
        lblSolarEquivalent.textContent = `${results.deathSolarStr} (${results.deathSolarDayOfWeek})`;
        
        let previewTitle = `${title} : ${name}`;
        if (currentAge >= 1) {
            previewTitle += ` (Sinh năm ${birthYear}-${currentBirthCanChi}, ${currentAgeTerm} ${currentAge} tuổi)`;
        }
        lblPreviewBadge.textContent = previewTitle;
        
        lblNgayAl.textContent = String(lDay).padStart(2, '0');
        lblThangAl.textContent = String(lMonth).padStart(2, '0') + (lLeap ? " (Nhuận)" : "");
        lblNamAl.textContent = lYear;
        lblCanChiAl.textContent = results.lunarYearCanChi;
        lblGioMat.textContent = gioMat;

        // Solar death details
        const [sDay, sMonth, sYear] = results.deathSolar;
        lblNgayDl.textContent = String(sDay).padStart(2, '0');
        lblThangDl.textContent = String(sMonth).padStart(2, '0');
        lblNamDl.textContent = sYear;

        // Populate Table View
        previewTableBody.innerHTML = "";
        results.table.forEach(item => {
            const tr = document.createElement("tr");
            
            const tdEvent = document.createElement("td");
            tdEvent.textContent = item.event;
            tr.appendChild(tdEvent);

            const tdDayOfWeek = document.createElement("td");
            tdDayOfWeek.textContent = item.dayOfWeek;
            tr.appendChild(tdDayOfWeek);

            const tdAl = document.createElement("td");
            tdAl.textContent = item.dateAl;
            tr.appendChild(tdAl);

            const tdDl = document.createElement("td");
            tdDl.textContent = item.dateDl;
            tr.appendChild(tdDl);

            previewTableBody.appendChild(tr);
        });
        
        // Update active preview tab reactively
        if (typeof updateActiveTabPreview === "function") {
            updateActiveTabPreview();
        }
    }

    // Set up Event Listeners
    hoTenInput.addEventListener("input", runRecalculation);
    gioiTinhSelect.addEventListener("change", runRecalculation);
    phamDaoSelect.addEventListener("change", runRecalculation);
    gioMatSelect.addEventListener("change", runRecalculation);
    ngayAlSelect.addEventListener("change", runRecalculation);
    thangAlSelect.addEventListener("change", runRecalculation);
    namAlSelect.addEventListener("change", runRecalculation);
    nhuanCheckbox.addEventListener("change", runRecalculation);
    sinhNamSelect.addEventListener("change", runRecalculation);
    tieuDeXungSelect.addEventListener("change", runRecalculation);

    noiSinhHuyenSelect.addEventListener("change", () => {
        updateNoiSinhXa();
        runRecalculation();
    });
    noiSinhXaSelect.addEventListener("change", runRecalculation);
    soCungXaSelect.addEventListener("change", runRecalculation);

    // Link Nơi cúng cửu and Bàn cúng
    noiCungSelect.addEventListener("change", () => {
        if (noiCungSelect.value === "Thánh Thất") {
            banCungSelect.value = "Điện Tiền";
        } else if (noiCungSelect.value === "Gia Đường") {
            banCungSelect.value = "Thiên Bàn";
        }
    });

    banCungSelect.addEventListener("change", () => {
        if (banCungSelect.value === "Điện Tiền") {
            noiCungSelect.value = "Thánh Thất";
        } else if (banCungSelect.value === "Thiên Bàn") {
            noiCungSelect.value = "Gia Đường";
        }
    });

    // Form submission helper for dynamic file downloads (fully compatible with mobile WebViews/Zalo/FB)
    function submitFormDownload(actionUrl, postData) {
        const form = document.createElement("form");
        form.method = "POST";
        form.action = actionUrl;
        form.style.display = "none";

        for (const key in postData) {
            if (postData.hasOwnProperty(key)) {
                const input = document.createElement("input");
                input.type = "hidden";
                input.name = key;
                if (typeof postData[key] === "object") {
                    input.value = JSON.stringify(postData[key]);
                } else {
                    input.value = postData[key];
                }
                form.appendChild(input);
            }
        }

        document.body.appendChild(form);
        form.submit();
        
        setTimeout(() => {
            document.body.removeChild(form);
        }, 1000);
    }

    // Export DOCX handler
    btnExport.addEventListener("click", () => {
        if (!hoTenInput.value.trim()) {
            alert("Vui lòng nhập họ tên người mất!");
            hoTenInput.focus();
            return;
        }

        if (!currentCalculation) {
            alert("Không có dữ liệu tính toán hợp lệ!");
            return;
        }

        loadingOverlay.classList.remove("hidden");

        const postData = {
            hoTen: hoTenInput.value.trim(),
            phamDao: phamDaoSelect.value,
            gioMat: gioMatSelect.value,
            ngayMatAl: `${String(ngayAlSelect.value).padStart(2, '0')}/${String(thangAlSelect.value).padStart(2, '0')}/${namAlSelect.value} (ÂL)${nhuanCheckbox.checked ? ' (Nhuận)' : ''}`,
            lunarYearCanChi: currentCalculation.lunarYearCanChi,
            ngayMatDl: `${currentCalculation.deathSolarStr} (DL)`,
            tuoiMat: currentAge >= 1 ? currentAge : 0,
            namSinhAl: `${parseInt(sinhNamSelect.value)}-${currentBirthCanChi}`,
            tuoiTerm: currentAgeTerm,
            tableData: currentCalculation.table
        };

        submitFormDownload("/generate", postData);

        setTimeout(() => {
            loadingOverlay.classList.add("hidden");
        }, 1500);
    });

    // Export Sớ Cầu Siêu handler
    btnExportCauSieu.addEventListener("click", () => {
        if (!currentCalculation) {
            alert("Vui lòng thực hiện tính toán lịch trước!");
            return;
        }

        const selectedEvent = soTuanCuuSelect.value;
        const noiCung = noiCungSelect.value;
        const banCung = banCungSelect.value;

        // Find cúng date from calculated table
        const row = currentCalculation.table.find(item => item.event.startsWith(selectedEvent) || item.event.includes(selectedEvent));
        if (!row) {
            alert("Không tìm thấy ngày cúng cho tuần cửu đã chọn!");
            return;
        }

        // Parse dateAl string
        const parts = row.dateAl.split('/');
        if (parts.length < 3) {
            alert("Ngày âm lịch cúng không hợp lệ!");
            return;
        }

        const day = parseInt(parts[0]);
        const month = parseInt(parts[1]);
        const rest = parts[2];
        const leap = rest.includes('(Nhuận)');
        const year = rest.replace('(Nhuận)', '').trim();

        // Translate to Sino-Vietnamese
        const LUNAR_MONTH_SINO = {
            1: "Chánh ngoạt", 2: "Nhị ngoạt", 3: "Tam ngoạt", 4: "Tứ ngoạt", 5: "Ngũ ngoạt", 6: "Lục ngoạt",
            7: "Thất ngoạt", 8: "Bát ngoạt", 9: "Cửu ngoạt", 10: "Thập ngoạt", 11: "Thập Nhứt ngoạt", 12: "Thập Nhị ngoạt"
        };

        const LUNAR_DAY_SINO = {
            1: "sơ nhất nhựt", 2: "sơ nhị nhựt", 3: "sơ tam nhựt", 4: "sơ tứ nhựt", 5: "sơ ngũ nhựt",
            6: "sơ lục nhựt", 7: "sơ thất nhựt", 8: "sơ bát nhựt", 9: "sơ cửu nhựt", 10: "Thập nhựt",
            11: "Thập nhứt nhựt", 12: "Thập nhị nhựt", 13: "Thập tam nhựt", 14: "Thập tứ nhựt", 15: "Thập ngũ nhựt",
            16: "Thập lục nhựt", 17: "Thập thất nhựt", 18: "Thập bát nhựt", 19: "Thập cửu nhựt", 20: "Nhị thập nhựt",
            21: "Nhị thập nhứt nhựt", 22: "Nhị thập nhị nhựt", 23: "Nhị thập tam nhựt", 24: "Nhị thập  tứ nhựt", 25: "Nhị thập ngũ nhựt",
            26: "Nhị thập lục nhựt", 27: "Nhị thập thất nhựt", 28: "Nhị thập bát nhựt", 29: "Nhị thập cửu nhựt", 30: "Tam thập nhựt"
        };

        const daySino = LUNAR_DAY_SINO[day] || `${day} nhựt`;
        let monthSino = LUNAR_MONTH_SINO[month] || `${month} ngoạt`;
        if (leap) {
            monthSino = "Nhuận " + monthSino;
        }

        // Determine value for (6) and filename suffix
        let tuanCuuSino = "";
        let filenameSuffix = "";

        if (selectedEvent.includes("Nhứt cửu")) {
            tuanCuuSino = "Nhứt cửu";
            filenameSuffix = "nhứt cửu";
        } else if (selectedEvent.includes("Nhị cửu")) {
            tuanCuuSino = "Nhị cửu";
            filenameSuffix = "nhị cửu";
        } else if (selectedEvent.includes("Tam cửu")) {
            tuanCuuSino = "Tam cửu";
            filenameSuffix = "tam cửu";
        } else if (selectedEvent.includes("Tứ cửu")) {
            tuanCuuSino = "Tứ cửu";
            filenameSuffix = "tứ cửu";
        } else if (selectedEvent.includes("Ngũ cửu")) {
            tuanCuuSino = "Ngũ cửu";
            filenameSuffix = "ngũ cửu";
        } else if (selectedEvent.includes("Lục cửu")) {
            tuanCuuSino = "Lục cửu";
            filenameSuffix = "lục cửu";
        } else if (selectedEvent.includes("Thất cửu")) {
            tuanCuuSino = "Thất cửu";
            filenameSuffix = "thất cửu";
        } else if (selectedEvent.includes("Bát cửu")) {
            tuanCuuSino = "Bát cửu";
            filenameSuffix = "bát cửu";
        } else if (selectedEvent.includes("Cửu cửu")) {
            tuanCuuSino = "Chung cửu";
            filenameSuffix = "chung cửu";
        } else if (selectedEvent.includes("Tiểu tường")) {
            tuanCuuSino = "Tiểu tường";
            filenameSuffix = "Tiểu tường";
        } else if (selectedEvent.includes("Đại tường")) {
            tuanCuuSino = "Đại tường";
            filenameSuffix = "Đại tường";
        }

        const nameVal = hoTenInput.value.trim() || "Chưa nhập họ tên";
        const phamDao = phamDaoSelect.value;
        const gioiTinh = gioiTinhSelect.value;
        const nameParts = nameVal.split(/\s+/);
        const lastName = nameParts[nameParts.length - 1] || "";

        let val7 = nameVal;
        if (phamDao === "Đạo hữu") {
            val7 = nameVal;
        } else if (phamDao === "Lễ Sanh") {
            if (gioiTinh === "Nữ") {
                val7 = `Hương ${lastName}`;
            } else {
                val7 = `Lễ Sanh……….${lastName} Thanh`;
            }
        } else if (phamDao === "Chức việc") {
            val7 = ` ………………..${nameVal}`;
        }



        const val8 = nameVal;
        const val9 = toSinoVietnamese(currentAge);

        // Birthplace details
        const huyenSelected = noiSinhHuyenSelect.value;
        const xaSelected = noiSinhXaSelect.value;

        let huyenName = huyenSelected;
        let huyenType = "huyện";
        if (huyenSelected.startsWith("Huyện ")) {
            huyenName = huyenSelected.replace("Huyện ", "");
            huyenType = "huyện";
        } else if (huyenSelected.startsWith("Thành phố ")) {
            huyenName = huyenSelected.replace("Thành phố ", "");
            huyenType = "thành phố";
        }

        let xaName = xaSelected;
        let xaType = "xã";
        if (xaSelected.startsWith("Xã ")) {
            xaName = xaSelected.replace("Xã ", "");
            xaType = "xã";
        } else if (xaSelected.startsWith("Phường ")) {
            xaName = xaSelected.replace("Phường ", "");
            xaType = "phường";
        } else if (xaSelected.startsWith("Thị trấn ")) {
            xaName = xaSelected.replace("Thị trấn ", "");
            xaType = "thị trấn";
        }

        // Cúng location commune details after merger
        const cungXaSelected = soCungXaSelect.value;
        let cungXaName = cungXaSelected;
        let cungXaType = "xã";
        if (cungXaSelected.startsWith("Xã ")) {
            cungXaName = cungXaSelected.replace("Xã ", "");
            cungXaType = "xã";
        } else if (cungXaSelected.startsWith("Phường ")) {
            cungXaName = cungXaSelected.replace("Phường ", "");
            cungXaType = "phường";
        }

        // Time of death Sino calculations
        const lMonthVal = parseInt(thangAlSelect.value);
        const LUNAR_MONTH_NUM_SINO = {
            1: "Chánh", 2: "Nhị", 3: "Tam", 4: "Tứ", 5: "Ngũ", 6: "Lục",
            7: "Thất", 8: "Bát", 9: "Cửu", 10: "Thập", 11: "Thập Nhứt", 12: "Thập Nhị"
        };
        let val14 = LUNAR_MONTH_NUM_SINO[lMonthVal] || lMonthVal;
        if (nhuanCheckbox.checked) {
            val14 = "Nhuận " + val14;
        }

        const lDayVal = parseInt(ngayAlSelect.value);
        const daySinoFull = LUNAR_DAY_SINO[lDayVal] || `${lDayVal} nhựt`;
        const val15 = daySinoFull.replace(/\s*nhựt\s*$/, "");

        const val16 = gioMatSelect.value;
        const val13 = currentCalculation.lunarYearCanChi;

        const postData = {
            val1: year,
            val2: monthSino,
            val3: daySino,
            val4: noiCung,
            val5: banCung,
            val6: tuanCuuSino,
            val7: val7,
            val8: val8,
            val9: val9,
            val10: cungXaName,
            cungXaType: cungXaType,
            val11: huyenName,
            huyenType: huyenType,
            val12: xaName,
            xaType: xaType,
            val13: val13,
            val14: val14,
            val15: val15,
            val16: val16,
            filename: `Sớ cầu siêu ${filenameSuffix}`
        };

        loadingOverlay.classList.remove("hidden");

        submitFormDownload("/generate_causieu", postData);

        setTimeout(() => {
            loadingOverlay.classList.add("hidden");
        }, 1500);
    });

    // Export Linh Vị handler
    const btnExportLinhVi = document.getElementById("btn-export-linhvi");
    btnExportLinhVi.addEventListener("click", () => {
        if (!hoTenInput.value.trim()) {
            alert("Vui lòng nhập họ tên người mất!");
            hoTenInput.focus();
            return;
        }

        if (!currentCalculation) {
            alert("Vui lòng thực hiện tính toán lịch trước!");
            return;
        }

        const nameVal = hoTenInput.value.trim();
        const phamDao = phamDaoSelect.value;
        const gioiTinh = gioiTinhSelect.value;
        const nameParts = nameVal.split(/\s+/);
        const lastName = nameParts[nameParts.length - 1] || "";
        const firstName = nameParts[0] || "";

        // (1a), (1b) split birth year can chi
        const birthParts = currentBirthCanChi.split(/\s+/);
        const val1a = birthParts[0] || "";
        const val1b = birthParts[1] || "";

        // (2a), (2b) birth province (Bến Tre is fixed)
        const val2a = "Bến";
        const val2b = "Tre";

        // (3a), (3b) split birthplace district
        const huyenSelected = noiSinhHuyenSelect.value;
        let huyenName = huyenSelected.replace("Huyện ", "").replace("Thành phố ", "");
        const huyenParts = huyenName.split(/\s+/);
        let val3a = "";
        let val3b = "";
        if (huyenParts.length === 2) {
            val3a = huyenParts[0];
            val3b = huyenParts[1];
        } else if (huyenParts.length === 3) {
            val3a = huyenParts[0] + " " + huyenParts[1];
            val3b = huyenParts[2];
        } else {
            val3a = huyenParts[0] || "";
            val3b = "";
        }

        // (4a), (4b) split birthplace commune
        const xaSelected = noiSinhXaSelect.value;
        let xaName = xaSelected.replace("Xã ", "").replace("Phường ", "").replace("Thị trấn ", "");
        const xaParts = xaName.split(/\s+/);
        let val4a = "";
        let val4b = "";
        if (xaParts.length === 2) {
            val4a = xaParts[0];
            val4b = xaParts[1];
        } else if (xaParts.length === 3) {
            val4a = xaParts[0] + " " + xaParts[1];
            val4b = xaParts[2];
        } else if (xaParts.length === 4) {
            val4a = xaParts[0] + " " + xaParts[1];
            val4b = xaParts[2] + " " + xaParts[3];
        } else {
            val4a = xaParts[0] || "";
            val4b = "";
        }

        // (5a), (5b) split pham dao
        const phamParts = phamDao.split(/\s+/);
        const val5a = phamParts[0] || "";
        const val5b = phamParts[1] || "";

        // (6) Họ
        const val6 = firstName;

        // (7) Tên
        const val7 = lastName;

        // (8) Tuổi term (Dương, Linh, Thọ, Hạ Thọ, Trung Thọ, Thượng Thọ, Mạo)
        let val8 = "Thọ";
        const term = currentAgeTerm.toLowerCase();
        if (term.includes("dương")) val8 = "Dương";
        else if (term.includes("linh")) val8 = "Linh";
        else if (term.includes("trung thọ")) val8 = "Trung Thọ";
        else if (term.includes("thượng thọ")) val8 = "Thượng Thọ";
        else if (term.includes("hạ thọ")) val8 = "Hạ Thọ";
        else if (term.includes("mạo")) val8 = "Mạo";
        else if (term.includes("thọ")) val8 = "Thọ";

        // (9), wordThap, (10) split age Hán Việt
        const units = {
            1: "Nhứt", 2: "Nhị", 3: "Tam", 4: "Tứ", 5: "Ngũ",
            6: "Lục", 7: "Thất", 8: "Bát", 9: "Cửu"
        };
        let val9 = "";
        let val10 = "";
        let wordThap = "Thập";
        
        if (currentAge < 10) {
            val9 = "";
            wordThap = "";
            val10 = units[currentAge] || "";
        } else if (currentAge === 10) {
            val9 = "";
            wordThap = "Thập";
            val10 = "";
        } else if (currentAge < 20) {
            val9 = "";
            wordThap = "Thập";
            val10 = units[currentAge - 10] || "";
        } else {
            const chuc = Math.floor(currentAge / 10);
            const donVi = currentAge % 10;
            val9 = units[chuc] || "";
            wordThap = "Thập";
            val10 = donVi === 0 ? "" : (units[donVi] || "");
        }

        // (11a), (11b) split death year can chi
        const deathParts = currentCalculation.lunarYearCanChi.split(/\s+/);
        const val11a = deathParts[0] || "";
        const val11b = deathParts[1] || "";

        // (12) month of death Hán Việt
        const lMonthVal = parseInt(thangAlSelect.value);
        const LUNAR_MONTH_LINHVI = {
            1: "Chánh", 2: "Nhị", 3: "Tam", 4: "Tứ", 5: "Ngũ", 6: "Lục",
            7: "Thất", 8: "Bát", 9: "Cửu", 10: "Thập", 11: "Thập Nhứt", 12: "Nhứt Thập Nhị"
        };
        let val12 = LUNAR_MONTH_LINHVI[lMonthVal] || lMonthVal;
        if (nhuanCheckbox.checked) {
            val12 = "Nhuận " + val12;
        }

        // (13a), (13b) split day of death Hán Việt
        const lDayVal = parseInt(ngayAlSelect.value);
        const LUNAR_DAY_SINO = {
            1: "sơ nhất nhựt", 2: "sơ nhị nhựt", 3: "sơ tam nhựt", 4: "sơ tứ nhựt", 5: "sơ ngũ nhựt",
            6: "sơ lục nhựt", 7: "sơ thất nhựt", 8: "sơ bát nhựt", 9: "sơ cửu nhựt", 10: "Thập nhựt",
            11: "Thập nhứt nhựt", 12: "Thập nhị nhựt", 13: "Thập tam nhựt", 14: "Thập tứ nhựt", 15: "Thập ngũ nhựt",
            16: "Thập lục nhựt", 17: "Thập thất nhựt", 18: "Thập bát nhựt", 19: "Thập cửu nhựt", 20: "Nhị thập nhựt",
            21: "Nhị thập nhứt nhựt", 22: "Nhị thập nhị nhựt", 23: "Nhị thập tam nhựt", 24: "Nhị thập  tứ nhựt", 25: "Nhị thập ngũ nhựt",
            26: "Nhị thập lục nhựt", 27: "Nhị thập thất nhựt", 28: "Nhị thập bát nhựt", 29: "Nhị thập cửu nhựt", 30: "Tam thập nhựt"
        };
        const daySinoFull = LUNAR_DAY_SINO[lDayVal] || `${lDayVal} nhựt`;
        const dayText = daySinoFull.replace(/\s*nhựt\s*$/, "");
        const dayParts = dayText.split(/\s+/);
        let val13a = "";
        let val13b = "";
        if (dayParts.length === 1) {
            val13a = dayParts[0];
            val13b = "";
        } else if (dayParts.length === 2) {
            if (dayParts[1] === "thập") {
                val13a = dayParts[0] + " " + dayParts[1];
                val13b = "";
            } else {
                val13a = dayParts[0];
                val13b = dayParts[1];
            }
        } else if (dayParts.length === 3) {
            val13a = dayParts[0] + " " + dayParts[1];
            val13b = dayParts[2];
        }

        // (14) hour of death
        const val14 = gioMatSelect.value;

        // (15a), (15b) split cúng commune
        const cungXaSelected = soCungXaSelect.value;
        let cungXaName = cungXaSelected.replace("Xã ", "").replace("Phường ", "");
        const cungXaParts = cungXaName.split(/\s+/);
        let val15a = "";
        let val15b = "";
        if (cungXaParts.length === 2) {
            val15a = cungXaParts[0];
            val15b = cungXaParts[1];
        } else if (cungXaParts.length === 3) {
            val15a = cungXaParts[0] + " " + cungXaParts[1];
            val15b = cungXaParts[2];
        } else if (cungXaParts.length === 4) {
            val15a = cungXaParts[0] + " " + cungXaParts[1];
            val15b = cungXaParts[2] + " " + cungXaParts[3];
        } else {
            val15a = cungXaParts[0] || "";
            val15b = "";
        }

        const postData = {
            gioiTinh: gioiTinh,
            val1a: val1a,
            val1b: val1b,
            val2a: val2a,
            val2b: val2b,
            val3a: val3a,
            val3b: val3b,
            val4a: val4a,
            val4b: val4b,
            val5a: val5a,
            val5b: val5b,
            val6: val6,
            val7: val7,
            val8: val8,
            val9: val9,
            val10: val10,
            wordThap: wordThap,
            val11a: val11a,
            val11b: val11b,
            val12: val12,
            val13a: val13a,
            val13b: val13b,
            val14: val14,
            val15a: val15a,
            val15b: val15b
        };

        loadingOverlay.classList.remove("hidden");

        submitFormDownload("/generate_linhvi", postData);

        setTimeout(() => {
            loadingOverlay.classList.add("hidden");
        }, 1500);
    });

    // Render Sớ Cầu Siêu HTML Preview
    function renderCauSieuPreview() {
        if (!currentCalculation) return;

        const nameVal = hoTenInput.value.trim() || "CHƯA NHẬP HỌ TÊN";
        const phamDao = phamDaoSelect.value;
        const gioiTinh = gioiTinhSelect.value;
        const nameParts = nameVal.split(/\s+/);
        const lastName = nameParts[nameParts.length - 1] || "";

        // (1) Year
        const lYearVal = currentCalculation.lunarYearCanChi;

        // (2) Month
        const lMonthVal = parseInt(thangAlSelect.value);
        const LUNAR_MONTH_SINO = {
            1: "Chánh ngoạt", 2: "Nhị ngoạt", 3: "Tam ngoạt", 4: "Tứ ngoạt", 5: "Ngũ ngoạt",
            6: "Lục ngoạt", 7: "Thất ngoạt", 8: "Bát ngoạt", 9: "Cửu ngoạt", 10: "Thập ngoạt",
            11: "Thập nhứt ngoạt", 12: "Nhứt thập nhị ngoạt"
        };
        let monthSino = LUNAR_MONTH_SINO[lMonthVal] || `${lMonthVal} ngoạt`;
        if (nhuanCheckbox.checked) {
            monthSino = "Nhuận " + monthSino;
        }

        // (3) Day
        const lDayVal = parseInt(ngayAlSelect.value);
        const LUNAR_DAY_SINO = {
            1: "sơ nhất nhựt", 2: "sơ nhị nhựt", 3: "sơ tam nhựt", 4: "sơ tứ nhựt", 5: "sơ ngũ nhựt",
            6: "sơ lục nhựt", 7: "sơ thất nhựt", 8: "sơ bát nhựt", 9: "sơ cửu nhựt", 10: "Thập nhựt",
            11: "Thập nhứt nhựt", 12: "Thập nhị nhựt", 13: "Thập tam nhựt", 14: "Thập tứ nhựt", 15: "Thập ngũ nhựt",
            16: "Thập lục nhựt", 17: "Thập thất nhựt", 18: "Thập bát nhựt", 19: "Thập cửu nhựt", 20: "Nhị thập nhựt",
            21: "Nhị thập nhứt nhựt", 22: "Nhị thập nhị nhựt", 23: "Nhị thập tam nhựt", 24: "Nhị thập  tứ nhựt", 25: "Nhị thập ngũ nhựt",
            26: "Nhị thập lục nhựt", 27: "Nhị thập thất nhựt", 28: "Nhị thập bát nhựt", 29: "Nhị thập cửu nhựt", 30: "Tam thập nhựt"
        };
        const daySino = LUNAR_DAY_SINO[lDayVal] || `${lDayVal} nhựt`;

        // (4) Nơi cúng
        const noiCung = noiCungSelect.value;

        // (5) Bàn cúng
        const banCung = banCungSelect.value;

        // (6) Tuần cửu
        const currentTuanCuuSelect = document.getElementById("tuanCuuSelect");
        const val6Raw = currentTuanCuuSelect ? currentTuanCuuSelect.value : "Tam cửu";
        const val6 = val6Raw.charAt(0).toUpperCase() + val6Raw.slice(1);

        // (7) Phẩm đạo formatting
        let val7 = "";
        if (phamDao === "Đạo hữu") {
            val7 = nameVal;
        } else if (phamDao === "Lễ Sanh" && gioiTinh === "Nữ") {
            val7 = `Hương ${lastName}`;
        } else if (phamDao === "Lễ Sanh" && gioiTinh === "Nam") {
            val7 = `Lễ Sanh……….${lastName} Thanh`;
        } else if (phamDao === "Chức việc") {
            val7 = ` ………………..${nameVal}`;
        }

        // (8) Họ tên
        const val8 = nameVal;

        // (9) Tuổi Hán Việt
        const val9 = toSinoVietnamese(currentAge);

        // (10) Cúng xã
        const cungXaSelected = soCungXaSelect.value;
        let cungXaName = cungXaSelected.replace("Xã ", "").replace("Phường ", "");
        const cungXaType = cungXaSelected.includes("Phường") ? "phường" : "xã";

        // (11) Birthplace district
        const huyenSelected = noiSinhHuyenSelect.value;
        const huyenType = huyenSelected.includes("Thành phố") ? "thành phố" : "huyện";
        let huyenName = huyenSelected.replace("Huyện ", "").replace("Thành phố ", "");

        // (12) Birthplace commune
        const xaSelected = noiSinhXaSelect.value;
        const xaType = xaSelected.includes("Phường") ? "phường" : (xaSelected.includes("Thị trấn") ? "thị trấn" : "xã");
        let xaName = xaSelected.replace("Xã ", "").replace("Phường ", "").replace("Thị trấn ", "");

        // (13) Year of death Can Chi
        const deathYear = currentCalculation.lunarYearCanChi;

        // (14) Month of death Hán Việt
        const deathMonthMap = {
            1: "Chánh", 2: "Nhị", 3: "Tam", 4: "Tứ", 5: "Ngũ", 6: "Lục",
            7: "Thất", 8: "Bát", 9: "Cửu", 10: "Thập", 11: "Thập nhứt", 12: "Thập nhị"
        };
        let deathMonth = deathMonthMap[lMonthVal] || String(lMonthVal);
        if (nhuanCheckbox.checked) {
            deathMonth = "Nhuận " + deathMonth;
        }

        // (15) Day of death Hán Việt
        const daySinoFull = LUNAR_DAY_SINO[lDayVal] || `${lDayVal} nhựt`;
        const deathDay = daySinoFull.replace(/\s*nhựt\s*$/, "");

        // (16) Hour of death
        const deathHour = gioMatSelect.value;

        const html = `
            <p class="center-text">ĐẠI ĐẠO TAM KỲ PHỔ ĐỘ</p>
            <p class="center-text">Đệ Bá Nhứt Niên</p>
            <p class="center-text">TAM GIÁO QUI NGUYÊN, NGŨ CHI HIỆP NHỨT</p>
            <p class="center-text" style="font-weight: bold; margin-top: 1.5rem;">THỜI DUY</p>
            <p>Thiên vận <span class="highlight-val">${lYearVal.toUpperCase()}</span> niên <span class="highlight-val">${monthSino.toUpperCase()}</span>, <span class="highlight-val">${daySino.toUpperCase()}</span> Ngọ thời. Hiện tại Việt Nam Quốc, Bến Tre Tỉnh, <span class="highlight-val">${cungXaName.toUpperCase()}</span> ${cungXaType}, Cư trụ <span class="highlight-val">${noiCung.toUpperCase()}</span> Chi Trung.</p>
            <p>Kim Hữu Đệ tử thọ Thiên ân <span class="highlight-val">${phamDao.toUpperCase()}</span>, cộng đồng chư Chức sắc hiệp giữ Thiện Nam Tín Nữ đẳng, lãnh đái Hiếu quyến quì tại <span class="highlight-val">${banCung.toUpperCase()}</span> thành tâm trình tấu.</p>
            <p class="center-text" style="font-weight: bold; margin-top: 1.5rem;">HUỲNH KIM KHUYẾT NỘI</p>
            <p class="center-text" style="font-weight: bold;">HUYỀN KHUNG CAO THƯỢNG ĐẾ NGỌC HOÀNG ĐẠI THIÊN TÔN</p>
            <p class="center-text">Viết: Cao Đài Tiên Ông Đại Bồ Tát Ma Ha Tát</p>
            <p class="center-text">TAM KỲ PHỔ ĐỘ TAM TRẤN OAI NGHI</p>
            <p class="center-text">Thường Cư Nam Hải Quan Âm Như Lai</p>
            <p class="center-text">Lý Đại Tiên Trưởng Thái Bạch Kim Tinh</p>
            <p class="center-text">Hiệp Thiên Đại Đế Quan Thánh Đế Quân</p>
            <p class="center-text">Chư Phật Tiên Thánh Thần Liên Đài Chi Hạ</p>
            <p>Kim vì <span class="highlight-val">${val8.toUpperCase()}</span> bổn mạng ư <span class="highlight-val">${val9.toUpperCase()}</span> tuế thọ sanh tại Bến Tre Tỉnh <span class="highlight-val">${huyenName.toUpperCase()}</span> ${huyenType} <span class="highlight-val">${xaName.toUpperCase()}</span> ${xaType} hạnh ngộ Tam Kỳ giáo hóa dĩ qui giác lộ chơn truyền bất hạnh ư <span class="highlight-val">${deathYear.toUpperCase()}</span> niên <span class="highlight-val">${deathMonth.toUpperCase()}</span> ngoạt <span class="highlight-val">${deathDay.toUpperCase()}</span> nhựt <span class="highlight-val">${deathHour.toUpperCase()}</span> thời. Khí thế mạng chung kim chí <span class="highlight-val">${val6.toUpperCase()}</span> Chi Tuần.</p>
            <p>Tư thỉnh cầu chư Chức sắc hiệp giữ Thiện Nam Tín Nữ đẳng, nghiêm thiết Đàn tràng hương, đăng, hoa, trà, quả thanh chước vi nghi, thành tâm phụng hiến khẩn cầu siêu độ.</p>
            <p style="font-weight: bold; text-indent: 0; margin-top: 1.5rem;">PHỤC VỌNG</p>
            <p>Vô trung Từ Phụ phát hạ thâm ân xá giữ linh hồn <span class="highlight-val">${val7.toUpperCase()}</span> đắc thoát ly u khổ, siêu thăng thượng giới tự tại tiêu diêu.</p>
            <p>Thiên phong giữ thiện tín đẳng.</p>
            <p class="center-text" style="font-weight: bold; margin-top: 1.5rem;">THIỆT TRIÊM</p>
            <p class="center-text">Thiên ân bất thăng khẩn khất chi chí cẩn biểu khấu bái</p>
            <p style="text-align: center; text-indent: 0; white-space: pre; font-family: monospace; margin-top: 1.5rem;">            THƯỢNG TẤU                    DĨ VĂN</p>
        `;
        document.getElementById("causieu-preview-text").innerHTML = html;
        document.getElementById("causieu-event-badge").innerText = val6;
    }

    // Render Linh Vị HTML Preview
    function renderLinhViPreview() {
        if (!currentCalculation) return;

        const nameVal = hoTenInput.value.trim() || "CHƯA NHẬP HỌ TÊN";
        const phamDao = phamDaoSelect.value;
        const gioiTinh = gioiTinhSelect.value;
        const nameParts = nameVal.split(/\s+/);
        const lastName = nameParts[nameParts.length - 1] || "";
        const firstName = nameParts[0] || "";

        // (1a), (1b) split birth year can chi
        const birthParts = currentBirthCanChi.split(/\s+/);
        const val1a = birthParts[0] || "";
        const val1b = birthParts[1] || "";

        // (2a), (2b) birth province (Bến Tre is fixed)
        const val2a = "Bến";
        const val2b = "Tre";

        // (3a), (3b) split birthplace district
        const huyenSelected = noiSinhHuyenSelect.value;
        let huyenName = huyenSelected.replace("Huyện ", "").replace("Thành phố ", "");
        const huyenParts = huyenName.split(/\s+/);
        let val3a = "";
        let val3b = "";
        if (huyenParts.length === 2) {
            val3a = huyenParts[0];
            val3b = huyenParts[1];
        } else if (huyenParts.length === 3) {
            val3a = huyenParts[0] + " " + huyenParts[1];
            val3b = huyenParts[2];
        } else {
            val3a = huyenParts[0] || "";
            val3b = "";
        }

        // (4a), (4b) split birthplace commune
        const xaSelected = noiSinhXaSelect.value;
        let xaName = xaSelected.replace("Xã ", "").replace("Phường ", "").replace("Thị trấn ", "");
        const xaParts = xaName.split(/\s+/);
        let val4a = "";
        let val4b = "";
        if (xaParts.length === 2) {
            val4a = xaParts[0];
            val4b = xaParts[1];
        } else if (xaParts.length === 3) {
            val4a = xaParts[0] + " " + xaParts[1];
            val4b = xaParts[2];
        } else if (xaParts.length === 4) {
            val4a = xaParts[0] + " " + xaParts[1];
            val4b = xaParts[2] + " " + xaParts[3];
        } else {
            val4a = xaParts[0] || "";
            val4b = "";
        }

        // (5a), (5b) split pham dao
        const phamParts = phamDao.split(/\s+/);
        const val5a = phamParts[0] || "";
        const val5b = phamParts[1] || "";

        // (6) Họ
        const val6 = firstName;

        // (7) Tên
        const val7 = lastName;

        // (8) Tuổi term (Dương, Linh, Thọ, Hạ Thọ, Trung Thọ, Thượng Thọ, Mạo)
        let val8 = "Thọ";
        const term = currentAgeTerm.toLowerCase();
        if (term.includes("dương")) val8 = "Dương";
        else if (term.includes("linh")) val8 = "Linh";
        else if (term.includes("trung thọ")) val8 = "Trung Thọ";
        else if (term.includes("thượng thọ")) val8 = "Thượng Thọ";
        else if (term.includes("hạ thọ")) val8 = "Hạ Thọ";
        else if (term.includes("mạo")) val8 = "Mạo";
        else if (term.includes("thọ")) val8 = "Thọ";

        // (9), wordThap, (10) split age Hán Việt
        const units = {
            1: "Nhứt", 2: "Nhị", 3: "Tam", 4: "Tứ", 5: "Ngũ",
            6: "Lục", 7: "Thất", 8: "Bát", 9: "Cửu"
        };
        let val9 = "";
        let val10 = "";
        let wordThap = "Thập";
        
        if (currentAge < 10) {
            val9 = "";
            wordThap = "";
            val10 = units[currentAge] || "";
        } else if (currentAge === 10) {
            val9 = "";
            wordThap = "Thập";
            val10 = "";
        } else if (currentAge < 20) {
            val9 = "";
            wordThap = "Thập";
            val10 = units[currentAge - 10] || "";
        } else {
            const chuc = Math.floor(currentAge / 10);
            const donVi = currentAge % 10;
            val9 = units[chuc] || "";
            wordThap = "Thập";
            val10 = donVi === 0 ? "" : (units[donVi] || "");
        }

        // (11a), (11b) split death year can chi
        const deathParts = currentCalculation.lunarYearCanChi.split(/\s+/);
        const val11a = deathParts[0] || "";
        const val11b = deathParts[1] || "";

        // (12) month of death Hán Việt
        const lMonthVal = parseInt(thangAlSelect.value);
        const LUNAR_MONTH_LINHVI = {
            1: "Chánh", 2: "Nhị", 3: "Tam", 4: "Tứ", 5: "Ngũ", 6: "Lục",
            7: "Thất", 8: "Bát", 9: "Cửu", 10: "Thập", 11: "Thập Nhứt", 12: "Nhứt Thập Nhị"
        };
        let val12 = LUNAR_MONTH_LINHVI[lMonthVal] || lMonthVal;
        if (nhuanCheckbox.checked) {
            val12 = "Nhuận " + val12;
        }

        // (13a), (13b) split day of death Hán Việt
        const lDayVal = parseInt(ngayAlSelect.value);
        const LUNAR_DAY_SINO = {
            1: "sơ nhất nhựt", 2: "sơ nhị nhựt", 3: "sơ tam nhựt", 4: "sơ tứ nhựt", 5: "sơ ngũ nhựt",
            6: "sơ lục nhựt", 7: "sơ thất nhựt", 8: "sơ bát nhựt", 9: "sơ cửu nhựt", 10: "Thập nhựt",
            11: "Thập nhứt nhựt", 12: "Thập nhị nhựt", 13: "Thập tam nhựt", 14: "Thập tứ nhựt", 15: "Thập ngũ nhựt",
            16: "Thập lục nhựt", 17: "Thập thất nhựt", 18: "Thập bát nhựt", 19: "Thập cửu nhựt", 20: "Nhị thập nhựt",
            21: "Nhị thập nhứt nhựt", 22: "Nhị thập nhị nhựt", 23: "Nhị thập tam nhựt", 24: "Nhị thập  tứ nhựt", 25: "Nhị thập ngũ nhựt",
            26: "Nhị thập lục nhựt", 27: "Nhị thập thất nhựt", 28: "Nhị thập bát nhựt", 29: "Nhị thập cửu nhựt", 30: "Tam thập nhựt"
        };
        const daySinoFull = LUNAR_DAY_SINO[lDayVal] || `${lDayVal} nhựt`;
        const dayText = daySinoFull.replace(/\s*nhựt\s*$/, "");
        const dayParts = dayText.split(/\s+/);
        let val13a = "";
        let val13b = "";
        if (dayParts.length === 1) {
            val13a = dayParts[0];
            val13b = "";
        } else if (dayParts.length === 2) {
            if (dayParts[1] === "thập") {
                val13a = dayParts[0] + " " + dayParts[1];
                val13b = "";
            } else {
                val13a = dayParts[0];
                val13b = dayParts[1];
            }
        } else if (dayParts.length === 3) {
            val13a = dayParts[0] + " " + dayParts[1];
            val13b = dayParts[2];
        }

        // (14) hour of death
        const val14 = gioMatSelect.value;

        // (15a), (15b) split cúng commune
        const cungXaSelected = soCungXaSelect.value;
        let cungXaName = cungXaSelected.replace("Xã ", "").replace("Phường ", "");
        const cungXaParts = cungXaName.split(/\s+/);
        let val15a = "";
        let val15b = "";
        if (cungXaParts.length === 2) {
            val15a = cungXaParts[0];
            val15b = cungXaParts[1];
        } else if (cungXaParts.length === 3) {
            val15a = cungXaParts[0] + " " + cungXaParts[1];
            val15b = cungXaParts[2];
        } else if (cungXaParts.length === 4) {
            val15a = cungXaParts[0] + " " + cungXaParts[1];
            val15b = cungXaParts[2] + " " + cungXaParts[3];
        } else {
            val15a = cungXaParts[0] || "";
            val15b = "";
        }

        function wordSpan(txt, isHigh = true) {
            if (!txt) return "";
            const isSmall = txt.length > 3 || txt.includes(" ");
            return `<span class="linhvi-word ${isSmall ? 'small-font' : ''} ${isHigh ? 'highlight-val' : ''}">${txt.toUpperCase()}</span>`;
        }

        function staticSpan(txt) {
            return `<span class="linhvi-word">${txt.toUpperCase()}</span>`;
        }

        let col1Html = `
            ${staticSpan("SANH")}
            ${staticSpan("Ư")}
            ${wordSpan(val1a)}
            ${wordSpan(val1b)}
            ${staticSpan("NIÊN")}
            ${staticSpan("CANH")}
            ${gioiTinh === "Nữ" ? wordSpan(val2a) : staticSpan("BẾN")}
            ${gioiTinh === "Nữ" ? wordSpan(val2b) : staticSpan("TRE")}
            ${staticSpan("TỈNH")}
            ${wordSpan(val3a)}
            ${wordSpan(val3b)}
            ${staticSpan("HUYỆN")}
            ${wordSpan(val4a)}
            ${wordSpan(val4b)}
            ${staticSpan("XÃ")}
        `;

        let col2Html = `
            ${staticSpan("PHỤNG")}
            ${staticSpan("THỈNH")}
            ${staticSpan("VONG")}
            ${staticSpan("LINH")}
            ${wordSpan(val5a)}
            ${wordSpan(val5b)}
            ${staticSpan("TÁNH")}
            ${wordSpan(val6)}
            ${staticSpan(gioiTinh === "Nữ" ? "HIỆU" : "TỰ")}
            ${wordSpan(val7)}
            ${staticSpan("HƯỞNG")}
            ${wordSpan(val8)}
            ${wordSpan(val9)}
            ${wordThap ? wordSpan(wordThap) : ""}
            ${wordSpan(val10)}
            ${staticSpan("TUẾ")}
            ${staticSpan("NHỨT")}
            ${staticSpan("VỊ")}
            ${staticSpan("THẦN")}
            ${staticSpan("HỒN")}
            ${staticSpan("TỌA VỊ")}
        `;

        let col3Html = `
            ${staticSpan("TỬ")}
            ${staticSpan("Ư")}
            ${wordSpan(val11a)}
            ${wordSpan(val11b)}
            ${staticSpan("NIÊN")}
            ${wordSpan(val12)}
            ${staticSpan("NGOẠT")}
            ${wordSpan(val13a)}
            ${wordSpan(val13b)}
            ${staticSpan("NHỰT")}
            ${wordSpan(val14)}
            ${staticSpan("THỜI")}
            ${staticSpan("NHI")}
            ${staticSpan("CHUNG")}
            ${wordSpan(val15a)}
            ${wordSpan(val15b)}
            ${staticSpan("XÃ")}
        `;

        document.getElementById("linhvi-preview-tablet").innerHTML = `
            <div class="linhvi-column">${col1Html}</div>
            <div class="linhvi-column">${col2Html}</div>
            <div class="linhvi-column">${col3Html}</div>
        `;
        document.getElementById("linhvi-gender-badge").innerText = `Linh Vị ${gioiTinh}`;
    }

    // Update active tab preview helper
    function updateActiveTabPreview() {
        const activeTab = document.querySelector(".tab-btn.active");
        if (activeTab) {
            const tabId = activeTab.getAttribute("data-tab");
            if (tabId === "causieu") {
                renderCauSieuPreview();
            } else if (tabId === "linhvi") {
                renderLinhViPreview();
            }
        }
    }

    // Tab buttons and click handlers
    const tabButtons = document.querySelectorAll(".tab-btn");
    const tabPanels = document.querySelectorAll(".tab-panel");

    function switchTab(tabId) {
        tabButtons.forEach(btn => {
            if (btn.getAttribute("data-tab") === tabId) {
                btn.classList.add("active");
            } else {
                btn.classList.remove("active");
            }
        });
        tabPanels.forEach(panel => {
            if (panel.id === `panel-${tabId}`) {
                panel.classList.add("active");
            } else {
                panel.classList.remove("active");
            }
        });
    }

    tabButtons.forEach(btn => {
        btn.addEventListener("click", () => {
            const tabId = btn.getAttribute("data-tab");
            switchTab(tabId);
            if (tabId === "causieu") {
                renderCauSieuPreview();
            } else if (tabId === "linhvi") {
                renderLinhViPreview();
            }
        });
    });

    const btnViewCauSieu = document.getElementById("btn-view-causieu");
    const btnViewLinhVi = document.getElementById("btn-view-linhvi");

    btnViewCauSieu.addEventListener("click", () => {
        if (!hoTenInput.value.trim()) {
            alert("Vui lòng nhập họ tên người mất!");
            hoTenInput.focus();
            return;
        }
        if (!currentCalculation) {
            alert("Vui lòng thực hiện tính toán lịch trước!");
            return;
        }
        switchTab("causieu");
        renderCauSieuPreview();
    });

    btnViewLinhVi.addEventListener("click", () => {
        if (!hoTenInput.value.trim()) {
            alert("Vui lòng nhập họ tên người mất!");
            hoTenInput.focus();
            return;
        }
        if (!currentCalculation) {
            alert("Vui lòng thực hiện tính toán lịch trước!");
            return;
        }
        switchTab("linhvi");
        renderLinhViPreview();
    });

    // Run initial calculation on page load
    runRecalculation();
});

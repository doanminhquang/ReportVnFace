document.addEventListener('DOMContentLoaded', function () {
    const now = new Date();
    const currentMonth = now.getMonth() + 1;
    const currentYear = now.getFullYear();

    document.getElementById('monthInput').value = currentMonth < 10 ? '0' + currentMonth : currentMonth;
    document.getElementById('yearInput').value = currentYear;
    
    document.getElementById('firstCheckinTime').value = localStorage.getItem('firstCheckinTime') || '07:30';
    document.getElementById('lastCheckinTime').value = localStorage.getItem('lastCheckinTime') || '17:00';

    document.getElementById('firstCheckinTime').addEventListener('change', function(e) {
        localStorage.setItem('firstCheckinTime', e.target.value);
    });

    document.getElementById('lastCheckinTime').addEventListener('change', function(e) {
        localStorage.setItem('lastCheckinTime', e.target.value);
    });
});

document.getElementById('fetchDataBtn').addEventListener('click', function () {
    const month = document.getElementById('monthInput').value;
    const year = document.getElementById('yearInput').value;
    const firstCheckinTime = document.getElementById('firstCheckinTime').value;
    const lastCheckinTime = document.getElementById('lastCheckinTime').value;

    if (month && year) {
        chrome.tabs.query({ active: true, currentWindow: true }, function (tabs) {
            for (let i = 0; i < tabs.length; i++) {
                let tab = tabs[i];
                if (tab.url && tab.url.includes("https://console-vnface.vnpt.vn/")) {
                    chrome.scripting.executeScript({
                        target: { tabId: tab.id },
                        func: fetchCheckinDataForMonthAndYear,
                        args: [month, year, firstCheckinTime, lastCheckinTime]
                    });

                    window.close();
                    break;
                }
            }
        });
    }
});

async function fetchCheckinDataForMonthAndYear(month, year, firstCheckinTime, lastCheckinTime) {
    const FirstCheckinMinutes = convertTimeToMinutes(firstCheckinTime);
    const LastCheckinMinutes = convertTimeToMinutes(lastCheckinTime);

    async function fetchCheckinData(month, year) {
        const startDate = new Date(year, month - 1, 1);
        const endDate = new Date(year, month, 0);
    
        const formattedStartDate = startDate.toLocaleString("en-GB", { timeZone: "Asia/Ho_Chi_Minh" }).split(",")[0].replace(/\//g, "%2F");
        const formattedEndDate = endDate.toLocaleString("en-GB", { timeZone: "Asia/Ho_Chi_Minh" }).split(",")[0].replace(/\//g, "%2F");
    
        const url = `https://api-vnface.vnpt.vn/checkin-service/his-checkin/list-filter?startDate=${formattedStartDate}%2000:00:00&endDate=${formattedEndDate}%2023:59:59&keySearch=&page=1&maxSize=50&uuidDevice=&type=ALL&minFirstCheckin=00:00&maxFirstCheckin=23:59&minLastCheckin=00:00&maxLastCheckin=23:59`;
    
        try {
            const response = await fetch(url, {
                headers: {
                    "accept": "application/json, text/plain, */*",
                    "authorization": `Bearer ${localStorage.getItem("accessToken")}`,
                    "content-type": "application/json",
                },
            });
            return await response.json();
        } catch (error) {
            console.error("Không thể kết nối đến máy chủ:", error);
            return null;
        }
    }

    function downloadExcel(workbook, filename) {
        try {
            const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });
            const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        } catch (error) {
            console.error('Lỗi khi tạo file Excel:', error);
            alert(`Lỗi khi tạo file Excel: ${error.message}`);
        }
    }
    
    function convertTimeToMinutes(timeStr) {
        const [hours, minutes] = timeStr.split(':').map(Number);
        return hours * 60 + minutes;
    }
    
    function createCellStyle(fontSize, isBold, fillColor) {
        return {
            font: { name: 'Times New Roman', sz: fontSize, bold: isBold },
            fill: fillColor ? { fgColor: { rgb: fillColor } } : undefined,
            border: {
                top: { style: 'thin', color: { rgb: '000000' } },
                left: { style: 'thin', color: { rgb: '000000' } },
                bottom: { style: 'thin', color: { rgb: '000000' } },
                right: { style: 'thin', color: { rgb: '000000' } }
            }
        };
    }
    
    function createCell(value, fontSize, isBold, fillColor, type='s') {
        return {
            v: String(value),
            t: type,
            s: createCellStyle(fontSize, isBold, fillColor)
        };
    }
    
    function isCompliant(firstCheckin, lastCheckin) {
        const firstCheckinMinutes = timeToMinutes(firstCheckin);
        const lastCheckinMinutes = timeToMinutes(lastCheckin);
        return !(firstCheckinMinutes >= FirstCheckinMinutes || lastCheckinMinutes < LastCheckinMinutes);
    }
    
    function timeToMinutes(timeStr) {
        const [hours, minutes] = timeStr.split(":").map(Number);
        return hours * 60 + minutes;
    }
    
    function convertToExcel(data, compliantDays, totalDays) {
        const headers = [
            createCell('Ngày điểm danh', 14, true),
            createCell('Điểm danh lần đầu', 14, true),
            createCell('Điểm danh lần cuối', 14, true),
            createCell('Số lần điểm danh', 14, true)
        ];
    
        const rows = data.map(item => {
            const isCompliantFirstCheckin = timeToMinutes(item.firstCheckin) >= FirstCheckinMinutes;
            const isCompliantLastCheckin = timeToMinutes(item.lastCheckin) < LastCheckinMinutes;
            return [
                createCell(item.dateCheckin, 12, false),
                createCell(item.firstCheckin, 12, false, isCompliantFirstCheckin ? "FF2C2C" : "FFFFFF"),
                createCell(item.lastCheckin, 12, false, isCompliantLastCheckin ? "FF2C2C" : "FFFFFF"),
                createCell(item.totalCheckin, 12, false)                
            ];
        });
    
        const summaryRow = [
            createCell('Tuân thủ: ' + compliantDays + '/' + totalDays, 12, true),
            createCell('', 12, true),
            createCell('', 12, true),
            createCell('', 12, true)
        ];
    
        rows.push(summaryRow);
    
        return [headers, ...rows];
    }
    
    function createCalendarSheet(checkinInfo, month, year, FirstCheckinMinutes, LastCheckinMinutes) {
        function timeToMinutes(timeStr) {
            const [hours, minutes] = timeStr.split(":").map(Number);
            return hours * 60 + minutes;
        }
    
        function createCell(value, fontSize, isBold, fillColor) {
            return {
                v: value,
                t: 's',
                s: {
                    font: { name: 'Times New Roman', sz: fontSize, bold: isBold },
                    fill: fillColor ? { fgColor: { rgb: fillColor } } : undefined,
                    border: {
                        top: { style: 'thin', color: { rgb: '000000' } },
                        left: { style: 'thin', color: { rgb: '000000' } },
                        bottom: { style: 'thin', color: { rgb: '000000' } },
                        right: { style: 'thin', color: { rgb: '000000' } }
                    },
                    alignment: { wrapText: true, vertical: 'center', horizontal: 'center' }
                }
            };
        }
    
        const headers = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7', 'Chủ nhật'].map(day => 
            createCell(day, 12, true, 'CCCCCC')
        );
    
        const firstDay = new Date(year, month - 1, 1);
        const lastDay = new Date(year, month, 0);
        const totalDays = lastDay.getDate();
    
        let calendar = [headers];
        let currentWeek = Array(7).fill(createCell('', 11, false, 'FFFFFF'));
        let dayOfWeek = firstDay.getDay();
        dayOfWeek = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
    
        for (let day = 1; day <= totalDays; day++) {
            const dayData = checkinInfo.find(item => {
                const [itemDay] = item.dateCheckin.split('/');
                return parseInt(itemDay) === day;
            });
    
            let cellContent = day < 10 ? "0" + day.toString() + "/" + month + "/" + year: day.toString() + "/" + month + "/" + year;
            let fillColor = 'FFFFFF';
    
            if (dayData) {
                const isCompliantFirst = timeToMinutes(dayData.firstCheckin) >= FirstCheckinMinutes;
                const isCompliantLast = timeToMinutes(dayData.lastCheckin) < LastCheckinMinutes;
                
                cellContent = `${dayData.dateCheckin}\n${dayData.firstCheckin}\n${dayData.lastCheckin}`;
                
                if (isCompliantFirst || isCompliantLast) {
                    fillColor = 'FF2C2C';
                }
            }
    
            currentWeek[dayOfWeek] = createCell(cellContent, 11, false, fillColor);
    
            if (dayOfWeek === 6 || day === totalDays) {
                calendar.push([...currentWeek]);
                currentWeek = Array(7).fill(createCell('', 11, false, 'FFFFFF'));
            }
    
            dayOfWeek = (dayOfWeek + 1) % 7;
        }
    
        calendar.forEach((row, index) => {
            if (!calendar['!rows']) calendar['!rows'] = [];
            calendar['!rows'][index] = { hpt: 60 }; 
        });
    
        return calendar;
    }
    
    function s2ab(s) {
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i < s.length; i++) {
            view[i] = s.charCodeAt(i) & 0xFF;
        }
        return buf;
    }    

    try {
        const data = await fetchCheckinData(month, year);
        if (!data) {
            throw new Error('Không thể lấy dữ liệu từ máy chủ');
        }
        if (!data.object?.data?.length) {
            throw new Error('Không tìm thấy dữ liệu điểm danh trong khoảng thời gian đã chọn');
        }

        let compliantDays = 0;
        const totalDays = data.object.data.length;
        const checkinInfo = data.object.data.map(item => {
            const { dateCheckin, firstCheckin, lastCheckin, totalCheckin } = item;
            if (!isCompliant(firstCheckin, lastCheckin)) {
                compliantDays++;
            }
            return { dateCheckin, firstCheckin, lastCheckin, totalCheckin };
        });

        const userName = data.object.data[0].username;
        const fileName = `[${userName}]_checkin_vnface_${month}-${year}.xlsx`;

        const workbook = XLSX.utils.book_new();

        const sheetData = convertToExcel(checkinInfo, compliantDays, totalDays);
        const ws = XLSX.utils.aoa_to_sheet(sheetData);
        ws['!cols'] = [{ wch: 20 }, { wch: 25 }, { wch: 25 }, { wch: 22 }];
        XLSX.utils.book_append_sheet(workbook, ws, 'Dữ liệu điểm danh vnFace');

        const calendarData = createCalendarSheet(checkinInfo, month, year, FirstCheckinMinutes, LastCheckinMinutes);
        const wsCalendar = XLSX.utils.aoa_to_sheet(calendarData);
        wsCalendar['!cols'] = Array(7).fill({ wch: 25 });
        XLSX.utils.book_append_sheet(workbook, wsCalendar, 'Lịch điểm danh');

        downloadExcel(workbook, fileName);
    } catch (error) {
        console.error('Lỗi xử lý dữ liệu:', error);
        alert(`Lỗi xử lý dữ liệu: ${error.message}`);
    }
}

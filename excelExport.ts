import * as XLSX from 'xlsx';
import { DowntimeReport } from './types';
import { calculateDowntimeCost } from './timeHelpers';
import { Product } from './data';

/**
 * Xuất toàn bộ dữ liệu hệ thống (Báo cáo và Sản phẩm) ra file Excel để backup
 */
export function exportFullSystemDataToExcel(reports: DowntimeReport[], products: Product[]) {
  const wb = XLSX.utils.book_new();

  // 1. Sheet Báo cáo dừng Line (Dữ liệu thô để re-import)
  const reportData = reports.map(r => ({
    'ID': r.id,
    'Ngày': r.date,
    'Ca': r.shift,
    'Chuyền': r.line,
    'Thiết bị': r.equipment,
    'Sản phẩm': r.productName,
    'Đơn giá': r.productUnitPrice,
    'Năng suất ĐM': r.standardRate,
    'Bắt đầu': r.startTime,
    'Kết thúc': r.endTime || '',
    'Thời gian': r.duration,
    'Nhóm nguyên nhân': r.reasonCategory,
    'Chi tiết': r.details,
    'Giải pháp': r.solution || '',
    'Người báo cáo': r.pic,
    'Trạng thái': r.status,
    'Ngày tạo': r.createdAt
  }));
  const wsReports = XLSX.utils.json_to_sheet(reportData);
  XLSX.utils.book_append_sheet(wb, wsReports, 'Báo cáo Dừng Line');

  // 2. Sheet Danh mục Sản phẩm
  const productData = products.map(p => ({
    'Tên Sản Phẩm': p.name,
    'Dây Chuyền': p.line,
    'Đơn Giá (VNĐ)': p.unitPrice,
    'Năng Suất Định Mức (SP/Giờ)': p.standardRate
  }));
  const wsProducts = XLSX.utils.json_to_sheet(productData);
  XLSX.utils.book_append_sheet(wb, wsProducts, 'Danh mục Sản phẩm');

  // 3. Xuất file
  const fileName = `Sao_luu_he_thong_Sunhouse_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(wb, fileName);
}

/**
 * Phân tích file Excel backup để lấy lại toàn bộ dữ liệu
 */
export async function importFullSystemDataFromExcel(file: File): Promise<{ reports: DowntimeReport[], products: Product[] }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        
        let reports: DowntimeReport[] = [];
        let products: Product[] = [];

        // Đọc sheet Báo cáo
        const wsReports = workbook.Sheets['Báo cáo Dừng Line'];
        if (wsReports) {
          const rawReports = XLSX.utils.sheet_to_json(wsReports) as any[];
          reports = rawReports.map(r => ({
            id: String(r['ID'] || Date.now() + Math.random()),
            date: String(r['Ngày']),
            shift: String(r['Ca']) as any,
            line: String(r['Chuyền']) as any,
            equipment: String(r['Thiết bị']),
            productName: String(r['Sản phẩm']),
            productUnitPrice: Number(r['Đơn giá']),
            standardRate: Number(r['Năng suất ĐM']),
            startTime: String(r['Bắt đầu']),
            endTime: r['Kết thúc'] ? String(r['Kết thúc']) : undefined,
            duration: Number(r['Thời gian']),
            reasonCategory: String(r['Nhóm nguyên nhân']) as any,
            details: String(r['Chi tiết']),
            solution: r['Giải pháp'] ? String(r['Giải pháp']) : undefined,
            pic: String(r['Người báo cáo']),
            status: String(r['Trạng thái']) as any,
            createdAt: r['Ngày tạo'] ? String(r['Ngày tạo']) : new Date().toISOString()
          }));
        }

        // Đọc sheet Sản phẩm
        const wsProducts = workbook.Sheets['Danh mục Sản phẩm'];
        if (wsProducts) {
          const rawProducts = XLSX.utils.sheet_to_json(wsProducts) as any[];
          products = rawProducts.map(p => ({
            name: String(p['Tên Sản Phẩm']),
            line: String(p['Dây Chuyền']) as any,
            unitPrice: Number(p['Đơn Giá (VNĐ)']),
            standardRate: Number(p['Năng Suất Định Mức (SP/Giờ)'])
          }));
        }

        resolve({ reports, products });
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = (error) => reject(error);
    reader.readAsBinaryString(file);
  });
}

/**
 * Xuất danh sách báo cáo dừng Line ra file Excel chuyên nghiệp
 * @param reports Danh sách báo cáo cần xuất
 * @param filtersDescription Chuỗi mô tả điều kiện lọc hiện tại để ghi vào báo cáo
 */
export function exportDowntimeReportsToExcel(reports: DowntimeReport[], filtersDescription: string) {
  // 1. Chuẩn bị dữ liệu thô dạng bảng phẳng
  const rawData = reports.map((r, index) => ({
    'STT': index + 1,
    'Ngày báo cáo': r.date,
    'Ca làm việc': r.shift,
    'Line sản xuất': r.line,
    'Thiết bị/Công đoạn': r.equipment,
    'Sản phẩm đang sản xuất': r.productName || 'N/A',
    'Đơn giá sản phẩm (VNĐ)': r.productUnitPrice || 0,
    'Năng suất định mức (Sp/giờ)': r.standardRate || 0,
    'Giờ bắt đầu': r.startTime,
    'Giờ kết thúc': r.endTime || 'Chưa khắc phục',
    'Thời gian dừng (phút)': r.duration,
    'Chi phí dừng Line (VNĐ)': r.status === 'Đang xử lý' ? 0 : calculateDowntimeCost(r.duration),
    'Tổn thất Doanh thu (VNĐ)': r.status === 'Đang xử lý' ? 0 : Math.round((r.duration / 60) * (r.standardRate || 0) * (r.productUnitPrice || 0)),
    'Nhóm nguyên nhân': r.reasonCategory,
    'Chi tiết sự cố': r.details,
    'Biện pháp khắc phục': r.solution || 'Đang xử lý',
    'Người báo cáo/PIC': r.pic,
    'Trạng thái': r.status,
  }));

  // 2. Tạo workbook và worksheet trống
  const wb = XLSX.utils.book_new();
  
  // 3. Tạo tiêu đề và thông tin chung trước bảng dữ liệu chính
  const titleRows = [
    ['CÔNG TY CỔ PHẦN TẬP ĐOÀN SUNHOUSE'],
    ['BÁO CÁO THỜI GIAN DỪNG LINE SẢN XUẤT (LINE DOWNTIME REPORT)'],
    [`Ngày xuất báo cáo: ${new Date().toLocaleDateString('vi-VN')} ${new Date().toLocaleTimeString('vi-VN')}`],
    [`Điều kiện lọc: ${filtersDescription}`],
    [], // Dòng trống cách biệt
  ];

  // 4. Tạo worksheet từ mảng tiêu đề trước
  const ws = XLSX.utils.aoa_to_sheet(titleRows);

  // 5. Thêm bảng dữ liệu chính tiếp nối vào dòng thứ 6 (index 5)
  XLSX.utils.sheet_add_json(ws, rawData, {
    origin: 'A6',
    skipHeader: false,
  });

  // 6. Tính toán các thông số tổng hợp để thêm vào cuối bảng
  const totalDowntime = reports.reduce((sum, r) => sum + r.duration, 0);
  const totalIncidents = reports.length;
  const avgDuration = totalIncidents > 0 ? Math.round(totalDowntime / totalIncidents) : 0;
  const activeIncidents = reports.filter(r => r.status === 'Đang xử lý').length;
  const totalCost = reports.reduce((sum, r) => sum + (r.status === 'Đang xử lý' ? 0 : calculateDowntimeCost(r.duration)), 0);
  const totalRevenueLoss = reports.reduce((sum, r) => sum + (r.status === 'Đang xử lý' ? 0 : Math.round((r.duration / 60) * (r.standardRate || 0) * (r.productUnitPrice || 0))), 0);

  const startSummaryRow = 6 + rawData.length + 2; // Cách bảng chính 2 dòng
  
  const summaryRows = [
    [],
    ['TỔNG HỢP THỐNG KÊ (SUMMARY REPORT)'],
    ['Tổng số vụ dừng Line:', totalIncidents, 'vụ'],
    ['Tổng thời gian dừng Line:', totalDowntime, 'phút'],
    ['Tổng chi phí dừng Line ước tính:', totalCost, 'VNĐ'],
    ['Tổng tổn thất doanh thu ước tính:', totalRevenueLoss, 'VNĐ'],
    ['Thời gian dừng trung bình (MTTR):', avgDuration, 'phút/vụ'],
    ['Sự cố chưa khắc phục (Đang xử lý):', activeIncidents, 'vụ'],
  ];

  XLSX.utils.sheet_add_aoa(ws, summaryRows, {
    origin: `A${startSummaryRow}`,
  });

  // 7. Cấu hình độ rộng các cột (Column Widths) tự động dựa trên nội dung
  const colWidths = [
    { wch: 6 },   // STT
    { wch: 15 },  // Ngày báo cáo
    { wch: 12 },  // Ca làm việc
    { wch: 20 },  // Line sản xuất
    { wch: 25 },  // Thiết bị/Công đoạn
    { wch: 35 },  // Sản phẩm đang sản xuất
    { wch: 22 },  // Đơn giá sản phẩm (VNĐ)
    { wch: 25 },  // Năng suất định mức (Sp/giờ)
    { wch: 12 },  // Giờ bắt đầu
    { wch: 15 },  // Giờ kết thúc
    { wch: 22 },  // Thời gian dừng (phút)
    { wch: 25 },  // Chi phí dừng Line (VNĐ)
    { wch: 25 },  // Tổn thất Doanh thu (VNĐ)
    { wch: 25 },  // Nhóm nguyên nhân
    { wch: 45 },  // Chi tiết sự cố
    { wch: 45 },  // Biện pháp khắc phục
    { wch: 25 },  // Người báo cáo/PIC
    { wch: 15 },  // Trạng thái
  ];
  ws['!cols'] = colWidths;

  // 8. Đính worksheet vào workbook
  XLSX.utils.book_append_sheet(wb, ws, 'Báo cáo Dừng Line');

  // 9. Xuất file và tải xuống trực tiếp ở trình duyệt
  const fileName = `Bao_cao_dung_line_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(wb, fileName);
}

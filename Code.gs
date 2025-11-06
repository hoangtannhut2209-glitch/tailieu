// Code.gs - Backend tối ưu với load-once strategy

// Hàm chính để serve web app
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Theo dõi tăng ca')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl("https://png.pngtree.com/png-clipart/20240531/original/pngtree-the-icon-of-timesheets-for-excel-vector-picture-image_15483131.png");
}

// Include files
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// LOAD ALL DATA - Gọi một lần duy nhất
// ==========================================
function getAllData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Chấm công');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: true,
        employees: [],
        attendanceData: {},
        timestamp: new Date().getTime()
      };
    }
    
    const employees = [];
    const attendanceData = {};
    
    // Parse tất cả dữ liệu trong một lần đọc
    for (let i = 1; i < data.length; i++) {
      const employeeId = data[i][0];
      
      // Thông tin nhân viên
      employees.push({
        id: employeeId,
        code: data[i][1],
        name: data[i][2],
        department: data[i][3],
        avatar: data[i][19] || '',
        username: data[i][16] || '',
        password: data[i][17] || '',
        role: data[i][18] || ''
      });
      
      // Dữ liệu chấm công 12 tháng
      attendanceData[employeeId] = {};
      for (let month = 1; month <= 12; month++) {
        const monthColumn = 4 + (month - 1);
        const jsonData = data[i][monthColumn] || '{}';
        try {
          attendanceData[employeeId][month] = JSON.parse(jsonData);
        } catch (e) {
          attendanceData[employeeId][month] = {};
        }
      }
    }
    
    return {
      success: true,
      employees: employees,
      attendanceData: attendanceData,
      timestamp: new Date().getTime()
    };
    
  } catch (error) {
    console.error('Error getting all data:', error);
    return {
      success: false,
      message: 'Lỗi tải dữ liệu: ' + error.toString(),
      employees: [],
      attendanceData: {}
    };
  }
}

// ==========================================
// ATTENDANCE CRUD OPERATIONS
// ==========================================

// Lưu dữ liệu chấm công cho một nhân viên
function saveAttendanceData(employeeId, month, year, attendanceData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Chấm công');
    const data = sheet.getDataRange().getValues();
    
    let employeeRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === employeeId) {
        employeeRow = i + 1;
        break;
      }
    }
    
    if (employeeRow === -1) {
      throw new Error('Không tìm thấy Nhân sự');
    }
    
    const monthColumn = 4 + month;
    
    // Format JSON
    let jsonString;
    if (Object.keys(attendanceData).length === 0) {
      jsonString = '{}';
    } else {
      const entries = Object.entries(attendanceData);
      const formattedEntries = entries.map(([day, dayData]) => {
        let dayEntry = `"${day}":{"status":"${dayData.status}","timestamp":"${dayData.timestamp}"`;
        
        if (dayData.overtimeHours && dayData.overtimeHours > 0) {
          dayEntry += `,"overtimeHours":${dayData.overtimeHours},"overtimeMultiplier":${dayData.overtimeMultiplier}`;
        }
        
        dayEntry += '}';
        return dayEntry;
      });
      jsonString = `{${formattedEntries.join(',\n')}}`;
    }
    
    sheet.getRange(employeeRow, monthColumn).setValue(jsonString);
    
    return { success: true, message: 'Đã lưu thành công' };
    
  } catch (error) {
    console.error('Error saving attendance:', error);
    return { success: false, message: 'Lỗi lưu dữ liệu: ' + error.toString() };
  }
}

// Bulk save - Lưu nhiều nhân viên cùng lúc
function saveBulkAttendanceData(bulkData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Chấm công');
    const data = sheet.getDataRange().getValues();
    const results = [];
    
    for (const item of bulkData) {
      try {
        const { employeeId, month, year, attendanceData } = item;
        
        let employeeRow = -1;
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] === employeeId) {
            employeeRow = i + 1;
            break;
          }
        }
        
        if (employeeRow === -1) {
          results.push({ 
            employeeId, 
            success: false, 
            message: 'Không tìm thấy nhân viên' 
          });
          continue;
        }
        
        const monthColumn = 4 + month;
        
        let jsonString;
        if (Object.keys(attendanceData).length === 0) {
          jsonString = '{}';
        } else {
          const entries = Object.entries(attendanceData);
          const formattedEntries = entries.map(([day, dayData]) => {
            let dayEntry = `"${day}":{"status":"${dayData.status}","timestamp":"${dayData.timestamp}"`;
            
            if (dayData.overtimeHours && dayData.overtimeHours > 0) {
              dayEntry += `,"overtimeHours":${dayData.overtimeHours},"overtimeMultiplier":${dayData.overtimeMultiplier}`;
            }
            
            dayEntry += '}';
            return dayEntry;
          });
          jsonString = `{${formattedEntries.join(',\n')}}`;
        }
        
        sheet.getRange(employeeRow, monthColumn).setValue(jsonString);
        
        results.push({ 
          employeeId, 
          success: true, 
          message: 'Đã lưu thành công' 
        });
        
      } catch (itemError) {
        console.error(`Error processing ${item.employeeId}:`, itemError);
        results.push({ 
          employeeId: item.employeeId, 
          success: false, 
          message: itemError.toString() 
        });
      }
    }
    
    return { 
      success: true, 
      message: `Đã xử lý ${bulkData.length} bản ghi`, 
      results: results 
    };
    
  } catch (error) {
    console.error('Error in bulk save:', error);
    return { 
      success: false, 
      message: 'Lỗi bulk save: ' + error.toString(),
      results: []
    };
  }
}

// ==========================================
// EMPLOYEE CRUD OPERATIONS
// ==========================================

// Thêm Nhân sự mới
function addEmployee(employeeData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Chấm công');
    const data = sheet.getDataRange().getValues();
    
    let maxId = 0;
    for (let i = 1; i < data.length; i++) {
      const id = data[i][0];
      if (id && id.startsWith('ID')) {
        const num = parseInt(id.substring(2));
        if (num > maxId) maxId = num;
      }
    }
    
    const newId = `ID${String(maxId + 1).padStart(3, '0')}`;
    
    const newRow = [
      newId,
      employeeData.code || newId,
      employeeData.name,
      employeeData.department,
      '{}', '{}', '{}', '{}', '{}', '{}',
      '{}', '{}', '{}', '{}', '{}', '{}',
      employeeData.username || '',
      employeeData.password || '',
      employeeData.role || '',
      employeeData.avatar || ''
    ];
    
    sheet.appendRow(newRow);
    
    return { 
      success: true, 
      message: 'Đã thêm Nhân sự thành công', 
      id: newId,
      employee: {
        id: newId,
        code: employeeData.code || newId,
        name: employeeData.name,
        department: employeeData.department,
        username: employeeData.username || '',
        password: employeeData.password || '',
        role: employeeData.role || ''
      }
    };
  } catch (error) {
    console.error('Error adding employee:', error);
    return { success: false, message: error.toString() };
  }
}

// Cập nhật Nhân sự
function updateEmployee(employeeId, employeeData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Chấm công');
    const data = sheet.getDataRange().getValues();
    
    let employeeRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === employeeId) {
        employeeRow = i + 1;
        break;
      }
    }
    
    if (employeeRow === -1) {
      throw new Error('Không tìm thấy Nhân sự');
    }
    
    sheet.getRange(employeeRow, 2).setValue(employeeData.code || employeeId);
    sheet.getRange(employeeRow, 3).setValue(employeeData.name);
    sheet.getRange(employeeRow, 4).setValue(employeeData.department);
    sheet.getRange(employeeRow, 17).setValue(employeeData.username || '');
    sheet.getRange(employeeRow, 18).setValue(employeeData.password || '');
    sheet.getRange(employeeRow, 19).setValue(employeeData.role || '');
    sheet.getRange(employeeRow, 20).setValue(employeeData.avatar || '');
    
    return { 
      success: true, 
      message: 'Đã cập nhật thành công',
      employee: {
        id: employeeId,
        code: employeeData.code || employeeId,
        name: employeeData.name,
        department: employeeData.department,
        username: employeeData.username || '',
        password: employeeData.password || '',
        role: employeeData.role || ''
      }
    };
  } catch (error) {
    console.error('Error updating employee:', error);
    return { success: false, message: error.toString() };
  }
}

// Xóa Nhân sự
function deleteEmployee(employeeId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Chấm công');
    const data = sheet.getDataRange().getValues();
    
    let employeeRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === employeeId) {
        employeeRow = i + 1;
        break;
      }
    }
    
    if (employeeRow === -1) {
      throw new Error('Không tìm thấy Nhân sự');
    }
    
    sheet.deleteRow(employeeRow);
    
    return { success: true, message: 'Đã xóa Nhân sự thành công' };
  } catch (error) {
    console.error('Error deleting employee:', error);
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// AUTHENTICATION
// ==========================================

function authenticateUser(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Chấm công');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { success: false, message: 'Không có dữ liệu người dùng' };
    }
    
    for (let i = 1; i < data.length; i++) {
      const userLogin = data[i][16];
      const userPassword = data[i][17];
      const userRole = data[i][18];
      
      if (userLogin === username && userPassword === password) {
        const sessionToken = Utilities.getUuid();
        const sessionData = {
          userId: data[i][0],
          username: username,
          name: data[i][2],
          role: userRole,
          loginTime: new Date().getTime()
        };
        
        PropertiesService.getScriptProperties().setProperty(
          'session_' + sessionToken, 
          JSON.stringify(sessionData)
        );
        
        return { 
          success: true, 
          token: sessionToken,
          user: sessionData
        };
      }
    }
    
    return { success: false, message: 'Tên đăng nhập hoặc mật khẩu không đúng' };
  } catch (error) {
    console.error('Error authenticating user:', error);
    return { success: false, message: 'Lỗi xác thực: ' + error.toString() };
  }
}

function validateSession(token) {
  try {
    const sessionData = PropertiesService.getScriptProperties().getProperty('session_' + token);
    if (!sessionData) {
      return { success: false, message: 'Session không tồn tại' };
    }
    
    const session = JSON.parse(sessionData);
    const now = new Date().getTime();
    const sessionAge = now - session.loginTime;
    
    if (sessionAge > 24 * 60 * 60 * 1000) {
      PropertiesService.getScriptProperties().deleteProperty('session_' + token);
      return { success: false, message: 'Session đã hết hạn' };
    }
    
    return { success: true, user: session };
  } catch (error) {
    console.error('Error validating session:', error);
    return { success: false, message: 'Lỗi validate session' };
  }
}

function logoutUser(token) {
  try {
    PropertiesService.getScriptProperties().deleteProperty('session_' + token);
    return { success: true, message: 'Đã đăng xuất thành công' };
  } catch (error) {
    console.error('Error logging out:', error);
    return { success: false, message: 'Lỗi đăng xuất' };
  }
}
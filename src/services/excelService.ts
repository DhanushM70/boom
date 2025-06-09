import * as XLSX from 'xlsx';
import { SystemData, BorrowRequest, Component, User, LoginSession } from '../types';

export class ExcelService {
  private static instance: ExcelService;

  static getInstance(): ExcelService {
    if (!ExcelService.instance) {
      ExcelService.instance = new ExcelService();
    }
    return ExcelService.instance;
  }

  exportToExcel(data: SystemData): void {
    const workbook = XLSX.utils.book_new();
    
    // Create styled sheets
    this.addSummarySheet(workbook, data);
    this.addRequestsSheet(workbook, data.requests);
    this.addComponentsSheet(workbook, data.components);
    this.addUsersSheet(workbook, data.users);
    this.addLoginSessionsSheet(workbook, data.loginSessions);

    // Generate filename with timestamp
    const timestamp = new Date().toISOString().split('T')[0];
    const filename = `Isaac-Asimov-Lab-Report-${timestamp}.xlsx`;

    // Write and download
    XLSX.writeFile(workbook, filename);
  }

  private addSummarySheet(workbook: XLSX.WorkBook, data: SystemData): void {
    const stats = this.calculateStats(data);
    const now = new Date();

    const summaryData = [
      ['Isaac Asimov Robotics Lab - System Report'],
      ['Generated on:', now.toLocaleString()],
      ['Report Period:', 'All Time'],
      [''],
      ['SYSTEM OVERVIEW'],
      ['Metric', 'Value', 'Status'],
      ['Total Users', stats.totalUsers, stats.totalUsers > 0 ? 'Active' : 'No Users'],
      ['Active Users', stats.activeUsers, `${((stats.activeUsers / stats.totalUsers) * 100).toFixed(1)}%`],
      ['Total Logins', stats.totalLogins, 'Cumulative'],
      ['Currently Online', stats.onlineUsers, stats.onlineUsers > 0 ? 'Users Online' : 'No Active Sessions'],
      [''],
      ['COMPONENT MANAGEMENT'],
      ['Metric', 'Value', 'Status'],
      ['Total Components', stats.totalComponents, 'Available Types'],
      ['Total Requests', stats.totalRequests, 'All Time'],
      ['Pending Requests', stats.pendingRequests, stats.pendingRequests > 0 ? 'Needs Attention' : 'All Clear'],
      ['Overdue Items', stats.overdueItems, stats.overdueItems > 0 ? 'Action Required' : 'On Track'],
      [''],
      ['REQUEST STATUS BREAKDOWN'],
      ['Status', 'Count', 'Percentage'],
      ['Pending', data.requests.filter(r => r.status === 'pending').length, `${((data.requests.filter(r => r.status === 'pending').length / Math.max(data.requests.length, 1)) * 100).toFixed(1)}%`],
      ['Approved', data.requests.filter(r => r.status === 'approved').length, `${((data.requests.filter(r => r.status === 'approved').length / Math.max(data.requests.length, 1)) * 100).toFixed(1)}%`],
      ['Rejected', data.requests.filter(r => r.status === 'rejected').length, `${((data.requests.filter(r => r.status === 'rejected').length / Math.max(data.requests.length, 1)) * 100).toFixed(1)}%`],
      ['Returned', data.requests.filter(r => r.status === 'returned').length, `${((data.requests.filter(r => r.status === 'returned').length / Math.max(data.requests.length, 1)) * 100).toFixed(1)}%`],
      [''],
      ['COMPONENT CATEGORIES'],
      ['Category', 'Total Units', 'Available', 'Utilization'],
      ...this.getComponentCategorySummary(data.components)
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(summaryData);
    
    // Apply styling
    this.styleWorksheet(worksheet, {
      headerRow: 0,
      titleStyle: { font: { bold: true, size: 16 }, fill: { fgColor: { rgb: "00CED1" } } },
      headerStyle: { font: { bold: true }, fill: { fgColor: { rgb: "E6F3FF" } } },
      dataStyle: { alignment: { horizontal: "left" } }
    });

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Summary');
  }

  private addRequestsSheet(workbook: XLSX.WorkBook, requests: BorrowRequest[]): void {
    const headers = [
      'Request ID',
      'Request Date',
      'Student Name',
      'Roll Number',
      'Mobile',
      'Component',
      'Quantity',
      'Due Date',
      'Status',
      'Approved By',
      'Approved Date',
      'Returned Date',
      'Days Overdue',
      'Notes'
    ];

    const rows = requests.map(request => {
      const dueDate = new Date(request.dueDate);
      const now = new Date();
      const daysOverdue = request.status === 'approved' && dueDate < now 
        ? Math.ceil((now.getTime() - dueDate.getTime()) / (1000 * 60 * 60 * 24))
        : 0;

      return [
        request.id,
        new Date(request.requestDate).toLocaleDateString(),
        request.studentName,
        request.rollNo,
        request.mobile,
        request.componentName,
        request.quantity,
        new Date(request.dueDate).toLocaleDateString(),
        request.status.toUpperCase(),
        request.approvedBy || '',
        request.approvedAt ? new Date(request.approvedAt).toLocaleDateString() : '',
        request.returnedAt ? new Date(request.returnedAt).toLocaleDateString() : '',
        daysOverdue > 0 ? daysOverdue : '',
        request.notes || ''
      ];
    });

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    
    // Apply conditional formatting for status
    this.styleWorksheet(worksheet, {
      headerRow: 0,
      headerStyle: { font: { bold: true }, fill: { fgColor: { rgb: "4CAF50" } } },
      conditionalFormatting: {
        statusColumn: 8, // Status column
        overdueColumn: 12 // Days Overdue column
      }
    });

    // Set column widths
    worksheet['!cols'] = [
      { width: 15 }, { width: 12 }, { width: 20 }, { width: 15 },
      { width: 15 }, { width: 25 }, { width: 10 }, { width: 12 },
      { width: 12 }, { width: 15 }, { width: 12 }, { width: 12 },
      { width: 12 }, { width: 30 }
    ];

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Requests');
  }

  private addComponentsSheet(workbook: XLSX.WorkBook, components: Component[]): void {
    const headers = [
      'Component ID',
      'Name',
      'Category',
      'Total Quantity',
      'Available Quantity',
      'Borrowed Quantity',
      'Utilization %',
      'Stock Status',
      'Description'
    ];

    const rows = components.map(component => {
      const borrowed = component.totalQuantity - component.availableQuantity;
      const utilization = component.totalQuantity > 0 
        ? ((borrowed / component.totalQuantity) * 100).toFixed(1)
        : '0';
      
      const stockStatus = component.availableQuantity === 0 ? 'OUT OF STOCK' :
                         component.availableQuantity < component.totalQuantity * 0.2 ? 'LOW STOCK' :
                         component.availableQuantity < component.totalQuantity * 0.5 ? 'MEDIUM STOCK' : 'GOOD STOCK';

      return [
        component.id,
        component.name,
        component.category,
        component.totalQuantity,
        component.availableQuantity,
        borrowed,
        `${utilization}%`,
        stockStatus,
        component.description || ''
      ];
    });

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    
    this.styleWorksheet(worksheet, {
      headerRow: 0,
      headerStyle: { font: { bold: true }, fill: { fgColor: { rgb: "FF9800" } } }
    });

    worksheet['!cols'] = [
      { width: 15 }, { width: 25 }, { width: 15 }, { width: 12 },
      { width: 15 }, { width: 15 }, { width: 12 }, { width: 15 }, { width: 40 }
    ];

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Components');
  }

  private addUsersSheet(workbook: XLSX.WorkBook, users: User[]): void {
    const headers = [
      'User ID',
      'Name',
      'Email',
      'Role',
      'Registration Date',
      'Last Login',
      'Total Logins',
      'Currently Active',
      'Account Status'
    ];

    const rows = users.map(user => [
      user.id,
      user.name,
      user.email,
      user.role.toUpperCase(),
      new Date(user.registeredAt).toLocaleDateString(),
      user.lastLoginAt ? new Date(user.lastLoginAt).toLocaleDateString() : 'Never',
      user.loginCount || 0,
      user.isActive ? 'Yes' : 'No',
      user.loginCount && user.loginCount > 0 ? 'Active' : 'Inactive'
    ]);

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    
    this.styleWorksheet(worksheet, {
      headerRow: 0,
      headerStyle: { font: { bold: true }, fill: { fgColor: { rgb: "9C27B0" } } }
    });

    worksheet['!cols'] = [
      { width: 15 }, { width: 25 }, { width: 30 }, { width: 10 },
      { width: 15 }, { width: 15 }, { width: 12 }, { width: 15 }, { width: 15 }
    ];

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Users');
  }

  private addLoginSessionsSheet(workbook: XLSX.WorkBook, sessions: LoginSession[]): void {
    const headers = [
      'Session ID',
      'User Name',
      'Email',
      'Role',
      'Login Time',
      'Logout Time',
      'Duration (minutes)',
      'Device',
      'Status',
      'Session Quality'
    ];

    const rows = sessions.map(session => {
      const duration = session.sessionDuration 
        ? Math.round(session.sessionDuration / 60000)
        : session.isActive 
          ? Math.round((new Date().getTime() - new Date(session.loginTime).getTime()) / 60000)
          : 0;

      const sessionQuality = duration > 60 ? 'Long Session' :
                            duration > 30 ? 'Medium Session' :
                            duration > 5 ? 'Short Session' : 'Brief Session';

      return [
        session.id,
        session.userName,
        session.userEmail,
        session.userRole.toUpperCase(),
        new Date(session.loginTime).toLocaleString(),
        session.logoutTime ? new Date(session.logoutTime).toLocaleString() : 'Active',
        duration,
        session.deviceInfo || 'Unknown',
        session.isActive ? 'Active' : 'Ended',
        sessionQuality
      ];
    });

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    
    this.styleWorksheet(worksheet, {
      headerRow: 0,
      headerStyle: { font: { bold: true }, fill: { fgColor: { rgb: "2196F3" } } }
    });

    worksheet['!cols'] = [
      { width: 20 }, { width: 20 }, { width: 30 }, { width: 10 },
      { width: 20 }, { width: 20 }, { width: 15 }, { width: 15 }, { width: 10 }, { width: 15 }
    ];

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Login Sessions');
  }

  private styleWorksheet(worksheet: XLSX.WorkSheet, options: any): void {
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    
    // Style header row
    if (options.headerRow !== undefined) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: options.headerRow, c: col });
        if (!worksheet[cellAddress]) continue;
        
        worksheet[cellAddress].s = options.headerStyle || {
          font: { bold: true },
          fill: { fgColor: { rgb: "E6F3FF" } },
          alignment: { horizontal: "center" }
        };
      }
    }

    // Apply title style to first row if specified
    if (options.titleStyle && range.e.r >= 0) {
      const titleCell = worksheet['A1'];
      if (titleCell) {
        titleCell.s = options.titleStyle;
      }
    }
  }

  private calculateStats(data: SystemData): any {
    const now = new Date();
    const overdueItems = data.requests.filter(r => 
      r.status === 'approved' && new Date(r.dueDate) < now
    );

    return {
      totalUsers: data.users.length,
      activeUsers: data.users.filter(u => u.isActive).length,
      totalLogins: data.users.reduce((sum, u) => sum + (u.loginCount || 0), 0),
      onlineUsers: data.loginSessions.filter(s => s.isActive).length,
      totalRequests: data.requests.length,
      pendingRequests: data.requests.filter(r => r.status === 'pending').length,
      totalComponents: data.components.length,
      overdueItems: overdueItems.length
    };
  }

  private getComponentCategorySummary(components: Component[]): string[][] {
    const categories = components.reduce((acc, comp) => {
      if (!acc[comp.category]) {
        acc[comp.category] = { total: 0, available: 0 };
      }
      acc[comp.category].total += comp.totalQuantity;
      acc[comp.category].available += comp.availableQuantity;
      return acc;
    }, {} as Record<string, { total: number; available: number }>);

    return Object.entries(categories).map(([category, data]) => {
      const utilization = data.total > 0 ? (((data.total - data.available) / data.total) * 100).toFixed(1) : '0';
      return [
        category,
        data.total.toString(),
        data.available.toString(),
        `${utilization}%`
      ];
    });
  }

  generatePreviewData(data: SystemData): any {
    const stats = this.calculateStats(data);
    
    return {
      summary: {
        totalUsers: stats.totalUsers,
        activeUsers: stats.activeUsers,
        totalRequests: stats.totalRequests,
        pendingRequests: stats.pendingRequests,
        totalComponents: stats.totalComponents,
        overdueItems: stats.overdueItems
      },
      recentRequests: data.requests
        .sort((a, b) => new Date(b.requestDate).getTime() - new Date(a.requestDate).getTime())
        .slice(0, 10),
      componentUtilization: this.getComponentCategorySummary(data.components),
      userActivity: data.users
        .filter(u => u.loginCount && u.loginCount > 0)
        .sort((a, b) => (b.loginCount || 0) - (a.loginCount || 0))
        .slice(0, 10)
    };
  }
}

export const excelService = ExcelService.getInstance();
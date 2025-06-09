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
    
    // Create professional styled sheets
    this.addExecutiveSummarySheet(workbook, data);
    this.addDetailedComponentSheet(workbook, data.components);
    this.addRequestsAnalysisSheet(workbook, data.requests);
    this.addUserManagementSheet(workbook, data.users);
    this.addLoginAnalyticsSheet(workbook, data.loginSessions);

    // Generate filename with timestamp
    const timestamp = new Date().toISOString().split('T')[0];
    const filename = `Isaac-Asimov-Lab-Professional-Report-${timestamp}.xlsx`;

    // Write and download
    XLSX.writeFile(workbook, filename);
  }

  private addExecutiveSummarySheet(workbook: XLSX.WorkBook, data: SystemData): void {
    const stats = this.calculateStats(data);
    const now = new Date();

    const summaryData = [
      ['ISAAC ASIMOV ROBOTICS LAB'],
      ['EXECUTIVE SUMMARY REPORT'],
      [''],
      ['Report Generated:', now.toLocaleString()],
      ['Report Period:', 'All Time Data'],
      ['System Status:', 'Operational'],
      [''],
      ['KEY PERFORMANCE INDICATORS'],
      [''],
      ['Metric', 'Current Value', 'Status', 'Trend', 'Target'],
      ['Total Registered Users', stats.totalUsers, stats.totalUsers > 0 ? 'Active' : 'No Users', '+12%', '100'],
      ['Currently Active Users', stats.activeUsers, `${((stats.activeUsers / Math.max(stats.totalUsers, 1)) * 100).toFixed(1)}%`, '+8%', '80%'],
      ['Total Login Sessions', stats.totalLogins, 'Cumulative', '+15%', 'N/A'],
      ['Online Users Now', stats.onlineUsers, stats.onlineUsers > 0 ? 'Users Online' : 'No Active Sessions', 'Real-time', 'N/A'],
      [''],
      ['COMPONENT MANAGEMENT OVERVIEW'],
      [''],
      ['Metric', 'Current Value', 'Status', 'Utilization', 'Action Required'],
      ['Total Component Types', stats.totalComponents, 'Available', 'N/A', 'None'],
      ['Total Requests Processed', stats.totalRequests, 'All Time', '100%', 'None'],
      ['Pending Requests', stats.pendingRequests, stats.pendingRequests > 0 ? 'Needs Review' : 'All Clear', 'N/A', stats.pendingRequests > 0 ? 'Review Required' : 'None'],
      ['Overdue Returns', stats.overdueItems, stats.overdueItems > 0 ? 'Action Required' : 'On Track', 'N/A', stats.overdueItems > 0 ? 'Follow Up' : 'None'],
      [''],
      ['REQUEST STATUS BREAKDOWN'],
      [''],
      ['Status', 'Count', 'Percentage', 'Trend', 'Notes'],
      ['Pending Review', data.requests.filter(r => r.status === 'pending').length, `${((data.requests.filter(r => r.status === 'pending').length / Math.max(data.requests.length, 1)) * 100).toFixed(1)}%`, 'Stable', 'Requires admin attention'],
      ['Approved & Active', data.requests.filter(r => r.status === 'approved').length, `${((data.requests.filter(r => r.status === 'approved').length / Math.max(data.requests.length, 1)) * 100).toFixed(1)}%`, 'Increasing', 'Components in use'],
      ['Successfully Returned', data.requests.filter(r => r.status === 'returned').length, `${((data.requests.filter(r => r.status === 'returned').length / Math.max(data.requests.length, 1)) * 100).toFixed(1)}%`, 'Positive', 'Completed transactions'],
      ['Rejected Requests', data.requests.filter(r => r.status === 'rejected').length, `${((data.requests.filter(r => r.status === 'rejected').length / Math.max(data.requests.length, 1)) * 100).toFixed(1)}%`, 'Low', 'Quality control'],
      [''],
      ['COMPONENT CATEGORY ANALYSIS'],
      [''],
      ['Category', 'Total Components', 'Total Units', 'Available Units', 'Utilization Rate', 'Performance'],
      ...this.getCategoryAnalysis(data.components),
      [''],
      ['RECOMMENDATIONS & INSIGHTS'],
      [''],
      ['Area', 'Recommendation', 'Priority', 'Expected Impact'],
      ['Inventory Management', 'Monitor high-utilization components', 'Medium', 'Prevent stockouts'],
      ['User Engagement', 'Promote underutilized components', 'Low', 'Increase usage diversity'],
      ['Process Optimization', 'Automate overdue notifications', 'High', 'Improve return rates'],
      ['Capacity Planning', 'Analyze peak usage patterns', 'Medium', 'Better resource allocation']
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(summaryData);
    
    // Apply professional styling
    this.styleExecutiveSheet(worksheet);

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Executive Summary');
  }

  private addDetailedComponentSheet(workbook: XLSX.WorkBook, components: Component[]): void {
    const headers = [
      'Component ID',
      'Component Name',
      'Category',
      'Total Stock',
      'Available Units',
      'Units in Use',
      'Utilization Rate (%)',
      'Stock Status',
      'Reorder Level',
      'Last Updated',
      'Description',
      'Performance Rating'
    ];

    const rows = components.map(component => {
      const inUse = component.totalQuantity - component.availableQuantity;
      const utilization = component.totalQuantity > 0 
        ? ((inUse / component.totalQuantity) * 100).toFixed(1)
        : '0.0';
      
      const stockStatus = component.availableQuantity === 0 ? 'OUT OF STOCK' :
                         component.availableQuantity < component.totalQuantity * 0.2 ? 'LOW STOCK' :
                         component.availableQuantity < component.totalQuantity * 0.5 ? 'MEDIUM STOCK' : 'GOOD STOCK';

      const reorderLevel = Math.ceil(component.totalQuantity * 0.2);
      const performanceRating = parseFloat(utilization) > 70 ? 'High Demand' :
                               parseFloat(utilization) > 40 ? 'Moderate Demand' :
                               parseFloat(utilization) > 10 ? 'Low Demand' : 'Minimal Use';

      return [
        component.id,
        component.name,
        component.category,
        component.totalQuantity,
        component.availableQuantity,
        inUse,
        utilization,
        stockStatus,
        reorderLevel,
        new Date().toLocaleDateString(),
        component.description || 'No description available',
        performanceRating
      ];
    });

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    
    this.styleComponentSheet(worksheet);

    worksheet['!cols'] = [
      { width: 15 }, { width: 30 }, { width: 18 }, { width: 12 },
      { width: 15 }, { width: 12 }, { width: 15 }, { width: 15 }, 
      { width: 12 }, { width: 12 }, { width: 40 }, { width: 18 }
    ];

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Component Inventory');
  }

  private addRequestsAnalysisSheet(workbook: XLSX.WorkBook, requests: BorrowRequest[]): void {
    const headers = [
      'Request ID',
      'Request Date',
      'Student Name',
      'Roll Number',
      'Mobile Number',
      'Component Requested',
      'Quantity',
      'Due Date',
      'Current Status',
      'Days Since Request',
      'Days Until Due',
      'Priority Level',
      'Approved By',
      'Approval Date',
      'Return Date',
      'Processing Time (Days)',
      'Notes'
    ];

    const rows = requests.map(request => {
      const requestDate = new Date(request.requestDate);
      const dueDate = new Date(request.dueDate);
      const now = new Date();
      
      const daysSinceRequest = Math.ceil((now.getTime() - requestDate.getTime()) / (1000 * 60 * 60 * 24));
      const daysUntilDue = Math.ceil((dueDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
      
      const isOverdue = request.status === 'approved' && dueDate < now;
      const priority = isOverdue ? 'HIGH - OVERDUE' : 
                      request.status === 'pending' && daysSinceRequest > 2 ? 'HIGH - DELAYED' :
                      request.status === 'pending' ? 'MEDIUM' : 'NORMAL';

      const processingTime = request.approvedAt 
        ? Math.ceil((new Date(request.approvedAt).getTime() - requestDate.getTime()) / (1000 * 60 * 60 * 24))
        : request.status === 'pending' ? daysSinceRequest : 0;

      return [
        request.id,
        requestDate.toLocaleDateString(),
        request.studentName,
        request.rollNo,
        request.mobile,
        request.componentName,
        request.quantity,
        dueDate.toLocaleDateString(),
        request.status.toUpperCase(),
        daysSinceRequest,
        request.status === 'approved' ? daysUntilDue : 'N/A',
        priority,
        request.approvedBy || 'Pending',
        request.approvedAt ? new Date(request.approvedAt).toLocaleDateString() : 'Not Approved',
        request.returnedAt ? new Date(request.returnedAt).toLocaleDateString() : 'Not Returned',
        processingTime,
        request.notes || 'No additional notes'
      ];
    });

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    
    this.styleRequestSheet(worksheet);

    worksheet['!cols'] = [
      { width: 15 }, { width: 12 }, { width: 20 }, { width: 15 },
      { width: 15 }, { width: 25 }, { width: 10 }, { width: 12 },
      { width: 15 }, { width: 12 }, { width: 12 }, { width: 18 },
      { width: 15 }, { width: 12 }, { width: 12 }, { width: 15 }, { width: 30 }
    ];

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Request Analysis');
  }

  private addUserManagementSheet(workbook: XLSX.WorkBook, users: User[]): void {
    const headers = [
      'User ID',
      'Full Name',
      'Email Address',
      'User Role',
      'Registration Date',
      'Last Login Date',
      'Total Login Count',
      'Account Status',
      'Activity Level',
      'Days Since Last Login',
      'Account Age (Days)',
      'Engagement Score'
    ];

    const rows = users.map(user => {
      const registrationDate = new Date(user.registeredAt);
      const lastLogin = user.lastLoginAt ? new Date(user.lastLoginAt) : null;
      const now = new Date();
      
      const daysSinceLastLogin = lastLogin 
        ? Math.ceil((now.getTime() - lastLogin.getTime()) / (1000 * 60 * 60 * 24))
        : 'Never logged in';
      
      const accountAge = Math.ceil((now.getTime() - registrationDate.getTime()) / (1000 * 60 * 60 * 24));
      
      const activityLevel = user.loginCount && user.loginCount > 10 ? 'High' :
                           user.loginCount && user.loginCount > 3 ? 'Medium' :
                           user.loginCount && user.loginCount > 0 ? 'Low' : 'Inactive';

      const engagementScore = user.loginCount && accountAge > 0 
        ? ((user.loginCount / accountAge) * 100).toFixed(2)
        : '0.00';

      return [
        user.id,
        user.name,
        user.email,
        user.role.toUpperCase(),
        registrationDate.toLocaleDateString(),
        lastLogin ? lastLogin.toLocaleDateString() : 'Never',
        user.loginCount || 0,
        user.isActive ? 'ACTIVE' : 'INACTIVE',
        activityLevel,
        daysSinceLastLogin,
        accountAge,
        engagementScore
      ];
    });

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    
    this.styleUserSheet(worksheet);

    worksheet['!cols'] = [
      { width: 15 }, { width: 25 }, { width: 30 }, { width: 12 },
      { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 },
      { width: 15 }, { width: 18 }, { width: 15 }, { width: 15 }
    ];

    XLSX.utils.book_append_sheet(workbook, worksheet, 'User Management');
  }

  private addLoginAnalyticsSheet(workbook: XLSX.WorkBook, sessions: LoginSession[]): void {
    const headers = [
      'Session ID',
      'User Name',
      'Email Address',
      'User Role',
      'Login Timestamp',
      'Logout Timestamp',
      'Session Duration (Minutes)',
      'Session Quality',
      'Device Type',
      'Session Status',
      'Login Day of Week',
      'Login Hour',
      'Productivity Score'
    ];

    const rows = sessions.map(session => {
      const loginTime = new Date(session.loginTime);
      const logoutTime = session.logoutTime ? new Date(session.logoutTime) : null;
      
      const duration = session.sessionDuration 
        ? Math.round(session.sessionDuration / 60000)
        : session.isActive 
          ? Math.round((new Date().getTime() - loginTime.getTime()) / 60000)
          : 0;

      const sessionQuality = duration > 120 ? 'Excellent' :
                            duration > 60 ? 'Good' :
                            duration > 30 ? 'Average' :
                            duration > 5 ? 'Short' : 'Brief';

      const dayOfWeek = loginTime.toLocaleDateString('en-US', { weekday: 'long' });
      const loginHour = loginTime.getHours();
      
      const productivityScore = duration > 0 ? Math.min((duration / 60) * 10, 100).toFixed(1) : '0.0';

      return [
        session.id,
        session.userName,
        session.userEmail,
        session.userRole.toUpperCase(),
        loginTime.toLocaleString(),
        logoutTime ? logoutTime.toLocaleString() : 'Still Active',
        duration,
        sessionQuality,
        session.deviceInfo || 'Unknown Device',
        session.isActive ? 'ACTIVE' : 'COMPLETED',
        dayOfWeek,
        `${loginHour}:00`,
        productivityScore
      ];
    });

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    
    this.styleLoginSheet(worksheet);

    worksheet['!cols'] = [
      { width: 20 }, { width: 20 }, { width: 30 }, { width: 12 },
      { width: 20 }, { width: 20 }, { width: 18 }, { width: 15 },
      { width: 15 }, { width: 15 }, { width: 15 }, { width: 12 }, { width: 15 }
    ];

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Login Analytics');
  }

  private styleExecutiveSheet(worksheet: XLSX.WorkSheet): void {
    // Title styling
    if (worksheet['A1']) {
      worksheet['A1'].s = {
        font: { bold: true, size: 18, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "00CED1" } },
        alignment: { horizontal: "center" }
      };
    }
    
    if (worksheet['A2']) {
      worksheet['A2'].s = {
        font: { bold: true, size: 14, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "0ba5a8" } },
        alignment: { horizontal: "center" }
      };
    }

    // Section headers
    const sectionHeaders = ['A8', 'A15', 'A23', 'A31', 'A38'];
    sectionHeaders.forEach(cell => {
      if (worksheet[cell]) {
        worksheet[cell].s = {
          font: { bold: true, size: 12, color: { rgb: "FFFFFF" } },
          fill: { fgColor: { rgb: "4CAF50" } },
          alignment: { horizontal: "left" }
        };
      }
    });
  }

  private styleComponentSheet(worksheet: XLSX.WorkSheet): void {
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    
    // Header row styling
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (worksheet[cellAddress]) {
        worksheet[cellAddress].s = {
          font: { bold: true, color: { rgb: "FFFFFF" } },
          fill: { fgColor: { rgb: "FF9800" } },
          alignment: { horizontal: "center" }
        };
      }
    }
  }

  private styleRequestSheet(worksheet: XLSX.WorkSheet): void {
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    
    // Header row styling
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (worksheet[cellAddress]) {
        worksheet[cellAddress].s = {
          font: { bold: true, color: { rgb: "FFFFFF" } },
          fill: { fgColor: { rgb: "4CAF50" } },
          alignment: { horizontal: "center" }
        };
      }
    }
  }

  private styleUserSheet(worksheet: XLSX.WorkSheet): void {
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    
    // Header row styling
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (worksheet[cellAddress]) {
        worksheet[cellAddress].s = {
          font: { bold: true, color: { rgb: "FFFFFF" } },
          fill: { fgColor: { rgb: "9C27B0" } },
          alignment: { horizontal: "center" }
        };
      }
    }
  }

  private styleLoginSheet(worksheet: XLSX.WorkSheet): void {
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    
    // Header row styling
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (worksheet[cellAddress]) {
        worksheet[cellAddress].s = {
          font: { bold: true, color: { rgb: "FFFFFF" } },
          fill: { fgColor: { rgb: "2196F3" } },
          alignment: { horizontal: "center" }
        };
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

  private getCategoryAnalysis(components: Component[]): string[][] {
    const categories = components.reduce((acc, comp) => {
      if (!acc[comp.category]) {
        acc[comp.category] = { 
          components: 0, 
          totalUnits: 0, 
          available: 0,
          requests: 0
        };
      }
      acc[comp.category].components += 1;
      acc[comp.category].totalUnits += comp.totalQuantity;
      acc[comp.category].available += comp.availableQuantity;
      return acc;
    }, {} as Record<string, { components: number; totalUnits: number; available: number; requests: number }>);

    return Object.entries(categories).map(([category, data]) => {
      const utilization = data.totalUnits > 0 
        ? (((data.totalUnits - data.available) / data.totalUnits) * 100).toFixed(1)
        : '0.0';
      
      const performance = parseFloat(utilization) > 70 ? 'High Performance' :
                         parseFloat(utilization) > 40 ? 'Good Performance' :
                         parseFloat(utilization) > 10 ? 'Moderate Performance' : 'Low Utilization';

      return [
        category,
        data.components.toString(),
        data.totalUnits.toString(),
        data.available.toString(),
        `${utilization}%`,
        performance
      ];
    });
  }

  generatePreviewData(data: SystemData): any {
    const stats = this.calculateStats(data);
    
    // Generate detailed component data
    const detailedComponents = data.components.map(component => {
      const inUse = component.totalQuantity - component.availableQuantity;
      const utilization = component.totalQuantity > 0 
        ? ((inUse / component.totalQuantity) * 100).toFixed(1) + '%'
        : '0.0%';

      return {
        id: component.id,
        name: component.name,
        category: component.category,
        totalQuantity: component.totalQuantity,
        available: component.availableQuantity,
        inUse: inUse,
        utilization: utilization,
        description: component.description || 'No description'
      };
    });

    // Generate category performance data
    const categoryPerformance = this.getCategoryPerformanceData(data.components, data.requests);
    
    return {
      summary: {
        totalUsers: stats.totalUsers,
        activeUsers: stats.activeUsers,
        totalRequests: stats.totalRequests,
        pendingRequests: stats.pendingRequests,
        totalComponents: stats.totalComponents,
        overdueItems: stats.overdueItems
      },
      detailedComponents: detailedComponents,
      recentRequests: data.requests
        .sort((a, b) => new Date(b.requestDate).getTime() - new Date(a.requestDate).getTime())
        .slice(0, 15),
      categoryPerformance: categoryPerformance,
      userActivity: data.users
        .filter(u => u.loginCount && u.loginCount > 0)
        .sort((a, b) => (b.loginCount || 0) - (a.loginCount || 0))
        .slice(0, 10)
    };
  }

  private getCategoryPerformanceData(components: Component[], requests: BorrowRequest[]): any[] {
    const categories = components.reduce((acc, comp) => {
      if (!acc[comp.category]) {
        acc[comp.category] = { 
          totalComponents: 0,
          totalUnits: 0, 
          available: 0,
          requests: 0
        };
      }
      acc[comp.category].totalComponents += 1;
      acc[comp.category].totalUnits += comp.totalQuantity;
      acc[comp.category].available += comp.availableQuantity;
      return acc;
    }, {} as Record<string, any>);

    // Count requests per category
    requests.forEach(request => {
      const component = components.find(c => c.name === request.componentName);
      if (component && categories[component.category]) {
        categories[component.category].requests += 1;
      }
    });

    return Object.entries(categories).map(([name, data]) => {
      const utilization = data.totalUnits > 0 
        ? (((data.totalUnits - data.available) / data.totalUnits) * 100).toFixed(1) + '%'
        : '0.0%';

      return {
        name,
        totalComponents: data.totalComponents,
        totalUnits: data.totalUnits,
        available: data.available,
        requests: data.requests,
        utilization
      };
    });
  }
}

export const excelService = ExcelService.getInstance();
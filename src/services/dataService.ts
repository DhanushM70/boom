import { SystemData, User, Component, BorrowRequest, Notification, LoginSession, SystemStats } from '../types';

class DataService {
  private storageKey = 'isaacLabData';

  private getDefaultData(): SystemData {
    return {
      users: [
        {
          id: 'admin-1',
          name: 'Administrator',
          email: 'admin@issacasimov.in',
          role: 'admin',
          registeredAt: new Date().toISOString(),
          loginCount: 0,
          isActive: false
        }
      ],
      components: [
        {
          id: 'comp-1',
          name: 'Arduino Uno R3',
          totalQuantity: 25,
          availableQuantity: 25,
          category: 'Microcontroller',
          description: 'Arduino Uno R3 development board'
        },
        {
          id: 'comp-2',
          name: 'L298N Motor Driver',
          totalQuantity: 15,
          availableQuantity: 15,
          category: 'Motor Driver',
          description: 'Dual H-Bridge Motor Driver'
        },
        {
          id: 'comp-3',
          name: 'Ultrasonic Sensor HC-SR04',
          totalQuantity: 20,
          availableQuantity: 20,
          category: 'Sensor',
          description: 'Ultrasonic distance sensor'
        },
        {
          id: 'comp-4',
          name: 'Servo Motor SG90',
          totalQuantity: 30,
          availableQuantity: 30,
          category: 'Actuator',
          description: '9g micro servo motor'
        },
        {
          id: 'comp-5',
          name: 'ESP32 Development Board',
          totalQuantity: 12,
          availableQuantity: 12,
          category: 'Microcontroller',
          description: 'WiFi and Bluetooth enabled microcontroller'
        }
      ],
      requests: [],
      notifications: [],
      loginSessions: []
    };
  }

  getData(): SystemData {
    try {
      const data = localStorage.getItem(this.storageKey);
      if (data) {
        const parsedData = JSON.parse(data);
        // Ensure loginSessions exists for backward compatibility
        if (!parsedData.loginSessions) {
          parsedData.loginSessions = [];
        }
        return parsedData;
      }
    } catch (error) {
      console.error('Error loading data:', error);
    }
    return this.getDefaultData();
  }

  saveData(data: SystemData): void {
    try {
      localStorage.setItem(this.storageKey, JSON.stringify(data));
    } catch (error) {
      console.error('Error saving data:', error);
    }
  }

  // User operations
  addUser(user: User): void {
    const data = this.getData();
    user.loginCount = 0;
    user.isActive = false;
    data.users.push(user);
    this.saveData(data);
  }

  updateUser(user: User): void {
    const data = this.getData();
    const index = data.users.findIndex(u => u.id === user.id);
    if (index !== -1) {
      data.users[index] = user;
      this.saveData(data);
    }
  }

  getUser(email: string): User | undefined {
    const data = this.getData();
    return data.users.find(user => user.email === email);
  }

  authenticateUser(email: string, password: string): User | null {
    const expectedPassword = email === 'admin@issacasimov.in' ? 'ralab' : 'issacasimov';
    
    if (password !== expectedPassword) {
      return null;
    }

    let user = this.getUser(email);
    
    if (!user && email.endsWith('@issacasimov.in') && email !== 'admin@issacasimov.in') {
      // Create new student user
      const name = email.split('@')[0].replace(/\./g, ' ').replace(/\b\w/g, l => l.toUpperCase());
      user = {
        id: `user-${Date.now()}`,
        name,
        email,
        role: 'student',
        registeredAt: new Date().toISOString(),
        loginCount: 0,
        isActive: true
      };
      this.addUser(user);
    }

    if (user) {
      // Update login statistics
      user.lastLoginAt = new Date().toISOString();
      user.loginCount = (user.loginCount || 0) + 1;
      user.isActive = true;
      this.updateUser(user);

      // Create login session
      this.createLoginSession(user);
    }

    return user || null;
  }

  // Login session management
  createLoginSession(user: User): LoginSession {
    const session: LoginSession = {
      id: `session-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
      userId: user.id,
      userEmail: user.email,
      userName: user.name,
      userRole: user.role,
      loginTime: new Date().toISOString(),
      ipAddress: this.getClientIP(),
      userAgent: navigator.userAgent,
      deviceInfo: this.getDeviceInfo(),
      isActive: true
    };

    const data = this.getData();
    data.loginSessions.push(session);
    this.saveData(data);

    return session;
  }

  endLoginSession(userId: string): void {
    const data = this.getData();
    const activeSessions = data.loginSessions.filter(s => s.userId === userId && s.isActive);
    
    activeSessions.forEach(session => {
      session.logoutTime = new Date().toISOString();
      session.isActive = false;
      session.sessionDuration = new Date().getTime() - new Date(session.loginTime).getTime();
    });

    // Update user active status
    const user = data.users.find(u => u.id === userId);
    if (user) {
      user.isActive = false;
    }

    this.saveData(data);
  }

  getLoginSessions(): LoginSession[] {
    return this.getData().loginSessions;
  }

  getActiveUsers(): User[] {
    return this.getData().users.filter(u => u.isActive);
  }

  private getClientIP(): string {
    // In a real application, you'd get this from your backend
    return 'Unknown';
  }

  private getDeviceInfo(): string {
    const ua = navigator.userAgent;
    if (ua.includes('Mobile')) return 'Mobile Device';
    if (ua.includes('Tablet')) return 'Tablet';
    return 'Desktop';
  }

  // Component operations
  getComponents(): Component[] {
    return this.getData().components;
  }

  updateComponent(component: Component): void {
    const data = this.getData();
    const index = data.components.findIndex(c => c.id === component.id);
    if (index !== -1) {
      data.components[index] = component;
      this.saveData(data);
    }
  }

  addComponent(component: Component): void {
    const data = this.getData();
    data.components.push(component);
    this.saveData(data);
  }

  // Request operations
  addRequest(request: BorrowRequest): void {
    const data = this.getData();
    data.requests.push(request);
    this.saveData(data);
  }

  updateRequest(request: BorrowRequest): void {
    const data = this.getData();
    const index = data.requests.findIndex(r => r.id === request.id);
    if (index !== -1) {
      data.requests[index] = request;
      this.saveData(data);
    }
  }

  getRequests(): BorrowRequest[] {
    return this.getData().requests;
  }

  getUserRequests(userId: string): BorrowRequest[] {
    return this.getData().requests.filter(r => r.studentId === userId);
  }

  // Notification operations
  addNotification(notification: Notification): void {
    const data = this.getData();
    data.notifications.push(notification);
    this.saveData(data);
  }

  getUserNotifications(userId: string): Notification[] {
    return this.getData().notifications.filter(n => n.userId === userId);
  }

  markNotificationAsRead(notificationId: string): void {
    const data = this.getData();
    const notification = data.notifications.find(n => n.id === notificationId);
    if (notification) {
      notification.read = true;
      this.saveData(data);
    }
  }

  // System statistics
  getSystemStats(): SystemStats {
    const data = this.getData();
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

  // Enhanced Export functionality
  exportToExcel(): void {
    const data = this.getData();
    const workbook = this.createWorkbook(data);
    this.downloadExcel(workbook, `Isaac-Asimov-Lab-Report-${new Date().toISOString().split('T')[0]}.xlsx`);
  }

  private createWorkbook(data: SystemData): any {
    // Create a comprehensive Excel workbook with multiple sheets
    const workbook = {
      SheetNames: ['Summary', 'Requests', 'Components', 'Users', 'Login Sessions'],
      Sheets: {}
    };

    // Summary Sheet
    workbook.Sheets['Summary'] = this.createSummarySheet(data);
    
    // Requests Sheet
    workbook.Sheets['Requests'] = this.createRequestsSheet(data.requests);
    
    // Components Sheet
    workbook.Sheets['Components'] = this.createComponentsSheet(data.components);
    
    // Users Sheet
    workbook.Sheets['Users'] = this.createUsersSheet(data.users);
    
    // Login Sessions Sheet
    workbook.Sheets['Login Sessions'] = this.createLoginSessionsSheet(data.loginSessions);

    return workbook;
  }

  private createSummarySheet(data: SystemData): any {
    const stats = this.getSystemStats();
    const now = new Date();
    
    const summaryData = [
      ['Isaac Asimov Robotics Lab - System Report'],
      ['Generated on:', now.toLocaleString()],
      [''],
      ['SYSTEM OVERVIEW'],
      ['Total Users', stats.totalUsers],
      ['Active Users', stats.activeUsers],
      ['Total Logins', stats.totalLogins],
      ['Currently Online', stats.onlineUsers],
      [''],
      ['COMPONENT MANAGEMENT'],
      ['Total Components', stats.totalComponents],
      ['Total Requests', stats.totalRequests],
      ['Pending Requests', stats.pendingRequests],
      ['Overdue Items', stats.overdueItems],
      [''],
      ['REQUEST STATUS BREAKDOWN'],
      ['Pending', data.requests.filter(r => r.status === 'pending').length],
      ['Approved', data.requests.filter(r => r.status === 'approved').length],
      ['Rejected', data.requests.filter(r => r.status === 'rejected').length],
      ['Returned', data.requests.filter(r => r.status === 'returned').length],
      [''],
      ['COMPONENT CATEGORIES'],
      ...this.getComponentCategorySummary(data.components)
    ];

    return this.arrayToSheet(summaryData);
  }

  private createRequestsSheet(requests: BorrowRequest[]): any {
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
      'Notes',
      'Days Overdue'
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
        request.notes || '',
        daysOverdue > 0 ? daysOverdue : ''
      ];
    });

    return this.arrayToSheet([headers, ...rows]);
  }

  private createComponentsSheet(components: Component[]): any {
    const headers = [
      'Component ID',
      'Name',
      'Category',
      'Total Quantity',
      'Available Quantity',
      'Borrowed Quantity',
      'Utilization %',
      'Description'
    ];

    const rows = components.map(component => {
      const borrowed = component.totalQuantity - component.availableQuantity;
      const utilization = component.totalQuantity > 0 
        ? ((borrowed / component.totalQuantity) * 100).toFixed(1)
        : '0';

      return [
        component.id,
        component.name,
        component.category,
        component.totalQuantity,
        component.availableQuantity,
        borrowed,
        `${utilization}%`,
        component.description || ''
      ];
    });

    return this.arrayToSheet([headers, ...rows]);
  }

  private createUsersSheet(users: User[]): any {
    const headers = [
      'User ID',
      'Name',
      'Email',
      'Role',
      'Registration Date',
      'Last Login',
      'Total Logins',
      'Currently Active'
    ];

    const rows = users.map(user => [
      user.id,
      user.name,
      user.email,
      user.role.toUpperCase(),
      new Date(user.registeredAt).toLocaleDateString(),
      user.lastLoginAt ? new Date(user.lastLoginAt).toLocaleDateString() : 'Never',
      user.loginCount || 0,
      user.isActive ? 'Yes' : 'No'
    ]);

    return this.arrayToSheet([headers, ...rows]);
  }

  private createLoginSessionsSheet(sessions: LoginSession[]): any {
    const headers = [
      'Session ID',
      'User Name',
      'Email',
      'Role',
      'Login Time',
      'Logout Time',
      'Duration (minutes)',
      'Device',
      'Status'
    ];

    const rows = sessions.map(session => {
      const duration = session.sessionDuration 
        ? Math.round(session.sessionDuration / 60000)
        : session.isActive 
          ? Math.round((new Date().getTime() - new Date(session.loginTime).getTime()) / 60000)
          : 0;

      return [
        session.id,
        session.userName,
        session.userEmail,
        session.userRole.toUpperCase(),
        new Date(session.loginTime).toLocaleString(),
        session.logoutTime ? new Date(session.logoutTime).toLocaleString() : 'Active',
        duration,
        session.deviceInfo || 'Unknown',
        session.isActive ? 'Active' : 'Ended'
      ];
    });

    return this.arrayToSheet([headers, ...rows]);
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

    return Object.entries(categories).map(([category, data]) => [
      category,
      `${data.available}/${data.total} available`
    ]);
  }

  private arrayToSheet(data: any[][]): any {
    const sheet: any = {};
    const range = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };

    for (let R = 0; R < data.length; R++) {
      for (let C = 0; C < data[R].length; C++) {
        if (range.s.r > R) range.s.r = R;
        if (range.s.c > C) range.s.c = C;
        if (range.e.r < R) range.e.r = R;
        if (range.e.c < C) range.e.c = C;

        const cell: any = { v: data[R][C] };
        
        if (cell.v == null) continue;
        
        const cell_ref = this.encodeCell({ c: C, r: R });
        
        if (typeof cell.v === 'number') cell.t = 'n';
        else if (typeof cell.v === 'boolean') cell.t = 'b';
        else if (cell.v instanceof Date) {
          cell.t = 'n';
          cell.z = 'mm/dd/yyyy';
          cell.v = this.dateToSerial(cell.v);
        } else cell.t = 's';

        sheet[cell_ref] = cell;
      }
    }
    
    if (range.s.c < 10000000) sheet['!ref'] = this.encodeRange(range);
    return sheet;
  }

  private encodeCell(cell: { c: number; r: number }): string {
    return this.encodeCol(cell.c) + this.encodeRow(cell.r);
  }

  private encodeCol(col: number): string {
    let s = '';
    for (++col; col; col = Math.floor((col - 1) / 26)) {
      s = String.fromCharCode(((col - 1) % 26) + 65) + s;
    }
    return s;
  }

  private encodeRow(row: number): string {
    return (row + 1).toString();
  }

  private encodeRange(range: any): string {
    return this.encodeCell(range.s) + ':' + this.encodeCell(range.e);
  }

  private dateToSerial(date: Date): number {
    return (date.getTime() - new Date(1900, 0, 1).getTime()) / (24 * 60 * 60 * 1000) + 1;
  }

  private downloadExcel(workbook: any, filename: string): void {
    const wbout = this.writeWorkbook(workbook);
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    const link = document.createElement('a');
    if (link.download !== undefined) {
      const url = URL.createObjectURL(blob);
      link.setAttribute('href', url);
      link.setAttribute('download', filename);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  }

  private writeWorkbook(workbook: any): ArrayBuffer {
    // Simple XLSX writer implementation
    const zip = this.createZip();
    
    // Add basic XLSX structure
    zip.addFile('[Content_Types].xml', this.createContentTypes());
    zip.addFile('_rels/.rels', this.createRels());
    zip.addFile('xl/_rels/workbook.xml.rels', this.createWorkbookRels(workbook));
    zip.addFile('xl/workbook.xml', this.createWorkbookXml(workbook));
    zip.addFile('xl/sharedStrings.xml', this.createSharedStrings(workbook));
    
    // Add worksheets
    workbook.SheetNames.forEach((name: string, index: number) => {
      zip.addFile(`xl/worksheets/sheet${index + 1}.xml`, this.createWorksheet(workbook.Sheets[name]));
    });

    return zip.generate();
  }

  private createZip(): any {
    // Simplified ZIP implementation for XLSX
    const files: Record<string, string> = {};
    
    return {
      addFile: (path: string, content: string) => {
        files[path] = content;
      },
      generate: (): ArrayBuffer => {
        // Convert to simple CSV for now since full XLSX is complex
        const csvContent = this.workbookToCSV(files);
        const encoder = new TextEncoder();
        return encoder.encode(csvContent).buffer;
      }
    };
  }

  private workbookToCSV(files: Record<string, string>): string {
    // Fallback to CSV format
    const data = this.getData();
    return this.exportToCSV();
  }

  private createContentTypes(): string {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>';
  }

  private createRels(): string {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  }

  private createWorkbookRels(workbook: any): string {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  }

  private createWorkbookXml(workbook: any): string {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></workbook>';
  }

  private createSharedStrings(workbook: any): string {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></sst>';
  }

  private createWorksheet(sheet: any): string {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>';
  }

  // Enhanced CSV Export
  exportToCSV(): string {
    const data = this.getData();
    const requests = data.requests;
    
    const headers = [
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
        new Date(request.requestDate).toLocaleDateString(),
        request.studentName,
        request.rollNo,
        request.mobile,
        request.componentName,
        request.quantity.toString(),
        new Date(request.dueDate).toLocaleDateString(),
        request.status.charAt(0).toUpperCase() + request.status.slice(1),
        request.approvedBy || '',
        request.approvedAt ? new Date(request.approvedAt).toLocaleDateString() : '',
        request.returnedAt ? new Date(request.returnedAt).toLocaleDateString() : '',
        daysOverdue > 0 ? daysOverdue.toString() : '',
        request.notes || ''
      ];
    });

    const csvContent = [headers, ...rows]
      .map(row => row.map(cell => `"${cell}"`).join(','))
      .join('\n');

    return csvContent;
  }

  exportLoginSessionsCSV(): string {
    const sessions = this.getLoginSessions();
    
    const headers = [
      'Login Time',
      'Logout Time',
      'User Name',
      'User Email',
      'Role',
      'Device Info',
      'Session Duration (minutes)',
      'Status'
    ];

    const rows = sessions.map(session => [
      new Date(session.loginTime).toLocaleString(),
      session.logoutTime ? new Date(session.logoutTime).toLocaleString() : 'Active',
      session.userName,
      session.userEmail,
      session.userRole,
      session.deviceInfo || 'Unknown',
      session.sessionDuration ? Math.round(session.sessionDuration / 60000).toString() : 'Active',
      session.isActive ? 'Active' : 'Ended'
    ]);

    const csvContent = [headers, ...rows]
      .map(row => row.map(cell => `"${cell}"`).join(','))
      .join('\n');

    return csvContent;
  }
}

export const dataService = new DataService();
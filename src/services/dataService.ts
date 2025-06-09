import { SystemData, User, Component, BorrowRequest, Notification } from '../types';

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
      notifications: []
    };
  }

  getData(): SystemData {
    try {
      const data = localStorage.getItem(this.storageKey);
      if (data) {
        return JSON.parse(data);
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
    data.users.push(user);
    this.saveData(data);
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
      };
      this.addUser(user);
    }

    return user || null;
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

  // Export functionality
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
      'Returned Date'
    ];

    const rows = requests.map(request => [
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
      request.returnedAt ? new Date(request.returnedAt).toLocaleDateString() : ''
    ]);

    const csvContent = [headers, ...rows]
      .map(row => row.map(cell => `"${cell}"`).join(','))
      .join('\n');

    return csvContent;
  }
}

export const dataService = new DataService();
export interface User {
  id: string;
  name: string;
  email: string;
  role: 'student' | 'admin';
  rollNo?: string;
  mobile?: string;
  registeredAt: string;
}

export interface Component {
  id: string;
  name: string;
  totalQuantity: number;
  availableQuantity: number;
  category: string;
  description?: string;
}

export interface BorrowRequest {
  id: string;
  studentId: string;
  studentName: string;
  rollNo: string;
  mobile: string;
  componentId: string;
  componentName: string;
  quantity: number;
  requestDate: string;
  dueDate: string;
  status: 'pending' | 'approved' | 'rejected' | 'returned';
  approvedBy?: string;
  approvedAt?: string;
  returnedAt?: string;
  notes?: string;
}

export interface Notification {
  id: string;
  userId: string;
  title: string;
  message: string;
  type: 'info' | 'success' | 'warning' | 'error';
  read: boolean;
  createdAt: string;
}

export interface SystemData {
  users: User[];
  components: Component[];
  requests: BorrowRequest[];
  notifications: Notification[];
}
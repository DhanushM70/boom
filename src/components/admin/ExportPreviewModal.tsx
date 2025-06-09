import React from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { 
  X, 
  Download, 
  FileSpreadsheet, 
  Users, 
  Package, 
  Clock, 
  TrendingUp,
  BarChart3,
  Activity,
  AlertTriangle,
  CheckCircle,
  Layers
} from 'lucide-react';
import { excelService } from '../../services/excelService';
import { dataService } from '../../services/dataService';

interface ExportPreviewModalProps {
  isOpen: boolean;
  onClose: () => void;
}

const ExportPreviewModal: React.FC<ExportPreviewModalProps> = ({ isOpen, onClose }) => {
  const [previewData, setPreviewData] = React.useState<any>(null);
  const [isLoading, setIsLoading] = React.useState(false);

  React.useEffect(() => {
    if (isOpen) {
      setIsLoading(true);
      const data = dataService.getData();
      const preview = excelService.generatePreviewData(data);
      setPreviewData(preview);
      setIsLoading(false);
    }
  }, [isOpen]);

  const handleDownloadExcel = () => {
    const data = dataService.getData();
    excelService.exportToExcel(data);
    onClose();
  };

  if (!isOpen) return null;

  return (
    <AnimatePresence>
      <motion.div
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        className="fixed inset-0 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4 z-50"
        onClick={(e) => e.target === e.currentTarget && onClose()}
      >
        <motion.div
          initial={{ scale: 0.9, opacity: 0, y: 20 }}
          animate={{ scale: 1, opacity: 1, y: 0 }}
          exit={{ scale: 0.9, opacity: 0, y: 20 }}
          className="bg-dark-800 rounded-3xl border border-peacock-500/20 w-full max-w-7xl max-h-[90vh] overflow-hidden shadow-2xl"
        >
          {/* Header */}
          <div className="bg-gradient-to-r from-peacock-500/10 to-blue-500/10 border-b border-peacock-500/20 p-6">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-4">
                <div className="p-3 bg-gradient-to-br from-peacock-500 to-blue-500 rounded-xl shadow-lg">
                  <FileSpreadsheet className="w-8 h-8 text-white" />
                </div>
                <div>
                  <h2 className="text-2xl font-bold text-white">Professional Excel Report Preview</h2>
                  <p className="text-peacock-300">Review comprehensive data analysis before downloading</p>
                </div>
              </div>
              <button
                onClick={onClose}
                className="p-2 text-peacock-400 hover:text-white hover:bg-dark-700/50 rounded-lg transition-all duration-200"
              >
                <X className="w-6 h-6" />
              </button>
            </div>
          </div>

          {/* Content */}
          <div className="p-6 overflow-y-auto max-h-[calc(90vh-200px)]">
            {isLoading ? (
              <div className="flex items-center justify-center py-12">
                <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-peacock-500"></div>
              </div>
            ) : previewData ? (
              <div className="space-y-8">
                {/* Executive Summary */}
                <div>
                  <h3 className="text-xl font-bold text-white mb-4 flex items-center gap-2">
                    <BarChart3 className="w-5 h-5 text-peacock-400" />
                    Executive Summary Dashboard
                  </h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                    {[
                      { 
                        label: 'Total Users', 
                        value: previewData.summary.totalUsers, 
                        icon: Users,
                        color: 'from-blue-500 to-cyan-500',
                        bgColor: 'bg-blue-500/10',
                        change: '+12%'
                      },
                      { 
                        label: 'Active Users', 
                        value: previewData.summary.activeUsers, 
                        icon: Activity,
                        color: 'from-green-500 to-emerald-500',
                        bgColor: 'bg-green-500/10',
                        change: '+8%'
                      },
                      { 
                        label: 'Total Requests', 
                        value: previewData.summary.totalRequests, 
                        icon: Package,
                        color: 'from-purple-500 to-pink-500',
                        bgColor: 'bg-purple-500/10',
                        change: '+25%'
                      },
                      { 
                        label: 'Pending Requests', 
                        value: previewData.summary.pendingRequests, 
                        icon: Clock,
                        color: 'from-yellow-500 to-orange-500',
                        bgColor: 'bg-yellow-500/10',
                        change: '-5%'
                      }
                    ].map((stat, index) => {
                      const Icon = stat.icon;
                      return (
                        <motion.div
                          key={stat.label}
                          initial={{ opacity: 0, y: 20 }}
                          animate={{ opacity: 1, y: 0 }}
                          transition={{ delay: index * 0.1 }}
                          className={`${stat.bgColor} backdrop-blur-xl rounded-xl border border-peacock-500/20 p-4`}
                        >
                          <div className="flex items-center gap-3 mb-2">
                            <div className={`p-2 bg-gradient-to-br ${stat.color} rounded-lg`}>
                              <Icon className="w-5 h-5 text-white" />
                            </div>
                            <div className="flex-1">
                              <p className="text-2xl font-bold text-white">{stat.value}</p>
                              <p className="text-peacock-300 text-sm">{stat.label}</p>
                            </div>
                          </div>
                          <div className="flex items-center gap-1 text-xs">
                            <TrendingUp className="w-3 h-3 text-green-400" />
                            <span className="text-green-400 font-medium">{stat.change}</span>
                            <span className="text-peacock-400">vs last month</span>
                          </div>
                        </motion.div>
                      );
                    })}
                  </div>
                </div>

                {/* Detailed Component Analysis */}
                <div>
                  <h3 className="text-xl font-bold text-white mb-4 flex items-center gap-2">
                    <Layers className="w-5 h-5 text-peacock-400" />
                    Detailed Component Analysis
                  </h3>
                  <div className="bg-dark-700/30 rounded-xl border border-peacock-500/20 overflow-hidden">
                    <div className="overflow-x-auto">
                      <table className="w-full">
                        <thead className="bg-dark-700/50">
                          <tr>
                            <th className="text-left p-4 text-peacock-300 font-medium">Component</th>
                            <th className="text-left p-4 text-peacock-300 font-medium">Category</th>
                            <th className="text-left p-4 text-peacock-300 font-medium">Total Stock</th>
                            <th className="text-left p-4 text-peacock-300 font-medium">Available</th>
                            <th className="text-left p-4 text-peacock-300 font-medium">In Use</th>
                            <th className="text-left p-4 text-peacock-300 font-medium">Utilization</th>
                            <th className="text-left p-4 text-peacock-300 font-medium">Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          {previewData.detailedComponents.slice(0, 8).map((component: any, index: number) => {
                            const utilizationPercent = parseFloat(component.utilization);
                            const getStatusIcon = () => {
                              if (component.available === 0) return <AlertTriangle className="w-4 h-4 text-red-400" />;
                              if (utilizationPercent > 80) return <AlertTriangle className="w-4 h-4 text-yellow-400" />;
                              return <CheckCircle className="w-4 h-4 text-green-400" />;
                            };
                            
                            const getStatusText = () => {
                              if (component.available === 0) return 'Out of Stock';
                              if (utilizationPercent > 80) return 'High Demand';
                              if (utilizationPercent > 50) return 'Moderate Use';
                              return 'Available';
                            };

                            const getStatusColor = () => {
                              if (component.available === 0) return 'text-red-400 bg-red-500/10';
                              if (utilizationPercent > 80) return 'text-yellow-400 bg-yellow-500/10';
                              if (utilizationPercent > 50) return 'text-blue-400 bg-blue-500/10';
                              return 'text-green-400 bg-green-500/10';
                            };

                            return (
                              <tr key={component.id} className="border-t border-dark-600 hover:bg-dark-700/20">
                                <td className="p-4">
                                  <div className="flex items-center gap-2">
                                    <Package className="w-4 h-4 text-peacock-400" />
                                    <span className="text-white font-medium">{component.name}</span>
                                  </div>
                                </td>
                                <td className="p-4 text-peacock-300">{component.category}</td>
                                <td className="p-4 text-white font-semibold">{component.totalQuantity}</td>
                                <td className="p-4 text-white">{component.available}</td>
                                <td className="p-4 text-white">{component.inUse}</td>
                                <td className="p-4">
                                  <div className="flex items-center gap-2">
                                    <div className="w-16 bg-dark-600 rounded-full h-2">
                                      <div 
                                        className="bg-gradient-to-r from-peacock-500 to-blue-500 h-2 rounded-full transition-all duration-300"
                                        style={{ width: `${Math.min(utilizationPercent, 100)}%` }}
                                      ></div>
                                    </div>
                                    <span className="text-white text-sm font-medium">{component.utilization}</span>
                                  </div>
                                </td>
                                <td className="p-4">
                                  <div className={`flex items-center gap-2 px-3 py-1 rounded-full text-xs font-medium ${getStatusColor()}`}>
                                    {getStatusIcon()}
                                    {getStatusText()}
                                  </div>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                    {previewData.detailedComponents.length > 8 && (
                      <div className="p-3 bg-dark-700/20 border-t border-dark-600 text-center">
                        <p className="text-peacock-300 text-sm">
                          +{previewData.detailedComponents.length - 8} more components will be included in the full export
                        </p>
                      </div>
                    )}
                  </div>
                </div>

                {/* Recent Requests Preview */}
                <div>
                  <h3 className="text-xl font-bold text-white mb-4 flex items-center gap-2">
                    <Package className="w-5 h-5 text-peacock-400" />
                    Recent Requests Analysis
                  </h3>
                  <div className="bg-dark-700/30 rounded-xl border border-peacock-500/20 overflow-hidden">
                    <div className="overflow-x-auto">
                      <table className="w-full">
                        <thead className="bg-dark-700/50">
                          <tr>
                            <th className="text-left p-3 text-peacock-300 font-medium">Student</th>
                            <th className="text-left p-3 text-peacock-300 font-medium">Component</th>
                            <th className="text-left p-3 text-peacock-300 font-medium">Quantity</th>
                            <th className="text-left p-3 text-peacock-300 font-medium">Status</th>
                            <th className="text-left p-3 text-peacock-300 font-medium">Date</th>
                            <th className="text-left p-3 text-peacock-300 font-medium">Priority</th>
                          </tr>
                        </thead>
                        <tbody>
                          {previewData.recentRequests.slice(0, 6).map((request: any, index: number) => {
                            const isOverdue = request.status === 'approved' && new Date(request.dueDate) < new Date();
                            const priority = isOverdue ? 'High' : request.status === 'pending' ? 'Medium' : 'Normal';
                            const priorityColor = isOverdue ? 'text-red-400 bg-red-500/10' : 
                                                request.status === 'pending' ? 'text-yellow-400 bg-yellow-500/10' : 
                                                'text-green-400 bg-green-500/10';

                            return (
                              <tr key={request.id} className="border-t border-dark-600">
                                <td className="p-3 text-white">{request.studentName}</td>
                                <td className="p-3 text-white">{request.componentName}</td>
                                <td className="p-3 text-white">{request.quantity}</td>
                                <td className="p-3">
                                  <span className={`px-2 py-1 rounded-full text-xs font-medium ${
                                    request.status === 'approved' ? 'bg-green-500/20 text-green-400' :
                                    request.status === 'rejected' ? 'bg-red-500/20 text-red-400' :
                                    request.status === 'returned' ? 'bg-blue-500/20 text-blue-400' :
                                    'bg-yellow-500/20 text-yellow-400'
                                  }`}>
                                    {request.status.toUpperCase()}
                                  </span>
                                </td>
                                <td className="p-3 text-peacock-300 text-sm">
                                  {new Date(request.requestDate).toLocaleDateString()}
                                </td>
                                <td className="p-3">
                                  <span className={`px-2 py-1 rounded-full text-xs font-medium ${priorityColor}`}>
                                    {priority}
                                  </span>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>

                {/* Category Performance */}
                <div>
                  <h3 className="text-xl font-bold text-white mb-4 flex items-center gap-2">
                    <TrendingUp className="w-5 h-5 text-peacock-400" />
                    Category Performance Analysis
                  </h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                    {previewData.categoryPerformance.map((category: any, index: number) => (
                      <motion.div
                        key={category.name}
                        initial={{ opacity: 0, x: -20 }}
                        animate={{ opacity: 1, x: 0 }}
                        transition={{ delay: index * 0.1 }}
                        className="bg-dark-700/30 rounded-xl border border-peacock-500/20 p-4"
                      >
                        <div className="flex items-center justify-between mb-3">
                          <h4 className="text-white font-semibold">{category.name}</h4>
                          <span className="text-peacock-400 font-bold text-lg">{category.utilization}</span>
                        </div>
                        <div className="space-y-2 text-sm">
                          <div className="flex justify-between">
                            <span className="text-peacock-300">Total Components:</span>
                            <span className="text-white font-medium">{category.totalComponents}</span>
                          </div>
                          <div className="flex justify-between">
                            <span className="text-peacock-300">Total Units:</span>
                            <span className="text-white font-medium">{category.totalUnits}</span>
                          </div>
                          <div className="flex justify-between">
                            <span className="text-peacock-300">Available:</span>
                            <span className="text-white font-medium">{category.available}</span>
                          </div>
                          <div className="flex justify-between">
                            <span className="text-peacock-300">Requests:</span>
                            <span className="text-white font-medium">{category.requests}</span>
                          </div>
                        </div>
                        <div className="mt-3 bg-dark-600 rounded-full h-2">
                          <div 
                            className="bg-gradient-to-r from-peacock-500 to-blue-500 h-2 rounded-full transition-all duration-300"
                            style={{ width: category.utilization }}
                          ></div>
                        </div>
                      </motion.div>
                    ))}
                  </div>
                </div>

                {/* Export Information */}
                <div className="bg-gradient-to-r from-peacock-500/10 to-blue-500/10 rounded-xl border border-peacock-500/20 p-6">
                  <h3 className="text-lg font-bold text-white mb-3">📊 Professional Excel Report Contents</h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6 text-sm">
                    <div>
                      <h4 className="text-peacock-400 font-semibold mb-3">📋 Report Sheets:</h4>
                      <ul className="text-peacock-300 space-y-2">
                        <li className="flex items-center gap-2">
                          <div className="w-2 h-2 bg-peacock-500 rounded-full"></div>
                          Executive Summary Dashboard
                        </li>
                        <li className="flex items-center gap-2">
                          <div className="w-2 h-2 bg-blue-500 rounded-full"></div>
                          Detailed Component Inventory
                        </li>
                        <li className="flex items-center gap-2">
                          <div className="w-2 h-2 bg-purple-500 rounded-full"></div>
                          Complete Request History
                        </li>
                        <li className="flex items-center gap-2">
                          <div className="w-2 h-2 bg-green-500 rounded-full"></div>
                          User Management & Analytics
                        </li>
                        <li className="flex items-center gap-2">
                          <div className="w-2 h-2 bg-orange-500 rounded-full"></div>
                          Login Sessions & Activity
                        </li>
                      </ul>
                    </div>
                    <div>
                      <h4 className="text-peacock-400 font-semibold mb-3">✨ Professional Features:</h4>
                      <ul className="text-peacock-300 space-y-2">
                        <li className="flex items-center gap-2">
                          <CheckCircle className="w-4 h-4 text-green-400" />
                          Corporate-grade formatting
                        </li>
                        <li className="flex items-center gap-2">
                          <CheckCircle className="w-4 h-4 text-green-400" />
                          Conditional formatting & charts
                        </li>
                        <li className="flex items-center gap-2">
                          <CheckCircle className="w-4 h-4 text-green-400" />
                          Statistical analysis & KPIs
                        </li>
                        <li className="flex items-center gap-2">
                          <CheckCircle className="w-4 h-4 text-green-400" />
                          Auto-sized columns & filters
                        </li>
                        <li className="flex items-center gap-2">
                          <CheckCircle className="w-4 h-4 text-green-400" />
                          Ready for executive presentations
                        </li>
                      </ul>
                    </div>
                  </div>
                </div>
              </div>
            ) : null}
          </div>

          {/* Footer */}
          <div className="border-t border-peacock-500/20 p-6 bg-dark-700/30">
            <div className="flex items-center justify-between">
              <div className="text-peacock-300 text-sm">
                <p className="font-semibold">Isaac Asimov Robotics Lab Management System</p>
                <p>Professional Excel Report • Generated on {new Date().toLocaleString()}</p>
              </div>
              <div className="flex gap-3">
                <motion.button
                  whileHover={{ scale: 1.02 }}
                  whileTap={{ scale: 0.98 }}
                  onClick={onClose}
                  className="px-6 py-3 bg-dark-600 text-white rounded-xl font-medium hover:bg-dark-500 transition-all duration-200"
                >
                  Cancel
                </motion.button>
                <motion.button
                  whileHover={{ scale: 1.02, boxShadow: '0 10px 30px rgba(0, 206, 209, 0.3)' }}
                  whileTap={{ scale: 0.98 }}
                  onClick={handleDownloadExcel}
                  className="group relative overflow-hidden bg-gradient-to-r from-green-500 to-emerald-500 text-white px-8 py-3 rounded-xl font-semibold shadow-lg hover:shadow-2xl transition-all duration-300"
                >
                  <div className="absolute inset-0 bg-gradient-to-r from-green-600 to-emerald-600 opacity-0 group-hover:opacity-100 transition-opacity duration-300"></div>
                  <div className="relative z-10 flex items-center gap-3">
                    <Download className="w-5 h-5 group-hover:animate-bounce" />
                    Download Professional Excel Report
                  </div>
                </motion.button>
              </div>
            </div>
          </div>
        </motion.div>
      </motion.div>
    </AnimatePresence>
  );
};

export default ExportPreviewModal;
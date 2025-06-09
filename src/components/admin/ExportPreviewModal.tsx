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
  Activity
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
          className="bg-dark-800 rounded-3xl border border-peacock-500/20 w-full max-w-6xl max-h-[90vh] overflow-hidden shadow-2xl"
        >
          {/* Header */}
          <div className="bg-gradient-to-r from-peacock-500/10 to-blue-500/10 border-b border-peacock-500/20 p-6">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-4">
                <div className="p-3 bg-gradient-to-br from-peacock-500 to-blue-500 rounded-xl shadow-lg">
                  <FileSpreadsheet className="w-8 h-8 text-white" />
                </div>
                <div>
                  <h2 className="text-2xl font-bold text-white">Export Preview</h2>
                  <p className="text-peacock-300">Review your data before downloading the Excel report</p>
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
                {/* Summary Cards */}
                <div>
                  <h3 className="text-xl font-bold text-white mb-4 flex items-center gap-2">
                    <BarChart3 className="w-5 h-5 text-peacock-400" />
                    System Overview
                  </h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                    {[
                      { 
                        label: 'Total Users', 
                        value: previewData.summary.totalUsers, 
                        icon: Users,
                        color: 'from-blue-500 to-cyan-500',
                        bgColor: 'bg-blue-500/10'
                      },
                      { 
                        label: 'Active Users', 
                        value: previewData.summary.activeUsers, 
                        icon: Activity,
                        color: 'from-green-500 to-emerald-500',
                        bgColor: 'bg-green-500/10'
                      },
                      { 
                        label: 'Total Requests', 
                        value: previewData.summary.totalRequests, 
                        icon: Package,
                        color: 'from-purple-500 to-pink-500',
                        bgColor: 'bg-purple-500/10'
                      },
                      { 
                        label: 'Pending Requests', 
                        value: previewData.summary.pendingRequests, 
                        icon: Clock,
                        color: 'from-yellow-500 to-orange-500',
                        bgColor: 'bg-yellow-500/10'
                      },
                      { 
                        label: 'Total Components', 
                        value: previewData.summary.totalComponents, 
                        icon: Package,
                        color: 'from-peacock-500 to-blue-500',
                        bgColor: 'bg-peacock-500/10'
                      },
                      { 
                        label: 'Overdue Items', 
                        value: previewData.summary.overdueItems, 
                        icon: TrendingUp,
                        color: 'from-red-500 to-red-600',
                        bgColor: 'bg-red-500/10'
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
                          <div className="flex items-center gap-3">
                            <div className={`p-2 bg-gradient-to-br ${stat.color} rounded-lg`}>
                              <Icon className="w-5 h-5 text-white" />
                            </div>
                            <div>
                              <p className="text-2xl font-bold text-white">{stat.value}</p>
                              <p className="text-peacock-300 text-sm">{stat.label}</p>
                            </div>
                          </div>
                        </motion.div>
                      );
                    })}
                  </div>
                </div>

                {/* Recent Requests Preview */}
                <div>
                  <h3 className="text-xl font-bold text-white mb-4 flex items-center gap-2">
                    <Package className="w-5 h-5 text-peacock-400" />
                    Recent Requests (Preview)
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
                          </tr>
                        </thead>
                        <tbody>
                          {previewData.recentRequests.slice(0, 5).map((request: any, index: number) => (
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
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    {previewData.recentRequests.length > 5 && (
                      <div className="p-3 bg-dark-700/20 border-t border-dark-600 text-center">
                        <p className="text-peacock-300 text-sm">
                          +{previewData.recentRequests.length - 5} more requests will be included in the full export
                        </p>
                      </div>
                    )}
                  </div>
                </div>

                {/* Component Utilization */}
                <div>
                  <h3 className="text-xl font-bold text-white mb-4 flex items-center gap-2">
                    <TrendingUp className="w-5 h-5 text-peacock-400" />
                    Component Utilization by Category
                  </h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {previewData.componentUtilization.map((category: any, index: number) => (
                      <motion.div
                        key={category[0]}
                        initial={{ opacity: 0, x: -20 }}
                        animate={{ opacity: 1, x: 0 }}
                        transition={{ delay: index * 0.1 }}
                        className="bg-dark-700/30 rounded-xl border border-peacock-500/20 p-4"
                      >
                        <div className="flex items-center justify-between mb-2">
                          <h4 className="text-white font-semibold">{category[0]}</h4>
                          <span className="text-peacock-400 font-bold">{category[3]}</span>
                        </div>
                        <div className="flex justify-between text-sm">
                          <span className="text-peacock-300">Total: {category[1]}</span>
                          <span className="text-peacock-300">Available: {category[2]}</span>
                        </div>
                        <div className="mt-2 bg-dark-600 rounded-full h-2">
                          <div 
                            className="bg-gradient-to-r from-peacock-500 to-blue-500 h-2 rounded-full transition-all duration-300"
                            style={{ width: category[3] }}
                          ></div>
                        </div>
                      </motion.div>
                    ))}
                  </div>
                </div>

                {/* Export Information */}
                <div className="bg-gradient-to-r from-peacock-500/10 to-blue-500/10 rounded-xl border border-peacock-500/20 p-6">
                  <h3 className="text-lg font-bold text-white mb-3">📊 What's Included in Your Excel Export</h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-sm">
                    <div>
                      <h4 className="text-peacock-400 font-semibold mb-2">📋 Sheets Included:</h4>
                      <ul className="text-peacock-300 space-y-1">
                        <li>• Summary Dashboard</li>
                        <li>• Complete Requests History</li>
                        <li>• Component Inventory</li>
                        <li>• User Management</li>
                        <li>• Login Sessions</li>
                      </ul>
                    </div>
                    <div>
                      <h4 className="text-peacock-400 font-semibold mb-2">✨ Features:</h4>
                      <ul className="text-peacock-300 space-y-1">
                        <li>• Professional formatting</li>
                        <li>• Color-coded headers</li>
                        <li>• Conditional formatting</li>
                        <li>• Auto-sized columns</li>
                        <li>• Statistical analysis</li>
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
                <p>Report generated on {new Date().toLocaleString()}</p>
                <p>Isaac Asimov Robotics Lab Management System</p>
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
                    Download Excel Report
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
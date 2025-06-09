import React, { useState, useEffect } from 'react';
import { motion } from 'framer-motion';
import { 
  Users, 
  UserCheck, 
  Activity, 
  Clock, 
  Download,
  Smartphone,
  Monitor,
  Tablet,
  TrendingUp,
  Calendar
} from 'lucide-react';
import { dataService } from '../../services/dataService';
import { LoginSession, SystemStats } from '../../types';

const UserAnalytics: React.FC = () => {
  const [stats, setStats] = useState<SystemStats>({
    totalUsers: 0,
    activeUsers: 0,
    totalLogins: 0,
    onlineUsers: 0,
    totalRequests: 0,
    pendingRequests: 0,
    totalComponents: 0,
    overdueItems: 0
  });
  const [loginSessions, setLoginSessions] = useState<LoginSession[]>([]);
  const [timeFilter, setTimeFilter] = useState<'today' | 'week' | 'month' | 'all'>('today');

  useEffect(() => {
    loadData();
    const interval = setInterval(loadData, 30000); // Update every 30 seconds
    return () => clearInterval(interval);
  }, []);

  const loadData = () => {
    setStats(dataService.getSystemStats());
    setLoginSessions(dataService.getLoginSessions());
  };

  const exportLoginData = () => {
    const csvContent = dataService.exportLoginSessionsCSV();
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    
    if (link.download !== undefined) {
      const url = URL.createObjectURL(blob);
      link.setAttribute('href', url);
      link.setAttribute('download', `login-sessions-${new Date().toISOString().split('T')[0]}.csv`);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  const getFilteredSessions = () => {
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    const monthAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);

    return loginSessions.filter(session => {
      const loginDate = new Date(session.loginTime);
      switch (timeFilter) {
        case 'today':
          return loginDate >= today;
        case 'week':
          return loginDate >= weekAgo;
        case 'month':
          return loginDate >= monthAgo;
        default:
          return true;
      }
    }).sort((a, b) => new Date(b.loginTime).getTime() - new Date(a.loginTime).getTime());
  };

  const getDeviceStats = () => {
    const sessions = getFilteredSessions();
    const deviceCounts = sessions.reduce((acc, session) => {
      const device = session.deviceInfo || 'Unknown';
      acc[device] = (acc[device] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    return deviceCounts;
  };

  const getDeviceIcon = (device: string) => {
    if (device.includes('Mobile')) return Smartphone;
    if (device.includes('Tablet')) return Tablet;
    return Monitor;
  };

  const filteredSessions = getFilteredSessions();
  const deviceStats = getDeviceStats();

  return (
    <div className="space-y-8">
      {/* Header */}
      <motion.div
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        className="bg-gradient-to-r from-blue-500/10 to-purple-500/10 backdrop-blur-xl rounded-2xl border border-blue-500/20 p-6"
      >
        <div className="flex items-center gap-4">
          <div className="p-3 bg-gradient-to-br from-blue-500 to-purple-500 rounded-xl shadow-lg">
            <Activity className="w-8 h-8 text-white" />
          </div>
          <div>
            <h2 className="text-2xl font-bold text-white mb-1">User Analytics</h2>
            <p className="text-blue-200">Monitor user activity and login patterns</p>
          </div>
        </div>
      </motion.div>

      {/* Stats Cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
        {[
          { 
            title: 'Total Users', 
            value: stats.totalUsers, 
            icon: Users, 
            color: 'from-blue-500 to-cyan-500',
            bgColor: 'bg-blue-500/10',
            borderColor: 'border-blue-500/30'
          },
          { 
            title: 'Active Users', 
            value: stats.activeUsers, 
            icon: UserCheck, 
            color: 'from-green-500 to-emerald-500',
            bgColor: 'bg-green-500/10',
            borderColor: 'border-green-500/30'
          },
          { 
            title: 'Total Logins', 
            value: stats.totalLogins, 
            icon: TrendingUp, 
            color: 'from-purple-500 to-pink-500',
            bgColor: 'bg-purple-500/10',
            borderColor: 'border-purple-500/30'
          },
          { 
            title: 'Online Now', 
            value: stats.onlineUsers, 
            icon: Activity, 
            color: 'from-orange-500 to-red-500',
            bgColor: 'bg-orange-500/10',
            borderColor: 'border-orange-500/30'
          }
        ].map((stat, index) => {
          const Icon = stat.icon;
          return (
            <motion.div
              key={stat.title}
              initial={{ opacity: 0, y: 20, scale: 0.9 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              transition={{ delay: index * 0.1 }}
              whileHover={{ scale: 1.02, y: -5 }}
              className={`relative overflow-hidden ${stat.bgColor} backdrop-blur-xl rounded-2xl border ${stat.borderColor} p-6 group hover:shadow-2xl transition-all duration-300`}
            >
              <div className="absolute inset-0 bg-gradient-to-br from-white/5 to-transparent opacity-0 group-hover:opacity-100 transition-opacity duration-300"></div>
              <div className="relative z-10">
                <div className="flex items-center justify-between mb-4">
                  <div className={`p-3 bg-gradient-to-br ${stat.color} rounded-xl shadow-lg group-hover:animate-pulse`}>
                    <Icon className="w-6 h-6 text-white" />
                  </div>
                  <motion.div
                    initial={{ scale: 0 }}
                    animate={{ scale: 1 }}
                    transition={{ delay: index * 0.1 + 0.3 }}
                    className="text-right"
                  >
                    <p className="text-3xl font-bold text-white">{stat.value}</p>
                  </motion.div>
                </div>
                <h3 className="text-peacock-200 font-medium">{stat.title}</h3>
              </div>
            </motion.div>
          );
        })}
      </div>

      {/* Device Statistics */}
      <motion.div
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ delay: 0.4 }}
        className="bg-dark-800/50 backdrop-blur-xl rounded-2xl border border-peacock-500/20 p-6"
      >
        <h3 className="text-xl font-bold text-white mb-4">Device Usage</h3>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          {Object.entries(deviceStats).map(([device, count], index) => {
            const DeviceIcon = getDeviceIcon(device);
            return (
              <motion.div
                key={device}
                initial={{ opacity: 0, scale: 0.9 }}
                animate={{ opacity: 1, scale: 1 }}
                transition={{ delay: index * 0.1 }}
                className="flex items-center gap-3 p-4 bg-dark-700/30 rounded-xl border border-dark-600"
              >
                <DeviceIcon className="w-6 h-6 text-peacock-400" />
                <div>
                  <p className="text-white font-semibold">{count}</p>
                  <p className="text-peacock-300 text-sm">{device}</p>
                </div>
              </motion.div>
            );
          })}
        </div>
      </motion.div>

      {/* Time Filter and Export */}
      <div className="flex flex-col sm:flex-row gap-4 justify-between items-center">
        <div className="flex gap-2 bg-dark-800/30 p-2 rounded-xl">
          {[
            { key: 'today', label: 'Today' },
            { key: 'week', label: 'This Week' },
            { key: 'month', label: 'This Month' },
            { key: 'all', label: 'All Time' }
          ].map((filter) => (
            <motion.button
              key={filter.key}
              whileHover={{ scale: 1.02 }}
              whileTap={{ scale: 0.98 }}
              onClick={() => setTimeFilter(filter.key as any)}
              className={`px-4 py-2 rounded-lg font-medium transition-all duration-200 ${
                timeFilter === filter.key
                  ? 'bg-peacock-500 text-white shadow-lg'
                  : 'text-peacock-300 hover:text-white hover:bg-dark-700/50'
              }`}
            >
              {filter.label}
            </motion.button>
          ))}
        </div>

        <motion.button
          whileHover={{ scale: 1.05 }}
          whileTap={{ scale: 0.95 }}
          onClick={exportLoginData}
          className="flex items-center gap-2 bg-gradient-to-r from-peacock-500 to-blue-500 text-white px-6 py-3 rounded-xl font-semibold shadow-lg hover:shadow-xl transition-all duration-300"
        >
          <Download className="w-5 h-5" />
          Export Login Data
        </motion.button>
      </div>

      {/* Login Sessions Table */}
      <motion.div
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ delay: 0.6 }}
        className="bg-dark-800/50 backdrop-blur-xl rounded-2xl border border-peacock-500/20 overflow-hidden"
      >
        <div className="p-6 border-b border-dark-700">
          <h3 className="text-xl font-bold text-white">Recent Login Sessions</h3>
          <p className="text-peacock-300 text-sm mt-1">
            Showing {filteredSessions.length} sessions for {timeFilter === 'all' ? 'all time' : timeFilter}
          </p>
        </div>
        
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead className="bg-dark-700/50">
              <tr>
                <th className="text-left p-4 text-peacock-300 font-medium">User</th>
                <th className="text-left p-4 text-peacock-300 font-medium">Login Time</th>
                <th className="text-left p-4 text-peacock-300 font-medium">Device</th>
                <th className="text-left p-4 text-peacock-300 font-medium">Duration</th>
                <th className="text-left p-4 text-peacock-300 font-medium">Status</th>
              </tr>
            </thead>
            <tbody>
              {filteredSessions.slice(0, 20).map((session, index) => {
                const DeviceIcon = getDeviceIcon(session.deviceInfo || 'Unknown');
                const duration = session.sessionDuration 
                  ? Math.round(session.sessionDuration / 60000) 
                  : Math.round((new Date().getTime() - new Date(session.loginTime).getTime()) / 60000);
                
                return (
                  <motion.tr
                    key={session.id}
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: index * 0.03 }}
                    className="border-t border-dark-700 hover:bg-dark-700/30 transition-colors duration-200"
                  >
                    <td className="p-4">
                      <div>
                        <p className="text-white font-medium">{session.userName}</p>
                        <p className="text-peacock-300 text-sm">{session.userEmail}</p>
                        <span className={`inline-block px-2 py-1 rounded-full text-xs font-medium ${
                          session.userRole === 'admin' 
                            ? 'bg-purple-500/20 text-purple-400' 
                            : 'bg-blue-500/20 text-blue-400'
                        }`}>
                          {session.userRole}
                        </span>
                      </div>
                    </td>
                    <td className="p-4">
                      <div className="flex items-center gap-2">
                        <Calendar className="w-4 h-4 text-peacock-400" />
                        <div>
                          <p className="text-white text-sm">{new Date(session.loginTime).toLocaleDateString()}</p>
                          <p className="text-peacock-300 text-xs">{new Date(session.loginTime).toLocaleTimeString()}</p>
                        </div>
                      </div>
                    </td>
                    <td className="p-4">
                      <div className="flex items-center gap-2">
                        <DeviceIcon className="w-4 h-4 text-peacock-400" />
                        <span className="text-white text-sm">{session.deviceInfo || 'Unknown'}</span>
                      </div>
                    </td>
                    <td className="p-4">
                      <div className="flex items-center gap-2">
                        <Clock className="w-4 h-4 text-peacock-400" />
                        <span className="text-white text-sm">{duration} min</span>
                      </div>
                    </td>
                    <td className="p-4">
                      <span className={`px-3 py-1 rounded-full text-xs font-medium ${
                        session.isActive 
                          ? 'bg-green-500/20 text-green-400 border border-green-500/30' 
                          : 'bg-gray-500/20 text-gray-400 border border-gray-500/30'
                      }`}>
                        {session.isActive ? 'Active' : 'Ended'}
                      </span>
                    </td>
                  </motion.tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </motion.div>

      {filteredSessions.length === 0 && (
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="text-center py-12"
        >
          <Activity className="w-16 h-16 text-peacock-400 mx-auto mb-4" />
          <h3 className="text-xl font-semibold text-white mb-2">No Login Sessions</h3>
          <p className="text-peacock-300">No login sessions found for the selected time period.</p>
        </motion.div>
      )}
    </div>
  );
};

export default UserAnalytics;
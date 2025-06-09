import React from 'react';
import { motion } from 'framer-motion';
import { LogOut, Cpu } from 'lucide-react';
import { useAuth } from '../context/AuthContext';
import NotificationBell from './NotificationBell';

const Header: React.FC = () => {
  const { user, logout } = useAuth();

  if (!user) return null;

  return (
    <motion.header
      initial={{ y: -20, opacity: 0 }}
      animate={{ y: 0, opacity: 1 }}
      className="bg-dark-800/90 backdrop-blur-lg border-b border-dark-700 sticky top-0 z-40"
    >
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex items-center justify-between h-16">
          <div className="flex items-center gap-3">
            <div className="flex items-center justify-center w-10 h-10 bg-gradient-to-r from-peacock-500 to-blue-500 rounded-lg">
              <Cpu className="w-6 h-6 text-white" />
            </div>
            <div>
              <h1 className="text-white font-bold text-lg">Isaac Asimov Lab</h1>
              <p className="text-peacock-300 text-sm">
                {user.role === 'admin' ? 'Admin Dashboard' : 'Student Portal'}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-4">
            <NotificationBell />
            
            <div className="flex items-center gap-3">
              <div className="text-right">
                <p className="text-white font-medium">{user.name}</p>
                <p className="text-peacock-300 text-sm">{user.email}</p>
              </div>
              
              <motion.button
                whileHover={{ scale: 1.05 }}
                whileTap={{ scale: 0.95 }}
                onClick={logout}
                className="p-2 text-peacock-300 hover:text-peacock-200 hover:bg-dark-700 rounded-lg transition-all duration-200"
                title="Logout"
              >
                <LogOut className="w-5 h-5" />
              </motion.button>
            </div>
          </div>
        </div>
      </div>
    </motion.header>
  );
};

export default Header;
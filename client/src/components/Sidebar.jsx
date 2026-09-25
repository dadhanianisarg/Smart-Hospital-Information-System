import React from 'react';
import { 
  LayoutDashboard, 
  Users, 
  UserCheck, 
  Building2, 
  CalendarCheck, 
  FileText, 
  Share2, 
  Database 
} from 'lucide-react';

export default function Sidebar({ activePage, setActivePage }) {
  const menuItems = [
    { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
    { id: 'patients', label: 'Patients', icon: Users },
    { id: 'doctors', label: 'Doctors', icon: UserCheck },
    { id: 'departments', label: 'Departments', icon: Building2 },
    { id: 'appointments', label: 'Appointments', icon: CalendarCheck },
    { id: 'prescriptions', label: 'Prescriptions', icon: FileText },
    { id: 'graph-explorer', label: 'Graph Explorer', icon: Share2 },
  ];

  return (
    <aside className="sidebar">
      <div className="sidebar-header">
        <div className="sidebar-title">
          <Database size={22} />
          <span>SmartHospital</span>
        </div>
        <div className="sidebar-subtitle">
          MongoDB + Neo4j NoSQL
        </div>
      </div>

      <nav className="sidebar-nav">
        {menuItems.map((item) => {
          const Icon = item.icon;
          const isActive = activePage === item.id;
          return (
            <button
              key={item.id}
              className={`nav-item ${isActive ? 'active' : ''}`}
              onClick={() => setActivePage(item.id)}
            >
              <Icon size={18} />
              <span>{item.label}</span>
            </button>
          );
        })}
      </nav>

      <div className="sidebar-footer">
        <div style={{ fontWeight: 600, color: '#94a3b8', marginBottom: '4px' }}>Team Members:</div>
        <div>• Dadhania Nisarg (23BCE2364)</div>
        <div>• Madhav Sah (23BCE0868)</div>
        <div>• Arnav Dewan (23BCE0351)</div>
        <div style={{ marginTop: '4px', color: '#38bdf8' }}>BCSE406L – Review 2</div>
      </div>
    </aside>
  );
}

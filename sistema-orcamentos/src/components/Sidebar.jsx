import { useState } from "react";
import { Link } from "react-router-dom";
import { FiHome, FiUsers, FiBox, FiFileText, FiSearch, FiSettings, FiChevronLeft, FiChevronRight } from "react-icons/fi";

export default function Sidebar() {
  const [open, setOpen] = useState(true);

  const menus = [
    { name: "Dashboard", icon: <FiHome />, path: "/" },
    { name: "Clientes", icon: <FiUsers />, path: "/clientes" },
    { name: "Produtos", icon: <FiBox />, path: "/produtos" },
    { name: "Orçamentos", icon: <FiFileText />, path: "/orcamentos" },
    { name: "Consulta", icon: <FiSearch />, path: "/consulta" },
    { name: "Configurações", icon: <FiSettings />, path: "/config" },
  ];

  return (
    <aside className={`bg-gray-900 text-white h-screen p-4 pt-6 relative transition-all ${open ? "w-60" : "w-20"}`}>
      {/* Botão toggle */}
      <button 
        className="absolute -right-3 top-9 bg-blue-600 rounded-full p-1"
        onClick={() => setOpen(!open)}
      >
        {open ? <FiChevronLeft /> : <FiChevronRight />}
      </button>

      <div className="flex flex-col gap-6 mt-8">
        {menus.map((menu, i) => (
          <Link key={i} to={menu.path} className="flex items-center gap-3 hover:bg-gray-700 p-2 rounded">
            <span className="text-lg">{menu.icon}</span>
            {open && <span>{menu.name}</span>}
          </Link>
        ))}
      </div>
    </aside>
  );
}

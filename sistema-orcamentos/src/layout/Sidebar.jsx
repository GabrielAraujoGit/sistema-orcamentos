import { Link } from "react-router-dom";
import { FiHome, FiUsers, FiBox, FiFileText, FiSearch, FiSettings } from "react-icons/fi";

export default function Sidebar() {
  return (
    <aside className="w-60 bg-gray-900 text-white min-h-screen p-4">
      <ul className="space-y-3">
        <li>
          <Link to="/" className="flex items-center gap-2 hover:underline">
            <FiHome /> Dashboard
          </Link>
        </li>
        <li>
          <Link to="/clientes" className="flex items-center gap-2 hover:underline">
            <FiUsers /> Clientes
          </Link>
        </li>
        <li>
          <Link to="/produtos" className="flex items-center gap-2 hover:underline">
            <FiBox /> Produtos
          </Link>
        </li>
        <li>
          <Link to="/orcamentos" className="flex items-center gap-2 hover:underline">
            <FiFileText /> Orçamentos
          </Link>
        </li>
        <li>
          <Link to="/consulta" className="flex items-center gap-2 hover:underline">
            <FiSearch /> Consultar
          </Link>
        </li>
        <li>
          <Link to="/config" className="flex items-center gap-2 hover:underline">
            <FiSettings /> Configurações
          </Link>
        </li>
      </ul>
    </aside>
  );
}

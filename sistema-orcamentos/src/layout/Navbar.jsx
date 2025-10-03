import { FiSearch, FiUser } from "react-icons/fi";

export default function Navbar() {
  return (
    <header className="bg-white h-14 shadow flex items-center justify-between px-6">
      <h1 className="text-lg font-bold text-gray-700">💼 Sistema ERP</h1>
      
      <div className="flex items-center gap-4">
        {/* Campo de pesquisa */}
        <div className="relative">
          <FiSearch className="absolute left-2 top-2 text-gray-500" />
          <input
            type="text"
            placeholder="Pesquisar..."
            className="border rounded pl-8 pr-3 py-1 focus:outline-blue-500"
          />
        </div>

        {/* Usuário */}
        <div className="flex items-center gap-2">
          <FiUser className="text-xl" />
          <span className="text-gray-700">Admin</span>
        </div>
      </div>
    </header>
  );
}

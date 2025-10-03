import { useState } from "react";
import { FiPlus, FiEdit2, FiTrash2 } from "react-icons/fi";
import FormCliente from "../components/FormCliente";

export default function Clientes() {
  const [clientes, setClientes] = useState([
    { id: 1, nome: "Empresa X", email: "contato@empresa.com", telefone: "(11) 99999-9999" },
    { id: 2, nome: "Cliente Y", email: "cliente@y.com", telefone: "(21) 98888-8888" },
  ]);

  const [showModal, setShowModal] = useState(false);
  const [editingCliente, setEditingCliente] = useState(null);

  const handleAdd = () => {
    setEditingCliente(null);
    setShowModal(true);
  };

  const handleEdit = (cliente) => {
    setEditingCliente(cliente);
    setShowModal(true);
  };

  const handleDelete = (id) => {
    setClientes(clientes.filter((c) => c.id !== id));
  };

  const handleSave = (cliente) => {
    if (cliente.id) {
      // edição
      setClientes(clientes.map((c) => (c.id === cliente.id ? cliente : c)));
    } else {
      // novo cadastro
      setClientes([...clientes, { ...cliente, id: Date.now() }]);
    }
    setShowModal(false);
  };

  return (
    <div className="max-w-6xl mx-auto">
      <div className="flex justify-between items-center mb-6">
        <h2 className="text-2xl font-bold text-gray-700">👤 Clientes</h2>
        <button
          onClick={handleAdd}
          className="flex items-center gap-2 bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded"
        >
          <FiPlus /> Novo Cliente
        </button>
      </div>

      {/* Tabela */}
      <div className="bg-white shadow rounded-xl overflow-hidden">
        <table className="w-full border-collapse">
          <thead className="bg-gray-100 text-gray-600">
            <tr>
              <th className="p-3 text-left">Nome</th>
              <th className="p-3 text-left">Email</th>
              <th className="p-3 text-left">Telefone</th>
              <th className="p-3 text-center">Ações</th>
            </tr>
          </thead>
          <tbody>
            {clientes.map((c) => (
              <tr key={c.id} className="border-t hover:bg-gray-50">
                <td className="p-3">{c.nome}</td>
                <td className="p-3">{c.email}</td>
                <td className="p-3">{c.telefone}</td>
                <td className="p-3 flex justify-center gap-3">
                  <button
                    onClick={() => handleEdit(c)}
                    className="text-blue-600 hover:text-blue-800"
                  >
                    <FiEdit2 />
                  </button>
                  <button
                    onClick={() => handleDelete(c.id)}
                    className="text-red-600 hover:text-red-800"
                  >
                    <FiTrash2 />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Modal de cadastro/edição */}
      {showModal && (
        <FormCliente
          cliente={editingCliente}
          onClose={() => setShowModal(false)}
          onSave={handleSave}
        />
      )}
    </div>
  );
}

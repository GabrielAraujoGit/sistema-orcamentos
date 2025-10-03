import { FiUsers, FiBox, FiFileText, FiDollarSign } from "react-icons/fi";
import { PieChart, Pie, Cell, Tooltip, Legend } from "recharts";

const data = [
  { name: "Em Aberto", value: 12 },
  { name: "Aprovado", value: 18 },
  { name: "Rejeitado", value: 4 },
  { name: "Cancelado", value: 2 },
];

const COLORS = ["#2563eb", "#16a34a", "#dc2626", "#f97316"];

export default function Dashboard() {
  return (
    <div className="max-w-7xl mx-auto space-y-10">
      {/* Cabeçalho */}
      <header>
        <h1 className="text-3xl font-bold text-gray-800">Dashboard</h1>
        <p className="text-gray-500">Visão geral do sistema de orçamentos</p>
      </header>

      {/* Cards de Resumo (grid centralizado) */}
      <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6">
        <CardResumo icon={<FiUsers size={28} />} title="Clientes" value="120" color="blue" />
        <CardResumo icon={<FiBox size={28} />} title="Produtos" value="80" color="green" />
        <CardResumo icon={<FiFileText size={28} />} title="Orçamentos" value="45" color="yellow" />
        <CardResumo icon={<FiDollarSign size={28} />} title="Receita" value="R$ 120k" color="purple" />
      </section>

      {/* Indicadores (lado a lado) */}
      <section className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Gráfico */}
        <div className="bg-white shadow rounded-xl p-6 flex justify-center">
          <div>
            <h3 className="text-lg font-semibold text-gray-700 mb-4 text-center">
              Status dos Orçamentos
            </h3>
            <PieChart width={380} height={250}>
              <Pie data={data} dataKey="value" outerRadius={100} label>
                {data.map((_, i) => (
                  <Cell key={i} fill={COLORS[i % COLORS.length]} />
                ))}
              </Pie>
              <Tooltip />
              <Legend />
            </PieChart>
          </div>
        </div>

        {/* Últimos Orçamentos */}
        <div className="bg-white shadow rounded-xl p-6">
          <h3 className="text-lg font-semibold text-gray-700 mb-4">
            Últimos Orçamentos
          </h3>
          <ul className="divide-y divide-gray-200">
            <li className="py-3 flex justify-between">
              <span>Empresa X</span>
              <span className="text-blue-600 font-medium">Em Aberto</span>
            </li>
            <li className="py-3 flex justify-between">
              <span>Cliente Y</span>
              <span className="text-green-600 font-medium">Aprovado</span>
            </li>
            <li className="py-3 flex justify-between">
              <span>Empresa Z</span>
              <span className="text-red-600 font-medium">Rejeitado</span>
            </li>
          </ul>
        </div>
      </section>

      {/* Tabela de movimentações (rodapé do dashboard) */}
      <section className="bg-white shadow rounded-xl p-6">
        <h3 className="text-lg font-semibold text-gray-700 mb-4">
          Movimentações Recentes
        </h3>
        <table className="w-full border-collapse">
          <thead className="bg-gray-100 text-gray-600">
            <tr>
              <th className="p-3 text-left">#</th>
              <th className="p-3 text-left">Cliente</th>
              <th className="p-3 text-left">Valor</th>
              <th className="p-3 text-left">Status</th>
            </tr>
          </thead>
          <tbody>
            <tr className="border-t hover:bg-gray-50">
              <td className="p-3">1</td>
              <td className="p-3">Empresa XPTO</td>
              <td className="p-3">R$ 5.000</td>
              <td className="p-3 text-blue-600 font-medium">Em Aberto</td>
            </tr>
            <tr className="border-t hover:bg-gray-50">
              <td className="p-3">2</td>
              <td className="p-3">Cliente Y</td>
              <td className="p-3">R$ 3.200</td>
              <td className="p-3 text-green-600 font-medium">Aprovado</td>
            </tr>
            <tr className="border-t hover:bg-gray-50">
              <td className="p-3">3</td>
              <td className="p-3">Empresa Z</td>
              <td className="p-3">R$ 1.800</td>
              <td className="p-3 text-red-600 font-medium">Rejeitado</td>
            </tr>
          </tbody>
        </table>
      </section>
    </div>
  );
}

/* 🔹 Componente para os cards de resumo */
function CardResumo({ icon, title, value, color }) {
  const colorClasses = {
    blue: "bg-blue-100 text-blue-600",
    green: "bg-green-100 text-green-600",
    yellow: "bg-yellow-100 text-yellow-600",
    purple: "bg-purple-100 text-purple-600",
  };

  return (
    <div className="bg-white shadow rounded-xl p-6 flex items-center gap-4">
      <div className={`${colorClasses[color]} p-4 rounded-lg`}>{icon}</div>
      <div>
        <p className="text-gray-500">{title}</p>
        <p className="text-2xl font-bold">{value}</p>
      </div>
    </div>
  );
}

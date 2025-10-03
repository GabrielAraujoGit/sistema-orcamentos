export default function Dashboard() {
  return (
    <div>
      <h2 className="text-2xl font-bold mb-6">Dashboard</h2>
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <div className="bg-white shadow rounded p-4">
          <h3 className="text-gray-500">Clientes</h3>
          <p className="text-2xl font-bold">120</p>
        </div>
        <div className="bg-white shadow rounded p-4">
          <h3 className="text-gray-500">Produtos</h3>
          <p className="text-2xl font-bold">80</p>
        </div>
        <div className="bg-white shadow rounded p-4">
          <h3 className="text-gray-500">Orçamentos</h3>
          <p className="text-2xl font-bold">45</p>
        </div>
      </div>
    </div>
  );
}

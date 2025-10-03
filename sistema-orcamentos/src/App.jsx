import { BrowserRouter as Router, Routes, Route } from "react-router-dom";
import Layout from "./layout/Layout";

import Dashboard from "./pages/Dashboard";
import Clientes from "./pages/Clientes";
import Produtos from "./pages/Produtos";
import Orcamentos from "./pages/Orcamentos";
import ConsultaOrcamentos from "./pages/ConsultaOrcamentos";

export default function App() {
  return (
    <Router>
      <Layout>
        <Routes>
          <Route path="/" element={<Dashboard />} />
          <Route path="/clientes" element={<Clientes />} />
          <Route path="/produtos" element={<Produtos />} />
          <Route path="/orcamentos" element={<Orcamentos />} />
          <Route path="/consulta" element={<ConsultaOrcamentos />} />
          <Route path="*" element={<h2>Página não encontrada</h2>} />
        </Routes>
      </Layout>
    </Router>
  );
}

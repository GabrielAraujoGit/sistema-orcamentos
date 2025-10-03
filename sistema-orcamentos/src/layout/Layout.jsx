import Navbar from "./Navbar";
import Sidebar from "./Sidebar";

export default function Layout({ children }) {
  return (
    <div className="flex h-screen">
      <Sidebar />
      <div className="flex flex-col flex-1">
        <Navbar />
            <main className="p-6 bg-gray-100 flex-1 overflow-auto">
                <div className="max-w-7xl mx-auto">{children}</div> 
        </main>
      </div>
    </div>
  );
}

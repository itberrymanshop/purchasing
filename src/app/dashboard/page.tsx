import Link from "next/link";

const rows = [
  {
    sku: "GE0008",
    name: "Timbangan Buah Digital 40KG",
    stock: 245,
    sales: 182,
    coverage: 41,
    status: "AMAN",
    kind: "Pareto MP & GR",
  },
  {
    sku: "HC0210",
    name: "Cetakan Press Kulit Pangsit 2 In 1",
    stock: 62,
    sales: 128,
    coverage: 14,
    status: "ORDER",
    kind: "Pareto Marketplace",
  },
  {
    sku: "HG0077",
    name: "Rak Sepatu Sandal 4 Susun",
    stock: 318,
    sales: 74,
    coverage: 96,
    status: "MONITOR",
    kind: "Pareto MP & GR",
  },
  {
    sku: "100024",
    name: "Kipas Angin Listrik Karakter SX-168",
    stock: 0,
    sales: 0,
    coverage: 0,
    status: "DISCONTINUE",
    kind: "Aktif",
  },
];

export default function Dashboard() {
  const metrics = [
    ["Rp 1,24 M", "Total omzet"],
    ["59,7%", "Kontribusi PM"],
    ["45,6%", "Kontribusi impor"],
    ["24", "Perlu order"],
  ];

  return (
    <main className="shell">
      <header className="topbar">
        <div>
          <span className="eyebrow">LAPORAN AKTIF</span>
          <h1>Dashboard purchasing</h1>
        </div>
        <nav>
          <Link href="/">Upload</Link>
          <Link href="/dashboard">Dashboard</Link>
          <Link href="/audit">Audit</Link>
        </nav>
      </header>
      <div className="update">Update terakhir: 08 September 2026 · laporan berhasil diproses</div>
      <section className="metrics">
        {metrics.map(([value, label]) => (
          <div className="metric" key={label}>
            <strong>{value}</strong>
            <span>{label}</span>
          </div>
        ))}
      </section>
      <section className="panel">
        <div className="panel-head">
          <div>
            <span className="eyebrow">INVENTORY OVERVIEW</span>
            <h2>Detail stok dan penjualan</h2>
          </div>
          <select defaultValue="all" aria-label="Filter status barang">
            <option value="all">Semua status</option>
            <option>ORDER</option>
            <option>AMAN</option>
            <option>MONITOR</option>
            <option>DISCONTINUE</option>
          </select>
        </div>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>SKU</th>
                <th>Nama barang</th>
                <th>Stok dapat dijual</th>
                <th>Penjualan bulan lalu</th>
                <th>Coverage</th>
                <th>Status</th>
                <th>Pareto</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => (
                <tr key={row.sku}>
                  <td className="code">{row.sku}</td>
                  <td>{row.name}</td>
                  <td>{row.stock.toLocaleString("id-ID")}</td>
                  <td>{row.sales.toLocaleString("id-ID")}</td>
                  <td>{row.coverage} hari</td>
                  <td>
                    <span className={`status ${row.status.toLowerCase()}`}>{row.status}</span>
                  </td>
                  <td>{row.kind}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </main>
  );
}

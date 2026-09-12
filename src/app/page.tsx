import Link from "next/link";
import { UploadForm } from "@/components/upload-form";

export default function Home() {
  return (
    <main className="shell">
      <header className="topbar"><div><span className="eyebrow">PURCHASING CONTROL</span><h1>Analisa Stok Pembelian</h1></div><nav><Link href="/dashboard">Dashboard</Link><Link href="/audit">Audit</Link></nav></header>
      <section className="hero"><div><span className="pill">Laporan harian</span><h2>Upload satu template, dapatkan laporan matang.</h2><p>Template berisi 8 sheet sumber. Sistem validasi SKU, hitung penjualan, coverage, proyeksi, dan status order.</p><a className="download-link" href="/templates/template-purchasing.xlsx" download>Download template Excel</a></div><div className="hero-stat"><strong>8</strong><span>sheet sumber</span></div></section>
      <UploadForm />
    </main>
  );
}

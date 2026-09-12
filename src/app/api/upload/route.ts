import { NextResponse } from "next/server";
import * as XLSX from "xlsx";

const groups: Record<string, string[]> = {
  stok: ["stok", "ketersediaan"],
  penjualan: ["penjualan"],
  harga: ["harga"],
  discontinue: ["discontinue"],
  impor: ["impor", "masuk"],
  ppn: ["ppn", "pajak"],
  pareto: ["pareto"],
};
export async function POST(request: Request) {
  const form = await request.formData();
  const file = form.get("file");
  if (!(file instanceof File))
    return NextResponse.json(
      { message: "File template wajib dipilih." },
      { status: 400 },
    );
  if (file.size > 25 * 1024 * 1024)
    return NextResponse.json(
      { message: "Ukuran file melebihi 25 MB." },
      { status: 400 },
    );
  try {
    const buffer = Buffer.from(await file.arrayBuffer());
    const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
    const names = workbook.SheetNames.map((x) => x.toLowerCase());
    const missing = Object.entries(groups)
      .filter(
        ([, aliases]) =>
          !names.some((name) => aliases.some((alias) => name.includes(alias))),
      )
      .map(([key]) => key);
    if (missing.length)
      return NextResponse.json(
        {
          message: `Sheet belum lengkap: ${missing.join(", ")}.`,
          sheets: workbook.SheetNames,
        },
        { status: 422 },
      );
    const counts = Object.fromEntries(
      workbook.SheetNames.map((name) => [
        name,
        XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1 }).length,
      ]),
    );
    return NextResponse.json({
      message: `Validasi berhasil. ${workbook.SheetNames.length} sheet siap diproses.`,
      sheets: counts,
    });
  } catch {
    return NextResponse.json(
      { message: "File Excel tidak dapat dibaca." },
      { status: 422 },
    );
  }
}

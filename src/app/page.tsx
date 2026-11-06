"use client";

import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

const APPS_SCRIPT_URL =
  "https://script.google.com/macros/s/AKfycbzD7pWIe451V7d05xxqk0P7Fg3d-WYdrMv1owJxZZDP7CiR33EyULSknQVTxLbBbKyq5A/exec";

interface KesItem {
  kategori: string;
  kumpulan: string;
  kodKes: string;
  jenisKes: string;
}

export default function Home() {
  const [loading, setLoading] = useState(true);
  const [dataKes, setDataKes] = useState<KesItem[]>([]);
  const [senaraiNegeri, setSenaraiNegeri] = useState<string[]>([]);
  const [senaraiStatus, setSenaraiStatus] = useState<string[]>([]);

  // Tetapan Borang
  const [tajuk, setTajuk] = useState("STATISTIK KES MAHKAMAH SYARIAH");
  const [tahunMula, setTahunMula] = useState(2021);
  const [tahunAkhir, setTahunAkhir] = useState(2025);
  const [formatKes, setFormatKes] = useState<"kedua" | "kod" | "jenis">("kedua");
  
  // Pilihan Struktur Sumbu Data
  const [susunanBaris, setSusunanBaris] = useState<"negeri_kes" | "kes_negeri" | "kes_sahaja">("negeri_kes");

  const [selectedKes, setSelectedKes] = useState<string[]>([]);
  const [selectedNegeri, setSelectedNegeri] = useState<string[]>([]);
  const [selectedStatus, setSelectedStatus] = useState<string[]>([]);
  const [carianKes, setCarianKes] = useState("");

  useEffect(() => {
    fetch(APPS_SCRIPT_URL)
      .then((res) => res.json())
      .then((data) => {
        setDataKes(data.kes || []);
        setSenaraiNegeri(data.negeri || []);
        setSenaraiStatus(data.status || []);
        setSelectedNegeri(data.negeri || []);
        setSelectedStatus(data.status || []);
        setLoading(false);
      })
      .catch((err) => {
        console.error("Gagal membaca Google Sheet:", err);
        setLoading(false);
      });
  }, []);

  const handlePilihSemuaKes = () => {
    if (selectedKes.length === dataKes.length) {
      setSelectedKes([]);
    } else {
      setSelectedKes(dataKes.map((k) => k.kodKes));
    }
  };

  const handleToggleKes = (kod: string) => {
    setSelectedKes((prev) =>
      prev.includes(kod) ? prev.filter((k) => k !== kod) : [...prev, kod]
    );
  };

  const getLabelKes = (item: KesItem) => {
    if (formatKes === "kod") return item.kodKes;
    if (formatKes === "jenis") return item.jenisKes;
    return `${item.kodKes} - ${item.jenisKes}`;
  };

  const generateExcel = () => {
    const wb = XLSX.utils.book_new();
    const wsData: any[][] = [];
    const merges: XLSX.Range[] = [];

    const tahunArr: number[] = [];
    for (let t = tahunMula; t <= tahunAkhir; t++) {
      tahunArr.push(t);
    }

    const negeriTerpilih = senaraiNegeri.filter((n) => selectedNegeri.includes(n));
    const statusTerpilih = senaraiStatus.filter((s) => selectedStatus.includes(s));
    const kesTerpilih = dataKes.filter((k) => selectedKes.includes(k.kodKes));

    // Tajuk Dokumen
    wsData.push([tajuk.toUpperCase()]);
    wsData.push([`TEMPOH / JULAT TAHUN: ${tahunMula} - ${tahunAkhir}`]);
    wsData.push([]); // Baris pemisah

    const startRowIdx = wsData.length; // Baris header bermula

    // --- STRUKTUR FORMAT GAMBAR 4: NEGERI -> KES DI BARIS KIRI ---
    if (susunanBaris === "negeri_kes") {
      const headerRow1 = ["Negeri", "Jenis Kes"];
      const headerRow2 = ["", ""];

      let colPtr = 2;
      tahunArr.forEach((thn) => {
        headerRow1.push(thn.toString());
        // Isi ruang kosong untuk merge tahun
        for (let i = 1; i < statusTerpilih.length; i++) {
          headerRow1.push("");
        }
        // Rekod gabungan sel horizontal untuk tahun
        if (statusTerpilih.length > 1) {
          merges.push({
            s: { r: startRowIdx, c: colPtr },
            e: { r: startRowIdx, c: colPtr + statusTerpilih.length - 1 },
          });
        }
        statusTerpilih.forEach((st) => {
          headerRow2.push(st);
        });
        colPtr += statusTerpilih.length;
      });

      // Merge vertikal untuk Header 'Negeri' & 'Jenis Kes'
      merges.push({ s: { r: startRowIdx, c: 0 }, e: { r: startRowIdx + 1, c: 0 } });
      merges.push({ s: { r: startRowIdx, c: 1 }, e: { r: startRowIdx + 1, c: 1 } });

      wsData.push(headerRow1);
      wsData.push(headerRow2);

      // Isi Data Mengikut Blok Negeri
      let currentRow = wsData.length;
      negeriTerpilih.forEach((neg) => {
        const startNegeriRow = currentRow;

        kesTerpilih.forEach((kes) => {
          const rowData = ["", getLabelKes(kes)];
          for (let i = 2; i < headerRow1.length; i++) rowData.push("");
          wsData.push(rowData);
          currentRow++;
        });

        // Baris Jumlah per Negeri
        const jumlahRow = ["", "Jumlah"];
        for (let i = 2; i < headerRow1.length; i++) jumlahRow.push("");
        wsData.push(jumlahRow);
        currentRow++;

        // Set nama negeri di sel baris pertama dan gabungkan sel negeri ke bawah
        wsData[startNegeriRow][0] = neg;
        merges.push({
          s: { r: startNegeriRow, c: 0 },
          e: { r: currentRow - 1, c: 0 },
        });
      });
    }

    // --- STRUKTUR: KES DI KIRI (SATU KOLOM KES SAHAJA) ---
    else if (susunanBaris === "kes_sahaja") {
      const headerRow1 = ["Kod & Jenis Kes"];
      const headerRow2 = [""];

      let colPtr = 1;
      negeriTerpilih.forEach((neg) => {
        headerRow1.push(neg);
        for (let i = 1; i < tahunArr.length * statusTerpilih.length; i++) {
          headerRow1.push("");
        }
        merges.push({
          s: { r: startRowIdx, c: colPtr },
          e: { r: startRowIdx, c: colPtr + (tahunArr.length * statusTerpilih.length) - 1 },
        });

        tahunArr.forEach((thn) => {
          statusTerpilih.forEach((st) => {
            headerRow2.push(`${thn} (${st})`);
          });
        });
        colPtr += tahunArr.length * statusTerpilih.length;
      });

      merges.push({ s: { r: startRowIdx, c: 0 }, e: { r: startRowIdx + 1, c: 0 } });
      wsData.push(headerRow1);
      wsData.push(headerRow2);

      kesTerpilih.forEach((kes) => {
        const row = [getLabelKes(kes)];
        for (let i = 1; i < headerRow1.length; i++) row.push("");
        wsData.push(row);
      });
    }

    // --- STRUKTUR: KES -> NEGERI ---
    else {
      const headerRow1 = ["Jenis Kes", "Negeri"];
      const headerRow2 = ["", ""];

      let colPtr = 2;
      tahunArr.forEach((thn) => {
        headerRow1.push(thn.toString());
        for (let i = 1; i < statusTerpilih.length; i++) headerRow1.push("");
        if (statusTerpilih.length > 1) {
          merges.push({
            s: { r: startRowIdx, c: colPtr },
            e: { r: startRowIdx, c: colPtr + statusTerpilih.length - 1 },
          });
        }
        statusTerpilih.forEach((st) => headerRow2.push(st));
        colPtr += statusTerpilih.length;
      });

      merges.push({ s: { r: startRowIdx, c: 0 }, e: { r: startRowIdx + 1, c: 0 } });
      merges.push({ s: { r: startRowIdx, c: 1 }, e: { r: startRowIdx + 1, c: 1 } });
      wsData.push(headerRow1);
      wsData.push(headerRow2);

      let currentRow = wsData.length;
      kesTerpilih.forEach((kes) => {
        const startKesRow = currentRow;
        negeriTerpilih.forEach((neg) => {
          const rowData = ["", neg];
          for (let i = 2; i < headerRow1.length; i++) rowData.push("");
          wsData.push(rowData);
          currentRow++;
        });

        wsData[startKesRow][0] = getLabelKes(kes);
        merges.push({
          s: { r: startKesRow, c: 0 },
          e: { r: currentRow - 1, c: 0 },
        });
      });
    }

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws["!merges"] = merges;

    // Laraskan Lebar Kolom Secara Automatik (Auto Width)
    const colWidths = wsData[startRowIdx]?.map((_, colIdx) => {
      let maxLen = 12;
      wsData.forEach((row) => {
        const cell = row[colIdx];
        if (cell) {
          const str = cell.toString();
          if (str.length > maxLen) maxLen = Math.min(str.length + 3, 48);
        }
      });
      return { wch: maxLen };
    });
    ws["!cols"] = colWidths;

    XLSX.utils.book_append_sheet(wb, ws, "Templat Statistik");

    const excelBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([excelBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    saveAs(blob, `Template_Statistik_${Date.now()}.xlsx`);
  };

  const kesDitapis = dataKes.filter(
    (k) =>
      k.kodKes.toLowerCase().includes(carianKes.toLowerCase()) ||
      k.jenisKes.toLowerCase().includes(carianKes.toLowerCase())
  );

  if (loading) {
    return (
      <div className="flex h-screen items-center justify-center font-sans">
        <p className="text-base font-medium text-slate-600">
          Menyambung ke Google Sheets...
        </p>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 p-6 md:p-10 font-sans text-slate-800">
      <div className="max-w-4xl mx-auto bg-white p-6 md:p-8 rounded-xl shadow border border-slate-200">
        <h1 className="text-2xl font-bold mb-6 text-slate-900 border-b pb-3">
          Janaan Templat Statistik Kes
        </h1>

        {/* Bahagian 1: Parameter Asas */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
          <div className="md:col-span-2">
            <label className="block text-sm font-semibold mb-1">Tajuk Templat</label>
            <input
              type="text"
              className="w-full border rounded-lg p-2.5 outline-none focus:ring-2 focus:ring-blue-500"
              value={tajuk}
              onChange={(e) => setTajuk(e.target.value)}
            />
          </div>

          <div>
            <label className="block text-sm font-semibold mb-1">Tahun Mula</label>
            <input
              type="number"
              className="w-full border rounded-lg p-2 outline-none focus:ring-2 focus:ring-blue-500"
              value={tahunMula}
              onChange={(e) => setTahunMula(Number(e.target.value))}
            />
          </div>

          <div>
            <label className="block text-sm font-semibold mb-1">Tahun Akhir</label>
            <input
              type="number"
              className="w-full border rounded-lg p-2 outline-none focus:ring-2 focus:ring-blue-500"
              value={tahunAkhir}
              onChange={(e) => setTahunAkhir(Number(e.target.value))}
            />
          </div>

          <div>
            <label className="block text-sm font-semibold mb-1">Format Teks Kes</label>
            <select
              className="w-full border rounded-lg p-2 bg-white outline-none focus:ring-2 focus:ring-blue-500"
              value={formatKes}
              onChange={(e) => setFormatKes(e.target.value as any)}
            >
              <option value="kedua">Kedua-duanya (Contoh: 251 - Nafkah Anak)</option>
              <option value="kod">Kod Sahaja (Contoh: 251)</option>
              <option value="jenis">Jenis Kes Sahaja</option>
            </select>
          </div>

          <div>
            <label className="block text-sm font-semibold mb-1">
              Susunan Matriks Data (Baris Menegak)
            </label>
            <select
              className="w-full border rounded-lg p-2 bg-white outline-none focus:ring-2 focus:ring-blue-500"
              value={susunanBaris}
              onChange={(e) => setSusunanBaris(e.target.value as any)}
            >
              <option value="negeri_kes">Negeri ➔ Jenis Kes (Format Tahun ke Kanan)</option>
              <option value="kes_negeri">Jenis Kes ➔ Negeri</option>
              <option value="kes_sahaja">Jenis Kes Sahaja (Negeri di bahagian Atas)</option>
            </select>
          </div>
        </div>

        {/* Bahagian 2: Pilihan Negeri & Status */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
          <div className="border rounded-lg p-4 bg-slate-50">
            <h2 className="text-sm font-semibold mb-2">Pilihan Negeri ({selectedNegeri.length})</h2>
            <div className="max-h-36 overflow-y-auto space-y-1 text-sm">
              {senaraiNegeri.map((neg) => (
                <label key={neg} className="flex items-center space-x-2">
                  <input
                    type="checkbox"
                    checked={selectedNegeri.includes(neg)}
                    onChange={() =>
                      setSelectedNegeri((prev) =>
                        prev.includes(neg) ? prev.filter((n) => n !== neg) : [...prev, neg]
                      )
                    }
                  />
                  <span>{neg}</span>
                </label>
              ))}
            </div>
          </div>

          <div className="border rounded-lg p-4 bg-slate-50">
            <h2 className="text-sm font-semibold mb-2">Pilihan Status ({selectedStatus.length})</h2>
            <div className="max-h-36 overflow-y-auto space-y-1 text-sm">
              {senaraiStatus.map((st) => (
                <label key={st} className="flex items-center space-x-2">
                  <input
                    type="checkbox"
                    checked={selectedStatus.includes(st)}
                    onChange={() =>
                      setSelectedStatus((prev) =>
                        prev.includes(st) ? prev.filter((s) => s !== st) : [...prev, st]
                      )
                    }
                  />
                  <span>{st}</span>
                </label>
              ))}
            </div>
          </div>
        </div>

        {/* Bahagian 3: Pilihan Kod Kes */}
        <div className="mb-6">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-3 gap-2">
            <h2 className="text-sm font-semibold">
              Senarai Kes (Dipilih: {selectedKes.length} / {dataKes.length})
            </h2>
            <div className="flex gap-2 w-full sm:w-auto">
              <input
                type="text"
                placeholder="Cari kod atau jenis kes..."
                className="border rounded-lg px-3 py-1.5 text-sm w-full sm:w-60 outline-none focus:ring-2 focus:ring-blue-500"
                value={carianKes}
                onChange={(e) => setCarianKes(e.target.value)}
              />
              <button
                type="button"
                onClick={handlePilihSemuaKes}
                className="text-xs bg-slate-200 hover:bg-slate-300 font-medium px-3 py-1.5 rounded-lg whitespace-nowrap"
              >
                {selectedKes.length === dataKes.length ? "Nyahpilih" : "Pilih Semua"}
              </button>
            </div>
          </div>

          <div className="border rounded-lg max-h-60 overflow-y-auto p-2 divide-y divide-slate-100 bg-white">
            {kesDitapis.map((item) => (
              <label
                key={item.kodKes}
                className="flex items-center justify-between p-2 hover:bg-slate-50 rounded cursor-pointer text-sm"
              >
                <div className="flex items-center space-x-3">
                  <input
                    type="checkbox"
                    checked={selectedKes.includes(item.kodKes)}
                    onChange={() => handleToggleKes(item.kodKes)}
                  />
                  <span className="font-semibold text-blue-600">[{item.kodKes}]</span>
                  <span className="text-slate-700">{item.jenisKes}</span>
                </div>
                <span className="text-xs text-slate-400 bg-slate-100 px-2 py-0.5 rounded">
                  {item.kategori}
                </span>
              </label>
            ))}
          </div>
        </div>

        {/* Butang Tindakan */}
        <button
          onClick={generateExcel}
          disabled={selectedKes.length === 0}
          className={`w-full py-3 rounded-lg text-white font-semibold transition ${
            selectedKes.length === 0
              ? "bg-slate-400 cursor-not-allowed"
              : "bg-blue-600 hover:bg-blue-700 shadow"
          }`}
        >
          Muat Turun Templat Excel (.xlsx)
        </button>
      </div>
    </div>
  );
}
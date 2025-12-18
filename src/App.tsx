import React, { useState, useEffect, useRef } from "react";
import { BrowserRouter as Router, Route, Routes, Link } from "react-router-dom";
import jsPDF from "jspdf"; // Tambahkan impor ini
import autoTable from "jspdf-autotable";
import SignatureCanvas from "react-signature-canvas";

// Extend jsPDF type untuk mendukung lastAutoTable dari plugin autotable
declare module "jspdf" {
  interface jsPDF {
    lastAutoTable: {
      finalY: number;
    };
  }
}

interface RowData {
  [key: string]: string;
}

interface SheetInfo {
  sheetName: string;
  mapel: string;
  semester: string;
  kelas: string;
}

interface SchoolData {
  namaSekolah: string;
  npsn: string;
  alamatSekolah: string;
  kodePos: string;
  desaKelurahan: string;
  kabKota: string;
  provinsi: string;
  tahunPelajaran: string; // ✅ TAMBAHAN BARU
  tanggalRapor: string; // ✅ TAMBAHAN BARU
  namaKepsek: string;
  nipKepsek: string;
  ttdKepsek: string;
  namaGuru: string;
  nipGuru: string;
  ttdGuru: string;
}

interface KehadiranData {
  [key: string]: string;
}

const endpoint =
  "https://script.google.com/macros/s/AKfycbyNbHB6GUV9IYyf9GRz791b3AyHonok9gPXIGHzQu8WgNysFPIL3qwo4uUs1NmjSiSb/exec";

const throttle = (func: Function, delay: number) => {
  let timeoutId: ReturnType<typeof setTimeout> | null = null;
  let lastRan: number = 0;

  return function (this: any, ...args: any[]) {
    const now = Date.now();

    if (now - lastRan >= delay) {
      func.apply(this, args);
      lastRan = now;
    } else {
      if (timeoutId) clearTimeout(timeoutId);
      timeoutId = setTimeout(() => {
        func.apply(this, args);
        lastRan = Date.now();
      }, delay - (now - lastRan));
    }
  };
};

// Context untuk pre-loading data rekap
interface RekapContextType {
  rekapData: RekapData[];
  availableSheets: SheetInfo[];
  schoolData: SchoolData | null;
  kehadiranData: KehadiranData[];
  loading: boolean;
  error: string | null;
  refreshRekapData: (silent?: boolean) => Promise<void>;
}

const RekapContext = React.createContext<RekapContextType | null>(null);

const RekapProvider: React.FC<{ children: React.ReactNode }> = ({
  children,
}) => {
  const [rekapData, setRekapData] = useState<RekapData[]>([]);
  const [availableSheets, setAvailableSheets] = useState<SheetInfo[]>([]);
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);
  const [kehadiranData, setKehadiranData] = useState<KehadiranData[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const getCatatanByRanking = (ranking: number): string => {
    if (ranking === 1) {
      return "Pertahankan prestasi ananda!";
    } else if (ranking >= 2 && ranking <= 5) {
      return "Sudah baik, namun tingkatkan lagi prestrasi ananda!";
    } else if (ranking >= 6 && ranking <= 10) {
      return "Fokus, rajin dan lebih semangat lagi!";
    } else if (ranking >= 11 && ranking <= 15) {
      return "Tingkatkan semangat ananda sewaktu belajar!";
    } else if (ranking >= 16) {
      return "Lebih rajin lagi mengulang pelajaran di rumah ya!";
    }
    return "";
  };

  const refreshRekapData = async (silent: boolean = false) => {
    if (!silent) {
      setLoading(true);
    }
    setError(null);
    try {
      // Fetch school data
      const schoolResponse = await fetch(`${endpoint}?action=schoolData`);
      if (schoolResponse.ok) {
        const schoolJson = await schoolResponse.json();
        if (
          schoolJson.success &&
          schoolJson.data &&
          schoolJson.data.length > 0
        ) {
          setSchoolData(schoolJson.data[0]);
        }
      }

      // Fetch kehadiran data
      const kehadiranResponse = await fetch(`${endpoint}?sheet=DataKehadiran`);
      if (kehadiranResponse.ok) {
        const kehadiranJson = await kehadiranResponse.json();
        setKehadiranData(kehadiranJson.slice(1));
      }

      // Fetch list sheets
      const sheetsResponse = await fetch(`${endpoint}?action=listSheets`);
      if (!sheetsResponse.ok) throw new Error("Failed to fetch sheet list");
      const sheets: SheetInfo[] = await sheetsResponse.json();
      setAvailableSheets(sheets);

      // Fetch data dari setiap sheet
      const allDataPromises = sheets.map(async (sheet) => {
        const response = await fetch(`${endpoint}?sheet=${sheet.sheetName}`);
        if (!response.ok) throw new Error(`Failed to fetch ${sheet.sheetName}`);
        const jsonData = await response.json();
        return { mapel: sheet.mapel, data: jsonData.slice(1) };
      });

      const allData = await Promise.all(allDataPromises);

      // Gabungkan data per siswa
      const siswaMap: { [nama: string]: RekapData } = {};
      allData.forEach(({ mapel, data }) => {
        data.forEach((row: any) => {
          const nama = row.Data4;
          const kelas = row.Data3;
          const nilai = parseFloat(row.Data24) || null;

          if (!nama || nama.trim() === "") {
            return;
          }

          if (!siswaMap[nama]) {
            siswaMap[nama] = {
              nama,
              kelas,
              nilaiMapel: {},
              jumlah: 0,
              rataRata: 0,
              ranking: 0,
              catatan: "",
            };
          }
          siswaMap[nama].nilaiMapel[mapel] = nilai;
        });
      });

      // Hitung jumlah dan rata-rata
      const siswaArray = Object.keys(siswaMap).map((key) => siswaMap[key]);
      const rekapArray = siswaArray.map((siswa: RekapData) => {
        const nilaiValues = Object.keys(siswa.nilaiMapel).map(
          (k) => siswa.nilaiMapel[k]
        );
        const nilaiList = nilaiValues.filter((n): n is number => n !== null);

        const jumlah =
          nilaiList.length > 0
            ? nilaiList.reduce((a: number, b: number) => a + b, 0)
            : 0;

        const rataRata = nilaiList.length > 0 ? jumlah / nilaiList.length : 0;

        return {
          ...siswa,
          jumlah: parseFloat(jumlah.toFixed(2)),
          rataRata: parseFloat(rataRata.toFixed(2)),
          ranking: 0,
          catatan: "",
        };
      });

      // Sort dan assign ranking
      rekapArray.sort((a, b) => b.jumlah - a.jumlah);

      let currentRank = 1;
      for (let i = 0; i < rekapArray.length; i++) {
        if (i === 0) {
          rekapArray[i].ranking = currentRank;
        } else {
          if (rekapArray[i].jumlah === rekapArray[i - 1].jumlah) {
            rekapArray[i].ranking = rekapArray[i - 1].ranking;
          } else {
            currentRank = i + 1;
            rekapArray[i].ranking = currentRank;
          }
        }
        rekapArray[i].catatan = getCatatanByRanking(rekapArray[i].ranking);
      }

      setRekapData(rekapArray);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Unknown error");
    } finally {
      if (!silent) {
        setLoading(false);
      }
    }
  };

  // Load data saat pertama kali component mount
  useEffect(() => {
    refreshRekapData();
  }, []);

  return (
    <RekapContext.Provider
      value={{
        rekapData,
        availableSheets,
        schoolData,
        kehadiranData,
        loading,
        error,
        refreshRekapData,
      }}
    >
      {children}
    </RekapContext.Provider>
  );
};

// Hook untuk menggunakan context
const useRekapData = () => {
  const context = React.useContext(RekapContext);
  if (!context) {
    throw new Error("useRekapData must be used within RekapProvider");
  }
  return context;
};

const InputNilai = () => {
  const [data, setData] = useState<RowData[]>([]);
  const [changedRows, setChangedRows] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [selectedSheet, setSelectedSheet] = useState<string>("MAPEL101");
  const [availableSheets, setAvailableSheets] = useState<SheetInfo[]>([]);
  const [showTPPopup, setShowTPPopup] = useState(false);
  const [selectedTP, setSelectedTP] = useState<string>("");
  const [tpDetails, setTPDetails] = useState<any>(null);
  const [loadingTP, setLoadingTP] = useState(false);
  const [showFloatingButton, setShowFloatingButton] = useState(false);
  const [floatingButtonPosition, setFloatingButtonPosition] = useState({
    top: 0,
    left: 0,
    visible: true, // Tambahkan flag visible
  });
  const [activeInput, setActiveInput] = useState<{
    rowIndex: number;
    colIndex: number;
  } | null>(null);
  const [isProcessingClick, setIsProcessingClick] = useState(false);
  const [showDescPopup, setShowDescPopup] = useState(false);
  const [selectedStudentDesc, setSelectedStudentDesc] = useState<{
    nama: string;
    descMin: string;
    descMax: string;
    tpMin: string;
    tpMax: string;
    nilaiMin: string;
    nilaiMax: string;
  } | null>(null);
  const [isLoadingDesc, setIsLoadingDesc] = useState(false);

  // useEffect #1: Fetch daftar semua sheet MAPEL (hanya sekali saat component mount)
  useEffect(() => {
    const fetchSheetList = async () => {
      try {
        const response = await fetch(`${endpoint}?action=listSheets`);
        if (!response.ok) throw new Error("Failed to fetch sheet list");
        const sheets = await response.json();
        setAvailableSheets(sheets);

        // Set sheet pertama sebagai default jika ada
        if (sheets.length > 0) {
          setSelectedSheet(sheets[0].sheetName);
        }
      } catch (err) {
        console.error("Error fetching sheets:", err);
        setError("Gagal memuat daftar sheet");
      }
    };
    fetchSheetList();
  }, []);

  // useEffect #2: Fetch data dari sheet yang dipilih
  useEffect(() => {
    const fetchData = async () => {
      if (!selectedSheet) return;

      setLoading(true);
      setError(null);
      try {
        const response = await fetch(`${endpoint}?sheet=${selectedSheet}`);
        if (!response.ok) {
          throw new Error("Network response was not ok");
        }
        const jsonData = await response.json();

        // ✅ FILTER: Hanya ambil baris yang memiliki nama siswa (Data4 tidak kosong)
        if (jsonData.length > 0) {
          const headers = jsonData[0];
          const filteredData = jsonData.slice(1).filter((row: any) => {
            return row.Data4 && row.Data4.trim() !== "";
          });
          setData([headers, ...filteredData]);
        } else {
          setData(jsonData);
        }
      } catch (err) {
        setError(err instanceof Error ? err.message : "Unknown error");
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, [selectedSheet]);

  const reloadDataKehadiran = async () => {
    try {
      const response = await fetch(`${endpoint}?sheet=${selectedSheet}`);
      if (!response.ok) {
        throw new Error("Failed to reload data");
      }
      const jsonData = await response.json();

      // ✅ FILTER: Hanya ambil baris yang memiliki nama siswa (Data4 tidak kosong)
      if (jsonData.length > 0) {
        const headers = jsonData[0];
        const filteredData = jsonData.slice(1).filter((row: any) => {
          return row.Data4 && row.Data4.trim() !== "";
        });
        const cleanedData = [headers, ...filteredData];
        setData(cleanedData);
        return cleanedData;
      } else {
        setData(jsonData);
        return jsonData;
      } // Return data untuk digunakan langsung
    } catch (err) {
      console.error("Error reloading data:", err);
      return null;
    }
  };

  // useEffect #3: Update posisi tombol saat scroll dan cek visibility
  useEffect(() => {
    const updateButtonPosition = () => {
      if (showFloatingButton && activeInput) {
        const { rowIndex, colIndex } = activeInput;
        const input = document.getElementById(
          `input-${rowIndex}-${colIndex}`
        ) as HTMLInputElement;

        if (input) {
          const rect = input.getBoundingClientRect();
          const tableContainer = document.getElementById(
            "table-scroll-container"
          );

          if (tableContainer) {
            const containerRect = tableContainer.getBoundingClientRect();

            // Dapatkan tinggi header yang sebenarnya
            const thead = tableContainer.querySelector("thead");
            const headerHeight = thead ? thead.offsetHeight : 40;

            // Cek apakah input masih terlihat dalam container
            // Input harus berada di bawah header (tidak tertutup)
            const inputTopInContainer = rect.top - containerRect.top;
            const inputBottomInContainer = rect.bottom - containerRect.top;

            const isVisibleInContainer =
              inputTopInContainer >= headerHeight && // Di bawah header
              inputBottomInContainer > headerHeight && // Minimal sebagian terlihat
              rect.bottom <= containerRect.bottom && // Tidak melewati batas bawah
              rect.left >= containerRect.left - 100 && // Toleransi horizontal
              rect.right <= window.innerWidth + 100;

            // Selalu update posisi tombol (bahkan saat hidden)
            setFloatingButtonPosition({
              top: rect.top + rect.height / 2 - 28,
              left: rect.right + 10,
              visible: isVisibleInContainer,
            });
          }
        }
      }
    };

    const handleScroll = throttle(updateButtonPosition, 16);
    const tableContainer = document.getElementById("table-scroll-container");

    if (tableContainer) {
      tableContainer.addEventListener("scroll", handleScroll as any, {
        passive: true,
      });
    }

    window.addEventListener("scroll", handleScroll as any, { passive: true });

    return () => {
      if (tableContainer) {
        tableContainer.removeEventListener("scroll", handleScroll as any);
      }
      window.removeEventListener("scroll", handleScroll as any);
    };
  }, [showFloatingButton, activeInput]);

  const handleInputChange = (
    rowIndex: number,
    header: string,
    value: string
  ) => {
    const updatedData = [...data];
    updatedData[rowIndex + 1][header] = value;
    setData(updatedData);
    setChangedRows((prev) => new Set([...Array.from(prev), rowIndex]));
  };

  const handleKeyDown = (
    e: React.KeyboardEvent<HTMLInputElement>,
    rowIndex: number,
    colIndex: number,
    actualDataLength: number
  ) => {
    if (e.key === "Enter") {
      e.preventDefault();
      const nextRow = rowIndex + 1;
      if (nextRow < actualDataLength) {
        const nextInput = document.getElementById(
          `input-${nextRow}-${colIndex}`
        ) as HTMLInputElement | null;
        if (nextInput) {
          nextInput.focus();
          nextInput.select();
        }
      }
    }
  };

  const handleSaveAll = async () => {
    if (changedRows.size === 0) {
      alert("No changes to save!");
      return;
    }

    setIsSaving(true);

    const updates: Array<{ rowIndex: number; values: string[] }> = [];
    changedRows.forEach((rowIndex) => {
      const rowData = data[rowIndex + 1];
      const values = headers.map((header) => rowData[header] || "");
      updates.push({
        rowIndex: rowIndex + 3,
        values: values,
      });
    });

    try {
      const requestBody = {
        action: "update_bulk",
        sheetName: selectedSheet,
        updates: updates,
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const responseText = await response.text();

      try {
        const responseJson = JSON.parse(responseText);
        if (responseJson.error) {
          throw new Error(responseJson.error);
        }
      } catch (parseError) {
        console.log(
          "Could not parse response, but request might be successful"
        );
      }

      alert("All changes saved successfully!");
      setChangedRows(new Set());

      // ✅ TAMBAHKAN KODE INI - Reload data setelah save berhasil
      setLoading(true);
      try {
        const refreshResponse = await fetch(
          `${endpoint}?sheet=${selectedSheet}`
        );
        if (!refreshResponse.ok) {
          throw new Error("Failed to refresh data");
        }
        const jsonData = await refreshResponse.json();

        // ✅ FILTER: Hanya ambil baris yang memiliki nama siswa
        if (jsonData.length > 0) {
          const headers = jsonData[0];
          const filteredData = jsonData.slice(1).filter((row: any) => {
            return row.Data4 && row.Data4.trim() !== "";
          });
          setData([headers, ...filteredData]);
        } else {
          setData(jsonData);
        }
      } catch (refreshError) {
        console.error("Error refreshing data:", refreshError);
        alert("Data saved but failed to refresh. Please reload the page.");
      } finally {
        setLoading(false);
      }
      // ✅ AKHIR TAMBAHAN
    } catch (err) {
      console.error("=== ERROR DETAILS ===");
      console.error(err);
      alert(
        "Error updating rows: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  const handleSheetChange = (newSheet: string) => {
    setSelectedSheet(newSheet);
    setChangedRows(new Set()); // Reset perubahan saat ganti sheet
  };

  if (loading)
    return (
      <div
        style={{
          textAlign: "center",
          fontSize: "18px",
          color: "#666",
          padding: "20px",
        }}
      >
        Loading...
      </div>
    );
  if (error)
    return (
      <div
        style={{
          textAlign: "center",
          fontSize: "18px",
          color: "red",
          padding: "20px",
        }}
      >
        Error: {error}
      </div>
    );
  if (data.length === 0)
    return (
      <div
        style={{
          textAlign: "center",
          fontSize: "18px",
          color: "#666",
          padding: "20px",
        }}
      >
        No data available
      </div>
    );

  const headers = [
    "Data1",
    "Data2",
    "Data3",
    "Data4",
    "Data5",
    "Data6",
    "Data7",
    "Data8",
    "Data9",
    "Data10",
    "Data11",
    "Data12",
    "Data13",
    "Data14",
    "Data15",
    "Data16",
    "Data17",
    "Data18",
    "Data19",
    "Data20",
    "Data21",
    "Data22",
    "Data23",
    "Data24", // ✅ TAMPILKAN
    "Data25", // ✅ TAMPILKAN
    "Data26",
    "Data27",
    "Data28",
    "Data29",
    "Data30", // ✅ TAMBAH BARU
    "Data31", // ✅ TAMBAH BARU
  ];

  const displayHeaders = headers.map((header) => data[0][header] || "");

  const readOnlyHeaders = new Set([
    "Data1",
    "Data2",
    "Data3",
    "Data4",
    "Data20",
    "Data21",
    "Data23",
    "Data24",
    "Data25",
  ]);

  const conditionalHeaders = [
    "Data5",
    "Data6",
    "Data7",
    "Data8",
    "Data9",
    "Data10",
    "Data11",
    "Data12",
    "Data13",
    "Data14",
    "Data15",
    "Data16",
    "Data17",
    "Data18",
    "Data19",
    "Data22",
  ];

  const fixedWidthHeaders = new Set([
    "Data5",
    "Data6",
    "Data7",
    "Data8",
    "Data9",
    "Data10",
    "Data11",
    "Data12",
    "Data13",
    "Data14",
    "Data15",
    "Data16",
    "Data17",
    "Data18",
    "Data19",
    "Data20",
    "Data21",
    "Data22",
    "Data23",
  ]);

  const frozenHeaders = new Set(["Data4"]);
  const hiddenHeaders = new Set([
    "Data1",
    "Data2",
    "Data3",
    "Data26", // TP Min
    "Data27", // TP Max
    "Data28", // Nilai Min
    "Data29", // Nilai Max
    "Data30", // ✅ TAMBAH BARU
    "Data31", // ✅ TAMBAH BARU
  ]);

  const visibleHeaders = headers.filter((header, index) => {
    // Data24 dan Data25 sekarang DITAMPILKAN, tidak di-filter
    // if (header === "Data24" || header === "Data25") {
    //   return false;
    // }

    if (hiddenHeaders.has(header)) {
      return false;
    }

    if (hiddenHeaders.has(header)) {
      return false;
    }

    if (conditionalHeaders.indexOf(header) !== -1) {
      return displayHeaders[index] !== "-";
    }
    return true;
  });

  const visibleDisplayHeaders = visibleHeaders.map(
    (header) => data[0][header] || ""
  );

  const actualData = data.slice(1);

  const getColumnWidth = (header: string): string => {
    if (header === "Data4") return "120px";
    if (header === "Data20") return "120px";
    if (header === "Data21") return "100px";
    if (header === "Data22") return "100px";
    if (header === "Data23") return "100px";
    if (fixedWidthHeaders.has(header)) return "50px";
    return "90px";
  };

  const getFrozenLeftPosition = (header: string): number => {
    if (header === "Data4") {
      return 80;
    }
    return 0;
  };

  const fetchTPDetails = async (
    tpCode: string,
    mapel: string,
    rowIndex: number
  ) => {
    console.log("Fetching TP:", tpCode, "for Mapel:", mapel);

    setLoadingTP(true);
    setShowTPPopup(true);
    setSelectedTP(tpCode);

    try {
      const url = `${endpoint}?sheet=DataTP&tp=${encodeURIComponent(
        tpCode
      )}&mapel=${encodeURIComponent(mapel)}`;
      console.log("Request URL:", url);

      const response = await fetch(url);
      if (!response.ok) throw new Error("Failed to fetch TP details");

      const tpData = await response.json();
      console.log("Response data:", tpData);

      // Langsung set data TP tanpa deskripsi
      setTPDetails(tpData);
    } catch (err) {
      console.error("Error fetching TP details:", err);
      setTPDetails({ error: "Gagal memuat data" });
    } finally {
      setLoadingTP(false);
    }
  };

  const handleFloatingArrowClick = () => {
    // Prevent double execution
    if (isProcessingClick) return;

    setIsProcessingClick(true);

    if (activeInput) {
      const { rowIndex, colIndex } = activeInput;
      const nextRow = rowIndex + 1;
      if (nextRow < actualData.length) {
        const nextInput = document.getElementById(
          `input-${nextRow}-${colIndex}`
        ) as HTMLInputElement | null;
        if (nextInput) {
          nextInput.focus();
          nextInput.select();
        }
      }
    }

    // Reset flag setelah delay
    setTimeout(() => {
      setIsProcessingClick(false);
    }, 300);
  };

  const updateFloatingButtonPosition = (
    element: HTMLInputElement,
    rowIndex: number,
    colIndex: number,
    forceShow: boolean = true
  ) => {
    const rect = element.getBoundingClientRect();

    setFloatingButtonPosition({
      top: rect.top + rect.height / 2 - 28,
      left: rect.right + 10,
      visible: true, // Set visible saat pertama kali focus
    });
    setActiveInput({ rowIndex, colIndex });

    if (forceShow) {
      setShowFloatingButton(rowIndex < actualData.length - 1);
    }
  };

  return (
    <div style={{ padding: "10px", margin: "0 auto", maxWidth: "100vw" }}>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "15px",
          fontSize: "20px",
        }}
      >
        Data Editor - Multi Sheet
      </h1>

      {/* Dropdown Pilih Sheet */}
      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <label style={{ fontSize: "14px", color: "#666", marginRight: "10px" }}>
          Pilih Mapel:
        </label>
        <select
          value={selectedSheet}
          onChange={(e) => handleSheetChange(e.target.value)}
          style={{
            padding: "10px 15px",
            fontSize: "16px",
            borderRadius: "4px",
            border: "1px solid #ddd",
            minWidth: "300px",
            cursor: "pointer",
            backgroundColor: "white",
          }}
        >
          {availableSheets.map((sheet, index) => (
            <option key={index} value={sheet.sheetName}>
              {sheet.mapel} - {sheet.kelas} (Semester {sheet.semester})
            </option>
          ))}
        </select>
      </div>

      {/* Info Sheet yang Sedang Dibuka */}
      <div
        style={{
          textAlign: "center",
          marginBottom: "10px",
          fontSize: "16px",
          color: "#333",
        }}
      >
        Mapel: {actualData[0]?.Data1 || "N/A"} | Kelas:{" "}
        {actualData[0]?.Data3 || "N/A"} | Semester:{" "}
        {actualData[0]?.Data2 || "N/A"}
      </div>

      {/* Tombol Save */}
      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <button
          onClick={handleSaveAll}
          disabled={isSaving}
          style={{
            padding: "12px 24px",
            backgroundColor: isSaving ? "#ccc" : "#4CAF50",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: isSaving ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            width: "100%",
            maxWidth: "300px",
          }}
          onMouseOver={(e) =>
            !isSaving &&
            ((e.target as HTMLButtonElement).style.backgroundColor = "#45a049")
          }
          onMouseOut={(e) =>
            !isSaving &&
            ((e.target as HTMLButtonElement).style.backgroundColor = "#4CAF50")
          }
        >
          {isSaving ? "Memproses..." : `Save All Changes (${changedRows.size})`}
        </button>
      </div>

      {/* Table */}
      <div
        id="table-scroll-container" // ← TAMBAHKAN INI
        style={{
          overflowX: "auto",
          overflowY: "auto",
          maxHeight: "calc(100vh - 250px)",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          position: "relative",
          WebkitOverflowScrolling: "touch",
          transform: "translateZ(0)",
        }}
      >
        <table
          style={{
            borderCollapse: "separate",
            borderSpacing: 0,
            minWidth: "100%",
            width: "max-content",
            tableLayout: "fixed",
          }}
        >
          <thead style={{ position: "sticky", top: 0, zIndex: 100 }}>
            <tr style={{ backgroundColor: "#f4f4f4" }}>
              <th
                style={{
                  padding: "8px 4px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "35px",
                  minWidth: "35px",
                  position: "sticky",
                  left: 0,
                  top: 0,
                  backgroundColor: "#f4f4f4",
                  zIndex: 3,
                  boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                  fontSize: "12px",
                }}
              >
                No.
              </th>
              <th
                style={{
                  padding: "8px 4px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "45px",
                  minWidth: "45px",
                  position: "sticky",
                  left: "35px", // ← Ubah dari 50px ke 35px (mengikuti lebar kolom No)
                  top: 0,
                  backgroundColor: "#f4f4f4",
                  zIndex: 3,
                  boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                  fontSize: "12px",
                }}
              >
                Desc
              </th>
              {visibleDisplayHeaders.map((header, index) => {
                const currentHeader = visibleHeaders[index];
                const isFrozen = frozenHeaders.has(currentHeader);
                const leftPos = isFrozen
                  ? getFrozenLeftPosition(currentHeader)
                  : "auto";
                const colWidth = getColumnWidth(currentHeader);

                return (
                  <th
                    key={index}
                    onClick={(e) => {
                      // Cek apakah ini kolom TP (Data5-Data19)
                      if (
                        conditionalHeaders.indexOf(currentHeader) !== -1 &&
                        ["Data20", "Data21", "Data22", "Data23"].indexOf(
                          currentHeader
                        ) === -1 &&
                        displayHeaders[headers.indexOf(currentHeader)] !== "-"
                      ) {
                        const tpCode =
                          displayHeaders[headers.indexOf(currentHeader)];
                        const mapel = actualData[0]?.Data1 || "";

                        // Cari baris pertama untuk mendapatkan deskripsi (karena deskripsi sama untuk semua siswa di TP yang sama)
                        fetchTPDetails(tpCode, mapel, 0);
                      }
                    }}
                    style={{
                      padding: "8px 4px",
                      textAlign: "center",
                      borderBottom: "2px solid #ddd",
                      fontWeight: "bold",
                      width: colWidth,
                      minWidth: colWidth,
                      maxWidth: colWidth,
                      position: "sticky",
                      left: isFrozen ? leftPos : "auto",
                      backgroundColor: "#f4f4f4",
                      zIndex: isFrozen ? 2 : 1,
                      boxShadow: isFrozen
                        ? "2px 0 5px rgba(0,0,0,0.1)"
                        : "none",
                      fontSize: "12px",
                      whiteSpace: "nowrap",
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      cursor:
                        conditionalHeaders.indexOf(currentHeader) !== -1 &&
                        ["Data20", "Data21", "Data22", "Data23"].indexOf(
                          currentHeader
                        ) === -1 &&
                        displayHeaders[headers.indexOf(currentHeader)] !== "-"
                          ? "pointer"
                          : "default",
                      transition: "background-color 0.2s",
                    }}
                    onMouseEnter={(e) => {
                      if (
                        conditionalHeaders.indexOf(currentHeader) !== -1 &&
                        ["Data20", "Data21", "Data22", "Data23"].indexOf(
                          currentHeader
                        ) === -1 &&
                        displayHeaders[headers.indexOf(currentHeader)] !== "-"
                      ) {
                        (e.target as HTMLElement).style.backgroundColor =
                          "#e0e0e0";
                      }
                    }}
                    onMouseLeave={(e) => {
                      (e.target as HTMLElement).style.backgroundColor =
                        "#f4f4f4";
                    }}
                  >
                    {header}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {actualData.map((row, rowIndex) => (
              <tr
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                }}
              >
                <td
                  style={{
                    padding: "6px 4px",
                    borderBottom: "1px solid #eee",
                    textAlign: "center",
                    fontWeight: "bold",
                    color: "#666",
                    width: "35px",
                    minWidth: "35px",
                    position: "sticky",
                    left: 0,
                    backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                    zIndex: 2,
                    boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                    fontSize: "12px",
                  }}
                >
                  {rowIndex + 1}
                </td>
                <td
                  style={{
                    padding: "4px",
                    borderBottom: "1px solid #eee",
                    textAlign: "center",
                    width: "45px",
                    minWidth: "45px",
                    position: "sticky",
                    left: "35px",
                    backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                    zIndex: 2,
                    boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                  }}
                >
                  <button
                    onClick={async () => {
                      setIsLoadingDesc(true); // Tampilkan loading
                      const freshData = await reloadDataKehadiran();
                      setIsLoadingDesc(false); // Sembunyikan loading

                      if (freshData && freshData.length > rowIndex + 1) {
                        const freshRow = freshData[rowIndex + 1];
                        setSelectedStudentDesc({
                          nama: freshRow.Data4 || "",
                          descMin: freshRow.Data26 || "Tidak ada deskripsi",
                          descMax: freshRow.Data27 || "Tidak ada deskripsi",
                          tpMin: freshRow.Data28 || "-",
                          tpMax: freshRow.Data29 || "-",
                          nilaiMin: freshRow.Data30 || "-",
                          nilaiMax: freshRow.Data31 || "-",
                        });
                        setShowDescPopup(true);
                      }
                    }}
                    disabled={isLoadingDesc}
                    style={{
                      width: "100%",
                      padding: "6px",
                      backgroundColor: isLoadingDesc ? "#ccc" : "#2196F3",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                      cursor: isLoadingDesc ? "not-allowed" : "pointer",
                      fontSize: "20px",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      fontWeight: "bold",
                    }}
                  >
                    {isLoadingDesc ? "⏳" : "±"}
                  </button>
                </td>
                {visibleHeaders.map((header, colIndex) => {
                  const isFrozen = frozenHeaders.has(header);
                  const leftPos = isFrozen
                    ? getFrozenLeftPosition(header)
                    : "auto";
                  const colWidth = getColumnWidth(header);

                  return (
                    <td
                      key={colIndex}
                      style={{
                        padding: "4px",
                        borderBottom: "1px solid #eee",
                        width: colWidth,
                        minWidth: colWidth,
                        maxWidth: colWidth,
                        position: isFrozen ? "sticky" : "static",
                        left: isFrozen ? leftPos : "auto",
                        backgroundColor: isFrozen
                          ? rowIndex % 2 === 0
                            ? "#fff"
                            : "#f9f9f9"
                          : "transparent",
                        zIndex: isFrozen ? 1 : 0,
                        boxShadow: isFrozen
                          ? "2px 0 5px rgba(0,0,0,0.1)"
                          : "none",
                      }}
                    >
                      {readOnlyHeaders.has(header) ? (
                        <div
                          style={{
                            padding: "4px 2px",
                            color: "#666",
                            fontWeight: "normal",
                            fontSize: "12px",
                            whiteSpace: "nowrap",
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            textAlign: header === "Data4" ? "left" : "center",
                          }}
                        >
                          {row[header] || ""}
                        </div>
                      ) : (
                        <input
                          id={`input-${rowIndex}-${colIndex}`}
                          type="text"
                          inputMode="decimal"
                          pattern="[0-9]*"
                          value={row[header] || ""}
                          onChange={(e) =>
                            handleInputChange(rowIndex, header, e.target.value)
                          }
                          onKeyDown={(e) =>
                            handleKeyDown(
                              e,
                              rowIndex,
                              colIndex,
                              actualData.length
                            )
                          }
                          onFocus={(e) => {
                            e.target.select();
                            updateFloatingButtonPosition(
                              e.target,
                              rowIndex,
                              colIndex,
                              true
                            );
                          }}
                          onBlur={(e) => {
                            // Cek apakah blur karena klik tombol arrow
                            const relatedTarget =
                              e.relatedTarget as HTMLElement;
                            if (
                              !relatedTarget ||
                              relatedTarget.tagName !== "BUTTON"
                            ) {
                              // Delay untuk memberi waktu jika tombol diklik
                              setTimeout(() => {
                                // Cek lagi apakah ada input yang sedang focus
                                const activeElement = document.activeElement;
                                const isInputFocused =
                                  activeElement?.tagName === "INPUT" &&
                                  activeElement?.id.startsWith("input-");
                                if (!isInputFocused) {
                                  setShowFloatingButton(false);
                                }
                              }, 150);
                            }
                          }}
                          style={{
                            width: "100%",
                            padding: "4px 2px",
                            border: "1px solid #ddd",
                            borderRadius: "3px",
                            boxSizing: "border-box",
                            backgroundColor: "white",
                            cursor: "text",
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                            fontSize: "12px",
                            textAlign: "center",
                          }}
                        />
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
        {/* Popup Modal untuk TP Details */}
        {showTPPopup && (
          <div
            style={{
              position: "fixed",
              top: 0,
              left: 0,
              right: 0,
              bottom: 0,
              backgroundColor: "rgba(0, 0, 0, 0.5)",
              display: "flex",
              justifyContent: "center",
              alignItems: "center",
              zIndex: 1000,
            }}
            onClick={() => setShowTPPopup(false)}
          >
            <div
              style={{
                backgroundColor: "white",
                borderRadius: "8px",
                padding: "20px",
                maxWidth: "600px",
                width: "90%",
                maxHeight: "80vh",
                overflowY: "auto",
                boxShadow: "0 4px 20px rgba(0,0,0,0.3)",
              }}
              onClick={(e) => e.stopPropagation()}
            >
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  marginBottom: "15px",
                  borderBottom: "2px solid #4CAF50",
                  paddingBottom: "10px",
                }}
              >
                <h2 style={{ margin: 0, color: "#333", fontSize: "18px" }}>
                  Rincian TP: {selectedTP}
                </h2>
                <button
                  onClick={() => setShowTPPopup(false)}
                  style={{
                    backgroundColor: "#f44336",
                    color: "white",
                    border: "none",
                    borderRadius: "4px",
                    padding: "8px 16px",
                    cursor: "pointer",
                    fontSize: "14px",
                    fontWeight: "bold",
                  }}
                >
                  Tutup
                </button>
              </div>

              {loadingTP ? (
                <div
                  style={{
                    textAlign: "center",
                    padding: "20px",
                    color: "#666",
                  }}
                >
                  Loading...
                </div>
              ) : tpDetails ? (
                <div>
                  <div style={{ marginBottom: "15px" }}>
                    <strong style={{ color: "#4CAF50" }}>Mapel:</strong>{" "}
                    <span style={{ color: "#333" }}>
                      {tpDetails.mapel || "N/A"}
                    </span>
                  </div>
                  <div style={{ marginBottom: "15px" }}>
                    <strong style={{ color: "#4CAF50" }}>TP:</strong>{" "}
                    <span style={{ color: "#333" }}>
                      {tpDetails.tp || "N/A"}
                    </span>
                  </div>
                  {/* TAMBAHAN: BAB */}
                  <div style={{ marginBottom: "15px" }}>
                    <strong style={{ color: "#4CAF50" }}>BAB:</strong>{" "}
                    <span style={{ color: "#333" }}>
                      {tpDetails.bab || "N/A"}
                    </span>
                  </div>
                  {/* AKHIR TAMBAHAN */}
                  <div style={{ marginBottom: "15px" }}>
                    <strong style={{ color: "#4CAF50" }}>Semester:</strong>{" "}
                    <span style={{ color: "#333" }}>
                      {tpDetails.semester || "N/A"}
                    </span>
                  </div>
                  <div style={{ marginBottom: "15px" }}>
                    <strong style={{ color: "#4CAF50" }}>Kelas:</strong>{" "}
                    <span style={{ color: "#333" }}>
                      {tpDetails.kelas || "N/A"}
                    </span>
                  </div>
                  <div>
                    <strong style={{ color: "#4CAF50" }}>Rincian TP:</strong>
                    <p
                      style={{
                        marginTop: "10px",
                        lineHeight: "1.6",
                        color: "#333",
                        backgroundColor: "#f9f9f9",
                        padding: "15px",
                        borderRadius: "4px",
                        border: "1px solid #e0e0e0",
                      }}
                    >
                      {tpDetails.rincian || "Tidak ada rincian"}
                    </p>
                  </div>
                </div>
              ) : (
                <div
                  style={{
                    textAlign: "center",
                    padding: "20px",
                    color: "#f44336",
                  }}
                >
                  Data tidak ditemukan
                </div>
              )}
            </div>
          </div>
        )}
      </div>
      {/* Popup Modal untuk Deskripsi */}
      {showDescPopup && selectedStudentDesc && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0, 0, 0, 0.5)",
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            zIndex: 1000,
          }}
          onClick={() => setShowDescPopup(false)}
        >
          <div
            style={{
              backgroundColor: "white",
              borderRadius: "8px",
              padding: "20px",
              maxWidth: "700px",
              width: "90%",
              maxHeight: "80vh",
              overflowY: "auto",
              boxShadow: "0 4px 20px rgba(0,0,0,0.3)",
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: "15px",
                borderBottom: "2px solid #2196F3",
                paddingBottom: "10px",
              }}
            >
              <h2 style={{ margin: 0, color: "#333", fontSize: "18px" }}>
                Deskripsi: {selectedStudentDesc.nama}
              </h2>
              <button
                onClick={() => setShowDescPopup(false)}
                style={{
                  backgroundColor: "#f44336",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  padding: "8px 16px",
                  cursor: "pointer",
                  fontSize: "14px",
                  fontWeight: "bold",
                }}
              >
                Tutup
              </button>
            </div>

            {/* Bagian TP Terendah dan Tertinggi */}
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "1fr 1fr",
                gap: "15px",
                marginBottom: "20px",
              }}
            >
              <div
                style={{
                  backgroundColor: "#ffebee",
                  padding: "15px",
                  borderRadius: "8px",
                  border: "2px solid #f44336",
                }}
              >
                <h3
                  style={{
                    color: "#f44336",
                    fontSize: "14px",
                    marginBottom: "8px",
                    display: "flex",
                    alignItems: "center",
                    gap: "5px",
                  }}
                >
                  <span style={{ fontSize: "20px" }}>📉</span> TP Terendah
                </h3>
                <p
                  style={{
                    fontSize: "24px",
                    fontWeight: "bold",
                    color: "#c62828",
                    margin: 0,
                    textAlign: "center",
                  }}
                >
                  {selectedStudentDesc.tpMin}
                </p>
              </div>

              <div
                style={{
                  backgroundColor: "#e8f5e9",
                  padding: "15px",
                  borderRadius: "8px",
                  border: "2px solid #4CAF50",
                }}
              >
                <h3
                  style={{
                    color: "#4CAF50",
                    fontSize: "14px",
                    marginBottom: "8px",
                    display: "flex",
                    alignItems: "center",
                    gap: "5px",
                  }}
                >
                  <span style={{ fontSize: "20px" }}>📈</span> TP Tertinggi
                </h3>
                <p
                  style={{
                    fontSize: "24px",
                    fontWeight: "bold",
                    color: "#2e7d32",
                    margin: 0,
                    textAlign: "center",
                  }}
                >
                  {selectedStudentDesc.tpMax}
                </p>
              </div>
            </div>

            {/* Bagian Nilai Terendah dan Tertinggi */}
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "1fr 1fr",
                gap: "15px",
                marginBottom: "20px",
              }}
            >
              <div
                style={{
                  backgroundColor: "#fff3e0",
                  padding: "15px",
                  borderRadius: "8px",
                  border: "2px solid #ff9800",
                }}
              >
                <h3
                  style={{
                    color: "#ff9800",
                    fontSize: "14px",
                    marginBottom: "8px",
                    display: "flex",
                    alignItems: "center",
                    gap: "5px",
                  }}
                >
                  <span style={{ fontSize: "20px" }}>📊</span> Nilai Terendah
                </h3>
                <p
                  style={{
                    fontSize: "24px",
                    fontWeight: "bold",
                    color: "#e65100",
                    margin: 0,
                    textAlign: "center",
                  }}
                >
                  {selectedStudentDesc.nilaiMin}
                </p>
              </div>

              <div
                style={{
                  backgroundColor: "#e3f2fd",
                  padding: "15px",
                  borderRadius: "8px",
                  border: "2px solid #2196F3",
                }}
              >
                <h3
                  style={{
                    color: "#2196F3",
                    fontSize: "14px",
                    marginBottom: "8px",
                    display: "flex",
                    alignItems: "center",
                    gap: "5px",
                  }}
                >
                  <span style={{ fontSize: "20px" }}>🎯</span> Nilai Tertinggi
                </h3>
                <p
                  style={{
                    fontSize: "24px",
                    fontWeight: "bold",
                    color: "#1565c0",
                    margin: 0,
                    textAlign: "center",
                  }}
                >
                  {selectedStudentDesc.nilaiMax}
                </p>
              </div>
            </div>

            {/* Deskripsi Minimal */}
            <div style={{ marginBottom: "20px" }}>
              <h3
                style={{
                  color: "#ff9800",
                  fontSize: "16px",
                  marginBottom: "10px",
                  display: "flex",
                  alignItems: "center",
                  gap: "5px",
                }}
              >
                <span style={{ fontSize: "18px" }}>⚠️</span> Deskripsi Minimal
              </h3>
              <p
                style={{
                  lineHeight: "1.6",
                  color: "#333",
                  backgroundColor: "#fff3cd",
                  padding: "15px",
                  borderRadius: "4px",
                  border: "1px solid #ffc107",
                  margin: 0,
                }}
              >
                {selectedStudentDesc.descMin}
              </p>
            </div>

            {/* Deskripsi Maksimal */}
            <div>
              <h3
                style={{
                  color: "#4CAF50",
                  fontSize: "16px",
                  marginBottom: "10px",
                  display: "flex",
                  alignItems: "center",
                  gap: "5px",
                }}
              >
                <span style={{ fontSize: "18px" }}>✅</span> Deskripsi Maksimal
              </h3>
              <p
                style={{
                  lineHeight: "1.6",
                  color: "#333",
                  backgroundColor: "#d4edda",
                  padding: "15px",
                  borderRadius: "4px",
                  border: "1px solid #28a745",
                  margin: 0,
                }}
              >
                {selectedStudentDesc.descMax}
              </p>
            </div>
          </div>
        </div>
      )}
      {/* Floating Arrow Button - Dynamic Position */}
      {showFloatingButton && (
        <button
          onClick={(e) => {
            e.preventDefault();
            e.stopPropagation();
            handleFloatingArrowClick();
          }}
          style={{
            position: "fixed",
            top: `${floatingButtonPosition.top}px`,
            left: `${floatingButtonPosition.left}px`,
            width: "56px",
            height: "56px",
            borderRadius: "50%",
            backgroundColor: "#4CAF50",
            color: "white",
            border: "none",
            cursor: "pointer",
            boxShadow: "0 4px 12px rgba(0,0,0,0.4)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: "28px",
            fontWeight: "bold",
            zIndex: 1001,
            transition: "all 0.2s ease",
            pointerEvents: floatingButtonPosition.visible ? "auto" : "none", // Disable click saat hidden
            opacity: floatingButtonPosition.visible ? 1 : 0, // Hide dengan opacity
            visibility: floatingButtonPosition.visible ? "visible" : "hidden", // Hide dengan visibility
            touchAction: "manipulation",
            WebkitTapHighlightColor: "transparent",
          }}
          onMouseEnter={(e) => {
            if (floatingButtonPosition.visible) {
              (e.target as HTMLButtonElement).style.backgroundColor = "#45a049";
              (e.target as HTMLButtonElement).style.transform = "scale(1.1)";
            }
          }}
          onMouseLeave={(e) => {
            (e.target as HTMLButtonElement).style.backgroundColor = "#4CAF50";
            (e.target as HTMLButtonElement).style.transform = "scale(1)";
          }}
        >
          ↓
        </button>
      )}
    </div>
  );
};

const DataSekolah = () => {
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);
  const [namaKepsek, setNamaKepsek] = useState("");
  const [nipKepsek, setNipKepsek] = useState("");
  const [namaGuru, setNamaGuru] = useState("");
  const [nipGuru, setNipGuru] = useState("");
  const [ttdKepsek, setTtdKepsek] = useState("");
  const [ttdGuru, setTtdGuru] = useState("");
  const [loading, setLoading] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [isKepsekSigning, setIsKepsekSigning] = useState(false);
  const [isGuruSigning, setIsGuruSigning] = useState(false);
  const kepsekSigCanvas = useRef<SignatureCanvas>(null);
  const guruSigCanvas = useRef<SignatureCanvas>(null);
  const [namaSekolah, setNamaSekolah] = useState("");
  const [npsn, setNpsn] = useState("");
  const [alamatSekolah, setAlamatSekolah] = useState("");
  const [kodePos, setKodePos] = useState("");
  const [desaKelurahan, setDesaKelurahan] = useState("");
  const [kabKota, setKabKota] = useState("");
  const [provinsi, setProvinsi] = useState("");
  const [tahunPelajaran, setTahunPelajaran] = useState(""); // ✅ TAMBAHAN BARU
  const [tanggalRapor, setTanggalRapor] = useState(""); // ✅ TAMBAHAN BARU

  useEffect(() => {
    fetch(`${endpoint}?action=schoolData`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success && data.data && data.data.length > 0) {
          const record = data.data[0];
          setSchoolData(record);

          setNamaSekolah(record.namaSekolah || "");
          setNpsn(record.npsn || "");
          setAlamatSekolah(record.alamatSekolah || "");
          setKodePos(record.kodePos || "");
          setDesaKelurahan(record.desaKelurahan || "");
          setKabKota(record.kabKota || "");
          setProvinsi(record.provinsi || "");
          setTahunPelajaran(record.tahunPelajaran || ""); // ✅ TAMBAHAN BARU
          setTanggalRapor(record.tanggalRapor || ""); // ✅ TAMBAHAN BARU

          setNamaKepsek(record.namaKepsek);
          setNipKepsek(record.nipKepsek);
          setTtdKepsek(record.ttdKepsek);
          setNamaGuru(record.namaGuru);
          setNipGuru(record.nipGuru);
          setTtdGuru(record.ttdGuru);
        }
        setLoading(false);
      });
  }, []);

  const handleSave = () => {
    if (!namaSekolah || !namaKepsek || !nipKepsek || !namaGuru || !nipGuru) {
      alert(
        "⚠️ Nama Sekolah, Nama & NIP Kepsek, dan Nama & NIP Guru wajib diisi!"
      );
      return;
    }

    setIsSaving(true);

    const data: SchoolData = {
      namaSekolah,
      npsn,
      alamatSekolah,
      kodePos,
      desaKelurahan,
      kabKota,
      provinsi,
      tahunPelajaran, // ✅ TAMBAHAN BARU
      tanggalRapor, // ✅ TAMBAHAN BARU
      namaKepsek,
      nipKepsek,
      ttdKepsek: ttdKepsek || "",
      namaGuru,
      nipGuru,
      ttdGuru: ttdGuru || "",
    };

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: {
        "Content-Type": "text/plain;charset=utf-8",
      },
      body: JSON.stringify({
        action: "schoolData",
        ...data,
      }),
    })
      .then(() => {
        alert("✅ Data sekolah berhasil diperbarui!");
        setIsSaving(false);
        window.location.reload();
      })
      .catch((error) => {
        console.error("Error saving school data:", error);
        alert("❌ Gagal memperbarui data sekolah.");
        setIsSaving(false);
      });
  };

  // Handler functions untuk signature
  const handleClearKepsekSignature = () => kepsekSigCanvas.current?.clear();
  const handleClearGuruSignature = () => guruSigCanvas.current?.clear();

  const handleSaveKepsekSignature = () => {
    const signature = kepsekSigCanvas.current?.toDataURL("image/png");
    if (signature && !kepsekSigCanvas.current?.isEmpty()) {
      setTtdKepsek(signature);
      setIsKepsekSigning(false);
    } else {
      alert("⚠️ Tanda tangan kepala sekolah kosong!");
    }
  };

  const handleSaveGuruSignature = () => {
    const signature = guruSigCanvas.current?.toDataURL("image/png");
    if (signature && !guruSigCanvas.current?.isEmpty()) {
      setTtdGuru(signature);
      setIsGuruSigning(false);
    } else {
      alert("⚠️ Tanda tangan guru kosong!");
    }
  };

  if (loading) {
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        Memuat data sekolah...
      </div>
    );
  }

  return (
    <div style={{ maxWidth: "800px", margin: "0 auto", padding: "20px" }}>
      <div
        style={{
          backgroundColor: "white",
          padding: "24px",
          borderRadius: "8px",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
        }}
      >
        <h2
          style={{
            fontSize: "24px",
            fontWeight: "bold",
            textAlign: "center",
            color: "#2563eb",
            marginBottom: "24px",
          }}
        >
          🏫 Data Sekolah
        </h2>

        <div style={{ marginBottom: "24px" }}>
          <h3
            style={{
              fontSize: "18px",
              fontWeight: "600",
              marginBottom: "12px",
              color: "#1e40af",
            }}
          >
            Informasi Sekolah
          </h3>

          <input
            type="text"
            placeholder="Nama Sekolah *"
            value={namaSekolah}
            onChange={(e) => setNamaSekolah(e.target.value)}
            disabled={isSaving}
            style={{
              width: "100%",
              padding: "10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              marginBottom: "8px",
            }}
          />

          <input
            type="text"
            placeholder="NPSN"
            value={npsn}
            onChange={(e) => setNpsn(e.target.value)}
            disabled={isSaving}
            style={{
              width: "100%",
              padding: "10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              marginBottom: "8px",
            }}
          />

          <textarea
            placeholder="Alamat Sekolah"
            value={alamatSekolah}
            onChange={(e) => setAlamatSekolah(e.target.value)}
            disabled={isSaving}
            rows={2}
            style={{
              width: "100%",
              padding: "10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              marginBottom: "8px",
              resize: "vertical",
            }}
          />

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gap: "8px",
            }}
          >
            <input
              type="text"
              placeholder="Kode Pos"
              value={kodePos}
              onChange={(e) => setKodePos(e.target.value)}
              disabled={isSaving}
              style={{
                padding: "10px",
                border: "1px solid #ddd",
                borderRadius: "4px",
              }}
            />

            <input
              type="text"
              placeholder="Desa/Kelurahan"
              value={desaKelurahan}
              onChange={(e) => setDesaKelurahan(e.target.value)}
              disabled={isSaving}
              style={{
                padding: "10px",
                border: "1px solid #ddd",
                borderRadius: "4px",
              }}
            />
          </div>

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gap: "8px",
              marginTop: "8px",
            }}
          >
            <input
              type="text"
              placeholder="Kabupaten/Kota"
              value={kabKota}
              onChange={(e) => setKabKota(e.target.value)}
              disabled={isSaving}
              style={{
                padding: "10px",
                border: "1px solid #ddd",
                borderRadius: "4px",
              }}
            />

            <input
              type="text"
              placeholder="Provinsi"
              value={provinsi}
              onChange={(e) => setProvinsi(e.target.value)}
              disabled={isSaving}
              style={{
                padding: "10px",
                border: "1px solid #ddd",
                borderRadius: "4px",
              }}
            />

            {/* ✅ TAMBAHAN BARU - Input Tahun Pelajaran */}
            <input
              type="text"
              placeholder="Tahun Pelajaran (contoh: 2024/2025)"
              value={tahunPelajaran}
              onChange={(e) => setTahunPelajaran(e.target.value)}
              disabled={isSaving}
              style={{
                padding: "10px",
                border: "1px solid #ddd",
                borderRadius: "4px",
              }}
            />

            {/* Input Tanggal Rapor dengan Date Picker */}
            <div style={{ position: "relative", marginBottom: "8px" }}>
              <input
                type="date"
                value={
                  tanggalRapor
                    ? (() => {
                        // Convert dd/mm/yyyy to yyyy-mm-dd for input[type="date"]
                        const parts = tanggalRapor.split("/");
                        if (parts.length === 3) {
                          return `${parts[2]}-${parts[1]}-${parts[0]}`;
                        }
                        return "";
                      })()
                    : ""
                }
                onChange={(e) => {
                  // Convert yyyy-mm-dd to dd/mm/yyyy
                  const value = e.target.value;
                  if (value) {
                    const parts = value.split("-");
                    setTanggalRapor(`${parts[2]}/${parts[1]}/${parts[0]}`);
                  } else {
                    setTanggalRapor("");
                  }
                }}
                disabled={isSaving}
                style={{
                  padding: "10px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  width: "100%",
                }}
              />
              {/* Label helper untuk menunjukkan format yang disimpan */}
              {tanggalRapor && (
                <div
                  style={{
                    fontSize: "11px",
                    color: "#4CAF50",
                    marginTop: "4px",
                  }}
                >
                  ✓ Tersimpan: {tanggalRapor}
                </div>
              )}
            </div>
          </div>
        </div>

        <hr
          style={{
            margin: "24px 0",
            border: "none",
            borderTop: "1px solid #e5e7eb",
          }}
        />

        {/* Kepala Sekolah */}
        <div style={{ marginBottom: "24px" }}>
          <h3
            style={{
              fontSize: "18px",
              fontWeight: "600",
              marginBottom: "12px",
            }}
          >
            Kepala Sekolah
          </h3>
          <input
            type="text"
            placeholder="Nama Kepala Sekolah"
            value={namaKepsek}
            onChange={(e) => setNamaKepsek(e.target.value)}
            disabled={isSaving}
            style={{
              width: "100%",
              padding: "10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              marginBottom: "8px",
            }}
          />
          <input
            type="text"
            placeholder="NIP Kepala Sekolah"
            value={nipKepsek}
            onChange={(e) => setNipKepsek(e.target.value)}
            disabled={isSaving}
            style={{
              width: "100%",
              padding: "10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              marginBottom: "12px",
            }}
          />

          <p style={{ fontSize: "14px", color: "#666", marginBottom: "8px" }}>
            Tanda Tangan Kepala Sekolah
          </p>
          <div style={{ position: "relative" }}>
            <SignatureCanvas
              ref={kepsekSigCanvas}
              penColor="black"
              canvasProps={{
                style: {
                  width: "100%",
                  height: "200px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  opacity: !isKepsekSigning || isSaving ? 0.5 : 1,
                  pointerEvents: !isKepsekSigning || isSaving ? "none" : "auto",
                },
              }}
            />
            {!isKepsekSigning && (
              <div
                style={{
                  position: "absolute",
                  top: 0,
                  left: 0,
                  right: 0,
                  bottom: 0,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  backgroundColor: "rgba(200,200,200,0.3)",
                }}
              >
                <span style={{ color: "#666" }}>
                  Klik "Mulai Tanda Tangan" untuk mengaktifkan
                </span>
              </div>
            )}
          </div>
          <div style={{ display: "flex", gap: "8px", marginTop: "8px" }}>
            {!isKepsekSigning && (
              <button
                onClick={() => setIsKepsekSigning(true)}
                disabled={isSaving}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#10b981",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                }}
              >
                ✍️ Mulai Tanda Tangan
              </button>
            )}
            {isKepsekSigning && (
              <button
                onClick={handleSaveKepsekSignature}
                disabled={isSaving}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#3b82f6",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                }}
              >
                💾 Simpan Tanda Tangan
              </button>
            )}
            <button
              onClick={handleClearKepsekSignature}
              disabled={!isKepsekSigning || isSaving}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ef4444",
                color: "white",
                border: "none",
                borderRadius: "4px",
                cursor: "pointer",
                opacity: !isKepsekSigning || isSaving ? 0.5 : 1,
              }}
            >
              🗑️ Hapus TTD
            </button>
          </div>
          {ttdKepsek && (
            <img
              src={ttdKepsek}
              alt="TTD Kepsek"
              style={{
                marginTop: "12px",
                maxWidth: "200px",
                height: "80px",
                border: "1px solid #ddd",
                borderRadius: "4px",
              }}
            />
          )}
        </div>

        {/* Guru - Similar structure */}
        <div style={{ marginBottom: "24px" }}>
          <h3
            style={{
              fontSize: "18px",
              fontWeight: "600",
              marginBottom: "12px",
            }}
          >
            Guru Kelas
          </h3>
          <input
            type="text"
            placeholder="Nama Guru"
            value={namaGuru}
            onChange={(e) => setNamaGuru(e.target.value)}
            disabled={isSaving}
            style={{
              width: "100%",
              padding: "10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              marginBottom: "8px",
            }}
          />
          <input
            type="text"
            placeholder="NIP Guru"
            value={nipGuru}
            onChange={(e) => setNipGuru(e.target.value)}
            disabled={isSaving}
            style={{
              width: "100%",
              padding: "10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              marginBottom: "12px",
            }}
          />

          <p style={{ fontSize: "14px", color: "#666", marginBottom: "8px" }}>
            Tanda Tangan Guru
          </p>
          <div style={{ position: "relative" }}>
            <SignatureCanvas
              ref={guruSigCanvas}
              penColor="black"
              canvasProps={{
                style: {
                  width: "100%",
                  height: "200px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  opacity: !isGuruSigning || isSaving ? 0.5 : 1,
                  pointerEvents: !isGuruSigning || isSaving ? "none" : "auto",
                },
              }}
            />
            {!isGuruSigning && (
              <div
                style={{
                  position: "absolute",
                  top: 0,
                  left: 0,
                  right: 0,
                  bottom: 0,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  backgroundColor: "rgba(200,200,200,0.3)",
                }}
              >
                <span style={{ color: "#666" }}>
                  Klik "Mulai Tanda Tangan" untuk mengaktifkan
                </span>
              </div>
            )}
          </div>
          <div style={{ display: "flex", gap: "8px", marginTop: "8px" }}>
            {!isGuruSigning && (
              <button
                onClick={() => setIsGuruSigning(true)}
                disabled={isSaving}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#10b981",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                }}
              >
                ✍️ Mulai Tanda Tangan
              </button>
            )}
            {isGuruSigning && (
              <button
                onClick={handleSaveGuruSignature}
                disabled={isSaving}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#3b82f6",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                }}
              >
                💾 Simpan Tanda Tangan
              </button>
            )}
            <button
              onClick={handleClearGuruSignature}
              disabled={!isGuruSigning || isSaving}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ef4444",
                color: "white",
                border: "none",
                borderRadius: "4px",
                cursor: "pointer",
                opacity: !isGuruSigning || isSaving ? 0.5 : 1,
              }}
            >
              🗑️ Hapus TTD
            </button>
          </div>
          {ttdGuru && (
            <img
              src={ttdGuru}
              alt="TTD Guru"
              style={{
                marginTop: "12px",
                maxWidth: "200px",
                height: "80px",
                border: "1px solid #ddd",
                borderRadius: "4px",
              }}
            />
          )}
        </div>

        <div style={{ textAlign: "center" }}>
          <button
            onClick={handleSave}
            disabled={isSaving}
            style={{
              padding: "12px 24px",
              backgroundColor: isSaving ? "#93c5fd" : "#2563eb",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: isSaving ? "not-allowed" : "pointer",
              fontWeight: "600",
            }}
          >
            {isSaving ? "⏳ Menyimpan..." : "💾 Simpan Data Sekolah"}
          </button>
        </div>
      </div>
    </div>
  );
};

interface RekapData {
  nama: string;
  kelas: string;
  nilaiMapel: { [mapel: string]: number | null }; // Nilai per mapel, dinamis
  jumlah: number;
  rataRata: number;
  ranking: number;
  catatan: string;
}

const RekapNilai = () => {
  const {
    rekapData,
    availableSheets,
    schoolData,
    kehadiranData,
    loading,
    error,
    refreshRekapData,
  } = useRekapData();
  const [isAutoRefreshEnabled, setIsAutoRefreshEnabled] = useState(true);
  const [lastUpdated, setLastUpdated] = useState<Date>(new Date());
  const refreshInterval = 5000; // 5 detik (statis)
  const [downloadingId, setDownloadingId] = useState<string | null>(null);

  // Auto-refresh data setiap interval tertentu (silent mode)
  useEffect(() => {
    if (!isAutoRefreshEnabled) return;

    const intervalId = setInterval(async () => {
      try {
        await refreshRekapData(true); // true = silent mode, tanpa loading
        setLastUpdated(new Date());
      } catch (error) {
        console.error("Auto-refresh error:", error);
      }
    }, refreshInterval);

    return () => clearInterval(intervalId);
  }, [isAutoRefreshEnabled, refreshInterval, refreshRekapData]);

  // Manual refresh (dengan loading indicator)
  const handleManualRefresh = async () => {
    try {
      await refreshRekapData(false); // false = show loading
      setLastUpdated(new Date());
    } catch (error) {
      console.error("Manual refresh error:", error);
    }
  };

  // CSS untuk animasi loading
  useEffect(() => {
    const style = document.createElement("style");
    style.textContent = `
      @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }
    `;
    document.head.appendChild(style);
    return () => {
      document.head.removeChild(style);
    };
  }, []);

  if (loading)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        Loading Rekap...
      </div>
    );
  if (error)
    return (
      <div style={{ textAlign: "center", color: "red", padding: "20px" }}>
        Error: {error}
      </div>
    );
  if (rekapData.length === 0)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        No data available
      </div>
    );

  const mapelColumns = availableSheets.map((sheet) => sheet.mapel);

  const downloadRaporPDF = async (siswa: RekapData) => {
    setDownloadingId(siswa.nama);

    console.log("=== START PDF GENERATION ===");
    console.log("Siswa:", siswa.nama);

    try {
      const doc = new jsPDF();

      const fetchNISNISN = async (
        namaSiswa: string
      ): Promise<{ nis: string; nisn: string }> => {
        try {
          const response = await fetch(`${endpoint}?sheet=DataSiswa`);
          if (!response.ok) return { nis: "-", nisn: "-" };
          const siswaData = await response.json();
          const siswaRecord = siswaData
            .slice(1)
            .find((row: any) => row.Data1 === namaSiswa);
          return {
            nis: siswaRecord?.Data3 || "-",
            nisn: siswaRecord?.Data4 || "-",
          };
        } catch (error) {
          console.log("Error fetching NIS/NISN:", error);
          return { nis: "-", nisn: "-" };
        }
      };

      const { nis, nisn } = await fetchNISNISN(siswa.nama);

      // Header
      doc.setFontSize(16);
      doc.setFont("helvetica", "bold");
      doc.text("LAPORAN HASIL BELAJAR (RAPOR)", 105, 20, { align: "center" });

      // Data Siswa
      doc.setFontSize(10);
      doc.setFont("helvetica", "normal");

      const leftCol = 20;
      const rightCol = 130;
      const leftColTTD = 25;
      const centerColTTD = 100;
      const rightColTTD = 150;
      let y = 35;

      doc.text("Nama Peserta Didik", leftCol, y);
      doc.text(": " + siswa.nama.toUpperCase(), leftCol + 50, y);
      doc.text("Kelas", rightCol, y);
      doc.text(": " + (siswa.kelas || "-"), rightCol + 30, y);

      y += 7;
      doc.text("NISN/NIS", leftCol, y);
      doc.text(`: ${nisn} / ${nis}`, leftCol + 50, y);
      doc.text("Fase", rightCol, y);
      doc.text(": C", rightCol + 30, y);

      y += 7;
      doc.text("Nama Sekolah", leftCol, y);
      doc.text(
        ": " + (schoolData?.namaSekolah || "UPT SD NEGERI 2 BATANG"),
        leftCol + 50,
        y
      );
      doc.text("Semester", rightCol, y);
      doc.text(": 1", rightCol + 30, y);

      y += 7;
      doc.text("Alamat Sekolah", leftCol, y);
      // ✅ PERBAIKAN: Gabungkan alamat lengkap dalam 1 baris
      const alamatLengkap = `${
        schoolData?.alamatSekolah || "Desa Bungeng, Kecamatan Batang"
      }, ${schoolData?.kabKota || ""}`;
      doc.text(": " + alamatLengkap, leftCol + 50, y);
      doc.text("Tahun Pelajaran", rightCol, y);
      doc.text(
        ": " + (schoolData?.tahunPelajaran || "2023/2024"),
        rightCol + 30,
        y
      );

      // ✅ HAPUS baris y += 7 berikutnya dan baris kabupaten terpisah

      // Fetch deskripsi untuk setiap mapel
      const deskripsiPromises = availableSheets.map(async (sheet) => {
        const response = await fetch(`${endpoint}?sheet=${sheet.sheetName}`);
        if (!response.ok)
          return { mapel: sheet.mapel, descMin: "", descMax: "" };

        const jsonData = await response.json();
        const siswaData = jsonData
          .slice(1)
          .find((row: any) => row.Data4 === siswa.nama);

        return {
          mapel: sheet.mapel,
          descMin: siswaData?.Data26 || "",
          descMax: siswaData?.Data27 || "",
        };
      });

      const deskripsiData = await Promise.all(deskripsiPromises);

      // Ganti fungsi cleanText yang ada (sekitar baris 2066) dengan ini:
      const cleanText = (text: string): string => {
        if (!text) return "";

        // Hapus semua spasi berlebihan (termasuk yang di tengah kata)
        return text
          .replace(/\s+/g, " ") // Multiple spaces jadi single space
          .replace(/\s([.,;:!?])/g, "$1") // Hapus spasi sebelum tanda baca
          .trim();
      };

      // Tabel Nilai
      y += 10;
      const mapelColumns = availableSheets.map((sheet) => sheet.mapel);
      const tableData = mapelColumns.map((mapel, index) => {
        const nilai = siswa.nilaiMapel[mapel];
        const desc = deskripsiData.find((d) => d.mapel === mapel);

        let capaianText = "";
        if (desc?.descMax) {
          capaianText += cleanText(desc.descMax); // ✅ TAMBAHKAN cleanText()
        }
        if (desc?.descMin) {
          if (capaianText) capaianText += "\n\n";
          capaianText += cleanText(desc.descMin); // ✅ TAMBAHKAN cleanText()
        }
        if (!capaianText) {
          capaianText = "-";
        }

        let nilaiText = "-";
        if (nilai !== null && nilai !== undefined) {
          nilaiText = String(nilai);
        }

        return [index + 1, mapel, nilaiText, capaianText];
      });

      autoTable(doc, {
        startY: y,
        head: [["No.", "Mata Pelajaran", "Nilai Akhir", "Capaian Kompetensi"]],
        body: tableData,
        theme: "grid",
        headStyles: {
          fillColor: [200, 200, 200],
          textColor: 0,
          fontStyle: "bold",
          halign: "center",
        },
        columnStyles: {
          0: { cellWidth: 15, halign: "center" },
          1: { cellWidth: 50 },
          2: { cellWidth: 25, halign: "center" },
          3: {
            cellWidth: 90,
            cellPadding: 3, // ✅ TAMBAHKAN
            overflow: "linebreak", // ✅ TAMBAHKAN
            valign: "top", // ✅ TAMBAHKAN
          },
        },
        styles: {
          fontSize: 9,
          cellPadding: 3,
          overflow: "linebreak", // ✅ TAMBAHKAN
          cellWidth: "wrap", // ✅ TAMBAHKAN
        },
        rowPageBreak: "avoid",
        pageBreak: "auto",
      });

      let additionalY = doc.lastAutoTable.finalY + 10;

      // KOKURIKULER
      try {
        const kokurikulerResponse = await fetch(
          `${endpoint}?sheet=DataKokurikuler`
        );

        let kokurikulerText = "-";

        if (kokurikulerResponse.ok) {
          const kokurikulerJson = await kokurikulerResponse.json();
          const studentKokurikuler = kokurikulerJson
            .slice(1)
            .find((k: any) => k.Data1 === siswa.nama);

          if (studentKokurikuler && studentKokurikuler.Data10) {
            kokurikulerText = studentKokurikuler.Data10;
          }
        }

        const remainingSpace = 297 - additionalY;
        const estimatedTableHeight = 30;

        if (remainingSpace < estimatedTableHeight + 60) {
          doc.addPage();
          additionalY = 20;
        }

        doc.setFont("helvetica", "bold");
        doc.text("Kokurikuler", leftCol, additionalY);
        doc.setFont("helvetica", "normal");
        additionalY += 7;

        autoTable(doc, {
          startY: additionalY,
          head: [["Deskripsi Kokurikuler"]],
          body: [[kokurikulerText]],
          theme: "grid",
          headStyles: {
            fillColor: [200, 200, 200],
            textColor: 0,
            fontStyle: "bold",
            halign: "center",
          },
          columnStyles: {
            0: { cellWidth: 140, halign: "left" },
          },
          styles: {
            fontSize: 9,
            cellPadding: 5,
          },
          margin: { left: leftCol },
        });

        additionalY = doc.lastAutoTable.finalY + 10;
      } catch (error) {
        console.log("Error fetching kokurikuler data:", error);
      }

      // EKSTRAKURIKULER
      try {
        const ekstrakurikulerResponse = await fetch(
          `${endpoint}?sheet=DataEkstrakurikuler`
        );

        let ekstrakurikulerData: any[] = [];

        if (ekstrakurikulerResponse.ok) {
          const ekstrakurikulerJson = await ekstrakurikulerResponse.json();
          const studentEkstrakurikuler = ekstrakurikulerJson
            .slice(1)
            .find((k: any) => k.Data1 === siswa.nama);

          if (studentEkstrakurikuler) {
            if (studentEkstrakurikuler.Data2 || studentEkstrakurikuler.Data3) {
              ekstrakurikulerData.push([
                studentEkstrakurikuler.Data2 || "-",
                studentEkstrakurikuler.Data3 || "-",
              ]);
            }

            if (studentEkstrakurikuler.Data4 || studentEkstrakurikuler.Data5) {
              ekstrakurikulerData.push([
                studentEkstrakurikuler.Data4 || "-",
                studentEkstrakurikuler.Data5 || "-",
              ]);
            }

            if (studentEkstrakurikuler.Data6) {
              ekstrakurikulerData.push([
                studentEkstrakurikuler.Data6 || "-",
                "-",
              ]);
            }
          }
        }

        if (ekstrakurikulerData.length === 0) {
          ekstrakurikulerData = [["-", "-"]];
        }

        const remainingSpace = 297 - additionalY;
        const estimatedTableHeight = 20 + ekstrakurikulerData.length * 8;

        if (remainingSpace < estimatedTableHeight + 60) {
          doc.addPage();
          additionalY = 20;
        }

        doc.setFont("helvetica", "bold");
        doc.text("Ekstrakurikuler", leftCol, additionalY);
        doc.setFont("helvetica", "normal");
        additionalY += 7;

        autoTable(doc, {
          startY: additionalY,
          head: [["Ekstrakurikuler", "Keterangan"]],
          body: ekstrakurikulerData,
          theme: "grid",
          headStyles: {
            fillColor: [200, 200, 200],
            textColor: 0,
            fontStyle: "bold",
            halign: "center",
          },
          columnStyles: {
            0: { cellWidth: 70, halign: "left" },
            1: { cellWidth: 70, halign: "left" },
          },
          styles: {
            fontSize: 9,
            cellPadding: 5,
          },
          margin: { left: leftCol },
        });

        additionalY = doc.lastAutoTable.finalY + 10;
      } catch (error) {
        console.log("Error fetching ekstrakurikuler data:", error);
      }

      // CATATAN GURU
      const remainingSpaceBeforeCatatan = 297 - additionalY;
      const estimatedCatatanHeight = 25;

      if (remainingSpaceBeforeCatatan < estimatedCatatanHeight + 60) {
        doc.addPage();
        additionalY = 20;
      }

      doc.setFont("helvetica", "bold");
      doc.text("Catatan Guru", leftCol, additionalY);
      doc.setFont("helvetica", "normal");
      additionalY += 7;

      autoTable(doc, {
        startY: additionalY,
        head: [["Catatan"]],
        body: [[siswa.catatan || "-"]],
        theme: "grid",
        headStyles: {
          fillColor: [200, 200, 200],
          textColor: 0,
          fontStyle: "bold",
          halign: "center",
        },
        columnStyles: {
          0: { cellWidth: 140, halign: "left" },
        },
        styles: {
          fontSize: 9,
          cellPadding: 5,
        },
        margin: { left: leftCol },
      });

      additionalY = doc.lastAutoTable.finalY + 10;

      // ===== CEK RUANG UNTUK KETIDAKHADIRAN + TANDA TANGAN =====
      const studentKehadiran = kehadiranData.find(
        (k) => k.Data1 === siswa.nama
      );

      // Hitung perkiraan tinggi yang dibutuhkan:
      // Ketidakhadiran (header + tabel) ≈ 40px
      // Tanda Tangan (dengan TTD + Kepsek di bawah) ≈ 130px
      // Total minimal yang dibutuhkan ≈ 170px
      const requiredSpace = 170; // ✅ UBAH DARI 90 MENJADI 170
      const remainingSpaceBeforeKehadiran = 297 - additionalY;

      // ✅ Jika ruang tidak cukup, pindah ke halaman baru
      if (remainingSpaceBeforeKehadiran < requiredSpace) {
        doc.addPage();
        additionalY = 20;
      }

      // ===== DATA KEHADIRAN =====
      if (studentKehadiran) {
        doc.setFont("helvetica", "bold");
        doc.text("Ketidakhadiran", leftCol, additionalY);
        doc.setFont("helvetica", "normal");
        additionalY += 7;

        autoTable(doc, {
          startY: additionalY,
          head: [["Keterangan", "Jumlah Hari"]],
          body: [
            ["Sakit", `${studentKehadiran.Data7 || "0"} hari`],
            ["Izin", `${studentKehadiran.Data6 || "0"} hari`],
            ["Tanpa Keterangan", `${studentKehadiran.Data5 || "0"} hari`],
          ],
          theme: "grid",
          headStyles: {
            fillColor: [200, 200, 200],
            textColor: 0,
            fontStyle: "bold",
            halign: "center",
          },
          columnStyles: {
            0: { cellWidth: 70, halign: "left" },
            1: { cellWidth: 70, halign: "center" },
          },
          styles: {
            fontSize: 9,
            cellPadding: 3,
          },
          margin: { left: leftCol },
        });

        additionalY = doc.lastAutoTable.finalY + 15;
      }

      // ===== TANDA TANGAN =====
      // ✅ Ambil nama orang tua dari DataSiswa
      const fetchNamaOrtu = async (namaSiswa: string): Promise<string> => {
        try {
          const response = await fetch(`${endpoint}?sheet=DataSiswa`);
          if (!response.ok) return "-";
          const siswaData = await response.json();
          const siswaRecord = siswaData
            .slice(1)
            .find((row: any) => row.Data1 === namaSiswa);
          return siswaRecord?.Data5 || "-";
        } catch (error) {
          console.log("Error fetching nama ortu:", error);
          return "-";
        }
      };

      const namaOrtu = await fetchNamaOrtu(siswa.nama);

      const ttdY = additionalY;
      doc.setFontSize(10);
      doc.setFont("helvetica", "normal");

      // ✅ UBAH: Pindahkan tanggal ke atas tanda tangan guru (kolom kanan)
      const tanggalRapor = schoolData?.tanggalRapor || "23 Desember 2023";
      doc.text(`Bungeng, ${tanggalRapor}`, rightColTTD, ttdY + 10, {
        // ← TAMBAHKAN + 5 (atau nilai lain)
        align: "left",
      }); // ✅ Gunakan rightColTTD dan align left

      // ===== KOLOM KIRI - ORANG TUA / WALI =====
      doc.text("Mengetahui :", leftColTTD, ttdY + 10);
      doc.text("Orang Tua / Wali,", leftColTTD, ttdY + 15);
      doc.text(namaOrtu || "_______________", leftColTTD, ttdY + 40);

      // ===== KOLOM KANAN ATAS - WALI KELAS =====
      doc.setFontSize(10);
      doc.text("Wali Kelas,", rightColTTD, ttdY + 15);

      if (schoolData?.ttdGuru) {
        try {
          doc.addImage(
            schoolData.ttdGuru,
            "PNG",
            rightColTTD - 4,
            ttdY + 17,
            40,
            20
          );
        } catch (error) {
          console.log("Error adding guru signature:", error);
        }
      }

      doc.setFont("helvetica", "bold");
      const namaGuru = schoolData?.namaGuru || "_______________";
      doc.text(namaGuru, rightColTTD, ttdY + 40);

      // ✅ TAMBAHKAN UNDERLINE UNTUK NAMA GURU
      doc.setLineWidth(0.3);
      const guruTextWidth = doc.getTextWidth(namaGuru);
      doc.line(rightColTTD, ttdY + 41, rightColTTD + guruTextWidth, ttdY + 41);

      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.text(
        `NIP. ${schoolData?.nipGuru || "_______________"}`,
        rightColTTD,
        ttdY + 45
      );

      // ===== KOLOM TENGAH BAWAH - KEPALA SEKOLAH =====
      const kepsekY = ttdY + 55;

      doc.setFontSize(10);
      doc.setFont("helvetica", "normal");
      doc.text("Mengetahui,", centerColTTD, kepsekY, { align: "center" });
      doc.text("Kepala Sekolah", centerColTTD, kepsekY + 5, {
        align: "center",
      });

      if (schoolData?.ttdKepsek) {
        try {
          doc.addImage(
            schoolData.ttdKepsek,
            "PNG",
            centerColTTD - 20,
            kepsekY + 7,
            40,
            20
          );
        } catch (error) {
          console.log("Error adding kepsek signature:", error);
        }
      }

      doc.setFont("helvetica", "bold");
      const namaKepsek = schoolData?.namaKepsek || "_______________";
      doc.text(namaKepsek, centerColTTD, kepsekY + 30, { align: "center" });

      // ✅ TAMBAHKAN UNDERLINE UNTUK NAMA KEPSEK
      doc.setLineWidth(0.3);
      const kepsekTextWidth = doc.getTextWidth(namaKepsek);
      doc.line(
        centerColTTD - kepsekTextWidth / 2,
        kepsekY + 31,
        centerColTTD + kepsekTextWidth / 2,
        kepsekY + 31
      );

      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.text(
        `NIP. ${schoolData?.nipKepsek || "_______________"}`,
        centerColTTD,
        kepsekY + 35,
        { align: "center" }
      );
      // ===== KOLOM KANAN - WALI KELAS =====
      doc.setFontSize(10);
      doc.text("Wali Kelas,", rightColTTD, ttdY + 15);

      if (schoolData?.ttdGuru) {
        try {
          doc.addImage(
            schoolData.ttdGuru,
            "PNG",
            rightColTTD - 4,
            ttdY + 17,
            40,
            20
          );
        } catch (error) {
          console.log("Error adding guru signature:", error);
        }
      }

      doc.setFont("helvetica", "bold");
      doc.text(
        schoolData?.namaGuru || "_______________",
        rightColTTD,
        ttdY + 40
      );
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.text(
        `NIP. ${schoolData?.nipGuru || "_______________"}`,
        rightColTTD,
        ttdY + 45
      );

      // Save PDF
      doc.save(`Rapor_${siswa.nama.replace(/\s+/g, "_")}.pdf`);
    } catch (error) {
      console.error("=== PDF ERROR DETAILS ===");
      console.error("Error:", error);

      const errorMessage =
        error instanceof Error ? error.message : String(error);
      const errorStack = error instanceof Error ? error.stack : "";

      console.error("Error message:", errorMessage);
      console.error("Error stack:", errorStack);
      console.error("Siswa data:", siswa);

      alert(
        `Gagal membuat PDF untuk ${siswa.nama}\n\nError: ${errorMessage}\n\nCek console untuk detail lengkap.`
      );
    } finally {
      setDownloadingId(null);
    }
  };

  return (
    <div>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "15px",
          fontSize: "20px",
        }}
      >
        Rekap Nilai Siswa
      </h1>

      {/* Panel Kontrol Auto-Refresh */}
      <div
        style={{
          backgroundColor: "#f0f8ff",
          padding: "15px",
          borderRadius: "8px",
          marginBottom: "15px",
          display: "flex",
          flexDirection: "column",
          gap: "10px",
          boxShadow: "0 2px 5px rgba(0,0,0,0.1)",
        }}
      >
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            flexWrap: "wrap",
            gap: "10px",
          }}
        >
          <div
            style={{
              fontSize: "12px",
              color: "#666",
              fontStyle: "italic",
            }}
          >
            💡 Tip: Data akan diperbarui otomatis sesuai interval yang dipilih.
            Matikan auto-update jika tidak diperlukan untuk menghemat bandwidth.
          </div>

          <div
            style={{
              fontSize: "12px",
              color: "#666",
              display: "flex",
              alignItems: "center",
              gap: "5px",
            }}
          >
            <span>⏰ Terakhir diperbarui:</span>
            <span style={{ fontWeight: "bold", color: "#4CAF50" }}>
              {lastUpdated.toLocaleTimeString("id-ID")}
            </span>
          </div>
        </div>

        <div
          style={{
            fontSize: "12px",
            color: "#666",
            fontStyle: "italic",
          }}
        >
          💡 Tip: Data akan diperbarui otomatis sesuai interval yang dipilih.
          Matikan auto-update jika tidak diperlukan untuk menghemat bandwidth.
        </div>
      </div>

      <div
        style={{
          overflowX: "auto",
          maxHeight: "calc(100vh - 150px)",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          position: "relative",
        }}
      >
        <table
          style={{
            borderCollapse: "separate",
            borderSpacing: 0,
            minWidth: "100%",
            tableLayout: "fixed",
          }}
        >
          <thead
            style={{
              position: "sticky",
              top: 0,
              backgroundColor: "#f4f4f4",
              zIndex: 100,
            }}
          >
            <tr>
              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  width: "50px",
                  position: "sticky",
                  left: 0,
                  backgroundColor: "#f4f4f4",
                  zIndex: 3,
                  boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                }}
              >
                No
              </th>
              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  width: "200px",
                  position: "sticky",
                  left: "50px",
                  backgroundColor: "#f4f4f4",
                  zIndex: 3,
                  boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                }}
              >
                Nama
              </th>
              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  width: "100px",
                }}
              >
                Kelas
              </th>
              {mapelColumns.map((mapel, index) => (
                <th
                  key={index}
                  style={{
                    padding: "8px",
                    textAlign: "center",
                    borderBottom: "2px solid #ddd",
                    width: "100px",
                  }}
                >
                  {mapel.toUpperCase()}
                </th>
              ))}
              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  width: "100px",
                }}
              >
                JUMLAH
              </th>
              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  width: "100px",
                }}
              >
                RATA-RATA
              </th>
              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  width: "120px",
                }}
              >
                RANKING
              </th>
              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  width: "120px",
                }}
              >
                AKSI
              </th>
            </tr>
          </thead>
          <tbody>
            {rekapData.map((siswa, index) => (
              <tr
                key={index}
                style={{
                  backgroundColor: index % 2 === 0 ? "#fff" : "#f9f9f9",
                }}
              >
                <td
                  style={{
                    padding: "8px",
                    textAlign: "center",
                    borderBottom: "1px solid #eee",
                    position: "sticky",
                    left: 0,
                    backgroundColor: index % 2 === 0 ? "#fff" : "#f9f9f9",
                    zIndex: 2,
                    boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                  }}
                >
                  {index + 1}
                </td>
                <td
                  style={{
                    padding: "8px",
                    textAlign: "left",
                    borderBottom: "1px solid #eee",
                    position: "sticky",
                    left: "50px",
                    backgroundColor: index % 2 === 0 ? "#fff" : "#f9f9f9",
                    zIndex: 2,
                    boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                  }}
                >
                  {siswa.nama}
                </td>
                <td
                  style={{
                    padding: "8px",
                    textAlign: "center",
                    borderBottom: "1px solid #eee",
                  }}
                >
                  {siswa.kelas}
                </td>
                {mapelColumns.map((mapel, colIndex) => (
                  <td
                    key={colIndex}
                    style={{
                      padding: "8px",
                      textAlign: "center",
                      borderBottom: "1px solid #eee",
                    }}
                  >
                    {siswa.nilaiMapel[mapel] ?? "-"}
                  </td>
                ))}
                <td
                  style={{
                    padding: "8px",
                    textAlign: "center",
                    borderBottom: "1px solid #eee",
                    fontWeight: "bold",
                  }}
                >
                  {siswa.jumlah}
                </td>
                <td
                  style={{
                    padding: "8px",
                    textAlign: "center",
                    borderBottom: "1px solid #eee",
                    fontWeight: "bold",
                  }}
                >
                  {siswa.rataRata}
                </td>
                <td
                  style={{
                    padding: "8px",
                    textAlign: "center",
                    borderBottom: "1px solid #eee",
                    fontWeight: "bold",
                    color: siswa.ranking <= 3 ? "#FFD700" : "#333",
                    fontSize: siswa.ranking <= 3 ? "16px" : "14px",
                  }}
                >
                  {siswa.ranking}
                </td>
                <td
                  style={{
                    padding: "8px",
                    textAlign: "center",
                    borderBottom: "1px solid #eee",
                  }}
                >
                  <button
                    onClick={() => downloadRaporPDF(siswa)}
                    disabled={downloadingId !== null}
                    style={{
                      padding: "6px 12px",
                      backgroundColor:
                        downloadingId !== null
                          ? "#ccc"
                          : downloadingId === siswa.nama
                          ? "#ff9800"
                          : "#e74c3c",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                      cursor:
                        downloadingId !== null ? "not-allowed" : "pointer",
                      fontSize: "12px",
                      fontWeight: "bold",
                      display: "flex",
                      alignItems: "center",
                      gap: "5px",
                      margin: "0 auto",
                      opacity:
                        downloadingId !== null && downloadingId !== siswa.nama
                          ? 0.5
                          : 1,
                    }}
                    onMouseEnter={(e) => {
                      if (downloadingId === null) {
                        (e.target as HTMLButtonElement).style.backgroundColor =
                          "#c0392b";
                      }
                    }}
                    onMouseLeave={(e) => {
                      if (downloadingId === null) {
                        (e.target as HTMLButtonElement).style.backgroundColor =
                          "#e74c3c";
                      }
                    }}
                  >
                    {downloadingId === siswa.nama ? (
                      <>
                        <span
                          style={{
                            display: "inline-block",
                            animation: "spin 1s linear infinite",
                          }}
                        >
                          ⏳
                        </span>
                        Loading...
                      </>
                    ) : (
                      <>📄 PDF</>
                    )}
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

const DataKehadiran = () => {
  const [data, setData] = useState<RowData[]>([]);
  const [changedRows, setChangedRows] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [showFloatingButton, setShowFloatingButton] = useState(false);
  const [floatingButtonPosition, setFloatingButtonPosition] = useState({
    top: 0,
    left: 0,
    visible: true,
  });
  const [activeInput, setActiveInput] = useState<{
    rowIndex: number;
    colIndex: number;
  } | null>(null);
  const [isProcessingClick, setIsProcessingClick] = useState(false);
  const [showStudentPopup, setShowStudentPopup] = useState(false);
  const [selectedStudent, setSelectedStudent] = useState<{
    nama: string;
    kelas: string;
    nisn: string;
  } | null>(null);

  // Fetch data kehadiran
  useEffect(() => {
    const fetchData = async () => {
      setLoading(true);
      setError(null);
      try {
        const response = await fetch(`${endpoint}?sheet=DataKehadiran`);
        if (!response.ok) {
          throw new Error("Network response was not ok");
        }
        const jsonData = await response.json();
        setData(jsonData);
      } catch (err) {
        setError(err instanceof Error ? err.message : "Unknown error");
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, []);

  // Update posisi tombol floating
  useEffect(() => {
    const updateButtonPosition = () => {
      if (showFloatingButton && activeInput) {
        const { rowIndex, colIndex } = activeInput;
        const input = document.getElementById(
          `kehadiran-input-${rowIndex}-${colIndex}`
        ) as HTMLInputElement;

        if (input) {
          const rect = input.getBoundingClientRect();
          const tableContainer = document.getElementById(
            "kehadiran-table-container"
          );

          if (tableContainer) {
            const containerRect = tableContainer.getBoundingClientRect();
            const thead = tableContainer.querySelector("thead");
            const headerHeight = thead ? thead.offsetHeight : 40;

            const inputTopInContainer = rect.top - containerRect.top;
            const inputBottomInContainer = rect.bottom - containerRect.top;

            const isVisibleInContainer =
              inputTopInContainer >= headerHeight &&
              inputBottomInContainer > headerHeight &&
              rect.bottom <= containerRect.bottom &&
              rect.left >= containerRect.left - 100 &&
              rect.right <= window.innerWidth + 100;

            setFloatingButtonPosition({
              top: rect.top + rect.height / 2 - 28,
              left: rect.right + 10,
              visible: isVisibleInContainer,
            });
          }
        }
      }
    };

    const handleScroll = throttle(updateButtonPosition, 16);
    const tableContainer = document.getElementById("kehadiran-table-container");

    if (tableContainer) {
      tableContainer.addEventListener("scroll", handleScroll as any, {
        passive: true,
      });
    }

    window.addEventListener("scroll", handleScroll as any, { passive: true });

    return () => {
      if (tableContainer) {
        tableContainer.removeEventListener("scroll", handleScroll as any);
      }
      window.removeEventListener("scroll", handleScroll as any);
    };
  }, [showFloatingButton, activeInput]);

  const handleInputChange = (
    rowIndex: number,
    header: string,
    value: string
  ) => {
    const updatedData = [...data];
    updatedData[rowIndex + 1][header] = value;
    setData(updatedData);
    setChangedRows((prev) => new Set([...Array.from(prev), rowIndex]));
  };

  const handleKeyDown = (
    e: React.KeyboardEvent<HTMLInputElement>,
    rowIndex: number,
    colIndex: number,
    actualDataLength: number
  ) => {
    if (e.key === "Enter") {
      e.preventDefault();
      const nextRow = rowIndex + 1;
      if (nextRow < actualDataLength) {
        const nextInput = document.getElementById(
          `kehadiran-input-${nextRow}-${colIndex}`
        ) as HTMLInputElement | null;
        if (nextInput) {
          nextInput.focus();
          nextInput.select();
        }
      }
    }
  };

  const handleSaveAll = async () => {
    if (changedRows.size === 0) {
      alert("Tidak ada perubahan untuk disimpan!");
      return;
    }

    setIsSaving(true);

    const headers = [
      "Data1",
      "Data2",
      "Data3",
      "Data4",
      "Data5",
      "Data6",
      "Data7",
      "Data8",
    ];
    const updates: Array<{ rowIndex: number; values: string[] }> = [];

    changedRows.forEach((rowIndex) => {
      const rowData = data[rowIndex + 1];
      const values = headers.map((header) => rowData[header] || "");
      updates.push({
        rowIndex: rowIndex + 3,
        values: values,
      });
    });

    try {
      const requestBody = {
        action: "update_bulk",
        sheetName: "DataKehadiran",
        updates: updates,
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      alert("Semua perubahan berhasil disimpan!");
      setChangedRows(new Set());
    } catch (err) {
      console.error("Error:", err);
      alert(
        "Error menyimpan data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  const handleFloatingArrowClick = () => {
    if (isProcessingClick) return;
    setIsProcessingClick(true);

    if (activeInput) {
      const { rowIndex, colIndex } = activeInput;
      const nextRow = rowIndex + 1;
      if (nextRow < actualData.length) {
        const nextInput = document.getElementById(
          `kehadiran-input-${nextRow}-${colIndex}`
        ) as HTMLInputElement | null;
        if (nextInput) {
          nextInput.focus();
          nextInput.select();
        }
      }
    }

    setTimeout(() => {
      setIsProcessingClick(false);
    }, 300);
  };

  const updateFloatingButtonPosition = (
    element: HTMLInputElement,
    rowIndex: number,
    colIndex: number,
    forceShow: boolean = true
  ) => {
    const rect = element.getBoundingClientRect();

    setFloatingButtonPosition({
      top: rect.top + rect.height / 2 - 28,
      left: rect.right + 10,
      visible: true,
    });
    setActiveInput({ rowIndex, colIndex });

    if (forceShow) {
      setShowFloatingButton(rowIndex < actualData.length - 1);
    }
  };

  if (loading)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>Loading...</div>
    );
  if (error)
    return (
      <div style={{ textAlign: "center", color: "red", padding: "20px" }}>
        Error: {error}
      </div>
    );
  if (data.length === 0)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        No data available
      </div>
    );

  const headers = [
    "Data1",
    "Data2",
    "Data3",
    "Data4",
    "Data5",
    "Data6",
    "Data7",
    "Data8",
  ];
  const displayHeaders = headers.map((header) => data[0][header] || "");
  const actualData = data.slice(1);

  // Data1=No, Data2=Kelas, Data3=Nama, Data4=Hadir, Data5=Alpha, Data6=Izin, Data7=Sakit, Data8=Total Pertemuan
  const readOnlyHeaders = new Set(["Data1", "Data2", "Data3", "Data8"]);
  const editableHeaders = ["Data4", "Data5", "Data6", "Data7"]; // Hadir, Alpha, Izin, Sakit
  const hiddenHeaders = new Set(["Data2", "Data3"]); // Data1 asumsikan NISN/No, Data2=Kelas
  const visibleHeaders = headers.filter((header) => !hiddenHeaders.has(header));
  const visibleDisplayHeaders = visibleHeaders.map(
    (header) => data[0][header] || ""
  );

  return (
    <div style={{ padding: "10px", margin: "0 auto", maxWidth: "100vw" }}>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "15px",
          fontSize: "20px",
        }}
      >
        📋 Data Kehadiran Siswa
      </h1>

      {/* Tombol Save */}
      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <button
          onClick={handleSaveAll}
          disabled={isSaving}
          style={{
            padding: "12px 24px",
            backgroundColor: isSaving ? "#ccc" : "#4CAF50",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: isSaving ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            width: "100%",
            maxWidth: "300px",
          }}
        >
          {isSaving ? "Memproses..." : `Simpan Perubahan (${changedRows.size})`}
        </button>
      </div>

      {/* Table */}
      <div
        id="kehadiran-table-container"
        style={{
          overflowX: "auto",
          overflowY: "auto",
          maxHeight: "calc(100vh - 200px)",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          position: "relative",
        }}
      >
        <table
          style={{
            borderCollapse: "separate",
            borderSpacing: 0,
            minWidth: "100%",
            width: "max-content",
            tableLayout: "fixed",
          }}
        >
          <thead style={{ position: "sticky", top: 0, zIndex: 100 }}>
            <tr style={{ backgroundColor: "#f4f4f4" }}>
              {visibleDisplayHeaders.map((header, index) => {
                const currentHeader = visibleHeaders[index];
                return (
                  <th
                    key={index}
                    style={{
                      padding: "8px 4px",
                      textAlign: "center",
                      borderBottom: "2px solid #ddd",
                      fontWeight: "bold",
                      width: currentHeader === "Data1" ? "200px" : "100px",
                      minWidth: currentHeader === "Data1" ? "200px" : "100px",
                      position: currentHeader === "Data1" ? "sticky" : "static",
                      left: currentHeader === "Data1" ? 0 : "auto",
                      backgroundColor: "#f4f4f4",
                      zIndex: currentHeader === "Data1" ? 2 : 1,
                      boxShadow:
                        currentHeader === "Data1"
                          ? "2px 0 5px rgba(0,0,0,0.1)"
                          : "none",
                      fontSize: "12px",
                    }}
                  >
                    {header}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {actualData.map((row, rowIndex) => (
              <tr
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                }}
              >
                {visibleHeaders.map((header, colIndex) => {
                  const isNama = header === "Data1";
                  const isEditable = editableHeaders.indexOf(header) !== -1; // Sudah dari instruksi sebelumnya
                  return (
                    <td
                      key={colIndex}
                      style={{
                        padding: "4px",
                        borderBottom: "1px solid #eee",
                        position: header === "Data1" ? "sticky" : "static",
                        left: header === "Data1" ? 0 : "auto",
                        backgroundColor:
                          header === "Data1"
                            ? rowIndex % 2 === 0
                              ? "#fff"
                              : "#f9f9f9"
                            : "transparent",
                        zIndex: header === "Data1" ? 1 : 0,
                        boxShadow:
                          header === "Data1"
                            ? "2px 0 5px rgba(0,0,0,0.1)"
                            : "none",
                      }}
                    >
                      {readOnlyHeaders.has(header) || !isEditable ? (
                        <div
                          style={{
                            padding: "4px 2px",
                            color: "#666",
                            fontSize: "12px",
                            textAlign: isNama ? "left" : "center",
                            cursor: header === "Data1" ? "pointer" : "default",
                          }}
                          onClick={() => {
                            if (header === "Data1") {
                              setSelectedStudent({
                                nama: row.Data3 || "",
                                kelas: row.Data2 || "N/A",
                                nisn: row.Data1 || "N/A",
                              });
                              setShowStudentPopup(true);
                            }
                          }}
                        >
                          {row[header] || ""}
                        </div>
                      ) : (
                        <input
                          id={`kehadiran-input-${rowIndex}-${colIndex}`}
                          type="text"
                          inputMode="decimal"
                          pattern="[0-9]*"
                          value={row[header] || ""}
                          onChange={(e) =>
                            handleInputChange(rowIndex, header, e.target.value)
                          }
                          onKeyDown={(e) =>
                            handleKeyDown(
                              e,
                              rowIndex,
                              colIndex,
                              actualData.length
                            )
                          }
                          onFocus={(e) => {
                            e.target.select();
                            updateFloatingButtonPosition(
                              e.target,
                              rowIndex,
                              colIndex,
                              true
                            );
                          }}
                          onBlur={(e) => {
                            const relatedTarget =
                              e.relatedTarget as HTMLElement;
                            if (
                              !relatedTarget ||
                              relatedTarget.tagName !== "BUTTON"
                            ) {
                              setTimeout(() => {
                                const activeElement = document.activeElement;
                                const isInputFocused =
                                  activeElement?.tagName === "INPUT" &&
                                  activeElement?.id.startsWith(
                                    "kehadiran-input-"
                                  );
                                if (!isInputFocused) {
                                  setShowFloatingButton(false);
                                }
                              }, 150);
                            }
                          }}
                          style={{
                            width: "100%",
                            padding: "4px 2px",
                            border: "1px solid #ddd",
                            borderRadius: "3px",
                            boxSizing: "border-box",
                            backgroundColor: "white",
                            fontSize: "12px",
                            textAlign: "center",
                          }}
                        />
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Floating Arrow Button */}
      {showFloatingButton && (
        <button
          onClick={handleFloatingArrowClick}
          style={{
            position: "fixed",
            top: `${floatingButtonPosition.top}px`,
            left: `${floatingButtonPosition.left}px`,
            width: "56px",
            height: "56px",
            borderRadius: "50%",
            backgroundColor: "#4CAF50",
            color: "white",
            border: "none",
            cursor: "pointer",
            boxShadow: "0 4px 12px rgba(0,0,0,0.4)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: "28px",
            fontWeight: "bold",
            zIndex: 1001,
            transition: "all 0.2s ease",
            pointerEvents: floatingButtonPosition.visible ? "auto" : "none",
            opacity: floatingButtonPosition.visible ? 1 : 0,
            visibility: floatingButtonPosition.visible ? "visible" : "hidden",
          }}
        >
          ↓
        </button>
      )}
      {showStudentPopup && selectedStudent && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0, 0, 0, 0.5)",
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            zIndex: 1000,
          }}
          onClick={() => setShowStudentPopup(false)}
        >
          <div
            style={{
              backgroundColor: "white",
              borderRadius: "8px",
              padding: "20px",
              maxWidth: "400px",
              width: "90%",
              boxShadow: "0 4px 20px rgba(0,0,0,0.3)",
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: "15px",
                borderBottom: "2px solid #2196F3",
                paddingBottom: "10px",
              }}
            >
              <h2 style={{ margin: 0, color: "#333", fontSize: "18px" }}>
                Detail Siswa: {selectedStudent.nama}
              </h2>
              <button
                onClick={() => setShowStudentPopup(false)}
                style={{
                  backgroundColor: "#f44336",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  padding: "8px 16px",
                  cursor: "pointer",
                  fontSize: "14px",
                  fontWeight: "bold",
                }}
              >
                Tutup
              </button>
            </div>
            <div style={{ marginBottom: "15px" }}>
              <strong style={{ color: "#4CAF50" }}>Nama:</strong>{" "}
              <span style={{ color: "#333" }}>{selectedStudent.nama}</span>
            </div>
            <div style={{ marginBottom: "15px" }}>
              <strong style={{ color: "#4CAF50" }}>Kelas:</strong>{" "}
              <span style={{ color: "#333" }}>{selectedStudent.kelas}</span>
            </div>
            <div style={{ marginBottom: "15px" }}>
              <strong style={{ color: "#4CAF50" }}>NISN:</strong>{" "}
              <span style={{ color: "#333" }}>{selectedStudent.nisn}</span>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

const InputTP = () => {
  const [data, setData] = useState<RowData[]>([]);
  const [changedRows, setChangedRows] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [isAddingNew, setIsAddingNew] = useState(false);
  const [newTP, setNewTP] = useState({
    mapel: "",
    tp: "",
    rincian: "",
    bab: "",
    semester: "",
    kelas: "",
  });
  const [availableMapel, setAvailableMapel] = useState<string[]>([]);

  // Fetch data dari DataTP
  useEffect(() => {
    const fetchData = async () => {
      setLoading(true);
      setError(null);
      try {
        const response = await fetch(`${endpoint}?sheet=DataTP`);
        if (!response.ok) {
          throw new Error("Network response was not ok");
        }
        const jsonData = await response.json();
        setData(jsonData);

        // Extract unique mapel untuk dropdown
        const actualData = jsonData.slice(1);
        const mapelSet = new Set<string>();
        actualData.forEach((row: any) => {
          if (row.Data1) {
            mapelSet.add(row.Data1);
          }
        });
        setAvailableMapel(Array.from(mapelSet).sort());
      } catch (err) {
        setError(err instanceof Error ? err.message : "Unknown error");
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, []);

  // PERBAIKAN: Function untuk auto-generate TP berdasarkan MAPEL dan BAB
  const getNextTP = (mapel: string, bab: string): string => {
    if (!mapel || !bab) return "";

    const actualData = data.slice(1);

    // Filter data berdasarkan MAPEL saja
    const filteredData = actualData.filter((row: any) => row.Data1 === mapel);

    // Jika belum ada data untuk mapel ini, mulai dari [bab].1
    if (filteredData.length === 0) {
      return `${bab}.1`;
    }

    // Ambil semua TP dari mapel yang sama yang dimulai dengan bab yang dipilih
    const tpList = filteredData
      .map((row: any) => row.Data2)
      .filter((tp: string) => {
        if (!tp || typeof tp !== "string") return false;
        // Cek apakah TP dimulai dengan bab yang sama
        const tpBab = tp.split(".")[0];
        return tpBab === bab;
      })
      .sort((a: string, b: string) => {
        const [aMain, aSub] = a.split(".").map(Number);
        const [bMain, bSub] = b.split(".").map(Number);
        if (aMain !== bMain) return aMain - bMain;
        return aSub - bSub;
      });

    // Jika tidak ada TP dengan bab ini, mulai dari [bab].1
    if (tpList.length === 0) {
      return `${bab}.1`;
    }

    // Ambil TP terakhir dan increment sub number
    const lastTP = tpList[tpList.length - 1];
    const [mainNum, subNum] = lastTP.split(".").map(Number);

    const nextTP = `${mainNum}.${subNum + 1}`;
    return nextTP;
  };

  // PERBAIKAN: Handle perubahan Mapel
  const handleMapelChange = (selectedMapel: string) => {
    setNewTP((prev) => {
      const actualData = data.slice(1);
      const mapelData = actualData.find(
        (row: any) => row.Data1 === selectedMapel
      );

      return {
        ...prev,
        mapel: selectedMapel,
        tp: "", // Reset TP, tunggu BAB diisi dulu
        bab: "",
        semester: mapelData?.Data5 || "",
        kelas: mapelData?.Data6 || "",
      };
    });
  };

  // Handle perubahan BAB
  const handleBabChange = (selectedBab: string) => {
    setNewTP((prev) => ({
      ...prev,
      bab: selectedBab,
      tp: prev.mapel && selectedBab ? getNextTP(prev.mapel, selectedBab) : "",
    }));
  };

  const handleInputChange = (
    rowIndex: number,
    header: string,
    value: string
  ) => {
    const updatedData = [...data];
    updatedData[rowIndex + 1][header] = value;
    setData(updatedData);
    setChangedRows((prev) => new Set([...Array.from(prev), rowIndex]));
  };

  const handleSaveAll = async () => {
    if (changedRows.size === 0) {
      alert("Tidak ada perubahan untuk disimpan!");
      return;
    }

    setIsSaving(true);

    const headers = ["Data1", "Data2", "Data3", "Data4", "Data5", "Data6"];
    const updates: Array<{ rowIndex: number; values: string[] }> = [];

    changedRows.forEach((rowIndex) => {
      const rowData = data[rowIndex + 1];
      const values = headers.map((header) => rowData[header] || "");
      updates.push({
        rowIndex: rowIndex + 3,
        values: values,
      });
    });

    try {
      const requestBody = {
        action: "update_bulk",
        sheetName: "DataTP",
        updates: updates,
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      alert("Semua perubahan berhasil disimpan!");
      setChangedRows(new Set());

      const refreshResponse = await fetch(`${endpoint}?sheet=DataTP`);
      const refreshedData = await refreshResponse.json();
      setData(refreshedData);
    } catch (err) {
      console.error("Error:", err);
      alert(
        "Error menyimpan data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  const handleAddNew = async () => {
    if (
      !newTP.mapel ||
      !newTP.tp ||
      !newTP.rincian ||
      !newTP.bab ||
      !newTP.semester ||
      !newTP.kelas
    ) {
      alert("⚠️ Semua field wajib diisi!");
      return;
    }

    setIsSaving(true);

    try {
      const requestBody = {
        action: "add_tp",
        data: newTP,
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      alert("✅ Data TP baru berhasil ditambahkan!");

      setNewTP({
        mapel: "",
        tp: "",
        rincian: "",
        bab: "",
        semester: "",
        kelas: "",
      });
      setIsAddingNew(false);

      const refreshResponse = await fetch(`${endpoint}?sheet=DataTP`);
      const refreshedData = await refreshResponse.json();
      setData(refreshedData);

      // Update available mapel
      const actualData = refreshedData.slice(1);
      const mapelSet = new Set<string>();
      actualData.forEach((row: any) => {
        if (row.Data1) {
          mapelSet.add(row.Data1);
        }
      });
      setAvailableMapel(Array.from(mapelSet).sort());
    } catch (err) {
      console.error("Error:", err);
      alert(
        "Error menambah data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  if (loading)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>Loading...</div>
    );
  if (error)
    return (
      <div style={{ textAlign: "center", color: "red", padding: "20px" }}>
        Error: {error}
      </div>
    );
  if (data.length === 0)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        No data available
      </div>
    );

  const headers = ["Data1", "Data2", "Data3", "Data4", "Data5", "Data6"];
  const displayHeaders = headers.map((header) => data[0][header] || "");
  const actualData = data.slice(1);

  return (
    <div style={{ padding: "10px", margin: "0 auto", maxWidth: "100vw" }}>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "15px",
          fontSize: "20px",
        }}
      >
        📚 Data Tujuan Pembelajaran (TP)
      </h1>

      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <button
          onClick={() => setIsAddingNew(!isAddingNew)}
          style={{
            padding: "12px 24px",
            backgroundColor: isAddingNew ? "#f44336" : "#2196F3",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            marginRight: "10px",
          }}
        >
          {isAddingNew ? "❌ Batal Tambah" : "➕ Tambah Data Baru"}
        </button>

        <button
          onClick={handleSaveAll}
          disabled={isSaving || changedRows.size === 0}
          style={{
            padding: "12px 24px",
            backgroundColor:
              isSaving || changedRows.size === 0 ? "#ccc" : "#4CAF50",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor:
              isSaving || changedRows.size === 0 ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
          }}
        >
          {isSaving
            ? "Memproses..."
            : `💾 Simpan Perubahan (${changedRows.size})`}
        </button>
      </div>

      {/* FORM YANG DIPERBARUI - Urutan: Mapel > BAB > TP (auto) */}
      {isAddingNew && (
        <div
          style={{
            backgroundColor: "#f0f8ff",
            padding: "20px",
            borderRadius: "8px",
            marginBottom: "20px",
            boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          }}
        >
          <h3 style={{ marginBottom: "15px", color: "#2196F3" }}>
            Form Tambah TP Baru
          </h3>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gap: "15px",
            }}
          >
            {/* LANGKAH 1: MAPEL */}
            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                1️⃣ Mapel:
              </label>
              <select
                value={newTP.mapel}
                onChange={(e) => {
                  const value = e.target.value;
                  if (value === "__custom__") {
                    const customMapel = prompt("Masukkan nama Mapel baru:");
                    if (customMapel) {
                      setNewTP({
                        mapel: customMapel.toUpperCase(),
                        tp: "",
                        rincian: "",
                        bab: "",
                        semester: "",
                        kelas: "",
                      });
                    }
                  } else {
                    handleMapelChange(value);
                  }
                }}
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                }}
              >
                <option value="">-- Pilih Mapel --</option>
                {availableMapel.map((mapel, index) => (
                  <option key={index} value={mapel}>
                    {mapel}
                  </option>
                ))}
                <option value="__custom__">➕ Tambah Mapel Baru...</option>
              </select>
            </div>

            {/* LANGKAH 2: BAB */}
            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                2️⃣ BAB:
              </label>
              <input
                type="text"
                value={newTP.bab}
                onChange={(e) => handleBabChange(e.target.value)}
                placeholder="Contoh: 1, 2, 3, dst"
                disabled={!newTP.mapel}
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  backgroundColor: !newTP.mapel ? "#f5f5f5" : "white",
                  cursor: !newTP.mapel ? "not-allowed" : "text",
                }}
              />
              {!newTP.mapel && (
                <small style={{ color: "#999", fontSize: "11px" }}>
                  Pilih Mapel terlebih dahulu
                </small>
              )}
            </div>

            {/* LANGKAH 3: TP (AUTO) */}
            <div style={{ gridColumn: "1 / -1" }}>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                3️⃣ TP:
                {newTP.mapel && newTP.bab && (
                  <span
                    style={{
                      color: "#4CAF50",
                      fontSize: "12px",
                      marginLeft: "8px",
                    }}
                  >
                    ✓ Otomatis terisi berdasarkan BAB
                  </span>
                )}
              </label>
              <input
                type="text"
                value={newTP.tp}
                readOnly
                placeholder="TP akan terisi otomatis setelah Mapel & BAB diisi"
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "2px solid " + (newTP.tp ? "#4CAF50" : "#ddd"),
                  borderRadius: "4px",
                  backgroundColor: newTP.tp ? "#e8f5e9" : "#f5f5f5",
                  fontSize: "16px",
                  fontWeight: "bold",
                  color: newTP.tp ? "#2e7d32" : "#999",
                  textAlign: "center",
                  cursor: "not-allowed",
                }}
              />
            </div>

            {/* RINCIAN TP */}
            <div style={{ gridColumn: "1 / -1" }}>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                Rincian TP:
              </label>
              <textarea
                value={newTP.rincian}
                onChange={(e) =>
                  setNewTP({ ...newTP, rincian: e.target.value })
                }
                placeholder="Jelaskan rincian tujuan pembelajaran..."
                rows={3}
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  resize: "vertical",
                }}
              />
            </div>

            {/* SEMESTER & KELAS */}
            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                Semester:
                {newTP.mapel && (
                  <span
                    style={{
                      color: "#2196F3",
                      fontSize: "12px",
                      marginLeft: "5px",
                    }}
                  >
                    (Auto-fill)
                  </span>
                )}
              </label>
              <input
                type="text"
                value={newTP.semester}
                onChange={(e) =>
                  setNewTP({ ...newTP, semester: e.target.value })
                }
                placeholder="Contoh: 1"
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  backgroundColor: newTP.mapel ? "#f0f8ff" : "white",
                }}
              />
            </div>
            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                Kelas:
                {newTP.mapel && (
                  <span
                    style={{
                      color: "#2196F3",
                      fontSize: "12px",
                      marginLeft: "5px",
                    }}
                  >
                    (Auto-fill)
                  </span>
                )}
              </label>
              <input
                type="text"
                value={newTP.kelas}
                onChange={(e) => setNewTP({ ...newTP, kelas: e.target.value })}
                placeholder="Contoh: 6"
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  backgroundColor: newTP.mapel ? "#f0f8ff" : "white",
                }}
              />
            </div>
          </div>

          {/* Info Helper */}
          {newTP.mapel && newTP.bab && newTP.tp && (
            <div
              style={{
                marginTop: "15px",
                padding: "12px",
                backgroundColor: "#e8f5e9",
                border: "1px solid #4CAF50",
                borderRadius: "4px",
              }}
            >
              <strong style={{ color: "#2e7d32" }}>✓ Siap disimpan:</strong>
              <div
                style={{ marginTop: "5px", fontSize: "14px", color: "#333" }}
              >
                {newTP.mapel} - BAB {newTP.bab} - TP {newTP.tp}
              </div>
            </div>
          )}

          <button
            onClick={handleAddNew}
            disabled={isSaving}
            style={{
              marginTop: "15px",
              padding: "12px 24px",
              backgroundColor: isSaving ? "#ccc" : "#4CAF50",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: isSaving ? "not-allowed" : "pointer",
              fontWeight: "bold",
              fontSize: "16px",
            }}
          >
            {isSaving ? "Menyimpan..." : "💾 Simpan Data Baru"}
          </button>
        </div>
      )}

      {/* Tabel Data TP */}
      <div
        style={{
          overflowX: "auto",
          overflowY: "auto",
          maxHeight: "calc(100vh - 300px)",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          position: "relative",
        }}
      >
        <table
          style={{
            borderCollapse: "separate",
            borderSpacing: 0,
            minWidth: "100%",
            width: "max-content",
          }}
        >
          <thead style={{ position: "sticky", top: 0, zIndex: 100 }}>
            <tr style={{ backgroundColor: "#f4f4f4" }}>
              {displayHeaders.map((header, index) => (
                <th
                  key={index}
                  style={{
                    padding: "8px",
                    textAlign: "center",
                    borderBottom: "2px solid #ddd",
                    fontWeight: "bold",
                    minWidth: index === 2 ? "300px" : "120px",
                    backgroundColor: "#f4f4f4",
                    fontSize: "12px",
                  }}
                >
                  {header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {actualData.map((row, rowIndex) => (
              <tr
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                }}
              >
                {headers.map((header, colIndex) => (
                  <td
                    key={colIndex}
                    style={{
                      padding: "8px",
                      borderBottom: "1px solid #eee",
                    }}
                  >
                    <input
                      type="text"
                      value={row[header] || ""}
                      onChange={(e) =>
                        handleInputChange(rowIndex, header, e.target.value)
                      }
                      style={{
                        width: "100%",
                        padding: "6px",
                        border: "1px solid #ddd",
                        borderRadius: "3px",
                        fontSize: "12px",
                      }}
                    />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

const DataMapel = () => {
  const [data, setData] = useState<RowData[]>([]);
  const [changedRows, setChangedRows] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [isAddingNew, setIsAddingNew] = useState(false);
  const [newMapel, setNewMapel] = useState("");

  useEffect(() => {
    const fetchData = async () => {
      setLoading(true);
      setError(null);
      try {
        const response = await fetch(`${endpoint}?sheet=DataMapel`);
        if (!response.ok) throw new Error("Network response was not ok");
        const jsonData = await response.json();
        setData(jsonData);
      } catch (err) {
        setError(err instanceof Error ? err.message : "Unknown error");
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, []);

  const handleInputChange = (
    rowIndex: number,
    header: string,
    value: string
  ) => {
    const updatedData = [...data];
    updatedData[rowIndex + 1][header] = value;
    setData(updatedData);
    setChangedRows((prev) => new Set([...Array.from(prev), rowIndex]));
  };

  const handleSaveAll = async () => {
    if (changedRows.size === 0) {
      alert("Tidak ada perubahan untuk disimpan!");
      return;
    }

    setIsSaving(true);
    const headers = ["Data1"];
    const updates: Array<{ rowIndex: number; values: string[] }> = [];

    changedRows.forEach((rowIndex) => {
      const rowData = data[rowIndex + 1];
      const values = headers.map((header) => rowData[header] || "");
      updates.push({ rowIndex: rowIndex + 3, values: values });
    });

    try {
      const requestBody = {
        action: "update_bulk",
        sheetName: "DataMapel",
        updates: updates,
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);

      alert("Semua perubahan berhasil disimpan!");
      setChangedRows(new Set());

      const refreshResponse = await fetch(`${endpoint}?sheet=DataMapel`);
      const refreshedData = await refreshResponse.json();
      setData(refreshedData);
    } catch (err) {
      console.error("Error:", err);
      alert(
        "Error menyimpan data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  const handleAddNew = async () => {
    if (!newMapel.trim()) {
      alert("⚠️ Nama Mata Pelajaran wajib diisi!");
      return;
    }

    setIsSaving(true);

    try {
      const requestBody = {
        action: "add_mapel",
        mapel: newMapel.toUpperCase().trim(),
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);

      alert("✅ Mata Pelajaran baru berhasil ditambahkan!");
      setNewMapel("");
      setIsAddingNew(false);

      const refreshResponse = await fetch(`${endpoint}?sheet=DataMapel`);
      const refreshedData = await refreshResponse.json();
      setData(refreshedData);
    } catch (err) {
      console.error("Error:", err);
      alert(
        "Error menambah data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  if (loading)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>Loading...</div>
    );
  if (error)
    return (
      <div style={{ textAlign: "center", color: "red", padding: "20px" }}>
        Error: {error}
      </div>
    );
  if (data.length === 0)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        No data available
      </div>
    );

  const actualData = data.slice(1);

  return (
    <div style={{ padding: "10px", margin: "0 auto", maxWidth: "800px" }}>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "15px",
          fontSize: "20px",
        }}
      >
        📚 Data Mata Pelajaran
      </h1>

      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <button
          onClick={() => setIsAddingNew(!isAddingNew)}
          style={{
            padding: "12px 24px",
            backgroundColor: isAddingNew ? "#f44336" : "#2196F3",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            marginRight: "10px",
          }}
        >
          {isAddingNew ? "❌ Batal Tambah" : "➕ Tambah Mata Pelajaran Baru"}
        </button>

        <button
          onClick={handleSaveAll}
          disabled={isSaving || changedRows.size === 0}
          style={{
            padding: "12px 24px",
            backgroundColor:
              isSaving || changedRows.size === 0 ? "#ccc" : "#4CAF50",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor:
              isSaving || changedRows.size === 0 ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
          }}
        >
          {isSaving
            ? "Memproses..."
            : `💾 Simpan Perubahan (${changedRows.size})`}
        </button>
      </div>

      {isAddingNew && (
        <div
          style={{
            backgroundColor: "#f0f8ff",
            padding: "20px",
            borderRadius: "8px",
            marginBottom: "20px",
            boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          }}
        >
          <h3 style={{ marginBottom: "15px", color: "#2196F3" }}>
            Form Tambah Mata Pelajaran Baru
          </h3>
          <input
            type="text"
            value={newMapel}
            onChange={(e) => setNewMapel(e.target.value)}
            placeholder="Contoh: BAHASA INGGRIS"
            style={{
              width: "100%",
              padding: "12px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              fontSize: "16px",
              marginBottom: "15px",
            }}
          />
          <button
            onClick={handleAddNew}
            disabled={isSaving}
            style={{
              padding: "12px 24px",
              backgroundColor: isSaving ? "#ccc" : "#4CAF50",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: isSaving ? "not-allowed" : "pointer",
              fontWeight: "bold",
              fontSize: "16px",
            }}
          >
            {isSaving ? "Menyimpan..." : "💾 Simpan Mata Pelajaran Baru"}
          </button>
        </div>
      )}

      <div
        style={{
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          overflow: "hidden",
        }}
      >
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead style={{ backgroundColor: "#f4f4f4" }}>
            <tr>
              <th
                style={{
                  padding: "12px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  width: "80px",
                }}
              >
                No
              </th>
              <th
                style={{
                  padding: "12px",
                  textAlign: "left",
                  borderBottom: "2px solid #ddd",
                }}
              >
                Nama Mata Pelajaran
              </th>
            </tr>
          </thead>
          <tbody>
            {actualData.map((row, rowIndex) => (
              <tr
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                }}
              >
                <td
                  style={{
                    padding: "12px",
                    textAlign: "center",
                    borderBottom: "1px solid #eee",
                  }}
                >
                  {rowIndex + 1}
                </td>
                <td style={{ padding: "8px", borderBottom: "1px solid #eee" }}>
                  <input
                    type="text"
                    value={row.Data1 || ""}
                    onChange={(e) =>
                      handleInputChange(rowIndex, "Data1", e.target.value)
                    }
                    style={{
                      width: "100%",
                      padding: "8px",
                      border: "1px solid #ddd",
                      borderRadius: "4px",
                      fontSize: "14px",
                    }}
                  />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

const DataKokurikuler = () => {
  const [data, setData] = useState<RowData[]>([]);
  const [changedRows, setChangedRows] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);

  useEffect(() => {
    const fetchData = async () => {
      setLoading(true);
      setError(null);
      try {
        const response = await fetch(`${endpoint}?sheet=DataKokurikuler`);
        if (!response.ok) throw new Error("Network response was not ok");
        const jsonData = await response.json();
        setData(jsonData);
      } catch (err) {
        setError(err instanceof Error ? err.message : "Unknown error");
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, []);

  const handleInputChange = (
    rowIndex: number,
    header: string,
    value: string
  ) => {
    const updatedData = [...data];
    updatedData[rowIndex + 1][header] = value;
    setData(updatedData);
    setChangedRows((prev) => new Set([...Array.from(prev), rowIndex]));
  };

  const handleSaveAll = async () => {
    if (changedRows.size === 0) {
      alert("Tidak ada perubahan untuk disimpan!");
      return;
    }

    setIsSaving(true);

    // Data10 (Deskripsi) tidak disertakan karena berisi formula
    const headers = [
      "Data1",
      "Data2",
      "Data3",
      "Data4",
      "Data5",
      "Data6",
      "Data7",
      "Data8",
      "Data9",
    ];
    const updates: Array<{ rowIndex: number; values: string[] }> = [];

    changedRows.forEach((rowIndex) => {
      const rowData = data[rowIndex + 1];
      const values = headers.map((header) => rowData[header] || "");
      updates.push({
        rowIndex: rowIndex + 3,
        values: values,
      });
    });

    try {
      const requestBody = {
        action: "update_bulk",
        sheetName: "DataKokurikuler",
        updates: updates,
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      alert("Semua perubahan berhasil disimpan!");
      setChangedRows(new Set());
    } catch (err) {
      console.error("Error:", err);
      alert(
        "Error menyimpan data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  if (loading)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>Loading...</div>
    );
  if (error)
    return (
      <div style={{ textAlign: "center", color: "red", padding: "20px" }}>
        Error: {error}
      </div>
    );
  if (data.length === 0)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        No data available
      </div>
    );

  const headers = [
    "Data1",
    "Data2",
    "Data3",
    "Data4",
    "Data5",
    "Data6",
    "Data7",
    "Data8",
    "Data9",
    "Data10",
  ];
  const displayHeaders = headers.map((header) => data[0][header] || "");
  const actualData = data.slice(1);

  // Data1=nama siswa (frozen, readonly)
  // Data2-Data9 = kolom nilai kokurikuler (editable)
  // Data10 = Deskripsi (readonly)
  const readOnlyHeaders = new Set(["Data1", "Data10"]);

  return (
    <div style={{ padding: "10px", margin: "0 auto", maxWidth: "100vw" }}>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "15px",
          fontSize: "20px",
        }}
      >
        🌟 Data Kokurikuler
      </h1>

      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <button
          onClick={handleSaveAll}
          disabled={isSaving}
          style={{
            padding: "12px 24px",
            backgroundColor: isSaving ? "#ccc" : "#4CAF50",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: isSaving ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            width: "100%",
            maxWidth: "300px",
          }}
        >
          {isSaving ? "Memproses..." : `Simpan Perubahan (${changedRows.size})`}
        </button>
      </div>

      <div
        style={{
          overflowX: "auto",
          overflowY: "auto",
          maxHeight: "calc(100vh - 200px)",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          position: "relative",
        }}
      >
        <table
          style={{
            borderCollapse: "separate",
            borderSpacing: 0,
            minWidth: "100%",
            width: "max-content",
            tableLayout: "fixed",
          }}
        >
          <thead style={{ position: "sticky", top: 0, zIndex: 100 }}>
            <tr style={{ backgroundColor: "#f4f4f4" }}>
              {displayHeaders.map((header, index) => {
                const currentHeader = headers[index];
                const isNameColumn = currentHeader === "Data1";
                return (
                  <th
                    key={index}
                    style={{
                      padding: "8px 4px",
                      textAlign: "center",
                      borderBottom: "2px solid #ddd",
                      fontWeight: "bold",
                      width: isNameColumn ? "200px" : "120px",
                      minWidth: isNameColumn ? "200px" : "120px",
                      position: isNameColumn ? "sticky" : "static",
                      left: isNameColumn ? 0 : "auto",
                      backgroundColor: "#f4f4f4",
                      zIndex: isNameColumn ? 2 : 1,
                      boxShadow: isNameColumn
                        ? "2px 0 5px rgba(0,0,0,0.1)"
                        : "none",
                      fontSize: "12px",
                    }}
                  >
                    {header}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {actualData.map((row, rowIndex) => (
              <tr
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                }}
              >
                {headers.map((header, colIndex) => {
                  const isNameColumn = header === "Data1";
                  const isReadOnly = readOnlyHeaders.has(header);
                  return (
                    <td
                      key={colIndex}
                      style={{
                        padding: "4px",
                        borderBottom: "1px solid #eee",
                        position: isNameColumn ? "sticky" : "static",
                        left: isNameColumn ? 0 : "auto",
                        backgroundColor: isNameColumn
                          ? rowIndex % 2 === 0
                            ? "#fff"
                            : "#f9f9f9"
                          : "transparent",
                        zIndex: isNameColumn ? 1 : 0,
                        boxShadow: isNameColumn
                          ? "2px 0 5px rgba(0,0,0,0.1)"
                          : "none",
                      }}
                    >
                      {isReadOnly ? (
                        <div
                          style={{
                            padding: "4px 2px",
                            color: "#666",
                            fontSize: "12px",
                            textAlign: isNameColumn ? "left" : "center",
                          }}
                        >
                          {row[header] || ""}
                        </div>
                      ) : (
                        <input
                          type="text"
                          inputMode="decimal"
                          value={row[header] || ""}
                          onChange={(e) =>
                            handleInputChange(rowIndex, header, e.target.value)
                          }
                          style={{
                            width: "100%",
                            padding: "4px 2px",
                            border: "1px solid #ddd",
                            borderRadius: "3px",
                            boxSizing: "border-box",
                            backgroundColor: "white",
                            fontSize: "12px",
                            textAlign: "center",
                          }}
                        />
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

const DataEkstrakurikuler = () => {
  const [data, setData] = useState<RowData[]>([]);
  const [changedRows, setChangedRows] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);

  useEffect(() => {
    const fetchData = async () => {
      setLoading(true);
      setError(null);
      try {
        const response = await fetch(`${endpoint}?sheet=DataEkstrakurikuler`);
        if (!response.ok) throw new Error("Network response was not ok");
        const jsonData = await response.json();
        setData(jsonData);
      } catch (err) {
        setError(err instanceof Error ? err.message : "Unknown error");
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, []);

  const handleInputChange = (
    rowIndex: number,
    header: string,
    value: string
  ) => {
    const updatedData = [...data];
    updatedData[rowIndex + 1][header] = value;
    setData(updatedData);
    setChangedRows((prev) => new Set([...Array.from(prev), rowIndex]));
  };

  const handleSaveAll = async () => {
    if (changedRows.size === 0) {
      alert("Tidak ada perubahan untuk disimpan!");
      return;
    }

    setIsSaving(true);

    const headers = ["Data1", "Data2", "Data3", "Data4", "Data5", "Data6"];
    const updates: Array<{ rowIndex: number; values: string[] }> = [];

    changedRows.forEach((rowIndex) => {
      const rowData = data[rowIndex + 1];
      const values = headers.map((header) => rowData[header] || "");
      updates.push({
        rowIndex: rowIndex + 3,
        values: values,
      });
    });

    try {
      const requestBody = {
        action: "update_bulk",
        sheetName: "DataEkstrakurikuler",
        updates: updates,
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      alert("Semua perubahan berhasil disimpan!");
      setChangedRows(new Set());
    } catch (err) {
      console.error("Error:", err);
      alert(
        "Error menyimpan data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  if (loading)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>Loading...</div>
    );
  if (error)
    return (
      <div style={{ textAlign: "center", color: "red", padding: "20px" }}>
        Error: {error}
      </div>
    );
  if (data.length === 0)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        No data available
      </div>
    );

  const headers = ["Data1", "Data2", "Data3", "Data4", "Data5", "Data6"];
  const displayHeaders = headers.map((header) => data[0][header] || "");
  const actualData = data.slice(1);

  const readOnlyHeaders = new Set(["Data1"]);

  return (
    <div style={{ padding: "10px", margin: "0 auto", maxWidth: "100vw" }}>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "15px",
          fontSize: "20px",
        }}
      >
        🎯 Data Ekstrakurikuler
      </h1>

      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <button
          onClick={handleSaveAll}
          disabled={isSaving}
          style={{
            padding: "12px 24px",
            backgroundColor: isSaving ? "#ccc" : "#4CAF50",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: isSaving ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            width: "100%",
            maxWidth: "300px",
          }}
        >
          {isSaving ? "Memproses..." : `Simpan Perubahan (${changedRows.size})`}
        </button>
      </div>

      <div
        style={{
          overflowX: "auto",
          overflowY: "auto",
          maxHeight: "calc(100vh - 200px)",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          position: "relative",
        }}
      >
        <table
          style={{
            borderCollapse: "separate",
            borderSpacing: 0,
            minWidth: "100%",
            width: "max-content",
            tableLayout: "fixed",
          }}
        >
          <thead style={{ position: "sticky", top: 0, zIndex: 100 }}>
            <tr style={{ backgroundColor: "#f4f4f4" }}>
              {displayHeaders.map((header, index) => {
                const currentHeader = headers[index];
                const isNameColumn = currentHeader === "Data1";
                return (
                  <th
                    key={index}
                    style={{
                      padding: "8px 4px",
                      textAlign: "center",
                      borderBottom: "2px solid #ddd",
                      fontWeight: "bold",
                      width: isNameColumn ? "200px" : "120px",
                      minWidth: isNameColumn ? "200px" : "120px",
                      position: isNameColumn ? "sticky" : "static",
                      left: isNameColumn ? 0 : "auto",
                      backgroundColor: "#f4f4f4",
                      zIndex: isNameColumn ? 2 : 1,
                      boxShadow: isNameColumn
                        ? "2px 0 5px rgba(0,0,0,0.1)"
                        : "none",
                      fontSize: "12px",
                    }}
                  >
                    {header}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {actualData.map((row, rowIndex) => (
              <tr
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                }}
              >
                {headers.map((header, colIndex) => {
                  const isNameColumn = header === "Data1";
                  const isReadOnly = readOnlyHeaders.has(header);
                  return (
                    <td
                      key={colIndex}
                      style={{
                        padding: "4px",
                        borderBottom: "1px solid #eee",
                        position: isNameColumn ? "sticky" : "static",
                        left: isNameColumn ? 0 : "auto",
                        backgroundColor: isNameColumn
                          ? rowIndex % 2 === 0
                            ? "#fff"
                            : "#f9f9f9"
                          : "transparent",
                        zIndex: isNameColumn ? 1 : 0,
                        boxShadow: isNameColumn
                          ? "2px 0 5px rgba(0,0,0,0.1)"
                          : "none",
                      }}
                    >
                      {isReadOnly ? (
                        <div
                          style={{
                            padding: "4px 2px",
                            color: "#666",
                            fontSize: "12px",
                            textAlign: isNameColumn ? "left" : "center",
                          }}
                        >
                          {row[header] || ""}
                        </div>
                      ) : (
                        <input
                          type="text"
                          inputMode="decimal"
                          value={row[header] || ""}
                          onChange={(e) =>
                            handleInputChange(rowIndex, header, e.target.value)
                          }
                          style={{
                            width: "100%",
                            padding: "4px 2px",
                            border: "1px solid #ddd",
                            borderRadius: "3px",
                            boxSizing: "border-box",
                            backgroundColor: "white",
                            fontSize: "12px",
                            textAlign: "center",
                          }}
                        />
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

const DataSiswa = () => {
  const [data, setData] = useState<RowData[]>([]);
  const [changedRows, setChangedRows] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [isAddingNew, setIsAddingNew] = useState(false);
  const [newSiswa, setNewSiswa] = useState({
    nama: "",
    kelas: "",
    nis: "",
    nisn: "",
    namaOrtu: "",
  });

  useEffect(() => {
    const fetchData = async () => {
      setLoading(true);
      setError(null);
      try {
        const response = await fetch(`${endpoint}?sheet=DataSiswa`);
        if (!response.ok) throw new Error("Network response was not ok");
        const jsonData = await response.json();
        setData(jsonData);
      } catch (err) {
        setError(err instanceof Error ? err.message : "Unknown error");
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, []);

  const handleInputChange = (
    rowIndex: number,
    header: string,
    value: string
  ) => {
    const updatedData = [...data];
    updatedData[rowIndex + 1][header] = value;
    setData(updatedData);
    setChangedRows((prev) => new Set([...Array.from(prev), rowIndex]));
  };

  const handleSaveAll = async () => {
    if (changedRows.size === 0) {
      alert("Tidak ada perubahan untuk disimpan!");
      return;
    }

    setIsSaving(true);
    const headers = ["Data1", "Data2", "Data3", "Data4", "Data5"];
    const updates: Array<{ rowIndex: number; values: string[] }> = [];

    changedRows.forEach((rowIndex) => {
      const rowData = data[rowIndex + 1];
      const values = headers.map((header) => rowData[header] || "");
      updates.push({ rowIndex: rowIndex + 3, values: values });
    });

    try {
      const requestBody = {
        action: "update_bulk",
        sheetName: "DataSiswa",
        updates: updates,
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);

      alert("Semua perubahan berhasil disimpan!");
      setChangedRows(new Set());

      const refreshResponse = await fetch(`${endpoint}?sheet=DataSiswa`);
      const refreshedData = await refreshResponse.json();
      setData(refreshedData);
    } catch (err) {
      console.error("Error:", err);
      alert(
        "Error menyimpan data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  const handleAddNew = async () => {
    if (!newSiswa.nama.trim() || !newSiswa.kelas.trim()) {
      alert("⚠️ Nama dan Kelas wajib diisi!");
      return;
    }

    setIsSaving(true);

    try {
      const requestBody = {
        action: "add_siswa",
        data: {
          nama: newSiswa.nama.trim(),
          kelas: newSiswa.kelas.trim(),
          nis: newSiswa.nis.trim(),
          nisn: newSiswa.nisn.trim(),
          namaOrtu: newSiswa.namaOrtu.trim(),
        },
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);

      alert("✅ Data siswa baru berhasil ditambahkan!");
      setNewSiswa({ nama: "", kelas: "", nis: "", nisn: "", namaOrtu: "" });
      setIsAddingNew(false);

      const refreshResponse = await fetch(`${endpoint}?sheet=DataSiswa`);
      const refreshedData = await refreshResponse.json();
      setData(refreshedData);
    } catch (err) {
      console.error("Error:", err);
      alert(
        "Error menambah data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  if (loading)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>Loading...</div>
    );
  if (error)
    return (
      <div style={{ textAlign: "center", color: "red", padding: "20px" }}>
        Error: {error}
      </div>
    );
  if (data.length === 0)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        No data available
      </div>
    );

  const headers = ["Data1", "Data2", "Data3", "Data4", "Data5"];
  const displayHeaders = headers.map((header) => data[0][header] || "");
  const actualData = data.slice(1);

  return (
    <div style={{ padding: "10px", margin: "0 auto", maxWidth: "100vw" }}>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "15px",
          fontSize: "20px",
        }}
      >
        👨‍🎓 Data Siswa
      </h1>

      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <button
          onClick={() => setIsAddingNew(!isAddingNew)}
          style={{
            padding: "12px 24px",
            backgroundColor: isAddingNew ? "#f44336" : "#2196F3",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            marginRight: "10px",
          }}
        >
          {isAddingNew ? "❌ Batal Tambah" : "➕ Tambah Siswa Baru"}
        </button>

        <button
          onClick={handleSaveAll}
          disabled={isSaving || changedRows.size === 0}
          style={{
            padding: "12px 24px",
            backgroundColor:
              isSaving || changedRows.size === 0 ? "#ccc" : "#4CAF50",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor:
              isSaving || changedRows.size === 0 ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
          }}
        >
          {isSaving
            ? "Memproses..."
            : `💾 Simpan Perubahan (${changedRows.size})`}
        </button>
      </div>

      {isAddingNew && (
        <div
          style={{
            backgroundColor: "#f0f8ff",
            padding: "20px",
            borderRadius: "8px",
            marginBottom: "20px",
            boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          }}
        >
          <h3 style={{ marginBottom: "15px", color: "#2196F3" }}>
            Form Tambah Siswa Baru
          </h3>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gap: "15px",
            }}
          >
            <input
              type="text"
              value={newSiswa.nama}
              onChange={(e) =>
                setNewSiswa({ ...newSiswa, nama: e.target.value })
              }
              placeholder="Nama Siswa *"
              style={{
                padding: "12px",
                border: "1px solid #ddd",
                borderRadius: "4px",
                fontSize: "16px",
              }}
            />
            <input
              type="text"
              value={newSiswa.kelas}
              onChange={(e) =>
                setNewSiswa({ ...newSiswa, kelas: e.target.value })
              }
              placeholder="Kelas *"
              style={{
                padding: "12px",
                border: "1px solid #ddd",
                borderRadius: "4px",
                fontSize: "16px",
              }}
            />
            <input
              type="text"
              value={newSiswa.nis}
              onChange={(e) =>
                setNewSiswa({ ...newSiswa, nis: e.target.value })
              }
              placeholder="NIS"
              style={{
                padding: "12px",
                border: "1px solid #ddd",
                borderRadius: "4px",
                fontSize: "16px",
              }}
            />
            <input
              type="text"
              value={newSiswa.nisn}
              onChange={(e) =>
                setNewSiswa({ ...newSiswa, nisn: e.target.value })
              }
              placeholder="NISN"
              style={{
                padding: "12px",
                border: "1px solid #ddd",
                borderRadius: "4px",
                fontSize: "16px",
              }}
            />
            <input
              type="text"
              value={newSiswa.namaOrtu}
              onChange={(e) =>
                setNewSiswa({ ...newSiswa, namaOrtu: e.target.value })
              }
              placeholder="Nama Orang Tua"
              style={{
                padding: "12px",
                border: "1px solid #ddd",
                borderRadius: "4px",
                fontSize: "16px",
                gridColumn: "1 / -1",
              }}
            />
          </div>
          <button
            onClick={handleAddNew}
            disabled={isSaving}
            style={{
              marginTop: "15px",
              padding: "12px 24px",
              backgroundColor: isSaving ? "#ccc" : "#4CAF50",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: isSaving ? "not-allowed" : "pointer",
              fontWeight: "bold",
              fontSize: "16px",
            }}
          >
            {isSaving ? "Menyimpan..." : "💾 Simpan Siswa Baru"}
          </button>
        </div>
      )}

      <div
        style={{
          overflowX: "auto",
          overflowY: "auto",
          maxHeight: "calc(100vh - 300px)",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          position: "relative",
        }}
      >
        <table
          style={{
            borderCollapse: "separate",
            borderSpacing: 0,
            minWidth: "100%",
            width: "max-content",
          }}
        >
          <thead style={{ position: "sticky", top: 0, zIndex: 100 }}>
            <tr style={{ backgroundColor: "#f4f4f4" }}>
              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "50px",
                }}
              >
                No
              </th>
              {displayHeaders.map((header, index) => (
                <th
                  key={index}
                  style={{
                    padding: "8px",
                    textAlign: "center",
                    borderBottom: "2px solid #ddd",
                    fontWeight: "bold",
                    minWidth: "150px",
                    backgroundColor: "#f4f4f4",
                    fontSize: "12px",
                  }}
                >
                  {header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {actualData.map((row, rowIndex) => (
              <tr
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                }}
              >
                <td
                  style={{
                    padding: "8px",
                    textAlign: "center",
                    borderBottom: "1px solid #eee",
                    fontWeight: "bold",
                    color: "#666",
                  }}
                >
                  {rowIndex + 1}
                </td>
                {headers.map((header, colIndex) => (
                  <td
                    key={colIndex}
                    style={{
                      padding: "8px",
                      borderBottom: "1px solid #eee",
                    }}
                  >
                    <input
                      type="text"
                      value={row[header] || ""}
                      onChange={(e) =>
                        handleInputChange(rowIndex, header, e.target.value)
                      }
                      style={{
                        width: "100%",
                        padding: "6px",
                        border: "1px solid #ddd",
                        borderRadius: "3px",
                        fontSize: "12px",
                      }}
                    />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

const App = () => {
  const [menuOpen, setMenuOpen] = useState(false);

  return (
    <RekapProvider>
      <Router>
        <div style={{ padding: "10px", margin: "0 auto", maxWidth: "100vw" }}>
          {/* Hamburger Menu Button */}
          <button
            onClick={() => setMenuOpen(!menuOpen)}
            style={{
              position: "fixed",
              top: "15px",
              left: "15px",
              width: "50px",
              height: "50px",
              backgroundColor: "#4CAF50",
              color: "white",
              border: "none",
              borderRadius: "8px",
              cursor: "pointer",
              zIndex: 1002,
              display: "flex",
              flexDirection: "column",
              justifyContent: "center",
              alignItems: "center",
              gap: "5px",
              boxShadow: "0 2px 10px rgba(0,0,0,0.2)",
            }}
          >
            <div
              style={{
                width: "25px",
                height: "3px",
                backgroundColor: "white",
                borderRadius: "2px",
              }}
            ></div>
            <div
              style={{
                width: "25px",
                height: "3px",
                backgroundColor: "white",
                borderRadius: "2px",
              }}
            ></div>
            <div
              style={{
                width: "25px",
                height: "3px",
                backgroundColor: "white",
                borderRadius: "2px",
              }}
            ></div>
          </button>

          {/* Menu Overlay */}
          {menuOpen && (
            <div
              style={{
                position: "fixed",
                top: 0,
                left: 0,
                right: 0,
                bottom: 0,
                backgroundColor: "rgba(0, 0, 0, 0.5)",
                zIndex: 1001,
              }}
              onClick={() => setMenuOpen(false)}
            >
              {/* Menu Panel */}
              <div
                style={{
                  position: "fixed",
                  top: 0,
                  left: 0,
                  width: "280px",
                  height: "100vh",
                  backgroundColor: "white",
                  boxShadow: "2px 0 10px rgba(0,0,0,0.3)",
                  padding: "80px 20px 20px 20px",
                  display: "flex",
                  flexDirection: "column",
                  gap: "10px",
                  overflowY: "auto", // ✅ TAMBAHKAN INI
                  overflowX: "hidden", // ✅ TAMBAHKAN INI
                }}
                onClick={(e) => e.stopPropagation()}
              >
                <h2
                  style={{
                    margin: "0 0 20px 0",
                    color: "#333",
                    fontSize: "20px",
                  }}
                >
                  📚 Menu
                </h2>

                <Link
                  to="/data-siswa"
                  onClick={() => setMenuOpen(false)}
                  style={{
                    padding: "15px 20px",
                    backgroundColor: "#f0f0f0",
                    borderRadius: "8px",
                    textDecoration: "none",
                    color: "#333",
                    fontWeight: "500",
                    transition: "background-color 0.2s",
                  }}
                  onMouseEnter={(e) =>
                    (e.currentTarget.style.backgroundColor = "#4CAF50")
                  }
                  onMouseLeave={(e) =>
                    (e.currentTarget.style.backgroundColor = "#f0f0f0")
                  }
                >
                  👨‍🎓 Data Siswa
                </Link>

                <Link
                  to="/"
                  onClick={() => setMenuOpen(false)}
                  style={{
                    padding: "15px 20px",
                    backgroundColor: "#f0f0f0",
                    borderRadius: "8px",
                    textDecoration: "none",
                    color: "#333",
                    fontWeight: "500",
                    transition: "background-color 0.2s",
                  }}
                  onMouseEnter={(e) =>
                    (e.currentTarget.style.backgroundColor = "#4CAF50")
                  }
                  onMouseLeave={(e) =>
                    (e.currentTarget.style.backgroundColor = "#f0f0f0")
                  }
                >
                  📝 Input Nilai
                </Link>

                <Link
                  to="/input-tp"
                  onClick={() => setMenuOpen(false)}
                  style={{
                    padding: "15px 20px",
                    backgroundColor: "#f0f0f0",
                    borderRadius: "8px",
                    textDecoration: "none",
                    color: "#333",
                    fontWeight: "500",
                    transition: "background-color 0.2s",
                  }}
                  onMouseEnter={(e) =>
                    (e.currentTarget.style.backgroundColor = "#4CAF50")
                  }
                  onMouseLeave={(e) =>
                    (e.currentTarget.style.backgroundColor = "#f0f0f0")
                  }
                >
                  📚 Input TP
                </Link>

                <Link
                  to="/data-mapel"
                  onClick={() => setMenuOpen(false)}
                  style={{
                    padding: "15px 20px",
                    backgroundColor: "#f0f0f0",
                    borderRadius: "8px",
                    textDecoration: "none",
                    color: "#333",
                    fontWeight: "500",
                    transition: "background-color 0.2s",
                  }}
                  onMouseEnter={(e) =>
                    (e.currentTarget.style.backgroundColor = "#4CAF50")
                  }
                  onMouseLeave={(e) =>
                    (e.currentTarget.style.backgroundColor = "#f0f0f0")
                  }
                >
                  📖 Data Mata Pelajaran
                </Link>

                <Link
                  to="/kehadiran"
                  onClick={() => setMenuOpen(false)}
                  style={{
                    padding: "15px 20px",
                    backgroundColor: "#f0f0f0",
                    borderRadius: "8px",
                    textDecoration: "none",
                    color: "#333",
                    fontWeight: "500",
                    transition: "background-color 0.2s",
                  }}
                  onMouseEnter={(e) =>
                    (e.currentTarget.style.backgroundColor = "#4CAF50")
                  }
                  onMouseLeave={(e) =>
                    (e.currentTarget.style.backgroundColor = "#f0f0f0")
                  }
                >
                  📋 Data Kehadiran
                </Link>

                <Link
                  to="/kokurikuler"
                  onClick={() => setMenuOpen(false)}
                  style={{
                    padding: "15px 20px",
                    backgroundColor: "#f0f0f0",
                    borderRadius: "8px",
                    textDecoration: "none",
                    color: "#333",
                    fontWeight: "500",
                    transition: "background-color 0.2s",
                  }}
                  onMouseEnter={(e) =>
                    (e.currentTarget.style.backgroundColor = "#4CAF50")
                  }
                  onMouseLeave={(e) =>
                    (e.currentTarget.style.backgroundColor = "#f0f0f0")
                  }
                >
                  🌟 Data Kokurikuler
                </Link>

                <Link
                  to="/ekstrakurikuler"
                  onClick={() => setMenuOpen(false)}
                  style={{
                    padding: "15px 20px",
                    backgroundColor: "#f0f0f0",
                    borderRadius: "8px",
                    textDecoration: "none",
                    color: "#333",
                    fontWeight: "500",
                    transition: "background-color 0.2s",
                  }}
                  onMouseEnter={(e) =>
                    (e.currentTarget.style.backgroundColor = "#4CAF50")
                  }
                  onMouseLeave={(e) =>
                    (e.currentTarget.style.backgroundColor = "#f0f0f0")
                  }
                >
                  🎯 Data Ekstrakurikuler
                </Link>

                <Link
                  to="/rekap"
                  onClick={() => setMenuOpen(false)}
                  style={{
                    padding: "15px 20px",
                    backgroundColor: "#f0f0f0",
                    borderRadius: "8px",
                    textDecoration: "none",
                    color: "#333",
                    fontWeight: "500",
                    transition: "background-color 0.2s",
                  }}
                  onMouseEnter={(e) =>
                    (e.currentTarget.style.backgroundColor = "#4CAF50")
                  }
                  onMouseLeave={(e) =>
                    (e.currentTarget.style.backgroundColor = "#f0f0f0")
                  }
                >
                  📊 Rekap Nilai
                </Link>

                <Link
                  to="/data-sekolah"
                  onClick={() => setMenuOpen(false)}
                  style={{
                    padding: "15px 20px",
                    backgroundColor: "#f0f0f0",
                    borderRadius: "8px",
                    textDecoration: "none",
                    color: "#333",
                    fontWeight: "500",
                    transition: "background-color 0.2s",
                  }}
                  onMouseEnter={(e) =>
                    (e.currentTarget.style.backgroundColor = "#4CAF50")
                  }
                  onMouseLeave={(e) =>
                    (e.currentTarget.style.backgroundColor = "#f0f0f0")
                  }
                >
                  🏫 Data Sekolah
                </Link>
              </div>
            </div>
          )}

          <Routes>
            <Route path="/data-siswa" element={<DataSiswa />} />
            <Route path="/" element={<InputNilai />} />
            <Route path="/kehadiran" element={<DataKehadiran />} />
            <Route path="/kokurikuler" element={<DataKokurikuler />} />
            <Route path="/ekstrakurikuler" element={<DataEkstrakurikuler />} />
            <Route path="/input-tp" element={<InputTP />} />
            <Route path="/data-mapel" element={<DataMapel />} />
            <Route path="/rekap" element={<RekapNilai />} />
            <Route path="/data-sekolah" element={<DataSekolah />} />
          </Routes>
        </div>
      </Router>
    </RekapProvider>
  );
};

export default App;

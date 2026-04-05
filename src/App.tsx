import React, { useState, useEffect, useCallback, useRef } from 'react';
import mammoth from 'mammoth';
import { DEFAULT_STAFF, SHIFTS } from './constants';
import { fmtVN, fmtIn, dayN, timNghi, timThay, abbrev, xacDinhCa } from './utils/shiftHelpers';
import { buildMultiLeaveResults, Leave, ResultItem } from './utils/multiLeaveAlgorithm';
import { exportWord, generateWordBlob, exportSwapDoc, generateSwapBlob } from './utils/wordExport';
import { renderAsync } from 'docx-preview';

export default function App() {
  const [staffData, setStaffData] = useState<string[][]>(() => {
    const saved = localStorage.getItem('sd');
    const ver = localStorage.getItem('sv');
    if (ver !== 'v4') {
      localStorage.removeItem('sd');
      localStorage.setItem('sv', 'v4');
      return DEFAULT_STAFF;
    }
    return saved ? JSON.parse(saved) : DEFAULT_STAFF;
  });

  const [ngayBatDau, setNgayBatDau] = useState(() => fmtIn(new Date()));
  const [ngayKetThuc, setNgayKetThuc] = useState(() => {
    const next = new Date();
    next.setDate(next.getDate() + 7);
    return fmtIn(next);
  });
  const [chucDanh, setChucDanh] = useState('');
  const [kipNghi, setKipNghi] = useState('');
  const [additionalLeaves, setAdditionalLeaves] = useState<{ kip: string, start: string, end: string, chucDanh: string }[]>([]);
  const [showStaff, setShowStaff] = useState(false);
  const [alert, setAlert] = useState<string | null>(null);
  const [currentResult, setCurrentResult] = useState<any>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const [previewBlob, setPreviewBlob] = useState<Blob | null>(null);
  const previewRef = useRef<HTMLDivElement>(null);
  const [isParsing, setIsParsing] = useState(false);
  const [config, setConfig] = useState({
    soVanBan: '',
    ngayKy: '',
    nguoiKy: 'Nguyễn Văn Nghị'
  });

  const [isGoogleAuth, setIsGoogleAuth] = useState(false);
  const [isUpdatingSheets, setIsUpdatingSheets] = useState(false);

  // Manual Swap State
  const [swapData, setSwapData] = useState({
    date1: fmtIn(new Date()),
    date2: fmtIn(new Date()),
    person1: '',
    person2: '',
    shift1: 'N',
    shift2: 'K'
  });
  const [isPreviewingSwap, setIsPreviewingSwap] = useState(false);
  const [swapChucDanh, setSwapChucDanh] = useState(staffData[0]?.[0] || '');

  useEffect(() => {
    const checkAuth = async () => {
      try {
        const res = await fetch('/api/auth/status');
        const data = await res.json();
        setIsGoogleAuth(data.authenticated);
      } catch (e) {
        console.error("Auth check failed", e);
      }
    };
    checkAuth();

    const handleMessage = (e: MessageEvent) => {
      if (e.data?.type === 'OAUTH_AUTH_SUCCESS') {
        setIsGoogleAuth(true);
        setAlert('✅ Kết nối Google thành công!');
      }
    };
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, []);

  const handleConnectGoogle = () => {
    const width = 500, height = 600;
    const left = (window.innerWidth - width) / 2;
    const top = (window.innerHeight - height) / 2;
    window.open('/api/auth/google', 'google-auth', `width=${width},height=${height},left=${left},top=${top}`);
  };

  const updateGoogleSheets = async () => {
    if (!isGoogleAuth || !currentResult) return;
    setIsUpdatingSheets(true);
    try {
      const updateMap: Record<string, string> = {}; // key: "name|date"

      // 1. Process all leaves (Absent people)
      currentResult.allResults.forEach((res: any) => {
        const start = new Date(res.start);
        const end = new Date(res.end);
        let d = new Date(start);
        while (d <= end) {
          const dateStr = fmtIn(d);
          const originalShift = xacDinhCa(d, res.kip);
          // Only change N, C, K to F. Keep O as O.
          const finalShift = (originalShift === 'N' || originalShift === 'C' || originalShift === 'K') ? 'F' : originalShift;
          
          if (!res.ten.includes('THIẾU NHÂN SỰ')) {
            const name = res.ten.trim().normalize('NFC');
            updateMap[`${name}|${dateStr}`] = finalShift;
          }
          d.setDate(d.getDate() + 1);
        }

        // 2. Process specific assignments for each leave
        res.ketQua.forEach((item: any) => {
          const dateStr = fmtIn(new Date(item.ngay));
          
          // Replacement person works the shift
          if (item.nguoiThay && item.nguoiThay !== 'N/A' && !item.nguoiThay.includes('THIẾU NHÂN SỰ')) {
            const name = item.nguoiThay.trim().normalize('NFC');
            updateMap[`${name}|${dateStr}`] = item.ca;
          }

          // If it's a CK Swap, the absent person works the other shift
          if (item.isCKSwap && item.swapAbsentTen && !item.swapAbsentTen.includes('THIẾU NHÂN SỰ')) {
            const name = item.swapAbsentTen.trim().normalize('NFC');
            const absentShift = xacDinhCa(new Date(item.ngay), item.kipThay);
            updateMap[`${name}|${dateStr}`] = absentShift;
          }

          // If someone is relieved
          if (item.relievedTen && !item.relievedTen.includes('THIẾU NHÂN SỰ')) {
            const name = item.relievedTen.trim().normalize('NFC');
            updateMap[`${name}|${dateStr}`] = 'O';
          }
        });
      });

      // 3. Process extraRows (Additional adjustments)
      currentResult.extraRows.forEach((row: any) => {
        const dateStr = fmtIn(new Date(row.ngay));
        
        if (row.nguoiThay && row.nguoiThay !== 'N/A' && !row.nguoiThay.includes('THIẾU NHÂN SỰ')) {
          const name = row.nguoiThay.trim().normalize('NFC');
          updateMap[`${name}|${dateStr}`] = row.ca;
        }
        
        if (row.relievedTen && !row.relievedTen.includes('THIẾU NHÂN SỰ')) {
          const name = row.relievedTen.trim().normalize('NFC');
          updateMap[`${name}|${dateStr}`] = 'O';
        }

        if (row.absentTen && !row.absentTen.includes('THIẾU NHÂN SỰ')) {
          const name = row.absentTen.trim().normalize('NFC');
          // Only set to F if not already assigned a working shift (like in a swap)
          const currentVal = updateMap[`${name}|${dateStr}`];
          if (!currentVal || currentVal === 'F') {
            const originalShift = xacDinhCa(new Date(row.ngay), row.absentKip);
            const finalShift = (originalShift === 'N' || originalShift === 'C' || originalShift === 'K') ? 'F' : originalShift;
            updateMap[`${name}|${dateStr}`] = finalShift;
          }
        }
      });

      const updates = Object.entries(updateMap).map(([key, shift]) => {
        const [name, date] = key.split('|');
        return { name, date, shift };
      });

      console.log("Sending updates to Google Sheets:", updates);

      const res = await fetch('/api/sheets/update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          spreadsheetId: '1HgGW-FvoGQXtj7V_JCMD-7Tuue0rTIM-bmohGmgqm6I',
          updates
        })
      });
      const data = await res.json();
      if (data.success) {
        setAlert('✅ Đã cập nhật Google Sheets thành công (Ca nghỉ: F)');
      } else {
        setAlert('⚠ Lỗi cập nhật Google Sheets: ' + (data.error || 'Unknown error'));
      }
    } catch (e) {
      console.error(e);
      setAlert('⚠ Lỗi kết nối Google Sheets');
    } finally {
      setIsUpdatingSheets(false);
    }
  };

  const handleUpdateStaff = (r: number, c: number, val: string) => {
    const newData = [...staffData];
    newData[r][c] = val.trim();
    setStaffData(newData);
    localStorage.setItem('sd', JSON.stringify(newData));
  };

  const addConcurrentLeave = () => {
    if (!chucDanh) {
      setAlert('⚠ Vui lòng chọn chức danh trước!');
      return;
    }
    setAdditionalLeaves([...additionalLeaves, { kip: '', start: '', end: '', chucDanh: chucDanh || (staffData[0] ? staffData[0][0] : '') }]);
  };

  const removeConcurrentLeave = (idx: number) => {
    const newList = [...additionalLeaves];
    newList.splice(idx, 1);
    setAdditionalLeaves(newList);
  };

  const updateConcurrentLeave = (idx: number, field: string, val: string) => {
    const newList = [...additionalLeaves];
    (newList[idx] as any)[field] = val;
    setAdditionalLeaves(newList);
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setIsParsing(true);
    setAlert(null);
    
    let mainSet = !!(chucDanh && kipNghi);
    const newAdditionalLeaves = [...additionalLeaves];
    let successCount = 0;
    let errorCount = 0;

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      try {
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.extractRawText({ arrayBuffer });
        const text = result.value;

        // Extraction logic
        const nameMatch = text.match(/Tên tôi là:\s*(.*)/i);
        const positionMatch = text.match(/Chức vụ:\s*(.*)/i);
        const dateMatch = text.match(/Thời gian:\s*Từ ngày\s*(\d{2}\/\d{2}\/\d{4})\s*đến hết ngày\s*(\d{2}\/\d{2}\/\d{4})/i);

        if (!nameMatch || !positionMatch || !dateMatch) {
          errorCount++;
          continue;
        }

        const extractedName = nameMatch[1].trim();
        const extractedPosition = positionMatch[1].trim();
        const startDateStr = dateMatch[1].trim();
        const endDateStr = dateMatch[2].trim();

        // Parse dates (DD/MM/YYYY to YYYY-MM-DD)
        const parseDate = (dStr: string) => {
          const [d, m, y] = dStr.split('/');
          return `${y}-${m}-${d}`;
        };

        const startISO = parseDate(startDateStr);
        const endISO = parseDate(endDateStr);

        // Try to find the person in staffData to get the correct title and kip
        let foundTitle = '';
        let foundKip = '';

        for (let r = 0; r < staffData.length; r++) {
          for (let c = 1; c <= 5; c++) {
            if (staffData[r][c]?.toLowerCase() === extractedName.toLowerCase()) {
              foundTitle = staffData[r][0];
              foundKip = String(c);
              break;
            }
          }
          if (foundTitle) break;
        }

        // If name not found, try to extract kip from position string (e.g., "Kíp2")
        if (!foundKip) {
          const kipMatch = extractedPosition.match(/Kíp\s*(\d)/i);
          if (kipMatch) foundKip = kipMatch[1];
        }

        // If title not found, try to match extractedPosition with staffData titles
        if (!foundTitle) {
          const sortedTitles = [...staffData].sort((a, b) => b[0].length - a[0].length);
          const matchedTitle = sortedTitles.find(r => extractedPosition.toLowerCase().includes(r[0].toLowerCase()));
          if (matchedTitle) foundTitle = matchedTitle[0];
        }

        if (!mainSet) {
          setNgayBatDau(startISO);
          setNgayKetThuc(endISO);
          setChucDanh(foundTitle);
          setKipNghi(foundKip);
          mainSet = true;
          successCount++;
        } else {
          newAdditionalLeaves.push({
            kip: foundKip,
            start: startISO,
            end: endISO,
            chucDanh: foundTitle
          });
          successCount++;
        }
      } catch (err) {
        console.error(err);
        errorCount++;
      }
    }

    setAdditionalLeaves(newAdditionalLeaves);

    if (successCount > 0) {
      setAlert(`✅ Đã trích xuất thành công ${successCount} đơn nghỉ phép.${errorCount > 0 ? ` (Thất bại ${errorCount} file)` : ''}`);
    } else if (errorCount > 0) {
      setAlert(`⚠ Không thể trích xuất thông tin từ ${errorCount} file. Vui lòng kiểm tra định dạng.`);
    }

    setIsParsing(false);
    e.target.value = '';
  };

  const taoLich = () => {
    setAlert(null);
    if (!ngayBatDau || !ngayKetThuc || !chucDanh || !kipNghi) {
      setAlert('⚠ Vui lòng nhập đầy đủ thông tin!');
      return;
    }

    const start = new Date(ngayBatDau + 'T00:00:00');
    const end = new Date(ngayKetThuc + 'T00:00:00');
    const kip = +kipNghi;

    if (start > end) {
      setAlert('⚠ Ngày bắt đầu phải trước ngày kết thúc!');
      return;
    }

    const ten = timNghi(chucDanh, kip, staffData);
    if (!ten) {
      setAlert(`⚠ Không tìm thấy chức danh "${chucDanh}" trong Kíp ${kip}!`);
      return;
    }

    const allLeaves: Leave[] = [{ kip, start, end, ten, chucDanh }];
    const addErr: string[] = [];
    additionalLeaves.forEach((al, idx) => {
      if (!al.kip && !al.start && !al.end) return;
      if (!al.kip || !al.start || !al.end || !al.chucDanh) {
        addErr.push(`Người nghỉ #${idx + 2} thiếu thông tin`);
        return;
      }
      const alKip = +al.kip;
      const alChucDanh = al.chucDanh;

      // Check for duplicate kip within the same chucDanh
      if (allLeaves.some(l => l.kip === alKip && l.chucDanh === alChucDanh)) {
        addErr.push(`Kíp ${alKip} của chức danh ${alChucDanh} đã được thêm`);
        return;
      }

      const alStart = new Date(al.start + 'T00:00:00');
      const alEnd = new Date(al.end + 'T00:00:00');
      if (alStart > alEnd) {
        addErr.push(`Người nghỉ #${idx + 2} ngày bắt đầu sau ngày kết thúc`);
        return;
      }
      const alTen = timNghi(alChucDanh, alKip, staffData);
      if (!alTen) {
        addErr.push(`Không tìm thấy "${alChucDanh}" trong Kíp ${alKip}`);
        return;
      }
      allLeaves.push({ kip: alKip, start: alStart, end: alEnd, ten: alTen, chucDanh: alChucDanh });
    });

    if (addErr.length) {
      setAlert('⚠ ' + addErr.join(' | '));
      return;
    }

    setIsProcessing(true);
    setTimeout(() => {
      // Tự động phát hiện và bổ sung các vị trí thiếu nhân sự (ô trống trong danh sách)
      // Nếu một chức danh có người nghỉ phép, các kíp đang thiếu người ở chức danh đó cũng sẽ được xếp lịch thay
      const minStart = new Date(Math.min(...allLeaves.map(l => l.start.getTime())));
      const maxEnd = new Date(Math.max(...allLeaves.map(l => l.end.getTime())));
      const involvedTitles = Array.from(new Set(allLeaves.map(l => l.chucDanh)));
      
      involvedTitles.forEach(title => {
        const row = staffData.find(r => r[0] === title);
        if (row) {
          for (let k = 1; k <= 5; k++) {
            const staffName = row[k] ? row[k].trim() : '';
            if (staffName === '') {
              if (!allLeaves.some(l => l.kip === k && l.chucDanh === title)) {
                allLeaves.push({
                  kip: k,
                  start: minStart,
                  end: maxEnd,
                  ten: `THIẾU NHÂN SỰ (Kíp ${k})`,
                  chucDanh: title
                });
              }
            }
          }
        }
      });

      // Group leaves by job title
      const groups = allLeaves.reduce((acc, l) => {
        if (!acc[l.chucDanh]) acc[l.chucDanh] = [];
        acc[l.chucDanh].push(l);
        return acc;
      }, {} as Record<string, Leave[]>);

      let mergedResults: any[] = [];
      let mergedExtraRows: any[] = [];
      let mergedHasConflict = false;
      const mergedCoverCount: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };

      Object.keys(groups).forEach(cd => {
        const buildResult = buildMultiLeaveResults(groups[cd], cd, staffData);
        mergedResults = [...mergedResults, ...buildResult.results];
        mergedExtraRows = [...mergedExtraRows, ...buildResult.extraRows];
        if (buildResult.hasConflict) mergedHasConflict = true;
        for (let k = 1; k <= 5; k++) {
          mergedCoverCount[k] += (buildResult.coverCount[k] || 0);
        }
      });

      setCurrentResult({
        ten, chucDanh, kip, start, end,
        ketQua: mergedResults[0].ketQua,
        allResults: mergedResults,
        extraRows: mergedExtraRows,
        hasConflict: mergedHasConflict,
        coverCount: mergedCoverCount,
        isMulti: allLeaves.length > 1
      });
      setIsProcessing(false);
    }, 250);
  };

  const handleExportWord = async () => {
    setIsProcessing(true);
    await exportWord(currentResult, config);
    if (isGoogleAuth) {
      await updateGoogleSheets();
    }
    setIsProcessing(false);
  };

  useEffect(() => {
    if (showPreview) {
      document.body.style.overflow = 'hidden';
    } else {
      document.body.style.overflow = 'unset';
    }
    return () => {
      document.body.style.overflow = 'unset';
    };
  }, [showPreview]);

  const handlePreviewWord = async () => {
    setIsProcessing(true);
    const blob = await generateWordBlob(currentResult, config);
    if (blob) {
      setPreviewBlob(blob);
      setIsPreviewingSwap(false);
      setShowPreview(true);
    }
    setIsProcessing(false);
  };

  const handlePreviewSwap = async () => {
    setIsProcessing(true);
    const blob = await generateSwapBlob(swapData, config);
    if (blob) {
      setPreviewBlob(blob);
      setIsPreviewingSwap(true);
      setShowPreview(true);
    }
    setIsProcessing(false);
  };

  const updateSwapGoogleSheets = async () => {
    if (!isGoogleAuth || !swapData.person1 || !swapData.person2) return;
    setIsUpdatingSheets(true);
    try {
      const updateMap: Record<string, string> = {}; // key: "name|date"

      const p1 = swapData.person1.trim().normalize('NFC');
      const p2 = swapData.person2.trim().normalize('NFC');

      if (swapData.date1 === swapData.date2) {
        // Same day swap: P1 takes P2's shift, P2 takes P1's shift
        updateMap[`${p1}|${swapData.date1}`] = swapData.shift2;
        updateMap[`${p2}|${swapData.date1}`] = swapData.shift1;
      } else {
        // Different day swap
        updateMap[`${p1}|${swapData.date1}`] = 'O';
        updateMap[`${p1}|${swapData.date2}`] = swapData.shift2;
        updateMap[`${p2}|${swapData.date1}`] = swapData.shift1;
        updateMap[`${p2}|${swapData.date2}`] = 'O';
      }

      const updates = Object.entries(updateMap).map(([key, shift]) => {
        const [name, date] = key.split('|');
        return { name, date, shift };
      });

      console.log("Sending manual swap updates to Google Sheets:", updates);

      const res = await fetch('/api/sheets/update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          spreadsheetId: '1HgGW-FvoGQXtj7V_JCMD-7Tuue0rTIM-bmohGmgqm6I',
          updates
        })
      });
      const data = await res.json();
      if (data.success) {
        setAlert('✅ Đã cập nhật Google Sheets thành công cho Lịch Đổi Ca');
      } else {
        setAlert('⚠ Lỗi cập nhật Google Sheets: ' + (data.error || 'Unknown error'));
      }
    } catch (err) {
      console.error(err);
      setAlert('⚠ Lỗi kết nối Google Sheets');
    } finally {
      setIsUpdatingSheets(false);
    }
  };

  const handleExportSwap = async () => {
    setIsProcessing(true);
    await exportSwapDoc(swapData, config);
    if (isGoogleAuth) {
      await updateSwapGoogleSheets();
    }
    setIsProcessing(false);
  };

  useEffect(() => {
    if (showPreview && previewBlob && previewRef.current) {
      renderAsync(previewBlob, previewRef.current)
        .catch(err => console.error("Preview error:", err));
    }
  }, [showPreview, previewBlob]);

  const handleExportCSV = () => {
    if (!currentResult) return;
    const d = currentResult;
    let csv = '\uFEFF';
    csv += 'CÔNG TY THỦY ĐIỆN IALY - BẢNG PHÂN CÔNG TRỰC THAY NGHỈ PHÉP\n\n';
    csv += `Người nghỉ:,${d.ten}\nChức danh:,${d.chucDanh}\nKíp:,Kíp ${d.kip}\n`;
    csv += `Thời gian:,${fmtVN(d.start)} đến ${fmtVN(d.end)}\n\n`;
    csv += 'STT,Ngày,Ca nghỉ,Kíp thay,Người đi thay\n';
    d.ketQua.forEach((it: any, i: number) => {
      csv += `${i + 1},${fmtVN(it.ngay)},${it.ca},Kíp ${it.kipThay},${it.nguoiThay}\n`;
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Lich_nghi_${d.ten.replace(/\s+/g, '_')}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // Flatten rows for display
  const shiftOrd: Record<string, number> = { N: 0, C: 1, K: 2 };
  const allRows: any[] = [];
  if (currentResult) {
    currentResult.allResults.forEach((res: any) => {
      res.ketQua.forEach((it: any) => {
        allRows.push({
          ngay: it.ngay, ca: it.ca,
          absentKip: res.kip,
          absentTen: res.ten,
          chucDanh: res.chucDanh,
          kipThay: it.kipThay, nguoiThay: it.nguoiThay,
          relievedKip: it.relievedKip, relievedTen: it.relievedTen,
          isConflict: it.isConflict, conflictNote: it.conflictNote || '',
          isOverlapDay: it.isOverlapDay,
          isCKChain: false,
          isSwap: false
        });
      });
    });
    (currentResult.extraRows || []).forEach((it: any) => {
      if (it.isSwapCRow) return;
      allRows.push({
        ngay: it.ngay, ca: it.ca,
        absentKip: it.absentKip, absentTen: it.absentTen,
        chucDanh: it.chucDanh,
        kipThay: it.kipThay, nguoiThay: it.nguoiThay,
        relievedKip: it.relievedKip, relievedTen: it.relievedTen,
        isConflict: it.isConflict, conflictNote: it.conflictNote || '',
        isOverlapDay: it.isOverlapDay, 
        isCKChain: it.isCKChain,
        isSwap: it.isSwap || false, 
        isSwapCRow: false
      });
    });
    allRows.sort((a, b) => {
      if (a.ngay < b.ngay) return -1;
      if (a.ngay > b.ngay) return 1;
      return (shiftOrd[a.ca] || 0) - (shiftOrd[b.ca] || 0);
    });
  }

  const coverStats: Record<number, any> = {};
  allRows.forEach(row => {
    const k = row.kipThay;
    const cd = row.chucDanh;
    if (!coverStats[k]) coverStats[k] = { total: 0, N: 0, C: 0, K: 0, byCD: {} };
    coverStats[k].total++;
    coverStats[k][row.ca]++;
    if (!coverStats[k].byCD[cd]) coverStats[k].byCD[cd] = 0;
    coverStats[k].byCD[cd]++;
  });

  const isEvenDistribution = () => {
    const counts = Object.values(coverStats).map(s => s.total);
    if (counts.length === 0) return true;
    return (Math.max(...counts) - Math.min(...counts)) <= 1;
  };

  return (
    <div className="wrap">
      <header>
        <div className="tag">Công ty Thủy Điện Ialy &middot; Phân Xưởng Vận Hành Ialy</div>
        <h1>Lịch Trực Thay Ca Nghỉ Phép</h1>
        <p className="sub">Hệ thống phân công tự động lịch trực ca vận hành nghỉ phép</p>
      </header>

      <div className="card">
        <div className="ctitle">Tải lên đơn xin nghỉ phép</div>
        <div className="p-4 border-2 border-dashed border-var(--acc-light) rounded-xl text-center hover:bg-var(--acc-light) transition-colors cursor-pointer relative">
          <input 
            type="file" 
            accept=".docx" 
            multiple
            onChange={handleFileUpload} 
            className="absolute inset-0 opacity-0 cursor-pointer"
            disabled={isParsing}
          />
          <div className="flex flex-col items-center gap-2">
            <span className="text-2xl">{isParsing ? '⌛' : '📄'}</span>
            <span className="text-[13px] font-medium">
              {isParsing ? 'Đang xử lý...' : 'Nhấn để chọn hoặc kéo thả các file Word (.docx)'}
            </span>
            <span className="text-[11px] text-var(--txt2)">Hệ thống sẽ lấy thông tin từ các đơn nghỉ phép </span>
          </div>
        </div>
      </div>

      <div className="card">
        <div className="ctitle">Thông tin nghỉ phép</div>
        {alert && <div className={`alert ${alert.startsWith('✅') ? 'asuc' : 'aerr'}`}>{alert}</div>}
        <div className="g2">
          <div className="field">
            <label>Ngày bắt đầu nghỉ</label>
            <input type="date" value={ngayBatDau} onChange={e => setNgayBatDau(e.target.value)} />
          </div>
          <div className="field">
            <label>Ngày kết thúc nghỉ</label>
            <input type="date" value={ngayKetThuc} onChange={e => setNgayKetThuc(e.target.value)} />
          </div>
          <div className="field">
            <label>Chức danh người nghỉ</label>
            <select value={chucDanh} onChange={e => setChucDanh(e.target.value)}>
              <option value="">-- Chọn chức danh --</option>
              {staffData.map(r => <option key={r[0]} value={r[0]}>{r[0]}</option>)}
            </select>
          </div>
          <div className="field">
            <label>Kíp nghỉ</label>
            <select value={kipNghi} onChange={e => setKipNghi(e.target.value)}>
              <option value="">-- Chọn kíp --</option>
              {[1, 2, 3, 4, 5].map(k => <option key={k} value={k}>Kíp {k}</option>)}
            </select>
          </div>
        </div>

        <div className="divider"></div>
        <div className="flex items-center justify-between mb-3">
          <span className="text-[12px] text-var(--txt2) font-semibold uppercase tracking-wider">
            Người nghỉ đồng thời <span className="font-normal">(có thể khác chức danh)</span>
          </span>
          <button className="staff-toggle font-bold text-[13px]" onClick={addConcurrentLeave}>+ Thêm người nghỉ</button>
        </div>

        <div id="concurrent-list">
          {additionalLeaves.map((leave, idx) => (
            <div key={idx} className="concurrent-item">
              <div className="absolute top-2 right-2.5">
                <button onClick={() => removeConcurrentLeave(idx)} className="text-var(--acc3) cursor-pointer text-lg leading-none" title="Xóa">✕</button>
              </div>
              <div className="text-[11px] font-bold text-var(--acc) uppercase tracking-wider mb-3">Người nghỉ #{idx + 2}</div>
              <div className="g2">
                <div className="field">
                  <label>Chức danh</label>
                  <select value={leave.chucDanh} onChange={e => updateConcurrentLeave(idx, 'chucDanh', e.target.value)}>
                    <option value="">-- Chức danh --</option>
                    {staffData.map(r => <option key={r[0]} value={r[0]}>{r[0]}</option>)}
                  </select>
                </div>
                <div className="field">
                  <label>Kíp nghỉ</label>
                  <select value={leave.kip} onChange={e => updateConcurrentLeave(idx, 'kip', e.target.value)}>
                    <option value="">-- Kíp --</option>
                    {[1, 2, 3, 4, 5].map(k => <option key={k} value={k}>Kíp {k}</option>)}
                  </select>
                </div>
                <div className="field">
                  <label>Ngày bắt đầu</label>
                  <input type="date" value={leave.start} onChange={e => updateConcurrentLeave(idx, 'start', e.target.value)} />
                </div>
                <div className="field">
                  <label>Ngày kết thúc</label>
                  <input type="date" value={leave.end} onChange={e => updateConcurrentLeave(idx, 'end', e.target.value)} />
                </div>
              </div>
            </div>
          ))}
        </div>

        <button className="btn btn-primary" onClick={taoLich} disabled={isProcessing}>
          {isProcessing ? <span className="spin mr-2"></span> : '⚡'} Tạo Lịch Thay Ca
        </button>
      </div>

      <div className="card">
        <div className="ctitle">
          Nhân sự của kíp 
          <button className="staff-toggle" onClick={() => setShowStaff(!showStaff)}>
            {showStaff ? 'Thu gọn ▲' : 'Chỉnh sửa ▼'}
          </button>
        </div>
        <p className="text-[13px] text-var(--txt2)">Nhấn "Chỉnh sửa" để cập nhật tên nhân viên trong từng kíp.</p>
        {showStaff && (
          <div className="staff-wrap">
            <table className="st">
              <thead>
                <tr>
                  <th>Chức danh</th>
                  <th>Kíp 1</th>
                  <th>Kíp 2</th>
                  <th>Kíp 3</th>
                  <th>Kíp 4</th>
                  <th>Kíp 5</th>
                </tr>
              </thead>
              <tbody>
                {staffData.map((row, r) => (
                  <tr key={r}>
                    <td>{row[0]}</td>
                    {[1, 2, 3, 4, 5].map(c => (
                      <td key={c}>
                        <input 
                          value={row[c] || ''} 
                          onChange={e => handleUpdateStaff(r, c, e.target.value)}
                        />
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      <div className="card">
        <div className="ctitle">Tạo lịch đổi ca thủ công</div>
        <p className="text-[13px] text-var(--txt2) mb-4">Sử dụng chức năng này để tạo nhanh văn bản "Lịch đổi ca" giữa 2 nhân viên.</p>
        <div className="g2">
          <div className="field">
            <label>Chức danh</label>
            <select value={swapChucDanh} onChange={e => {
              setSwapChucDanh(e.target.value);
              setSwapData({...swapData, person1: '', person2: ''});
            }}>
              {staffData.map(r => <option key={r[0]} value={r[0]}>{r[0]}</option>)}
            </select>
          </div>
          <div className="field"></div>
          <div className="field">
            <label>Ngày đổi ca 1</label>
            <input 
              type="date" 
              value={swapData.date1} 
              onChange={e => setSwapData({...swapData, date1: e.target.value})} 
            />
          </div>
          <div className="field">
            <label>Ngày đổi ca 2</label>
            <input 
              type="date" 
              value={swapData.date2} 
              onChange={e => setSwapData({...swapData, date2: e.target.value})} 
            />
          </div>
          <div className="field">
            <label>Người đổi ca (P1)</label>
            <input 
              type="text" 
              list="staff-list"
              placeholder="Chọn hoặc nhập tên"
              value={swapData.person1} 
              onChange={e => setSwapData({...swapData, person1: e.target.value})} 
            />
          </div>
          <div className="field">
            <label>Người đi ca thay (P2)</label>
            <input 
              type="text" 
              list="staff-list"
              placeholder="Chọn hoặc nhập tên"
              value={swapData.person2} 
              onChange={e => setSwapData({...swapData, person2: e.target.value})} 
            />
          </div>
          <datalist id="staff-list">
            {staffData.find(r => r[0] === swapChucDanh)?.slice(1).filter(Boolean).map(name => (
              <option key={name} value={name} />
            ))}
          </datalist>
          <div className="field">
            <label>Ca của P1 (nghỉ)</label>
            <select value={swapData.shift1} onChange={e => setSwapData({...swapData, shift1: e.target.value})}>
              <option value="N">Ca N</option>
              <option value="C">Ca C</option>
              <option value="K">Ca K</option>
            </select>
          </div>
          <div className="field">
            <label>Ca của P2 (nghỉ)</label>
            <select value={swapData.shift2} onChange={e => setSwapData({...swapData, shift2: e.target.value})}>
              <option value="N">Ca N</option>
              <option value="C">Ca C</option>
              <option value="K">Ca K</option>
            </select>
          </div>
        </div>
        <div className="flex gap-3 mt-4">
          <button className="btn btn-primary flex-1" onClick={handlePreviewSwap} disabled={isProcessing || !swapData.person1 || !swapData.person2}>
            {isProcessing ? <span className="spin mr-2"></span> : '📝'} Xem trước Lịch Đổi Ca
          </button>
          <button className="btn btn-secondary flex-1" onClick={handleExportSwap} disabled={!swapData.person1 || !swapData.person2}>
            📥 Xuất File Word
          </button>
        </div>
      </div>

      {currentResult && (
        <div id="results">
          <div className="card">
            <div className="res-hdr">
              <div className="ctitle mb-0">Kết quả phân công</div>
              <div className="ex-row">
                <button className="btn-ex btn-word" onClick={handlePreviewWord} disabled={isProcessing}>
                  {isProcessing ? <span className="spin spinw mr-2"></span> : '📝'} Xem trước Word
                </button>
                <button className="btn-ex btn-word" onClick={handleExportWord} disabled={isProcessing}>
                  {isProcessing ? <span className="spin spinw mr-2"></span> : '📝'} Xuất Word
                </button>
                {!isGoogleAuth ? (
                  <button className="btn-ex bg-[#4285F4] text-white hover:bg-[#357ae8]" onClick={handleConnectGoogle}>
                    🔗 Kết nối Google Sheets
                  </button>
                ) : (
                  <button 
                    className={`btn-ex ${isUpdatingSheets ? 'bg-gray-400' : 'bg-[#34a853]'} text-white`} 
                    onClick={updateGoogleSheets}
                    disabled={isUpdatingSheets}
                  >
                    {isUpdatingSheets ? <span className="spin spinw mr-2"></span> : '🔄'} 
                    {isUpdatingSheets ? 'Đang đồng bộ...' : 'Đồng bộ Google Sheets'}
                  </button>
                )}
              </div>
            </div>

            <div className="igrid">
              {currentResult.isMulti ? (
                <>
                  <div className="iitem"><div className="ilbl">Chức danh</div><div className="ival">{currentResult.chucDanh}</div></div>
                  <div className="iitem"><div className="ilbl">Số người nghỉ</div><div className="ival">{currentResult.allResults.length} người</div></div>
                  <div className="iitem"><div className="ilbl">Người nghỉ</div><div className="ival">{currentResult.allResults.map((r: any) => `${r.ten} (K${r.kip})`).join(' | ')}</div></div>
                  <div className="iitem"><div className="ilbl">Tổng ca thay</div><div className="ival">{allRows.length} ca</div></div>
                </>
              ) : (
                <>
                  <div className="iitem"><div className="ilbl">Người nghỉ</div><div className="ival">{currentResult.ten}</div></div>
                  <div className="iitem"><div className="ilbl">Chức danh</div><div className="ival">{currentResult.chucDanh}</div></div>
                  <div className="iitem"><div className="ilbl">Kíp nghỉ</div><div className="ival">Kíp {currentResult.kip}</div></div>
                  <div className="iitem"><div className="ilbl">Thời gian</div><div className="ival">{fmtVN(currentResult.start)} → {fmtVN(currentResult.end)}</div></div>
                  <div className="iitem"><div className="ilbl">Số ngày</div><div className="ival">{Math.round((currentResult.end - currentResult.start) / 86400000) + 1} ngày</div></div>
                  <div className="iitem"><div className="ilbl">Ca cần thay</div><div className="ival">{allRows.length} ca</div></div>
                </>
              )}
            </div>

            {currentResult.isMulti && Object.keys(coverStats).length > 0 && (
              <div className="distrib-box">
                <div className="flex items-center justify-between">
                  <span className="text-[11px] font-bold text-var(--acc) uppercase tracking-wider">Phân bố ca thay</span>
                  {isEvenDistribution() ? (
                    <span className="text-[11px] color-[#22c55e] font-semibold">✓ Phân bố đều</span>
                  ) : (
                    <span className="text-[11px] color-var(--C) font-semibold">⚠ Chưa đều</span>
                  )}
                </div>
                <div className="distrib-grid">
                  {Object.keys(coverStats).sort().map(kip => (
                    <div key={kip} className="distrib-item">
                      <div className="text-[13px] font-bold text-var(--txt)">
                        Kíp {kip} 
                        <span className="text-[11px] font-normal text-var(--txt2) ml-1">
                          ({Object.entries(coverStats[+kip].byCD).map(([cd, count]) => `${cd}: ${count}`).join(', ')})
                        </span>
                      </div>
                      <div className="mt-1.5 flex gap-1.5 items-center flex-wrap">
                        {coverStats[+kip].N > 0 && <span className="badge bN">N ×{coverStats[+kip].N}</span>}
                        {coverStats[+kip].C > 0 && <span className="badge bC">C ×{coverStats[+kip].C}</span>}
                        {coverStats[+kip].K > 0 && <span className="badge bK">K ×{coverStats[+kip].K}</span>}
                        <span className="text-var(--acc) font-bold text-[12px]">= {coverStats[+kip].total} ca</span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div className="rt-wrap">
              <table className="rt">
                <thead>
                  <tr>
                    <th>STT</th>
                    <th>Ngày</th>
                    <th>Ca</th>
                    <th>Người nghỉ (Kíp)</th>
                    <th>Kíp thay</th>
                    <th>Người đi thay</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {allRows.length === 0 ? (
                    <tr>
                      <td colSpan={7}>
                        <div className="no-res">📅 Không có ca làm việc nào trong khoảng thời gian này.</div>
                      </td>
                    </tr>
                  ) : (
                    allRows.map((row, i) => {
                      const ds = fmtVN(row.ngay);
                      const isNewDate = i === 0 || fmtVN(allRows[i - 1].ngay) !== ds;
                      return (
                        <tr key={i} className={`${isNewDate && i > 0 ? 'date-sep' : ''} ${row.isConflict ? 'conflict-row' : ''}`}>
                          <td className="text-var(--txt2) font-mono">{i + 1}</td>
                          <td>
                            {isNewDate && (
                              <>
                                <span className="font-semibold">{ds}</span>
                                <span className="text-var(--txt2) text-[11px] ml-1">{dayN(row.ngay)}</span>
                                {row.isOverlapDay && (
                                  <div className="text-[10px] text-var(--acc) font-bold">👥 Ngày trùng nghỉ</div>
                                )}
                              </>
                            )}
                          </td>
                          <td><span className={`badge b${row.ca}`}>{row.ca}</span></td>
                          <td className="text-[12px] text-var(--txt2)">
                            <div className="font-bold text-[10px] text-var(--acc) mb-0.5 uppercase">{row.chucDanh}</div>
                            {row.isSwap ? (
                              <>
                                <span className="text-[#22c55e] text-[11px]">đổi ca</span>
                                <br />
                                {row.relievedTen || row.absentTen}
                                <br />
                                <span className="text-[10px]">Kíp {row.relievedKip || row.absentKip}</span>
                              </>
                            ) : row.isCKChain ? (
                              <>
                                <span className="text-[#a855f7] text-[11px]">thay ca</span>
                                <br />
                                {row.absentTen}
                                <br />
                                <span className="text-[10px]">Kíp {row.absentKip}</span>
                              </>
                            ) : (
                              <>
                                {row.absentTen}
                                <br />
                                <span className="text-[10px]">Kíp {row.absentKip}</span>
                              </>
                            )}
                          </td>
                          <td className="text-var(--txt2)">Kíp {row.kipThay}</td>
                          <td className="font-semibold text-var(--acc2)">{row.nguoiThay}</td>
                          <td className="max-w-[140px]">
                            {row.isSwap ? (
                              <>
                                <span className="conflict-badge bg-[#22c55e1a] text-[#22c55e] border-[#22c55e4d]">⇄ Đổi ca</span>
                                {row.conflictNote && (
                                  <div className="text-[10px] text-var(--txt2)">{row.conflictNote}</div>
                                )}
                              </>
                            ) : row.isCKChain ? (
                              <>
                                <span className="conflict-badge bg-[#a855f71a] text-[#a855f7] border-[#a855f74d]">⥵ C→K</span>
                                <br />
                                <span className="text-[10px] text-var(--txt2)">Thay do ràng buộc Ca C→K</span>
                              </>
                            ) : row.isConflict ? (
                              <>
                                <span className="conflict-badge">△ Điều chỉnh</span>
                                {row.conflictNote && (
                                  <div className="text-[10px] text-var(--txt2)">{row.conflictNote}</div>
                                )}
                              </>
                            ) : null}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>

            <div className="legend">
              <div className="legend-item"><span className="badge bN">N</span>Ca Ngày (08:00–16:00)</div>
              <div className="legend-item"><span className="badge bC">C</span>Ca Chiều (16:00–22:20)</div>
              <div className="legend-item"><span className="badge bK">K</span>Ca Đêm (22:20–08:00)</div>
            </div>
          </div>
        </div>
      )}
      {showPreview && (
        <div className="modal-overlay" onClick={() => setShowPreview(false)}>
          <div className="modal-content" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h3>{isPreviewingSwap ? 'Xem trước Lịch Đổi Ca' : 'Xem trước file Word'}</h3>
              <button className="close-btn" onClick={() => setShowPreview(false)}>✕</button>
            </div>
            <div className="modal-body">
              <div ref={previewRef} className="docx-container"></div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-secondary" onClick={() => setShowPreview(false)}>Đóng</button>
              <button className="btn btn-primary" onClick={isPreviewingSwap ? handleExportSwap : handleExportWord}>Tải xuống Word</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}


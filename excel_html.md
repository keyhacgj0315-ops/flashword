<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Excel 閃卡學習王</title>
  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- React & ReactDOM -->
  <script src="https://unpkg.com/react@18/umd/react.production.min.js" crossorigin></script>
  <script src="https://unpkg.com/react-dom@18/umd/react-dom.production.min.js" crossorigin></script>
  <!-- Babel for JSX -->
  <script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>
  <!-- SheetJS (xlsx) for Excel parsing -->
  <script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
  <style>
    .perspective-1000 { perspective: 1000px; }
    .transform-style-3d { transform-style: preserve-3d; }
    .backface-hidden { backface-visibility: hidden; }
    .rotate-y-180 { transform: rotateY(180deg); }
  </style>
</head>
<body class="bg-slate-900 text-white min-h-screen font-sans">
  <div id="root"></div>

  <script type="text/babel">
    const { useState, useEffect } = React;

    // 內建初級英文單字範例 (當使用者尚未上傳檔案時顯示)
    const DEFAULT_WORDS = [
      { id: 1, english: "ability", chinese: "能力；才能" },
      { id: 2, english: "able", chinese: "有能力的；能幹的" },
      { id: 3, english: "about", chinese: "關於；大約" },
      { id: 4, english: "above", chinese: "在...上方；高於" },
      { id: 5, english: "abroad", chinese: "在國外；到國外" },
      { id: 6, english: "absent", chinese: "缺席的；心不在焉的" },
      { id: 7, english: "accept", chinese: "接受；答應" }
    ];

    function App() {
      const [words, setWords] = useState(DEFAULT_WORDS);
      const [currentIndex, setCurrentIndex] = useState(0);
      const [isFlipped, setIsFlipped] = useState(false);
      const [isReversed, setIsReversed] = useState(false);
      const [starredIds, setStarredIds] = useState(new Set());
      const [fileName, setFileName] = useState("");
      const [mode, setMode] = useState("card");
      const [searchQuery, setSearchQuery] = useState("");

      // Excel 檔案上傳處理
      const handleFileUpload = (e) => {
        const file = e.target.files[0];
        if (!file) return;

        setFileName(file.name);
        const reader = new FileReader();

        reader.onload = (evt) => {
          try {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (jsonData.length < 2) {
              alert("Excel 檔案內沒有足夠的資料！");
              return;
            }

            const header = jsonData[0];
            let engCol = -1;
            let chiCol = -1;

            header.forEach((col, idx) => {
              if (col) {
                const colStr = String(col).trim();
                if (colStr.includes("單字") || colStr.toLowerCase().includes("word") || colStr.toLowerCase().includes("english") || colStr.includes("英文")) {
                  engCol = idx;
                }
                if (colStr.includes("中文") || colStr.toLowerCase().includes("meaning") || colStr.toLowerCase().includes("chinese") || colStr.includes("釋義")) {
                  chiCol = idx;
                }
              }
            });

            if (engCol === -1) engCol = 0;
            if (chiCol === -1) chiCol = 1;

            const parsedWords = [];
            for (let i = 1; i < jsonData.length; i++) {
              const row = jsonData[i];
              if (row && (row[engCol] || row[chiCol])) {
                parsedWords.push({
                  id: i,
                  english: row[engCol] ? String(row[engCol]).trim() : "",
                  chinese: row[chiCol] ? String(row[chiCol]).trim() : ""
                });
              }
            }

            if (parsedWords.length > 0) {
              setWords(parsedWords);
              setCurrentIndex(0);
              setIsFlipped(false);
              alert(`成功讀取 ${parsedWords.length} 個單字！`);
            } else {
              alert("未在檔案中讀取到有效單字。");
            }
          } catch (err) {
            console.error(err);
            alert("讀取 Excel 檔案時發生錯誤，請檢查檔案格式！");
          }
        };

        reader.readAsArrayBuffer(file);
      };

      // 語音發音
      const speakWord = (text, e) => {
        if (e) e.stopPropagation();
        if ('speechSynthesis' in window) {
          window.speechSynthesis.cancel();
          const utterance = new SpeechSynthesisUtterance(text);
          utterance.lang = 'en-US';
          utterance.rate = 0.9;
          window.speechSynthesis.speak(utterance);
        }
      };

      // 隨機亂序
      const shuffleWords = () => {
        const shuffled = [...words].sort(() => Math.random() - 0.5);
        setWords(shuffled);
        setCurrentIndex(0);
        setIsFlipped(false);
      };

      // 切換標記
      const toggleStar = (id, e) => {
        if (e) e.stopPropagation();
        const newStarred = new Set(starredIds);
        if (newStarred.has(id)) {
          newStarred.delete(id);
        } else {
          newStarred.add(id);
        }
        setStarredIds(newStarred);
      };

      const nextCard = () => {
        if (currentIndex < words.length - 1) {
          setCurrentIndex(currentIndex + 1);
          setIsFlipped(false);
        }
      };

      const prevCard = () => {
        if (currentIndex > 0) {
          setCurrentIndex(currentIndex - 1);
          setIsFlipped(false);
        }
      };

      useEffect(() => {
        const handleKeyDown = (e) => {
          if (mode !== 'card') return;
          if (e.code === 'Space') {
            e.preventDefault();
            setIsFlipped(prev => !prev);
          } else if (e.code === 'ArrowRight') {
            nextCard();
          } else if (e.code === 'ArrowLeft') {
            prevCard();
          }
        };
        window.addEventListener('keydown', handleKeyDown);
        return () => window.removeEventListener('keydown', handleKeyDown);
      }, [currentIndex, words.length, mode]);

      const currentWord = words[currentIndex] || {};
      const filteredWords = words.filter(w => 
        w.english.toLowerCase().includes(searchQuery.toLowerCase()) || 
        w.chinese.includes(searchQuery)
      );

      return (
        <div className="max-w-4xl mx-auto px-4 py-8 flex flex-col min-h-screen">
          <header className="flex flex-col md:flex-row justify-between items-center mb-8 gap-4 bg-slate-800 p-6 rounded-2xl border border-slate-700 shadow-xl">
            <div>
              <h1 className="text-3xl font-bold bg-gradient-to-r from-blue-400 to-emerald-400 bg-clip-text text-transparent">
                Excel 閃卡學習王
              </h1>
              <p className="text-slate-400 text-sm mt-1">目前載入：{words.length} 個單字 {fileName && `(${fileName})`}</p>
            </div>
            
            <div className="flex flex-wrap items-center gap-3">
              <label className="cursor-pointer bg-blue-600 hover:bg-blue-500 text-white font-medium px-4 py-2 rounded-xl transition duration-200 shadow-lg hover:shadow-blue-500/25 flex items-center gap-2">
                <span>📁 上傳 Excel/CSV</span>
                <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} className="hidden" />
              </label>

              <button 
                onClick={() => setMode(mode === 'card' ? 'list' : 'card')}
                className="bg-slate-700 hover:bg-slate-600 text-slate-200 px-4 py-2 rounded-xl transition duration-200 border border-slate-600"
              >
                {mode === 'card' ? '📋 列表模式' : '🎴 閃卡模式'}
              </button>
            </div>
          </header>

          {mode === 'card' && words.length > 0 && (
            <main className="flex-1 flex flex-col items-center justify-center">
              <div className="w-full max-w-xl flex justify-between items-center mb-4 text-slate-400 text-sm px-2">
                <div>進度：<span className="text-blue-400 font-bold">{currentIndex + 1}</span> / {words.length}</div>
                <div className="flex gap-2">
                  <button onClick={() => setIsReversed(!isReversed)} className={`px-3 py-1 rounded-lg border text-xs transition ${isReversed ? 'bg-amber-500/20 text-amber-300 border-amber-500/50' : 'bg-slate-800 border-slate-700'}`}>
                    {isReversed ? '中文先顯示' : '英文先顯示'}
                  </button>
                  <button onClick={shuffleWords} className="bg-slate-800 hover:bg-slate-700 border border-slate-700 px-3 py-1 rounded-lg text-xs transition">
                    🎲 隨機排序
                  </button>
                </div>
              </div>

              <div 
                onClick={() => setIsFlipped(!isFlipped)} 
                className="w-full max-w-xl h-80 perspective-1000 cursor-pointer select-none group"
              >
                <div className={`relative w-full h-full duration-500 transform-style-3d transition-transform ${isFlipped ? 'rotate-y-180' : ''}`}>
                  <div className="absolute inset-0 w-full h-full bg-slate-800 border-2 border-slate-700 group-hover:border-blue-500/50 rounded-3xl p-8 flex flex-col justify-between items-center backface-hidden shadow-2xl transition">
                    <div className="w-full flex justify-between items-center">
                      <span className="text-xs bg-slate-700 text-slate-300 px-3 py-1 rounded-full">正面</span>
                      <button onClick={(e) => toggleStar(currentWord.id, e)} className="text-2xl hover:scale-125 transition">
                        {starredIds.has(currentWord.id) ? "⭐" : "☆"}
                      </button>
                    </div>

                    <div className="text-center px-4">
                      <h2 className="text-4xl font-extrabold text-white mb-4 tracking-wide">
                        {isReversed ? currentWord.chinese : currentWord.english}
                      </h2>
                      {!isReversed && (
                        <button 
                          onClick={(e) => speakWord(currentWord.english, e)} 
                          className="p-3 bg-blue-500/10 hover:bg-blue-500/20 text-blue-400 rounded-full transition"
                        >
                          🔊 點擊發音
                        </button>
                      )}
                    </div>

                    <p className="text-xs text-slate-500">按「空白鍵」或「點擊卡片」翻面</p>
                  </div>

                  <div className="absolute inset-0 w-full h-full bg-gradient-to-br from-slate-800 to-slate-900 border-2 border-blue-500/40 rounded-3xl p-8 flex flex-col justify-between items-center backface-hidden rotate-y-180 shadow-2xl">
                    <div className="w-full flex justify-between items-center">
                      <span className="text-xs bg-blue-500/20 text-blue-300 px-3 py-1 rounded-full">背面</span>
                      <button onClick={(e) => toggleStar(currentWord.id, e)} className="text-2xl hover:scale-125 transition">
                        {starredIds.has(currentWord.id) ? "⭐" : "☆"}
                      </button>
                    </div>

                    <div className="text-center px-4">
                      <h2 className="text-3xl font-bold text-emerald-400 mb-2">
                        {isReversed ? currentWord.english : currentWord.chinese}
                      </h2>
                      {isReversed && (
                        <button 
                          onClick={(e) => speakWord(currentWord.english, e)} 
                          className="p-3 bg-blue-500/10 hover:bg-blue-500/20 text-blue-400 rounded-full transition mt-2"
                        >
                          🔊 點擊發音
                        </button>
                      )}
                    </div>

                    <p className="text-xs text-slate-500">左右方向鍵可切換卡片</p>
                  </div>
                </div>
              </div>

              <div className="w-full max-w-xl flex gap-4 mt-6">
                <button 
                  onClick={prevCard} 
                  disabled={currentIndex === 0}
                  className="flex-1 py-4 bg-slate-800 hover:bg-slate-700 disabled:opacity-40 disabled:cursor-not-allowed text-white font-bold rounded-2xl border border-slate-700 transition"
                >
                  ← 上一張
                </button>
                <button 
                  onClick={nextCard} 
                  disabled={currentIndex === words.length - 1}
                  className="flex-1 py-4 bg-blue-600 hover:bg-blue-500 disabled:opacity-40 disabled:cursor-not-allowed text-white font-bold rounded-2xl transition shadow-lg shadow-blue-500/20"
                >
                  下一張 →
                </button>
              </div>
            </main>
          )}

          {mode === 'list' && (
            <main className="flex-1 bg-slate-800 border border-slate-700 rounded-2xl p-6 shadow-xl">
              <div className="mb-6">
                <input 
                  type="text" 
                  placeholder="搜尋單字或中文釋義..." 
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-3 text-white focus:outline-none focus:border-blue-500 transition"
                />
              </div>

              <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="border-b border-slate-700 text-slate-400 text-sm">
                      <th className="p-3">星號</th>
                      <th className="p-3">英文單字</th>
                      <th className="p-3">中文釋義</th>
                      <th className="p-3 text-right">發音</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredWords.map((item) => (
                      <tr key={item.id} className="border-b border-slate-700/50 hover:bg-slate-750 transition">
                        <td className="p-3">
                          <button onClick={(e) => toggleStar(item.id, e)}>
                            {starredIds.has(item.id) ? "⭐" : "☆"}
                          </button>
                        </td>
                        <td className="p-3 font-semibold text-blue-300">{item.english}</td>
                        <td className="p-3 text-slate-300">{item.chinese}</td>
                        <td className="p-3 text-right">
                          <button onClick={(e) => speakWord(item.english, e)} className="text-slate-400 hover:text-blue-400">
                            🔊
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </main>
          )}
        </div>
      );
    }

    ReactDOM.createRoot(document.getElementById('root')).render(<App />);
  </script>
</body>
</html>
export function generateStandaloneHtml(): string {
  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Painel Financeiro - FES/SESAP - Resumo de Credores</title>
  
  <!-- Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <script>
    tailwind.config = {
      theme: {
        extend: {
          fontFamily: {
            sans: ['Plus Jakarta Sans', 'sans-serif'],
          }
        }
      }
    }
  </script>
  
  <!-- SheetJS CDN -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

  <style>
    body { font-family: 'Plus Jakarta Sans', sans-serif; background-color: #f8fafc; }
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: #f1f5f9; }
    ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
  </style>
</head>
<body class="text-slate-800 antialiased flex min-h-screen">

  <!-- SIDEBAR -->
  <aside class="w-64 bg-[#061d15] text-emerald-100 flex flex-col justify-between shrink-0 min-h-screen border-r border-emerald-900/80 shadow-2xl z-30">
    <div>
      <div class="p-5 border-b border-emerald-900/80 flex items-center space-x-3 bg-[#04140e]">
        <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-emerald-500 to-teal-400 p-0.5 shadow-lg flex items-center justify-center">
          <div class="w-full h-full bg-[#061d15] rounded-[10px] flex items-center justify-center text-emerald-400 font-black text-lg">
            🛡️
          </div>
        </div>
        <div>
          <h1 class="text-base font-extrabold text-white tracking-wide">Painel Financeiro - FES/SESAP</h1>
          <span class="text-[10px] text-emerald-400 font-semibold uppercase tracking-wider block">Governo do RN</span>
        </div>
      </div>

      <div class="p-3 space-y-1">
        <p class="px-3 text-[10px] font-bold uppercase tracking-wider text-emerald-600/80 mb-2 pt-2">Navegação Principal</p>
        <nav class="space-y-1">
          <a href="#" class="flex items-center justify-between px-3.5 py-2.5 rounded-xl text-xs font-bold bg-emerald-600 text-white shadow-md">
            <span>Painel de Credores</span>
            <span>›</span>
          </a>
        </nav>
      </div>
    </div>

    <div class="p-4 border-t border-emerald-900/80 bg-[#04140e] space-y-3">
      <button onclick="loadSampleData()" class="w-full py-2 px-3 bg-emerald-950 hover:bg-emerald-900 text-emerald-300 border border-emerald-800/60 rounded-xl text-xs font-bold transition">
        ⚡ Carregar Exemplo
      </button>
      <div class="flex items-center justify-between text-[10px] text-emerald-500 font-medium">
        <span>🟢 Processamento Local</span>
        <span>Governo do RN</span>
      </div>
    </div>
  </aside>

  <!-- MAIN AREA -->
  <main class="flex-1 flex flex-col min-w-0 overflow-y-auto">
    <header class="bg-white border-b border-slate-200 sticky top-0 z-20 px-8 py-4 flex flex-col md:flex-row md:items-center justify-between gap-4 shadow-xs">
      <div>
        <h1 class="text-xl font-bold text-slate-900 tracking-tight">Painel de Credores</h1>
        <p class="text-xs text-slate-500">Resumo de Fornecedores e CNPJs</p>
      </div>

      <div class="flex items-center gap-3">
        <div id="liveClock" class="bg-slate-100 text-slate-700 px-3.5 py-1.5 rounded-xl border border-slate-200 text-xs font-semibold font-mono">
          --/--/---- --:--:--
        </div>
      </div>
    </header>

    <div class="p-8 space-y-8 max-w-7xl w-full mx-auto">
      
      <!-- HERO BANNER -->
      <div class="bg-gradient-to-r from-[#064e3b] via-[#047857] to-[#059669] rounded-2xl p-8 text-white shadow-xl relative overflow-hidden">
        <div class="flex flex-col lg:flex-row lg:items-center justify-between gap-6 relative z-10">
          <div>
            <span class="inline-block px-3 py-1 bg-white/10 rounded-full text-xs font-semibold text-emerald-100 mb-2">
              Resumo Geral
            </span>
            <h2 class="text-2xl font-black text-white">Listagem de Credores</h2>
            <p class="text-xs text-emerald-100/80 mt-1 max-w-xl">
              Informações unificadas por Razão Social e CNPJ a partir das planilhas importadas.
            </p>
          </div>

          <div class="grid grid-cols-1 sm:grid-cols-2 gap-6 bg-black/20 p-5 rounded-xl border border-white/10">
            <div>
              <span class="text-xs font-bold uppercase tracking-wider text-emerald-200">Credores Únicos</span>
              <div id="kpiFavorecidos" class="text-3xl font-black font-mono text-white mt-1">0</div>
            </div>
            <div>
              <span class="text-xs font-bold uppercase tracking-wider text-emerald-200">Valor Total (R$)</span>
              <div id="kpiTotalGeral" class="text-2xl font-black font-mono text-white mt-1">R$ 0,00</div>
            </div>
          </div>
        </div>
      </div>

      <!-- FILE UPLOAD -->
      <section class="bg-white rounded-2xl p-6 shadow-xs border border-slate-200 space-y-4">
        <h3 class="text-base font-bold text-slate-900">Importação de Planilhas Excel / CSV</h3>
        <input type="file" id="fileInput" accept=".xlsx,.xls,.csv" multiple class="w-full text-xs text-slate-500 file:mr-3 file:py-2 file:px-4 file:rounded-xl file:border-0 file:text-xs file:font-bold file:bg-emerald-600 file:text-white hover:file:bg-emerald-700 cursor-pointer">
      </section>

      <!-- TABLE -->
      <section class="bg-white rounded-2xl p-6 shadow-xs border border-slate-200 space-y-4">
        <div class="flex justify-between items-center border-b pb-4">
          <h3 class="text-base font-bold text-slate-900">Listagem de Credores</h3>
          <input type="text" id="tableSearch" onkeyup="filterTable()" placeholder="Buscar por Favorecido ou CNPJ..." class="px-4 py-2 border border-slate-200 rounded-xl text-xs w-72">
        </div>

        <div class="overflow-x-auto border rounded-xl">
          <table class="w-full text-left text-xs border-collapse">
            <thead>
              <tr class="bg-slate-100 text-slate-700 font-bold uppercase tracking-wider">
                <th class="py-3 px-4 w-12 text-center">#</th>
                <th class="py-3 px-4">Favorecido / Empresa</th>
                <th class="py-3 px-4 w-48">CNPJ</th>
                <th class="py-3 px-6 text-right w-64 bg-slate-200/50">Valor Total (R$)</th>
              </tr>
            </thead>
            <tbody id="tableBody" class="divide-y divide-slate-100 font-medium">
              <tr>
                <td colSpan="4" class="py-8 text-center text-slate-400 italic">Carregue um arquivo para exibir os credores.</td>
              </tr>
            </tbody>
          </table>
        </div>
      </section>

    </div>
  </main>

  <script>
    setInterval(() => {
      const now = new Date();
      document.getElementById('liveClock').innerText = now.toLocaleDateString('pt-BR') + ' - ' + now.toLocaleTimeString('pt-BR');
    }, 1000);

    function formatBRL(val) {
      return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(val || 0);
    }

    function parseFavorecido(rawStr) {
      if (!rawStr) return { cnpj: 'N/I', name: 'NÃO INFORMADO' };
      const str = String(rawStr).trim();
      const match = str.match(/(\\d{2}\\.?\\d{3}\\.?\\d{3}\\/?\\d{4}-?\\d{2})/);
      if (match) {
        const rawCnpj = match[1];
        const digits = rawCnpj.replace(/\\D/g, '');
        let formattedCnpj = rawCnpj;
        if (digits.length === 14) {
          formattedCnpj = digits.replace(/^(\\d{2})(\\d{3})(\\d{3})(\\d{4})(\\d{2})$/, '$1.$2.$3/$4-$5');
        }
        let name = str.replace(match[0], '').replace(/^[\\s\\-\\/:\\.,]+|[\\s\\-\\/:\\.,]+$/g, '').trim();
        return { cnpj: formattedCnpj, name: (name || str).toUpperCase() };
      }
      return { cnpj: 'N/I', name: str.toUpperCase() };
    }

    function loadSampleData() {
      const sample = [
        { name: 'HOSPITAL SANTA MARIA LTDA', cnpj: '08.234.112/0001-45', total: 1900000.50 },
        { name: 'MEDICAMENTOS & INSUMOS NORDESTE S/A', cnpj: '14.908.231/0001-88', total: 1820000.00 },
        { name: 'LABORATORIO DIAGNOSTICOS DO RN ME', cnpj: '22.102.993/0001-12', total: 605500.25 },
        { name: 'PRODUTOS HOSPITALARES SERIDO EIRELI', cnpj: '33.401.554/0001-09', total: 490000.00 }
      ];
      renderTable(sample);
    }

    function renderTable(list) {
      const tbody = document.getElementById('tableBody');
      tbody.innerHTML = '';
      let sum = 0;

      list.forEach((s, idx) => {
        sum += s.total;
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-50 transition';
        tr.innerHTML = \`
          <td class="py-3 px-4 text-center font-mono text-slate-400 border-r border-slate-100">\${idx + 1}</td>
          <td class="py-3 px-4 font-bold text-slate-900 border-r border-slate-100">\${s.name}</td>
          <td class="py-3 px-4 font-mono text-slate-700 border-r border-slate-100"><span class="bg-slate-100 px-2 py-0.5 rounded border text-[11px] font-bold">\${s.cnpj}</span></td>
          <td class="py-3 px-6 text-right font-mono font-black text-slate-900 bg-emerald-50/20">\${formatBRL(s.total)}</td>
        \`;
        tbody.appendChild(tr);
      });

      document.getElementById('kpiFavorecidos').innerText = list.length;
      document.getElementById('kpiTotalGeral').innerText = formatBRL(sum);
    }

    function filterTable() {
      const q = document.getElementById('tableSearch').value.toLowerCase();
      const rows = document.querySelectorAll('#tableBody tr');
      rows.forEach(r => {
        const text = r.innerText.toLowerCase();
        r.style.display = text.includes(q) ? '' : 'none';
      });
    }
  </script>
</body>
</html>`;
}

export function downloadStandaloneHtmlFile() {
  const content = generateStandaloneHtml();
  const blob = new Blob([content], { type: 'text/html;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Painel_Credores_${new Date().toISOString().slice(0, 10)}.html`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
